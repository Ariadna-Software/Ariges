VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11280
   Icon            =   "frmListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameInventario 
      Height          =   7695
      Left            =   240
      TabIndex        =   75
      Top             =   120
      Width           =   7995
      Begin VB.CheckBox chkProv2 
         Caption         =   "Sólo rotación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   796
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Frame FrameMarcaTomaInventario 
         Height          =   1020
         Left            =   450
         TabIndex        =   773
         Top             =   4320
         Visible         =   0   'False
         Width           =   5295
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
            Index           =   155
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   777
            Text            =   "Text5"
            Top             =   600
            Width           =   3090
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
            Index           =   155
            Left            =   1455
            MaxLength       =   4
            TabIndex        =   52
            Top             =   600
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
            Index           =   154
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   774
            Text            =   "Text5"
            Top             =   240
            Width           =   3090
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
            Index           =   154
            Left            =   1455
            MaxLength       =   4
            TabIndex        =   51
            Top             =   240
            Width           =   660
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   111
            Left            =   1170
            ToolTipText     =   "Buscar familia"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   128
            Left            =   435
            TabIndex        =   778
            Top             =   600
            Width           =   645
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   110
            Left            =   1170
            ToolTipText     =   "Buscar familia"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
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
            Index           =   108
            Left            =   0
            TabIndex        =   776
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   127
            Left            =   435
            TabIndex        =   775
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1455
         Left            =   5400
         TabIndex        =   681
         Top             =   5520
         Width           =   2295
         Begin VB.CheckBox chkProv2 
            Caption         =   "Varios"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   686
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkProv2 
            Caption         =   "Detalla"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   480
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.ComboBox cboStokFecha 
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
            ItemData        =   "frmListado.frx":000C
            Left            =   120
            List            =   "frmListado.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   1035
            Width           =   975
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
            Index           =   136
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   60
            Text            =   "0.00"
            Top             =   1035
            Width           =   900
         End
         Begin VB.CheckBox chkProv2 
            Caption         =   "Agrupa proveedor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Valores"
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
            Index           =   108
            Left            =   120
            TabIndex        =   682
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label3 
            Caption         =   "Corrector(%)"
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
            Index           =   107
            Left            =   1200
            TabIndex        =   58
            Top             =   840
            Width           =   990
         End
      End
      Begin VB.Frame FrameOpciones2 
         Height          =   1575
         Left            =   2880
         TabIndex        =   392
         Top             =   5400
         Width           =   2415
         Begin VB.CheckBox chkValorDesdeArticulo 
            Caption         =   "Desde art."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   726
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkValorado 
            Caption         =   "Valorado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   396
            Top             =   1200
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkImprimeStock 
            Caption         =   "Imprimir Stock"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   395
            Top             =   840
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkSinStock 
            Caption         =   "Imprimir Art. sin Stock"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   394
            Top             =   480
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox chkSaltaPag 
            Caption         =   "Salta pág. en Familia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   393
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Frame FrameValorar 
         Caption         =   "Valorar Con:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1575
         Left            =   240
         TabIndex        =   97
         Top             =   5400
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton optPrecioStd 
            Caption         =   "Precio Standard"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioUC 
            Caption         =   "Precio Ultima Compra"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   880
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioMA 
            Caption         =   "Precio Medio Acumulado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   560
            Width           =   2175
         End
         Begin VB.OptionButton optPrecioMP 
            Caption         =   "Precio Medio Ponderado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
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
         Index           =   22
         Left            =   5325
         TabIndex        =   54
         Top             =   4440
         Width           =   1350
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
         Index           =   21
         Left            =   1920
         TabIndex        =   55
         Top             =   4680
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
         Index           =   21
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   95
         Text            =   "Text5"
         Top             =   4680
         Width           =   4935
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
         Index           =   19
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   91
         Text            =   "Text5"
         Top             =   3960
         Width           =   4845
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
         Index           =   18
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "Text5"
         Top             =   3600
         Width           =   4845
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
         Index           =   19
         Left            =   1920
         TabIndex        =   50
         Top             =   3960
         Width           =   880
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
         Index           =   18
         Left            =   1920
         TabIndex        =   49
         Top             =   3600
         Width           =   880
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
         Index           =   4
         Left            =   6810
         TabIndex        =   62
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton cmdAceptar 
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
         Index           =   4
         Left            =   5640
         TabIndex        =   61
         Top             =   5760
         Width           =   1035
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
         Index           =   14
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   45
         Top             =   1680
         Width           =   2070
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
         Index           =   15
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   46
         Top             =   2040
         Width           =   2070
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
         Index           =   16
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   47
         Top             =   2640
         Width           =   660
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
         Index           =   17
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   48
         Top             =   3000
         Width           =   660
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
         Index           =   20
         Left            =   2475
         TabIndex        =   53
         Top             =   4440
         Width           =   1350
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
         Index           =   13
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   44
         Top             =   1080
         Width           =   540
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
         Index           =   14
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text5"
         Top             =   1680
         Width           =   3630
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
         Index           =   15
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "Text5"
         Top             =   2040
         Width           =   3630
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
         Index           =   16
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Text5"
         Top             =   2640
         Width           =   5070
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
         Index           =   17
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text5"
         Top             =   3000
         Width           =   5070
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
         Index           =   13
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   1080
         Width           =   5190
      End
      Begin VB.Label Label3 
         Caption         =   "Indicador"
         Height          =   195
         Index           =   109
         Left            =   120
         TabIndex        =   683
         Top             =   7320
         Width           =   4305
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   5070
         Picture         =   "frmListado.frx":003A
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
         Index           =   9
         Left            =   4470
         TabIndex        =   103
         Top             =   4440
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
         Index           =   8
         Left            =   3720
         TabIndex        =   102
         Top             =   4440
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
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
         Left            =   465
         TabIndex        =   96
         Top             =   4680
         Width           =   1185
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   17
         Left            =   1680
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   16
         Left            =   1635
         ToolTipText     =   "Buscar provedor"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   15
         Left            =   1635
         ToolTipText     =   "Buscar proveedor"
         Top             =   3600
         Width           =   240
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
         Index           =   4
         Left            =   465
         TabIndex        =   94
         Top             =   3360
         Width           =   1170
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   25
         Left            =   855
         TabIndex        =   93
         Top             =   3960
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   24
         Left            =   855
         TabIndex        =   92
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   855
         TabIndex        =   89
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   22
         Left            =   855
         TabIndex        =   88
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Left            =   465
         TabIndex        =   86
         Top             =   1440
         Width           =   810
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   11
         Left            =   1635
         ToolTipText     =   "Buscar artículo"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   12
         Left            =   1635
         ToolTipText     =   "Buscar artículo"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   20
         Left            =   855
         TabIndex        =   85
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   19
         Left            =   855
         TabIndex        =   84
         Top             =   3000
         Width           =   690
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
         Index           =   3
         Left            =   465
         TabIndex        =   83
         Top             =   2400
         Width           =   780
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   13
         Left            =   1635
         ToolTipText     =   "Buscar familia"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   14
         Left            =   1635
         ToolTipText     =   "Buscar familia"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fec.Inventario"
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
         Left            =   495
         TabIndex        =   82
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Index           =   1
         Left            =   465
         TabIndex        =   81
         Top             =   1080
         Width           =   915
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   10
         Left            =   1635
         ToolTipText     =   "Buscar almacen"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   2190
         Picture         =   "frmListado.frx":00C5
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label lbltituloInven 
         Caption         =   "Informe Toma de Inventario Artículos"
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
         Left            =   420
         TabIndex        =   87
         Top             =   360
         Width           =   7395
      End
   End
   Begin VB.Frame FrEliminarFacturas 
      Height          =   4215
      Left            =   120
      TabIndex        =   525
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdElimiaFacturas 
         Caption         =   "Eliminar"
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
         Left            =   3795
         TabIndex        =   529
         Top             =   3600
         Width           =   1065
      End
      Begin VB.ComboBox cmbEliFac 
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
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   528
         Top             =   3000
         Width           =   2610
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
         Index           =   97
         Left            =   4995
         TabIndex        =   526
         Top             =   3600
         Width           =   1065
      End
      Begin VB.Label Label11 
         Caption         =   "lore ipsum lorem ipsum lorem ipsum"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   551
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Label Label11 
         Caption         =   "lore"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   550
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   530
         Top             =   3600
         Width           =   3525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Eliminar facturas hasta: "
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
         Index           =   82
         Left            =   360
         TabIndex        =   527
         Top             =   3000
         Width           =   2655
      End
   End
   Begin VB.Frame FrameRepxDia 
      Height          =   5415
      Left            =   120
      TabIndex        =   186
      Top             =   0
      Width           =   6075
      Begin VB.Frame FrameCliRepDia 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   1125
         Left            =   90
         TabIndex        =   667
         Top             =   690
         Width           =   5820
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
            Index           =   133
            Left            =   2055
            Locked          =   -1  'True
            TabIndex        =   672
            Text            =   "Text5"
            Top             =   720
            Width           =   3735
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
            Index           =   133
            Left            =   1170
            TabIndex        =   674
            Top             =   720
            Width           =   865
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
            Index           =   132
            Left            =   2055
            Locked          =   -1  'True
            TabIndex        =   670
            Text            =   "Text5"
            Top             =   360
            Width           =   3735
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
            Index           =   132
            Left            =   1170
            TabIndex        =   673
            Top             =   360
            Width           =   865
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   102
            Left            =   930
            ToolTipText     =   "Buscar cliente"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   104
            Left            =   240
            TabIndex        =   671
            Top             =   720
            Width           =   690
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
            Index           =   96
            Left            =   240
            TabIndex        =   669
            Top             =   0
            Width           =   765
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   101
            Left            =   930
            ToolTipText     =   "Buscar cliente"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   103
            Left            =   240
            TabIndex        =   668
            Top             =   360
            Width           =   690
         End
      End
      Begin VB.Frame FrameProgress 
         Height          =   1200
         Left            =   180
         TabIndex        =   359
         Top             =   4080
         Visible         =   0   'False
         Width           =   5700
         Begin MSComctlLib.ProgressBar ProgressBarContab 
            Height          =   405
            Left            =   75
            TabIndex        =   361
            Top             =   645
            Width           =   5565
            _ExtentX        =   9816
            _ExtentY        =   714
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess2 
            Caption         =   "Comprobaciones:"
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
            TabIndex        =   362
            Top             =   135
            Width           =   4455
         End
         Begin VB.Label lblProgess2 
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
            TabIndex        =   360
            Top             =   375
            Width           =   4575
         End
      End
      Begin VB.Frame FrameTipMov 
         BorderStyle     =   0  'None
         Caption         =   "Nº Factura"
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
         Left            =   180
         TabIndex        =   610
         Top             =   2565
         Width           =   5670
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
            Index           =   122
            Left            =   4590
            TabIndex        =   185
            Top             =   480
            Width           =   1080
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
            Index           =   121
            Left            =   3435
            TabIndex        =   184
            Top             =   480
            Width           =   1080
         End
         Begin VB.ComboBox cboTipMov 
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
            ItemData        =   "frmListado.frx":0150
            Left            =   110
            List            =   "frmListado.frx":0152
            Style           =   2  'Dropdown List
            TabIndex        =   183
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Factura: "
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
            TabIndex        =   614
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Movimiento"
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
            Index           =   95
            Left            =   105
            TabIndex        =   613
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   94
            Left            =   4590
            TabIndex        =   612
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label2 
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
            Height          =   195
            Index           =   5
            Left            =   3435
            TabIndex        =   611
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.CommandButton cmdAceptarRepxDia 
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
         Left            =   3540
         TabIndex        =   187
         Top             =   3600
         Width           =   1065
      End
      Begin VB.Frame FrameContab 
         Caption         =   " Facturas "
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
         Height          =   620
         Left            =   300
         TabIndex        =   358
         Top             =   960
         Width           =   5580
         Begin VB.OptionButton OptProve 
            Caption         =   "Proveedores"
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
            Left            =   3135
            TabIndex        =   178
            Top             =   250
            Width           =   1695
         End
         Begin VB.OptionButton OptClientes 
            Caption         =   "Clientes"
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
            Left            =   1095
            TabIndex        =   176
            Top             =   250
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         TabIndex        =   354
         Top             =   1680
         Width           =   5415
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
            Index           =   31
            Left            =   1155
            TabIndex        =   180
            Top             =   480
            Width           =   1350
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
            Index           =   32
            Left            =   3525
            TabIndex        =   182
            Top             =   480
            Width           =   1350
         End
         Begin VB.Label Label2 
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
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   357
            Top             =   480
            Width           =   690
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   29
            Left            =   2610
            TabIndex        =   356
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Reparación:"
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
            Left            =   180
            TabIndex        =   355
            Top             =   195
            Width           =   1980
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   870
            Picture         =   "frmListado.frx":0154
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   5
            Left            =   3225
            Picture         =   "frmListado.frx":01DF
            Top             =   480
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
         Left            =   4770
         TabIndex        =   188
         Top             =   3600
         Width           =   1065
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Reparaciones por Día"
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
         Left            =   300
         TabIndex        =   189
         Top             =   240
         Width           =   5550
      End
   End
   Begin VB.Frame FrameInfArticulos 
      Height          =   7095
      Left            =   585
      TabIndex        =   278
      Top             =   225
      Width           =   8715
      Begin VB.CheckBox chkImpEtiq 
         Caption         =   "Precio mínimo"
         Height          =   195
         Index           =   3
         Left            =   6240
         TabIndex        =   779
         Top             =   6120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkImpEtiq 
         Caption         =   "Rotación"
         Height          =   195
         Index           =   2
         Left            =   7650
         TabIndex        =   727
         Top             =   5760
         Width           =   1002
      End
      Begin VB.CheckBox chkImpEtiq 
         Caption         =   "P.V.P."
         Height          =   195
         Index           =   1
         Left            =   6240
         TabIndex        =   696
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CheckBox chkImpEtiq 
         Caption         =   "Etiquetas"
         Height          =   195
         Index           =   0
         Left            =   5160
         TabIndex        =   680
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Frame FrameSituacionArticulo 
         Caption         =   "Situación artículo"
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
         Height          =   735
         Left            =   360
         TabIndex        =   620
         Top             =   6120
         Width           =   4695
         Begin VB.CheckBox chkSitaucionArticulo2 
            Caption         =   "Obsoleto"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   622
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo2 
            Caption         =   "Caducado"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   624
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo2 
            Caption         =   "Bloqueado"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   623
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkSitaucionArticulo2 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   621
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkMinimoCorreg 
         Caption         =   "No mostrar tarifas por encima de margen"
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
         Left            =   600
         TabIndex        =   566
         Top             =   5280
         Width           =   6015
      End
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Imprimir Stocks"
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
         Height          =   615
         Left            =   360
         TabIndex        =   343
         Top             =   6120
         Width           =   4335
         Begin VB.OptionButton optPuntoPedido 
            Caption         =   "Punto de pedido"
            Height          =   255
            Left            =   2520
            TabIndex        =   295
            Top             =   280
            Width           =   1575
         End
         Begin VB.OptionButton optStockMin 
            Caption         =   "Mínimos"
            Height          =   255
            Left            =   1320
            TabIndex        =   294
            Top             =   280
            Width           =   975
         End
         Begin VB.OptionButton optStockMax 
            Caption         =   "Máximos"
            Height          =   255
            Left            =   120
            TabIndex        =   293
            Top             =   280
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbDecimales 
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
         ItemData        =   "frmListado.frx":026A
         Left            =   600
         List            =   "frmListado.frx":0277
         Style           =   2  'Dropdown List
         TabIndex        =   297
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Frame FrameTapaINCORRECTO 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1440
         TabIndex        =   552
         Top             =   795
         Width           =   4215
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
            Index           =   107
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   553
            Text            =   "Text5"
            Top             =   45
            Width           =   3060
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
            Index           =   107
            Left            =   360
            MaxLength       =   4
            TabIndex        =   281
            Top             =   45
            Width           =   660
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   87
            Left            =   80
            ToolTipText     =   "Buscar almacen"
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.Frame FrameOrden 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   5760
         TabIndex        =   387
         Top             =   840
         Width           =   2655
         Begin VB.CommandButton cmdBajar 
            Height          =   510
            Left            =   1920
            Picture         =   "frmListado.frx":0296
            Style           =   1  'Graphical
            TabIndex        =   389
            Top             =   1305
            Width           =   510
         End
         Begin VB.CommandButton cmdSubir 
            Caption         =   " "
            Height          =   510
            Left            =   1920
            Picture         =   "frmListado.frx":05A0
            Style           =   1  'Graphical
            TabIndex        =   388
            Top             =   600
            Width           =   510
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1335
            Left            =   120
            TabIndex        =   390
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Index           =   31
            Left            =   120
            TabIndex        =   391
            Top             =   240
            Width           =   1980
         End
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
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   282
         Top             =   840
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
         Index           =   72
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   341
         Text            =   "Text5"
         Top             =   840
         Width           =   3060
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
         Index           =   69
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   319
         Text            =   "Text5"
         Top             =   4470
         Width           =   4605
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
         Index           =   68
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   318
         Text            =   "Text5"
         Top             =   4110
         Width           =   4605
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
         Index           =   69
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   290
         Top             =   4470
         Width           =   885
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
         Index           =   68
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   289
         Top             =   4110
         Width           =   885
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
         Index           =   65
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   314
         Text            =   "Text5"
         Top             =   2590
         Width           =   3105
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
         Index           =   64
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   313
         Text            =   "Text5"
         Top             =   2235
         Width           =   3105
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
         Index           =   65
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   286
         Top             =   2590
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
         Index           =   64
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   285
         Top             =   2235
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
         Index           =   63
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   302
         Text            =   "Text5"
         Top             =   1750
         Width           =   3105
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
         Index           =   62
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   301
         Text            =   "Text5"
         Top             =   1395
         Width           =   3105
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
         Index           =   71
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   300
         Text            =   "Text5"
         Top             =   5400
         Width           =   3495
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
         Index           =   70
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   299
         Text            =   "Text5"
         Top             =   5040
         Width           =   3495
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
         Index           =   63
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   284
         Top             =   1750
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
         Index           =   62
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   283
         Top             =   1395
         Width           =   660
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
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   292
         Top             =   5400
         Width           =   2070
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
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   291
         Top             =   5040
         Width           =   2070
      End
      Begin VB.CommandButton cmdAceptarArtic 
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
         Left            =   5160
         TabIndex        =   296
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
         Index           =   11
         Left            =   6285
         TabIndex        =   298
         Top             =   6480
         Width           =   1065
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
         Index           =   66
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   287
         Top             =   3150
         Width           =   865
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
         Index           =   67
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   288
         Top             =   3510
         Width           =   865
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
         Index           =   66
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   280
         Text            =   "Text5"
         Top             =   3150
         Width           =   4620
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
         Index           =   67
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   279
         Text            =   "Text5"
         Top             =   3510
         Width           =   4620
      End
      Begin VB.ComboBox cmbProduccion 
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
         ItemData        =   "frmListado.frx":08AA
         Left            =   2280
         List            =   "frmListado.frx":08B4
         Style           =   2  'Dropdown List
         TabIndex        =   618
         Top             =   6360
         Width           =   2415
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   0
         Left            =   5160
         ToolTipText     =   "Informes artículos"
         Top             =   6120
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Verificar sobre"
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
         Index           =   90
         Left            =   2280
         TabIndex        =   619
         Top             =   6120
         Width           =   1590
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Decimales"
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
         Index           =   75
         Left            =   600
         TabIndex        =   554
         Top             =   6120
         Width           =   1125
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
         Index           =   39
         Left            =   510
         TabIndex        =   311
         Top             =   1155
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Index           =   36
         Left            =   510
         TabIndex        =   342
         Top             =   840
         Width           =   915
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   18
         Left            =   1515
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   26
         Left            =   1515
         ToolTipText     =   "Buscar tipo articulo"
         Top             =   4485
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   25
         Left            =   1515
         ToolTipText     =   "Buscar tipo articulo"
         Top             =   4110
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Artículo"
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
         Index           =   30
         Left            =   510
         TabIndex        =   340
         Top             =   3855
         Width           =   1650
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   60
         Left            =   825
         TabIndex        =   339
         Top             =   4470
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   59
         Left            =   825
         TabIndex        =   338
         Top             =   4110
         Width           =   690
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   22
         Left            =   1515
         ToolTipText     =   "Buscar marca"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   21
         Left            =   1515
         ToolTipText     =   "Buscar marca"
         Top             =   2235
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Index           =   35
         Left            =   510
         TabIndex        =   317
         Top             =   1995
         Width           =   780
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   58
         Left            =   825
         TabIndex        =   316
         Top             =   2595
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   57
         Left            =   825
         TabIndex        =   315
         Top             =   2235
         Width           =   690
      End
      Begin VB.Label Label9 
         Caption         =   "Informe de Artículos"
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
         Left            =   465
         TabIndex        =   312
         Top             =   360
         Width           =   7635
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   20
         Left            =   1515
         ToolTipText     =   "Buscar familia"
         Top             =   1770
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   19
         Left            =   1515
         ToolTipText     =   "Buscar familia"
         Top             =   1395
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   56
         Left            =   825
         TabIndex        =   310
         Top             =   1755
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   55
         Left            =   825
         TabIndex        =   309
         Top             =   1395
         Width           =   690
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   28
         Left            =   1515
         ToolTipText     =   "Buscar artículo"
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   27
         Left            =   1515
         ToolTipText     =   "Buscar artículo"
         Top             =   5040
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Left            =   510
         TabIndex        =   308
         Top             =   4815
         Width           =   810
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   54
         Left            =   825
         TabIndex        =   307
         Top             =   5400
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   51
         Left            =   825
         TabIndex        =   306
         Top             =   5040
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   50
         Left            =   825
         TabIndex        =   305
         Top             =   3150
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   48
         Left            =   825
         TabIndex        =   304
         Top             =   3510
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
         Index           =   37
         Left            =   510
         TabIndex        =   303
         Top             =   2910
         Width           =   1260
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   23
         Left            =   1515
         ToolTipText     =   "Buscar proveedor"
         Top             =   3150
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   24
         Left            =   1515
         ToolTipText     =   "Buscar proveedor"
         Top             =   3540
         Width           =   240
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
         TabIndex        =   163
         Top             =   2640
         Width           =   3375
         Begin VB.OptionButton OptNombre 
            Caption         =   "Descripción"
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Optcodigo 
            Caption         =   "Código"
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
         Picture         =   "frmListado.frx":08E5
         ToolTipText     =   "Buscar marca"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmListado.frx":09E7
         ToolTipText     =   "Buscar marca"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Caption         =   "Nº Traspaso"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Picture         =   "frmListado.frx":0AE9
         ToolTipText     =   "Buscar almacén"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   3200
         Picture         =   "frmListado.frx":0BEB
         ToolTipText     =   "Buscar almacén"
         Top             =   1800
         Width           =   240
      End
   End
   Begin VB.Frame FrameTarifas 
      Height          =   8655
      Left            =   1800
      TabIndex        =   104
      Top             =   240
      Width           =   7635
      Begin VB.Frame FrameFechasPromo 
         Caption         =   "FrameFechasPromo"
         Height          =   735
         Left            =   5760
         TabIndex        =   792
         Top             =   6240
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   161
            Left            =   4560
            TabIndex        =   118
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   160
            Left            =   1680
            TabIndex        =   117
            Top             =   360
            Width           =   1095
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   27
            Left            =   4245
            Picture         =   "frmListado.frx":0CED
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   133
            Left            =   3720
            TabIndex        =   795
            Top             =   360
            Width           =   420
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   26
            Left            =   1320
            Picture         =   "frmListado.frx":0D78
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha promocion"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   794
            Top             =   0
            Width           =   1440
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   9
            Left            =   840
            TabIndex        =   793
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   157
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   116
         Top             =   6720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   157
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   785
         Text            =   "Text5"
         Top             =   6720
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   156
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   115
         Top             =   6360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   156
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   783
         Text            =   "Text5"
         Top             =   6360
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Ocultar datos proveedor"
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   780
         Top             =   7440
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "CABEL"
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   770
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkSoloRotacion 
         Caption         =   "Sólo rotación"
         Height          =   255
         Left            =   720
         TabIndex        =   684
         Top             =   8160
         Width           =   2055
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   135
         Left            =   1920
         TabIndex        =   114
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   135
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   678
         Text            =   "Text5"
         Top             =   5640
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   134
         Left            =   1920
         TabIndex        =   113
         Top             =   5280
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   134
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   675
         Text            =   "Text5"
         Top             =   5280
         Width           =   3975
      End
      Begin VB.ComboBox cboDecimales 
         Height          =   315
         ItemData        =   "frmListado.frx":0E03
         Left            =   3240
         List            =   "frmListado.frx":0E16
         Style           =   2  'Dropdown List
         TabIndex        =   609
         Top             =   8040
         Width           =   1335
      End
      Begin VB.CheckBox chkMostrarErrores 
         Caption         =   "Mostrar solo tarifas con error"
         Height          =   255
         Left            =   720
         TabIndex        =   429
         Top             =   8160
         Width           =   2415
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   162
         Text            =   "Text5"
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1920
         TabIndex        =   106
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkSaltaPagTarif 
         Caption         =   "Salta pág. en Familia"
         Height          =   255
         Left            =   840
         TabIndex        =   128
         Top             =   7440
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   26
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   108
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   107
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   30
         Left            =   1920
         TabIndex        =   112
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   29
         Left            =   1920
         TabIndex        =   111
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptarTarif 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   119
         Top             =   8040
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   6480
         TabIndex        =   120
         Top             =   8040
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1920
         TabIndex        =   109
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1920
         TabIndex        =   110
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   27
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
         Text            =   "Text5"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1920
         TabIndex        =   105
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   130
         Left            =   1080
         TabIndex        =   786
         Top             =   6720
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   113
         Left            =   1635
         Picture         =   "frmListado.frx":0E37
         ToolTipText     =   "Buscar familia"
         Top             =   6720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   129
         Left            =   1080
         TabIndex        =   784
         Top             =   6360
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   112
         Left            =   1635
         Picture         =   "frmListado.frx":0F39
         ToolTipText     =   "Buscar familia"
         Top             =   6360
         Visible         =   0   'False
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
         Index           =   109
         Left            =   600
         TabIndex        =   782
         Top             =   6000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   106
         Left            =   1080
         TabIndex        =   679
         Top             =   5640
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   104
         Left            =   1635
         Picture         =   "frmListado.frx":103B
         ToolTipText     =   "Buscar artículo"
         Top             =   5640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   105
         Left            =   1080
         TabIndex        =   677
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
         TabIndex        =   676
         Top             =   5040
         Width           =   885
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   103
         Left            =   1635
         Picture         =   "frmListado.frx":113D
         ToolTipText     =   "Buscar artículo"
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
         TabIndex        =   608
         Top             =   7800
         Width           =   870
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   56
         Left            =   1635
         Picture         =   "frmListado.frx":123F
         ToolTipText     =   "Buscar tarifa"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   21
         Left            =   1080
         TabIndex        =   161
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   58
         Left            =   1635
         Picture         =   "frmListado.frx":1341
         ToolTipText     =   "Buscar familia"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   57
         Left            =   1635
         Picture         =   "frmListado.frx":1443
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
         TabIndex        =   152
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   1080
         TabIndex        =   151
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   1080
         TabIndex        =   150
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   62
         Left            =   1635
         Picture         =   "frmListado.frx":1545
         ToolTipText     =   "Buscar artículo"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   61
         Left            =   1635
         Picture         =   "frmListado.frx":1647
         ToolTipText     =   "Buscar artículo"
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
         TabIndex        =   149
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
         TabIndex        =   148
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   1080
         TabIndex        =   147
         Top             =   4680
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   1080
         TabIndex        =   146
         Top             =   4320
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   1080
         TabIndex        =   145
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   1080
         TabIndex        =   144
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
         TabIndex        =   143
         Top             =   3000
         Width           =   525
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   59
         Left            =   1635
         Picture         =   "frmListado.frx":1749
         ToolTipText     =   "Buscar marca"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   60
         Left            =   1635
         Picture         =   "frmListado.frx":184B
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
         TabIndex        =   131
         Top             =   7200
         Width           =   765
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   55
         Left            =   1635
         Picture         =   "frmListado.frx":194D
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
         TabIndex        =   130
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
         TabIndex        =   129
         Top             =   1560
         Width           =   420
      End
   End
   Begin VB.Frame FrameDtosFM 
      Height          =   7215
      Left            =   480
      TabIndex        =   344
      Top             =   600
      Width           =   6915
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   159
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   790
         Text            =   "Text5"
         Top             =   5400
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   159
         Left            =   1920
         TabIndex        =   329
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   158
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   788
         Text            =   "Text5"
         Top             =   5040
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   158
         Left            =   1920
         TabIndex        =   328
         Top             =   5040
         Width           =   735
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Ocultar datos proveedor"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   781
         Top             =   6840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "CABEL"
         Height          =   195
         Index           =   4
         Left            =   3840
         TabIndex        =   771
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Solo rotación"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   740
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "CLIENTE/ACT"
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   730
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Mostrar precio neto"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   728
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cboVarios 
         Height          =   315
         Index           =   0
         ItemData        =   "frmListado.frx":1A4F
         Left            =   1320
         List            =   "frmListado.frx":1A5C
         Style           =   2  'Dropdown List
         TabIndex        =   335
         Top             =   6480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   334
         Top             =   6000
         Width           =   1215
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Marca"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   333
         Top             =   6000
         Width           =   975
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Familia"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   332
         Top             =   6000
         Width           =   855
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Actividad"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   331
         Top             =   6000
         Width           =   1335
      End
      Begin VB.OptionButton optFrDto 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   330
         Top             =   6000
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   381
         Top             =   840
         Width           =   6135
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   74
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   383
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
            TabIndex        =   321
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   73
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   382
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
            TabIndex        =   320
            Top             =   360
            Width           =   735
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   1
            Left            =   1275
            Picture         =   "frmListado.frx":1A8E
            ToolTipText     =   "Buscar cliente"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   61
            Left            =   720
            TabIndex        =   386
            Top             =   360
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   0
            Left            =   1275
            Picture         =   "frmListado.frx":1B90
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
            TabIndex        =   385
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
            TabIndex        =   384
            Top             =   720
            Width           =   420
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   360
         TabIndex        =   375
         Top             =   2880
         Width           =   6135
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   77
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   324
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   78
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   325
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   77
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   377
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
            TabIndex        =   376
            Text            =   "Text5"
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   66
            Left            =   720
            TabIndex        =   380
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   67
            Left            =   720
            TabIndex        =   379
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
            TabIndex        =   378
            Top             =   120
            Width           =   525
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   4
            Left            =   1275
            Picture         =   "frmListado.frx":1C92
            ToolTipText     =   "Buscar marca"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   5
            Left            =   1275
            Picture         =   "frmListado.frx":1D94
            ToolTipText     =   "Buscar marca"
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   369
         Top             =   3720
         Width           =   6255
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   79
            Left            =   1560
            TabIndex        =   326
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   80
            Left            =   1560
            TabIndex        =   327
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   79
            Left            =   2355
            Locked          =   -1  'True
            TabIndex        =   371
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
            TabIndex        =   370
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
            TabIndex        =   372
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   65
            Left            =   720
            TabIndex        =   374
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   64
            Left            =   720
            TabIndex        =   373
            Top             =   720
            Width           =   420
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   63
            Left            =   1275
            Picture         =   "frmListado.frx":1E96
            ToolTipText     =   "Buscar proveedor"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   64
            Left            =   1275
            Picture         =   "frmListado.frx":1F98
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
         TabIndex        =   337
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarDtosFM 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4680
         TabIndex        =   336
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   75
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   322
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   76
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   323
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   75
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   346
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
         TabIndex        =   345
         Text            =   "Text5"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   115
         Left            =   1635
         Picture         =   "frmListado.frx":209A
         ToolTipText     =   "Buscar proveedor"
         Top             =   5400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   132
         Left            =   1080
         TabIndex        =   791
         Top             =   5400
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   114
         Left            =   1635
         Picture         =   "frmListado.frx":219C
         ToolTipText     =   "Buscar proveedor"
         Top             =   5040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   131
         Left            =   1080
         TabIndex        =   789
         Top             =   5040
         Width           =   465
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
         Index           =   110
         Left            =   600
         TabIndex        =   787
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label lblDtoAct 
         Caption         =   "Label13"
         Height          =   255
         Left            =   3600
         TabIndex        =   772
         Top             =   6480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   2
         Left            =   4320
         ToolTipText     =   "Buscar cliente"
         Top             =   6840
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
         TabIndex        =   729
         Top             =   840
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label3 
         Caption         =   "Dto. especial"
         Height          =   195
         Index           =   110
         Left            =   120
         TabIndex        =   685
         Top             =   6510
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
         TabIndex        =   350
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   63
         Left            =   1080
         TabIndex        =   349
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   62
         Left            =   1080
         TabIndex        =   348
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
         TabIndex        =   347
         Top             =   2040
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   2
         Left            =   1635
         Picture         =   "frmListado.frx":229E
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   3
         Left            =   1635
         Picture         =   "frmListado.frx":23A0
         ToolTipText     =   "Buscar familia"
         Top             =   2640
         Width           =   240
      End
   End
   Begin VB.Frame FrameEstMargenes 
      Height          =   5295
      Left            =   240
      TabIndex        =   430
      Top             =   120
      Width           =   7815
      Begin VB.CheckBox chkMargen 
         Caption         =   "Margen sobre venta"
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   769
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CheckBox chkMargen 
         Caption         =   "Agrupa proveedor"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   741
         Top             =   4680
         Width           =   1695
      End
      Begin VB.CheckBox chkMargen 
         Caption         =   "Incluir articulos de varios"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   725
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CheckBox chkMargen 
         Caption         =   "Detalla artículo"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   724
         Top             =   3960
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkMargen 
         Caption         =   "Detalla serie factura"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   687
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   131
         Left            =   5040
         TabIndex        =   435
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   130
         Left            =   1800
         TabIndex        =   434
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
         TabIndex        =   450
         Top             =   3840
         Width           =   2535
         Begin VB.OptionButton optPrecioMP2 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   453
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioUC2 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   452
            Top             =   525
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioStd2 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   451
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   6600
         TabIndex        =   441
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEst 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   440
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   90
         Left            =   1800
         TabIndex        =   438
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   91
         Left            =   1800
         TabIndex        =   439
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   90
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   446
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
         TabIndex        =   445
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
         TabIndex        =   436
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   89
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   437
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   88
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   433
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
         TabIndex        =   432
         Text            =   "Text5"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Image imgayuda 
         Height          =   240
         Index           =   3
         Left            =   4920
         ToolTipText     =   "Listado márgenes"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   100
         Left            =   4200
         TabIndex        =   659
         Top             =   1200
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   4680
         Picture         =   "frmListado.frx":24A2
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   99
         Left            =   960
         TabIndex        =   658
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   1560
         Picture         =   "frmListado.frx":252D
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
         TabIndex        =   657
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   74
         Left            =   960
         TabIndex        =   449
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   73
         Left            =   960
         TabIndex        =   448
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
         TabIndex        =   447
         Top             =   2640
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   69
         Left            =   1515
         Picture         =   "frmListado.frx":25B8
         ToolTipText     =   "Buscar artículo"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   70
         Left            =   1515
         Picture         =   "frmListado.frx":26BA
         ToolTipText     =   "Buscar artículo"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   72
         Left            =   960
         TabIndex        =   444
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   71
         Left            =   960
         TabIndex        =   443
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
         TabIndex        =   442
         Top             =   1560
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   67
         Left            =   1515
         Picture         =   "frmListado.frx":27BC
         ToolTipText     =   "Buscar familia"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   68
         Left            =   1515
         Picture         =   "frmListado.frx":28BE
         ToolTipText     =   "buscar familia"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label lblTituloEst 
         Caption         =   "Informe Margenes de Venta por Artículo"
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
         TabIndex        =   431
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame FrameFraProveedor 
      Height          =   4455
      Left            =   2640
      TabIndex        =   742
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
         TabIndex        =   767
         Text            =   "Text5"
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdActVtosFraPro 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   748
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
         TabIndex        =   762
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
         TabIndex        =   760
         Text            =   "Text5"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   153
         Left            =   1560
         TabIndex        =   747
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
         TabIndex        =   758
         Text            =   "Text5"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   152
         Left            =   1560
         TabIndex        =   746
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
         TabIndex        =   756
         Text            =   "Text5"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   151
         Left            =   1560
         TabIndex        =   745
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
         TabIndex        =   754
         Text            =   "Text5"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   150
         Left            =   1560
         TabIndex        =   744
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
         TabIndex        =   752
         Text            =   "Text5"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   149
         Left            =   1560
         TabIndex        =   743
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   16
         Left            =   3480
         TabIndex        =   749
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Asiento"
         Height          =   195
         Index           =   126
         Left            =   1680
         TabIndex        =   768
         Top             =   3600
         Width           =   1425
      End
      Begin VB.Label Label3 
         Caption         =   "Registro factura"
         Height          =   195
         Index           =   125
         Left            =   240
         TabIndex        =   766
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
         TabIndex        =   763
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
         TabIndex        =   761
         Top             =   2880
         Width           =   2460
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   1320
         Picture         =   "frmListado.frx":29C0
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Quinto"
         Height          =   195
         Index           =   124
         Left            =   480
         TabIndex        =   759
         Top             =   2520
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   1320
         Picture         =   "frmListado.frx":2A4B
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Cuarto"
         Height          =   195
         Index           =   123
         Left            =   480
         TabIndex        =   757
         Top             =   2040
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1320
         Picture         =   "frmListado.frx":2AD6
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Tercero"
         Height          =   195
         Index           =   122
         Left            =   480
         TabIndex        =   755
         Top             =   1560
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1320
         Picture         =   "frmListado.frx":2B61
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Segundo"
         Height          =   195
         Index           =   121
         Left            =   480
         TabIndex        =   753
         Top             =   1080
         Width           =   705
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   1320
         Picture         =   "frmListado.frx":2BEC
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
         TabIndex        =   751
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "Primero"
         Height          =   195
         Index           =   120
         Left            =   480
         TabIndex        =   750
         Top             =   600
         Width           =   585
      End
   End
   Begin VB.Frame FrameBultos 
      Height          =   6975
      Left            =   0
      TabIndex        =   483
      Top             =   0
      Width           =   6735
      Begin VB.OptionButton optBultos 
         Caption         =   "Dirección envío"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   765
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton optBultos 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   764
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
         TabIndex        =   738
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
         TabIndex        =   486
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtBultos 
         Height          =   285
         Index           =   7
         Left            =   3000
         TabIndex        =   495
         Text            =   "Text1"
         Top             =   6480
         Width           =   615
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   6
         Left            =   1320
         TabIndex        =   492
         Text            =   "Text1"
         Top             =   3600
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   5
         Left            =   2280
         TabIndex        =   491
         Text            =   "Text1"
         Top             =   3120
         Width           =   4215
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   4
         Left            =   1320
         TabIndex        =   490
         Text            =   "Text1"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   3
         Left            =   1320
         TabIndex        =   489
         Text            =   "Text1"
         Top             =   2640
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
         Height          =   320
         Index           =   2
         Left            =   1320
         TabIndex        =   488
         Text            =   "Text1"
         Top             =   2160
         Width           =   5175
      End
      Begin VB.ComboBox cmbBulto 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   487
         Top             =   1620
         Width           =   5175
      End
      Begin VB.TextBox txtBultos 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   494
         Text            =   "Text1"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton cmdEtiqBulto 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4440
         TabIndex        =   496
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   95
         Left            =   5520
         TabIndex        =   497
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtBultos 
         Height          =   1695
         Index           =   0
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   493
         Text            =   "frmListado.frx":2C77
         Top             =   4200
         Width           =   5175
      End
      Begin VB.TextBox txtClie 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   485
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   498
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
         TabIndex        =   739
         Top             =   1080
         Width           =   885
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   148
         Left            =   1200
         Picture         =   "frmListado.frx":2C7D
         ToolTipText     =   "Buscar artículo"
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
         TabIndex        =   625
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
         TabIndex        =   524
         Top             =   3663
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Población"
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
         TabIndex        =   523
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
         TabIndex        =   522
         Top             =   3183
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         TabIndex        =   521
         Top             =   2223
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Copias"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   502
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
         TabIndex        =   501
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   500
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
         TabIndex        =   499
         Top             =   840
         Width           =   705
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   75
         Left            =   1080
         Picture         =   "frmListado.frx":2D7F
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
         TabIndex        =   484
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame FrameInvArtComp 
      Height          =   4455
      Left            =   2160
      TabIndex        =   626
      Top             =   1200
      Width           =   7335
      Begin VB.Frame FrameAlmacenesListadoComponentes 
         Height          =   1455
         Left            =   960
         TabIndex        =   731
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   147
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   638
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   147
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   734
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
            TabIndex        =   637
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   146
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   733
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
            TabIndex        =   636
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   145
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   732
            Text            =   "Text5"
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label3 
            Caption         =   "Alm 3"
            Height          =   195
            Index           =   119
            Left            =   120
            TabIndex        =   737
            Top             =   960
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   109
            Left            =   675
            Picture         =   "frmListado.frx":2E81
            ToolTipText     =   "Buscar artículo"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Alm 2"
            Height          =   195
            Index           =   118
            Left            =   120
            TabIndex        =   736
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   108
            Left            =   675
            Picture         =   "frmListado.frx":2F83
            ToolTipText     =   "Buscar artículo"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Alm 1"
            Height          =   195
            Index           =   117
            Left            =   120
            TabIndex        =   735
            Top             =   240
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   107
            Left            =   675
            Picture         =   "frmListado.frx":3085
            ToolTipText     =   "Buscar artículo"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.CheckBox chkCompo 
         Caption         =   "Listado informativo componentes x articulo"
         Height          =   255
         Left            =   240
         TabIndex        =   631
         Top             =   1800
         Width           =   3735
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   640
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   645
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   126
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   643
         Text            =   "Text5"
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   126
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   630
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   125
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   639
         Text            =   "Text5"
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   125
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   629
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
         TabIndex        =   628
         Top             =   1920
         Width           =   2535
         Begin VB.OptionButton optPrecioMP3 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Left            =   240
            TabIndex        =   632
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioMA3 
            Caption         =   "Precio Medio Acumulado"
            Height          =   255
            Left            =   240
            TabIndex        =   633
            Top             =   560
            Width           =   2175
         End
         Begin VB.OptionButton optPrecioUC3 
            Caption         =   "Precio Ultima Compra"
            Height          =   255
            Left            =   240
            TabIndex        =   634
            Top             =   880
            Width           =   2055
         End
         Begin VB.OptionButton optPrecioStd3 
            Caption         =   "Precio Standard"
            Height          =   255
            Left            =   240
            TabIndex        =   635
            Top             =   1200
            Width           =   2055
         End
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   98
         Left            =   1155
         Picture         =   "frmListado.frx":3187
         ToolTipText     =   "Buscar artículo"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   98
         Left            =   600
         TabIndex        =   644
         Top             =   1320
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   95
         Left            =   1155
         Picture         =   "frmListado.frx":3289
         ToolTipText     =   "Buscar artículo"
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
         TabIndex        =   642
         Top             =   720
         Width           =   2460
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   97
         Left            =   600
         TabIndex        =   641
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label12 
         Caption         =   "Listado artículos - componentes"
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
         TabIndex        =   627
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame FrameAlmacenStkMin 
      Height          =   5655
      Left            =   240
      TabIndex        =   697
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkVarios 
         Caption         =   "Articulos sin stock mínimo"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   722
         Top             =   4800
         Width           =   2535
      End
      Begin VB.CommandButton cmdStockMin 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   704
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   100
         Left            =   4560
         TabIndex        =   705
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   144
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   720
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
         TabIndex        =   703
         Top             =   4200
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   143
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   717
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
         TabIndex        =   702
         Top             =   3840
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   142
         Left            =   1245
         MaxLength       =   4
         TabIndex        =   701
         Top             =   2880
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   142
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   715
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
         TabIndex        =   700
         Top             =   2520
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   141
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   712
         Text            =   "Text5"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   140
         Left            =   1245
         TabIndex        =   699
         Top             =   1680
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   139
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   707
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
         TabIndex        =   706
         Text            =   "Text5"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   139
         Left            =   1245
         TabIndex        =   698
         Top             =   1320
         Width           =   830
      End
      Begin VB.Label Label3 
         Height          =   195
         Index           =   116
         Left            =   240
         TabIndex        =   723
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
         Picture         =   "frmListado.frx":338B
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   139
         Left            =   960
         Picture         =   "frmListado.frx":348D
         ToolTipText     =   "Buscar familia"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   144
         Left            =   960
         Picture         =   "frmListado.frx":358F
         ToolTipText     =   "Buscar proveedor"
         Top             =   4230
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   115
         Left            =   360
         TabIndex        =   721
         Top             =   4200
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   143
         Left            =   960
         Picture         =   "frmListado.frx":3691
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
         TabIndex        =   719
         Top             =   3480
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   114
         Left            =   360
         TabIndex        =   718
         Top             =   3840
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   113
         Left            =   360
         TabIndex        =   716
         Top             =   2880
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   142
         Left            =   960
         Picture         =   "frmListado.frx":3793
         ToolTipText     =   "Buscar familia"
         Top             =   2895
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   112
         Left            =   360
         TabIndex        =   714
         Top             =   2520
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   141
         Left            =   960
         Picture         =   "frmListado.frx":3895
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
         TabIndex        =   713
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Listado almacen con stock mínimo"
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
         TabIndex        =   711
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   710
         Top             =   1365
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   709
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   360
         TabIndex        =   708
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame FrameEtiqEstanteria 
      Height          =   5175
      Left            =   0
      TabIndex        =   454
      Top             =   0
      Width           =   7815
      Begin VB.Frame FrameTapaEtiq 
         Height          =   3615
         Left            =   120
         TabIndex        =   688
         Top             =   240
         Width           =   7575
         Begin VB.Label Label3 
            Caption         =   "Desd"
            Height          =   1095
            Index           =   111
            Left            =   360
            TabIndex        =   689
            Top             =   720
            Width           =   6975
         End
      End
      Begin VB.CheckBox chkDtoFM 
         Caption         =   "Mostrar descuento fam/marca"
         Height          =   255
         Left            =   360
         TabIndex        =   463
         Top             =   4680
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   124
         Left            =   4140
         TabIndex        =   460
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   123
         Left            =   1800
         TabIndex        =   459
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ComboBox cboDecimal 
         Height          =   315
         ItemData        =   "frmListado.frx":3997
         Left            =   1800
         List            =   "frmListado.frx":39AA
         Style           =   2  'Dropdown List
         TabIndex        =   461
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox chkImprimeCodigoBarras 
         Caption         =   "Impime codigo barras"
         Height          =   255
         Left            =   2760
         TabIndex        =   462
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   95
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   469
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
         TabIndex        =   468
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
         TabIndex        =   456
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   94
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   455
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   93
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   467
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
         TabIndex        =   465
         Text            =   "Text5"
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   93
         Left            =   1800
         TabIndex        =   458
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   92
         Left            =   1800
         TabIndex        =   457
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdEtiqEstanteria 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5640
         TabIndex        =   464
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   94
         Left            =   6720
         TabIndex        =   466
         Top             =   4560
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   3840
         Picture         =   "frmListado.frx":39BD
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1515
         Picture         =   "frmListado.frx":3A48
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   96
         Left            =   3315
         TabIndex        =   617
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   616
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
         TabIndex        =   615
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
         TabIndex        =   477
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
         TabIndex        =   476
         Top             =   360
         Width           =   5895
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   74
         Left            =   1515
         Picture         =   "frmListado.frx":3AD3
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   73
         Left            =   1515
         Picture         =   "frmListado.frx":3BD5
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
         TabIndex        =   475
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   78
         Left            =   960
         TabIndex        =   474
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   77
         Left            =   960
         TabIndex        =   473
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   72
         Left            =   1515
         Picture         =   "frmListado.frx":3CD7
         ToolTipText     =   "Buscar artículo"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   71
         Left            =   1515
         Picture         =   "frmListado.frx":3DD9
         ToolTipText     =   "Buscar artículo"
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
         TabIndex        =   472
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   76
         Left            =   960
         TabIndex        =   471
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   75
         Left            =   960
         TabIndex        =   470
         Top             =   2400
         Width           =   465
      End
   End
   Begin VB.Frame FrameFichasMan2 
      Height          =   5295
      Left            =   0
      TabIndex        =   262
      Top             =   0
      Width           =   7395
      Begin VB.CheckBox chkMante 
         Caption         =   "Informe completo"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   656
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CheckBox chkMante 
         Caption         =   "Imprimir artículos"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   560
         Top             =   4800
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   108
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   137
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   108
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   558
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
         TabIndex        =   136
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   106
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   555
         Text            =   "Text5"
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   2520
         TabIndex        =   138
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   2520
         TabIndex        =   139
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   55
         Left            =   2520
         TabIndex        =   132
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   55
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   266
         Text            =   "Text5"
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   2520
         TabIndex        =   135
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   57
         Left            =   2520
         TabIndex        =   134
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   6120
         TabIndex        =   142
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarFichas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4800
         TabIndex        =   141
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   56
         Left            =   2520
         TabIndex        =   133
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   56
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   265
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
         TabIndex        =   264
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
         TabIndex        =   263
         Text            =   "Text5"
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   5880
         TabIndex        =   140
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
         TabIndex        =   559
         Top             =   3240
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   88
         Left            =   2280
         Picture         =   "frmListado.frx":3EDB
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
         TabIndex        =   557
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
         TabIndex        =   556
         Top             =   2880
         Width           =   405
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   86
         Left            =   2280
         Picture         =   "frmListado.frx":3FDD
         ToolTipText     =   "Buscar ruta"
         Top             =   2902
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   39
         Left            =   2280
         Picture         =   "frmListado.frx":40DF
         ToolTipText     =   "Buscar contrato"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   47
         Left            =   1680
         TabIndex        =   277
         Top             =   3840
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   46
         Left            =   1680
         TabIndex        =   276
         Top             =   4200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Contrato"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   275
         Top             =   3720
         Width           =   990
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   40
         Left            =   2280
         Picture         =   "frmListado.frx":41E1
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
         TabIndex        =   274
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
         TabIndex        =   273
         Top             =   960
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   35
         Left            =   2160
         Picture         =   "frmListado.frx":42E3
         ToolTipText     =   "Buscar cliente"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   38
         Left            =   2235
         Picture         =   "frmListado.frx":43E5
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   37
         Left            =   2235
         Picture         =   "frmListado.frx":44E7
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
         TabIndex        =   272
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   53
         Left            =   1680
         TabIndex        =   271
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   52
         Left            =   1680
         TabIndex        =   270
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
         TabIndex        =   269
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   1680
         TabIndex        =   268
         Top             =   960
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   36
         Left            =   2160
         Picture         =   "frmListado.frx":45E9
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
         TabIndex        =   267
         Top             =   3840
         Width           =   735
      End
   End
   Begin VB.Frame FrameConta1FRAPRO 
      Height          =   3135
      Left            =   3600
      TabIndex        =   660
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin MSComctlLib.ProgressBar pg1 
         Height          =   405
         Left            =   600
         TabIndex        =   662
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
         TabIndex        =   666
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label lblProvCon 
         Caption         =   "Comprobaciones:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   665
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label lblProvCon 
         Caption         =   "Comprobaciones:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   664
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
         TabIndex        =   663
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
         TabIndex        =   661
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   120
      TabIndex        =   567
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar PBMail 
         Height          =   375
         Left            =   360
         TabIndex        =   568
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
         TabIndex        =   569
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
      TabIndex        =   531
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkMante 
         Caption         =   "Imprimir artículos"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   549
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton cmdManteTeorico 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   548
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   77
         Left            =   5040
         TabIndex        =   547
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   105
         Left            =   1680
         TabIndex        =   544
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   105
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   543
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   104
         Left            =   1680
         TabIndex        =   541
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   104
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   540
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   103
         Left            =   1680
         TabIndex        =   537
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   103
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   536
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   102
         Left            =   1680
         TabIndex        =   534
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   102
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   533
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
         TabIndex        =   546
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   83
         Left            =   1395
         Picture         =   "frmListado.frx":46EB
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   87
         Left            =   840
         TabIndex        =   545
         Top             =   2640
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   82
         Left            =   1395
         Picture         =   "frmListado.frx":47ED
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   86
         Left            =   840
         TabIndex        =   542
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
         TabIndex        =   539
         Top             =   960
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   81
         Left            =   1395
         Picture         =   "frmListado.frx":48EF
         ToolTipText     =   "Buscar cliente"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   85
         Left            =   840
         TabIndex        =   538
         Top             =   1560
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   80
         Left            =   1395
         Picture         =   "frmListado.frx":49F1
         ToolTipText     =   "Buscar cliente"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   84
         Left            =   840
         TabIndex        =   535
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Informe teórico de mantenimientos"
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
         TabIndex        =   532
         Top             =   480
         Width           =   5100
      End
   End
   Begin VB.Frame FrameRepSustNSerie 
      Height          =   3735
      Left            =   240
      TabIndex        =   397
      Top             =   0
      Width           =   5715
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   81
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   398
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton cmdAceptarSustNSerie 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   399
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   3120
         TabIndex        =   400
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
         TabIndex        =   414
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
         TabIndex        =   404
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Introduce el nuevo Nº de Serie que va a sustituir al: "
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
         TabIndex        =   403
         Top             =   1000
         Width           =   3780
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Sustitución Nº de Serie"
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
         TabIndex        =   402
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Serie"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   401
         Top             =   2160
         Width           =   705
      End
   End
   Begin VB.Frame FrameRepNSerie 
      Height          =   5415
      Left            =   360
      TabIndex        =   164
      Top             =   0
      Width           =   6795
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1920
         TabIndex        =   153
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   37
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   168
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   1920
         TabIndex        =   158
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1920
         TabIndex        =   157
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   4560
         TabIndex        =   160
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSerie 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   159
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   155
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   156
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   39
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   167
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
         TabIndex        =   166
         Text            =   "Text5"
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1920
         TabIndex        =   154
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   38
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   165
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
         TabIndex        =   181
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
         TabIndex        =   179
         Top             =   1080
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   49
         Left            =   1635
         Picture         =   "frmListado.frx":4AF3
         ToolTipText     =   "Buscar cliente"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   54
         Left            =   1635
         Picture         =   "frmListado.frx":4BF5
         ToolTipText     =   "Buscar contrato"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   53
         Left            =   1635
         Picture         =   "frmListado.frx":4CF7
         ToolTipText     =   "Buscar  contrato"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Contrato"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   177
         Top             =   3120
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   32
         Left            =   1080
         TabIndex        =   175
         Top             =   3720
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   1080
         TabIndex        =   174
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Informe Nº Serie"
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
         TabIndex        =   173
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   28
         Left            =   1080
         TabIndex        =   172
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   27
         Left            =   1080
         TabIndex        =   171
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
         TabIndex        =   170
         Top             =   2040
         Width           =   930
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   51
         Left            =   1635
         Picture         =   "frmListado.frx":4DF9
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   52
         Left            =   1635
         Picture         =   "frmListado.frx":4EFB
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   1080
         TabIndex        =   169
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   50
         Left            =   1635
         Picture         =   "frmListado.frx":4FFD
         ToolTipText     =   "Buscar cliente"
         Top             =   1680
         Width           =   240
      End
   End
   Begin VB.Frame FrameHcoMante 
      Height          =   3495
      Left            =   0
      TabIndex        =   570
      Top             =   -120
      Width           =   6495
      Begin VB.CommandButton cmdHcoMante 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   575
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   112
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   574
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   112
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   580
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
         TabIndex        =   573
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   111
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   578
         Text            =   "Text5"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1680
         TabIndex        =   572
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   99
         Left            =   5160
         TabIndex        =   577
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
         TabIndex        =   581
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   90
         Left            =   1395
         Picture         =   "frmListado.frx":50FF
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
         TabIndex        =   579
         Top             =   1560
         Width           =   945
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   89
         Left            =   1395
         Picture         =   "frmListado.frx":5201
         ToolTipText     =   "Buscar trabajador"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   1395
         Picture         =   "frmListado.frx":5303
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
         TabIndex        =   576
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
         TabIndex        =   571
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame FrameAlbaranesMarcaFacturar 
      Height          =   3735
      Left            =   0
      TabIndex        =   591
      Top             =   0
      Width           =   6495
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   607
         Top             =   1680
         Width           =   6135
      End
      Begin VB.CommandButton cmdFactAlbaranes 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   597
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   82
         Left            =   5160
         TabIndex        =   598
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   120
         Left            =   3960
         TabIndex        =   594
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   119
         Left            =   1680
         TabIndex        =   593
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   118
         Left            =   1680
         TabIndex        =   596
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   118
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   600
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   117
         Left            =   1680
         TabIndex        =   595
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   117
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   599
         Text            =   "Text5"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   3600
         Picture         =   "frmListado.frx":538E
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
         TabIndex        =   606
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
         TabIndex        =   605
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   93
         Left            =   720
         TabIndex        =   604
         Top             =   1440
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1320
         Picture         =   "frmListado.frx":5419
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
         TabIndex        =   603
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
         TabIndex        =   602
         Top             =   1920
         Width           =   585
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   94
         Left            =   1395
         Picture         =   "frmListado.frx":54A4
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   92
         Left            =   840
         TabIndex        =   601
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   93
         Left            =   1395
         Picture         =   "frmListado.frx":55A6
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
         TabIndex        =   592
         Top             =   360
         Width           =   5775
      End
   End
   Begin VB.Frame FrameRepxClien 
      Height          =   5415
      Left            =   240
      TabIndex        =   192
      Top             =   240
      Width           =   6795
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   3720
         TabIndex        =   351
         Top             =   3240
         Width           =   2415
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   4
            TabIndex        =   199
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
            TabIndex        =   353
            Top             =   420
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar equipos con más de:"
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
            TabIndex        =   352
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
         TabIndex        =   205
         Text            =   "Text5"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   1920
         TabIndex        =   194
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   36
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   204
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
         TabIndex        =   203
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
         TabIndex        =   196
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   195
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdAceptarRepxClien 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   200
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   5040
         TabIndex        =   201
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1920
         TabIndex        =   197
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1920
         TabIndex        =   198
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   33
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   202
         Text            =   "Text5"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   1920
         TabIndex        =   193
         Top             =   1320
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   1640
         Picture         =   "frmListado.frx":56A8
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1640
         Picture         =   "frmListado.frx":5733
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   7
         Left            =   1635
         Picture         =   "frmListado.frx":57BE
         ToolTipText     =   "Buscar cliente"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   37
         Left            =   1080
         TabIndex        =   215
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   9
         Left            =   1635
         Picture         =   "frmListado.frx":58C0
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   8
         Left            =   1635
         Picture         =   "frmListado.frx":59C2
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
         TabIndex        =   214
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   36
         Left            =   1080
         TabIndex        =   213
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   35
         Left            =   1080
         TabIndex        =   212
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
         TabIndex        =   211
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   1080
         TabIndex        =   210
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   30
         Left            =   1080
         TabIndex        =   209
         Top             =   3720
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Index           =   21
         Left            =   600
         TabIndex        =   208
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   6
         Left            =   1635
         Picture         =   "frmListado.frx":5AC4
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
         TabIndex        =   207
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
         TabIndex        =   206
         Top             =   1680
         Width           =   420
      End
   End
   Begin VB.Frame FrameFrecuencia 
      Height          =   3855
      Left            =   120
      TabIndex        =   503
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame13 
         Height          =   615
         Left            =   360
         TabIndex        =   653
         Top             =   2880
         Width           =   2655
         Begin VB.OptionButton OptFrecFicha 
            Caption         =   "Ficha"
            Height          =   255
            Left            =   120
            TabIndex        =   655
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptFrecResumen 
            Caption         =   "Resumen"
            Height          =   255
            Left            =   1320
            TabIndex        =   654
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
         TabIndex        =   514
         Text            =   "Text5"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   99
         Left            =   1320
         TabIndex        =   506
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   101
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   513
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
         TabIndex        =   512
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
         TabIndex        =   508
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
         TabIndex        =   507
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
         TabIndex        =   511
         Text            =   "Text5"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   98
         Left            =   1320
         TabIndex        =   505
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdFrecuencias 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   509
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   96
         Left            =   4800
         TabIndex        =   510
         Top             =   3120
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   77
         Left            =   1035
         Picture         =   "frmListado.frx":5BC6
         ToolTipText     =   "Buscar cliente"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   81
         Left            =   480
         TabIndex        =   520
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   79
         Left            =   1035
         Picture         =   "frmListado.frx":5CC8
         ToolTipText     =   "Buscar dir/dpto"
         Top             =   2400
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   78
         Left            =   1035
         Picture         =   "frmListado.frx":5DCA
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
         TabIndex        =   519
         Top             =   1800
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   80
         Left            =   480
         TabIndex        =   518
         Top             =   2400
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   79
         Left            =   480
         TabIndex        =   517
         Top             =   2040
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   76
         Left            =   1035
         Picture         =   "frmListado.frx":5ECC
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
         TabIndex        =   516
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
         TabIndex        =   515
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
         TabIndex        =   504
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame FrameListAvisosPtes 
      Height          =   4815
      Left            =   0
      TabIndex        =   415
      Top             =   0
      Width           =   6315
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   97
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   410
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   97
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   481
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
         TabIndex        =   409
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   96
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   478
         Text            =   "Text5"
         Top             =   3120
         Width           =   3375
      End
      Begin VB.ComboBox cboSituaAviso 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   411
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   82
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   405
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarAviPtes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   412
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   4800
         TabIndex        =   413
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   83
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   406
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   84
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   417
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
         TabIndex        =   407
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   85
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   416
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
         TabIndex        =   408
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
         TabIndex        =   482
         Top             =   3480
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   97
         Left            =   1440
         Picture         =   "frmListado.frx":5FCE
         ToolTipText     =   "Buscar tecnico"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   59
         Left            =   600
         TabIndex        =   480
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
         TabIndex        =   479
         Top             =   3120
         Width           =   450
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   96
         Left            =   1440
         Picture         =   "frmListado.frx":60D0
         ToolTipText     =   "Buscar tecnico"
         Top             =   3120
         Width           =   240
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
         Index           =   52
         Left            =   600
         TabIndex        =   425
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
         TabIndex        =   424
         Top             =   1080
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmListado.frx":61D2
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Avisos de avería pendientes"
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
         TabIndex        =   423
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
         TabIndex        =   422
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
         TabIndex        =   421
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3960
         Picture         =   "frmListado.frx":625D
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   84
         Left            =   1440
         Picture         =   "frmListado.frx":62E8
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
         TabIndex        =   420
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
         TabIndex        =   419
         Top             =   1920
         Width           =   450
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   85
         Left            =   1440
         Picture         =   "frmListado.frx":63EA
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
         TabIndex        =   418
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
         Picture         =   "frmListado.frx":64EC
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   740
         Width           =   585
      End
      Begin VB.CommandButton cmdSelTodos 
         Height          =   435
         Left            =   9720
         Picture         =   "frmListado.frx":6BD6
         Style           =   1  'Graphical
         TabIndex        =   190
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         Picture         =   "frmListado.frx":72C0
         ToolTipText     =   "Cliente"
         Top             =   4560
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   66
         Left            =   1200
         Picture         =   "frmListado.frx":73C2
         Top             =   4920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   70
         Left            =   600
         TabIndex        =   428
         Top             =   4560
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   69
         Left            =   600
         TabIndex        =   427
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
         TabIndex        =   426
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
         TabIndex        =   68
         Top             =   960
         Width           =   1755
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3315
         Picture         =   "frmListado.frx":74C4
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1155
         Picture         =   "frmListado.frx":754F
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   34
         Left            =   1155
         Picture         =   "frmListado.frx":75DA
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   33
         Left            =   1155
         Picture         =   "frmListado.frx":76DC
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
         TabIndex        =   67
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   66
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   65
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
         TabIndex        =   64
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   2760
         TabIndex        =   63
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
         Picture         =   "frmListado.frx":77DE
         ToolTipText     =   "Buscar familia"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   31
         Left            =   1155
         Picture         =   "frmListado.frx":78E0
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
         Picture         =   "frmListado.frx":79E2
         ToolTipText     =   "Buscar artículo"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   29
         Left            =   1155
         Picture         =   "frmListado.frx":7AE4
         ToolTipText     =   "Buscar artículo"
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
         Caption         =   "Informes Movimiento Artículos"
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
   Begin VB.Frame FrameMantenimientos 
      Height          =   7695
      Left            =   3480
      TabIndex        =   216
      Top             =   0
      Width           =   6735
      Begin VB.Frame FrameRuta 
         Height          =   1095
         Left            =   600
         TabIndex        =   690
         Top             =   4800
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   138
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   694
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
            TabIndex        =   229
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   137
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   691
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
            TabIndex        =   228
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   106
            Left            =   1080
            Picture         =   "frmListado.frx":7BE6
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
            TabIndex        =   695
            Top             =   720
            Width           =   420
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   105
            Left            =   1080
            Picture         =   "frmListado.frx":7CE8
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
            TabIndex        =   693
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
            TabIndex        =   692
            Top             =   285
            Width           =   450
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   360
         TabIndex        =   564
         Top             =   5880
         Width           =   6255
         Begin VB.CheckBox chkMante 
            Caption         =   "Copia remitente"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   235
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   109
            Left            =   1440
            TabIndex        =   236
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optMante 
            Caption         =   "Comercial"
            Height          =   195
            Index           =   1
            Left            =   4200
            TabIndex        =   234
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optMante 
            Caption         =   "Administracion"
            Height          =   195
            Index           =   0
            Left            =   2520
            TabIndex        =   233
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.CheckBox chkMante 
            Caption         =   "Enviar e-mail"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   232
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
            TabIndex        =   565
            Top             =   720
            Width           =   990
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   109
            Left            =   1155
            Picture         =   "frmListado.frx":7DEA
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Frame FrameManteActi 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         TabIndex        =   646
         Top             =   4800
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   129
            Left            =   1800
            TabIndex        =   240
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   127
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   648
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   127
            Left            =   1800
            TabIndex        =   226
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   128
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   647
            Text            =   "Text5"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   128
            Left            =   1800
            TabIndex        =   227
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Postal"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   652
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   102
            Left            =   960
            TabIndex        =   651
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
            TabIndex        =   650
            Top             =   0
            Width           =   795
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   99
            Left            =   1515
            Picture         =   "frmListado.frx":7E75
            ToolTipText     =   "Buscar actividad"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   101
            Left            =   960
            TabIndex        =   649
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   100
            Left            =   1515
            Picture         =   "frmListado.frx":7F77
            ToolTipText     =   "Buscar actividad"
            Top             =   600
            Width           =   240
         End
      End
      Begin VB.Frame FrameManteAnu 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         TabIndex        =   582
         Top             =   4800
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   116
            Left            =   5040
            TabIndex        =   239
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   115
            Left            =   2400
            TabIndex        =   238
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   114
            Left            =   1800
            TabIndex        =   231
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   114
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   586
            Text            =   "Text5"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   113
            Left            =   1800
            TabIndex        =   230
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   113
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   583
            Text            =   "Text5"
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   91
            Left            =   4200
            TabIndex        =   590
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   90
            Left            =   1560
            TabIndex        =   589
            Top             =   1080
            Width           =   465
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   14
            Left            =   4680
            Picture         =   "frmListado.frx":8079
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   13
            Left            =   2160
            Picture         =   "frmListado.frx":8104
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
            TabIndex        =   588
            Top             =   1080
            Width           =   915
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   92
            Left            =   1515
            Picture         =   "frmListado.frx":818F
            ToolTipText     =   "Buscar motivo baja"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   89
            Left            =   960
            TabIndex        =   587
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   91
            Left            =   1515
            Picture         =   "frmListado.frx":8291
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
            TabIndex        =   585
            Top             =   0
            Width           =   1005
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   88
            Left            =   960
            TabIndex        =   584
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   480
         TabIndex        =   561
         Top             =   5880
         Width           =   5895
         Begin VB.ComboBox cboTipoList 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   562
            Tag             =   "Tipo Facturación|N|N|||scaalb|tipofact||N|"
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
            TabIndex        =   563
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   45
         Left            =   1920
         TabIndex        =   218
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   45
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   258
         Text            =   "Text5"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   46
         Left            =   1920
         TabIndex        =   219
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   46
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   257
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
         TabIndex        =   256
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
         TabIndex        =   255
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
         TabIndex        =   244
         Text            =   "Text5"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   48
         Left            =   1920
         TabIndex        =   221
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   50
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   243
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
         TabIndex        =   242
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
         TabIndex        =   223
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   49
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   222
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton cmdAceptarMante 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   237
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5280
         TabIndex        =   241
         Top             =   7080
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   51
         Left            =   1920
         TabIndex        =   224
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   52
         Left            =   1920
         TabIndex        =   225
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   47
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   217
         Text            =   "Text5"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   47
         Left            =   1920
         TabIndex        =   220
         Top             =   2160
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   0
         Left            =   480
         TabIndex        =   363
         Top             =   4800
         Width           =   5415
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   1560
            TabIndex        =   366
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   54
            Left            =   3840
            TabIndex        =   365
            Top             =   360
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   9
            Left            =   3555
            Picture         =   "frmListado.frx":8393
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   8
            Left            =   1275
            Picture         =   "frmListado.frx":841E
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   195
            Index           =   44
            Left            =   720
            TabIndex        =   368
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   45
            Left            =   3000
            TabIndex        =   367
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
            TabIndex        =   364
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
         TabIndex        =   261
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
         TabIndex        =   260
         Top             =   960
         Width           =   420
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   41
         Left            =   1635
         Picture         =   "frmListado.frx":84A9
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   43
         Left            =   1080
         TabIndex        =   259
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   42
         Left            =   1635
         Picture         =   "frmListado.frx":85AB
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   44
         Left            =   1635
         Picture         =   "frmListado.frx":86AD
         ToolTipText     =   "Buscar cliente"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   42
         Left            =   1080
         TabIndex        =   254
         Top             =   2160
         Width           =   465
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   46
         Left            =   1635
         Picture         =   "frmListado.frx":87AF
         ToolTipText     =   "Buscar agente"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   45
         Left            =   1635
         Picture         =   "frmListado.frx":88B1
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
         TabIndex        =   253
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   41
         Left            =   1080
         TabIndex        =   252
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   40
         Left            =   1080
         TabIndex        =   251
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
         TabIndex        =   250
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   39
         Left            =   1080
         TabIndex        =   249
         Top             =   4080
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   38
         Left            =   1080
         TabIndex        =   248
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
         TabIndex        =   247
         Top             =   3840
         Width           =   1155
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   47
         Left            =   1635
         Picture         =   "frmListado.frx":89B3
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   48
         Left            =   1635
         Picture         =   "frmListado.frx":8AB5
         ToolTipText     =   "Buscar tipo contrato"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   43
         Left            =   1635
         Picture         =   "frmListado.frx":8BB7
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
         TabIndex        =   246
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
         TabIndex        =   245
         Top             =   2520
         Width           =   420
      End
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   11000
      Picture         =   "frmListado.frx":8CB9
      Tag             =   "-1"
      ToolTipText     =   "Buscar almacén"
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ====  MODIFICACIONES  ==========================================
' ====  [16/09/2009] LAURA : Añadir el frame "FrameInvArtComp" para sacar listado articulos con componentes
' ====  [17/09/2009] LAURA : Cambiar conexion CONTAB si FAZ
' ================================================================


Public OpcionListado As Integer

    '==== Listados de ALMACEN ====
    '=============================
    ' 1 .- Listados Marcas.
    ' 2 .- Listado de Almacenes Propios
    ' 3 .- Listado de Tipos de Unidad
    ' 4 .- Listado de Tipos de Artículos
    ' 5 .- Listado de Familias de artículos
    
    ' 6 .- Listado de Artículos
    ' 7 .- Informe de Traspaso de Almacenes
    ' 8 .- Informe de Movimientos de Almacen
    ' 9 .- Listado Busquedas de movimientos de Artículos
    '10 .-
    
    '11 .- Listado de Articulos con componentes ' ====  [16/09/2009] LAURA
    '12 .- Listado Toma de Inventario Articulos
    '13 .- Listado de Diferencias de Inventario Articulos
    '14 .- Actualizar Diferencias de Inventario (No IMPRIME INFORME)
    '15 .- Listado de Articulos Inactivos.
    
    '16 .- Listado Valoracion de Stocks Inventariados
    '17 .- Listado Valoración Stocks
    '18 .- Informe Stocks Maximos y Minimos
    '19 .- Informe de Stocks a una fecha
    
    '110 .- Listado de Ubicaciones
    
    
    
    
    '==== Listados de FACTURACION ====
    '=================================
    '20 .- Listado de Actividades de Clientes
    '21 .- Listado de Zonas de Clientes
    '22 .- Listado de Rutas de Asistencia
    '23 .- Listado de Formas de Envío
    '24 .- Listado de Tarifas Ventas
    '25 .-
    
    '26 .-
    '27 .- Listado de Situaciones Especiales
    '28 .- Informe de Tarifas de Articulos
    '29 .- Informe de Promociones de Tarifas
    '30 .- Informe de Precios Especiales
    
    '31 .- Informe de Ofertas
    '32 .- Informe de Recordatorio de Ofertas
    '33 .- Informe de Valoración de Ofertas
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
    '247 .- Corrección de errores y acutalizacion de tarifas
    
    
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
    '63 .- Listado Reparaciones por Día
    '64 .- Listado Reparaciones por Cliente
    '65 .- Listado motivos baja equipos
    
    '406 .- Listado Frecuencia de reparaciones
    '407 .- Sustitución Nº de Serie
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
    '74 .- Prefacturación Mantenimientos
    '75 .- Facturación de Mantenimientos
    '76 .- IGUAL QUE EL 70 pero en ANULADOS
        
        
        
    '77 .- Informe teórico de mantenimientos
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
    
    '92 .- Informe de Gastos Técnicos
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
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

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
Private WithEvents frmMtoFamilia As frmBasico2 'frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmMtoProveedor As frmBasico2 '%=%=frmComProveedores
Attribute frmMtoProveedor.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmBasico2
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Attribute frmMtoCartasOfe.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmBasico2 'frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoClientes As frmBasico2
Attribute frmMtoClientes.VB_VarHelpID = -1
Private WithEvents frmMtoMotivos As frmRepMotivosPend
Attribute frmMtoMotivos.VB_VarHelpID = -1
Private WithEvents frmMtoAgentes As frmBasico2 '%=%=frmFacAgentesCom
Attribute frmMtoAgentes.VB_VarHelpID = -1
Private WithEvents frmMtoTiposCon As frmManTiposContrato
Attribute frmMtoTiposCon.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmIncidencias
Attribute frmMtoIncid.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmPWD As frmMensajes
Attribute frmPWD.VB_VarHelpID = -1


'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------





'Para ademas de insertarlas en la conta, que las contabilice (pase a hsaldos)
'es decir, en el momento que inserta en cabfact tb insertaremos en hlinapu, hacabapu, hsaldos y hsaldosanal (si procede)










Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
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
            
            chkImpEtiq(3).visible = False
        Else
            chkImpEtiq(1).Caption = "P.V.P."
            chkImpEtiq(1).Value = 0
            chkImpEtiq(3).visible = True
            chkImpEtiq(3).Value = 0
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
            Codigo = "{sartic.codartic}"
        Else
            cadNomRPT = "rAlmArtCompVer.rpt"
            conSubRPT = False
            Codigo = "{sarti1.codarti1}"
        End If
        Screen.MousePointer = vbHourglass
        DoEvents
        
        'Añadir el parametro de Empresa
        cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = 1
        
        
        
        'Septiembre 2014. Herbelca
        If Me.chkCompo.Value = 1 Then
            cadAux = ""
          
            For bytPrecio = 1 To 3
                
                cadParam = cadParam & "Almacen" & bytPrecio & "=" & Val(txtCodigo(144 + bytPrecio).Text) & "|"
                numParam = numParam + 1
                'empieza en el 145
                If Trim(txtCodigo(144 + bytPrecio).Text) <> "" Then
                    cadAux = cadAux & "      " & Trim(txtCodigo(144 + bytPrecio)) & " " & txtNombre(144 + bytPrecio).Text
                End If
            Next
            If cadAux = "" Then cadAux = "ERROR ALMACENES"
            cadParam = cadParam & "pAlmacenes=""Almacenes:  " & Trim(cadAux) & """|"
            numParam = numParam + 1
        End If
        If Trim(txtCodigo(125).Text) <> "" Or Trim(txtCodigo(126).Text) <> "" Then
            cadFormula = CadenaDesdeHasta(txtCodigo(125).Text, txtCodigo(126).Text, Codigo, "T")
            
            If cadFormula <> "" Then
                If cadFormula = "Error" Then Exit Sub
                cadAux = ""
                If txtCodigo(125).Text <> "" Then cadAux = "Desde: " & txtCodigo(125).Text & " " & txtNombre(125).Text
                If txtCodigo(126).Text <> "" Then
                    If cadAux <> "" Then cadAux = cadAux & "  -  "
                    cadAux = cadAux & " Hasta: " & txtCodigo(126).Text & " " & txtNombre(126).Text
                End If
                cadParam = cadParam & "pDesde=""" & cadAux & """|"
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
        cadParam = cadParam & "pPrecio=" & bytPrecio & "|"
        numParam = numParam + 1
    ' ====
   
   
    Case 1 'Frame Listados
        If Me.Optcodigo.Value = True Then
            cadAux = Orden1
        Else
            cadAux = Orden2
        End If
        cadParam = "|pOrden=" & cadAux & "|"
        numParam = 1
        
        'Añadir el parametro de Empresa
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        If Trim(txtCodigo(1).Text) <> "" Or Trim(txtCodigo(2).Text) <> "" Then
            'Cadena para seleccion Desde y Hasta
            If OpcionListado = 4 Or OpcionListado = 110 Then
                '4: Listado Tipos de Articulos, 110: List. Ubicaciones
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, Codigo, "T")
            Else
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, Codigo, "N")
            End If
            
            If cadFormula <> "" Then
                If cadFormula = "Error" Then Exit Sub
                cadAux = ""
                If txtCodigo(1).Text <> "" Then cadAux = "Desde: " & txtCodigo(1).Text & " " & txtNombre(1).Text
                If txtCodigo(2).Text <> "" Then
                    If cadAux <> "" Then cadAux = cadAux & "  -  "
                    cadAux = cadAux & " Hasta: " & txtCodigo(2).Text & " " & txtNombre(2).Text
                End If
                cadParam = cadParam & "pDesde=""" & cadAux & """|"
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
        
        cadParam = "|"
        If Not PonerParamEmpresa(cadParam, numParam) Then Exit Sub
        If PonerParamRPT2(indRPT, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
            'Cadena para seleccion Desde y Hasta DOCUMENTO
            '----------------------------------------------
            If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
                If Not PonerDesdeHasta(Codigo, "N", 3, 4, "") Then Exit Sub
            End If
        
            If Not HayRegParaInforme(cadAux, cadSelect) Then Exit Sub
        End If
                       
                   
                   
    '========= Frame Listado Movimiento de Artículos ========================
    Case 3 'Frame Listado Movimiento de Artículos
        'Nombre fichero .rpt a Imprimir
        
        indRPT = 75
        If Not PonerParamRPT2(indRPT, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNomRPT = "rAlmMovim.rpt"
        
        
        
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
                MsgBox "La Actualización de Inventario se ha realizado correctamente.", vbInformation
            End If
        Else
            MsgBox "El campo Trabajador debe tener valor", vbInformation
            PonerFoco txtCodigo(21)
            Exit Sub
        End If
        
   Else 'Listados
   
   

   
   
   
'        If OpcionListado = 19 Then cadFormula = ""
        If OpcionListado = 19 Then cadFormula = "({tmpstockfec.codusu} =" & vUsu.Codigo & ")"
        
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
    
    

    
    
    If MsgBox("¿Impresión correcta para Actualizar Inventario?", vbQuestion + vbYesNo) = vbYes Then
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
        cadParam = ""
        For Opcion = 0 To 3
            If Me.chkSitaucionArticulo2(Opcion).Value = 1 Then cadParam = cadParam & "O"
        Next
        If cadParam = "" Then
            MsgBox "Seleccione la situacion del articulo", vbExclamation
            Exit Sub
        End If
        Opcion = 0
        

        If Me.chkImpEtiq(0).Value = 0 And Me.chkImpEtiq(1).Value = 0 And Me.chkImpEtiq(3).Value = 1 Then
            'MsgBox "Debe marcar la opcion PVP para que salga el precio mínimo", vbExclamation
            'Exit Sub
            chkImpEtiq(1).Value = 1
        End If
    End Select
    
    '===================================================
    '============ PARAMETROS ===========================
    cadParam = "|"
    'Empresa
    cadParam = cadParam & "pEmpresa=""" & vParam.NombreEmpresa & """|"
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
        
        cadParam = cadParam & devuelve & """|"
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
        cadParam = cadParam & devuelve
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
    
    If chkImpEtiq(3).Value = 1 Then
        'PREcio minimo
        campo = "{sartic.preciominvta} > 0 "
        AnyadirAFormula cadFormula, campo
        AnyadirAFormula cadSelect, campo
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
        
            'Añadir a la formula el chk
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
            'Para precio minimo NO lanzamos
            'If chkImpEtiq(3).Value = 0 Then
                cadTitulo = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.Codigo
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
           ' End If
        
            cadTitulo = "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo
            conn.Execute cadTitulo
    
            'Añadire los tipos de IVA a esta tabla para los posibles links
            cadTitulo = "INSERT INTO tmpinformes(codusu,codigo1)  select " & vUsu.Codigo & ",codigiva from tmpnseries,sartic"
            cadTitulo = cadTitulo & " WHERE codusu = " & vUsu.Codigo & " AND tmpnseries.codartic=sartic.codartic"
            cadTitulo = cadTitulo & " GROUP BY codigiva"
            conn.Execute cadTitulo
            
            Espera 0.2
             'Abrimos los IVAS en conta
            Set miRsAux = New ADODB.Recordset
            cadTitulo = "Select codigo1 from tmpinformes WHERE codusu = " & vUsu.Codigo
            miRsAux.Open cadTitulo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                cadTitulo = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", miRsAux!Codigo1)
                cadTitulo = TransformaComasPuntos(cadTitulo)
                cadTitulo = "UPDATE tmpinformes SET porcen1= " & cadTitulo
                cadTitulo = cadTitulo & " WHERE codusu = " & vUsu.Codigo & " AND codigo1 = " & miRsAux!Codigo1
                conn.Execute cadTitulo
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            
            
            If vParamAplic.NumeroInstalacion = vbTaxco Then
                cadTitulo = "Select tmpnseries.*,ubialmac from tmpnseries left join salmac ON tmpnseries.codartic=salmac.codartic "
                cadTitulo = cadTitulo & " and codalmac=1 "
                cadTitulo = cadTitulo & " WHERE codusu = " & vUsu.Codigo
                
                miRsAux.Open cadTitulo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not miRsAux.EOF
                    cadFormula = miRsAux!nummante
                    cadTitulo = DBLet(miRsAux!ubialmac, "T")
                    Codigo = ""
                    
                    'numero de etiquetas
                    
                    If Val(cadFormula) > 0 Then
                        indCodigo = Val(cadFormula) - 1
                    Else
                        indCodigo = miRsAux!cantidad - 1
                    End If
                    While indCodigo <> 0
                        'tmpnseries(codusu,codartic,numserie,numlinealb,numlinea)
                        Codigo = Codigo & ", (" & vUsu.Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!numSerie, "T")
                        Codigo = Codigo & "," & indCodigo & ",1)"
                        indCodigo = indCodigo - 1
                    Wend
                    If Codigo <> "" Then
                        Codigo = Mid(Codigo, 2)
                        Codigo = "INSERT INTO tmpnseries(codusu,codartic,numserie,numlinealb,numlinea) VALUES " & Codigo
                        ejecutar Codigo, False
                        Espera 0.15
                    End If
                    
                    
                    If cadTitulo <> "" Then
                        If Asc(cadTitulo) = 13 Then cadTitulo = ""
                    End If
                    If cadTitulo <> "" Then
                    
                        cadTitulo = "UPDATE tmpnseries SET nummante= " & DBSet(miRsAux!ubialmac, "T")
                        cadTitulo = cadTitulo & " WHERE codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(miRsAux!codArtic, "T")
                        conn.Execute cadTitulo
                    End If
                    
                    
                    
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
            
            End If
            
            Set miRsAux = Nothing
            
    
        
        
            'Es para imprimir etiquetas
            'rAlmArticulosPVP.rpt
            If chkImpEtiq(0).Value = 0 Then
                'OK. PVPV
                If Not PonerParamRPT2(IIf(chkImpEtiq(3).Value = 1, 87, 80), cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNomRPT = "rAlmArticulosPVP.rpt"
                cadTitulo = IIf(chkImpEtiq(3).Value = 1, "Articulos PVP IVA con precio mínimo", "Articulos PVP IVA")
            Else
                If Not PonerParamRPT2(23, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNomRPT = "rEtiArticulo.rpt"
                cadTitulo = "Etiquetas articulos gral."
                cadParam = "|pImprimeBarras=""1""|numerodecimales=2|"
                numParam = 2
            End If
            
            cadFrom = " tmpnseries   "
            cadSelect = "{tmpnseries.codusu} =" & vUsu.Codigo
            cadFormula = "{tmpnseries.codusu} =" & vUsu.Codigo
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
                            cadParam = cadParam & campo & "|"
                            numParam = numParam + 1
                            
                            campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                            cadParam = cadParam & campo & "|"
                            numParam = numParam + 1
                        Case 2 'El Group3 es el Proveedor
                            campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                            cadParam = cadParam & campo & "|"
                            numParam = numParam + 1
                            
                            campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                            cadParam = cadParam & campo & "|"
                            numParam = numParam + 1
                        Case 3, 0 'El Group4 es el Proveedor
                                  '0 'El Group1 es el Proveedor
                            campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                            cadParam = cadParam & campo & "|"
                            numParam = numParam + 1
                            
                            campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """"
                            cadParam = cadParam & campo & "|"
                            numParam = numParam + 1
                            
                            If Opcion = 0 Then
                                campo = "pTitulo3=""" & ListView2.ListItems(4).Text & """"
                                cadParam = cadParam & campo & "|"
                                numParam = numParam + 1
                            End If
                    End Select
                   
                    
                    cadTitulo = "Listado de Artículos"
                    campo = "pOrden=" & Opcion
                    cadParam = cadParam & campo & "|"
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
        cadParam = cadParam & "codusu= " & vUsu.Codigo & "|"
        numParam = numParam + 1
    
        'Añadir el Parametro de Stocks Maximos o Minimos
        If Me.optStockMax.Value = True Then
            campo = "0"
        Else
            If optPuntoPedido.Value Then
                campo = "2"
            Else
                campo = "1"
            End If
        End If
        cadParam = cadParam & "pStockMax=" & campo & "|"
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
        
        
    indCodigo = 0
    If OpcionListado = 6 Then
        If chkImpEtiq(0).Value = 1 Then
            OpcionListado = 513   'para que imprmia etiquetas directamente
            indCodigo = 1 'Indicamos que hemos cambiado
        End If
    End If
    
    LlamarImprimir False
    
    If indCodigo = 1 Then OpcionListado = 6
    
    
End Sub


Private Sub cmdAceptarAviPtes_Click()
'409: Listado Avisos averias pendientes
Dim tabla As String
Dim campo As String, Cad As String

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    tabla = "scaavi"
    cadTitulo = "Listado Avisos de averías Pendientes"
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
    Cad = "pDHSitua=""Situación: "
    If Me.cboSituaAviso.ListIndex = -1 Or Me.cboSituaAviso.ListIndex = 0 Then
        Cad = Cad & "Todas" & """|"
    Else
        Cad = Cad & Me.cboSituaAviso.List(Me.cboSituaAviso.ListIndex) & """|"
        campo = "{" & tabla & ".situacio}=" & Me.cboSituaAviso.ListIndex - 1
        
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    End If
    cadParam = cadParam & Cad
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
        Cad = "pDHTecni=""Técnico: "
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
        lblDtoAct.visible = False
        Exit Sub
    End If
    
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
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
    
    
    
    If OpcionListado = 54 Then
        If txtCodigo(158).Text <> "" Or txtCodigo(159).Text <> "" Then
            campo = "{sclien.codagent}"
            'If OpcionListado = 309 Then campo = "{sartic.codfamia}"
            Cad = "     Agente: "
            If Not PonerDesdeHasta(campo, "N", 158, 159, Cad) Then Exit Sub
            Orden1 = Trim(Orden1 & Cad)
        End If
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
        cadParam = cadParam & Cad
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
        cadParam = cadParam & Cad & "|"
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
            cadParam = cadParam & Cad & " ""|"
        End If
    End If
    Cad = ""
    
    '==============================================================
    If OpcionListado = 54 Then
        
        'En herbelca NO dejo continuar si no pone algun desde hasta
        'If vParamAplic.AlmacenB > 90 Then
        If vParamAplic.NumeroInstalacion = vbHerbelca Then
            If cadSelect = "" Then
                MsgBox "Escriba algun criterio de busqueda", vbExclamation
                Exit Sub
            End If
        End If
        
        If Me.optFrDto(0).Value Then
            cadNomRPT = "rFacDtosFM.rpt"
            campo = "sdtofm.codclien"
        ElseIf Me.optFrDto(1).Value Then
            cadNomRPT = "rFacDtosFMAct.rpt"
            campo = "sdtofm.codactiv"
        ElseIf Me.optFrDto(4).Value Then
            'Nuevo proveedor
            cadNomRPT = "rFacDtosFMprov.rpt"
            campo = "sdtofm.codclien"
        
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
            cadFormula = cadFormula & " ({" & campo & "}>0)"
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
        
        
        
        tabla = tabla & " INNER JOIN sclien ON sdtofm.codclien=sclien.codclien"

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
    
    If OpcionListado = 54 Then
        If Me.chkVarios(6).Value = 1 Then
            'OCULTAR PROVEEDOR
            cadParam = cadParam & "pOcultaProv=1|"
            numParam = numParam + 1
        End If
    End If
    
    If OpcionListado = 309 Then
        If Me.chkVarios(1).Value Then
            'Cargaremos tmpInformes con el sdtopm aplicado sobre el precio del articulo
            Orden1 = tabla
            HazCalculoPrecioNetoProve
            
            cadTitulo = "Precio neto proveedor"
            cadFormula = "({tmpinformes.codusu} = " & vUsu.Codigo & ")"
            cadNomRPT = "rComPreciosNeto.rpt"
        End If
    Else
        cadTitulo = ""
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
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    If chkMargen(0).Value = 0 Then
        'Para los que detalla fracion NO envio pDetalla
        cadParam = cadParam & "Detalla= " & chkMargen(1).Value & "|"
        numParam = numParam + 1
    End If
    
    ' Septiembre 2015
    ' Sin marcar (como estaba), marcado sobre las ventas(Herbelca)... y creo que es lo mas logico
    tabla = 1
    If chkMargen(4).Value = 1 Then tabla = "0"   'paremtro en rpt: MargenSobreCoste
    'Para los que detalla fracion NO envio pDetalla
    cadParam = cadParam & "MargenSobreCoste= " & tabla & "|"
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
    desPrecio = "Valoración coste: "
    If Me.optPrecioMP2.Value Then
        opcPrecio = "{slifac.preciomp}" 'precio medio ponderado
        desPrecio = desPrecio & "Precio medio ponderado"
    ElseIf Me.optPrecioUC2.Value Then
        opcPrecio = "{slifac.preciouc}" 'precio ultima compra
        desPrecio = desPrecio & "Precio última compra"
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
    
    cadParam = cadParam & "pCampo=" & opcPrecio & "|"
    'Le pong las fechas(si es k las han puesto)
    desPrecio = Trim(desPrecio & "          " & param)
    If chkMargen(4).Value = 1 Then desPrecio = desPrecio & "[% Sobre vta]"
    cadParam = cadParam & "pDesCampo=""" & desPrecio & """|"
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
    
    'Cadena para seleccion D/H artículo
    '--------------------------------------------
    If txtCodigo(90).Text <> "" Or txtCodigo(91).Text <> "" Then
        campo = "{slifac.codartic}"
        param = "pDHArticulo=""Artículo: "
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
    cadParam = "|"
    
    
    If chkMante(4).Value = 1 Then
        'Enero 2010
        'Informe completo
        indRPT = 38
    Else
        indRPT = 13
    End If
    If Not PonerParamRPT2(indRPT, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
    'Ejercicio
    cadParam = cadParam & "pEjercicio=""" & txtCodigo(61).Text & """|"
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
    
    'Cadena para seleccion D/H Nº Mantenimiento
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
    cadParam = cadParam & "ImprimeArticulo=" & Abs(Me.chkMante(1).Value) & "|"
    numParam = numParam + 1
    LlamarImprimir True
End Sub


Private Sub cmdAceptarMante_Click()
'Listado de Mantenimientos
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim Codigo  As String

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
        Codigo = "scaman"
        If OpcionListado = 76 Then
            'ANULADOS    rManListManImporteAnu.rpt
            cadTitulo = cadTitulo & " Anulados"
            Codigo = Codigo & "a"
            cadNomRPT = cadNomRPT & "Anu"
        End If
        cadNomRPT = cadNomRPT & ".RPT"
    Case 71
        cadNomRPT = "rManListRevisiones.rpt"
        Codigo = "scaman"
        cadTitulo = "Informe Revisiones"
    Case 78
    
        'PEqueña comprobacion.
        'Fecha obligatoria
        If txtCodigo(109).Text = "" Then
            MsgBox "Debe indicar la fecha", vbExclamation
            Exit Sub
        End If
    
    
        If Not PonerParamRPT2(21, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
        Codigo = "scaman"
    Case 79
        If Not PonerParamRPT2(45, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
        Codigo = "scaman"
    End Select
    cadFrom = "(" & Codigo & " INNER JOIN sclien ON " & Codigo & ".codclien=sclien.codclien) "
      
      
      
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
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
        campo = "{" & Codigo & ".codclien}"
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
        campo = "{" & Codigo & ".codtipco}"
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
        cadParam = cadParam & "pDHTipoCon=""" & Orden1 & """|"
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
            campo = "{" & Codigo & ".codincid}"
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
                cadParam = cadParam & devuelve & """|"
                numParam = numParam + 1
            End If
        End If
        
        If txtCodigo(54).Text <> "" Then
            devuelve = "Date(" & Year(txtCodigo(54).Text) & "," & Month(txtCodigo(54).Text) & "," & Day(txtCodigo(54).Text) & ")"
            If devuelve <> "" Then
                devuelve = "pHFecha=" & devuelve & "|"
                cadParam = cadParam & devuelve & """|"
                numParam = numParam + 1
            End If
        End If
    End If
        
        

        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    'cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    'Esto lo hago siempre para gene temporales
    conn.Execute "DELETE from tmpnlotes where codusu =" & vUsu.Codigo
    
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
        cadFormula = "({tmpinformes.codusu} = " & vUsu.Codigo & ")"
        conSubRPT = True
    End If
    devuelve = ""
    If OpcionListado = 78 Then
        'Añado la fecha
        cadParam = cadParam & "|FechaImp=""" & txtCodigo(109).Text & """|"
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
                devuelve = devuelve & "    - " & miRsAux!NomClien & vbCrLf
            Else
                'INSERTAMOS
                NumRegElim = NumRegElim + 1
                Codigo = "insert into tmpnlotes (codusu,numalbar,fechaalb,codprove,numlinea,codartic) values ("
                Codigo = Codigo & vUsu.Codigo & ",1,'" & Format(txtCodigo(109).Text, FormatoFecha) & "'," & miRsAux!codClien & ","
                Codigo = Codigo & NumRegElim & ",'" & miRsAux!nummante & "')"
                conn.Execute Codigo
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
            devuelve = "Clientes sin mail: " & vbCrLf & devuelve & "¿Continuar?"
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
            cadSelect = cadSelect & " as email,tmpnlotes.* from tmpnlotes,sclien where codusu = " & vUsu.Codigo & " and codclien=codprove"
        
            
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
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(37).Text <> "" Or txtCodigo(38).Text <> "" Then
        campo = Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        Cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 37, 38, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Del DEPARTAMENTO
    '--------------------------------------------
    If txtCodigo(39).Text <> "" Or txtCodigo(40).Text <> "" Then
        campo = Codigo & ".coddirec}"
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
    
    'Cadena para seleccion Nº CONTRATO
    '--------------------------------------------
    If txtCodigo(41).Text <> "" Or txtCodigo(42).Text <> "" Then
        campo = Codigo & ".nummante}"
        'Parametro Desde/Hasta Cliente
        Cad = "pDHContrato=""Nº Mantenimiento: "
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
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
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
            devuelve = "pDHDpto=""Dirección: "
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
        
        'Nº de Reparaciones, Añadirlo como parametro
        '----------------------------------------------
        cadParam = cadParam & "pNumVeces=" & txtCodigo(0).Text & "|"
        numParam = numParam + 1
        
        On Error GoTo EFrecu
        'Insertar en la tabla temporal tmpInformes el total de reparaciones para cada
        'codartic, numserie para el criterio de seleccion introducid
        devuelve = "INSERT INTO tmpinformes(codusu,nombre1,nombre2,campo1) "
        devuelve = devuelve & "SELECT " & vUsu.Codigo & ", codartic,numserie,count(numserie) as campo1 from schrep "
        devuelve = devuelve & " WHERE " & cadSelect
        devuelve = devuelve & " group by codartic,numserie"
        conn.Execute devuelve
        
        'Eliminamos de la tabla aquellos registros que no superen el nº de reparaciones introducido
        devuelve = "DELETE FROM tmpinformes where codusu=" & vUsu.Codigo & " and campo1<=" & txtCodigo(0).Text
        conn.Execute devuelve
        
        'Volver a comprobar que hay registro a mostrar para ello miramos en la
        'tabla tmpInformes que supere el nº de reparaciones a mostrar
        cadSelect = "codusu=" & vUsu.Codigo
        If Not HayRegParaInforme("tmpinformes", cadSelect) Then
            BorrarTempInformes
            Exit Sub
        End If
    End If
    
    LlamarImprimir False
    
    'Eliminar de la tabla temporal
    If OpcionListado = 406 Then BorrarTempInformes
    
EFrecu:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo nº de reparaciones.", Err.Description
End Sub


Private Sub cmdAceptarRepxDia_Click()
'Reparaciones por Dia
Dim devuelve As String
Dim param As String
Dim TotalMante As Integer
Dim RS As ADODB.Recordset
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
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    Select Case OpcionListado
        Case 63
            'Codigo = "{scarep.fecentre}"
            Codigo = "{scarep.fecrepar}"   '09/12/2010  Estan cambiadas en el form
            param = "pDHFecha=""Fecha Rep.: "
            NomTabla = "scarep"
            cadNomRPT = "rRepReparxDia.rpt"
            conSubRPT = True
            cadTitulo = "Reparaciones por día"
        Case 73
            'Añadir el parametro total Mantenim. si estamos en Informe de Altas
            devuelve = "SELECT DISTINCT COUNT(*) FROM scaman "
            Set RS = New ADODB.Recordset
            RS.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                TotalMante = RS.Fields(0).Value
                cadParam = cadParam & "pTotalMante=" & TotalMante & "|"
                numParam = numParam + 1
            End If
            RS.Close
            Set RS = Nothing
            
            'Añadir el Total Mantenim. del Periodo anterior
            fecha1 = Day(txtCodigo(31).Text) & "/" & Month(txtCodigo(31).Text) & "/" & Year(txtCodigo(31).Text) - 1
            fecha2 = Day(txtCodigo(32).Text) & "/" & Month(txtCodigo(32).Text) & "/" & Year(txtCodigo(32).Text) - 1
            Codigo = "scaman.fechaini"
            devuelve = CadenaDesdeHastaBD(fecha1, fecha2, Codigo, "F")
            If devuelve <> "" And devuelve <> "Error" Then
                devuelve = "SELECT DISTINCT COUNT(*) FROM scaman WHERE " & devuelve
                Set RS = New ADODB.Recordset
                RS.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    TotalMante = RS.Fields(0).Value
                    cadParam = cadParam & "pTotalAnte=" & TotalMante & "|"
                    numParam = numParam + 1
                End If
                RS.Close
                Set RS = Nothing
            End If
            
            '================= FORMULA =========================
            Codigo = "{scaman.fechaini}"
            param = "pDHFecha=""Fecha: "
            NomTabla = "scaman"
            cadNomRPT = "rManListAltas.rpt"
            cadTitulo = "Informe Altas Mantenimientos"
        
        Case 223
            param = ""
            If Me.OptClientes Then
                Codigo = "{scafac.fecfactu}"
                NomTabla = "scafac"
            Else
                Codigo = "{scafpc.fecrecep}"
                NomTabla = "scafpc"
            End If
    End Select
   
        
    '===================================================
    '================= FORMULA =========================
    
    '== Cadena para seleccion Desde y Hasta FECHA ==
    If OpcionListado = 223 Then
        'El usuario de B solo puede contabilzar facturas de B
        If Val(vUsu.AlmacenPorDefecto2) = vParamAplic.AlmacenB Then
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
        'contabilidad par ello mirar en la BD de la Conta los parámetros
        If Not ComprobarFechasConta(31) Then Exit Sub
        If Not ComprobarFechasConta(32) Then Exit Sub
    End If
    
    devuelve = CadenaDesdeHasta(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F", "Fecha Factura")
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Parametro D/H Fecha
    If devuelve <> "" And param <> "" Then
        cadParam = cadParam & AnyadirParametroDH(param, 31, 32) & """|"
        numParam = numParam + 1
    End If
    
    
    '## LAURA 20/06/2008
    '## Añadir frame de selec. factuar en contabilizar
    '- cadena para select en BDatos
    cadSelect = CadenaDesdeHastaBD(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F")
    
    
    
    
    'DAVID###
    'Rep x dia. Añadimos desde hasta cliente
    If OpcionListado = 63 Then
        devuelve = CadenaDesdeHasta(txtCodigo(132).Text, txtCodigo(133).Text, "{scarep.codclien}", "N", "Cliente")
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Parametro D/H Fecha
        If devuelve <> "" Then
            devuelve = ""
            cadParam = cadParam & AnyadirParametroDH("pDHCliente=""Cliente: ", 132, 133) & """|"
            numParam = numParam + 1
        End If
    
    End If
    
    '== Cadena para seleccion Desde y Hasta NºFactura ==
    If OpcionListado = 223 Then
        '- comprobar: si nº factura tienen valor tipoMov tb
        If txtCodigo(121).Text <> "" Or txtCodigo(122).Text <> "" Then
            If Me.cboTipMov.ListIndex = -1 Then
                MsgBox "Debe seleccionar el tipo de movimiento si quiere contabilizar Desde/Hasta Nº Factura.", vbInformation
                Exit Sub
            End If
            
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) = "" Then
                MsgBox "Debe seleccionar el tipo de movimiento si quiere contabilizar Desde/Hasta Nº Factura.", vbInformation
                Exit Sub
            End If
            
            '- añadir desde/hasta factura a cadena seleccion registros
            Codigo = "{scafac.numfactu}"
            devuelve = CadenaDesdeHasta(txtCodigo(121).Text, txtCodigo(122).Text, Codigo, "N", "Nº Factura")
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            'Parametro D/H nº factura
            If devuelve <> "" And param <> "" Then
                cadParam = cadParam & AnyadirParametroDH(param, 31, 32) & """|"
                numParam = numParam + 1
            End If
            ' añadir a la formula de bd
            devuelve = CadenaDesdeHastaBD(txtCodigo(121).Text, txtCodigo(122).Text, Codigo, "N")
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        End If
    
                
        '- añadir tipo movimiento a cadena seleccion
        If Me.cboTipMov.ListIndex >= 0 Then
            If Trim(Me.cboTipMov.List(Me.cboTipMov.ListIndex)) <> "" Then
                Codigo = "{scafac.codtipom}"
                devuelve = Mid(Me.cboTipMov.List(Me.cboTipMov.ListIndex), 1, 3)
                devuelve = Codigo & "=" & DBSet(devuelve, "T")
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
                     If Val(vUsu.AlmacenPorDefecto2) <> vParamAplic.AlmacenB Then cadSelect = cadSelect & " AND scafac.codtipom <> 'FAZ'"
                End If
            End If
        Else
            'CONTABILZIACION DE LOS TICKETS AGRUPADOS
            'Añado el tipom al cad select
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
'Sustitucion de un Nº de Serie que este en garantía por otro nº de serie.
Dim SQL As String
Dim RS As ADODB.Recordset

    txtCodigo(81).Text = Trim(txtCodigo(81).Text)
    
    If txtCodigo(81).Text <> "" Then
        'Comprobar que el nuevo nº de serie no existe ya
        SQL = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", txtCodigo(81).Text, "T", , "codartic", Me.CadTag, "T")
        If SQL <> "" Then
            MsgBox "Ya existe ese Nº de serie.", vbExclamation
            Exit Sub
        End If
        
        On Error GoTo ESustNSerie
        conn.BeginTrans
        
        'Insertar un registro con ese nº de serie y todos los valores que tenga el
        'num serie que sustituye
        SQL = "SELECT codartic,codtipar,codclien,coddirec,tieneman,nummante,ultrepar,fingaran,codtipom,numfactu,fechavta,numalbar,numline1,codprove,numalbpr,fechacom,numline2 FROM sserie "
        SQL = SQL & " WHERE numserie=" & DBSet(NumCod, "T") & " AND codartic=" & DBSet(CadTag, "T")
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If Not RS.EOF Then
            SQL = "(" & DBSet(txtCodigo(81).Text, "T") & ", " & DBSet(RS!codArtic, "T", "N") & "," & DBSet(RS!codTipar, "T", "N") & ","
            SQL = SQL & DBSet(RS!codClien, "N", "S") & "," & DBSet(RS!CodDirec, "N", "S") & "," & DBSet(RS!TieneMan, "N", "S") & ","
            SQL = SQL & DBSet(RS!nummante, "T", "S") & "," & DBSet(RS!ultrepar, "F", "S") & "," & DBSet(RS!fingaran, "F", "S") & ","
            SQL = SQL & DBSet(RS!codtipom, "T", "S") & "," & DBSet(RS!Numfactu, "N", "S") & "," & DBSet(RS!FechaVta, "F", "S") & ","
            SQL = SQL & DBSet(RS!Numalbar, "N", "S") & "," & DBSet(RS!numline1, "N", "S") & "," & DBSet(RS!Codprove, "N", "S") & ","
            SQL = SQL & DBSet(RS!numalbPr, "T", "S") & "," & DBSet(RS!fechacom, "F", "S") & "," & DBSet(RS!numline2, "N", "S") & ")"
        End If
        RS.Close
        Set RS = Nothing
        
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
        MsgBox "Debe introducir el Nº Serie por el que se sustituye.", vbInformation
        Exit Sub
    End If

ESustNSerie:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Sustitución Nº Serie.", Err.Description
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
        cadFrom = Codigo & " INNER JOIN sartic ON " & Codigo & ".codartic=sartic.codartic "
    Else
        cadFrom = Codigo
    End If
    
    
    If (OpcionListado = 30) Then cadFrom = cadFrom & " INNER JOIN sclien ON " & Codigo & ".codclien=sclien.codclien "
    
    'seleccionar solo los que tienen margen con error
    If OpcionListado = 245 Then
        If Me.chkMostrarErrores Then
            AnyadirAFormula cadSelect, " sartic.preciove <> sartic.preciouc + round((sartic.preciouc * if(isnull(sartic.margecom),0,sartic.margecom))/100,4)"
            AnyadirAFormula cadFormula, " {sartic.preciove} <> {sartic.preciouc} + round(({sartic.preciouc} * iif(IsNull({sartic.margecom}),0,{sartic.margecom}))/100,4)"
        End If
    End If
    
    
    If OpcionListado = 30 Then
        If Me.chkVarios(5).Value = 1 Then
            'OCULTAR PROVEEDOR
            cadParam = cadParam & "pOcultaProv=1|"
            numParam = numParam + 1
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
                If vParamAplic.ContabilidadNueva Then
                    Codigo = "pagos"
                Else
                    Codigo = "spagop"
                End If
                Codigo = "UPDATE " & Codigo & " set fecefect = " & DBSet(txtCodigo(149 + indCodigo).Text, "F")
                Codigo = Codigo & " WHERE " & cmdActVtosFraPro.Tag & " AND numorden =" & Label3(120 + indCodigo).Tag
                ConnConta.Execute Codigo
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
Dim i As Byte

    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = False
    Next i
End Sub

Private Sub cmdElimiaFacturas_Click()
Dim B As Boolean

'Igual hay que quitarlo


    'Proceso de borre de facturas
    If cmbEliFac.ListIndex < 0 Then Exit Sub
    
    
    
    'Tablas que voy a tener que borrar
    'Para que no se queden datos
    cadTitulo = String(60, "*") & vbCrLf & vbCrLf
    cadTitulo = cadTitulo & " Se eliminarán los datos con fecha anterior a la solicitada de: " & vbCrLf
    cadTitulo = cadTitulo & " CLIENTES: " & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "Albaranes, hco albaranes, ofertas, hco ofertas, pedidos, hco pedidos" & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "facturas, hco facturas, ventas tpv, reparaciones, hco reparaciones, produccion" & vbCrLf & vbCrLf & vbCrLf
    cadTitulo = cadTitulo & " PRVEEDORES: " & vbCrLf
    cadTitulo = cadTitulo & Space(20) & "Albaranes, hco albaranes,  pedidos, hco pedidos, facturas, hco facturas " & vbCrLf & vbCrLf & vbCrLf
    
    Codigo = cadTitulo & "El proceso es irreversible." & vbCrLf & vbCrLf & vbCrLf & "SEGURO QUE DESEA CONTINUAR?"
    
    'Reestablecer variables
    InicializarVbles
    cadTitulo = ""
    
    If MsgBox(Codigo, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
'-- cambiado por lo de abajo
'    Codigo = InputBox("Password seguridad")
'    Codigo = UCase(Codigo)
'    If Codigo <> "ARIADNA" Then Exit Sub
'++
    Codigo = ""
    Set frmPWD = New frmMensajes
    frmPWD.OpcionMensaje = 30
    frmPWD.Show vbModal
    Set frmPWD = Nothing
    If Codigo <> "ARIADNA" Then Exit Sub
    
    Label3(83).Caption = "Inicio del proceso del borre de facturas"
    Me.cmdElimiaFacturas.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    'Conn.BeginTrans
    B = BorrarFacturas
    'Conn.RollbackTrans
    'Volvemos a dejarlo todo como estaba
    Set miRsAux = Nothing
    Orden1 = ""
    Codigo = ""
    Label3(83).Caption = ""
    Me.cmdElimiaFacturas.Enabled = True
    Screen.MousePointer = vbDefault
    
    If B Then Unload Me
End Sub

Private Sub cmdEtiqBulto_Click()
Dim i As Integer

    cadParam = ""
    If OpcionListado = 95 Then
        If Me.txtClie.Text = "" Then cadParam = "Ponga el cliente"
    Else
        If Me.txtCodigo(148).Text = "" Or Me.txtNombre(148).Text = "" Then cadParam = "Ponga el proveedor"
    End If
    If cadParam <> "" Then
        MsgBox cadParam, vbExclamation
        Exit Sub
    End If
        
        
    
    If Val(txtBultos(1).Text) = 0 Then txtBultos(1).Text = "1"
    cadParam = "delete from tmpinformes where codusu =" & vUsu.Codigo
    conn.Execute cadParam
       
    numParam = 0
    
    Orden2 = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,nombre2,nombre3) VALUES (" & vUsu.Codigo & ","
    If OpcionListado = 95 Then
        cadParam = "," & vParam.Codigo & ",'" & DevNombreSQL(txtNombre(10).Text) & "')"
    Else
        cadParam = "," & vParam.Codigo & ",'" & DevNombreSQL(txtNombre(148).Text) & "')"
    End If
    cadFormula = ""
    If txtBultos(7).Text <> "" Then
        'Lleva etiquetas en blanco
        For i = 1 To Val(txtBultos(7).Text)
            '           secuencia               'El cliente a blancos
            numParam = numParam + 1
            cadFormula = numParam & ",''"
            cadFormula = Orden2 & cadFormula & cadParam
            conn.Execute cadFormula
        Next i
    End If
    For i = 1 To Val(txtBultos(1).Text)
          '           secuencia               'El cliente a blancos
            numParam = numParam + 1
            cadFormula = numParam & ",'" & txtClie.Text & "'"
            cadFormula = Orden2 & cadFormula & cadParam
            conn.Execute cadFormula
            
    Next i
    cadFormula = ""
       
    'Como puede llevar saltos de linea
    Orden2 = SaltosDeLinea(txtBultos(0).Text)
    'Le pasare los datos
    cadParam = ""
    numParam = 0
    If PonerParamRPT2(19, cadParam, numParam, cadNomRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        Orden1 = "0"    'Ahora siempre es CERO. No tiene sentido pasar al rpt

        'Metemos los campos de direccion
        cadParam = cadParam & "Dom=""" & txtBultos(2).Text & """|"
        cadParam = cadParam & "Pob=""" & txtBultos(3).Text & """|"
        cadParam = cadParam & "Pro=""" & Trim(txtBultos(4).Text & "      " & txtBultos(5).Text) & """|"
        
        
        'Si lleva departamento lo metere
        cadSelect = ""
        If cmbBulto.ListIndex > 0 Then
            'Ha cogido departamento
            i = InStr(1, cmbBulto.Text, ":")
            If i = 0 Then
                'NO  deberia pasar nunca
                MsgBox "Error nombre departamento", vbExclamation
            Else
                cadSelect = Trim(Mid(cmbBulto.Text, 1, i - 1))
                cadSelect = Replace(cadSelect, """", "'")
            End If
        End If
        cadParam = cadParam & "nomdirec=""" & cadSelect & """|"
        
        'AÑado la direccion que se ve
        cadParam = cadParam & "DireccionAlternativa=0|"
        'cadParam = cadParam & "DireccionAlternativa=" & Orden1 & "|"
        cadParam = cadParam & "Texto= """ & Orden2 & """|"
        numParam = numParam + 2
        cadSelect = "codusu=" & vUsu.Codigo
        cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
        LlamarImprimir True
        If Me.NumCod <> "" Then Unload Me
    End If
        
End Sub

'INTENTARE METERLO DENTRO DE OTRO PROC

'Abril 2010
'En una columna de tmpinforme voy a grabar el dto para la familia
'De moemnto pong la UNO a piñon
'Veremos si hay que pedir datos o no. De momento esta a piñon

Private Sub cmdEtiqEstanteria_Click()
'Estadistica margen ventas por articulo
Dim campo As String
Dim param As String
Dim tabla As String
Dim RS As ADODB.Recordset
Dim Li As Collection
Dim i As Integer
Dim Dto As Currency
Dim Precio As Currency
Dim Codfamia As Integer

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    cadParam = cadParam & "|pImprimeBarras=""" & Abs(Me.chkImprimeCodigoBarras.Value) & """|"
    numParam = numParam + 1
    cadParam = cadParam & "|numerodecimales=" & Me.cboDecimal.List(cboDecimal.ListIndex) & "|"
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
            
        'Cadena para seleccion D/H artículo
        '--------------------------------------------
        If txtCodigo(92).Text <> "" Or txtCodigo(93).Text <> "" Then
            campo = "{sartic.codartic}"
            param = "pDHArticulo=""Artículo: "
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
    tabla = "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo
    conn.Execute tabla
    
    'Añadire los tipos de IVA a esta tabla
    tabla = "INSERT INTO tmpinformes(codusu,codigo1)  select " & vUsu.Codigo & ",codigiva from sartic"
    If cadSelect <> "" Then tabla = tabla & " WHERE " & cadSelect
    tabla = tabla & " GROUP BY codigiva"
    conn.Execute tabla
    
    
    
    
    
    'AHora desde conta cargo los % de IVA desde la conta
    Set RS = New ADODB.Recordset
    tabla = "Select * from tmpinformes where codusu =" & vUsu.Codigo
    RS.Open tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Li = New Collection
    While Not RS.EOF
        Li.Add Val(RS.Fields(1))
        RS.MoveNext
    Wend
    RS.Close
    
    
    '
    
    'Abrimos los IVAS en conta
    tabla = "Select codigiva,porceiva from tiposiva"
    RS.Open tabla, ConnConta, adOpenKeyset, adLockOptimistic, adCmdText
    For i = 1 To Li.Count
        tabla = "codigiva = " & Li.Item(i)
        RS.Find tabla, , , 1
        If RS.EOF Then
            MsgBox "Tipo de IVA no encontrado en la contabilidad" & tabla, vbExclamation
            RS.Close
            Exit Sub
        Else
            tabla = "UPDATE tmpinformes SET porcen1 =" & TransformaComasPuntos(CStr(RS!PorceIVA))
            tabla = tabla & " WHERE codusu =" & vUsu.Codigo & " AND codigo1 = " & RS!Codigiva
            conn.Execute tabla
        End If
    Next i
    RS.Close
    Set Li = Nothing
    
    
    'Borramos los datos de la tabla donde iran los articulos
    tabla = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.Codigo
    conn.Execute tabla
    i = Me.cboDecimal.List(cboDecimal.ListIndex)
    If i = 0 Then
        tabla = "0"
    Else
        tabla = "#,##0." & Mid("0000", 1, i)
    End If
    frmMensajes.cadWHERE2 = tabla
    frmMensajes.cadWhere = cadSelect
    frmMensajes.vCampos = ""  'estaopcion en etiquetas es para mostrar las del almacen con punto de pedido indicado
    frmMensajes.OpcionMensaje = 15
    frmMensajes.Show vbModal
    
    'Si ha devuelto seleccionados
    tabla = " tmpnseries   "
    cadFormula = " codusu =" & vUsu.Codigo
    
    If Not HayRegParaInforme(tabla, cadFormula) Then Exit Sub
    
    
    If OpcionListado = 513 And vParamAplic.NumeroInstalacion = vbTaxco Then
        'La cantidad será la que tenga el albaran
        
        
        
            tabla = "Select codartic,cantidad,codalmac from slialp " & NumCod & " AND cantidad >=1 "
            tabla = tabla & " AND codartic IN (select codartic from tmpnseries WHERE codusu = " & vUsu.Codigo & ")"
            Set miRsAux = New ADODB.Recordset
            
            
            
            
            
            
            miRsAux.Open tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Codfamia = -1
            While Not miRsAux.EOF
                cadFormula = "nummante"
                Codigo = "codartic = " & DBSet(miRsAux!codArtic, "T") & " AND codusu "
                tabla = DevuelveDesdeBD(conAri, "numserie", "tmpnseries", Codigo, vUsu.Codigo, "N", cadFormula)
                CadTag = "codalmac=" & miRsAux!codAlmac & " AND codartic "
                CadTag = DevuelveDesdeBD(conAri, "ubialmac", "salmac", CadTag, miRsAux!codArtic, "T")
                
                Codigo = ""
                If tabla = "" Then
                    MsgBox "Error leyendo talba temporal", vbExclamation
                Else
                    If Val(cadFormula) > 0 Then
                        indCodigo = Val(cadFormula) - 1
                    Else
                        indCodigo = miRsAux!cantidad - 1
                    End If
                    While indCodigo <> 0
                        'tmpnseries(codusu,codartic,numserie,numlinealb,numlinea)
                        Codigo = Codigo & ", (" & vUsu.Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & DBSet(tabla, "T")
                        Codigo = Codigo & "," & indCodigo & ",1)"
                        indCodigo = indCodigo - 1
                    Wend
                    If Codigo <> "" Then
                        Codigo = Mid(Codigo, 2)
                        Codigo = "INSERT INTO tmpnseries(codusu,codartic,numserie,numlinealb,numlinea) VALUES " & Codigo
                        ejecutar Codigo, False
                    End If
                    
                    
                        Espera 0.2
                        Codigo = "UPDATE tmpnseries SET nummante=" & DBSet(CadTag, "T") & " WHERE codusu = " & vUsu.Codigo & " AND codartic=" & DBSet(miRsAux!codArtic, "T")
                        ejecutar Codigo, False
                    
                    CadTag = ""
                    
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
    
        
    End If
    
    'Para los articulos que hay que mostrar, si tienen dto hay que poner
    'cargalro
    If Me.chkDtoFM.Value = 1 Then
        'Cargo los dtos
        'A piñon para ALZIRA
        tabla = "select * from sdtofm where codactiv=1 and codclien is null and codmarca is null and codfamia >=0 order by codfamia "
        RS.Open tabla, conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        If Not RS.EOF Then
                tabla = "SELECT tmpinformes.codusu,`sartic`.`nomartic`, `sartic`.`preciove`, `tmpinformes`.`porcen1`, `sartic`.`codartic`,codfamia,codmarca,numlinea"
                tabla = tabla & " FROM   ((`tmpnseries` `tmpnseries` INNER JOIN `sartic` `sartic` ON `tmpnseries`.`codartic`=`sartic`.`codartic`)"
                tabla = tabla & " INNER JOIN `tmpinformes` `tmpinformes` ON (`sartic`.`codigiva`=`tmpinformes`.`codigo1`)"
                tabla = tabla & " AND (`tmpnseries`.`codusu`=`tmpinformes`.`codusu`)) Where tmpinformes.CodUsu = " & vUsu.Codigo & " ORDER BY codfamia,codmarca"
                Set miRsAux = New ADODB.Recordset
                
                
                miRsAux.Open tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Codfamia = -1
                While Not miRsAux.EOF
                    
                    If Codfamia <> miRsAux!Codfamia Then
                        'Hay que buscar
                        i = 1
                    Else
                        i = 0
                    End If
                    
                    
                    If i = 1 Then
                        Codfamia = miRsAux!Codfamia
                        Dto = 0
                        RS.MoveFirst
                        tabla = ""
                        While i = 1
                            If RS!Codfamia = Codfamia Then
                                'OK. ESte es. No muevo
                                i = 0 'salga
                                Dto = RS!dtoline1 + RS!dtoline2
                            Else
                                If RS!Codfamia > Codfamia Then RS.MoveLast
                                RS.MoveNext
                            End If
                            If RS.EOF Then i = 0
                        Wend
                    End If
                    If Not RS.EOF Then
                        'OK hay dto
                        
                        If Dto > 0 Then
                            Precio = DBLet(miRsAux!Porcen1, "N")
                            Precio = (miRsAux!PrecioVe * Precio) / 100
                            Precio = Precio + miRsAux!PrecioVe
                            Precio = (Precio * Dto) / 100
                            
                            If Precio > 0 Then
                                tabla = Format(Precio, FormatoCantidad)
                                
                                tabla = "update tmpnseries set nummante = '" & tabla & "' WHERE codusu = " & vUsu.Codigo
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
        
        
        RS.Close
    End If
    
    
    
    
    cadFormula = "({tmpnseries.codusu} =" & vUsu.Codigo & ")"
    
    campo = ""
    If Not PonerParamRPT2(23, cadParam, numParam, campo, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        cadNomRPT = "rEtiqEsta.rpt"
    Else
        cadNomRPT = campo
    End If
    
    LlamarImprimir True
    
    BorrarTempInformes
    
    'Borramos los datos de la tabla donde iran los articulos
    tabla = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.Codigo
    conn.Execute tabla
    
End Sub



Private Sub cmdFactAlbaranes_Click()
    Codigo = "¿Seguro que desea continuar?"
    If MsgBox(Codigo, vbYesNo + vbQuestion) = vbNo Then Exit Sub
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
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'Añadir parametro es departamento o direccion
    cadParam = cadParam & "|pDpto=" & vParamAplic.HayDeparNuevo & "|"
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
'        'AÑado la direccion que se ve
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
    Codigo = ""
    For indCodigo = 110 To 112
        If txtCodigo(indCodigo).Text = "" Then Codigo = Codigo & "M"
        If indCodigo > 110 Then If txtNombre(indCodigo).Text = "" Then Codigo = Codigo & "M"
    Next indCodigo
    If Codigo <> "" Then
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
Dim Codigo  As String

    InicializarVbles

    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    
        cadNomRPT = "rManListTeorico.rpt"
    
        
        cadTitulo = "Informe Mantenimientos"
        Codigo = "scaman"
    
    cadFrom = "(" & Codigo & " INNER JOIN sclien ON " & Codigo & ".codclien=sclien.codclien) "
      
    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(102).Text <> "" Or txtCodigo(103).Text <> "" Then
        campo = "{" & Codigo & ".codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 102, 103, devuelve) Then Exit Sub
    End If

    
    'Cadena para seleccion TIPO CONTRATO
    '--------------------------------------------
    If txtCodigo(104).Text <> "" Or txtCodigo(105).Text <> "" Then
        campo = "{" & Codigo & ".codtipco}"
        'Parametro Desde/Hasta Tipo Contrato
        devuelve = "pDHTipoCon=""Tipo Contrato: "
        If Not PonerDesdeHasta(campo, "T", 104, 105, devuelve) Then Exit Sub
    End If
       
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    'Si  detalla o no
    cadParam = cadParam & "Detallar=" & Abs(Me.chkMante(0).Value) & "|"
    numParam = numParam + 1

    
    LlamarImprimir False
End Sub

Private Sub cmdSelTodos_Click()
Dim i As Byte

    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = True
    Next i
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
        cadTitulo = "Almacén: "
        If Not PonerDesdeHasta(Orden1, "N", 139, 140, cadTitulo) Then Exit Sub
        Orden2 = cadTitulo
    End If
    If Me.chkVarios(0).Value Then Orden2 = Orden2 & "            * Sin stock mínimo"
    Orden2 = "|pDHZona=""" & Trim(Orden2) & """|"
    cadParam = cadParam & Orden2
    
    
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
        cadTitulo = "Listado Stock mínimo"
        cadNomRPT = "rAlmStockMinimos.rpt"
        conSubRPT = False
        cadFormula = "{tmpInformes.codusu}=" & vUsu.Codigo
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
            '4:Tipos de Artículos, 6:Artículos
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
        Case 9 'Informe Movimientos Artículos
            'PonerFoco txtCodigo(5)
            IndiceFoco = 5
        Case 11     '11: Listado de Articulos con componentes ' ====  [16/09/2009] LAURA
            IndiceFoco = 125
            Orden1 = "nomalmac"
            Codigo = DevuelveDesdeBD(conAri, "codalmac", "salmpr", "codalmac>0 AND 1", "1", , Orden1)
            txtCodigo(145).Text = Codigo
            txtNombre(145).Text = Orden1
            
            
        Case 12, 13, 14, 15, 16, 17, 19
                        '12: Listado Toma de Inventario Articulos
                        '13: Listado Diferencias de Inventario Articulos
                        '14: Actualizar Diferencias de Inventario (No IMPRIME INFORME)
                        '15: Listado Articulos Inactivos
                        '16: Listado Valoracion de Stocks Inventariados
                        '17: Listado Valoración Stocks
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
        Case 60 '60: Informe Reparacions - Nº Series
            'PonerFoco txtCodigo(37)
            IndiceFoco = 37
        Case 63
            IndiceFoco = 131
            
        Case 73
            '63: Listado Reparaciones x día
            IndiceFoco = 31
        
        
        Case 223
        
        
        
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
            
        '++ añadido lo del nivel
        Case 97
            If vUsu.Nivel > 1 Then
                MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
                Unload Me
                Exit Sub
            End If
        
            
            
        Case 309 '309:Listado precios de compra
            'PonerFoco txtCodigo(79)
            IndiceFoco = 79
        Case 407 'Sustitución Nº Serie
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
            
            
            If vParamAplic.NumeroInstalacion = vbTaxco Then cmdEtiqEstanteria_Click
            
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
    CargaIconosAyuda2
    'Ocultar todos los Frames de Formulario
    frameListado.visible = False
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
        H = 7215
        W = 6920
        PonerFrameVisible Me.FrameDtosFM, True, H, W
        ponerOptVisible True
        Me.Frame4.visible = True
        indFrame = 6
       ' txtCodigo(79).TabIndex = 318
       ' txtCodigo(80).TabIndex = 318
        cboVarios(0).visible = True
        optFrDto(0).Value = True
        
        'JUNIO 2014
        'Esta opcion es nueva. No debe poner nada
        Me.chkVarios(6).Value = 0
        chkVarios(6).visible = False
        
            Me.chkVarios(6).Top = 5760
            Me.chkVarios(6).visible = True

    Case 58 '58: listado Proveedores
        PonerFrameListadoVisible True, H, W
        Me.lblTitulo(1).Caption = "Listado Proveedores"
        indFrame = 1
        Codigo = "{sprove.codprove}"
        Orden1 = "{sprove.codprove}"
        Orden2 = "{sprove.nomprove}"
        
        
    'LISTADOS DE REPARACIONES
    '-------------------------
    Case 60 '60: Informe Nº Series
        H = 5415
        W = 6675
        PonerFrameVisible Me.FrameRepNSerie, True, H, W
        indFrame = 6
        Codigo = "{sserie"
        
     Case 61, 65  'Listados de Motivos Pend. Rep.
        PonerFrameListadoVisible True, H, W
        Me.lblTitulo(1).Caption = "Listado de Motivos"
        indFrame = 1
        If OpcionListado = 61 Then
            Codigo = "{smotre.codmotre}"
            Orden1 = "{smotre.codmotre}"
            Orden2 = "{smotre.nommotre}"
        Else
            Codigo = "{smotba.codmotiv}"
            Orden1 = "{smotba.codmotiv}"
            Orden2 = "{smotba.desmotiv}"
        End If
        
    Case 63, 73, 223, 224, 248
                '63: Listado Reparaciones por Día
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
            Cad = "(" & Cad & " OR codtipom IN ('FRT','FMO'))"
            Cad = Cad & " and not isnull(letraser) and trim(letraser)<>''"
            
            'Febrero 2011
            'Solo los usuarios de B podran contabilizar fras de B
            If Val(vUsu.AlmacenPorDefecto2) = vParamAplic.AlmacenB Then
                'Es usuario de B. Solo tienen b
                Cad = Cad & " and codtipom = 'FAZ'"
            Else
                'No ven el B
                Cad = Cad & " and codtipom <> 'FAZ'"
            End If
            CargarCombo_TipMov Me.cboTipMov, "stipom", "codtipom", "nomtipom", Cad, True
            
            If vParamAplic.NumeroInstalacion = vbEuler Then cboTipMov.AddItem "FPY - Proyectos"
                
            
            
            'Si es usuario de B solo ha cargado el B
            If Val(vUsu.AlmacenPorDefecto2) = vParamAplic.AlmacenB Then
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
             'Me.Label4(21).Caption = "Fecha Reparación:"
             txtCodigo(0).Text = "1"
        End If
        
        
        
    Case 82, 83
        
        'LIstado etiquetas estanterias
        H = Me.FrameAlbaranesMarcaFacturar.Height
        W = FrameAlbaranesMarcaFacturar.Width
        PonerFrameVisible Me.FrameAlbaranesMarcaFacturar, True, H, W
        indFrame = 82
        If OpcionListado = 82 Then
            cadTitulo = "Poner marca facturación"
            
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
            
            If vParamAplic.NumeroInstalacion = vbTaxco Then chkImprimeCodigoBarras.Value = 1
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
            '- Traer datos del Albaran: cliente, dpto, nº bultos
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
        Label11(0).Caption = "Este proceso es irreversible." & vbCrLf & " No debería haber nadie trabajando en esta empresa y " & vbCrLf & _
            "debería hacer una copia de seguridad."
        
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
'Form de Mantenimiento de Marcas de Artículos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoMotivos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Marcas de Artículos
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
'Form de Mantenimiento de Tipos de Artículo
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


Private Sub frmPWD_DatoSeleccionado(CadenaSeleccion As String)
    Codigo = CadenaSeleccion
End Sub

Private Sub imgAyuda_Click(Index As Integer)
    Select Case Index
    Case 0
        Codigo = "-->Etiquetas." & vbCrLf
        Codigo = Codigo & "-Si marca con 'stock mínimo' pedira un almacen y mostrara" & vbCrLf
        Codigo = Codigo & "los articulos que tenga valor para ese dato" & vbCrLf
        Codigo = Codigo & vbCrLf & "->PVP" & vbCrLf
        Codigo = Codigo & "--Listado PVP, con IVA" & vbCrLf
        Codigo = Codigo & "-Si seleciona precio minimo saldran tambien este ultimo "
    Case 1
        Codigo = "Mostrara los datos de stock minimo,maximo , punto de pedido y stock." & vbCrLf
        Codigo = Codigo & "Si marca sin 'stock mínimo' mostrará los articulos que tienen stock" & vbCrLf
        Codigo = Codigo & "y no tienen valor en el campo stock minimo" & vbCrLf
        
    Case 2
        Codigo = "Descuentos familia marca." & vbCrLf & vbCrLf
        Codigo = Codigo & "CLIENTE - ACTIVIDAD: Mostrara todos los descuentos del cliente y los de la actividad que le corresponda."
        Codigo = Codigo & " No tendra en cuenta el resto de desde/hasta, solo cliente"
        Codigo = Codigo & vbCrLf & vbCrLf
        Codigo = Codigo & "Resto opciones: Mostrara los descuentos desde la tabla de descuentos familia marca, teniendo en cuenta"
        Codigo = Codigo & " la opcion del proveedor de ocultar en listados descuento"
        Codigo = Codigo & vbCrLf & vbCrLf
        Codigo = Codigo & "Ocultar datos proveedor: Ordenando por cliente, ocultará las columnas relacionadas con el proveedor de Dtos y rappels"
    Case 3
        Codigo = "El procentaje de margen puede calcularse sobre el coste o sobre las ventas"
     
    End Select
    
    MsgBox Codigo, vbInformation
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
            
            Case 23 'Listado de Formas de Envío
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
'                Set frmMtoProveedor = New frmComProveedores
'                frmMtoProveedor.DatosADevolverBusqueda = "0|1"
'                frmMtoProveedor.Show vbModal
                Set frmMtoProveedor = New frmBasico2
                AyudaProveedores frmMtoProveedor, txtCodigo(indCodigo)
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
'            Set frmMtoFamilia = New frmAlmFamiliaArticulo
'            frmMtoFamilia.DatosADevolverBusqueda = "0|1"
'            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = New frmBasico2
            AyudaFamilias frmMtoFamilia, txtCodigo(indCodigo)
            Set frmMtoFamilia = Nothing
            
            
        Case 90, 91, 92
            indCodigo = 22 + Index
            Set frmMtoIncid = New frmIncidencias
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
        Case 4, 5, 21, 22, 59, 60, 110, 111 'cod. MARCA
            Select Case Index
                Case 4, 5: indCodigo = Index + 73
                Case 21, 22: indCodigo = Index + 43
                Case 59, 60:  indCodigo = Index - 32
                Case 110, 111:  indCodigo = Index + 44
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
            ' ====  [16/09/2009] LAURA : Listado Articulos con componentes (añade index 95 y 98)
            Select Case Index
                Case 11, 12: indCodigo = Index + 3
                Case 27, 28: indCodigo = Index + 43
                Case 29, 30: indCodigo = Index - 24
                Case 61, 62: indCodigo = Index - 32
                Case 69, 70, 71, 72: indCodigo = Index + 21
                ' ====  [16/09/2009] LAURA : Listado Articulos con componentes (añade index 95 y 98)
                Case 95: indCodigo = 125
                Case 98: indCodigo = 126
                ' ====
            End Select
            Set frmMtoArticulos = New frmBasico2
            'frmMtoArticulos.DatosADevolverBusqueda3 = "@1@" 'Abrimos en Modo Busqueda
            AyudaArticulos frmMtoArticulos, txtCodigo(indCodigo)
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
'            Set frmMtoProveedor = New frmComProveedores
'            frmMtoProveedor.DatosADevolverBusqueda = "0|1"
'            frmMtoProveedor.Show vbModal
            Set frmMtoProveedor = New frmBasico2
            AyudaProveedores frmMtoProveedor, txtCodigo(indCodigo)
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
'            Set frmMtoTraba = New frmAdmTrabajadores
'            frmMtoTraba.DatosADevolverBusqueda = "0|1"
'            frmMtoTraba.Show vbModal
            Set frmMtoTraba = New frmBasico2
            AyudaTrabajadores frmMtoTraba, txtCodigo(indCodigo)
            Set frmMtoTraba = Nothing
            
        Case 45, 46, 112 To 115 'cod. AGENTE
            If Index < 112 Then
                indCodigo = Index + 4
            Else
                
                indCodigo = Index + 44
            End If
'            Set frmMtoAgentes = New frmFacAgentesCom
'            frmMtoAgentes.DatosADevolverBusqueda = "0|1"
'            frmMtoAgentes.Show vbModal
            Set frmMtoAgentes = New frmBasico2
            AyudaAgentesComerciales frmMtoAgentes, txtCodigo(indCodigo), , True
            Set frmMtoAgentes = Nothing

        Case 37, 38, 47, 48, 82, 83 'cod. TIPO CONTRATO (= nº mantenimiento)
            Select Case Index
                Case 37, 38: indCodigo = Index + 20
                Case 47, 48: indCodigo = Index + 4
                Case 82, 83: indCodigo = Index + 22
            End Select
            Set frmMtoTiposCon = New frmManTiposContrato
            frmMtoTiposCon.DatosADevolverBusqueda = "0|1"
            frmMtoTiposCon.Show vbModal
            Set frmMtoTiposCon = Nothing
        
        Case 39, 40, 53, 54 'cod. Nº CONTRATO (= nº mantenimiento)

        
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
        Case 26, 27
            indCodigo = Index + 134
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
        Label2(2).Caption = "Fecha Recepción: "
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
Dim RS As ADODB.Recordset
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
                        Set RS = New ADODB.Recordset
                        Codigo = "select nomclien,domclien,sclien.codpobla as cpos,sclien.pobclien,proclien,"
                        If Me.optBultos(1).Value Then
                            'Direcciones de envio
                            'nomdiren   domdiren     pobdiren   pobdiren    prodiren
                            Codigo = Codigo & "nomdiren  nomdirec ,domdiren  domdirec ,pobdiren pobdirec ,sdirenvio.codpobla  ,prodiren prodirec, coddiren CodDirec"
                            Codigo = Codigo & ",clivario from sclien left join sdirenvio on sclien.codclien=sdirenvio.codclien "
                            
                        
                        Else
                            'Departamentos
                            Codigo = Codigo & " nomdirec ,  domdirec ,pobdirec ,sdirec.codpobla  ,prodirec, CodDirec"
                            Codigo = Codigo & ",clivario from sclien left join sdirec on sclien.codclien=sdirec.codclien "
    
                        End If
                                
                        Codigo = Codigo & " WHERE sclien.codclien =" & txtClie.Text
                        Codigo = Codigo & " order by 6"   'nomdirec nomdiren
                        RS.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        Orden1 = ""
                        
                        While Not RS.EOF
                            'Meto primero la direccion de la ficha
                            If Orden1 = "" Then
                                cmbBulto.AddItem "Ppal:  " & DBLet(RS.Fields(1), "T") & " - " & DBLet(RS.Fields(3), "T")
                                txtBultos(2).Tag = DBLet(RS.Fields(1), "T") & "|"
                                txtBultos(3).Tag = DBLet(RS.Fields(3), "T") & "|"
                                txtBultos(4).Tag = DBLet(RS.Fields(2), "T") & "|"
                                txtBultos(5).Tag = DBLet(RS.Fields(4), "T") & "|"
                                txtBultos(6).Tag = "|"
                                Orden1 = "T"
                                
                                Orden2 = RS!NomClien
                                Clivario = DBLet(RS!Clivario, "N") = 1
                            End If
                            'Las direcciones alternativas
                            If Not IsNull(RS!domdirec) Or Not IsNull(RS!domdirec) Then
                                'TIENE DIRECCION ALTERNATIVA
                                txtBultos(2).Tag = txtBultos(2).Tag & DBLet(RS!domdirec, "T") & "|"
                                txtBultos(3).Tag = txtBultos(3).Tag & DBLet(RS!pobdirec, "T") & "|"
                                txtBultos(4).Tag = txtBultos(4).Tag & DBLet(RS!codpobla, "T") & "|"
                                txtBultos(5).Tag = txtBultos(5).Tag & DBLet(RS!prodirec, "T") & "|"
                                txtBultos(6).Tag = txtBultos(6).Tag & "|"
'                                cmbBulto.AddItem "       " & DBLet(RS!domdirec, "T") & " - " & DBLet(RS!pobdirec, "T")
                                cmbBulto.AddItem DBLet(RS!nomdirec, "T") & ":   " & DBLet(RS!domdirec, "T") & " - " & DBLet(RS!pobdirec, "T")
                                If Me.CadTag = CStr(DBLet(RS!CodDirec, "N")) Then
                                    Ind = cmbBulto.ListCount - 1
                                End If
                            End If
                            RS.MoveNext
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
                        RS.Close
                        Set RS = Nothing

                        
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
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 13: KEYBusqueda KeyAscii, 10 'almacen
            Case 14: KEYBusqueda KeyAscii, 11 'articulo
            Case 15: KEYBusqueda KeyAscii, 12 'articulo
            Case 16: KEYBusqueda KeyAscii, 13 'familia
            Case 17: KEYBusqueda KeyAscii, 14 'familia
            Case 18: KEYBusqueda KeyAscii, 15 'proveedor
            Case 19: KEYBusqueda KeyAscii, 16 'proveedor
            Case 20: KEYFecha KeyAscii, 2 'fecha
            Case 22: KEYFecha KeyAscii, 3 'fecha
            Case 21:  KEYBusqueda KeyAscii, 17 'trabajador
            
            
            Case 154: KEYBusqueda KeyAscii, 110 'marca
            Case 155: KEYBusqueda KeyAscii, 111 'marca
            
            Case 107: KEYBusqueda KeyAscii, 87 'almacen
            Case 62: KEYBusqueda KeyAscii, 19 'familia
            Case 63: KEYBusqueda KeyAscii, 20 'familia
            Case 64: KEYBusqueda KeyAscii, 21 'marca
            Case 65: KEYBusqueda KeyAscii, 22 'marca
            Case 66: KEYBusqueda KeyAscii, 23 'proveedor
            Case 67: KEYBusqueda KeyAscii, 24 'proveedor
            Case 68: KEYBusqueda KeyAscii, 25 'tipo articulo
            Case 69: KEYBusqueda KeyAscii, 26 'tipo articulo
            Case 70: KEYBusqueda KeyAscii, 27 'articulo
            Case 71: KEYBusqueda KeyAscii, 28 'articulo
            
            Case 72: KEYBusqueda KeyAscii, 18 'almacen
            
            ' contabilizacion de facturas cliente
            Case 132: KEYBusqueda KeyAscii, 101 'cliente
            Case 133: KEYBusqueda KeyAscii, 102 'cliente
            Case 31: KEYFecha KeyAscii, 4 'fecha
            Case 32: KEYFecha KeyAscii, 5 'fecha
            
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscarG_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
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
    'informes. Según de donde llamemos código de una tabla u otra
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
                
            Case 4 'Listado Tipos Artículos
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), 1, "stipar", "nomtipar", "codtipar", "Tipo de Artículo", "T")
    
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
            
            Case 23 'Listado Formas de Envío
                EsNomCod = True
                tabla = "senvio"
                codCampo = "codenvio"
                NomCampo = "nomenvio"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Forma de Envío"
            
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
                Titulo = "Situación Especial"
            
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
            ' ====  [16/09/2009] LAURA : añade index 125,126
            EsNomCod = True
            tabla = "sartic"
            codCampo = "codartic"
            NomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
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
        Case 9, 10, 20, 22, 31, 32, 43, 44, 53, 54, 82, 83, 109, 110, 115, 116, 119, 120, 123, 124, 130, 131, 160, 161
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
        
        Case 27, 28, 64, 65, 77, 78, 154, 155 'MARCAS
            EsNomCod = True
            tabla = "smarca"
            codCampo = "codmarca"
            NomCampo = "nommarca"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Marca"
        
        Case 31 'Nº de Oferta
            If txtCodigo(Index).Text = "" Then Exit Sub
            codCampo = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", txtCodigo(Index).Text, "N")
            If codCampo = "" Then
                MsgBox "No existe el código de Oferta: " & NumCod, vbInformation
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
            'Desde/Hasta se ha seleccionado un único cliente
            If Index = 39 Or Index = 40 Then
                If txtCodigo(37).Text <> txtCodigo(38).Text Then
                    MsgBox "Seleccionar dpto/direc. solo cuando se seleccione un único cliente.", vbInformation
                    txtCodigo(Index).Text = ""
                    Exit Sub
                End If
            ElseIf Index = 35 Or Index = 36 Then
                If txtCodigo(33).Text <> txtCodigo(34).Text Then
                    MsgBox "Seleccionar dpto/direc. solo cuando se seleccione un único cliente.", vbInformation
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
'                    codCampo = " la Dirección "
'                End If
                codCampo = "No existe" & codCampo & txtCodigo(Index).Text & " para el cliente: "
                codCampo = codCampo & txtCodigo(Index - 2).Text & " - " & txtNombre(Index - 2).Text
                MsgBox codCampo, vbInformation
                txtCodigo(Index).Text = ""
                PonerFoco txtCodigo(Index)
            End If
        
        Case 41, 42, 59, 60 'Nº Contrato
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
        
        Case 49, 50, 156, 157, 158, 159 'Cod. AGENTE
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
            
        Case 61 'Año Ejercicio
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "El Ejercicio debe ser un Año", vbInformation
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
            
        Case 121, 122 'Nº Factura
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
    Conexion = conAri    'Conexión a BD: Ariges
    Select Case OpcionListado
        Case 7 'Traspaso de Almacenes
            Cad = Cad & "Nº Trasp|scatra|codtrasp|N|0000000|40·Almacen Origen|scatra|almaorig|N|000|20·Almacen Destino|scatra|almadest|N|000|20·Fecha|scatra|fechatra|F||20·"
            tabla = "scatra"
            Titulo = "Traspaso Almacenes"
        Case 8 'Movimientos de Almacen
            Cad = Cad & "Nº Movim.|scamov|codmovim|N|0000000|40·Almacen|scamov|codalmac|N|000|30·Fecha|scamov|fecmovim|F||30·"
            tabla = "scamov"
            Titulo = "Movimientos Almacen"
        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
                   '12: Inventario Articulos
                   '14:Actualizar Diferencias de Stock Inventariado
                   '16: Listado Valoracion stock inventariado
            Cad = Cad & "Código|sartic|codartic|T||30·Denominacion|sartic|nomartic|T||70·"
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
            Cad = Cad & "Codigo|sdirec|coddirec|N|000|15·"
            Cad = Cad & "Descripcion|sdirec|nomdirec|T||55·"
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
    PonerFrameVisible Me.frameListado, visible, H, W

    If visible = True Then
        Me.Optcodigo.Value = True
    End If
End Sub



Private Sub PonerFrameInventarioVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Inventario Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Inventario
Dim VerOpcion As Boolean


    chkValorDesdeArticulo.visible = False
    FrameMarcaTomaInventario.visible = False
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
        
        chkProv2(3).visible = False 'SOLO ROTACION
        If OpcionListado = 12 Then
            'Toma de inventario
            FrameMarcaTomaInventario.BorderStyle = 0
            FrameMarcaTomaInventario.visible = True
            Me.txtCodigo(20).Top = FrameMarcaTomaInventario.Top + FrameMarcaTomaInventario.Height + 230
            imgFecha(2).Top = txtCodigo(20).Top
            Label4(5).Top = txtCodigo(20).Top
            Me.cmdAceptar(4).Top = 7000
            
            If vParamAplic.NumeroInstalacion = vbHerbelca Then
                chkProv2(3).visible = True
                chkProv2(3).Top = txtCodigo(20).Top
                chkProv2(3).Left = txtCodigo(20).Left + txtCodigo(20).Width + 360
            End If
            
            H = 7400
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
        
        
        chkSaltaPag.Caption = "Salta pág. en Familia"
        If OpcionListado = 13 Then
            If vParamAplic.InventarioxProv Then chkSaltaPag.Caption = "Salta pág. en proveedor"
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
            Me.Label4(5).Caption = "Fec.Inactividad"
        ElseIf OpcionListado = 19 Then
            Me.Label4(5).Caption = "Fecha Stock"
        Else
            Me.Label4(5).Caption = "Fec.Inventario"
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
            Me.Label4(8).Left = 2230 '2280
            Me.imgFecha(2).Left = 2850 '2820
            Me.txtCodigo(20).Left = 3120
            Me.Label4(9).Left = 4660
            Me.imgFecha(3).Left = 5250 '5160
            Me.txtCodigo(22).Left = 5530 '5430
'            txtCodigo(22).TabIndex = 48
        End If
        
        
        '====================================
        'Activar o no los check de Opcion:
        VerOpcion = (OpcionListado = 12) Or (OpcionListado = 13) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19) Or OpcionListado = 15
                    '12: Toma de Inventario
                    '13: Listado Diferencias stock
        
        Me.FrameOpciones2.visible = VerOpcion
        If OpcionListado = 12 Then
            Me.FrameOpciones2.Top = 6400
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
    H = 8655  '7335
    If OpcionListado = 245 Then H = 5075
    W = 7635
    PonerFrameVisible Me.FrameTarifas, visible, H, W
    
    'CABEL
    chkVarios(3).visible = False
    chkVarios(3).Value = 0
    'Ocultar datos proveedor
    chkVarios(5).Value = 0
    chkVarios(5).visible = False
    If visible = True Then
        '====================================
        '28: Tarifas Precios 29: Promociones
        VerOpcion = (OpcionListado = 28) Or (OpcionListado = 29)
        Me.chkSaltaPagTarif.visible = VerOpcion
        Me.Label4(12).visible = VerOpcion
        
        'CABEL
        chkVarios(3).visible = OpcionListado = 28 And vParamAplic.NumeroInstalacion = vbHerbelca
        'Prees
        If OpcionListado = 30 And vParamAplic.NumeroInstalacion = vbHerbelca Then
            chkVarios(5).visible = True
            chkVarios(5).Left = 360
        End If
        
        '====================================
        If OpcionListado = 30 Then Me.Label4(11).Caption = "Cliente"
        
        
        FrameFechasPromo.visible = OpcionListado = 29
        chkSoloRotacion.visible = (OpcionListado = 28)
        '245: Control margenes tarifas
        '==================================
        VerOpcion = OpcionListado = 245 Or OpcionListado = 28
        Me.cboDecimales.visible = VerOpcion
        Label4(88).visible = VerOpcion
        If VerOpcion Then cboDecimales.ListIndex = 2
        VerOpcion = (OpcionListado = 245)
        Me.chkMostrarErrores.visible = VerOpcion
        
        
        Label4(109).visible = OpcionListado = 30
        For numParam = 0 To 1
            imgBuscarG(112 + numParam).visible = OpcionListado = 30
            Label3(129 + numParam).visible = OpcionListado = 30
            txtCodigo(156 + numParam).visible = OpcionListado = 30
            txtNombre(156 + numParam).visible = OpcionListado = 30
        Next
        
        
        
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
                Me.lblTitulo(0).Caption = "Reparaciones por Día"
                Me.Label2(2).Caption = "Fecha Reparación:"
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
Dim B As Boolean
        'Opciones: 70,71,78,79,76
    H = 7695
    W = 6875
    PonerFrameVisible Me.FrameMantenimientos, visible, H, W

    If visible = True Then
        B = (OpcionListado = 70)
        
        Me.cboTipoList.visible = B 'List. Mantenimientos
        Me.Label1(4).visible = B
        
        
        
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
Dim B As Boolean



    'Hay una opcion mas que mostrara este frame. la 247. Correccion de de tarifas e importes en articulos
    FrameTapaINCORRECTO.visible = False
    chkMinimoCorreg.visible = False
    B = (OpcionListado = 6)
    chkImpEtiq(0).visible = B
    chkImpEtiq(1).visible = B
    chkImpEtiq(3).visible = B
    Me.imgayuda(0).visible = B
    If B Then
        Me.Label9.Caption = "Informe de Artículos"
       
        W = 8715
    Else
        If OpcionListado = 18 Then
            Me.Label9.Caption = "Informe Stocks Máximos y Mínimos"
            Label4(36).Caption = "Almacén"
            W = 7495
        Else
            'NUEVA OCPION:  247
            'Corregir tarifas y eso
            chkMinimoCorreg.visible = True
            Me.Label9.Caption = "Verificación tarifas y P.V.P."
            FrameTapaINCORRECTO.visible = True
            Label4(36).Caption = "Tarifa"
            cmbDecimales.ListIndex = 0
            W = 7500 '7395
        End If
        
       
    End If
    H = 7095
    
    
    PonerFrameVisible Me.FrameInfArticulos, visible, H, W
    If visible = True Then
        'visible orden campos si opcionlistado=6
        Me.FrameOrden.visible = B
        Label4(36).visible = Not B

        Me.imgBuscarG(18).visible = Not B
        Me.txtCodigo(72).visible = Not B
        Me.txtNombre(72).visible = Not B
        
        'Visible Frame stocks Max Minimos si opcionlistado=18
        Me.optStockMax.Value = True
        Me.FrameStockMaxMin.visible = OpcionListado = 18
  
        FrameSituacionArticulo.visible = OpcionListado = 6
    
    
        'REajustes.
        'El articulo NO se muestra si la opcion es 247
        B = OpcionListado <> 247
        PonerLabelsArticulosFrameVisible B
        Label4(75).visible = Not B
        cmbDecimales.visible = Not B
        Label4(90).visible = Not B
        cmbDecimales.visible = Not B
    
    End If
    
    
    Me.cmdAceptarArtic.Top = H - cmdAceptarArtic.Height - 120
    cmdCancel(11).Top = H - cmdCancel(11).Height - 120
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
Dim RS As ADODB.Recordset
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", 800
    ListView1.ColumnHeaders.Add , , "Descripción", 2250
    
    SQL = "select codtipom,nomtipom from stipom where muevesto=1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = RS.Fields(0).Value
        ItmX.Checked = True
        ItmX.SubItems(1) = RS.Fields(1).Value
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
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
Dim i As Byte

    PonerFormulaYParametrosInf9 = False
    InicializarVbles
    
    'Parametro EMPRESA
    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = 1
        
    '-- Cadena para seleccion Desde y Hasta ARTICULO
    If txtCodigo(5).Text <> "" Or txtCodigo(6).Text <> "" Then
        Codigo = "{smoval.codartic}"
        devuelve = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(Codigo, "T", 5, 6, devuelve) Then Exit Function
    End If
                    
    '-- Cadena para seleccion Desde y Hasta FAMILIA
    If txtCodigo(7).Text <> "" Or txtCodigo(8).Text <> "" Then
        Codigo = "{sartic.codfamia}"
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 7, 8, devuelve) Then Exit Function
    End If
        
    '-- Cadena para seleccion Desde y Hasta ALMACEN
    If txtCodigo(11).Text <> "" Or txtCodigo(12).Text <> "" Then
        Codigo = "{smoval.codalmac}"
        devuelve = "pDHAlmacen=""Almacen: "
        If Not PonerDesdeHasta(Codigo, "N", 11, 12, devuelve) Then Exit Function
    End If
    
    
    '-- Cadena para seleccion Desde y Hasta CLIENTE/PROVEEDOR
    If txtCodigo(86).Text <> "" Or txtCodigo(87).Text <> "" Then
        Codigo = "{smoval.codigope}"
        devuelve = "pDHOperario=""Cliente/Proveedor/Trab.: "
        If Not PonerDesdeHasta(Codigo, "N", 86, 87, devuelve) Then Exit Function
    End If
    
        
'    cadSelect = QuitarCaracterACadena(cadFormula, "{")
'    cadSelect = QuitarCaracterACadena(cadSelect, "}")
        
    '=================================================
    '-- Cadena para seleccion Desde y Hasta FECHA
    If txtCodigo(9).Text <> "" Or txtCodigo(10).Text <> "" Then
        Codigo = "{smoval.fechamov}"
        devuelve = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(Codigo, "F", 9, 10, devuelve) Then Exit Function
    End If
        
    '-- seleccionar los articulos que tienen control de stock
    Codigo = "{sartic.ctrstock}=1"
    AnyadirAFormula cadFormula, Codigo
    AnyadirAFormula cadSelect, Codigo
        
        
    '-- Cadena de Seleccion TIPOS de MOVIMIENTOS
    Codigo = "{smoval.detamovi}"
    devuelve = ""
    'Si todos seleccionados no añadir la select
    todosMarcados = True
    i = 1
    While Not i > Me.ListView1.ListItems.Count And todosMarcados
        If Not Me.ListView1.ListItems(i).Checked Then todosMarcados = False
        i = i + 1
    Wend
    
    'si no estan todos seleccionados montar select de los seleccionados
    If Not todosMarcados Then
        Cad = ""
        devuelve = ""
        For i = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(i).Checked Then
                If Cad = "" Then
                    Cad = Me.ListView1.ListItems(i).Text
                Else
                    Cad = Cad & ", " & Me.ListView1.ListItems(i).Text
                End If
                If devuelve = "" Then
                    devuelve = Codigo & " = """ & Me.ListView1.ListItems(i).Text & """"
                Else
                    devuelve = devuelve & " or " & Codigo & " = """ & Me.ListView1.ListItems(i).Text & """"
                End If
            End If
        Next i

        If devuelve <> "" Then 'Hay algun movimiento marcado
            If cadFormula <> "" Then
                cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadSelect = cadSelect & " AND " & "(" & devuelve & ")"
                cadParam = cadParam
            Else
                cadFormula = "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadSelect = "(" & devuelve & ")"
            End If
            Cad = "pTiposMov=""Tipos Movimiento: " & Cad
            cadParam = cadParam & Cad & """|"
            numParam = numParam + 1
        Else 'Todos desmarcados
            Cad = ""
            For i = 1 To ListView1.ListItems.Count
                If Cad = "" Then
                    Cad = """" & ListView1.ListItems(i).Text & """"
                Else
                    Cad = Cad & ", """ & ListView1.ListItems(i).Text & """"
                End If
            Next i
            devuelve = Codigo & " NOT IN [" & Cad & "]"
            Cad = Codigo & " NOT IN (" & Cad & ")"
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
        MsgBox "Introduzca algún criterio de selección para el Informe.", vbInformation
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
    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
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
    Codigo = CodAux & "codalmac}"
    If Trim(txtCodigo(13).Text) <> "" Then _
    devuelve = Codigo & " = " & Val(txtCodigo(13).Text)
    If devuelve <> "" Then
        cadFormula = devuelve
        Cad = "pAlmacen= ""Almacen: " & Format(txtCodigo(13).Text, "000") & " " & txtNombre(13).Text
        
        If OpcionListado = 19 Then
            'QUE SALGA LA MARCA DE VARIOS
            If Me.chkProv2(2).Value = 1 Then Cad = Cad & " (VARIOS)"
        End If
        
        cadParam = cadParam & Cad & """|"
        numParam = numParam + 1
    End If
            
    'Cadena para seleccion Desde y Hasta ARTICULOS
    '----------------------------------------------
    If txtCodigo(14).Text <> "" Or txtCodigo(15).Text <> "" Then
        Codigo = CodAux & "codartic}"
        devuelve = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(Codigo, "T", 14, 15, devuelve) Then Exit Function
    End If
    
    'Cadena para seleccion Desde y Hasta FAMILIA
    '--------------------------------------------
    If txtCodigo(16).Text <> "" Or txtCodigo(17).Text <> "" Then
        Select Case OpcionListado
            Case 12, 15, 16, 17, 19: Codigo = "{sartic.codfamia}"
            Case Else: Codigo = "{sinven.codfamia}"
        End Select
        Cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 16, 17, Cad) Then Exit Function
    End If
    cadFrom = cadFrom & " INNER JOIN sartic ON " & CodAux & "codartic=sartic.codartic "
    'Enero 2008
    'David
    cadFormula = cadFormula & " AND {sartic.ctrstock} = 1"
    
    'If
    If OpcionListado = 12 And chkProv2(3).visible Then
        If Me.chkProv2(3).Value Then cadFormula = cadFormula & " AND {sartic.rotacion} = 1"
    End If
    
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
            Case 12, 15, 16, 17, 19: Codigo = "{sartic.codprove}"
            Case Else: Codigo = "{sinven.codprove}"
        End Select
        Cad = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(Codigo, "N", 18, 19, Cad) Then Exit Function
    End If
    

    'Agosto2017
    'Cadena para seleccion Desde y Hasta MARCA en toma inventario
    '------------------------------------------------------------
    If OpcionListado = 12 Then
        If txtCodigo(154).Text <> "" Or txtCodigo(155).Text <> "" Then
            Cad = "pDHMarca=""Marca: "
            Codigo = "{sartic.codmarca}"
            If Not PonerDesdeHasta(Codigo, "N", 154, 155, Cad) Then Exit Function
        End If
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
            Codigo = CodAux & "fechainv}"
            devuelve = CadenaDesdeHasta(txtCodigo(20).Text, txtCodigo(22).Text, Codigo, "F")
    
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
                cadParam = cadParam & Trim(Cad) & """|"
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
        cadParam = cadParam & Cad
        numParam = numParam + 1
        
        'Añadir a la formula de seleccion que no sea uno de la lista
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
                
            cadParam = cadParam & "pFechaStock=""" & Cad & """|"
            numParam = numParam + 1
            
                            
            'Si lleva factor conversion y solo negativos o positivos
            Cad = ""
            devuelve = "0"
            If Me.cboStokFecha.ListIndex > 0 Then Cad = "Sólo " & Me.cboStokFecha.Text
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
            
            cadParam = cadParam & "pdhValora=""" & Cad & """|"
            numParam = numParam + 1
                
            'Incremento
            cadParam = cadParam & "Incremento=" & devuelve & "|"
            numParam = numParam + 1
            
            
            'Detalla
            cadParam = cadParam & "detalla=" & Abs(Me.chkProv2(1).Value) & "|"
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
        cadParam = cadParam & "pSinStock=" & ImprStock & "|"
        numParam = numParam + 1
    End If
    
       
    '============================================
    '============= PARAMETROS ===================
    If OpcionListado = 12 Or OpcionListado = 15 Then
        '12: Toma de Inventario
        '15: Listado Articulos Inactivos
        cadParam = cadParam & "pFechaInve=""" & txtCodigo(20).Text & """|"
        numParam = numParam + 1
    End If
    
    If OpcionListado = 12 Then
        'Parámetro Imprime Stock (Si/No)
        If Me.chkImprimeStock.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pImprimeStock=" & ImprStock & "|"
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
        'Parámetro Salta Pag. en Familia (Si/No)
        If Me.chkSaltaPag.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pSaltaFamilia=" & ImprStock & "|"
        numParam = numParam + 1
    End If
    
    If OpcionListado = 16 Or OpcionListado = 17 Then '16: Valoración de Stocks Inventariados
                                                     '17: Valoración Stocks
        'Parámetro Valorado
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
        cadParam = cadParam & "pValorado=" & strValorado & "|"
        numParam = numParam + 1
        
        
        'Mayo 2013
        
        cadParam = cadParam & "pDesdeArticulo=" & Abs(Me.chkValorDesdeArticulo.Value) & "|"
        numParam = numParam + 1
        
    End If
    
    If (OpcionListado = 15) Or (OpcionListado = 16) Or (OpcionListado = 17) Or (OpcionListado = 19) Then
        'Parametro Precio de Valoracion, elegir un Precio para realizar la valoracion: canstock * precio
        If Me.optPrecioMP.Value Then bytPrecio = 1
        If Me.optPrecioMA.Value Then bytPrecio = 2
        If Me.optPrecioUC.Value Then bytPrecio = 3
        If Me.optPrecioStd.Value Then bytPrecio = 4
        cadParam = cadParam & "pPrecio=" & bytPrecio & "|"
        numParam = numParam + 1
    End If
    '=====================================================================
    
       
    'comprobar si hay registros para mostrar en el Informe antes de Abrirlo
    If Not HayRegParaInforme(cadFrom, cadSelect) Then
        Exit Function
    Else
        If OpcionListado = 12 Then
        If MsgBox("¿Actualizar diferencias?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        End If
    End If
    
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
        cadParam = cadParam & "kprecio= """ & strValorado & """|"
        numParam = numParam + 1
        strValorado = ""
    End If
    
    
    
    
    '12,13,16 inventario
    Select Case OpcionListado
    Case 12, 13, 16
        If Not vParamAplic.InventarioxProv Then
            If vParamAplic.InventarioCodigoArticulo Then
                cadParam = cadParam & "orden= 1|"
                numParam = numParam + 1
            End If
            
        End If
    End Select
    

    PonerFormulaYParametrosInf12 = True
End Function



Private Function PonerFormulaYParametrosInf28() As Boolean
'Informes de Descuentos y Tarifas
Dim Cad As String
Dim cadCodigo As String
Dim Aux As String

    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = 1
    
    PonerFormulaYParametrosInf28 = False
    
    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion Desde y Hasta TARIFA o D/H CLIENTE
    '--------------------------------------------
     If txtCodigo(23).Text <> "" Or txtCodigo(24).Text <> "" Then
        If OpcionListado = 30 Then 'Precios Especiales Cliente
            cadCodigo = Codigo & ".codclien}"
            Cad = "pDHCliente=""Cliente: "
        Else
            cadCodigo = Codigo & ".codlista}"
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
    
    
    
    'Promociones. Si ponen fecha
    If OpcionListado = 29 Then
        If txtCodigo(160).Text = "" Xor txtCodigo(161).Text = "" Then
            MsgBox "O indica fecha inicio y fecha fin, o no indica ninguna", vbExclamation 'HERBELCA
            Exit Function
        End If
        Cad = ""
        If txtCodigo(160).Text <> "" Then
            cadCodigo = CadenaDesdeHasta(txtCodigo(160).Text, "", "{spromo.fechaini}", "F")
            If cadCodigo = "Error" Then Exit Function
            If Not AnyadirAFormula(cadFormula, cadCodigo) Then Exit Function
            cadCodigo = CadenaDesdeHasta("", txtCodigo(161).Text, "{spromo.fechafin}", "F")
            If cadCodigo = "Error" Then Exit Function
            If Not AnyadirAFormula(cadFormula, cadCodigo) Then Exit Function
            cadCodigo = " spromo.fechaini=" & DBSet(txtCodigo(160).Text, "F") & " AND spromo.fechafin=" & DBSet(txtCodigo(161).Text, "F")
            If Not AnyadirAFormula(cadSelect, cadCodigo) Then Exit Function
            Aux = Trim(Aux & "    Fechas promocion: " & txtCodigo(160).Text & " - " & txtCodigo(161).Text)
        End If
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
        cadParam = cadParam & Cad
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
        
        If OpcionListado = 30 Then
            If txtCodigo(156).Text <> "" Or txtCodigo(157).Text <> "" Then
                cadCodigo = "{sclien.codagent}"
                Cad = "    Agente: "
                If Not PonerDesdeHasta(cadCodigo, "N", 156, 157, Cad) Then Exit Function
                Aux = Aux & Cad
            End If
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
            cadParam = cadParam & Cad
            numParam = numParam + 1
        End If
    End If
            
            
            
            
            
            
            
    'Cadena para seleccion Desde y Hasta ARTICULOS
    '----------------------------------------------
    If txtCodigo(29).Text <> "" Or txtCodigo(30).Text <> "" Then
        cadCodigo = Codigo & ".codartic}"
        Cad = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(cadCodigo, "T", 29, 30, Cad) Then Exit Function
    End If
 
 
 
 
 
 
 
    '=====================================================================
    '====   PARAMETROS
    If (OpcionListado = 28 Or OpcionListado = 29) Then
        'Parámetro Salta Pag. en Familia (Si/No)
        If Me.chkSaltaPagTarif.Value = 1 Then
            Cad = "True"
        Else
           Cad = "False"
        End If
        cadParam = cadParam & "pSaltaFamilia=" & Cad & "|"
        numParam = numParam + 1
    End If
       
    If OpcionListado = 245 Then
        'Parámetro mostrar solo tarifas con errores (Si/No)
        Cad = Abs(Val(Me.chkMostrarErrores.Value))
        cadParam = cadParam & "Suprimr=" & Cad & "|"
        numParam = numParam + 1
        'Decimales
    End If
    
    If OpcionListado = 245 Or OpcionListado = 28 Then
        If cboDecimales.ListIndex < 0 Then
            MsgBox "Seleccione decimales", vbExclamation
            Exit Function
        End If
        Cad = (cboDecimales.ItemData(Me.cboDecimales.ListIndex))
        cadParam = cadParam & "Decimales=" & Cad & "|"
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
'Además Actualiza la Tabla:salmac los campos:fechainv, horainve, statusin
Dim SQL As String, ADonde As String
Dim RS As ADODB.Recordset
Dim hora As Date

On Error GoTo EInventario:
   
'   If CrearTmpInventario(cadSelect) Then
   

        'Aqui empieza transaccion
        Screen.MousePointer = vbHourglass
        conn.BeginTrans
    
          
    
'        'Insertar en la tabla de Histórico: shinve
'        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
'        ADonde = "Insertando datos en Histórico. Tabla: shinve"
'        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
'        SQL = SQL & " SELECT salmac.codartic, salmac.codalmac, salmac.fechainv,salmac.horainve,salmac.stockinv "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
'        'si no se ha inventariado antes no lo pasamos al historico
'        SQL = SQL & " AND not isnull(salmac.fechainv) "
'        Conn.Execute SQL
'
        
        'Insertar en la tabla de Histórico: shinve
        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
        ADonde = "Insertando datos en Histórico. Tabla: shinve"
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
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RS.EOF
            
          
            
            
            
            
            'Actualizamos la tabla salmac ponemos statusin=1 para indicar que se
            'esta realizando inventario y bloquear los articulos para que no se puedan
            'realizar movimientos, traspasos, etc.
            'Actualizamos la Tabla: salmac los campos: fechainv, horainve
            ADonde = "Actualizando datos en Articulos x Almacen"
            SQL = "UPDATE salmac SET fechainv='" & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', "
            SQL = SQL & " horainve='" & Format(hora, "yyyy-mm-dd hh:mm:ss") & "', " & "statusin=1 "
            'MAYO 2013 preciomp,precioma,preciouc,preciost
            SQL = SQL & ", preciompin = " & DBSet(RS!PrecioMP, "N")
            SQL = SQL & ", preciomain = " & DBSet(RS!precioma, "N")
            SQL = SQL & ", precioucin = " & DBSet(RS!precioUC, "N")
            SQL = SQL & ", preciostin = " & DBSet(RS!preciost, "N")
            
            'SEPTIEMBRE 2013
            'Incializar stock al inventariar
            SQL = SQL & ", stockinv="
            If vParamAplic.IncializarStockEnInventario Then
                SQL = SQL & " 0"  'stockinv=0 inicializamos el stock del articulo
            Else
                SQL = SQL & " if(canstock>0,canstock,0)"
            End If
            SQL = SQL & " WHERE codartic=" & DBSet(RS.Fields(0).Value, "T") & " AND "
            SQL = SQL & "codalmac=" & RS.Fields(1).Value
            conn.Execute SQL
            RS.MoveNext
        Wend
    
        RS.Close
        Set RS = Nothing
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
    Screen.MousePointer = vbDefault
End Function


Private Function CrearTmpInventario(cadFormula As String) As Boolean
Dim SQL As String
Dim LineaCodartic  As String
Dim B As Boolean

    On Error GoTo ECrearInv
    
    B = False
    
    
    'De momento en FONTENAS
    LineaCodartic = "codartic varchar(16) NOT NULL default '',"
    If vParamAplic.NumeroInstalacion = 5 Then
       ' show full fields from `ariges1`.`sartic` where 1=1
        Set miRsAux = New ADODB.Recordset
        SQL = "show full fields from " & vEmpresa.BDAriges & ".sartic"
        miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While SQL <> ""
            If miRsAux.EOF Then
                SQL = ""
            Else
                
                If LCase(DBLet(miRsAux.Fields(0), "T")) = "codartic" Then
                    
                    'ejemplo:_`codartic` varchar(16) COLLATE latin1_spanish_ci NOT NULL DEFAULT '',
                    SQL = "codartic " & "varchar(16) COLLATE " & DBLet(miRsAux!collation, "T") & " NOT NULL default '',"
                    LineaCodartic = SQL
                    SQL = ""
                Else
                    miRsAux.MoveNext
                End If
            End If
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    
    SQL = "CREATE TEMPORARY TABLE tmpInven ( "
    SQL = SQL & LineaCodartic   '"codartic varchar(16) NOT NULL default '0', "
    SQL = SQL & "codalmac smallint(3) unsigned NOT NULL default '0', "
    SQL = SQL & "codfamia smallint(4) unsigned NOT NULL default '0', "
    SQL = SQL & "codprove int(6) unsigned NOT NULL default '0', "
    SQL = SQL & "fechainv date NOT NULL default '0000-00-00', "
    SQL = SQL & "horainve datetime NOT NULL default '0000-00-00 00:00:00', "
    SQL = SQL & "stockinv decimal(12,2) NOT NULL default '0.00')"
    conn.Execute SQL
    B = True
    
    
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
        B = False
        'Err.Clear
        MuestraError Err.Number, "Crear temporal inventario.", Err.Description
    End If
    CrearTmpInventario = B
End Function






Private Function ActualizarInventario() As Boolean
'-----------------------------------------------------------------
'* Modifica en la Tabla: salmac los campos: cansotck, fechainv, horainve,statusin de los articulos seleccionados
'y les asigna los valores de los campos: existenc, fechainv, horainve, false de la tabla: sinven
'* Elimina de la Tabla: sinven los registros seleccinados para actualizar
'* Inserta Movimientos de Articulos en la Tabla: smoval
'-------------------------------------------------------------------
Dim SQL As String, ADonde As String
Dim RS As ADODB.Recordset
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
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        bol = False
        ActualizarInventario = False
        MsgBox "No existen Registros en la Tabla: sinven para Actualizar Inventario.", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    
    'Obtener el contador para los movimientos del Almacen que se esta inventariando
    'A cada registro de la tabla sinven se le asignará un numero de linea.
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
    
    While Not RS.EOF And bol 'Para cada registro de la tabla sinven
    
        'Introducir Movimiento de Entrada/Salida si hay diferencia entre el
        'Stock del Sistema y el Stock Real Inventariado.
        '------------------------------------------------------------------
        ADonde = "Introduciendo Movimiento de Entrada/Salida. Tabla: smoval."
        DevStock = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", RS!codArtic, "T", , "codalmac", RS!codAlmac, "N")
        If DevStock <> "" Then
            CanStock = CLng(DevStock)
            Diferencia = RS!existenc - CanStock
            If Diferencia <> 0 Then 'Insertar Movimiento de Entrada/Salida en Almacen
                CadValues = DBSet(RS!codArtic, "T") & ", " & RS!codAlmac & ", '" & Format(RS!FechaINV, "yyyy-mm-dd") & "', '"
                CadValues = CadValues & Format(RS!HOraInve, "yyyy-mm-dd hh:mm:ss") & "', "
                bol = InsertarMovimArticulos(CadValues, RS!codArtic, Diferencia, LetraSerie, NumMovim, numlinea)
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
            ADonde = "Actualizando Stock de Artículos en Almacén. Tabla: salmac."
            SQL = "UPDATE salmac SET canstock=" & DBSet(RS!existenc, "N") & ", statusin=0"
            SQL = SQL & " WHERE codartic=" & DBSet(RS!codArtic, "T") & " AND codalmac=" & RS!codAlmac
            conn.Execute SQL
        End If

        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    If bol Then
'        'Pasamos la tabla de inventario real sinven al historico: shinve
'        'antes de eliminarla
'        ADonde = "Pasando registros de Inventario al Histórico: shinve."
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
            cantidad = -cantidad 'Sept 2019 la cantidad va en valor absoluto
            vImporte = Abs(vImporte)
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
Dim B As Boolean

        B = True
        '- campo almacen debe tener valor
        If Trim(txtCodigo(13).Text) = "" Then
             MsgBox "El campo Almacen debe tener valor.", vbInformation
             PonerFoco txtCodigo(13)
             B = False
        End If
    
        '- fecha de inventario debe tener valor
        If B Then
            If (OpcionListado = 12 Or OpcionListado = 15 Or OpcionListado = 19) And Trim(txtCodigo(20).Text) = "" Then
                 MsgBox "El campo Fecha debe tener valor.", vbInformation
                 PonerFoco txtCodigo(20)
                 B = False
            End If
        End If
        
        'informe 19: stocks a una fecha
        'la fecha tiene que ser < a fecha hoy
        If OpcionListado = 19 And txtCodigo(20).Text <> "" Then
            If Not CDate(txtCodigo(20).Text) < CDate(Format(Now, "dd/mm/yyyy")) Then
                If vParamAplic.NumeroInstalacion <> vbFenollar Then
                    MsgBox "La fecha stock tiene que ser anterior a la fecha de hoy.", vbInformation
                    PonerFoco txtCodigo(20)
                    B = False
                End If
            End If
        End If
        If OpcionListado = 19 And txtCodigo(103).Text <> "" Then
            If Not IsNumeric(txtCodigo(103).Text) Then
                MsgBox "Campo incremento incorrecto", vbExclamation
                txtCodigo(103).Text = "2"
                PonerFoco txtCodigo(103)
                B = False
            End If
        End If
        If B Then
            If OpcionListado = 16 Then
                If Me.chkValorado.Value = 0 And Me.chkValorDesdeArticulo.Value = 1 Then
                    MsgBox "No esta marcada la opcion de valorar. NO mostrará valoración alguna", vbExclamation
                End If
            End If
        End If
        ValidarCamposInventario = B
End Function



Private Function ListaArtActivos(cadWhere As String, FechaIn As String) As String
Dim RS As ADODB.Recordset
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
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
'        lista = lista & """" & RS.Fields(0).Value & """"
        Lista = Lista & DBSet(RS.Fields(0).Value, "T")
        RS.MoveNext
        If Not RS.EOF Then Lista = Lista & ", "
    Wend
    Lista = Lista & "]"
    ListaArtActivos = Lista
    RS.Close
    Set RS = Nothing
End Function



Private Sub ActualizarImprimir()
Dim i As Long
Dim Desde As Long, Hasta As Long
Dim SQL As String

    Select Case OpcionListado
    Case 7  'TRASPASO ALMACEN
        If frmVisReport.EstaImpreso = True Then
        'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
            If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
            If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
            For i = Desde To Hasta
                SQL = "UPDATE scatra SET situacio=1" 'Impreso
                SQL = SQL & " WHERE codtrasp=" & i
                conn.Execute SQL
            Next i
        End If
        
    Case 8  'MOVIMIENTO ALMACEN
        If frmVisReport.EstaImpreso = True Then
           'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
           If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
           If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
           For i = Desde To Hasta
                SQL = "UPDATE scamov SET situacio=1"
                SQL = SQL & " WHERE codmovim=" & i
                conn.Execute SQL
           Next i
        End If
    End Select
End Sub


Private Sub CargarComboTipoList()
'### Combo Tipo Facturación
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
'### Combo Tipo Facturación
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

    cboSituaAviso.AddItem "En reparación"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 2
    
    cboSituaAviso.AddItem "Pendiente"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 3
    
    cboSituaAviso.AddItem "Cerrado"
    cboSituaAviso.ItemData(cboSituaAviso.NewIndex) = 4

End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
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
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub LlamarImprimir(PonerNombrePDF As Boolean)
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .NombrePDF = ""
        .SeleccionaRPTCodigo = pRptvMultiInforme
        If PonerNombrePDF Then .NombrePDF = pPdfRpt
        If OpcionListado = 513 Then .SoloImprimir = True
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
            cadParam = cadParam & campo & "{sartic.codfamia}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
            Else
                cadParam = cadParam & NomCampo & " totext({sartic.codfamia},""0000"") & " & """ """ & " & {sfamia.nomfamia}" & "|"
            End If
            numParam = numParam + 1
        Case "Marca"
            cadParam = cadParam & campo & "{sartic.codmarca}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
            Else
                cadParam = cadParam & NomCampo & " totext({sartic.codmarca},""0000"") & " & """ """ & " & {smarca.nommarca}" & "|"
            End If
            numParam = numParam + 1
        Case "Proveedor"
            cadParam = cadParam & campo & "{sartic.codprove}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""PROVEEDOR: "" & " & " totext({sartic.codprove},""000000"") & " & """  """ & " & {sprove.nomprove}" & "|"
            Else
                cadParam = cadParam & NomCampo & " totext({sartic.codprove},""000000"") & " & """ """ & " & {sprove.nomprove}" & "|"
            End If
            numParam = numParam + 1
            PonerGrupo = numGrupo
        Case "Tipo Articulo"
            cadParam = cadParam & campo & "{sartic.codtipar}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""TIPO ARTICULO: "" & " & " {sartic.codtipar} & " & """  """ & " & {stipar.nomtipar}" & "|"
            Else
                cadParam = cadParam & NomCampo & " {sartic.codtipar} & " & """ """ & " & {stipar.nomtipar}" & "|"
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



Private Sub AbrirFrmActividades(Optional Indice As Integer)
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
         Set frmMtoClientes = New frmBasico2
            AyudaClientes frmMtoClientes, txtCodigo(indCodigo).Text
            Set frmMtoClientes = Nothing
    
End Sub


Private Function ComprobarFechasConta(Ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim RS As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(Ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set RS = New ADODB.Recordset
        RS.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not RS.EOF Then
            FechaIni = DBLet(RS!FechaIni, "F")
            '## LAURA 19/06/2008
'            FechaFin = DBLet(RS!FechaFin, "F") + 365
'            FechaFin = DateAdd("d", 365, DBLet(RS!FechaFin, "F"))
            FechaFin = DateAdd("yyyy", 1, DBLet(RS!FechaFin, "F"))
            '##
            
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(Ind).Text, FechaFin) Then
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtCodigo(Ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        RS.Close
        Set RS = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function



Private Function ContabilizarFacturas(cadTabla As String, cadWhere As String, ByRef PGB As ProgressBar, ByRef LblPg0 As Label, LblPg1 As Label, DesdeGenerarFraProveedor As Boolean) As Boolean
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim B As Boolean
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
     'contabilidad par ello mirar en la BD de la Conta los parámetros
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
                
                If Val(vUsu.AlmacenPorDefecto2) <> vParamAplic.AlmacenB Then SQL = SQL & " AND scafac.codtipom <> 'FAZ'"
                   
            End If
        End If
        
        If RegistrosAListar(SQL) > 0 Then
            If MsgBox("Hay Facturas anteriores sin contabilizar. " & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
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
    B = CrearTMPFacturas(cadTabla, cadWhere)
    If Not B Then Exit Function
            
            
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
        B = ComprobarLetraSerie(cadTabla)
    End If
    IncrementarProgres PGB, 10
    Me.Refresh
    If Not B Then Exit Function
    
    
    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If cadTabla = "scafac" Then
        LblPg1.Caption = "Comprobando Nº Facturas en contabilidad ..."
        LblPg1.Refresh
        SQL = "anofaccl>=" & Year(txtCodigo(31).Text) & " AND anofaccl<= " & Year(txtCodigo(32).Text)
        B = ComprobarNumFacturas_new(cadTabla, SQL)
    End If
    IncrementarProgres PGB, 20
    Me.Refresh
    If Not B Then Exit Function
    
    
    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: sclien.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    LblPg1.Caption = "Comprobando Cuentas Contables en contabilidad ..."
    LblPg1.Refresh
    B = ComprobarCtaContable_new(cadTabla, 1)
    IncrementarProgres PGB, 20
    Me.Refresh
    If Not B Then Exit Function
    
    
    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    If cadTabla = "scafac" Then
        LblPg1.Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    Else
        LblPg1.Caption = "Comprobando Cuentas Ctbles Compras en contabilidad ..."
    End If
    LblPg1.Refresh
    B = ComprobarCtaContable_new(cadTabla, 2)
    IncrementarProgres PGB, 20
    Me.Refresh
    If Not B Then Exit Function
    
    
    
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
    B = ComprobarTiposIVA(cadTabla)
    IncrementarProgres PGB, 10
    Me.Refresh
    If Not B Then Exit Function
    
    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    
    If vEmpresa.TieneAnalitica Then
       LblPg1.Caption = "Comprobando Contabilidad Analítica ..."
       LblPg1.Refresh
       B = ComprobarCtaContable_new(cadTabla, 3)
       If Not B Then Exit Function
       
       '(si tiene analítica requiere un centro de coste para insertar en conta.linfact)
       B = cadTabla = "scafac"
       If B And OptProve.Tag <> "" Then
        'NUEVO
        'CONTABUILZIACION AGRUPADA DE TIKETS
        
            CCoste2 = ComprobarCCosteTikAgrupado(cadWhere)
       Else
            CCoste2 = ComprobarCCoste(cadWhere, B)
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
        B = ComprobarCtaContable_new(cadTabla, 4)
        If Not B Then Exit Function
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
        'YA no hace la creacion de la AUTOFACTURA
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
    B = NuevasComprobacionesContabilizacion(cadTabla = "scafpc", cadWhere)
    If Not B Then Exit Function
    
    
    
    
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
    SQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    
    '---- Pasar las Facturas a la Contabilidad
    B = PasarFacturasAContab(cadTabla, CCoste2)
    
    
    
    '---- Mostrar ListView de posibles errores (si hay)
    If Not B Then
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
        
            If NumRegistros("Select count(*) from tmpinformes where codusu = " & vUsu.Codigo) > 0 Then
                InicializarVbles
                cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
                numParam = numParam + 1
                
                cadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
                numParam = numParam + 1
                cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
                cadNomRPT = "rContabPRO.rpt"
                conSubRPT = False
                cadTitulo = "Listado contabilizacion FRAPRO"
                
                LlamarImprimir True
            End If
        End If
    End If
    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact
    
    ContabilizarFacturas = True
End Function





'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador
Private Function PasarFacturasAContab(cadTabla As String, miCC As Byte) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim B As Boolean
Dim i As Integer
Dim Numfactu As Integer
Dim Codigo1 As String
Dim ContabilizacionAgrupadaTickets As Boolean

'ENERO 2009
Dim cContaFra As cContabilizarFacturas



    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    
    'Si escontailizacion de facturas de tickets agrupados
    ContabilizacionAgrupadaTickets = False
    If Me.OptProve.Tag <> "" Then ContabilizacionAgrupadaTickets = True
    
    Set RS = New ADODB.Recordset
    
    
    
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
    
    
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Numfactu = RS.Fields(0)
    Else
        Numfactu = 0
    End If
    RS.Close
    Set RS = Nothing


    'Enero 2009
    '------------------------------------------------------------
    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        SQL = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
        SQL = SQL & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        SQL = SQL & Space(50) & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    
    
    
    


    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If Numfactu > 0 Then
    
        Set RS = New ADODB.Recordset
    
        CargarProgres Me.ProgressBarContab, Numfactu
        
        
        'PreComproabacion de los asientos
        If cContaFra.RealizarContabilizacion Then
            SQL = "Select min(fecfactu) from tmpfactu"
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                If Not cContaFra.PreComprobacionNumeroAsiento(RS.Fields(0), Numfactu) Then
                    
                    'Para que la ventana siguiente muestr bien el error
                    SQL = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) VALUES ("
                    SQL = SQL & "'',0,'" & Format(RS.Fields(0), FormatoFecha) & "','Error contadores')"
                    
                    conn.Execute SQL
                    RS.Close
                    Err.Raise 6, , "Comprobacion numeros asiento"
                End If
            End If
            RS.Close
        End If
        
        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        SQL = "SELECT * "
        SQL = SQL & " FROM tmpFactu "
            

        RS.Open SQL, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        B = True
   
   
   
   
   
   
   
        
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not RS.EOF
        
            'Segun sea cli o pro
            If cadTabla = "scafac" Then
                SQL = cadTabla & "." & Codigo1 & "=" & DBSet(RS.Fields(0), "T") & " AND scafac.numfactu=" & RS!Numfactu
                SQL = SQL & " and scafac.fecfactu=" & DBSet(RS!FecFactu, "F")
                If PasarFactura(SQL, miCC, ContabilizacionAgrupadaTickets, cContaFra) = False And B Then B = False
            Else
                SQL = cadTabla & "." & Codigo1 & "=" & DBSet(RS.Fields(0), "N") & " and scafpc.numfactu=" & DBSet(RS!Numfactu, "T")
                SQL = SQL & " and scafpc.fecfactu=" & DBSet(RS!FecFactu, "F")
                If PasarFacturaProv(SQL, miCC, Orden2, cContaFra) = False And B Then B = False
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
            Me.lblProgess2(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & Numfactu & ")"
            Me.Refresh
            i = i + 1
            RS.MoveNext   'Siguiente factura
        Wend
        
        'Veremos si ha dado error la contabilizacion de factiras
        If cContaFra.TieneErrores Then cContaFra.MuestraErroresContabilizacion
        
        
        RS.Close
        Set RS = Nothing
    End If
    
EPasarFac:
    If Err.Number <> 0 Then B = False
    Set cContaFra = Nothing
    If B Then
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
            Codigo = "{smarca.codmarca}"
            Orden1 = "{smarca.codmarca}"
            Orden2 = "{smarca.nommarca}"
            cadTitulo = "Listado Marcas"
            cadNomRPT = "rAlmMarcas.rpt"
            conSubRPT = False
            
        Case 2   'Listado de Almacenes Propios
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado de Almacenes"
            indFrame = 1
            Codigo = "{salmpr.codalmac}"
            Orden1 = "{salmpr.codalmac}"
            Orden2 = "{salmpr.nomalmac}"
            cadTitulo = "Listado Almacenes Propios"
            cadNomRPT = "rAlmAPropios.rpt"
            conSubRPT = False
            
        Case 3   'Listado de Tipos de Unidad
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Unidad"
            indFrame = 1
            Codigo = "{sunida.codunida}"
            Orden1 = "{sunida.codunida}"
            Orden2 = "{sunida.nomunida}"
            cadTitulo = "Listado Tipos de Unidad"
            cadNomRPT = "rAlmTUnidad.rpt"
            conSubRPT = False
            
        Case 4   'Listado de Tipos de Artículos
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Artículos"
            indFrame = 1
            Codigo = "{stipar.codtipar}"
            Orden1 = "{stipar.codtipar}"
            Orden2 = "{stipar.nomtipar}"
            txtCodigo(1).Tag = CadTag
            txtCodigo(2).Tag = CadTag
            cadTitulo = "Listado Tipos de Artículos"
            cadNomRPT = "rAlmTArticulo.rpt"
            conSubRPT = False
            
        Case 6    'Listado de Artículo
            ponerFrameArticulosVisible True, H, W
            CargarListViewOrden
            Codigo = "{sartic"
            indFrame = 11
           
            
            
        Case 110   'Listados Ubicaciones Almacen
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Ubicaciones Almacén"
            indFrame = 1
            Codigo = "{subica.codubica}"
            Orden1 = "{subica.codubica}"
            Orden2 = "{subica.nomubica}"
            cadTitulo = "Listado Ubicaciones Almacén"
            cadNomRPT = "rAlmUbica.rpt"
            conSubRPT = False
            
        Case 18, 247 'Informe Stocks Maximos y Minimos   'OPCION: 247 es este tb
            ponerFrameArticulosVisible True, H, W
            Codigo = "{salmac"
            indFrame = 11
            cmbProduccion.ListIndex = 0
            cmbProduccion.visible = vParamAplic.Produccion
            Label4(90).visible = vParamAplic.Produccion
            
        Case 7, 8 '7: Informe de Traspasos de Almacen
                  '8: Informe de Movimientos de Almacen
            If OpcionListado = 7 Then
                Me.lblTitulo(2).Caption = "Informe Traspaso de Almacén"
                Me.Label2(1).Caption = "Nº Traspaso"
                Codigo = "{scatra.codtrasp}"
            Else
                Me.lblTitulo(2).Caption = "Informe Movimientos de Almacén"
                Me.Label2(1).Caption = "Nº Movimiento"
                Codigo = "{scamov.codmovim}"
            End If
            H = 3495
            W = 5835
            PonerFrameVisible Me.FrameInfAlmacen, True, H, W
            indFrame = 2
            If NumCod <> "" Then
                txtCodigo(3).Text = NumCod
                txtCodigo(4).Text = NumCod
            End If
            
        Case 9 'Informe Movimiento Artículos
            W = 10700
            H = 5775
            PonerFrameVisible Me.FrameMovArtic, True, H, W
            indFrame = 3
            Codigo = "{smoval.codartic}"
            cadTitulo = "Informe Movimientos Artículos"
            conSubRPT = True
            CargarListView
            
        ' ====  [16/09/2009] LAURA : Listado Articulos con componentes
        Case 11
            W = Me.FrameInvArtComp.Width
            H = Me.FrameInvArtComp.Height
            PonerFrameVisible Me.FrameInvArtComp, True, H, W
            Codigo = "{sartic.codartic}"
            cadTitulo = "Listado Artículos con Componentes"
        ' ====
            
        Case 12 '12: Listado Toma de Inventario Articulos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.chkImprimeStock.visible = True
            Me.lbltituloInven.Caption = "Listado Toma de Inventario Artículos"
            cadTitulo = "Toma Inventario Artículos"
            'codigo = "{salmac.codalmac}"
            
        Case 13 '13: Listado Diferencias de Inventario Articulos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Diferencias de Inventario Artículos"
            'codigo = "{sinven.codalmac}"
            cadTitulo = "Diferencias Inventario Artículos"
            
        Case 14 '14: Actualizar Direfencias Inventario (NO IMPRIME INFORME)
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Actualizar Diferencias de Inventario de Artículos"
            Me.Caption = "Inventario de Artículos"
            
        Case 15 '15: Listado de Articulos Inactivos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Artículos Inactivos"
            cadTitulo = "Listado Artículos Inactivos"
    
        Case 16 '16 .- Listado Valoracion de Stocks Inventariados
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoración Stocks Inventariados"
            cadTitulo = "Listado Valoración Stocks Inventariados"
            
        Case 17 '17 .- Listado Valoración Stocks
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoración Stocks"
            cadTitulo = "Listado Valoración Stocks"
            
        Case 19 '19 .- Inf. Stocks a una Fecha
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Informe Stocks a una Fecha"
            cadTitulo = "Stocks a una Fecha"
            
            
            If vParamAplic.NumeroInstalacion = vbFenollar Then
                txtCodigo(13).Text = "1"
                txtNombre(13).Text = PonerNombreDeCod(txtCodigo(13), conAri, "salmpr", "nomalmac", "codalmac", "", "T")
                txtCodigo(20).Text = Format(Now, "dd/mm/yyyy")
            End If
            
            
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
            Codigo = "{sactiv.codactiv}"
            Orden1 = "{sactiv.codactiv}"
            Orden2 = "{sactiv.nomactiv}"
            cadTitulo = "Listado Actividades de Clientes"
            cadNomRPT = "rFacActividades.rpt"
            
        Case 21    'Listado de Zonas de Clientes
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Zonas de Clientes"
            indFrame = 1
            Codigo = "{szonas.codzonas}"
            Orden1 = "{szonas.codzonas}"
            Orden2 = "{szonas.nomzonas}"
            cadTitulo = "Listado Zonas de Clientes"
            cadNomRPT = "rFacZonas.rpt"
        
        Case 22    'Listado de Rutas de Asistencia
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Rutas de Asistencia"
            indFrame = 1
            Codigo = "{srutas.codrutas}"
            Orden1 = "{srutas.codrutas}"
            Orden2 = "{srutas.nomrutas}"
            cadTitulo = "Listado Rutas de Asistencia"
            cadNomRPT = "rFacRutas.rpt"
            
        Case 23     'Listado de Tipos de Formas de Envío
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Formas de Envío"
            indFrame = 1
            Codigo = "{senvio.codenvio}"
            Orden1 = "{senvio.codenvio}"
            Orden2 = "{senvio.nomenvio}"
            cadTitulo = "Listado Formas de Envio"
            cadNomRPT = "rFacEnvio.rpt"
            
        Case 24    'Tarifas Venta
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tarifas Venta"
            indFrame = 1
            Codigo = "{starif.codlista}"
            Orden1 = "{starif.codlista}"
            Orden2 = "{starif.nomlista}"
            cadTitulo = "Listado Tarifas Venta"
            cadNomRPT = "rFacTarifasVen.rpt"
            
        Case 27     'Situaciones Especiales
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Situaciones Especiales"
            indFrame = 1
            Codigo = "{ssitua.codsitua}"
            Orden1 = "{ssitua.codsitua}"
            Orden2 = "{ssitua.nomsitua}"
            cadTitulo = "Listado Situaciones Especiales"
            cadNomRPT = "rFacSituaciones.rpt"
            
        Case 28    '28: Informe de Tarifas de Precios
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Tarifas de Artículos"
            Codigo = "{slista"
            indFrame = 5
            cadTitulo = "Listado Tarifas Articulos"
            
        Case 29  '29: Informe Promociones
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Promociones Tarifas"
            Codigo = "{spromo"
            indFrame = 5
            cadTitulo = "Listado Promociones de Tarifas"
            FrameFechasPromo.Top = 6000
            FrameFechasPromo.Left = 240
            FrameFechasPromo.BorderStyle = 0
        Case 30 '30: Informe Precios Especiales
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Precios Especiales Artículos"
            Codigo = "{sprees"
            indFrame = 5
            cadTitulo = "Listado Precios Especiales"
            
        Case 245, 247 '245: Informe control margenes tarifas
            indFrame = 5
            PonerFrameTarifasVisible True, H, W
            Me.lblTituloTarif.Caption = "Informe Control Margenes de Tarifas"
            Codigo = "{slista"
            cadTitulo = "Listado Control Margenes Tarifas"
            cboDecimales.ListIndex = 4
        Case 246 '246: Informe margen ventas x articulo
            indFrame = 15
            H = 5300
            W = 7820
            PonerFrameVisible Me.FrameEstMargenes, True, H, W
            cadTitulo = "Listado Margen ventas por artículo"
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
        Case 407 'Sustitución Num. serie
            H = 3700
            W = 5720
            PonerFrameVisible Me.FrameRepSustNSerie, True, H, W
            Me.lblNumSerie(0).Caption = "Nº Serie:   " & NumCod
            Me.lblNumSerie(1).Caption = "Artículo:   " & Me.CadTag
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
Dim i As Integer
    For i = 2 To 6
        Me.txtBultos(i).Text = ""
        Me.txtBultos(i).Tag = ""
    Next i
End Sub



Private Sub PonerCamposDireccionBultos(Indice As Integer)
Dim i As Integer

    'El indice mara el listindex del combo, por lo tanto sera indice + 1
    For i = 2 To 6
        Me.txtBultos(i).Text = RecuperaValor(Me.txtBultos(i).Tag, Indice + 1)
    Next i
End Sub


Private Sub PonerCamposAlbaran()
'Informe Etiquetas Bultos
'si en NumCod se ha pasado el nº de un Albaran cargar por defectos valores
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo ErrAlb
    
    '1) -- Buscar en la tabla de ALBARANES: PED -> ALV
    SQL = "SELECT codclien,coddirec, sum(l.numbultos) as totBultos"
    SQL = SQL & " FROM scaalb c INNER JOIN slialb l ON c.numalbar=l.numalbar and c.codtipom=l.codtipom"
    SQL = SQL & " WHERE c.numalbar=" & NumCod & " and c.codtipom='ALV'"
    SQL = SQL & " GROUP by c.numalbar,c.codtipom"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        txtClie.Text = RS!codClien
    
        CadTag = DBLet(RS!CodDirec, "T")
        
        txtBultos(1).Text = DBLet(RS!totbultos, "N")
        
        txtClie_LostFocus
    End If
    
    RS.Close
    Set RS = Nothing
    
    '2) Buscar en la tabla de FACTURAS PED -> FAV
    If txtClie.Text = "" Then
         'Comprobar en FACTURAS: x si se pasa de PED -> FAC
        SQL = "SELECT codclien,coddirec, sum(l.numbultos) as totBultos "
        SQL = SQL & " FROM (scafac c INNER JOIN scafac1 a ON c.numfactu=a.numfactu and c.codtipom=a.codtipom and c.fecfactu=a.fecfactu)"
        SQL = SQL & " INNER JOIN slifac l ON a.numfactu=l.numfactu and a.codtipom=l.codtipom and a.fecfactu=l.fecfactu and a.numalbar=l.numalbar and a.codtipoa=l.codtipoa"
        SQL = SQL & " WHERE a.numalbar=" & NumCod & " and a.codtipoa='ALV'"
        SQL = SQL & " GROUP BY a.numalbar,a.codtipoa"
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        If Not RS.EOF Then
            txtClie.Text = RS!codClien
        
            CadTag = DBLet(RS!CodDirec, "T")
            
            txtBultos(1).Text = DBLet(RS!totbultos, "N")
            
            txtClie_LostFocus
        End If
        RS.Close
        Set RS = Nothing
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
    Codigo = "select min(fecfactu) from scafac"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    F2 = DateAdd("yyyy", -5, CDate("01/01/" & Year(Now)))

    Codigo = F2
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then Codigo = miRsAux.Fields(0)
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Codigo = "31/12/" & Year(CDate(Codigo))
    
    While CDate(Codigo) < F2
        
        cmbEliFac.AddItem "     " & Format(CDate(Codigo), "dd/mm/yyyy")
        Codigo = CStr(DateAdd("yyyy", 1, CDate(Codigo)))
    
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
    Codigo = "Select count(*) from scafac where fecfactu<='" & Format(FechaBorre, FormatoFecha) & "' and intconta = 0"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Existen " & NumRegElim & " facturas sin contabilizar en esas fechas", vbExclamation
        Exit Function
    End If
    
        
    'lo mismo para proeedores
    Codigo = "Select count(*) from scafpc where fecfactu<='" & Format(FechaBorre, FormatoFecha) & "' and intconta = 0"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    Codigo = "slifac|scafac1|svenci|srecom|scafac|"
    For NumRegElim = 1 To 5
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Tabla CLI: " & Orden1
        Orden1 = "DELETE FROM " & Orden1 & " WHERE  fecfactu<='" & Format(FechaBorre, FormatoFecha) & "'"
        Me.Refresh
        DoEvents
        conn.Execute Orden1
    Next NumRegElim
    
    '---------------------------------------------------------------------------------
    'Albarananes CLIENTES.
    '--
    Codigo = "scaalb|schalb|slialb|slhalb|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE codtipom = '"
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!codtipom & "'  AND numalbar = " & miRsAux!Numalbar
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Borramos las cabceeras
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
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
    Codigo = "scaped|schped|sliped|slhped|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedcl<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numpedcl = "
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!NumPedcl
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Cabce
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
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
    Codigo = "scapre|schpre|slipre|slhpre|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecofert<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numofert = "
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!NumOfert
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Ofertas(C): " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1
        Orden1 = Orden1 & " WHERE  fecofert <='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    DoEvents
    
    
    Codigo = "scarep|schrep|slirep|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Reparaciones: " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecrepar<='" & Format(FechaBorre, FormatoFecha) & "'"
        If NumRegElim = 1 Then
            'Lineas de reparacion solo hay en scarep
            'En shrep no hay lineas
            miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
            Orden1 = "DELETE FROM " & Orden1 & " WHERE numrepar = "
            While Not miRsAux.EOF
                conn.Execute Orden1 & miRsAux!numrepar
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
        End If
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
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
        conn.Execute Orden1 & miRsAux!Codigo
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
    Codigo = "slifpc|scafpa|scafpc|"
    For NumRegElim = 1 To 3
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Tabla PRO: " & Orden1
        Orden1 = "DELETE FROM " & Orden1 & " WHERE  fecfactu<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
    Next NumRegElim
    
    
    
    
    Codigo = "slhalp|slialp|scaalp|schalp|"
    For NumRegElim = 1 To 4
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Albaranes prov: " & Orden1
        Label3(83).Refresh
        Orden1 = "DELETE from " & Orden1 & " WHERE  fechaalb<='" & Format(FechaBorre, FormatoFecha) & "'"
        conn.Execute Orden1
        
    Next
    DoEvents
    
    
    
    
    '-----------------------------------------------
    'Pedidos proveedor
    '--
    Codigo = "scappr|schppr|slippr|slhppr|"
    For NumRegElim = 1 To 2
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
        Label3(83).Caption = "Pedidos prov.(L): " & Orden1
        Label3(83).Refresh
        Orden1 = "Select * from " & Orden1
        Orden1 = Orden1 & " WHERE  fecpedpr<='" & Format(FechaBorre, FormatoFecha) & "'"
        miRsAux.Open Orden1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Orden1 = RecuperaValor(Codigo, NumRegElim + 2)
        Orden1 = "DELETE FROM " & Orden1 & " WHERE numpedpr = "
        While Not miRsAux.EOF
            conn.Execute Orden1 & miRsAux!numpedpr
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        Orden1 = RecuperaValor(Codigo, CInt(NumRegElim))
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
    
    cadSelect = "Select * from tmpnlotes where codusu =" & vUsu.Codigo & " ORDER BY codalmac,numalbar,codprove"
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
            .OtrosParametros = cadParam
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
        Codigo = "UPDATE scaalb set factursn = 1 "
        If NumCod <> "" Then cadSelect = " codtipom ='" & NumCod & "'"
        
        cadParam = "fechaalb"
        cadFormula = CadenaDesdeHastaBD(txtCodigo(117).Text, txtCodigo(118).Text, "codclien", "N")
        If cadFormula <> "" Then
            If cadSelect <> "" Then cadSelect = cadSelect & " AND "
            cadSelect = cadSelect & cadFormula
        End If
        

    Else
        'Hacer borrar avisos
        Codigo = "DELETE FROM scaavi"
        cadSelect = " situacio = 3"
        cadParam = "fechaavi"
    End If
    
    cadFormula = CadenaDesdeHastaBD(txtCodigo(119).Text, txtCodigo(120).Text, cadParam, "F")
    If cadFormula <> "" Then
        If cadSelect <> "" Then cadSelect = cadSelect & " AND "
        cadSelect = cadSelect & cadFormula
    End If
    
    If cadSelect <> "" Then cadSelect = " WHERE " & cadSelect
    Codigo = Codigo & cadSelect
    conn.Execute Codigo
    
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
                            C = "Factura contabilizada con fecha de recepción menor que ya existentes en contabilidad."
                            C = C & vbCrLf & vbCrLf & "¿Continuar?"
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
    s = "DELETE FROM tmpsliped where codusu = " & vUsu.Codigo
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
        s = s & ", (" & vUsu.Codigo & ",1," & J & "," & R!codAlmac & "," & DBSet(R!codArtic, "T")
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
        s = s & ", (" & vUsu.Codigo & ",2," & J & "," & R!codAlmac & "," & DBSet(R!codArtic, "T")
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

Private Sub CargaIconosAyuda2()
Dim i As Integer
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    
    For i = 10 To 28
        imgBuscarG(i).Picture = imgBuscar(0).Picture
    Next i
    
    imgBuscarG(87).Picture = imgBuscar(0).Picture
    
    imgBuscarG(101).Picture = imgBuscar(0).Picture
    imgBuscarG(102).Picture = imgBuscar(0).Picture
    
    imgBuscarG(110).Picture = imgBuscar(0).Picture
    imgBuscarG(111).Picture = imgBuscar(0).Picture
    
    
    Err.Clear
End Sub



Private Function HacerInfrStockMinimo() As Boolean

    On Error GoTo eHacerInfrStockMinimo
    
    Codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    
    Codigo = " FROM sartic,salmac WHERE sartic.codartic=salmac.codartic AND ctrstock=1 AND artvario  =0  "
    'El desde hasta
    If cadSelect <> "" Then
        cadTitulo = Replace(cadSelect, "{", "")
        cadTitulo = Replace(cadTitulo, "}", "")
        Codigo = Codigo & " AND " & cadTitulo
    End If
    
    If Me.chkVarios(0).Value = 0 Then
        'Normal. Listado de stcok minimos
        Codigo = Codigo & " AND stockmin>0"
        
        
    Else
        'Los que no tiene minimo y tienen stock
        Codigo = Codigo & " AND COALESCE(stockmin,0)<=0 and canstock>0"
        
    End If
    Codigo = "SELECT " & vUsu.Codigo & ",codprove,codfamia,codalmac,sartic.codartic,nomartic,stockmin,stockmax,puntoped,if(canstock<0,0,canstock) stock " & Codigo
    
    'codusu,campo2,codigo1,campo1,nombre1,nombre2,
    'codalmac,stockmin,stockmax,puntoped,canstock,numorden
    Codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,importe1,importe2,importe3,importe4) " & Codigo
    
    Label3(116).Caption = "Insertando en BD"
    Label3(116).Refresh
    conn.Execute Codigo
    DoEvents
    
    'Actualizamos la familia
    Set miRsAux = New ADODB.Recordset
    Codigo = "Select campo1 from tmpinformes WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
    Set miRsAux = New ADODB.Recordset
    Label3(116).Caption = "Leer familias"
    Label3(116).Refresh
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Codigo = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", miRsAux!campo1)
        Label3(116).Caption = Codigo
        Label3(116).Refresh
        Codigo = "UPDATE tmpinformes SET nombre3=" & DBSet(Codigo, "T") & " WHERE codusu =" & vUsu.Codigo & " AND campo1 = " & miRsAux!campo1
        conn.Execute Codigo
        miRsAux.MoveNext
        If (NumRegElim Mod 10) = 0 Then DoEvents
    Wend
    miRsAux.Close
   
   
    'marzo 2014
    'Añadir pedidos clientes de ese almacen
    
    If NumRegElim > 0 Then
        Label3(116).Caption = "Pedidos clientes"
        Label3(116).Refresh
        DoEvents
        
        Codigo = "select codalmac,codartic,sum(cantidad) cuantos from sliped where (codalmac,codartic) IN"
        Codigo = Codigo & "(select campo2,nombre1 from tmpinformes where codusu=" & vUsu.Codigo & " ) group by 1,2 ORDER BY 1,2"
        miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Label3(116).Caption = miRsAux!codAlmac & " " & miRsAux!codArtic
            Label3(116).Refresh
            
            
            Codigo = "UPDATE tmpinformes SET importe5=" & DBSet(DBLet(miRsAux!Cuantos, "N"), "N")
            Codigo = Codigo & " WHERE codusu =" & vUsu.Codigo & " AND campo2 = " & miRsAux!codAlmac
            Codigo = Codigo & " AND nombre1 = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute Codigo
   
   
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
Dim RS As ADODB.Recordset

    Label4(103).Caption = "Preparando datos"
    Label4(103).Refresh
    BorrarTempInformes
        
    Codigo = "SELECT slispr.*,nomartic,codfamia,codmarca from " & Orden1
    If cadSelect <> "" Then Codigo = Codigo & " WHERE " & cadSelect
    Codigo = Codigo & " ORDER BY codprove,codfamia,codmarca"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    Set RS = New ADODB.Recordset
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
            
            If cadSelect <> "" Then RS.Close
            RS.Open Orden1, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            cadSelect = Orden2
        End If
        
        '                                                                           rap1  rap2
        'codusu,codigo1,campo1 ,campo2,nombre1,nombre2,importe1,porcen1,porcen2,importe4,importe5
        '1er trozo del insert
        Codigo = Codigo & ", (" & vUsu.Codigo & "," & NumRegElim & "," & miRsAux!Codprove & "," & miRsAux!Codfamia
        Codigo = Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!NomArtic, "T") & ","
   
        cadTitulo = miRsAux!precioac
        If Not IsNull(miRsAux!fechanue) Then
            If miRsAux!fechanue <= Now Then cadTitulo = DBLet(miRsAux!precionu, "N")
        End If
        Codigo = Codigo & TransformaComasPuntos(cadTitulo) & ","
   
   
        
        If Not RS.EOF Then
            If miRsAux!dtopermi = 0 Then
                Codigo = Codigo & "0,0,"
            Else
                Codigo = Codigo & TransformaComasPuntos(CStr(DBLet(RS!dtoline1, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(RS!dtoline2, "N"))) & ","
            End If
            Codigo = Codigo & TransformaComasPuntos(CStr(DBLet(RS!Rap1, "N"))) & "," & TransformaComasPuntos(CStr(DBLet(RS!Rap2, "N"))) & ")"
            
        Else
            Codigo = Codigo & "0,0,0,0)"
        End If
       

        
        

        
        If (NumRegElim Mod 20) = 0 Then InsertaEnTmpHazCalculo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Set RS = Nothing
    If Codigo <> "" Then InsertaEnTmpHazCalculo
    

    Label4(103).Caption = "Calculando neto"
    Label4(103).Refresh
    'Seguntipo dto
    If vParamAplic.TipoDtos = 1 Then
        Codigo = "UPDATE tmpinformes Set importeb3=importeb1 * ((100 - porcen1) / 100) WHERE codusu =" & vUsu.Codigo
        conn.Execute Codigo
        Espera 0.25
        Codigo = "UPDATE tmpinformes Set importeb2=importeb3 * ((100 - porcen2) / 100) WHERE codusu =" & vUsu.Codigo
        conn.Execute Codigo
    Else
        Codigo = "UPDATE tmpinformes Set importeb2=importeb1 * ((100 - (porcen1+porcen2)) / 100) WHERE codusu =" & vUsu.Codigo
        conn.Execute Codigo
    End If
    Label4(103).Caption = ""
End Sub


Private Sub InsertaEnTmpHazCalculo()
    Codigo = Mid(Codigo, 2)
    Codigo = "" & Codigo
    'codusu,codigo1,campo1 ,campo2,nombre1,nombre2,importe1,porcen1,porcen2
    Codigo = "INSERT INTO tmpinformes(codusu,codigo1,campo1 ,campo2,nombre1,nombre2,importeb1,porcen1,porcen2,importe4,importe5) VALUES " & Codigo
    conn.Execute Codigo
    Codigo = ""
End Sub



Private Sub HacerListadoDtosCliente()



    lblDtoAct.visible = True
    lblDtoAct.Caption = "Prepara"
    lblDtoAct.Refresh
    'Vaciamos
    conn.Execute "DELETE FROM tmpinformes WHERE codusu =" & vUsu.Codigo
    
    
    'Solo tendra en cuenta el desde hasta cliente.
    'Sacara todos los dtos propios y los de su actividad(si es que estan dados de alta)
    Codigo = "SELECT codclien,codactiv FROM sclien WHERE codclien>=0 "
    If txtCodigo(73).Text <> "" Then Codigo = Codigo & " AND codclien >=" & txtCodigo(73).Text
    If txtCodigo(74).Text <> "" Then Codigo = Codigo & " AND codclien <=" & txtCodigo(74).Text
    
    
    
    If txtCodigo(158).Text <> "" Then Codigo = Codigo & " AND codagent >=" & txtCodigo(158).Text
    If txtCodigo(159).Text <> "" Then Codigo = Codigo & " AND codagent <=" & txtCodigo(159).Text
    
    
    
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 0
    miRsAux.Open Codigo, conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        miRsAux.MoveNext
    Wend
 
    
    If NumRegElim = 0 Then
        MsgBox "No existe datos para mostrar", vbExclamation
    Else
        If NumRegElim > 4 Then
            If MsgBox("El proceso puede llevar mucho tiempo. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    If NumRegElim > 0 Then
        DoEvents
        miRsAux.MoveFirst
        
        While Not miRsAux.EOF
            lblDtoAct.Caption = "Cl " & miRsAux!codClien
            lblDtoAct.Refresh
            Codigo = "insert into tmpinformes(codusu,codigo1,campo1,campo2,importe1,importe2,fecha1,porcen1)"
            Codigo = Codigo & " select " & vUsu.Codigo & ",codclien,codfamia,codmarca,dtoline1,dtoline2,fechadto,0"
            Codigo = Codigo & " from sdtofm where codclien=" & miRsAux!codClien
            conn.Execute Codigo
            Espera 0.2
    
    
            'Los que vienen de descuento
            Codigo = "insert into tmpinformes(codusu,codigo1,campo1,campo2,importe1,importe2,fecha1,porcen1)"
            Codigo = Codigo & " select " & vUsu.Codigo & "," & miRsAux!codClien & ",codfamia,null,dtoline1,"
            Codigo = Codigo & " dtoline2,fechadto,1 from sdtofm where codactiv=" & miRsAux!codactiv & " and not codfamia in ("
            Codigo = Codigo & " select campo1 from tmpinformes where codusu =" & vUsu.Codigo & " and codigo1=" & miRsAux!codClien & " and campo2 is null)"
            conn.Execute Codigo
            
            miRsAux.MoveNext
        
        Wend
    End If
    miRsAux.Close
    
    'Si tiene alguno de MARCA
     lblDtoAct.Caption = "Marca"
    lblDtoAct.Refresh
    Codigo = "Select campo2 from tmpinformes WHERE codusu =" & vUsu.Codigo & " AND campo2>=0 GROUP BY 1"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Codigo = DevuelveDesdeBD(conAri, "nommarca", "smarca", "codmarca", miRsAux.Fields(0))
        Codigo = "UPDATE tmpinformes SET nombre2=" & DBSet(Codigo, "T") & " WHERE codusu =" & vUsu.Codigo & " AND campo2 = " & miRsAux.Fields(0)
        conn.Execute Codigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
  
    
     lblDtoAct.Caption = "Activ"
    lblDtoAct.Refresh
    Codigo = "UPDATE tmpinformes SET nombre3='Activ.' WHERE codusu =" & vUsu.Codigo & " AND porcen1>=1"
    conn.Execute Codigo
    
    
    
    
    
    
    'Enero 2017
    'Lincaremos con sdtomp , para ello veremos para cada familia porveedor y marca (marca puede ser NULL seran dos procesos
    
    For NumRegElim = 1 To 2
        lblDtoAct.Caption = "Prov (" & NumRegElim & ")"
        lblDtoAct.Refresh
        Codigo = "select * from sdtomp where "
        If NumRegElim = 1 Then
            Codigo = Codigo & " codmarca is null "
        Else
            Codigo = Codigo & " codmarca>=0 "
        End If
        Codigo = Codigo & " AND (codprove,codfamia) in "
        Codigo = Codigo & " (select distinct codfamia,codprove from sfamia ,tmpinformes where "
        If NumRegElim = 1 Then
            Codigo = Codigo & " codmarca is null "
        Else
            Codigo = Codigo & " codmarca>=0 "
        End If
        Codigo = Codigo & " and codusu =" & vUsu.Codigo & " and tmpinformes.campo1=sfamia.codfamia )"
        
        miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            lblDtoAct.Caption = "Fam: " & miRsAux!Codfamia
            lblDtoAct.Refresh
            'importeb1,importeb2,importeb3,importeb4  << dtoline1,dtoline2,rap1,rap2
            Codigo = "UPDATE tmpinformes set importeb1=" & DBSet(miRsAux!dtoline1, "N", "S")
            Codigo = Codigo & ", importeb2 =" & DBSet(miRsAux!dtoline2, "N", "S")
            Codigo = Codigo & ", importeb3 =" & DBSet(miRsAux!Rap1, "N", "S")
            Codigo = Codigo & ", importeb4 =" & DBSet(miRsAux!Rap2, "N", "S")
            Codigo = Codigo & " WHERE codusu =" & vUsu.Codigo
            Codigo = Codigo & " AND campo1=" & miRsAux!Codfamia
            If NumRegElim = 1 Then
            Codigo = Codigo & " AND campo2 is null "
            Else
                Codigo = Codigo & " AND campo2 >=0 "
            End If
            conn.Execute Codigo
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    
    Next
    Set miRsAux = Nothing
    
    
    cadTitulo = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.Codigo))
    If Val(cadTitulo) > 0 Then
    
        cadTitulo = "Descuento cliente / actividad"
        cadFormula = "({tmpinformes.codusu} = " & vUsu.Codigo & ")"
        cadNomRPT = "rFacDtoCliACtiv.rpt"
        
        LlamarImprimir False
    End If
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub PonerDatosFacturaProveedorAcabadaRecepcionar()
          'Si genera la factura y la contabiliza (UNA A UNA),
            numParam = CByte(vbCritical)
            cadFormula = "importe1"
            cadParam = DevuelveDesdeBD(conAri, "codigo1", "tmpinformes", "codusu", vUsu.Codigo, "N", cadFormula)
             Orden1 = ""
            If cadParam <> "" Then
                If cadParam = "0" Then
                    cadParam = ""
                    Orden1 = ""
                Else
                    Orden1 = DevuelveDesdeBD(conAri, "importeb5", "tmpinformes", "codusu", vUsu.Codigo, "N")
                    'If Orden1 <> "" And Val(Orden1) <> 0 Then
                    '    Orden1 = vbCrLf & "Asiento: " & Val(Orden1)
                    'Else
                    '    Orden1 = ""
                    'End If
                    
                    
                End If
                
            End If
            
            Me.txtImporte(5).Text = cadParam
            txtImporte(6).Text = Orden1
          
           
            cadParam = "Numero de registro: " & cadParam & vbCrLf
                If cadFormula <> "importe1" Then
                    If cadFormula = "" Then cadFormula = "0" 'para que no de error Ene 2021
                    cadFormula = DevuelveDesdeBD(conAri, "codmacta", "sprove", "codprove", CStr(Val(CCur(cadFormula))))
                    Orden1 = ""
                Else
                    cadFormula = ""
                End If
                
                If cadFormula = "" Then
                    Orden1 = "Error en cuenta contable proveedor"
                Else
                        
                
                    Orden1 = DevuelveDesdeBD(conAri, "nombre1", "tmpinformes", "codusu", vUsu.Codigo, "N")
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
                                Me.txtImporte(indCodigo).Text = Format(miRsAux!impefect, FormatoImporte)
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
        Me.txtImporte(indCodigo).visible = visible
        Label3(120 + indCodigo).visible = visible
        Me.txtCodigo(149 + indCodigo).visible = visible
        imgFecha(21 + indCodigo).visible = visible
End Sub
