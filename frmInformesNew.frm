VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInformesNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe "
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   12300
   Icon            =   "frmInformesNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   7035
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   12090
      Begin VB.Frame FrameInfArticulosOrd 
         Caption         =   "Ordenación"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2490
         Left            =   7335
         TabIndex        =   128
         Top             =   3105
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CommandButton cmdSubir 
            Caption         =   " "
            Height          =   650
            Left            =   3105
            Picture         =   "frmInformesNew.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   131
            Top             =   585
            Width           =   510
         End
         Begin VB.CommandButton cmdBajar 
            Height          =   650
            Left            =   3105
            Picture         =   "frmInformesNew.frx":108E
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   1350
            Width           =   510
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1485
            Left            =   315
            TabIndex        =   130
            Top             =   540
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   2619
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
      End
      Begin VB.Frame FrameInfArticulosOpc 
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5385
         Left            =   7335
         TabIndex        =   91
         Top             =   225
         Visible         =   0   'False
         Width           =   4455
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
            Height          =   465
            Left            =   450
            TabIndex        =   98
            Top             =   2700
            Width           =   3945
         End
         Begin VB.CheckBox chkImpEtiq 
            Caption         =   "Etiquetas"
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
            Index           =   0
            Left            =   450
            TabIndex        =   97
            Top             =   3240
            Width           =   1320
         End
         Begin VB.CheckBox chkImpEtiq 
            Caption         =   "P.V.P."
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
            Index           =   1
            Left            =   2295
            TabIndex        =   96
            Top             =   3240
            Width           =   2055
         End
         Begin VB.CheckBox chkImpEtiq 
            Caption         =   "Rotación"
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
            Index           =   2
            Left            =   2295
            TabIndex        =   95
            Top             =   3600
            Width           =   1275
         End
         Begin VB.CheckBox chkImpEtiq 
            Caption         =   "Precio mínimo"
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
            Index           =   3
            Left            =   450
            TabIndex        =   94
            Top             =   3600
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.Frame FrameSituacionArticulo 
            Caption         =   "Situación artículo"
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
            Height          =   1275
            Left            =   315
            TabIndex        =   93
            Top             =   765
            Width           =   3705
            Begin VB.CheckBox chkSitaucionArticulo2 
               Caption         =   "Normal"
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
               Index           =   0
               Left            =   210
               TabIndex        =   114
               Top             =   360
               Value           =   1  'Checked
               Width           =   1680
            End
            Begin VB.CheckBox chkSitaucionArticulo2 
               Caption         =   "Bloqueado"
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
               Index           =   2
               Left            =   210
               TabIndex        =   116
               Top             =   720
               Value           =   1  'Checked
               Width           =   1680
            End
            Begin VB.CheckBox chkSitaucionArticulo2 
               Caption         =   "Caducado"
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
               Index           =   3
               Left            =   1995
               TabIndex        =   120
               Top             =   720
               Value           =   1  'Checked
               Width           =   1680
            End
            Begin VB.CheckBox chkSitaucionArticulo2 
               Caption         =   "Obsoleto"
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
               Index           =   1
               Left            =   1995
               TabIndex        =   118
               Top             =   360
               Value           =   1  'Checked
               Width           =   1680
            End
         End
         Begin VB.Frame FrameStockMaxMin 
            Caption         =   "Imprimir Stocks"
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
            Height          =   1380
            Left            =   315
            TabIndex        =   115
            Top             =   675
            Width           =   3705
            Begin VB.OptionButton optStockMax 
               Caption         =   "Máximos"
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
               Left            =   225
               TabIndex        =   121
               Top             =   315
               Width           =   1155
            End
            Begin VB.OptionButton optStockMin 
               Caption         =   "Mínimos"
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
               Left            =   225
               TabIndex        =   119
               Top             =   652
               Width           =   1110
            End
            Begin VB.OptionButton optPuntoPedido 
               Caption         =   "Punto de pedido"
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
               Left            =   225
               TabIndex        =   117
               Top             =   990
               Width           =   1980
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
            ItemData        =   "frmInformesNew.frx":2110
            Left            =   1935
            List            =   "frmInformesNew.frx":211D
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   1170
            Width           =   2415
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
            ItemData        =   "frmInformesNew.frx":213C
            Left            =   1935
            List            =   "frmInformesNew.frx":2146
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   1530
            Width           =   2415
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
            Left            =   315
            TabIndex        =   125
            Top             =   1575
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
            Left            =   315
            TabIndex        =   123
            Top             =   1215
            Width           =   1125
         End
         Begin VB.Image imgayuda 
            Height          =   240
            Index           =   0
            Left            =   4095
            ToolTipText     =   "Informes artículos"
            Top             =   3555
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "&Imprimir"
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
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Top             =   5760
         Width           =   1335
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
         Left            =   10440
         TabIndex        =   3
         Top             =   5805
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
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
         Index           =   1
         Left            =   8955
         TabIndex        =   23
         Top             =   5805
         Width           =   1335
      End
      Begin VB.Frame FrameInfArticulosSel 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6780
         Left            =   225
         TabIndex        =   65
         Top             =   225
         Visible         =   0   'False
         Width           =   6960
         Begin VB.Frame FrameTapaINCORRECTO 
            BorderStyle     =   0  'None
            Height          =   420
            Left            =   855
            TabIndex        =   99
            Top             =   495
            Width           =   5970
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
               TabIndex        =   101
               Top             =   45
               Width           =   615
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
               Index           =   107
               Left            =   990
               Locked          =   -1  'True
               TabIndex        =   100
               Text            =   "Text5"
               Top             =   45
               Width           =   4905
            End
            Begin VB.Image imgBuscarG 
               Height          =   240
               Index           =   87
               Left            =   90
               ToolTipText     =   "Buscar almacen"
               Top             =   60
               Width           =   240
            End
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   113
            Text            =   "Text5"
            Top             =   540
            Width           =   4860
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
            Left            =   1275
            MaxLength       =   4
            TabIndex        =   102
            Top             =   540
            Width           =   615
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
            Left            =   1230
            MaxLength       =   16
            TabIndex        =   111
            Top             =   5625
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
            Index           =   71
            Left            =   1230
            MaxLength       =   16
            TabIndex        =   112
            Top             =   6030
            Width           =   2070
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
            Left            =   3345
            Locked          =   -1  'True
            TabIndex        =   90
            Text            =   "Text5"
            Top             =   5625
            Width           =   3450
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
            Left            =   3345
            Locked          =   -1  'True
            TabIndex        =   89
            Text            =   "Text5"
            Top             =   6030
            Width           =   3450
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
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   85
            Text            =   "Text5"
            Top             =   3825
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
            Index           =   66
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   84
            Text            =   "Text5"
            Top             =   3420
            Width           =   4620
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
            Left            =   1230
            MaxLength       =   6
            TabIndex        =   108
            Top             =   3825
            Width           =   870
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
            Left            =   1230
            MaxLength       =   6
            TabIndex        =   107
            Text            =   "000000"
            Top             =   3420
            Width           =   870
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
            Left            =   1230
            MaxLength       =   4
            TabIndex        =   103
            Top             =   1170
            Width           =   615
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
            Left            =   1230
            MaxLength       =   4
            TabIndex        =   104
            Top             =   1575
            Width           =   615
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
            Left            =   1875
            Locked          =   -1  'True
            TabIndex        =   83
            Text            =   "Text5"
            Top             =   1170
            Width           =   4905
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
            Left            =   1875
            Locked          =   -1  'True
            TabIndex        =   82
            Text            =   "Text5"
            Top             =   1575
            Width           =   4905
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
            Left            =   1230
            MaxLength       =   4
            TabIndex        =   105
            Top             =   2325
            Width           =   615
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
            Left            =   1230
            MaxLength       =   4
            TabIndex        =   106
            Top             =   2730
            Width           =   615
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
            Left            =   1875
            Locked          =   -1  'True
            TabIndex        =   81
            Text            =   "Text5"
            Top             =   2325
            Width           =   4905
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
            Left            =   1875
            Locked          =   -1  'True
            TabIndex        =   80
            Text            =   "Text5"
            Top             =   2730
            Width           =   4905
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
            Left            =   1230
            MaxLength       =   8
            TabIndex        =   109
            Top             =   4560
            Width           =   1065
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
            Left            =   1230
            MaxLength       =   8
            TabIndex        =   110
            Text            =   "00000000"
            Top             =   4965
            Width           =   1065
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
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   79
            Text            =   "Text5"
            Top             =   4560
            Width           =   4425
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
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   78
            Text            =   "Text5"
            Top             =   4965
            Width           =   4425
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   18
            Left            =   990
            Top             =   540
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Height          =   285
            Index           =   36
            Left            =   180
            TabIndex        =   92
            Top             =   270
            Width           =   915
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   27
            Left            =   945
            ToolTipText     =   "Buscar artículo"
            Top             =   5625
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   225
            Index           =   28
            Left            =   945
            ToolTipText     =   "Buscar artículo"
            Top             =   6030
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
            Left            =   180
            TabIndex        =   88
            Top             =   5265
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
            Left            =   225
            TabIndex        =   87
            Top             =   5970
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
            Index           =   51
            Left            =   225
            TabIndex        =   86
            Top             =   5610
            Width           =   735
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   24
            Left            =   945
            ToolTipText     =   "Buscar proveedor"
            Top             =   3855
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   23
            Left            =   945
            ToolTipText     =   "Buscar proveedor"
            Top             =   3420
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   19
            Left            =   945
            ToolTipText     =   "Buscar familia"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   20
            Left            =   945
            ToolTipText     =   "Buscar familia"
            Top             =   1590
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   21
            Left            =   945
            ToolTipText     =   "Buscar marca"
            Top             =   2325
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   22
            Left            =   945
            ToolTipText     =   "Buscar marca"
            Top             =   2760
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   25
            Left            =   945
            ToolTipText     =   "Buscar tipo articulo"
            Top             =   4560
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   26
            Left            =   945
            ToolTipText     =   "Buscar tipo articulo"
            Top             =   4980
            Width           =   240
         End
         Begin VB.Label Label3 
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
            Height          =   285
            Index           =   23
            Left            =   180
            TabIndex        =   77
            Top             =   855
            Width           =   915
         End
         Begin VB.Label Label4 
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
            Index           =   9
            Left            =   225
            TabIndex        =   76
            Top             =   2295
            Width           =   600
         End
         Begin VB.Label Label4 
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
            Index           =   8
            Left            =   225
            TabIndex        =   75
            Top             =   2670
            Width           =   645
         End
         Begin VB.Label Label3 
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
            Height          =   285
            Index           =   22
            Left            =   180
            TabIndex        =   74
            Top             =   3060
            Width           =   3390
         End
         Begin VB.Label Label4 
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
            Index           =   3
            Left            =   180
            TabIndex        =   73
            Top             =   3375
            Width           =   690
         End
         Begin VB.Label Label4 
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
            Index           =   2
            Left            =   180
            TabIndex        =   72
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Artículo"
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
            Index           =   21
            Left            =   135
            TabIndex        =   71
            Top             =   4185
            Width           =   1320
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
            Index           =   20
            Left            =   225
            TabIndex        =   70
            Top             =   1545
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
            Index           =   19
            Left            =   225
            TabIndex        =   69
            Top             =   1185
            Width           =   690
         End
         Begin VB.Label Label3 
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
            Index           =   18
            Left            =   180
            TabIndex        =   68
            Top             =   1980
            Width           =   645
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
            Index           =   17
            Left            =   180
            TabIndex        =   67
            Top             =   4845
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
            Index           =   16
            Left            =   180
            TabIndex        =   66
            Top             =   4485
            Width           =   735
         End
      End
      Begin VB.Frame FrameMovArtSel 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5780
         Left            =   225
         TabIndex        =   24
         Top             =   315
         Width           =   6960
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
            Index           =   87
            Left            =   1275
            TabIndex        =   44
            Top             =   5130
            Width           =   855
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
            Index           =   86
            Left            =   1275
            TabIndex        =   43
            Top             =   4770
            Width           =   855
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
            Index           =   10
            Left            =   3825
            TabIndex        =   39
            Top             =   2880
            Width           =   1350
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
            Index           =   9
            Left            =   1260
            TabIndex        =   40
            Top             =   2880
            Width           =   1350
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
            Index           =   11
            Left            =   1275
            TabIndex        =   41
            Top             =   3585
            Width           =   615
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
            Index           =   12
            Left            =   1275
            TabIndex        =   42
            Top             =   3945
            Width           =   615
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
            Index           =   11
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "Text5"
            Top             =   3585
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
            Index           =   12
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "Text5"
            Top             =   3945
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
            Index           =   6
            Left            =   1275
            MaxLength       =   16
            TabIndex        =   34
            Top             =   1095
            Width           =   1815
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
            Index           =   6
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "Text5"
            Top             =   1095
            Width           =   3615
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
            Index           =   5
            Left            =   3105
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "Text5"
            Top             =   720
            Width           =   3615
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
            Index           =   5
            Left            =   1260
            MaxLength       =   16
            TabIndex        =   33
            Top             =   720
            Width           =   1815
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
            Index           =   7
            Left            =   1275
            MaxLength       =   4
            TabIndex        =   36
            Top             =   1845
            Width           =   615
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
            Index           =   8
            Left            =   1275
            MaxLength       =   4
            TabIndex        =   38
            Top             =   2205
            Width           =   615
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
            Index           =   7
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "Text5"
            Top             =   1845
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
            Index           =   8
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "Text5"
            Top             =   2205
            Width           =   4845
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   66
            Left            =   990
            Top             =   5130
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   65
            Left            =   990
            ToolTipText     =   "Cliente"
            Top             =   4770
            Visible         =   0   'False
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
            Index           =   12
            Left            =   225
            TabIndex        =   55
            Top             =   2895
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
            Index           =   9
            Left            =   2805
            TabIndex        =   54
            Top             =   2895
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            Index           =   10
            Left            =   225
            TabIndex        =   53
            Top             =   2610
            Width           =   630
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
            Index           =   7
            Left            =   225
            TabIndex        =   52
            Top             =   3585
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
            Index           =   6
            Left            =   225
            TabIndex        =   51
            Top             =   3945
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacen"
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
            Index           =   11
            Left            =   225
            TabIndex        =   50
            Top             =   3300
            Width           =   915
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   33
            Left            =   990
            ToolTipText     =   "Buscar almacen"
            Top             =   3585
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   34
            Left            =   990
            ToolTipText     =   "Buscar almacen"
            Top             =   3945
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   990
            Picture         =   "frmInformesNew.frx":2177
            Top             =   2895
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   3540
            Picture         =   "frmInformesNew.frx":2202
            Top             =   2895
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
            Index           =   5
            Left            =   225
            TabIndex        =   47
            Top             =   735
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
            Index           =   33
            Left            =   225
            TabIndex        =   46
            Top             =   1095
            Width           =   645
         End
         Begin VB.Label Label3 
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
            Index           =   4
            Left            =   225
            TabIndex        =   45
            Top             =   405
            Width           =   810
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   29
            Left            =   990
            ToolTipText     =   "Buscar artículo"
            Top             =   735
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   30
            Left            =   990
            ToolTipText     =   "Buscar artículo"
            Top             =   1095
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   31
            Left            =   990
            ToolTipText     =   "Buscar familia"
            Top             =   1845
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   32
            Left            =   990
            ToolTipText     =   "Buscar familia"
            Top             =   2205
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Index           =   7
            Left            =   225
            TabIndex        =   30
            Top             =   5100
            Width           =   645
         End
         Begin VB.Label Label4 
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
            Index           =   6
            Left            =   225
            TabIndex        =   29
            Top             =   4725
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente/Proveedor"
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
            Height          =   285
            Index           =   3
            Left            =   225
            TabIndex        =   28
            Top             =   4410
            Width           =   3390
         End
         Begin VB.Label Label4 
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
            Index           =   5
            Left            =   225
            TabIndex        =   27
            Top             =   2220
            Width           =   645
         End
         Begin VB.Label Label4 
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
            Index           =   4
            Left            =   225
            TabIndex        =   26
            Top             =   1845
            Width           =   600
         End
         Begin VB.Label Label3 
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
            Height          =   285
            Index           =   2
            Left            =   225
            TabIndex        =   25
            Top             =   1530
            Width           =   915
         End
      End
      Begin VB.Frame frameConceptoOrd 
         Caption         =   "Ordenación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5370
         Left            =   7335
         TabIndex        =   11
         Top             =   225
         Width           =   4455
         Begin VB.OptionButton OptNombre 
            Caption         =   "Descripción"
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
            Left            =   2100
            TabIndex        =   127
            Top             =   585
            Width           =   1455
         End
         Begin VB.OptionButton Optcodigo 
            Caption         =   "Código"
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
            Left            =   585
            TabIndex        =   126
            Top             =   585
            Width           =   1215
         End
      End
      Begin VB.Frame FrameMovArtOpc 
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5280
         Left            =   7335
         TabIndex        =   56
         Top             =   225
         Width           =   4455
         Begin MSComctlLib.ListView ListView1 
            Height          =   4230
            Left            =   270
            TabIndex        =   57
            Top             =   780
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   7461
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
         Begin VB.Image cmdSelTodos 
            Height          =   240
            Left            =   4005
            Picture         =   "frmInformesNew.frx":228D
            ToolTipText     =   "Puntear al Debe"
            Top             =   450
            Width           =   240
         End
         Begin VB.Image cmdDeselTodos 
            Height          =   240
            Index           =   0
            Left            =   3645
            Picture         =   "frmInformesNew.frx":23D7
            ToolTipText     =   "Quitar al Debe"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Movimiento"
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
            Index           =   13
            Left            =   270
            TabIndex        =   58
            Top             =   450
            Width           =   2235
         End
      End
      Begin VB.Frame FrameConceptoSel 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2490
         Left            =   225
         TabIndex        =   5
         Top             =   225
         Width           =   6915
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "Text5"
            Top             =   1320
            Width           =   4170
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "Text5"
            Top             =   945
            Width           =   4170
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
            Index           =   1
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   1
            Top             =   1320
            Width           =   830
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
            Index           =   0
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   0
            Top             =   945
            Width           =   830
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   900
            MouseIcon       =   "frmInformesNew.frx":2521
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   915
            MouseIcon       =   "frmInformesNew.frx":2673
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   945
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Index           =   12
            Left            =   225
            TabIndex        =   10
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label Label4 
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
            Index           =   13
            Left            =   225
            TabIndex        =   9
            Top             =   945
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Código"
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
            Height          =   375
            Index           =   8
            Left            =   180
            TabIndex        =   6
            Top             =   540
            Width           =   3120
         End
      End
      Begin VB.Frame FrameInfAlmacenesSel 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2490
         Left            =   225
         TabIndex        =   59
         Top             =   315
         Width           =   6960
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
            Index           =   4
            Left            =   1215
            TabIndex        =   64
            Top             =   1485
            Width           =   1275
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
            Index           =   3
            Left            =   1230
            TabIndex        =   63
            Top             =   1035
            Width           =   1275
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   930
            ToolTipText     =   "Buscar almacén"
            Top             =   1485
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   945
            ToolTipText     =   "Buscar almacén"
            Top             =   1035
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Nro Traspaso"
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
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   62
            Top             =   540
            Width           =   3120
         End
         Begin VB.Label Label4 
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
            Index           =   1
            Left            =   225
            TabIndex        =   61
            Top             =   1035
            Width           =   690
         End
         Begin VB.Label Label4 
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
            Index           =   0
            Left            =   225
            TabIndex        =   60
            Top             =   1485
            Width           =   645
         End
      End
      Begin VB.Frame FrameTipoSalida 
         Caption         =   "Tipo de salida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         Left            =   225
         TabIndex        =   12
         Top             =   3105
         Width           =   6960
         Begin VB.OptionButton optTipoSal 
            Caption         =   "Impresora"
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
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   585
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optTipoSal 
            Caption         =   "Archivo csv"
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
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   1065
            Width           =   1515
         End
         Begin VB.OptionButton optTipoSal 
            Caption         =   "PDF"
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
            Index           =   2
            Left            =   240
            TabIndex        =   20
            Top             =   1545
            Width           =   975
         End
         Begin VB.OptionButton optTipoSal 
            Caption         =   "eMail"
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
            Index           =   3
            Left            =   240
            TabIndex        =   19
            Top             =   2025
            Width           =   975
         End
         Begin VB.TextBox txtTipoSalida 
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
            Index           =   0
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   585
            Width           =   3345
         End
         Begin VB.TextBox txtTipoSalida 
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
            Index           =   1
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1065
            Width           =   4665
         End
         Begin VB.TextBox txtTipoSalida 
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
            Index           =   2
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1545
            Width           =   4665
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   0
            Left            =   6450
            TabIndex        =   15
            Top             =   1065
            Width           =   255
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   1
            Left            =   6450
            TabIndex        =   14
            Top             =   1545
            Width           =   255
         End
         Begin VB.CommandButton PushButtonImpr 
            Caption         =   "Propiedades"
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
            Left            =   5190
            TabIndex        =   13
            Top             =   585
            Width           =   1515
         End
      End
   End
End
Attribute VB_Name = "frmInformesNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit





Public OpcionListado As Integer
'1.  Marcas
'2.  Almacenes propios
'3.  Tipos de Unidad
'4.  Tipos de Artículo

'6.  Listado de Articulos

'7.  Traspaso Almacen
'8.  Movimientos Almacen
'9.  Informe de movimiento de articulos

'20. Actividades
'21. Zonas

'22. Rutas
'24. Tarifas Ventas

'27. Situaciones

'23.  Categorias

'58. Listado de proveedores


'==== Listados de REPARACIONES ====
'==================================
'61. Listado Motivos Pend. Rep.
'65. Listado motivos baja equipos



'110.  Ubicaciones

'999. Incidencias (Éste es nuevo no estaba en frmListado)

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto
Public EsHco As Boolean
    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMar As frmAlmMarcas 'marcas
Attribute frmMar.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios 'almacenes propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmTArt As frmAlmTipoArticulo 'tipo de articulos
Attribute frmTArt.VB_VarHelpID = -1
Private WithEvents frmTUni As frmAlmTipoUnidad 'tipo de unidad
Attribute frmTUni.VB_VarHelpID = -1
Private WithEvents frmUbi As frmAlmUbicaciones 'ubicaciones
Attribute frmUbi.VB_VarHelpID = -1
Private WithEvents frmCat As frmAlmCategorias 'categorias
Attribute frmCat.VB_VarHelpID = -1
Private WithEvents frmFam As frmBasico2 'familias
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmProv As frmBasico2 'Proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmBasico2 'frmAlmArticu2
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmMtoClientes As frmBasico2
Attribute frmMtoClientes.VB_VarHelpID = -1
Private WithEvents frmMov As frmBasico2
Attribute frmMov.VB_VarHelpID = -1
Private WithEvents frmMtoProveedor As frmBasico2
Attribute frmMtoProveedor.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmAct As frmFacActividades 'Actividades
Attribute frmAct.VB_VarHelpID = -1
Private WithEvents frmZon As frmFacZonas 'Zonas
Attribute frmZon.VB_VarHelpID = -1
Private WithEvents frmRut As frmFacRutas 'Rutas
Attribute frmRut.VB_VarHelpID = -1
Private WithEvents frmSit As frmFacSituaciones 'Situaciones
Attribute frmSit.VB_VarHelpID = -1
Private WithEvents frmInc As frmIncidencias 'Incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmTar As frmFacTarifas
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmRMot1 As frmRepMotivosPend
Attribute frmRMot1.VB_VarHelpID = -1
Private WithEvents frmRMot2 As frmRepMotivosBaja
Attribute frmRMot2.VB_VarHelpID = -1

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe

Dim NombreRPT As String

'Los reports
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadNombreRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private ConSubInforme As Boolean 'Si el informe tiene subreports

Private cadPDFrpt As String 'Nombre del informe a enviar por email
Private vMostrarTree As Boolean
Private ExportarPDF As Boolean
Private SoloImprimir As Boolean

Private HaPulsadoImprimir As Boolean


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rc As String
Dim RC2 As String

Dim cDesde As String
Dim cHasta As String

    MontaSQL = False
    
    If Not DatosOk Then Exit Function
    
    
    Select Case OpcionListado
        Case 1 ' marcas
            If Not PonerDesdeHasta2("{smarca.codmarca}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 2 ' almacenes propios
            If Not PonerDesdeHasta2("{salmpr.codalmac}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 3 ' unidades
            If Not PonerDesdeHasta2("{sunida.codunida}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 4 ' tipo de articulos
            If Not PonerDesdeHasta2("{stipar.codtipar}", "T", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 20 ' actividades
            If Not PonerDesdeHasta2("{sactiv.codactiv}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 21 ' zonas
            If Not PonerDesdeHasta2("{szonas.codzonas}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 22 ' rutas
            If Not PonerDesdeHasta2("{srutas.codrutas}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 110 ' ubicaciones
            If Not PonerDesdeHasta2("{subica.codubica}", "T", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 23 ' categorias
            If Not PonerDesdeHasta2("{scateg.codcateg}", "T", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 24 ' tarifas de ventas
            If Not PonerDesdeHasta2("{starif.codlista}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 27 ' situaciones
            If Not PonerDesdeHasta2("{ssitua.codsitua}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        
        Case 61 ' reparaciones motivos pendientes
            If Not PonerDesdeHasta2("{smotre.codmotre}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 65 ' reparaciones motivos de baja
            If Not PonerDesdeHasta2("{smotba.codmotiv}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        
        Case 999 ' incidencias
            If Not PonerDesdeHasta2("{sincid.codincid}", "T", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
        Case 7, 8
            'Cadena para seleccion Desde y Hasta DOCUMENTO
            '----------------------------------------------
            If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
                If Not PonerDesdeHasta(Codigo, "N", 3, 4, "") Then Exit Function
            End If
        Case 58 ' proveedores
            If Not PonerDesdeHasta2("{sprove.codprove}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
            
    End Select
    
    AnyadirAFormula cadFormula, cDesde
    AnyadirAFormula cadSelect, Replace(Replace(cDesde, "{", ""), "}", "")
        
    Select Case OpcionListado
        Case 9 ' movArticulos
            If Not PonerFormulaYParametrosInfMovArt() Then Exit Function
            
            'comprobar que hay datos para mostrar en el Informe
            tabla = "smoval INNER JOIN sartic ON smoval.codartic=sartic.codartic "
        
        Case 6 ' informe de articulos
        
            MontaSqlInfArticulos
            
    End Select
    
    
    MontaSQL = True
    
End Function


Private Sub MontaSqlInfArticulos()
Dim campo As String
Dim devuelve As String
Dim Opcion As Byte, numOp As Byte
Dim cadFrom As String

Dim PrevioArticulos As Boolean
Dim cadParam2 As String

        
        
    'Enero 2022
    If Not PonerParamRPT2(99, cadParam, numParam, cadNombreRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNombreRPT = "rAlmArticulosPVP.rpt"

    cadFrom = " sartic"
            
    cadParam2 = ""
    For Opcion = 0 To 3
        If Me.chkSitaucionArticulo2(Opcion).Value = 1 Then cadParam2 = cadParam2 & "O"
    Next
    If cadParam2 = "" Then
        MsgBox "Seleccione la situacion del articulo", vbExclamation
        Exit Sub
    End If
    Opcion = 0
    

    If Me.chkImpEtiq(0).Value = 0 And Me.chkImpEtiq(1).Value = 0 And Me.chkImpEtiq(3).Value = 1 Then
        'MsgBox "Debe marcar la opcion PVP para que salga el precio mínimo", vbExclamation
        'Exit Sub
        chkImpEtiq(1).Value = 1
    End If
    

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
    
    
    indCodigo = 0
    devuelve = ""
    If Me.chkSitaucionArticulo2(0).Value = 1 Then devuelve = "- NORMAL": indCodigo = indCodigo + 1
    If Me.chkSitaucionArticulo2(1).Value = 1 Then devuelve = devuelve & "- OBSOLETO": indCodigo = indCodigo + 1
    If Me.chkSitaucionArticulo2(2).Value = 1 Then devuelve = devuelve & "- BLOQUEADO": indCodigo = indCodigo + 1
    If Me.chkSitaucionArticulo2(3).Value = 1 Then devuelve = devuelve & "- CADUCADO": indCodigo = indCodigo + 1
    If indCodigo <> 4 Then
      cadTitulo = Trim(cadTitulo & "      Situacion: " & Mid(devuelve, 2))
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
                If Me.chkImpEtiq(0).Value = 0 And Me.chkImpEtiq(1).Value = 1 Then
                    'INFORME PVP
                    indCodigo = 0
                Else
                    If Val(cadFormula) > 0 Then
                        indCodigo = Val(cadFormula) - 1
                    Else
                        indCodigo = miRsAux!cantidad - 1
                    End If
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
            If Not PonerParamRPT2(IIf(chkImpEtiq(3).Value = 1, 87, 80), cadParam, numParam, cadNombreRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNombreRPT = "rAlmArticulosPVP.rpt"
            cadTitulo = IIf(chkImpEtiq(3).Value = 1, "Articulos PVP IVA con precio mínimo", "Articulos PVP IVA")
        Else
            If Not PonerParamRPT2(23, cadParam, numParam, cadNombreRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNombreRPT = "rEtiArticulo.rpt"
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
    
    cadSelect = cadFormula
    tabla = cadFrom

End Sub

Private Function CargarDatosFamiliasDtoEnTmp() As Boolean
Dim miSQL As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo eCargarDatosFamiliasDtoEnTmp
    CargarDatosFamiliasDtoEnTmp = False
    
    conn.Execute "DELETE from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute "DELETE from tmpcommand where codusu = " & vUsu.Codigo
    
    miSQL = "select sfamia.*, sfamiadtos.clasifica,nombre,dtoline1,nomprove "
    miSQL = miSQL & " From sfamiadtos, sfamiatipodto, sfamia left join sprove on sfamia.codprove=sprove.codprove"

    miSQL = miSQL & " Where sfamiadtos.clasifica = sfamiatipodto.clasifica"
    miSQL = miSQL & " AND  sfamia.codfamia=sfamiadtos.codfamia"
    'JUL 2013
    miSQL = miSQL & " AND sprove.OcultarEnListDto = 0"
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "{", "(")
        Codigo = Replace(Codigo, "}", ")")
        miSQL = miSQL & " AND " & Codigo
    End If
    miSQL = miSQL & " order by codfamia,clasifica"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    miSQL = ""
    While Not miRsAux.EOF
     
        'tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,importe1,importe2,importe3,cantidad)
        NumRegElim = NumRegElim + 1
        miSQL = miSQL & ", (" & vUsu.Codigo & "," & NumRegElim & "," & DBLet(miRsAux!CodProve, "N") & "," & miRsAux!Codfamia & ","
        miSQL = miSQL & DBSet(miRsAux!nomprove, "T", "N") & "," & DBSet(miRsAux!nomfamia, "T", "N") & ","
        
        'CERO O DOS
        If miRsAux!clasifica = 0 Then
            miSQL = miSQL & DBSet(miRsAux!dtoline1, "N") & ",0"
        Else
            miSQL = miSQL & "0," & DBSet(miRsAux!dtoline1, "N")
        End If
        miSQL = miSQL & "," & DBSet(miRsAux!maxdtopar, "N")
        miSQL = miSQL & "," & DBSet(miRsAux!Dtopmv, "N")
        miSQL = miSQL & ")"
        
        If (NumRegElim Mod 50) = 0 Then
            miSQL = Mid(miSQL, 2)
            miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,importe1,importe2,importe3,importe4) VALUES " & miSQL
            conn.Execute miSQL
            miSQL = ""
        End If
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If miSQL <> "" Then
        miSQL = Mid(miSQL, 2)
        miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,importe1,importe2,importe3,importe4) VALUES " & miSQL
        conn.Execute miSQL
    End If
    
    'Ahora cargaremos tmpcommand
    miSQL = "insert into tmpcommand(codusu,codprove,codfamia,nomprove,nomfamia,importel,rap1,rap2,cantidad) "
    miSQL = miSQL & " select codusu,campo1,campo2,nombre1,nombre2,importe3,sum(importe1),sum(importe2),max(importe4) from tmpinformes"
    miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " group by 1,2,3"
    conn.Execute miSQL
    
    If NumRegElim > 0 Then
        CargarDatosFamiliasDtoEnTmp = True
    Else
        MsgBox "Ningun dato generado", vbExclamation
    End If
eCargarDatosFamiliasDtoEnTmp:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function


Private Function DatosOk() As Boolean
Dim Sql As String
Dim B As Boolean
    
    B = True
    
    DatosOk = B

End Function


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

Private Sub cmdAccion_Click(Index As Integer)

    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV
    
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal
    End If


    'MONIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII.   Faltaba faltaba !!!!!
    If OpcionListado = 7 And HaPulsadoImprimir Then
        If NumCod <> "" Then
            cadParam = "UPDATE scatra SET situacio=1" 'Impreso
            cadParam = cadParam & " WHERE codtrasp=" & NumCod
            ejecutar cadParam, False
            cadParam = ""
        End If
    End If
End Sub

Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
Dim vIndCodigo As Integer
    
    vMostrarTree = False
    conSubRPT = False
        
    Select Case OpcionListado
        Case 1 'marcas
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={smarca.codmarca}|"
            Else
                cadParam = cadParam & "pOrden={smarca.nommarca}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rAlmMarcas.rpt"
        Case 2 'almacenes propios
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={salmpr.codalmac}|"
            Else
                cadParam = cadParam & "pOrden={salmpr.nomalmac}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rAlmAPropios.rpt"
        Case 3 'unidades
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={sunida.codunida}|"
            Else
                cadParam = cadParam & "pOrden={sunida.nomunida}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rAlmTUnidad.rpt"
        Case 4 'tipos de articulo
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={stipar.codtipar}|"
            Else
                cadParam = cadParam & "pOrden={stipar.nomtipar}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rAlmTArticulo.rpt"
        Case 20 'actividades
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={sactiv.codactiv}|"
            Else
                cadParam = cadParam & "pOrden={sactiv.nomactiv}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rFacActividades.rpt"
        Case 21 'zonas
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={szonas.codzonas}|"
            Else
                cadParam = cadParam & "pOrden={szonas.nomzonas}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rFacZonas.rpt"
        Case 22 'rutas
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={srutas.codrutas}|"
            Else
                cadParam = cadParam & "pOrden={srutas.nomrutas}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rFacRutas.rpt"
        Case 23 ' categorias
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={scateg.codcateg}|"
            Else
                cadParam = cadParam & "pOrden={scateg.descateg}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rAlmCategorias.rpt"
        Case 24 'tarifas de precios
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={starif.codlista}|"
            Else
                cadParam = cadParam & "pOrden={starif.nomlista}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rFacTarifasVen.rpt"
        Case 27 ' situaciones
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={ssitua.codsitua}|"
            Else
                cadParam = cadParam & "pOrden={ssitua.nomsitua}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rFacSituaciones.rpt"
        
        Case 61 'reparaciones motivos pendientes
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={smotre.codmotre}|"
            Else
                cadParam = cadParam & "pOrden={smotre.nommotre}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rRepMotivosPend.rpt"
            
        Case 65 'reparaciones motivos de baja
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={smotba.codmotiv}|"
            Else
                cadParam = cadParam & "pOrden={smotba.desmotiv}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rRepMotivosBaja.rpt"
        
        
        
        Case 999 ' incidencias
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={sincid.codincid}|"
            Else
                cadParam = cadParam & "pOrden={sincid.nomincid}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rIncidencias.rpt"
        Case 110 ' ubicaciones
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={subica.codubica}|"
            Else
                cadParam = cadParam & "pOrden={subica.nomubica}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rAlmUbica.rpt"
        Case 7, 8
            If OpcionListado = 7 Then '7: Traspaso Almacen
                If EsHco Then
                    indRPT = 2
                Else
                    indRPT = 1
                End If
            ElseIf OpcionListado = 8 Then '8: Movimientos Almacen
                If EsHco Then
                    indRPT = 4
                Else
                    indRPT = 3
                End If
            End If
        
            cadParam = "|"
            If Not PonerParamEmpresa(cadParam, numParam) Then Exit Sub
            If PonerParamRPT2(CByte(indRPT), cadParam, numParam, cadNombreRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
            
            End If
            
        Case 9 ' movimientos de articulos
            indRPT = 75
            If Not PonerParamRPT2(CByte(indRPT), cadParam, numParam, cadNombreRPT, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then cadNombreRPT = "rAlmMovim.rpt"
            
        Case 58 ' proveedores
            If Me.Optcodigo.Value Then
                cadParam = cadParam & "pOrden={sprove.codprove}|"
            Else
                cadParam = cadParam & "pOrden={sprove.nomprove}|"
            End If
            numParam = numParam + 1
            cadNombreRPT = "rComProve.rpt"
            
        Case 6 ' listado de articulos
            vIndCodigo = 0
            If chkImpEtiq(0).Value = 1 Then
                OpcionListado = 513   'para que imprmia etiquetas directamente
                vIndCodigo = 1 'Indicamos que hemos cambiado
                
                
                
            End If
    
    End Select
    
    ImprimeGeneral
    
        
    If vIndCodigo = 1 Then OpcionListado = 6
    
    
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook OpcionListado
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Sub AccionesCSV()
Dim Sql As String

    'Monto el SQL
    Select Case OpcionListado
        Case 1 'marcas
            Sql = "Select codmarca AS Código,nommarca as Descripcion "
            Sql = Sql & " From smarca "
        Case 2 'almacenes propios
            Sql = "Select codalmac AS Código,nomalmac as Descripcion "
            Sql = Sql & " From salmpr "
        Case 3 'tipos de unidad
            Sql = "Select codunida AS Código,nomunida as Descripcion, nomunbre as Abrev, tasareciclado as TasaReciclado, estrabajo as Trabajo "
            Sql = Sql & " From sunida "
        Case 4 'tipos de articulos
            Sql = "Select codtipar AS Código,nomtipar as Descripcion "
            Sql = Sql & " From stipar "
        Case 20 'actividades
            Sql = "Select codactiv AS Código,nomactiv as Descripcion "
            Sql = Sql & " From sactiv "
        Case 21 'tipos de articulos
            Sql = "Select codtipar AS Código,nomtipar as Descripcion "
            Sql = Sql & " From szonas "
        Case 22 'rutas
            Sql = "Select codrutas AS Código,nomrutas as Descripcion "
            Sql = Sql & " From srutas "
        Case 23 'categorias
            Sql = "Select codcateg AS Código,descateg as Descripcion, if (ctrlotes=0,'No','Sí') as CtrolLotes "
            Sql = Sql & " From scateg "
        Case 24 'tarifas
            Sql = "Select codlista AS Código,nomlista as Descripcion "
            Sql = Sql & " From starif "
        Case 27 'situaciones
            'tipositu,clioferped,ocultarbus,PermiteEfectAalb
            ' "S|cboclioferped|C|Blq ofe/ped|1450|;S|cboBusqweb|C|Busq. web|1450|;S|cbopermitefecti|C|Alb.efec|1450|;"
            Sql = "Select codsitua AS Código,nomsitua as Descripcion, if (tipositu=0,'No','Sí') as Bloquea,"
            Sql = Sql & "if (clioferped=0,'No','Sí') as BlqOfePed, if (ocultarbus=0,'No','Sí') as BusqWeb, "
            Sql = Sql & "if (permiteefectaalb=0,'No','Sí') as AlbEfec "
            Sql = Sql & " From ssitua "
            
        Case 61 'reparaciones motivos pendientes
            Sql = "Select codmotre AS Código,nommotre as Descripcion "
            Sql = Sql & " From smotre "
        Case 65 'reparaciones motivos baja
            Sql = "Select codmotiv AS Código,desmotiv as Descripcion "
            Sql = Sql & " From smotba "
            
        Case 999 ' incidencias
            Sql = "Select codincid AS Código,nomincid as Descripcion "
            Sql = Sql & " From sincid "
            
        Case 110 'ubicaciones
            Sql = "Select codubica AS Código,nomubica as Descripcion "
            Sql = Sql & " From subica "
        Case 7 ' familias/descuentos
           ' If chkVarios(1).Value = 1 Then
           '
           ' Else
           '
           ' End If
        Case 58 ' proveedores
            Sql = "Select codprove AS Código,nomprove as Nombre, domprove as Domicilio, codpobla as CPostal, pobprove as Poblacion, proprove as Provincia, nifprove as NIF, telprov1 as Telefono, codmacta as Cuenta, maiprov1 as Email1  "
            Sql = Sql & " From sprove "
        
        
    End Select
    
    If cadSelect <> "" Then Sql = Sql & " WHERE " & cadSelect
    
    If Me.Optcodigo.Value Then
        Sql = Sql & " ORDER BY 1 "
    Else
        Sql = Sql & " ORDER BY 2 "
    End If
        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDeselTodos_Click(Index As Integer)
Dim i As Byte

    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = False
    Next i

End Sub

Private Sub cmdSelTodos_Click()
Dim i As Byte

    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = True
    Next i
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
        Optcodigo.Value = True
        
        Select Case OpcionListado
            Case 6
                PonerFoco txtCodigo(62)
            Case 9
                PonerFoco txtCodigo(5)
        End Select
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

   'IMAGES para busqueda
    For H = 0 To 1
        Me.imgBuscar(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgBuscar(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next H
    For H = 3 To 4
        Me.imgBuscar(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgBuscar(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next H
     
    For H = 19 To 34
        Me.imgBuscarG(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgBuscarG(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next H
    
    For H = 87 To 87
        Me.imgBuscarG(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgBuscarG(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next H
    
    CargaIconosAyuda
     
    FrameCobros.visible = True
    
    '###Descomentar
'    CommitConexion
    
    FrameCobrosVisible True, H, W
    
    FrameConceptoVisible False
    FrameMovArtVisible False
    FrameInfAlmacenesVisible False
    FrameInfArticulosVisible False
    
    Select Case OpcionListado
        Case 1 ' listado de marcas
            FrameConceptoVisible True
            indFrame = 5
            tabla = "smarca"
            Me.Caption = "Informe de Marcas"
            Me.imgBuscar(0).ToolTipText = "Buscar marca"
            Me.imgBuscar(1).ToolTipText = "Buscar marca"
        Case 2 ' listado de almacenes propios
            FrameConceptoVisible True
            indFrame = 5
            tabla = "salmpr"
            Me.Caption = "Informe de Almacenes"
            Me.imgBuscar(0).ToolTipText = "Buscar almacen"
            Me.imgBuscar(1).ToolTipText = "Buscar almacen"
        Case 3 ' listado de tipos de unidad
            FrameConceptoVisible True
            indFrame = 5
            tabla = "sunida"
            Me.Caption = "Informe de Tipos Unidad"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        Case 4 ' listado de tipos de articulos
            FrameConceptoVisible True
            indFrame = 5
            tabla = "stipar"
            Me.Caption = "Informe de Tipos de Articulos"
            Me.imgBuscar(0).ToolTipText = "Buscar tipo"
            Me.imgBuscar(1).ToolTipText = "Buscar tipo"
        
        Case 999 ' listado de Situaciones
            FrameConceptoVisible True
            indFrame = 5
            tabla = "sincid"
            Me.Caption = "Informe de Incidencias"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        
        Case 20 ' listado de tipos de actividades
            FrameConceptoVisible True
            indFrame = 5
            tabla = "sactiv"
            Me.Caption = "Informe de Actividades"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        
        Case 21 ' listado de zonas de clientes
            FrameConceptoVisible True
            indFrame = 5
            tabla = "szonas"
            Me.Caption = "Informe de Zonas de clientes"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        
        Case 22 ' listado de rutas
            FrameConceptoVisible True
            indFrame = 5
            tabla = "srutas"
            Me.Caption = "Informe de Rutas"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        
        Case 24 ' listado de tarifas de ventas
            FrameConceptoVisible True
            indFrame = 5
            tabla = "starif"
            Me.Caption = "Informe de Tarifas Venta"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        
        Case 27 ' listado de situaciones
            FrameConceptoVisible True
            indFrame = 5
            tabla = "ssitua"
            Me.Caption = "Informe de Situaciones Especiales"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        
        Case 110 ' listado de ubicaciones
            FrameConceptoVisible True
            indFrame = 5
            tabla = "subica"
            Me.Caption = "Informe Ubicaciones Almacén"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        Case 23 ' listado de categorias
            FrameConceptoVisible True
            indFrame = 5
            tabla = "scateg"
            Me.Caption = "Informe Categorias"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        
        Case 6 ' informe de articulos
            ponerFrameArticulosVisible True, H, W
            FrameInfArticulosVisible True
            tabla = "sartic"
            Me.Caption = "Informe de Artículos"
            
            CargarListViewOrden
            
            AmpliarFrame 4020
            
            'De momento el csv no lo vemos
            VerCSV False
        
        Case 7, 8 ' informes de almacenes
            FrameInfAlmacenesVisible True
            txtCodigo(3).Text = NumCod
            txtCodigo(4).Text = NumCod
            
            If OpcionListado = 7 Then
                tabla = "scatra"
                Me.Caption = "Informe Traspaso de Almacen"
                Codigo = "{scatra.codtrasp}"
                If EsHco Then
                    tabla = "schtra"
                    Me.Caption = "Informe Histórico Traspaso de Almacen"
                    Codigo = "{schtra.codtrasp}"
                End If
                Me.Label3(0).Caption = "Nº Traspaso"
                imgBuscar(3).ToolTipText = "Nº Traspaso"
                imgBuscar(4).ToolTipText = "Nº Traspaso"
            Else
                tabla = "scamov"
                Me.Caption = "Informe Movimientos de Almacén"
                If EsHco Then
                    tabla = "schmov"
                    Me.Caption = "Informe Histórico Movimientos de Almacén"
                End If
                Me.Label3(0).Caption = "Nº Movimiento"
                imgBuscar(3).ToolTipText = "Nº Movimiento"
                imgBuscar(4).ToolTipText = "Nº Movimiento"
                If EsHco Then
                    Codigo = "{schmov.codmovim}"
                Else
                    Codigo = "{scamov.codmovim}"
                End If
            End If
            EnsancharFrame -4455
            
            'De momento el csv no lo vemos
            VerCSV False
            
        Case 9 ' listado de movimientos
            FrameMovArtVisible True
            AmpliarFrame 3020
            tabla = "smoval"
            Me.Caption = "Informe Movimientos Artículos"
            Codigo = "{smoval.codartic}"
            conSubRPT = True
            CargarListView
            
            'De momento el csv no lo vemos
            Me.optTipoSal(1).Enabled = False
            Me.PushButton2(0).Enabled = False
            
        Case 58 ' listado de proveedores
            FrameConceptoVisible True
            tabla = "sprove"
            Me.Caption = "Informe Proveedores"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
            indFrame = 1
            Codigo = "{sprove.codprove}"
            Orden1 = "{sprove.codprove}"
            Orden2 = "{sprove.nomprove}"
        
        Case 61 ' listado de reparaciones motivos pendientes
            FrameConceptoVisible True
            indFrame = 5
            tabla = "smotre"
            Me.Caption = "Motivos Pendientes Reparación"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        
        Case 65 ' listado de reparaciones motivos baja
            FrameConceptoVisible True
            indFrame = 5
            tabla = "smotba"
            Me.Caption = "Motivos Baja Equipos"
            Me.imgBuscar(0).ToolTipText = "Buscar codigo"
            Me.imgBuscar(1).ToolTipText = "Buscar codigo"
        
    End Select
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = Me.Width - 220
    Me.Height = Me.FrameCobros.Height + 20
End Sub

Private Sub EnsancharFrame(incremen As Long)
    Me.Width = Me.Width + incremen
    cmdAccion(1).Left = cmdAccion(1).Left + incremen
    cmdCancel.Left = cmdCancel.Left + incremen
End Sub

Private Sub VerCSV(Ver As Boolean)
    Me.optTipoSal(1).Enabled = Ver
    Me.PushButton2(0).Enabled = Ver
End Sub


Private Sub AmpliarFrame(incremen As Long)
    Me.FrameTipoSalida.Top = Me.FrameTipoSalida.Top + incremen
    Me.cmdAccion(0).Top = Me.cmdAccion(0).Top + incremen
    Me.cmdAccion(1).Top = Me.cmdAccion(1).Top + incremen
    Me.cmdCancel.Top = Me.cmdCancel.Top + incremen
    Me.FrameCobros.Height = Me.FrameCobros.Height + incremen
    Select Case OpcionListado
        Case 7, 8 ' informes de almacenes
        
        Case 9
            Me.FrameMovArtOpc.Top = Me.FrameMovArtSel.Top
            Me.FrameMovArtOpc.Height = FrameMovArtOpc.Height + incremen
            
        Case 6
            Me.FrameInfArticulosOpc.Height = FrameInfArticulosSel.Height
            Me.FrameInfArticulosOrd.Top = Me.FrameTipoSalida.Top
    End Select
    
    Me.Height = Me.FrameCobros.Height
End Sub


Private Sub frmAct_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'actividades
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmRMot1_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmRMot2_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmRut_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'situaciones
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'tarifas de precios
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmZon_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'zonas
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/MM/yyyy")
End Sub


Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object

    indCodigo = Index
    Select Case Index
        Case 0
            indCodigo = 6
    End Select
    
    'FECHA
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtCodigo(indCodigo).Text <> "" Then frmC.Fecha = CDate(txtCodigo(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub frmF_Selec(vFecha As Date)
 'Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmMov_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'nro movimiento
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "0000000")
End Sub

Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoClientes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    If indCodigo > 0 Then
        txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
'        'EL 0 es para el listado de bultos
'        Me.txtClie.Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
'        txtClie_LostFocus
    End If
End Sub

Private Sub frmMtoProveedor_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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

Private Sub imgBuscarG_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 19, 20, 31, 32 'cod. FAMILIA
            Select Case Index
                Case 19, 20: indCodigo = Index + 43
                Case 31, 32: indCodigo = Index - 24
            End Select
            
            Set frmFam = New frmBasico2
            AyudaFamilias frmFam, txtCodigo(indCodigo)
            Set frmFam = Nothing
            
        Case 33, 34 'cod. ALMACEN
            Select Case Index
                Case 33, 34: indCodigo = Index - 22
            End Select
            
            AbrirFrmAlmPropios indCodigo
        
        Case 25, 26 'cod TIPO ARTICULO
            indCodigo = Index + 43
            AbrirFrmTipoArt indCodigo
            
        Case 27, 28, 29, 30 'cod. ARTICULO
            ' ====  [16/09/2009] LAURA : Listado Articulos con componentes (añade index 95 y 98)
            Select Case Index
                Case 27: indCodigo = 70
                Case 28: indCodigo = 71
                Case 29: indCodigo = 5
                Case 30: indCodigo = 6
            End Select

            Set frmMtoArticulos = New frmBasico2
            AyudaArticulos frmMtoArticulos, txtCodigo(indCodigo)
            Set frmMtoArticulos = Nothing

        Case 4, 5, 21, 22, 59, 60, 110, 111 'cod. MARCA
            Select Case Index
                Case 4, 5: indCodigo = Index + 73
                Case 21, 22: indCodigo = Index + 43
                Case 59, 60:  indCodigo = Index - 32
                Case 110, 111:  indCodigo = Index + 44
            End Select
            
            AbrirFrmMarcas indCodigo
            
        Case 23, 24 ' proveedores
            Select Case Index
                Case 15, 16: indCodigo = Index + 3
                Case 23, 24: indCodigo = Index + 43
                Case 63, 64: indCodigo = Index + 16
                Case 103, 104: indCodigo = Index + 31
                Case 143, 144, 148: indCodigo = Index
            End Select
            Set frmMtoProveedor = New frmBasico2
            AyudaProveedores frmMtoProveedor, txtCodigo(indCodigo)
            Set frmMtoProveedor = Nothing
        
            
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

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'familias
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    If txtCodigo(indCodigo).Text <> "" Then txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTArt_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTUni_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmUbi_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCat_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1
            Select Case OpcionListado
                Case 1 'marcas
                    AbrirFrmMarcas (Index)
                Case 2 'almacenes propios
                    AbrirFrmAlmPropios (Index)
                Case 3 ' tipos de unidad
                    AbrirFrmTUnidad (Index)
                Case 4 'tipos de articulos
                    AbrirFrmTipoArt (Index)
                Case 20 'actividades
                    AbrirFrmActividades (Index)
                Case 21 'zonas
                    AbrirFrmZonas (Index)
                Case 22 'rutas
                    AbrirFrmRutas (Index)
                Case 24 ' tarifas de venta
                    AbrirFrmTarifasVta (Index)
                Case 27 'situaciones
                    AbrirFrmSituaciones (Index)
                
                Case 61 'motivos pendientes
                    AbrirFrmRepMotivosPend (Index)
                Case 65 'motivos de baja
                    AbrirFrmRepMotivosBaja (Index)

                Case 110 'ubicaciones
                    AbrirFrmAlmUbicaciones (Index)
                Case 23 'categorias
                    AbrirFrmAlmCategorias (Index)
                Case 58
                    AbrirFrmProveedores (Index)
                Case 999
                    AbrirFrmIncidencias (Index)
            End Select
        Case 3, 4
            If OpcionListado = 7 Or OpcionListado = 8 Then
'            Case 7, 8 '7: Informe de Traspasos de Almacenes
                  '8: Informe de Movimientos de Almacen
                indCodigo = Index
                  
                Set frmMov = New frmBasico2
                AyudaAlmMovimientosPrev frmMov, EsHco
                Set frmMov = Nothing
            End If
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
'        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
'        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
         frmPpal.CommonDialog1.Filter = "*.csv|*.csv"
         
    Else
        frmPpal.CommonDialog1.Filter = "*.pdf|*.pdf"
    End If
    frmPpal.CommonDialog1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmPpal.CommonDialog1.FilterIndex = 1
    frmPpal.CommonDialog1.ShowSave
    If frmPpal.CommonDialog1.FileTitle <> "" Then
        If Dir(frmPpal.CommonDialog1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmPpal.CommonDialog1.FileName
    End If

End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'codigo desde
            Case 1: KEYBusqueda KeyAscii, 1 'codigo hasta
            
            'listado de movimientos de articulos
            Case 5: KEYBusquedaG KeyAscii, 29 'articulo desde
            Case 6: KEYBusquedaG KeyAscii, 30 'articulo hasta
            Case 7: KEYBusquedaG KeyAscii, 31 'familia desde
            Case 8: KEYBusquedaG KeyAscii, 32 'familia hasta
            Case 11: KEYBusquedaG KeyAscii, 33 'almacen desde
            Case 12: KEYBusquedaG KeyAscii, 34 'almacen hasta
            Case 9: KEYFecha2 KeyAscii, 0 'fecha desde
            Case 10: KEYFecha2 KeyAscii, 1 'fecha hasta
            
            ' listado de articulos
            Case 107: KEYBusquedaG KeyAscii, 87 'almacen desde
            Case 62: KEYBusquedaG KeyAscii, 19 'familia desde
            Case 63: KEYBusquedaG KeyAscii, 20 'familia hasta
            Case 64: KEYBusquedaG KeyAscii, 21 'marca desde
            Case 65: KEYBusquedaG KeyAscii, 22 'marca hasta
            Case 66: KEYBusquedaG KeyAscii, 23 'proveedor desde
            Case 67: KEYBusquedaG KeyAscii, 24 'proveedor hasta
            Case 68: KEYBusquedaG KeyAscii, 25 'tipo de articulo desde
            Case 69: KEYBusquedaG KeyAscii, 26 'tipo de articulo hasta
            Case 70: KEYBusquedaG KeyAscii, 27 'articulo desde
            Case 71: KEYBusquedaG KeyAscii, 28 'articulo hasta
        
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYBusquedaG(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscarG_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub KEYFecha2(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1
            Select Case OpcionListado
                Case 1 ' marcas
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "smarca", "nommarca", "codmarca", , "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
                Case 2 ' almacenes propios
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "salmpr", "nomalmac", "codalmac", , "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
                Case 3 ' tipos de unidad
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sunida", "nomunida", "codunida", , "N")
                Case 4 ' tipos de articulos
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "stipar", "nomtipar", "codtipar", , "T")
                Case 20 ' actividades
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sactiv", "nomactiv", "codactiv", , "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
                Case 21 ' zonas
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "szonas", "nomzonas", "codzonas", , "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
                Case 22 ' rutas
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "srutas", "nomrutas", "codrutas", , "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
                Case 23 ' categorias
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "scateg", "descateg", "codcateg", , "T")
                Case 24 ' tarifas
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "starif", "nomlista", "codlista", , "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
                Case 27 ' situaciones
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "ssitua", "nomsitua", "codsitua", , "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
                
                Case 61 ' motivos pendientes equipos
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "smotre", "nommotre", "codmotre", , "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
                Case 65 ' motivos baja equipos
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "smotba", "desmotiv", "codmotiv", , "N")
                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
                
                Case 58 ' proveedores
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sprove", "nomprove", "codprove", , "N")
                Case 110 ' ubicaciones
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "subica", "nomubica", "codubica", , "T")
                Case 999 ' incidencias
                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sincid", "nomincid", "codincid", , "T")
            End Select
        
        Case 7, 8, 62, 63  ' familias
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sfamia", "nomfamia", "codfamia", , "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        
        Case 64, 65  ' marcas
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "smarca", "nommarca", "codmarca", , "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        
        Case 66, 67  ' proveedores
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sprove", "nomprove", "codprove", , "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 68, 69  ' Tipo de articulos
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "stipar", "nomtipar", "codtipar", , "T")
        
        Case 5, 6, 70, 71 ' articulos
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sartic", "nomartic", "codartic", , "T")
        
        Case 9, 10 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 11, 12, 72 ' almacenes
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "salmpr", "nomalmac", "codalmac", , "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
        Case 86, 87
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
                If (Index = 86 Or Index = 87) Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            End If
  
  End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los cobros a clientes por fecha vencimiento
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub


Private Sub FrameConceptoVisible(visible As Boolean)
    FrameConceptoSel.visible = visible
    frameConceptoOrd.visible = visible
    
    FrameConceptoSel.Enabled = visible
    frameConceptoOrd.Enabled = visible
End Sub

Private Sub FrameMovArtVisible(visible As Boolean)
    FrameMovArtSel.visible = visible
    FrameMovArtOpc.visible = visible
    
    FrameMovArtSel.Enabled = visible
    FrameMovArtOpc.Enabled = visible
End Sub

Private Sub FrameInfAlmacenesVisible(visible As Boolean)
    FrameInfAlmacenesSel.visible = visible
    FrameInfAlmacenesSel.Enabled = visible
End Sub

Private Sub FrameInfArticulosVisible(visible As Boolean)
    FrameInfArticulosSel.visible = visible
    FrameInfArticulosSel.Enabled = visible
    
    FrameInfArticulosOpc.visible = visible
    FrameInfArticulosOpc.Enabled = visible
    
    FrameInfArticulosOrd.visible = visible
    FrameInfArticulosOrd.Enabled = visible
End Sub

Private Sub AbrirFrmMarcas(Indice As Integer)
    indCodigo = Indice
    Set frmMar = New frmAlmMarcas
    frmMar.DatosADevolverBusqueda = "0|1|"
    frmMar.Show vbModal
    Set frmMar = Nothing
End Sub
 
Private Sub AbrirFrmAlmPropios(Indice As Integer)
    indCodigo = Indice
    Set frmAlm = New frmAlmAlPropios
    frmAlm.DatosADevolverBusqueda = "0|1|"
    frmAlm.Show vbModal
    Set frmAlm = Nothing
End Sub

Private Sub AbrirFrmTipoArt(Indice As Integer)
    indCodigo = Indice
    Set frmTArt = New frmAlmTipoArticulo
    frmTArt.DatosADevolverBusqueda = "0|1|"
    frmTArt.Show vbModal
    Set frmTArt = Nothing
End Sub

Private Sub AbrirFrmTUnidad(Indice As Integer)
    indCodigo = Indice
    Set frmTUni = New frmAlmTipoUnidad
    frmTUni.DatosADevolverBusqueda = "0|1|"
    frmTUni.Show vbModal
    Set frmTUni = Nothing
End Sub

Private Sub AbrirFrmAlmUbicaciones(Indice As Integer)
    indCodigo = Indice
    Set frmUbi = New frmAlmUbicaciones
    frmUbi.DatosADevolverBusqueda = "0|1|"
    frmUbi.Show vbModal
    Set frmUbi = Nothing
End Sub

Private Sub AbrirFrmAlmCategorias(Indice As Integer)
    indCodigo = Indice
    Set frmCat = New frmAlmCategorias
    frmCat.DatosADevolverBusqueda = "0|1|"
    frmCat.Show vbModal
    Set frmCat = Nothing
End Sub

Private Sub AbrirFrmFamilias(Indice As Integer)
    indCodigo = Indice
    Set frmFam = New frmBasico2
    AyudaFamilias frmFam, txtCodigo(Indice)
    Set frmFam = Nothing
End Sub

Private Sub AbrirFrmProveedores(Indice As Integer)
    indCodigo = Indice
    Set frmProv = New frmBasico2
    AyudaProveedores frmProv, txtCodigo(Indice)
    Set frmProv = Nothing
End Sub

Private Sub AbrirFrmActividades(Indice As Integer)
    indCodigo = Indice
    Set frmAct = New frmFacActividades
    frmAct.DatosADevolverBusqueda = "0|1|"
    frmAct.Show vbModal
    Set frmAct = Nothing
End Sub

Private Sub AbrirFrmZonas(Indice As Integer)
    indCodigo = Indice
    Set frmZon = New frmFacZonas
    frmZon.DatosADevolverBusqueda = "0|1|"
    frmZon.Show vbModal
    Set frmZon = Nothing
End Sub

Private Sub AbrirFrmRutas(Indice As Integer)
    indCodigo = Indice
    Set frmRut = New frmFacRutas
    frmRut.DatosADevolverBusqueda = "0|1|"
    frmRut.Show vbModal
    Set frmRut = Nothing
End Sub

Private Sub AbrirFrmTarifasVta(Indice As Integer)
    indCodigo = Indice
    Set frmTar = New frmFacTarifas
    frmTar.DatosADevolverBusqueda = "0|1|"
    frmTar.Show vbModal
    Set frmTar = Nothing
End Sub


Private Sub AbrirFrmSituaciones(Indice As Integer)
    indCodigo = Indice
    Set frmSit = New frmFacSituaciones
    frmSit.DatosADevolverBusqueda = "0|1|"
    frmSit.Show vbModal
    Set frmSit = Nothing
End Sub

Private Sub AbrirFrmIncidencias(Indice As Integer)
    indCodigo = Indice
    Set frmInc = New frmIncidencias
    frmInc.DatosADevolverBusqueda = "0|1|"
    frmInc.Show vbModal
    Set frmInc = Nothing
End Sub


Private Sub AbrirFrmRepMotivosPend(Indice As Integer)
    indCodigo = Indice
    Set frmRMot1 = New frmRepMotivosPend
    frmRMot1.DatosADevolverBusqueda = "0|1|"
    frmRMot1.Show vbModal
    Set frmRMot1 = Nothing
End Sub

Private Sub AbrirFrmRepMotivosBaja(Indice As Integer)
    indCodigo = Indice
    Set frmRMot2 = New frmRepMotivosBaja
    frmRMot2.DatosADevolverBusqueda = "0|1|"
    frmRMot2.Show vbModal
    Set frmRMot2 = Nothing
End Sub



Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

'#################################################
'###########    A Ñ A D I D O     ################  DE  NUEVA CONTA DE DAVID
'#################################################
Private Sub PonerDatosPorDefectoImpresion(ByRef formu As Form, SoloImpresora As Boolean, Optional NombreArchivoEx As String)
On Error Resume Next
'        AbiertoOtroFormEnListado = False
        
        formu.txtTipoSalida(0).Text = Printer.DeviceName
        If Err.Number <> 0 Then
            formu.txtTipoSalida(0).Text = "No hay impresora instalada"
            Err.Clear
        End If
        If SoloImpresora Then Exit Sub
        
        formu.txtTipoSalida(1).Text = App.Path & "\Exportar\" & NombreArchivoEx & ".csv"
        formu.txtTipoSalida(2).Text = App.Path & "\Exportar\" & NombreArchivoEx & ".pdf"
        
        If Err.Number <> 0 Then Err.Clear
    
End Sub


'PDF=true   CSV=false
Private Function EliminarDocum(PDF As Boolean) As Boolean
    On Error Resume Next
    If PDF Then
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    Else
        If Dir(App.Path & "\docum.csv", vbArchive) <> "" Then Kill App.Path & "\docum.csv"
    End If
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Err.Clear
        EliminarDocum = False
    Else
        EliminarDocum = True
    End If
End Function


Private Sub ponerLabelBotonImpresion(ByRef BotonAcept As CommandButton, ByRef BotonImpr As CommandButton, SelectorImpresion As Integer)
    On Error GoTo eponerLabelBotonImpresion
    If SelectorImpresion = 0 Then
        BotonAcept.Caption = "&Vista previa"
    Else
        BotonAcept.Caption = "&Aceptar"
    End If
    BotonImpr.visible = SelectorImpresion = 0
    
eponerLabelBotonImpresion:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function ImprimeGeneral() As Boolean
Dim cadPDFrpt As String

    Screen.MousePointer = vbHourglass


'    frmPpal.SkinFramework1.AutoApplyNewWindows = False
'    frmPpal.SkinFramework1.AutoApplyNewThreads = False

  
    HaPulsadoImprimir = False
    cadPDFrpt = cadNombreRPT
    With frmVisReport
        .Informe = App.Path & "\Informes\"
        If ExportarPDF Then
            'PDF
            .Informe = .Informe & cadPDFrpt
        Else
            'IMPRIMIR
            .Informe = .Informe & cadNombreRPT
        End If
        .FormulaSeleccion = cadFormula
        .SoloImprimir = False
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .ConSubInforme = ConSubInforme

        .NumCopias = 1

        .SoloImprimir = SoloImprimir
        .ExportarPDF = ExportarPDF
        .MostrarTree = vMostrarTree
        If OpcionListado = 513 Then
                .SoloImprimir = True
                
                'ETIQUETAS TAXCo
                If vParamAplic.NumeroInstalacion = vbTaxco Then .ForzarNombreImpresora = "GODEX500"
          
        End If
        .Show vbModal
        HaPulsadoImprimir = .EstaImpreso
        
      End With
    
    
'     'DAVID
'     frmPpal.SkinFramework1.AutoApplyNewWindows = True
'     frmPpal.SkinFramework1.AutoApplyNewThreads = True
    
End Function

Private Function CopiarFicheroASalida(csv As Boolean, Salida As String, Optional SinMensaje As Boolean) As Boolean
    CopiarFicheroASalida = False
    If Dir(Salida, vbArchive) <> "" Then
        If Not SinMensaje Then
            If Not csv Then
                If MsgBox("Fichero ya existe. ¿Reemplazar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
            End If
        End If
    End If
    
    
    If csv Then
        FileCopy App.Path & "\docum.csv", Salida
    Else
        FileCopy App.Path & "\docum.pdf", Salida
    End If
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Copiando " & Salida
    Else
        If Not SinMensaje Then
            MsgBox "Fichero:  " & Salida & vbCrLf & "Generado con éxito.", vbInformation
        End If
        CopiarFicheroASalida = True
    End If
End Function


Private Sub LanzaProgramaAbrirOutlook(outTipoDocumento As Integer, Optional emailDestinatario As String)
Dim NombrePDF As String
Dim Aux As String
Dim Lanza As String

    On Error GoTo ELanzaProgramaAbrirOutlook

    If Not PrepararCarpetasEnvioMail(True) Then Exit Sub
    
    If Not ExisteARIMAILGES Then Exit Sub

    'Primer tema. Copiar el docum.pdf con otro nombre mas significatiov
    Select Case outTipoDocumento
    Case 1 To 110
        'Marcas
        Aux = "Marcas|Almacenes Propios|Tipos Unidad|Tipos Artículos|||Movimientos Almacen|Traspaso Almacen|Movimientos Articulos||"
        Aux = Aux & "Actividades|Zonas|Rutas||||||||"
        Aux = Aux & "||Categorias|Tarifas|||Situaciones Especiales||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "|||||||Proveedores|||"
        Aux = Aux & "Motivos Pdtes Reparacion||||Motivos Baja Equipos||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "|||||||||Ubicaciones|"

        
            
        Aux = RecuperaValor(Aux, outTipoDocumento) & ".pdf"
             
    Case 999 ' incidencias
        Aux = "Incidencias.pdf"
    End Select
    NombrePDF = App.Path & "\temp\" & Aux
    If Dir(NombrePDF, vbArchive) <> "" Then Kill NombrePDF
    FileCopy App.Path & "\docum.pdf", NombrePDF
    
    Aux = FijaDireccionEmail(outTipoDocumento)
    If Aux = "" And emailDestinatario <> "" Then Aux = emailDestinatario
    Lanza = Aux & "|"
    Aux = ""
    Select Case outTipoDocumento
        
    Case 1 To 110
        Aux = "Marcas|AlmPropios|TiposUnidad|TipoArtículos|||MovimientosAlmacen|TraspasoAlmacen|MovimientosArticulos||"
        Aux = Aux & "Actividades|Zonas|Rutas||||||||"
        Aux = Aux & "||Categorias|Tarifas|||SituacionesEspeciales||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "|||||||Proveedores|||"
        Aux = Aux & "MotivosPdtesRep||||MotivosBajaEquipos||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "|||||||||Ubicaciones|"
        '--------------------------------------------------
        Aux = RecuperaValor(Aux, outTipoDocumento)
        
    Case 999
        Aux = "Incidencias"
    End Select
    Aux = vEmpresa.nomresum & ". " & Aux
    
    Lanza = Lanza & Aux & "|"
    
    'Aqui pondremos lo del texto del BODY
    Aux = ""
    Lanza = Lanza & Aux & "|"
    
    
    'Envio o mostrar
    Lanza = Lanza & "0"   '0. Display   1.  send
    
    'Campos reservados para el futuro
    Lanza = Lanza & "||||"
    
    'El/los adjuntos
    Lanza = Lanza & NombrePDF & "|"
    
    Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Aux, vbNormalFocus
    
    Exit Sub
ELanzaProgramaAbrirOutlook:
    MuestraError Err.Number, Err.Description
End Sub

Private Function FijaDireccionEmail(outTipoDocumento As Integer) As String
Dim campoemail As String
Dim otromail As String


    FijaDireccionEmail = ""
    campoemail = ""
    
'    If outTipoDocumento < 50 Then
''            'Para provedores
''            If outTipoDocumento = 51 Or outTipoDocumento = 52 Or outTipoDocumento = 53 Then
''                campoemail = "maiprov1"
''                otromail = "maiprov2"
''            Else
''                campoemail = "maiprov2"
''                otromail = "maiprov1"
''            End If
''            campoemail = DevuelveDesdeBDNew(cpconta, "proveedor", campoemail, "codprove", Me.outCodigoCliProv, "N", otromail)
'            If campoemail = "" Then campoemail = otromail
'        Else
'            'Para Socios
'            If outTipoDocumento >= 100 Then
'                campoemail = "maisocio"
'                otromail = "maisocio"
'            Else
'                campoemail = "maisocio"
'                otromail = "maisocio"
'            End If
''            campoemail = DevuelveDesdeBDNew(cAgro, "rsocios", campoemail, "codsocio", Me.outCodigoCliProv, "N") ' , otromail)
'            If campoemail = "" Then campoemail = otromail
'        End If
'    End If
    FijaDireccionEmail = campoemail
End Function


Private Function GeneraFicheroCSV(cadSQL As String, Salida As String, Optional OcultarMensajeCreacionCorrecta As Boolean) As Boolean
Dim NF As Integer
Dim i  As Integer

    On Error GoTo eGeneraFicheroCSV
    GeneraFicheroCSV = False
    
    
    If Dir(Salida, vbArchive) <> "" Then
        If MsgBox("El fichero ya existe. ¿Sobreescribir?", vbQuestion + vbYesNo) <> vbYes Then Exit Function
    End If
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Ningun dato generado", vbExclamation
        cadSQL = ""
    Else
        NF = FreeFile
        Open App.Path & "\docum.csv" For Output As #NF
        'Cabecera
        cadSQL = ""
        For i = 0 To miRsAux.Fields.Count - 1
            cadSQL = cadSQL & ";""" & miRsAux.Fields(i).Name & """"
        Next i
        Print #NF, Mid(cadSQL, 2)
    
    
        'Lineas
        While Not miRsAux.EOF
            cadSQL = ""
            For i = 0 To miRsAux.Fields.Count - 1
                cadSQL = cadSQL & ";""" & DBLet(miRsAux.Fields(i).Value, "T") & """"
            Next i
            Print #NF, Mid(cadSQL, 2)
            
            
            
            miRsAux.MoveNext
        Wend
        cadSQL = "OK"
    End If
    miRsAux.Close
    Close #NF

    If cadSQL = "OK" Then
        If CopiarFicheroASalida(True, Salida, OcultarMensajeCreacionCorrecta) Then GeneraFicheroCSV = True
    End If
    
    Exit Function
eGeneraFicheroCSV:
    MuestraError Err.Number, Err.Description
End Function


Private Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadSelect = ""
    cadParam = "|"
    numParam = 0
    cadNombreRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub



Private Function PonerDesdeHasta2(campo As String, Tipo As String, ByRef Desde As TextBox, ByRef DesD As TextBox, ByRef Hasta As TextBox, ByRef DesH As TextBox, param As String) As Boolean
Dim devuelve As String
Dim Cad As String
Dim Subtipo As String 'F: fecha   N: numero   T: texto  H: HORA



    PonerDesdeHasta2 = False
    
    Select Case Tipo
    Case "F", "FEC"
        'Campos fecha
        Subtipo = "F"
    
    Case "N"
        'concepto
        Subtipo = "N"
        
    Case "T"
        Subtipo = "T"
        
    End Select
    
    devuelve = CadenaDesdeHasta(CStr(Desde), CStr(Hasta), campo, Subtipo)
    If devuelve = "Error" Then
        PonerFoco Desde
        Exit Function
    End If
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If devuelve = "" Then
        PonerDesdeHasta2 = True
        Exit Function
    End If
    
    'QUITO LAS LLAVES
    devuelve = Replace(devuelve, "{", "")
    devuelve = Replace(devuelve, "}", "")
    
    If Subtipo <> "F" And Subtipo <> "FH" Then
        'Fecha para Crystal Report

        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(Desde.Text, Hasta.Text, campo, Subtipo)
        Cad = Replace(Cad, "{", "")
        Cad = Replace(Cad, "}", "")
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH2(param, Desde, Hasta, DesD, DesH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta2 = True
    End If
End Function


Private Function AnyadirParametroDH2(Cad As String, ByRef TextoDESDE As TextBox, TextoHasta As TextBox, ByRef TD As TextBox, ByRef TH As TextBox) As String
On Error Resume Next
    
    
    If Not TextoDESDE Is Nothing Then
         If TextoDESDE.Text <> "" Then
            Cad = Cad & "desde " & TextoDESDE.Text
'            If TD.Caption <> "" Then Cad = Cad & " - " & TD.Caption
        End If
    End If
    If Not TextoHasta Is Nothing Then
        If TextoHasta.Text <> "" Then
            Cad = Cad & "  hasta " & TextoHasta.Text
'            If TH.Caption <> "" Then Cad = Cad & " - " & TH.Caption
        End If
    End If
    
    AnyadirParametroDH2 = Cad
    If Err.Number <> 0 Then Err.Clear
End Function




Private Function ExisteARIMAILGES()
Dim Sql As String

    If Dir(App.Path & "\ArimailGes.exe") = "" Then
        MsgBox "No existe el programa ArimailGes.exe. Llame a Ariadna.", vbExclamation
        ExisteARIMAILGES = False
    Else
        ExisteARIMAILGES = True
    End If
End Function




Private Sub CargarListView()
'Carga el List View del frame: frameMovimArtic
'con los parametros de la tabla: stipom (Tipos de Movimientos)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", 1000
    ListView1.ColumnHeaders.Add , , "Descripción", 2650
    
    Sql = "select codtipom,nomtipom from stipom where muevesto=1"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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

Private Sub AbrirFrmClientes()
'Clientes

    

            Set frmMtoClientes = New frmBasico2
            AyudaClientes frmMtoClientes, txtCodigo(indCodigo).Text
            Set frmMtoClientes = Nothing
    
    
    
    
End Sub

Private Function PonerFormulaYParametrosInfMovArt() As Boolean
Dim Cad As String
Dim todosMarcados As Boolean
Dim devuelve As String
Dim i As Byte

    PonerFormulaYParametrosInfMovArt = False
'    InicializarVbles True
    
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
    PonerFormulaYParametrosInfMovArt = True
    
End Function

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


Private Sub CargarListViewOrden()
'Carga el List View del frame: frameInfArticulos
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Familia, MArca, Proveedor, Tipo de Articulo, Articulo
Dim ItmX As ListItem

    'Los encabezados
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Campo", 1900 '1600
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Familia"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Marca"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Proveedor"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Tipo Articulo"
End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView2
End Sub

Private Sub cmdSubir_Click()
'Subir el item seleccionado del listview2 una posicion
    SubirItemList Me.ListView2
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

End Function


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
        'Me.Label9.Caption = "Informe de Articulos"
       
        W = 8715
    Else
'%=%=cuando se haga el 18 y 247 activar
'        If OpcionListado = 18 Then
'            Me.Label9.Caption = "Informe Stocks Maximos y Minimos"
'            Label4(36).Caption = "Almacén"
'            W = 7495
'        Else
'            'NUEVA OCPION:  247
'            'Corregir tarifas y eso
'            chkMinimoCorreg.visible = True
'            Me.Label9.Caption = "Verificación tarifas y P.V.P."
'            FrameTapaINCORRECTO.visible = True
'            Label4(36).Caption = "Tarifa"
'            cmbDecimales.ListIndex = 0
'            W = 7395
'        End If
        
       
    End If
    H = 7095
    
'    PonerFrameVisible Me.FrameInfArticulos, visible, H, W
    
    If visible = True Then
        'visible orden campos si opcionlistado=6
        Me.FrameInfArticulosOrd.visible = B
        Label4(36).visible = Not B ' label almacen

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
        cmbProduccion.visible = Not B
    
    End If
    
    
'    Me.cmdAceptarArtic.Top = H - cmdAceptarArtic.Height - 120
'    cmdCancel(11).Top = H - cmdCancel(11).Height - 120
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

Private Sub CargaIconosAyuda()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    For Each Ima In Me.imgayuda
        Ima.Picture = frmPpal.imgListComun.ListImages(46).Picture
    Next
    Err.Clear
End Sub

