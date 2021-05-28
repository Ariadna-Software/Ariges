VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInformesNewOfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe "
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   15810
   Icon            =   "frmInformesNewOfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   15810
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
      Height          =   10410
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   15555
      Begin VB.Frame FrameInfClientesOrd 
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
         Left            =   10935
         TabIndex        =   82
         Top             =   6975
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CommandButton cmdBajar 
            Height          =   650
            Left            =   3105
            Picture         =   "frmInformesNewOfer.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   1395
            Width           =   510
         End
         Begin VB.CommandButton cmdSubir 
            Caption         =   " "
            Height          =   650
            Left            =   3105
            Picture         =   "frmInformesNewOfer.frx":108E
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   495
            Width           =   510
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Left            =   225
            TabIndex        =   85
            Top             =   495
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   2778
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
      Begin VB.Frame FrameInfClientesOpc 
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
         Height          =   6780
         Left            =   10935
         TabIndex        =   20
         Top             =   180
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Frame FrameVolumen 
            Caption         =   "Volumen de ventas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   2055
            Left            =   255
            TabIndex        =   78
            Top             =   495
            Visible         =   0   'False
            Width           =   4035
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
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   48
               Top             =   945
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
               Index           =   123
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   49
               Top             =   1425
               Width           =   1350
            End
            Begin VB.Label Label14 
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
               Index           =   24
               Left            =   330
               TabIndex        =   81
               Top             =   975
               Width           =   600
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   36
               Left            =   1200
               Picture         =   "frmInformesNewOfer.frx":2110
               Top             =   1425
               Width           =   240
            End
            Begin VB.Label Label14 
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
               Index           =   25
               Left            =   330
               TabIndex        =   80
               Top             =   1455
               Width           =   570
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   35
               Left            =   1200
               Picture         =   "frmInformesNewOfer.frx":219B
               Top             =   945
               Width           =   240
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Fechas cálculo"
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
               Index           =   26
               Left            =   210
               TabIndex        =   79
               Top             =   465
               Width           =   1470
            End
         End
         Begin VB.CheckBox chkVolumen 
            Caption         =   "Informe con volumen ventas"
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
            Left            =   255
            TabIndex        =   50
            Top             =   2775
            Width           =   3555
         End
         Begin VB.ComboBox cboOrdVolVta 
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
            ItemData        =   "frmInformesNewOfer.frx":2226
            Left            =   255
            List            =   "frmInformesNewOfer.frx":2230
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   3135
            Visible         =   0   'False
            Width           =   3435
         End
         Begin VB.CheckBox chkExportacion 
            Caption         =   "Formato exportación"
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
            Left            =   255
            TabIndex        =   76
            Top             =   4095
            Visible         =   0   'False
            Width           =   3075
         End
         Begin VB.Frame FrVolVetasCredito 
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   495
            Left            =   135
            TabIndex        =   73
            Top             =   3495
            Visible         =   0   'False
            Width           =   4215
            Begin VB.ComboBox cboClienteCredito 
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
               ItemData        =   "frmInformesNewOfer.frx":2251
               Left            =   960
               List            =   "frmInformesNewOfer.frx":2264
               Style           =   2  'Dropdown List
               TabIndex        =   74
               Top             =   120
               Width           =   2595
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Crédito"
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
               Index           =   27
               Left            =   120
               TabIndex        =   75
               Top             =   150
               Width           =   840
            End
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Poblacion / actividad"
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
            Left            =   255
            TabIndex        =   72
            Top             =   4935
            Width           =   3075
         End
         Begin VB.OptionButton optClienteLis 
            Caption         =   "Telefonos"
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
            Left            =   255
            TabIndex        =   71
            Top             =   4575
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.OptionButton optClienteLis 
            Caption         =   "Email"
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
            Left            =   1845
            TabIndex        =   70
            Top             =   4575
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.OptionButton optClienteLis 
            Caption         =   "F.Pago"
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
            Left            =   3015
            TabIndex        =   69
            Top             =   4575
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Image imgayuda 
            Height          =   255
            Index           =   0
            Left            =   3960
            ToolTipText     =   "Listados de clientes"
            Top             =   225
            Width           =   255
         End
      End
      Begin VB.Frame FrameInfClientesSel 
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
         Left            =   270
         TabIndex        =   15
         Top             =   180
         Visible         =   0   'False
         Width           =   10515
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
            Index           =   129
            Left            =   6540
            MaxLength       =   6
            TabIndex        =   42
            Top             =   5895
            Width           =   945
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
            Index           =   129
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "Text5"
            Top             =   5895
            Width           =   2760
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
            Index           =   130
            Left            =   6540
            MaxLength       =   6
            TabIndex        =   43
            Top             =   6300
            Width           =   945
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
            Index           =   130
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "Text5"
            Top             =   6300
            Width           =   2760
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
            Index           =   42
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "Text5"
            Top             =   6300
            Width           =   3135
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
            Index           =   41
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "Text5"
            Top             =   5895
            Width           =   3135
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
            Index           =   42
            Left            =   1140
            MaxLength       =   2
            TabIndex        =   41
            Top             =   6300
            Width           =   840
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
            Index           =   41
            Left            =   1140
            MaxLength       =   2
            TabIndex        =   40
            Top             =   5895
            Width           =   840
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
            Index           =   40
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "Text5"
            Top             =   4095
            Width           =   8220
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
            Index           =   39
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "Text5"
            Top             =   3690
            Width           =   8220
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
            Index           =   40
            Left            =   1140
            MaxLength       =   4
            TabIndex        =   37
            Top             =   4095
            Width           =   840
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
            Index           =   39
            Left            =   1140
            MaxLength       =   4
            TabIndex        =   36
            Top             =   3690
            Width           =   840
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
            Index           =   36
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "Text5"
            Top             =   1965
            Width           =   8220
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
            Index           =   35
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "Text5"
            Top             =   1560
            Width           =   8220
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
            Index           =   36
            Left            =   1140
            MaxLength       =   3
            TabIndex        =   33
            Top             =   1965
            Width           =   840
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
            Index           =   35
            Left            =   1140
            MaxLength       =   3
            TabIndex        =   32
            Top             =   1560
            Width           =   840
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
            Index           =   34
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "Text5"
            Top             =   945
            Width           =   8220
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
            Index           =   33
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "Text5"
            Top             =   540
            Width           =   8220
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
            Index           =   34
            Left            =   1140
            MaxLength       =   4
            TabIndex        =   31
            Top             =   945
            Width           =   840
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
            Index           =   33
            Left            =   1140
            MaxLength       =   4
            TabIndex        =   30
            Top             =   540
            Width           =   840
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
            Index           =   37
            Left            =   1140
            MaxLength       =   3
            TabIndex        =   34
            Top             =   2625
            Width           =   840
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
            Index           =   38
            Left            =   1140
            MaxLength       =   3
            TabIndex        =   35
            Top             =   3030
            Width           =   840
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
            Index           =   37
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "Text5"
            Top             =   2625
            Width           =   8220
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
            Index           =   38
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "Text5"
            Top             =   3030
            Width           =   8220
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
            Index           =   151
            Left            =   1140
            MaxLength       =   4
            TabIndex        =   38
            Top             =   4770
            Width           =   840
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
            Index           =   151
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "Text5"
            Top             =   4770
            Width           =   8220
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
            Index           =   152
            Left            =   1140
            MaxLength       =   4
            TabIndex        =   39
            Top             =   5160
            Width           =   840
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
            Index           =   152
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "Text5"
            Top             =   5160
            Width           =   8220
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   22
            Left            =   855
            Top             =   6315
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   21
            Left            =   855
            Top             =   5895
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   20
            Left            =   855
            Top             =   4110
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   19
            Left            =   855
            Top             =   3690
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   16
            Left            =   855
            Top             =   1995
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   15
            Left            =   855
            Top             =   1605
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   14
            Left            =   855
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   13
            Left            =   855
            Top             =   585
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   17
            Left            =   855
            Top             =   2625
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   18
            Left            =   855
            Top             =   3060
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   89
            Left            =   855
            Top             =   4770
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   90
            Left            =   855
            Top             =   5175
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
            Index           =   6
            Left            =   180
            TabIndex        =   68
            Top             =   5880
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
            Index           =   5
            Left            =   180
            TabIndex        =   67
            Top             =   6240
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
            Top             =   4755
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
            Index           =   17
            Left            =   180
            TabIndex        =   65
            Top             =   5115
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
            Index           =   19
            Left            =   180
            TabIndex        =   64
            Top             =   1590
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
            Index           =   20
            Left            =   180
            TabIndex        =   63
            Top             =   1950
            Width           =   645
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
            TabIndex        =   62
            Top             =   4065
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
            Index           =   3
            Left            =   180
            TabIndex        =   61
            Top             =   3690
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
            Index           =   8
            Left            =   180
            TabIndex        =   60
            Top             =   2985
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
            Index           =   9
            Left            =   180
            TabIndex        =   59
            Top             =   2610
            Width           =   600
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
            Index           =   0
            Left            =   180
            TabIndex        =   58
            Top             =   945
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
            Index           =   1
            Left            =   180
            TabIndex        =   57
            Top             =   585
            Width           =   690
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   69
            Left            =   6255
            Top             =   5940
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   70
            Left            =   6255
            Top             =   6300
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Situación"
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
            Index           =   45
            Left            =   180
            TabIndex        =   25
            Top             =   5580
            Width           =   975
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
            Index           =   3
            Left            =   5580
            TabIndex        =   24
            Top             =   5970
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
            Index           =   2
            Left            =   5580
            TabIndex        =   23
            Top             =   6330
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Código Postal"
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
            Index           =   101
            Left            =   5535
            TabIndex        =   22
            Top             =   5625
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Actividad"
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
            Index           =   49
            Left            =   180
            TabIndex        =   21
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Visitador"
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
            Index           =   127
            Left            =   180
            TabIndex        =   19
            Top             =   4455
            Width           =   960
         End
         Begin VB.Label Label4 
            Caption         =   "Zona"
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
            Index           =   48
            Left            =   180
            TabIndex        =   18
            Top             =   1260
            Width           =   915
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
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
            Index           =   47
            Left            =   180
            TabIndex        =   17
            Top             =   3375
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ruta"
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
            Index           =   51
            Left            =   180
            TabIndex        =   16
            Top             =   2295
            Width           =   510
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
         Left            =   270
         TabIndex        =   0
         Top             =   9675
         Width           =   1335
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
         Left            =   270
         TabIndex        =   3
         Top             =   6975
         Width           =   10515
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   585
            Width           =   6990
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
            TabIndex        =   8
            Top             =   1065
            Width           =   8310
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
            TabIndex        =   7
            Top             =   1545
            Width           =   8310
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   0
            Left            =   10095
            TabIndex        =   6
            Top             =   1065
            Width           =   255
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   1
            Left            =   10095
            TabIndex        =   5
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
            Left            =   8835
            TabIndex        =   4
            Top             =   585
            Width           =   1515
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
         Left            =   13995
         TabIndex        =   1
         Top             =   9630
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
         Left            =   12510
         TabIndex        =   14
         Top             =   9630
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInformesNewOfer"
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

'23.  Categorias

'58. Listado de proveedores

'110.  Ubicaciones

    '400:  Clientes potenciales.  Cartas
    '401:                "       Etiquetas
    '402        "               CRM

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto
Public EsHco As Boolean
    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Attribute frmMtoCartasOfe.VB_VarHelpID = -1
Private WithEvents frmMtoCliente As frmBasico2
Attribute frmMtoCliente.VB_VarHelpID = -1
Private WithEvents frmMtoProve As frmBasico2
Attribute frmMtoProve.VB_VarHelpID = -1
Private WithEvents frmMtoAgente As frmBasico2 '%=%=frmFacAgentesCom
Attribute frmMtoAgente.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmBasico2 'frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoActiv As frmFacActividades
Attribute frmMtoActiv.VB_VarHelpID = -1
Private WithEvents frmMtoZona As frmFacZonas
Attribute frmMtoZona.VB_VarHelpID = -1
Private WithEvents frmMtoRuta As frmFacRutas
Attribute frmMtoRuta.VB_VarHelpID = -1
Private WithEvents frmMtoSitua As frmFacSituaciones
Attribute frmMtoSitua.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmIncidencias
Attribute frmMtoIncid.VB_VarHelpID = -1
Private WithEvents frmMtoArtic As frmBasico2
Attribute frmMtoArtic.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmBasico2 'frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'codigo postal
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes para mostrar las etiquetas a imprimir
Attribute frmMen.VB_VarHelpID = -1

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
Private Titulo As String 'Titulo para la ventana frmImprimir
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
Dim SQL As String
Dim Sql2 As String
Dim Rc As String
Dim RC2 As String

Dim cDesde As String
Dim cHasta As String
Dim B As Boolean

    MontaSQL = False
    
    If Not DatosOk Then Exit Function
    
    Select Case OpcionListado
        Case 47 ' informe de clientes
'            If Not PonerDesdeHasta2("{sprove.codprove}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Código: ") Then Exit Function
    End Select
    
    AnyadirAFormula cadFormula, cDesde
    AnyadirAFormula cadSelect, Replace(Replace(cDesde, "{", ""), "}", "")
        
    B = False
    
    Select Case OpcionListado
        Case 47 ' informe de clientes
            B = MontaSqlInfClientes
    End Select
    
    MontaSQL = B
    
End Function


Private Function MontaSqlInfClientes() As Boolean
Dim campo As String, devuelve As String
Dim numOp As Byte
Dim B As Boolean
    
    MontaSqlInfClientes = False
    
    
    Codigo = ""
    If txtCodigo(33).Text <> "" Or txtCodigo(34).Text <> "" Then
        campo = "{sclien.codactiv}"
        'Parametro Desde/Hasta Actividad
        devuelve = " Actividad: "
        If Not PonerDesdeHasta(campo, "N", 33, 34, devuelve) Then Exit Function
        Codigo = devuelve
    End If
    If Me.chkVolumen.Value = 1 Then
        devuelve = ""
        If Me.txtCodigo(122).Text <> "" Then devuelve = devuelve & " desde " & Format(txtCodigo(122).Text, "dd/mm/yyyy")
        If Me.txtCodigo(123).Text <> "" Then devuelve = devuelve & " hasta " & Format(txtCodigo(123).Text, "dd/mm/yyyy")
        
        
        
        'Si tiene marcado credito
        If vParamAplic.OperacionesAseguradas Then
            If Me.cboClienteCredito.ListIndex > 0 Then
                devuelve = devuelve & " Credito:"
                If Me.cboClienteCredito.ListIndex = 1 Then
                    devuelve = devuelve & " Privado"
                    campo = "({sclien.credipriv} = 1)"
                    
                ElseIf Me.cboClienteCredito.ListIndex = 2 Then
                    devuelve = devuelve & " Aseguradora"
                    campo = "({sclien.credipriv} = 0)"
                ElseIf Me.cboClienteCredito.ListIndex = 3 Then
                    devuelve = devuelve & " Estudiado"
                    campo = "({sclien.credipriv} = 2)"
                Else
                    devuelve = devuelve & " NO asignado"
                    campo = "({sclien.credipriv} = 9)"
                End If
                If cadFormula <> "" Then cadFormula = cadFormula & " AND "
                If cadSelect <> "" Then cadSelect = cadSelect & " AND "
                cadFormula = cadFormula & campo
                cadSelect = cadSelect & campo
                
            End If
        End If
        devuelve = "              Fecha ventas: " & devuelve
        Codigo = Trim(Codigo & devuelve)
    End If
    If Codigo <> "" Then
        Codigo = "pDHActividad=""" & Codigo & """|"
        cadParam = cadParam & Codigo
        numParam = numParam + 1
    End If
    'Cadena para seleccion D/H ZONA
    '--------------------------------------------
     If txtCodigo(35).Text <> "" Or txtCodigo(36).Text <> "" Then
        campo = "{sclien.codzonas}"
        'Parametro Desde/Hasta Zona
        devuelve = "pDHZona=""Zona: "
        If Not PonerDesdeHasta(campo, "N", 35, 36, devuelve) Then Exit Function
    End If
    
    'Cadena para seleccion D/H RUTA
    '--------------------------------------------
     If txtCodigo(37).Text <> "" Or txtCodigo(38).Text <> "" Then
        campo = "{sclien.codrutas}"
        'Parametro Desde/Hasta Ruta
        devuelve = "pDHRuta=""Ruta: "
        If Not PonerDesdeHasta(campo, "N", 37, 38, devuelve) Then Exit Function
    End If
    
    'Cadena para seleccion D/H AGENTE    y VISITADOR
    '--------------------------------------------
    Titulo = ""
    If txtCodigo(39).Text <> "" Or txtCodigo(40).Text <> "" Then
        campo = "{sclien.codagent}"
        'Parametro Desde/Hasta Agente
        devuelve = "Agente: "
        If Not PonerDesdeHasta(campo, "N", 39, 40, devuelve) Then Exit Function
        Titulo = Replace(devuelve, """", "")
    End If
    If txtCodigo(151).Text <> "" Or txtCodigo(152).Text <> "" Then
        campo = "{sclien.visitador}"
        'Parametro Desde/Hasta Agente
        devuelve = "Visitador: "
        If Not PonerDesdeHasta(campo, "N", 151, 152, devuelve) Then Exit Function
        Titulo = Trim(Titulo & "    " & Replace(devuelve, """", ""))
    End If
    
    
    If Titulo <> "" Then
        devuelve = "pDHAgente="" " & Titulo & """|"
        cadParam = cadParam & devuelve
        numParam = numParam + 1
        Titulo = ""
    End If
    
    
    'Cadena para seleccion D/H SITUACION
    '--------------------------------------------
    Titulo = ""
    If txtCodigo(41).Text <> "" Or txtCodigo(42).Text <> "" Then
        campo = "{sclien.codsitua}"
        'Parametro Desde/Hasta Situacion
        'devuelve = "pDHSituacion=""Situación: "  '
        devuelve = "Situación: "
        If Not PonerDesdeHasta(campo, "N", 41, 42, devuelve) Then Exit Function
        Titulo = Replace(devuelve, """", "")
    End If
    
    
    If txtCodigo(129).Text <> "" Or txtCodigo(130).Text <> "" Then
            campo = "{sclien.codpobla}"
            'Parametro Desde/Hasta Agente
            devuelve = "C.Postal: "
            If Not PonerDesdeHasta(campo, "T", 129, 130, devuelve) Then Exit Function
            Titulo = Trim(Titulo & "    " & Replace(devuelve, """", ""))
    End If
 
    If Titulo <> "" Then
        devuelve = "pDHSituacion="" " & Titulo & """|"
        cadParam = cadParam & devuelve
        numParam = numParam + 1
        Titulo = ""
    End If
    
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
    If Me.chkVolumen.Value = 0 Then
        If Me.chkVolumen.Value = 0 Then
            numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
            numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
            numOp = PonerGrupo(3, ListView1.ListItems(3).Text)
            numOp = PonerGrupo(4, ListView1.ListItems(4).Text)
        End If
    End If


    cadSelect = cadFormula
    tabla = "sclien"


    MontaSqlInfClientes = True

End Function

Private Function DatosOk() As Boolean
Dim SQL As String
Dim B As Boolean
    
    B = True
    
    DatosOk = B

End Function


Private Sub chkVarios_Click(Index As Integer)
    If Index = 3 Then
        'Listado de clientes
        If Me.chkVarios(3).Value = 1 And chkVolumen.Value = 1 Then
            
            MsgBox "No puede marcar a la vez 'con volumen ventas' y 'poblacion'", vbExclamation
            chkVolumen.Value = 0
            chkVolumen_Click
        End If
    End If
End Sub

Private Sub chkVolumen_Click()
   
        'Listado de clientes
        If Me.chkVarios(3).Value = 1 And chkVolumen.Value = 1 Then
            MsgBox "No puede marcar a la vez 'con volumen ventas' y 'poblacion'", vbExclamation
            Me.chkVarios(3).Value = 0
        End If
 


    FrameVolumen.visible = chkVolumen.Value = 1
    cboOrdVolVta.visible = chkVolumen.Value = 1
    optClienteLis(0).visible = FrameVolumen.visible
    optClienteLis(1).visible = FrameVolumen.visible
    optClienteLis(2).visible = FrameVolumen.visible
    FrVolVetasCredito.visible = (chkVolumen.Value = 1) And vParamAplic.OperacionesAseguradas
    
    If Me.chkVolumen.Value = 0 Then Me.chkExportacion.visible = False

End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim B As Boolean

    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    If OpcionListado = 47 Then
        Screen.MousePointer = vbHourglass
        B = True
        If Me.chkVolumen.Value = 1 Then B = CalculaVolumenVtas_
        Screen.MousePointer = vbDefault
        
        If Not B Then Exit Sub
    
    End If
    
    
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

End Sub

Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    Select Case OpcionListado
        Case 47 ' listado de clientes
            ACInformeClientes
        
            
    End Select
    
    ImprimeGeneral
    
        
    If indCodigo = 1 Then OpcionListado = 6
    
    
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook OpcionListado
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub ACInformeClientes()
Dim campo As String, devuelve As String
Dim numOp As Byte
Dim B As Boolean

    If Me.chkVarios(3).Value = 1 Then
        'rFacClienCP.rpt
        NombreRPT = "rFacClienCP.rpt"
    Else
        If Me.chkVolumen.Value = 0 Then
            'ESTE ES EL NORMAL
            NombreRPT = "rFacClientes.rpt"
        Else
            'Añadimos codusu
            If cadFormula <> "" Then cadFormula = cadFormula & " AND "
            cadFormula = cadFormula & " ({tmpstockfec.codusu} = " & vUsu.Codigo & " )"
            
            'Añadimos el de emial
            devuelve = 0
            If Me.optClienteLis(1).Value Then devuelve = 1
            If Me.optClienteLis(2).Value Then devuelve = 2
            devuelve = "MuestrEmail=" & devuelve & "|"
            cadParam = cadParam & devuelve
            numParam = numParam + 1
        
            'Le calculo el volumen de ventas
            If cboOrdVolVta.ListIndex <= 0 Then
                NombreRPT = "rFacClienAgeVol.rpt"
                  
            Else
                If Me.chkExportacion.Value = 1 Then
                    NombreRPT = "rFacClienAgeExp.rpt"
                Else
                    NombreRPT = "rFacClienAgeVol2.rpt"
                End If
                
            End If
        End If
    End If
    cadNombreRPT = NombreRPT
    
End Sub

Private Sub AccionesCSV()
Dim SQL As String

    'Monto el SQL
    Select Case OpcionListado
        Case 47 ' listado de clientes
            SQL = "Select codprove AS Código,nomprove as Nombre, domprove as Domicilio, codpobla as CPostal, pobprove as Poblacion, proprove as Provincia, nifprove as NIF, telprov1 as Telefono, codmacta as Cuenta, maiprov1 as Email1  "
            SQL = SQL & " From sprove "
        
        
    End Select
    
    If cadSelect <> "" Then SQL = SQL & " WHERE " & cadSelect
    
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDeselTodos_Click(Index As Integer)
Dim I As Byte

    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = False
    Next I

End Sub

Private Sub cmdSelTodos_Click()
Dim I As Byte

    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = True
    Next I
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
        
        Select Case OpcionListado
            Case 47 '47: Informe de Clientes
                PonerFoco txtCodigo(33)
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
    For H = 13 To 22
        Me.imgBuscarOfer(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgBuscarOfer(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next H
    For H = 69 To 70
        Me.imgBuscarOfer(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgBuscarOfer(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next H
    For H = 89 To 90
        Me.imgBuscarOfer(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgBuscarOfer(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next H
    
     
    FrameCobros.visible = True
    
    '###Descomentar
'    CommitConexion
    
    FrameCobrosVisible True, H, W
    
    FrameInfClientesVisible False
    
    Select Case OpcionListado
        Case 47 ' informe de clientes
            FrameInfClientesVisible True
            Me.Caption = "Informe de Clientes"
            
            CargarListViewOrden
            indFrame = 6
            'Viloumen de ventas
            FrameVolumen.visible = False
            Me.chkVolumen.Value = 0
            'fijo el año actual
            txtCodigo(122).Text = "01/01/" & Year(Now)
            txtCodigo(123).Text = Format(Now, "dd/mm/yyyy")
            Label4(51).Caption = IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Asociación", "Ruta")
            
            'De momento el csv no lo vemos
            VerCSV False
        
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
        
    End Select
    
    Me.Height = Me.FrameCobros.Height
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

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de cod Postal
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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

Private Sub frmMtoActiv_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Actividades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAgente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agentes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCliente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Incidencias
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProve_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoRuta_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Rutas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoSitua_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Situaciones
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Trabajadores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoZona_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Zonas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub




Private Sub imgAyuda_Click(Index As Integer)
Dim Ayuda As String

    'Sera las ayuda. Tampoco queiero la biblia, pero,
    'si un "pelin" de ayuda no me vendria mal a mi, imaginemos a el cliente final
    Codigo = vbCrLf & Space(10)
    Select Case Index
    Case 0
        Ayuda = ""
        Ayuda = Ayuda & vbCrLf & " --> VOLUMEN "
        Ayuda = Ayuda & Codigo & "Agrupado por agente muestra ademas de los datos basicos"
        Ayuda = Ayuda & Codigo & "el volumen de ventas entre las fechas seleccionadas "
        Ayuda = Ayuda & Codigo & "y el credito que tenga. Si marca agrupado no 'salta'"
        Ayuda = Ayuda & Codigo & "por zona,ruta."
        Ayuda = Ayuda & Codigo & "  -Telefonos / mail / Forma de pago:  Muestra en listado, o telefonos o email o la forma de pago(en el agrupado) "
        Ayuda = Ayuda & Codigo & "  -Formato exportacion: facilita exportación excel"
        
        Ayuda = Ayuda & vbCrLf & vbCrLf & " --> Poblacion / actividad "
        'Ordenado por codpobla,activadad. Solo rompe por codpostal
        Ayuda = Ayuda & Codigo & "Agrupado por codigo postal, muestra los datos basicos y la actividad"
        
        
        Ayuda = Ayuda & vbCrLf & vbCrLf & " --> Normal "
        Ayuda = Ayuda & Codigo & "Ordenado segun la seleccion mostrara los datos basicos:"
        Ayuda = Ayuda & Codigo & "ruta,zona,agente,codigo,nombre,domicilio,nif,telefono"
        
    
    End Select
    Ayuda = imgayuda(Index).ToolTipText & vbCrLf & String(46, "=") & vbCrLf & Ayuda
    MsgBox Ayuda, vbInformation
End Sub


Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1, 39, 40, 45, 60, 77 'Cod. Carta
            Select Case Index
                Case 0: indCodigo = 2
                Case 1: indCodigo = 13
                Case 39: indCodigo = 63
                Case 40: indCodigo = 64
                Case 45: indCodigo = 81
                Case 60: indCodigo = 116
                Case 77: indCodigo = 138
            End Select
            
            Set frmMtoCartasOfe = New frmFacCartasOferta
            frmMtoCartasOfe.DatosADevolverBusqueda = "0|1|"
            frmMtoCartasOfe.Show vbModal
            Set frmMtoCartasOfe = Nothing
            
        Case 2, 3, 9, 10, 23, 24, 46, 47, 52, 53, 56, 57, 63, 64, 78, 79, 85, 86, 93, 94, 99, 100, 147, 148 'Cod. CLIENTE
            Select Case Index
                Case 2, 3: indCodigo = 7 + Index
                Case 9, 10: indCodigo = 18 + Index
                Case 23, 24: indCodigo = Index + 20
                Case 46, 47: indCodigo = Index + 33
                Case 52, 53: indCodigo = Index + 44
                Case 56, 57: indCodigo = Index + 54
                Case 63, 64: indCodigo = Index + 57
                Case 78, 79: indCodigo = Index + 61
                Case 85, 86, 93, 94: indCodigo = Index + 62
                Case 99, 100: indCodigo = Index + 63
            End Select
            
            
            Set frmMtoCliente = New frmBasico2
            AyudaClientes frmMtoCliente, txtCodigo(indCodigo).Text
            Set frmMtoCliente = Nothing
    
            
            
        Case 4, 5, 6, 7, 11, 12, 19, 20, 25, 26, 80, 81, 87, 88, 89, 90 'Cod. AGENTE
            Select Case Index
                Case 4, 5: indCodigo = 7 + Index
                Case 5: indCodigo = 12
                Case 6, 7: indCodigo = 12 + Index
                Case 11, 12: indCodigo = 18 + Index
                Case 19, 20, 25, 26: indCodigo = 20 + Index
                Case 80, 81: indCodigo = Index + 61
                Case 87, 88, 89, 90: indCodigo = Index + 62
            End Select
            If OpcionListado <> 92 Then
'                Set frmMtoAgente = New frmFacAgentesCom
'                frmMtoAgente.DatosADevolverBusqueda = "0|1|"
'                frmMtoAgente.Show vbModal
                Set frmMtoAgente = New frmBasico2
                AyudaAgentesComerciales frmMtoAgente, txtCodigo(indCodigo), , True
                Set frmMtoAgente = Nothing
            ElseIf Index = 6 Or Index = 7 Then 'Gastos financieros (trabajador)
'                Set frmMtoTraba = New frmAdmTrabajadores
'                frmMtoTraba.DatosADevolverBusqueda = "0|1|"
'                frmMtoTraba.Show vbModal
                Set frmMtoTraba = New frmBasico2
                AyudaTrabajadores frmMtoTraba, txtCodigo(indCodigo)
                Set frmMtoTraba = Nothing
            End If
            
        Case 8, 27, 28, 61, 62 'cod. TRABAJADOR
            indCodigo = 24
            If Index = 27 Then
                indCodigo = 47
            ElseIf Index = 28 Then indCodigo = 51
            ElseIf Index > 28 Then indCodigo = (117 + 61) - Index
            End If
'            Set frmMtoTraba = New frmAdmTrabajadores
'            frmMtoTraba.DatosADevolverBusqueda = "0|1|"
'            frmMtoTraba.Show vbModal
            Set frmMtoTraba = New frmBasico2
            AyudaTrabajadores frmMtoTraba, txtCodigo(indCodigo)
            Set frmMtoTraba = Nothing
            
        Case 13, 14, 30, 31, 67, 68, 75, 76, 82, 83 'cod. ACTIVIDAD
            indCodigo = 20 + Index
            If Index = 30 Or Index = 31 Then indCodigo = Index + 23
            'If Index = 67 Or Index = 68 Then indCodigo = Index + 60
            If Index >= 67 Then indCodigo = Index + 61
            
            
            Set frmMtoActiv = New frmFacActividades
            frmMtoActiv.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmMtoActiv.Show vbModal
            Set frmMtoActiv = Nothing
            
        Case 15, 16, 91, 92 'cod. ZONA
            If Index < 91 Then
                indCodigo = 20 + Index
            Else
                indCodigo = 62 + Index
           
            End If
            Set frmMtoZona = New frmFacZonas
            frmMtoZona.DatosADevolverBusqueda = "0|1|"
            frmMtoZona.Show vbModal
            Set frmMtoZona = Nothing
            
         Case 17, 18, 95, 96 'cod. RUTA
            If Index < 95 Then
                indCodigo = 20 + Index
            Else
                indCodigo = 62 + Index
            End If
            Set frmMtoRuta = New frmFacRutas
            frmMtoRuta.DatosADevolverBusqueda = "0|1|"
            frmMtoRuta.Show vbModal
            Set frmMtoRuta = Nothing
            
        Case 21, 22, 34, 84 'cod. SITUACION
            indCodigo = 20 + Index
            If Index = 34 Then indCodigo = Index + 23
            If Index = 84 Then indCodigo = Index + 61
            Set frmMtoSitua = New frmFacSituaciones
            frmMtoSitua.DatosADevolverBusqueda = "0|1|"
            frmMtoSitua.Show vbModal
            Set frmMtoSitua = Nothing
            
        Case 29 'INCIDENCIAS
            indCodigo = 52
            Set frmMtoIncid = New frmIncidencias
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            txtCodigo(indCodigo).Text = ""
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
        Case 32, 33, 37, 38, 69, 70, 73, 74 'Cod POSTAL
            If Index < 69 Then
                indCodigo = Index + 23
            Else
                indCodigo = Index + 60
            End If
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0|1|"
            txtCodigo(indCodigo).Text = ""
            frmCP.Show vbModal
            Set frmCP = Nothing
            
        Case 35, 36, 41, 42, 48, 49, 65, 66 'cod. PROVEEDOR
            Select Case Index
                Case 35, 36: indCodigo = Index + 23
                Case 41, 42: indCodigo = Index + 24
                Case 48, 49: indCodigo = Index + 42
                Case 65, 66: indCodigo = Index + 59
            End Select
'            If Index = 35 Or Index = 36 Then indCodigo = Index + 23
'            If Index = 41 Or Index = 42 Then indCodigo = Index + 24
'            If Index = 48 Or Index = 49 Then indCodigo = Index + 42
            Set frmMtoProve = New frmBasico2
'            frmMtoProve.DatosADevolverBusqueda = "0|1|"
'            frmMtoProve.Show vbModal
            AyudaProveedores frmMtoProve, txtCodigo(indCodigo)
            Set frmMtoProve = Nothing
            
        Case 43, 44, 58, 59 'cod. ARTICULO
            If Index <= 44 Then
                indCodigo = Index + 24
            Else
                indCodigo = Index + 54  'En listado de vetnas x familia articulo
            End If
            Set frmMtoArtic = New frmBasico2
            'frmMtoArtic.DatosADevolverBusqueda3 = "@1@" 'Abrimos en Modo Busqueda
'            frmMtoArtic.DatosADevolverBusqueda = "0|1|"
'            frmMtoArtic.Show vbModal
            AyudaArticulos frmMtoArtic, txtCodigo(indCodigo)
            Set frmMtoArtic = Nothing
            
        Case 50, 51, 54, 55 'Cod. FAMILIA articulo
            Select Case Index
                Case 50, 51: indCodigo = Index + 44
                Case 54, 55: indCodigo = Index + 46
            End Select
'            Set frmMtoFamilia = New frmAlmFamiliaArticulo
'            frmMtoFamilia.DatosADevolverBusqueda = "0|1|"
'            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = New frmBasico2
            AyudaFamilias frmMtoFamilia, txtCodigo(indCodigo)
            Set frmMtoFamilia = Nothing
        Case 71, 72
            'Clientes potenciales
            AbrirBuscaGrid Index
            
        Case 97, 98
            'Bancos propios   - Formas pago
            AbrirBuscaGrid Index
            
        End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 1 'frameOfertas (indFrame=6)
            indCodigo = 3 'Desde
        Case 2 'frameOfertas (indFrame=6)
            indCodigo = 4 'Hasta
        Case 3 'frameRecordatorio Oferta
            indCodigo = 7 '(Desde)
        Case 4 'frameRecordatorio Oferta
            indCodigo = 8 '(Hasta)
        Case 5 'frameEfectuadas
            indCodigo = 16 'Desde
        Case 6 'frameEfectuadas
            indCodigo = 17 'Hasta
        Case 7 'frameTraspasoHco
            indCodigo = 22 'Desde
        Case 8 'frameTraspasoHco
            indCodigo = 23 'hasta
        Case 9, 10 'FrameGenerarPedido
            indCodigo = Index + 16
        Case 11, 12 'Frame Clientes Inactivos
            indCodigo = 20 + Index
        Case 13 'frame pasar pedido a Albaran de compras (a proveedor)
            indCodigo = 49
        Case 14
            indCodigo = 50
        Case 15, 16
            indCodigo = Index + 54
        Case 17 'Frame Factura Rectificariva
            indCodigo = 72
        Case 18, 19 'Ped. Compras
            indCodigo = Index + 56
        Case 20, 21 'Carta Pedidos
            indCodigo = Index + 57
        Case 22: indCodigo = Index + 60
        Case 23, 24 'Reimprimir facturas
            indCodigo = Index + 62
        Case 25, 26 'Cierre caja TPV
            indCodigo = Index + 63
        Case 27, 28 'Listados estadistica compras
            indCodigo = Index + 65
        Case 29, 30 'Estadistica ventas por familia
            indCodigo = Index + 69
   
        Case 31, 32 'Impresion etiq. clientes. Desde / hasta factura
            indCodigo = Index + 73
        Case 33, 34
            indCodigo = Index + 75
            
        Case 35, 36
            indCodigo = Index + 87
        Case 37
            indCodigo = 160
        Case 38, 39, 40, 41
            indCodigo = 126 + Index
        Case 42
            indCodigo = 168
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
            'listado de clientes
            Case 33: KEYBusqueda KeyAscii, 13 'actividad desde
            Case 34: KEYBusqueda KeyAscii, 14 'actividad hasta
            Case 35: KEYBusqueda KeyAscii, 15 'zona desde
            Case 36: KEYBusqueda KeyAscii, 16 'zona hasta
            Case 37: KEYBusqueda KeyAscii, 17 'ruta desde
            Case 38: KEYBusqueda KeyAscii, 18 'ruta hasta
            Case 39: KEYBusqueda KeyAscii, 19 'agente desde
            Case 40: KEYBusqueda KeyAscii, 20 'agente hasta
            Case 41: KEYBusqueda KeyAscii, 21 'situacion desde
            Case 42: KEYBusqueda KeyAscii, 22 'situacion hasta
            Case 151: KEYBusqueda KeyAscii, 89 'visitador desde
            Case 152: KEYBusqueda KeyAscii, 90 'visitador hasta
            Case 129: KEYBusqueda KeyAscii, 69 'poblacion desde
            Case 130: KEYBusqueda KeyAscii, 70 'poblacion hasta
            Case 122: KEYFecha KeyAscii, 35 'fecha desde
            Case 123: KEYFecha KeyAscii, 36 'fecha hasta
        
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscarOfer_Click (Indice)
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
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        'FECHA Desde Hasta
        Case 3, 4, 7, 8, 16, 17, 22, 23, 25, 26, 31, 32, 49, 50, 69, 70, 72, 74, 75, 77, 78, 82, 85, 86, 88, 89, 92, 93, 98, 99, 104, 105, 108, 109, 122, 123, 160, 164, 165, 166, 167, 168
            If txtCodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtCodigo(Index)
            
            'Fecha entrega para Pedido. Poner la semana
            If Index = 26 Then
                'Comprobar que fecha entrega es posterior a la del pedido
                If Not EsFechaIgualPosterior(txtCodigo(25).Text, txtCodigo(26).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then
                    txtCodigo(Index).Text = ""
                    PonerFoco txtCodigo(Index)
                Else
                    If Not IsDate(txtCodigo(26).Text) Then
                        txtNombre(4).Text = ""
                    Else
                        txtNombre(4).Text = CalculaSemana(CDate(txtCodigo(26).Text))
                    End If
                End If
            End If
            
        Case 2, 13, 63, 64, 81, 116, 138 'CARTA de la Oferta
            EsNomCod = True
            tabla = "scartas"
            codCampo = "codcarta"
            NomCampo = "descarta"
            Formato = "000"
            Titulo = "cod. de Carta"
                    
        Case 9, 10, 27, 28, 43, 44, 79, 80, 96, 97, 110, 111, 120, 121, 139, 140, 147, 148, 155, 156, 162, 163 'Cod. CLIENTE
            EsNomCod = True
            tabla = "sclien"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"

        Case 11, 12, 18, 19, 29, 30, 39, 40, 45, 46, 80, 81, 141, 142, 149, 150, 151, 152 'Cod. AGENTE
            EsNomCod = True
            Formato = "0000"
            If OpcionListado = 92 Then 'Gastos tecnicos
                If Index = 18 Or Index = 19 Then
                    'cod agente / cod. trabajador
                    tabla = "straba"
                    codCampo = "codtraba"
                    NomCampo = "nomtraba"
                    Titulo = "Trabajador"
                End If
            Else
                tabla = "sagent"
                codCampo = "codagent"
                NomCampo = "nomagent"
                Titulo = "Agente"
            End If
        
        Case 24, 47, 51, 117, 118 'Cod. TRABAJADOR
            EsNomCod = True
            tabla = "straba"
            codCampo = "codtraba"
            NomCampo = "nomtraba"
            Formato = "0000"
            Titulo = "Trabajador"
            
        Case 33, 34, 53, 54, 127, 128, 136, 137, 143, 144 'Cod ACTIVIDAD
            EsNomCod = True
            tabla = "sactiv"
            codCampo = "codactiv"
            NomCampo = "nomactiv"
            Formato = "000"
            Titulo = "Actividad de Cliente"
            
        Case 35, 36, 153, 154 'cod ZONA
            EsNomCod = True
            tabla = "szonas"
            codCampo = "codzonas"
            NomCampo = "nomzonas"
            Formato = "000"
            
        Case 37, 38, 157, 158, 162, 163 'cod RUTA
            EsNomCod = True
            tabla = "srutas"
            codCampo = "codrutas"
            NomCampo = "nomrutas"
            Formato = "000"
            Titulo = IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Asociacion", "Ruta de Asistencia")
                        
        Case 41, 42, 57, 145 'cod SITUACION
            EsNomCod = True
            tabla = "ssitua"
            codCampo = "codsitua"
            NomCampo = "nomsitua"
            Formato = "00"
            Titulo = "Situación Especial"
            
        Case 52 'cod. Incidencias
            EsNomCod = True
            tabla = "sincid"
            codCampo = "codincid"
            NomCampo = "nomincid"
            TipCampo = "T"
            Titulo = "Incidencias"
            
        Case 55, 56, 60, 61, 129, 130, 133, 134 'cod POSTAL
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "scpostal", "provincia", "cpostal", "CPostal")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = txtCodigo(Index).Text
            
         Case 58, 59, 65, 66, 90, 91, 124, 125 'Cod. PROVEEDOR
            EsNomCod = True
            tabla = "sprove"
            codCampo = "codprove"
            NomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
            
        Case 67, 68, 112, 113 'cod. ARTICULO
            EsNomCod = True
            tabla = "sartic"
            codCampo = "codartic"
            NomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
            
        Case 94, 95, 100, 101 'cod. FAMILIA articulos
            EsNomCod = True
            tabla = "sfamia"
            codCampo = "codfamia"
            NomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
        Case 126
            If Me.txtCodigo(Index).Text <> "" Then
                If Not PonerFormatoDecimal(txtCodigo(Index), 1) Then txtCodigo(Index).Text = ""
            End If
        Case 131, 132 'cliente potencial
            EsNomCod = True
            tabla = "sclipot"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Cli. potenciales"
        Case 159
            'Bancos propios
            EsNomCod = True
            tabla = "sbanpr"
            codCampo = "codbanpr"
            NomCampo = "nombanpr"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Bancos"
        Case 161
            'Bancos propios
            EsNomCod = True
            tabla = "sforpa"
            codCampo = "codforpa"
            NomCampo = "nomforpa"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Forma pago"
        
    End Select
    
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, Titulo, TipCampo)
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
            Else
                txtNombre(Index).Text = ""
                txtCodigo(Index).Text = "" 'Puesto el 25 de enero
            End If

            
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, tabla, NomCampo, codCampo, Titulo, TipCampo)
'            If tabla = "sincid" Then
'                If txtNombre(Index).Text = "" Then txtCodigo(Index).Text = ""
'            End If
            
        End If
    End If
    
    
    
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


Private Sub FrameInfClientesVisible(visible As Boolean)
    FrameInfClientesSel.visible = visible
    FrameInfClientesOpc.visible = visible
    FrameInfClientesOrd.visible = visible
    
    FrameInfClientesSel.Enabled = visible
    FrameInfClientesOpc.Enabled = visible
    FrameInfClientesOrd.Enabled = visible
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
        'Van de 10 en 10
        Aux = "Marcas|Almacenes Propios|Tipos Unidad|Tipos Artículos|||Movimientos Almacen|Traspaso Almacen|Movimientos Articulos||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||Categorias||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||Clientes||||"
        Aux = Aux & "|||||||Proveedores|||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "|||||||||Ubicaciones|"

        
            
        Aux = RecuperaValor(Aux, outTipoDocumento) & ".pdf"
        If Aux = ".pdf" Then Aux = "Documento.pdf"
        
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
        'Informe
        Aux = "Marcas|AlmPropios|TiposUnidad|TipoArtículos|||MovimientosAlmacen|TraspasoAlmacen|MovimientosArticulos||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||Categorias||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||Clientes||||"
        Aux = Aux & "|||||||Proveedores|||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "||||||||||"
        Aux = Aux & "|||||||||Ubicaciones|"
        '--------------------------------------------------
        Aux = RecuperaValor(Aux, outTipoDocumento)
        
        If Aux = "" Then Aux = "Informe/documento"
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
Dim I  As Integer

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
        For I = 0 To miRsAux.Fields.Count - 1
            cadSQL = cadSQL & ";""" & miRsAux.Fields(I).Name & """"
        Next I
        Print #NF, Mid(cadSQL, 2)
    
    
        'Lineas
        While Not miRsAux.EOF
            cadSQL = ""
            For I = 0 To miRsAux.Fields.Count - 1
                cadSQL = cadSQL & ";""" & DBLet(miRsAux.Fields(I).Value, "T") & """"
            Next I
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
Dim SQL As String

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
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", 1000
    ListView1.ColumnHeaders.Add , , "Descripción", 2650
    
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

Private Sub AbrirFrmClientes()
'Clientes

    

            Set frmMtoCliente = New frmBasico2
            AyudaClientes frmMtoCliente, txtCodigo(indCodigo).Text
            Set frmMtoCliente = Nothing
    
    
    
    
End Sub

Private Function PonerFormulaYParametrosInfMovArt() As Boolean
Dim Cad As String
Dim todosMarcados As Boolean
Dim devuelve As String
Dim I As Byte

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
                    devuelve = Codigo & " = """ & Me.ListView1.ListItems(I).Text & """"
                Else
                    devuelve = devuelve & " or " & Codigo & " = """ & Me.ListView1.ListItems(I).Text & """"
                End If
            End If
        Next I

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
            For I = 1 To ListView1.ListItems.Count
                If Cad = "" Then
                    Cad = """" & ListView1.ListItems(I).Text & """"
                Else
                    Cad = Cad & ", """ & ListView1.ListItems(I).Text & """"
                End If
            Next I
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
'Carga el List View del frame: frameClientes
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Actividad, Zona, Ruta, Agente, Situación
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Campo", 1500

    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Actividad"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Zona"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = IIf(vParamAplic.NumeroInstalacion <> vbHerbelca, "Ruta", "Asociacion")
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Agente"
End Sub


Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView1
End Sub

Private Sub cmdSubir_Click()
'Subir el item seleccionado del listview2 una posicion
    SubirItemList Me.ListView1
End Sub

Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
'Pone por que campos se van a AGrupar los datos en el Informe de Crystal Report
'El informe tiene definido 4 formulas a las cuales ahora le asignamos un campo
'de la tabla segun el orden seleccionado para el agrupamiento
Dim campo As String
Dim NomCampo As String

    campo = "pGroup" & numGrupo & "="
    NomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Actividad"
            cadParam = cadParam & campo & "{sclien.codactiv}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""ACTIVIDAD:  "" & " & " totext({sclien.codactiv},""000"") & " & """  """ & " & {sactiv.nomactiv}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codactiv},""000"") & " & """ """ & " & {sactiv.nomactiv}" & "|"
                cadParam = cadParam & NomCampo & "{sactiv.nomactiv}" & "|"
                cadParam = cadParam & "pTitulo" & numGrupo & "=""Actividad""" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
            
        Case "Zona"
            cadParam = cadParam & campo & "{sclien.codzonas}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""ZONA:  "" & " & " totext({sclien.codzonas},""000"") & " & """  """ & " & {szonas.nomzonas}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codzonas},""000"") & " & """ """ & " & {szonas.nomzonas}" & "|"
                cadParam = cadParam & NomCampo & "{szonas.nomzonas}" & "|"
                cadParam = cadParam & "pTitulo" & numGrupo & "=""Zona""" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
            
        Case "Ruta", "Asociacion"
            cadParam = cadParam & campo & "{sclien.codrutas}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " """ & UCase(cadgrupo) & ":  "" & " & " totext({sclien.codrutas},""000"") & " & """  """ & " & {srutas.nomrutas}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codrutas},""000"") & " & """ """ & " & {srutas.nomrutas}" & "|"
                cadParam = cadParam & NomCampo & "{srutas.nomrutas}" & "|"
              '  cadParam = cadParam & "pTitulo" & numGrupo & "=""Ruta""" & "|"
                 cadParam = cadParam & "pTitulo" & numGrupo & "=""" & IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Asociacion", "Ruta") & """" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
'            PonerGrupo = numGrupo
        Case "Agente"
            cadParam = cadParam & campo & "{sclien.codagent}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""AGENTE:  "" & " & " totext({sclien.codagent},""000000"") & " & """  """ & " & {sagent.nomagent}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codagent},""000000"") & " & """ """ & " & {sagent.nomagent}" & "|"
                cadParam = cadParam & NomCampo & "{sagent.nomagent}" & "|"
                cadParam = cadParam & "pTitulo" & numGrupo & "=""Agente""" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
'        Case "Situacion"
    End Select
End Function


Private Sub AbrirBuscaGrid(OP As Integer)
Dim indT As Integer
    Set frmB = New frmBuscaGrid
    cadFormula = "" 'Aqui metera el valor
    Select Case OP
    Case 71, 72
        indT = OP + 60
        frmB.vCampos = "Codigo|sclipot|codclien|T||20·Descripción|sclipot|nomclien|T||70·"
        frmB.vTabla = "sclipot"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Clientes potenciales"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
    
    Case 97
        'Banocs propios
        indT = 159
        frmB.vCampos = "Codigo|sbanpr|codbanpr|T||20·Descripción|sbanpr|nombanpr|T||70·"
        frmB.vTabla = "sbanpr"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "BANCOS"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
        
    Case 98
        indT = 161
        frmB.vCampos = "Codigo|sforpa|codforpa|T||20·Descripción|sforma|nomforpa|T||70·"
        frmB.vTabla = "sforpa"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Forma Pago"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
    End Select
    frmB.Show vbModal
    Set frmB = Nothing
    
    If cadFormula <> "" Then
        txtCodigo(indT).Text = Format(RecuperaValor(cadFormula, 1), "0000")
        txtNombre(indT).Text = RecuperaValor(cadFormula, 2)
    End If
End Sub


Private Function CalculaVolumenVtas_() As Boolean

On Error GoTo ECalculaVolumenVtas_
    CalculaVolumenVtas_ = False
    
    Codigo = "DELETE FROM tmpstockfec WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    'codigo = "Select " & vUsu.codigo & ",'DAVID' ,codclien,sum(totalfac)"  VA sin IVA
    Codigo = "Select " & vUsu.Codigo & ",'DAVID' ,codclien,sum(baseimp1 + if(baseimp2 is null, 0,baseimp2) + if(baseimp3 is null, 0,baseimp3))"
    Codigo = Codigo & " from scafac where codtipom<>'FAZ'"
    If Me.txtCodigo(122).Text <> "" Then Codigo = Codigo & " AND fecfactu>=" & DBSet(txtCodigo(122).Text, "F")
    If Me.txtCodigo(123).Text <> "" Then Codigo = Codigo & " AND fecfactu<=" & DBSet(txtCodigo(123).Text, "F")
    If cadSelect <> "" Then Codigo = Codigo & " AND codclien IN (Select codclien from sclien WHERE " & cadSelect & ")"
    Codigo = Codigo & " GROUP BY 1,2,3"
    
    Codigo = "INSERT INTO tmpstockfec(codusu,codartic,codalmac,stock)  " & Codigo
    conn.Execute Codigo
    
    
    'Metere los que no hubieran facturado
    Codigo = "Select " & vUsu.Codigo & ",'DAVID',codclien,0 from sclien WHERE 1=1 "
    If cadSelect <> "" Then Codigo = Codigo & " AND " & cadSelect
    Codigo = "INSERT IGNORE INTO tmpstockfec(codusu,codartic,codalmac,stock) " & Codigo
    conn.Execute Codigo
    
    
    CalculaVolumenVtas_ = True
ECalculaVolumenVtas_:
    If Err.Number <> 0 Then MuestraError Err.Number, "Calculando volumen ventas"
End Function



