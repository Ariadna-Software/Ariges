VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInformesNew3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe "
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   12300
   Icon            =   "frmInformesNew3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
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
      TabIndex        =   2
      Top             =   120
      Width           =   12090
      Begin VB.Frame FrameClientesVOpc 
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
         Height          =   5325
         Left            =   7335
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Frame FrameOrden2 
            Caption         =   "Orden 2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   450
            TabIndex        =   48
            Top             =   2430
            Width           =   3705
            Begin VB.OptionButton optVarios 
               Caption         =   "Nombre"
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
               Index           =   4
               Left            =   2220
               TabIndex        =   50
               Top             =   540
               Value           =   -1  'True
               Width           =   1290
            End
            Begin VB.OptionButton optVarios 
               Caption         =   "NIF"
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
               Index           =   5
               Left            =   270
               TabIndex        =   49
               Top             =   540
               Width           =   1290
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Orden 1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   405
            TabIndex        =   45
            Top             =   720
            Width           =   3750
            Begin VB.OptionButton optVarios 
               Caption         =   "Poblacion"
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
               Left            =   270
               TabIndex        =   47
               Top             =   495
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton optVarios 
               Caption         =   "NIF"
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
               Left            =   2265
               TabIndex        =   46
               Top             =   540
               Width           =   975
            End
         End
      End
      Begin VB.Frame FrameClientesVSel 
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
         Height          =   2760
         Left            =   225
         TabIndex        =   38
         Top             =   270
         Width           =   6915
         Begin VB.TextBox txtNumEntero 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1125
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   810
            Width           =   1095
         End
         Begin VB.TextBox txtNumEntero 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1125
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   1215
            Width           =   1095
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
            TabIndex        =   41
            Top             =   1185
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
            TabIndex        =   40
            Top             =   810
            Width           =   690
         End
         Begin VB.Label Label3 
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
            Height          =   285
            Index           =   3
            Left            =   225
            TabIndex        =   39
            Top             =   405
            Width           =   3120
         End
      End
      Begin VB.Frame FrameFamiliasOpc 
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
         Height          =   5325
         Left            =   7335
         TabIndex        =   37
         Top             =   270
         Width           =   4455
         Begin VB.CheckBox chkVarios 
            Caption         =   "Mostrar familias inactivas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   315
            TabIndex        =   27
            Top             =   1800
            Width           =   3630
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Mostrar descuentos"
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
            Left            =   315
            TabIndex        =   25
            Top             =   720
            Width           =   3510
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Dto particulares"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   315
            TabIndex        =   26
            Top             =   1215
            Width           =   3630
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
         TabIndex        =   0
         Top             =   5760
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
         Left            =   225
         TabIndex        =   4
         Top             =   3105
         Width           =   6915
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   1545
            Width           =   4665
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   0
            Left            =   6450
            TabIndex        =   7
            Top             =   1065
            Width           =   255
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   1
            Left            =   6450
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   585
            Width           =   1515
         End
      End
      Begin VB.Frame frameConceptoOpc 
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
         Height          =   5325
         Left            =   7335
         TabIndex        =   3
         Top             =   270
         Width           =   4455
         Begin VB.Frame frameOrdenar 
            Caption         =   "Ordenar por"
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
            Height          =   735
            Left            =   315
            TabIndex        =   15
            Top             =   495
            Width           =   3690
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
               Left            =   300
               TabIndex        =   17
               Top             =   240
               Width           =   1215
            End
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
               Left            =   1815
               TabIndex        =   16
               Top             =   240
               Width           =   1455
            End
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
         Left            =   10440
         TabIndex        =   1
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
         TabIndex        =   29
         Top             =   5805
         Width           =   1335
      End
      Begin VB.Frame FrameFamiliasSel 
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
         Height          =   2760
         Left            =   225
         TabIndex        =   18
         Top             =   270
         Width           =   6915
         Begin VB.TextBox txtFamia 
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
            Left            =   1215
            MaxLength       =   6
            TabIndex        =   24
            Top             =   2160
            Width           =   890
         End
         Begin VB.TextBox txtFamia 
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
            Index           =   6
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   23
            Top             =   1770
            Width           =   890
         End
         Begin VB.TextBox txtDescFamia 
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
            Index           =   7
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "Text5"
            Top             =   2160
            Width           =   4575
         End
         Begin VB.TextBox txtDescFamia 
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
            Index           =   6
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "Text5"
            Top             =   1770
            Width           =   4575
         End
         Begin VB.TextBox txtCodProve 
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
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   22
            Top             =   1125
            Width           =   890
         End
         Begin VB.TextBox txtCodProve 
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
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   21
            Top             =   735
            Width           =   890
         End
         Begin VB.TextBox txtDescProve 
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
            Index           =   12
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "Text5"
            Top             =   1125
            Width           =   4575
         End
         Begin VB.TextBox txtDescProve 
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
            Index           =   11
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "Text5"
            Top             =   735
            Width           =   4575
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
            Index           =   1
            Left            =   225
            TabIndex        =   36
            Top             =   1440
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
            Index           =   3
            Left            =   225
            TabIndex        =   35
            Top             =   1755
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
            Index           =   2
            Left            =   225
            TabIndex        =   34
            Top             =   2130
            Width           =   645
         End
         Begin VB.Image imgFamilia 
            Height          =   240
            Index           =   7
            Left            =   915
            MouseIcon       =   "frmInformesNew3.frx":000C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar familia"
            Top             =   2160
            Width           =   240
         End
         Begin VB.Image imgFamilia 
            Height          =   240
            Index           =   6
            Left            =   945
            MouseIcon       =   "frmInformesNew3.frx":015E
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar familia"
            Top             =   1755
            Width           =   240
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
            Index           =   0
            Left            =   225
            TabIndex        =   31
            Top             =   405
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
            TabIndex        =   30
            Top             =   720
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
            TabIndex        =   28
            Top             =   1095
            Width           =   645
         End
         Begin VB.Image imgProveedor 
            Height          =   240
            Index           =   12
            Left            =   915
            MouseIcon       =   "frmInformesNew3.frx":02B0
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar proveedor"
            Top             =   1125
            Width           =   240
         End
         Begin VB.Image imgProveedor 
            Height          =   240
            Index           =   11
            Left            =   915
            MouseIcon       =   "frmInformesNew3.frx":0402
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar Proveedor"
            Top             =   735
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frmInformesNew3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Integer

'12.  Familias de Artículos
'15.  Clientes Varios

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
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
Private WithEvents frmMtoArticulos As frmBasico2 'Basico2
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmMtoClientes As frmBasico2
Attribute frmMtoClientes.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1


Dim miSQL As String


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
Dim SQL As String
Dim SQL2 As String
Dim Rc As String
Dim RC2 As String

Dim cDesde As String
Dim cHasta As String

    MontaSQL = False
    
    If Not DatosOk Then Exit Function
    
    
    Select Case OpcionListado
        Case 12 ' familias
            If Not PonerDesdeHasta2("{sfamia.codfamia}", "N", Me.txtFamia(6), Me.txtDescFamia(6), Me.txtFamia(7), Me.txtDescFamia(7), "pDHProve=""Proveedor: ") Then Exit Function
            If Not PonerDesdeHasta2("{sfamia.codprove}", "N", Me.txtCodProve(11), Me.txtDescProve(11), Me.txtCodProve(12), Me.txtDescProve(12), "pDHFamia=""Familia: ") Then Exit Function
            If chkVarios(0).Value = 0 Then
                    
                    Codigo = IIf(cadFormula <> "", " AND ", "")
                    Codigo = Codigo & "({sfamia.Inactiva} = 0)"
                    cadFormula = cadFormula & Codigo
                    cadSelect = cadSelect & Codigo
                    cadParam = cadParam & "pdh=""-Solo activas""|"
                    numParam = numParam + 1
            End If
            
        Case 15 ' clientes varios
            If txtNumEntero(1).Text <> "" Or txtNumEntero(2).Text <> "" Then
                miSQL = "pdH1=""Cod. postal: "
                If txtNumEntero(1).Text <> "" Then
                    miSQL = miSQL & "  desde " & txtNumEntero(1).Text
                    Codigo = "{sclvar.codpobla} >= """ & txtNumEntero(1).Text & """"
                    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Function
                End If
                If txtNumEntero(2).Text <> "" Then
                    miSQL = miSQL & "  hasta " & txtNumEntero(2).Text
                    Codigo = "{sclvar.codpobla} <= """ & txtNumEntero(2).Text & """"
                    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Function
                End If
                cadParam = cadParam & miSQL & """|"
                numParam = numParam + 1
                
            End If
            
    End Select
    
    AnyadirAFormula cadFormula, cDesde
    AnyadirAFormula cadSelect, Replace(Replace(cDesde, "{", ""), "}", "")
        
    Select Case OpcionListado
        Case 12 'familias/descuentos
            If chkVarios(1).Value = 1 Then
                'Hacemos el select este y cargamos tmpinformes
                If Not CargarDatosFamiliasDtoEnTmp Then Exit Function
                
                
                cadFormula = "{tmpcommand.codusu} = " & vUsu.Codigo
                cadSelect = "tmpcommand.codusu = " & vUsu.Codigo
                
                
                cadParam = cadParam & "Particulares=" & chkVarios(2).Value & "|"
                numParam = numParam + 1
                
                tabla = "tmpcommand"
            Else
                tabla = "sfamia"
            End If
        
    End Select
    
    
    MontaSQL = True
    
End Function

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
        miSQL = miSQL & ", (" & vUsu.Codigo & "," & NumRegElim & "," & DBLet(miRsAux!Codprove, "N") & "," & miRsAux!Codfamia & ","
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
Dim SQL As String
Dim b As Boolean
    
    b = True
    
    Select Case OpcionListado
        Case 12 ' listado de familias
            If chkVarios(1).Value = 0 And chkVarios(2).Value = 1 Then
                MsgBox "El descuento particulares solo para listado con descuentos", vbExclamation
                chkVarios(2).Value = 0
            End If
    End Select
    
    DatosOk = b

End Function


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

End Sub

Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    Select Case OpcionListado
        Case 12 'familias/descuentos
            If chkVarios(1).Value = 0 Then
                'Listado familias tradicional
                cadTitulo = "Listado Familia de Artículos"
                cadNombreRPT = "rAlmFamArtic.rpt"
                cadPDFrpt = cadNombreRPT
            Else
                cadFormula = "{tmpcommand.codusu} = " & vUsu.Codigo
                cadTitulo = "Listado Familia / Descuentos"
                cadNombreRPT = "rFamiaDtos.rpt"
                cadPDFrpt = cadNombreRPT
            End If
        Case 15 ' clientes varios
            If optVarios(3).Value = 0 Then
                'Listado NIF
                cadTitulo = "List. clientes varios NIF"
                cadNombreRPT = "rFacClivarNIF.rpt"
                cadPDFrpt = cadNombreRPT
            Else
                cadTitulo = "List. clientes varios poblacion"
                cadNombreRPT = "rFacClivar.rpt"
                cadPDFrpt = cadNombreRPT
                
                If Me.optVarios(4).Value Then
                    Codigo = "nomclien"
                Else
                    Codigo = "nifclien"
                End If
                Codigo = "{sclvar." & Codigo & "}"
                cadParam = cadParam & "orden2=" & Codigo & "|"
                numParam = numParam + 1
            End If
    End Select
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook OpcionListado
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Sub AccionesCSV()
Dim SQL As String

    'Monto el SQL
    Select Case OpcionListado
        Case 1 'marcas
        
    End Select
    
    If cadSelect <> "" Then SQL = SQL & " WHERE " & cadSelect
    
    If Me.Optcodigo.Value Then
        SQL = SQL & " ORDER BY 1 "
    Else
        SQL = SQL & " ORDER BY 2 "
    End If
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 12
                PonerFoco txtCodProve(11)
            Case 15
                PonerFoco txtNumEntero(1)
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
    For H = 11 To 12
        Me.imgProveedor(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgProveedor(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next H
     
    For H = 6 To 7
        Me.imgFamilia(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        Me.imgFamilia(H).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next H
     
    FrameCobros.visible = True
    
    '###Descomentar
'    CommitConexion
    
    FrameCobrosVisible True, H, W
    
    FrameFamiliasVisible False
    FrameClientesVVisible False
    
    Select Case OpcionListado
        Case 12 ' listado de familias
            FrameFamiliasVisible True
            indFrame = 5
            tabla = "sfamia"
            Me.Caption = "Informe Familias-Descuentos"
            
            'De momento el csv no lo vemos
            Me.optTipoSal(1).Enabled = False
        
        Case 15 ' listado de familias
            FrameClientesVVisible True
            indFrame = 5
            tabla = "sclvar"
            Me.Caption = "Informe Clientes Varios"
            
            'De momento el csv no lo vemos
            Me.optTipoSal(1).Enabled = False
            
    End Select
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = Me.Width - 220
    Me.Height = Me.FrameCobros.Height + 20
End Sub

Private Sub AmpliarFrame(incremen As Long)
    Me.FrameTipoSalida.Top = Me.FrameTipoSalida.Top + incremen
    Me.cmdAccion(0).Top = Me.cmdAccion(0).Top + incremen
    Me.cmdAccion(1).Top = Me.cmdAccion(1).Top + incremen
    Me.cmdCancel.Top = Me.cmdCancel.Top + incremen
    Me.FrameCobros.Height = Me.FrameCobros.Height + incremen
    
    Me.Height = Me.FrameCobros.Height
End Sub


Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
    txtFamia(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'familias
    If txtFamia(indCodigo).Text <> "" Then txtFamia(indCodigo).Text = Format(txtFamia(indCodigo).Text, "000")
    txtDescFamia(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
    txtCodProve(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    If txtCodProve(indCodigo).Text <> "" Then txtCodProve(indCodigo).Text = Format(txtCodProve(indCodigo).Text, "000000")
    txtDescProve(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgProveedor_Click(Index As Integer)
    AbrirFrmProveedores (Index)
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

Private Sub optVarios_Click(Index As Integer)
    'Orden 2 visible si orden1 es Poblacion
    If Index = 3 Or Index = 2 Then
        FrameOrden2.visible = Index = 3
    End If
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

Private Sub KEYBusquedaFam(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFamilia_Click (Indice)
End Sub

Private Sub KEYBusquedaProv(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgProveedor_Click (Indice)
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


Private Sub FrameFamiliasVisible(visible As Boolean)
    FrameFamiliasSel.visible = visible
    FrameFamiliasOpc.visible = visible
    
    FrameFamiliasSel.Enabled = visible
    FrameFamiliasOpc.Enabled = visible
End Sub

Private Sub FrameClientesVVisible(visible As Boolean)
    FrameClientesVSel.visible = visible
    FrameClientesVOpc.visible = visible
    
    FrameClientesVSel.Enabled = visible
    FrameClientesVOpc.Enabled = visible
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

Private Sub AbrirFrmTArticulos(Indice As Integer)
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
    AyudaFamilias frmFam, txtFamia(Indice)
    Set frmFam = Nothing
End Sub

Private Sub AbrirFrmProveedores(Indice As Integer)
    indCodigo = Indice
    Set frmProv = New frmBasico2
    AyudaProveedores frmProv, txtCodProve(Indice)
    Set frmProv = Nothing
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
    Case 1 To 100
        'Marcas
        Aux = "|||||||||||Familias-Descuentos|"
        Aux = Aux & "||ClientesVarios||||||||||"
            
        Aux = RecuperaValor(Aux, outTipoDocumento) & ".pdf"
             
        
    End Select
    NombrePDF = App.Path & "\temp\" & Aux
    If Dir(NombrePDF, vbArchive) <> "" Then Kill NombrePDF
    FileCopy App.Path & "\docum.pdf", NombrePDF
    
    Aux = FijaDireccionEmail(outTipoDocumento)
    If Aux = "" And emailDestinatario <> "" Then Aux = emailDestinatario
    Lanza = Aux & "|"
    Aux = ""
    Select Case outTipoDocumento
        
    Case 1 To 100
        Aux = "|||||||||||FamiliasDescuentos|"
        Aux = Aux & "||ClientesVarios||||||||||"
        '--------------------------------------------------
        Aux = RecuperaValor(Aux, outTipoDocumento)
        
        
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
Dim cad As String
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
        cad = CadenaDesdeHastaBD(Desde.Text, Hasta.Text, campo, Subtipo)
        cad = Replace(cad, "{", "")
        cad = Replace(cad, "}", "")
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Function
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


Private Function AnyadirParametroDH2(cad As String, ByRef TextoDESDE As TextBox, TextoHasta As TextBox, ByRef TD As TextBox, ByRef TH As TextBox) As String
On Error Resume Next
    
    
    If Not TextoDESDE Is Nothing Then
         If TextoDESDE.Text <> "" Then
            cad = cad & "desde " & TextoDESDE.Text
'            If TD.Caption <> "" Then Cad = Cad & " - " & TD.Caption
        End If
    End If
    If Not TextoHasta Is Nothing Then
        If TextoHasta.Text <> "" Then
            cad = cad & "  hasta " & TextoHasta.Text
'            If TH.Caption <> "" Then Cad = Cad & " - " & TH.Caption
        End If
    End If
    
    AnyadirParametroDH2 = cad
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


Private Sub AbrirFrmClientes()
'Clientes
    Set frmMtoClientes = New frmBasico2
    AyudaClientes frmMtoClientes
    Set frmMtoClientes = Nothing
End Sub




Private Sub txtFamia_GotFocus(Index As Integer)
    ConseguirFoco txtFamia(Index), 3
End Sub

Private Sub txtFamia_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 6: KEYBusquedaFam KeyAscii, 6 'codigo desde
            Case 7: KEYBusquedaFam KeyAscii, 7 'codigo hasta
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtFamia_LostFocus(Index As Integer)
    txtFamia(Index).Text = Trim(txtFamia(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtFamia(Index).Text <> "" Then
        If IsNumeric(txtFamia(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtFamia(Index).Text, "N")
            If Codigo = "" Then MsgBox "El codigo no pertenece a ninguna familia", vbExclamation
        Else
            miSQL = "Campo numerico"
        End If
    End If
     
    Me.txtDescFamia(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtFamia(Index).Text = ""
        PonerFoco txtFamia(Index)
        'If Index = 16 Then
            
    End If
    
    
    If Index = 16 And Codigo = "" Then
        If txtFamia(Index).Text <> "" Then
            txtFamia(Index).Text = ""
            PonerFoco txtFamia(Index)
        End If
    End If
End Sub

Private Sub imgFamilia_Click(Index As Integer)
    AbrirFrmFamilias (Index)
End Sub

Private Sub txtCodProve_GotFocus(Index As Integer)
    ConseguirFoco txtCodProve(Index), 3
End Sub

Private Sub txtCodProve_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 11: KEYBusquedaProv KeyAscii, 11 'codigo desde
            Case 12: KEYBusquedaProv KeyAscii, 12 'codigo hasta
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtCodProve_LostFocus(Index As Integer)
    txtCodProve(Index).Text = Trim(txtCodProve(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtCodProve(Index).Text <> "" Then
        If IsNumeric(txtCodProve(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtCodProve(Index).Text, "N")
            If Codigo = "" Then
                If Index < 0 Then   'dE MOMENTO NO ES REQUERIDO PARA NINGUNO
                    'Codprove REQUERIDO
                    miSQL = "No existe proveedor"
                Else
                    MsgBox "El codigo no pertence a ningun proveedor", vbExclamation
                End If
            End If
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescProve(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtCodProve(Index).Text = ""
        PonerFoco txtCodProve(Index)
    End If
    
'    If Index = 8 Then CargaPrecioCompraProve
'    If Index = 13 Then CargarFamiliasProveedor
'    If Index = 29 Then HacerSimulacionPedidoProveedor
End Sub

Private Sub txtNumEntero_GotFocus(Index As Integer)
    ConseguirFoco txtNumEntero(Index), 3
End Sub

Private Sub txtNumEntero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumEntero_LostFocus(Index As Integer)
    txtNumEntero(Index).Text = Trim(txtNumEntero(Index).Text)
    If txtNumEntero(Index).Text <> "" Then
        If PonerFormatoEntero(txtNumEntero(Index)) Then

        Else
            txtNumEntero(Index).Text = ""
        End If
    End If
End Sub
