VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   21225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   21225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameActualizaSdtoFm 
      Height          =   3495
      Left            =   0
      TabIndex        =   27
      Top             =   1800
      Width           =   6375
      Begin VB.TextBox txtDescFamia 
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
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1920
         Width           =   3540
      End
      Begin VB.TextBox txtFamia 
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
         Left            =   1440
         TabIndex        =   30
         Top             =   1920
         Width           =   1140
      End
      Begin VB.CommandButton cmdSdtofmInsert 
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
         Left            =   3885
         TabIndex        =   32
         Top             =   2925
         Width           =   1065
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Insertar sólo los nuevos"
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
         TabIndex        =   31
         Top             =   2640
         Width           =   3090
      End
      Begin VB.TextBox txtActiv 
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
         Left            =   1440
         TabIndex        =   29
         Top             =   1440
         Width           =   780
      End
      Begin VB.TextBox txtDescActiv 
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
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1440
         Width           =   3915
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   1440
         TabIndex        =   28
         Top             =   840
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancelar 
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
         Left            =   5085
         TabIndex        =   33
         Top             =   2925
         Width           =   1065
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   0
         Left            =   1170
         Top             =   1920
         Width           =   240
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
         Index           =   18
         Left            =   150
         TabIndex        =   44
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label lblIndicador 
         Caption         =   "Label2"
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
         Left            =   240
         TabIndex        =   38
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label lblDpto 
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
         Index           =   2
         Left            =   150
         TabIndex        =   37
         Top             =   840
         Width           =   630
      End
      Begin VB.Image imgActividad 
         Height          =   240
         Index           =   1
         Left            =   1170
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   36
         Top             =   1440
         Width           =   990
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1170
         Picture         =   "frmListado5.frx":0000
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Actualizar descuentos familia/marca"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   195
         TabIndex        =   34
         Top             =   240
         Width           =   6000
      End
   End
   Begin VB.Frame FrameDtoAsginar 
      Height          =   3135
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdCargaDtoFamiliaActiv 
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
         Left            =   3570
         TabIndex        =   21
         Top             =   2520
         Width           =   1065
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Insertar sólo los nuevos"
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
         TabIndex        =   20
         Top             =   2280
         Width           =   3090
      End
      Begin VB.ComboBox cboTipoDescuento 
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
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1560
         Width           =   3795
      End
      Begin VB.TextBox txtDescActiv 
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
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   960
         Width           =   3420
      End
      Begin VB.TextBox txtActiv 
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
         Left            =   1680
         TabIndex        =   18
         Top             =   960
         Width           =   780
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   3
         Left            =   4785
         TabIndex        =   22
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo descuento"
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
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   1620
         Width           =   1635
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   39
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   990
      End
      Begin VB.Image imgActividad 
         Height          =   240
         Index           =   0
         Left            =   1320
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Actualizar desde familias"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Left            =   480
         TabIndex        =   23
         Top             =   240
         Width           =   5145
      End
   End
   Begin VB.Frame FrameACtualizaPrecioMinimo 
      Height          =   4335
      Left            =   360
      TabIndex        =   439
      Top             =   2280
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chkVarios 
         Caption         =   "Fecha menor o igual"
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
         Index           =   9
         Left            =   360
         TabIndex        =   451
         Top             =   3600
         Width           =   2775
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Eliminar promoción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   1680
         TabIndex        =   450
         Top             =   3000
         Width           =   3135
      End
      Begin VB.CommandButton cmdPrecioMinimo 
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
         Left            =   3240
         TabIndex        =   444
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Borrar precio minimo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   1680
         TabIndex        =   443
         Top             =   2520
         Width           =   3135
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Establecer precio mínimo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   1680
         TabIndex        =   442
         Top             =   2040
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   40
         Left            =   4680
         TabIndex        =   445
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   24
         Left            =   4200
         TabIndex        =   441
         Top             =   1283
         Width           =   1350
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   23
         Left            =   1560
         TabIndex        =   440
         Top             =   1283
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   38
         Left            =   3480
         TabIndex        =   449
         Top             =   1335
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   37
         Left            =   480
         TabIndex        =   448
         Top             =   1335
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   3870
         Picture         =   "frmListado5.frx":008B
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1200
         Picture         =   "frmListado5.frx":0116
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fechas promoción"
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
         Index           =   52
         Left            =   480
         TabIndex        =   447
         Top             =   960
         Width           =   1950
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Actualizar precio mínimo artículo"
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
         Index           =   39
         Left            =   840
         TabIndex        =   446
         Top             =   240
         Width           =   4785
      End
   End
   Begin VB.Frame FrameAsignarAlbaranesEuler 
      Height          =   9375
      Left            =   0
      TabIndex        =   476
      Top             =   -120
      Visible         =   0   'False
      Width           =   16935
      Begin VB.CommandButton cmdEstablecerAlbaranPrincipal 
         Height          =   375
         Index           =   1
         Left            =   3960
         Picture         =   "frmListado5.frx":01A1
         Style           =   1  'Graphical
         TabIndex        =   502
         ToolTipText     =   "Establecer albarán principal del proyecto"
         Top             =   3840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdEstablecerAlbaranPrincipal 
         Height          =   375
         Index           =   0
         Left            =   3960
         Picture         =   "frmListado5.frx":0BA3
         Style           =   1  'Graphical
         TabIndex        =   501
         ToolTipText     =   "Establecer albarán principal del proyecto"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.TreeView treeAlb 
         Height          =   2655
         Index           =   0
         Left            =   240
         TabIndex        =   481
         Top             =   720
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   4683
         _Version        =   393217
         Indentation     =   2646
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   15240
         TabIndex        =   478
         Top             =   8880
         Width           =   1095
      End
      Begin VB.CommandButton cmdAsignarAlbarEuler 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13920
         TabIndex        =   477
         Top             =   8880
         Width           =   1095
      End
      Begin MSComctlLib.TreeView treeAlb 
         Height          =   4335
         Index           =   1
         Left            =   240
         TabIndex        =   482
         Top             =   4320
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   7646
         _Version        =   393217
         Indentation     =   2646
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblIndicador 
         Caption         =   "lblIndicador 5"
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
         Index           =   8
         Left            =   240
         TabIndex        =   483
         Top             =   8880
         Width           =   4455
      End
      Begin VB.Label Label9 
         Caption         =   "Albaranes sin vincular"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   42
         Left            =   240
         TabIndex        =   480
         Top             =   3840
         Width           =   3435
      End
      Begin VB.Label Label9 
         Caption         =   "Albaranes en el proyecto"
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
         Index           =   41
         Left            =   240
         TabIndex        =   479
         Top             =   240
         Width           =   4365
      End
   End
   Begin VB.Frame FrameLineasAlbaFalsoEuler 
      Height          =   2055
      Left            =   105
      TabIndex        =   285
      Top             =   2280
      Visible         =   0   'False
      Width           =   13335
      Begin VB.TextBox txtNumero 
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
         Left            =   10920
         TabIndex        =   290
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   27
         Left            =   11940
         TabIndex        =   293
         Top             =   1440
         Width           =   1065
      End
      Begin VB.CommandButton cmdAceptarLinEspEuler 
         Caption         =   "Aceptar"
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
         Left            =   10740
         TabIndex        =   292
         Top             =   1440
         Width           =   1065
      End
      Begin VB.TextBox txtModificable 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Index           =   5
         Left            =   2640
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   287
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtNumero 
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
         Left            =   11640
         TabIndex        =   291
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtNumero 
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
         Left            =   9600
         TabIndex        =   289
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtNumero 
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
         Index           =   5
         Left            =   8400
         TabIndex        =   288
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtModificable 
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
         Left            =   240
         TabIndex        =   286
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Dto"
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
         Height          =   240
         Index           =   38
         Left            =   10920
         TabIndex        =   300
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "MODIFICAR"
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
         Height          =   240
         Index           =   37
         Left            =   240
         TabIndex        =   299
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         Height          =   240
         Index           =   36
         Left            =   11640
         TabIndex        =   298
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
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
         Height          =   240
         Index           =   35
         Left            =   9600
         TabIndex        =   297
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
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
         Height          =   240
         Index           =   33
         Left            =   8400
         TabIndex        =   296
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   32
         Left            =   2640
         TabIndex        =   295
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
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
         Height          =   240
         Index           =   31
         Left            =   240
         TabIndex        =   294
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame FrameCosteLin 
      Height          =   3735
      Left            =   7800
      TabIndex        =   410
      Top             =   -360
      Visible         =   0   'False
      Width           =   11535
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   421
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   9480
         TabIndex        =   411
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   38
         Left            =   9960
         TabIndex        =   412
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Index           =   8
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   416
         Top             =   2040
         Width           =   6615
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   415
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   414
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Precio venta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   34
         Left            =   7920
         TabIndex        =   422
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Coste"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   33
         Left            =   9840
         TabIndex        =   420
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ampliación línea"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   32
         Left            =   240
         TabIndex        =   419
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   31
         Left            =   2520
         TabIndex        =   418
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código artículo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   30
         Left            =   240
         TabIndex        =   417
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Precio coste articulo varios"
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
         Index           =   37
         Left            =   240
         TabIndex        =   413
         Top             =   360
         Width           =   3915
      End
   End
   Begin VB.Frame FrameTaxcoSvenciAlvic 
      Height          =   3375
      Left            =   3600
      TabIndex        =   491
      Top             =   3240
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   26
         Left            =   4680
         TabIndex        =   496
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   25
         Left            =   1680
         TabIndex        =   494
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   45
         Left            =   5160
         TabIndex        =   493
         Top             =   2640
         Width           =   1165
      End
      Begin VB.CommandButton cmdAjusteVtosFaccliALVIC 
         Caption         =   "Aceptar"
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
         Left            =   3720
         TabIndex        =   492
         Top             =   2640
         Width           =   1165
      End
      Begin VB.Label lblIndicador 
         Caption         =   "lblIndicador 9"
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
         Index           =   9
         Left            =   240
         TabIndex        =   500
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   55
         Left            =   240
         TabIndex        =   499
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Ajuste formas de pago en facturas ALVIC"
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
         Index           =   44
         Left            =   240
         TabIndex        =   498
         Top             =   360
         Width           =   5895
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   26
         Left            =   4320
         Picture         =   "frmListado5.frx":15A5
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   54
         Left            =   3720
         TabIndex        =   497
         Top             =   1560
         Width           =   330
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   1320
         Picture         =   "frmListado5.frx":1630
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   53
         Left            =   240
         TabIndex        =   495
         Top             =   1560
         Width           =   600
      End
   End
   Begin VB.Frame FrameCestaApp 
      Height          =   8895
      Left            =   2880
      TabIndex        =   486
      Top             =   0
      Visible         =   0   'False
      Width           =   12615
      Begin VB.CommandButton cmdCesta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         TabIndex        =   489
         Top             =   8280
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   44
         Left            =   11040
         TabIndex        =   488
         Top             =   8280
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   6975
         Index           =   15
         Left            =   240
         TabIndex        =   487
         Top             =   1080
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12303
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1781
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Articulo"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   9067
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   2485
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Stock"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   11
         Left            =   600
         Picture         =   "frmListado5.frx":16BB
         Top             =   8160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   10
         Left            =   240
         Picture         =   "frmListado5.frx":1805
         Top             =   8160
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Cestas artículos almacén"
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
         Index           =   43
         Left            =   240
         TabIndex        =   490
         Top             =   480
         Width           =   11685
      End
   End
   Begin VB.Frame FrameTaxcoGasolineraCambiCli 
      Height          =   9135
      Left            =   7800
      TabIndex        =   348
      Top             =   120
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox txtCliente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   1320
         TabIndex        =   347
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   353
         Top             =   960
         Width           =   6375
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   34
         Left            =   11280
         TabIndex        =   350
         Top             =   8640
         Width           =   975
      End
      Begin VB.CommandButton cmdTaxcoCambioCliente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   10200
         TabIndex        =   349
         Top             =   8640
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   6975
         Index           =   13
         Left            =   240
         TabIndex        =   351
         Top             =   1440
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12303
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Albaran"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2485
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "F.P."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Articulo"
            Object.Width           =   4198
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Descripcion"
            Object.Width           =   4805
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   9
         Left            =   11520
         Picture         =   "frmListado5.frx":194F
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   8
         Left            =   11880
         Picture         =   "frmListado5.frx":1A99
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   14
         Left            =   1080
         Picture         =   "frmListado5.frx":1BE3
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   19
         Left            =   240
         TabIndex        =   354
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Cambio cliente albaranes ALVIC"
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
         Height          =   345
         Index           =   33
         Left            =   3480
         TabIndex        =   352
         Top             =   240
         Width           =   5070
      End
   End
   Begin VB.Frame FrameCopiarPrecios 
      Height          =   4305
      Left            =   0
      TabIndex        =   452
      Top             =   0
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtDescFamia 
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
         Index           =   3
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   484
         Top             =   2925
         Width           =   5520
      End
      Begin VB.TextBox txtFamia 
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
         Left            =   270
         TabIndex        =   457
         Tag             =   "Cod. Proveedor|N|N|0|999999|slispr|codprove|000000|S|"
         Top             =   2925
         Width           =   840
      End
      Begin VB.TextBox txtDescProve 
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
         Index           =   3
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   461
         Top             =   2070
         Width           =   5520
      End
      Begin VB.TextBox txtProve 
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
         Left            =   270
         TabIndex        =   456
         Tag             =   "Cod. Proveedor|N|N|0|999999|slispr|codprove|000000|S|"
         Top             =   2070
         Width           =   840
      End
      Begin VB.TextBox txtProve 
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
         Left            =   285
         TabIndex        =   455
         Tag             =   "Cod. Proveedor|N|N|0|999999|slispr|codprove|000000|S|"
         Top             =   1260
         Width           =   840
      End
      Begin VB.TextBox txtDescProve 
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
         Index           =   2
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   454
         Top             =   1260
         Width           =   5520
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   41
         Left            =   5610
         TabIndex        =   459
         Top             =   3420
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepCopia 
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
         Left            =   4455
         TabIndex        =   458
         Top             =   3420
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Familia Origen"
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
         Left            =   300
         TabIndex        =   485
         Top             =   2610
         Width           =   1590
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   3
         Left            =   1965
         Picture         =   "frmListado5.frx":1CE5
         Tag             =   "-1"
         ToolTipText     =   "Buscar Proveedor"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor Destino"
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
         Left            =   300
         TabIndex        =   462
         Top             =   1755
         Width           =   2085
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   3
         Left            =   2370
         Picture         =   "frmListado5.frx":26E7
         Tag             =   "-1"
         ToolTipText     =   "Buscar Proveedor"
         Top             =   1755
         Width           =   240
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   2
         Left            =   2340
         Picture         =   "frmListado5.frx":30E9
         Tag             =   "-1"
         ToolTipText     =   "Buscar Proveedor"
         Top             =   945
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor Origen"
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
         Left            =   315
         TabIndex        =   460
         Top             =   945
         Width           =   1950
      End
      Begin VB.Label Label9 
         Caption         =   "Copiar Precios Proveedor"
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
         Index           =   40
         Left            =   315
         TabIndex        =   453
         Top             =   360
         Width           =   4365
      End
   End
   Begin VB.Frame FrameCambioCliente 
      Height          =   7815
      Left            =   6480
      TabIndex        =   355
      Top             =   180
      Visible         =   0   'False
      Width           =   9255
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   364
         Top             =   1560
         Width           =   6375
      End
      Begin VB.TextBox txtCliente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   1080
         TabIndex        =   357
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCambioCliente 
         Caption         =   "Aceptar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   358
         Top             =   7200
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   7920
         TabIndex        =   359
         Top             =   7200
         Width           =   975
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   361
         Top             =   960
         Width           =   6375
      End
      Begin VB.TextBox txtCliente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   1080
         TabIndex        =   356
         Top             =   960
         Width           =   1335
      End
      Begin MSComctlLib.ListView lw 
         Height          =   4815
         Index           =   14
         Left            =   120
         TabIndex        =   360
         Top             =   2160
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Referencia"
            Object.Width           =   11228
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Datos"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblIndicador 
         Caption         =   "lblIndicador 5"
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
         Index           =   5
         Left            =   120
         TabIndex        =   366
         Top             =   7320
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   21
         Left            =   120
         TabIndex        =   365
         Top             =   1560
         Width           =   630
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   16
         Left            =   840
         Picture         =   "frmListado5.frx":3AEB
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Cambio cliente ARIGES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Index           =   34
         Left            =   1920
         TabIndex        =   363
         Top             =   240
         Width           =   5070
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   20
         Left            =   120
         TabIndex        =   362
         Top             =   960
         Width           =   570
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   15
         Left            =   840
         Picture         =   "frmListado5.frx":3BED
         Top             =   960
         Width           =   240
      End
   End
   Begin VB.Frame FrameComparativoDtos 
      Height          =   6015
      Left            =   4200
      TabIndex        =   387
      Top             =   240
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox txtDescActiv 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   409
         Top             =   4560
         Width           =   4455
      End
      Begin VB.TextBox txtActiv 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   382
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtDescActiv 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   405
         Top             =   4148
         Width           =   4455
      End
      Begin VB.TextBox txtActiv 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   381
         Top             =   4148
         Width           =   1095
      End
      Begin VB.TextBox txtCodigoVario 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   8880
         TabIndex        =   386
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDesVario 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   402
         Top             =   4080
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox txtCodigoVario 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   7680
         TabIndex        =   385
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDesVario 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   399
         Top             =   3600
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox txtFamia 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   380
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   398
         Top             =   3240
         Width           =   4455
      End
      Begin VB.TextBox txtFamia 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   379
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   394
         Top             =   2880
         Width           =   4455
      End
      Begin VB.TextBox txtProve 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   378
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   392
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox txtDescProve 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   389
         Top             =   1440
         Width           =   5415
      End
      Begin VB.TextBox txtProve 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   377
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   37
         Left            =   6480
         TabIndex        =   384
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton cmdListaComparaDto 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   383
         Top             =   5280
         Width           =   975
      End
      Begin VB.Image imgActividad 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "frmListado5.frx":3CEF
         Top             =   4590
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   29
         Left            =   360
         TabIndex        =   408
         Top             =   4200
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   28
         Left            =   360
         TabIndex        =   407
         Top             =   4560
         Width           =   555
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   48
         Left            =   240
         TabIndex        =   406
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Image imgActividad 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmListado5.frx":3DF1
         Top             =   4185
         Width           =   240
      End
      Begin VB.Label lblIndicador 
         Caption         =   "Label2"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   404
         Top             =   5400
         Width           =   2895
      End
      Begin VB.Image imgCodigoVario 
         Height          =   240
         Index           =   1
         Left            =   8640
         Picture         =   "frmListado5.frx":3EF3
         Top             =   3960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   27
         Left            =   7920
         TabIndex        =   403
         Top             =   3960
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   47
         Left            =   6960
         TabIndex        =   401
         Top             =   3600
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Image imgCodigoVario 
         Height          =   240
         Index           =   0
         Left            =   7680
         Picture         =   "frmListado5.frx":3FF5
         Top             =   3120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   26
         Left            =   7560
         TabIndex        =   400
         Top             =   3360
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmListado5.frx":40F7
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   25
         Left            =   360
         TabIndex        =   397
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   24
         Left            =   360
         TabIndex        =   396
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   46
         Left            =   240
         TabIndex        =   395
         Top             =   2520
         Width           =   795
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmListado5.frx":41F9
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmListado5.frx":42FB
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   23
         Left            =   240
         TabIndex        =   393
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   45
         Left            =   240
         TabIndex        =   391
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   22
         Left            =   240
         TabIndex        =   390
         Top             =   1440
         Width           =   600
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmListado5.frx":43FD
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Listado comparativo descuentos "
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
         Index           =   36
         Left            =   2040
         TabIndex        =   388
         Top             =   360
         Width           =   4755
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameSelProveedores 
      Height          =   6615
      Left            =   6480
      TabIndex        =   91
      Top             =   360
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdSelProvee 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   95
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5160
         TabIndex        =   94
         Top             =   6120
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5535
         Index           =   1
         Left            =   240
         TabIndex        =   93
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   9763
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   600
         Picture         =   "frmListado5.frx":44FF
         Top             =   6240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "frmListado5.frx":4649
         Top             =   6240
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Seleccionar proveedores"
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
         Index           =   10
         Left            =   1440
         TabIndex        =   92
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Frame FrameTaxcoNuevaTaller 
      Height          =   6015
      Left            =   1440
      TabIndex        =   301
      Top             =   840
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdTaxcoNuevoEntradaVehiculo 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7560
         TabIndex        =   307
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtModificable 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   360
         MaxLength       =   15
         TabIndex        =   304
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   29
         Left            =   8880
         TabIndex        =   303
         Top             =   5400
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   3135
         Index           =   8
         Left            =   360
         TabIndex        =   306
         Top             =   2040
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5530
         SortKey         =   5
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "O:R."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Kms"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Observa"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Orden"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Histórico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   41
         Left            =   360
         TabIndex        =   308
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Matrícula"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   40
         Left            =   360
         TabIndex        =   305
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Entrada vehículo"
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
         Index           =   27
         Left            =   2040
         TabIndex        =   302
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FramePedCli_A_prov 
      Height          =   5655
      Left            =   720
      TabIndex        =   336
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
      Begin VB.CheckBox chkVarios 
         Caption         =   "Insertar precios y descuentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   341
         Top             =   5160
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   32
         Left            =   13800
         TabIndex        =   337
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdPedCli_A_prov 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   12720
         TabIndex        =   339
         Top             =   5040
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   3975
         Index           =   11
         Left            =   240
         TabIndex        =   338
         Top             =   960
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Arti."
            Object.Width           =   2962
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7761
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Dto1"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Dto2"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Importel"
            Object.Width           =   2141
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Prov."
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Nombre"
            Object.Width           =   3835
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Generar pedido proveedor desde pedido venta"
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
         Index           =   31
         Left            =   3480
         TabIndex        =   340
         Top             =   360
         Width           =   8175
      End
   End
   Begin VB.Frame FrameCoarval 
      Height          =   8295
      Left            =   1800
      TabIndex        =   274
      Top             =   120
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   26
         Left            =   11280
         TabIndex        =   278
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton cmdImpFraCoarval 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   9840
         TabIndex        =   277
         Top             =   7680
         Width           =   1215
      End
      Begin MSComctlLib.ListView lw 
         Height          =   6495
         Index           =   7
         Left            =   240
         TabIndex        =   276
         Top             =   960
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   11456
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblIndicador 
         Caption         =   "I"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   279
         Top             =   7920
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "Importar facturas coarval"
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
         Index           =   26
         Left            =   3960
         TabIndex        =   275
         Top             =   360
         Width           =   4830
      End
   End
   Begin VB.Frame FrameBusqPreviaPedFontenas 
      Height          =   8775
      Left            =   600
      TabIndex        =   342
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
      Begin MSComctlLib.ImageList imglistPed 
         Left            =   240
         Top             =   8040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListado5.frx":4793
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListado5.frx":AFF5
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmListado5.frx":11857
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   33
         Left            =   9840
         TabIndex        =   344
         Top             =   8040
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarPedFontenas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   8760
         TabIndex        =   343
         Top             =   8040
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   6615
         Index           =   12
         Left            =   240
         TabIndex        =   346
         Top             =   840
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   11668
         SortKey         =   6
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Estado"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Prioridad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nombre"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ordenfec"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ordenestado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ordenprioridad"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Pedidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   945
         Index           =   32
         Left            =   240
         TabIndex        =   345
         Top             =   360
         Width           =   2685
      End
   End
   Begin VB.Frame FrameCRMClieAccion 
      Height          =   5895
      Left            =   120
      TabIndex        =   207
      Top             =   0
      Visible         =   0   'False
      Width           =   12135
      Begin VB.Frame FrameAccionComerOrden 
         Caption         =   "Frame1"
         Height          =   975
         Left            =   120
         TabIndex        =   280
         Top             =   4560
         Width           =   5655
         Begin VB.CheckBox chkVarios 
            Caption         =   "Ocultar hora"
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   284
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton optVarios 
            Caption         =   "Accion comercial"
            Height          =   255
            Index           =   8
            Left            =   3600
            TabIndex        =   283
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optVarios 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   282
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblDpto 
            AutoSize        =   -1  'True
            Caption         =   "Ordenado"
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
            Left            =   120
            TabIndex        =   281
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   11
         Left            =   1320
         TabIndex        =   211
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   233
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   230
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   210
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdCRMClieAccion 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   9720
         TabIndex        =   217
         Top             =   5280
         Width           =   975
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "c"
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   215
         Top             =   4320
         Width           =   1455
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "c"
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   214
         Top             =   4320
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   12
         Left            =   4080
         TabIndex        =   213
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   212
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtDescAge 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   223
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   209
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtAgente 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   208
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDescAge 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   220
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   22
         Left            =   10800
         TabIndex        =   218
         Top             =   5280
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   3975
         Index           =   5
         Left            =   6120
         TabIndex        =   216
         Top             =   1080
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7011
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   11
         Left            =   1080
         Picture         =   "frmListado5.frx":180B9
         Top             =   2880
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   600
         TabIndex        =   234
         Top             =   2880
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   600
         TabIndex        =   232
         Top             =   2520
         Width           =   450
      End
      Begin VB.Label lblDpto 
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
         Index           =   27
         Left            =   240
         TabIndex        =   231
         Top             =   2160
         Width           =   585
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   10
         Left            =   1080
         Picture         =   "frmListado5.frx":181BB
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label lblIndicador 
         Caption         =   "I"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   229
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Acciones comerciales"
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
         Left            =   6120
         TabIndex        =   228
         Top             =   840
         Width           =   1815
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   7
         Left            =   11640
         Picture         =   "frmListado5.frx":182BD
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   6
         Left            =   11160
         Picture         =   "frmListado5.frx":18407
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   3720
         Picture         =   "frmListado5.frx":18551
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   15
         Left            =   3120
         TabIndex        =   227
         Top             =   3720
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   1080
         Picture         =   "frmListado5.frx":185DC
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   25
         Left            =   240
         TabIndex        =   226
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   600
         TabIndex        =   225
         Top             =   3720
         Width           =   465
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   600
         TabIndex        =   224
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmListado5.frx":18667
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgAgente 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmListado5.frx":18769
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   24
         Left            =   240
         TabIndex        =   222
         Top             =   840
         Width           =   615
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   221
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label9 
         Caption         =   "Clientes por acciones comerciales"
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
         Index           =   22
         Left            =   3480
         TabIndex        =   219
         Top             =   240
         Width           =   4830
      End
   End
   Begin VB.Frame FramePreviFacturaTaxo 
      Height          =   4695
      Left            =   4080
      TabIndex        =   317
      Top             =   240
      Width           =   6015
      Begin VB.CheckBox chkVerificarCtas 
         Caption         =   "Verificar cuentas credito"
         Height          =   255
         Left            =   360
         TabIndex        =   335
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton cmdPrevFraALVIC 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   324
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   31
         Left            =   4560
         TabIndex        =   325
         Top             =   4080
         Width           =   975
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Tienda"
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   323
         Top             =   3240
         Width           =   1455
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Gasolinera"
         Height          =   255
         Index           =   9
         Left            =   1440
         TabIndex        =   322
         Top             =   3240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   18
         Left            =   3960
         TabIndex        =   321
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   17
         Left            =   1320
         TabIndex        =   320
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   13
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   329
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   13
         Left            =   1200
         TabIndex        =   319
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   12
         Left            =   1200
         TabIndex        =   318
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   326
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "Prevision facturacion ALVIC"
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
         Index           =   30
         Left            =   1080
         TabIndex        =   334
         Top             =   360
         Width           =   4830
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   3000
         TabIndex        =   333
         Top             =   2400
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   3600
         Picture         =   "frmListado5.frx":1886B
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   480
         TabIndex        =   332
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   44
         Left            =   120
         TabIndex        =   331
         Top             =   2040
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   960
         Picture         =   "frmListado5.frx":188F6
         Top             =   2400
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   480
         TabIndex        =   330
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   13
         Left            =   960
         Picture         =   "frmListado5.frx":18981
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   12
         Left            =   960
         Picture         =   "frmListado5.frx":18A83
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   43
         Left            =   120
         TabIndex        =   328
         Top             =   840
         Width           =   585
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   480
         TabIndex        =   327
         Top             =   1200
         Width           =   450
      End
   End
   Begin VB.Frame FrameOrdenarLineas 
      Height          =   7095
      Left            =   4680
      TabIndex        =   236
      Top             =   360
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton cmdLinPed 
         Height          =   495
         Index           =   3
         Left            =   3240
         Picture         =   "frmListado5.frx":18B85
         Style           =   1  'Graphical
         TabIndex        =   244
         ToolTipText     =   "Ultimo"
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton cmdLinPed 
         Height          =   495
         Index           =   2
         Left            =   2400
         Picture         =   "frmListado5.frx":1A5F7
         Style           =   1  'Graphical
         TabIndex        =   243
         ToolTipText     =   "Sigiente"
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton cmdLinPed 
         Height          =   495
         Index           =   1
         Left            =   1200
         Picture         =   "frmListado5.frx":1C069
         Style           =   1  'Graphical
         TabIndex        =   242
         ToolTipText     =   "Anterior"
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton cmdLinPed 
         Height          =   495
         Index           =   0
         Left            =   360
         Picture         =   "frmListado5.frx":1DADB
         Style           =   1  'Graphical
         TabIndex        =   241
         ToolTipText     =   "Primero"
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton cmdOrdenarLineas 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   9480
         TabIndex        =   240
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   23
         Left            =   10560
         TabIndex        =   239
         Top             =   6480
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5295
         Index           =   6
         Left            =   360
         TabIndex        =   238
         Top             =   720
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9340
         SortKey         =   7
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lin"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Referencia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Titulo"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Pendiente"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Solicitadas"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Precio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Importe"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "oorden"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Ordenar lineas "
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
         Index           =   23
         Left            =   360
         TabIndex        =   237
         Top             =   240
         Width           =   4845
      End
   End
   Begin VB.Frame FrameListPedxDia 
      Height          =   2535
      Left            =   7320
      TabIndex        =   367
      Top             =   120
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdListPedDia 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   370
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   36
         Left            =   4440
         TabIndex        =   371
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   20
         Left            =   4200
         TabIndex        =   369
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   19
         Left            =   1440
         TabIndex        =   368
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblIndicador 
         Caption         =   "I"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   376
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   20
         Left            =   3840
         Picture         =   "frmListado5.frx":1F54D
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   3240
         TabIndex        =   375
         Top             =   1320
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   1080
         Picture         =   "frmListado5.frx":1F5D8
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   11
         Left            =   360
         TabIndex        =   374
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   373
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Listado pedidos por dia"
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
         Index           =   35
         Left            =   840
         TabIndex        =   372
         Top             =   360
         Width           =   4275
      End
   End
   Begin VB.Frame frameAlvic 
      Height          =   9015
      Left            =   8040
      TabIndex        =   309
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton cmdAlvic 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   313
         Top             =   8520
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   30
         Left            =   6600
         TabIndex        =   310
         Top             =   8520
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   3135
         Index           =   9
         Left            =   360
         TabIndex        =   311
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Forma de pago"
            Object.Width           =   7126
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lw 
         Height          =   3255
         Index           =   10
         Left            =   360
         TabIndex        =   316
         Top             =   5040
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Observacion"
            Object.Width           =   7832
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Infrmación adicional"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   345
         Index           =   29
         Left            =   360
         TabIndex        =   315
         Top             =   4560
         Width           =   6495
      End
      Begin VB.Label lblDpto 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   42
         Left            =   360
         TabIndex        =   314
         Top             =   3960
         Width           =   4815
      End
      Begin VB.Label Label9 
         Caption         =   "ALVIC"
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
         Index           =   28
         Left            =   360
         TabIndex        =   312
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame FramePtosCliente 
      Height          =   3495
      Left            =   4920
      TabIndex        =   186
      Top             =   2880
      Width           =   5535
      Begin VB.TextBox txtModificable 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   183
         Top             =   2040
         Width           =   3615
      End
      Begin VB.CommandButton cmdPuntosCliente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   184
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3960
         TabIndex        =   182
         Text            =   "Text1"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   181
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   19
         Left            =   4200
         TabIndex        =   185
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
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
         Index           =   21
         Left            =   240
         TabIndex        =   191
         Top             =   2085
         Width           =   690
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Puntos"
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
         Index           =   20
         Left            =   3120
         TabIndex        =   190
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   840
         Picture         =   "frmListado5.frx":1F663
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   19
         Left            =   240
         TabIndex        =   189
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lblDpto 
         Caption         =   "d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   18
         Left            =   240
         TabIndex        =   188
         Top             =   840
         Width           =   5040
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Insertar puntos cliente"
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
         Index           =   19
         Left            =   840
         TabIndex        =   187
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame FrameCreditoYCaucion 
      Height          =   2535
      Left            =   3240
      TabIndex        =   265
      Top             =   4560
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkVarios 
         Caption         =   "Fichero csv"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   273
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   16
         Left            =   3840
         TabIndex        =   268
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   15
         Left            =   1320
         TabIndex        =   267
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCreditoCaucion 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   269
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   25
         Left            =   4680
         TabIndex        =   270
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   2880
         TabIndex        =   272
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   3480
         Picture         =   "frmListado5.frx":1F6EE
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   360
         TabIndex        =   271
         Top             =   1200
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   960
         Picture         =   "frmListado5.frx":1F779
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Listado ventas crédito y caución"
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
         Index           =   25
         Left            =   240
         TabIndex        =   266
         Top             =   360
         Width           =   5280
      End
   End
   Begin VB.Frame FrameComprasTratamientos 
      Height          =   2535
      Left            =   0
      TabIndex        =   80
      Top             =   4560
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CheckBox chkVarios 
         Caption         =   "Detalla artículos"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   84
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   5
         Left            =   4080
         TabIndex        =   83
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdComprasTratamientos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3120
         TabIndex        =   85
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   4320
         TabIndex        =   86
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   82
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblIndicador 
         Caption         =   "I"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   90
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   9
         Left            =   3120
         TabIndex        =   89
         Top             =   1080
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   3720
         Picture         =   "frmListado5.frx":1F804
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   88
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   7
         Left            =   240
         TabIndex        =   87
         Top             =   720
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmListado5.frx":1F88F
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Ajuste compras tratamientos"
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
         Left            =   720
         TabIndex        =   81
         Top             =   240
         Width           =   4275
      End
   End
   Begin VB.Frame FrameListadoAlb 
      Height          =   4695
      Left            =   5640
      TabIndex        =   245
      Top             =   1440
      Visible         =   0   'False
      Width           =   6135
      Begin VB.ComboBox cboDestinoB 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmListado5.frx":1F91A
         Left            =   840
         List            =   "frmListado5.frx":1F921
         Style           =   2  'Dropdown List
         TabIndex        =   251
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chkAlb 
         Caption         =   "Facturados"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   253
         Top             =   3120
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkAlb 
         Caption         =   "Pendientes"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   252
         Top             =   3120
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   14
         Left            =   4440
         TabIndex        =   250
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   13
         Left            =   1440
         TabIndex        =   249
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   259
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   248
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   258
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   247
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdListadoAlb 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   254
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   24
         Left            =   4680
         TabIndex        =   255
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lblDestinoB 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
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
         Left            =   240
         TabIndex        =   264
         Top             =   3180
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   3480
         TabIndex        =   263
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   4080
         Picture         =   "frmListado5.frx":1F92E
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   16
         Left            =   600
         TabIndex        =   262
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   29
         Left            =   240
         TabIndex        =   261
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1080
         Picture         =   "frmListado5.frx":1F9B9
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   260
         Top             =   1440
         Width           =   420
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   9
         Left            =   840
         Picture         =   "frmListado5.frx":1FA44
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   8
         Left            =   840
         Picture         =   "frmListado5.frx":1FB46
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   28
         Left            =   120
         TabIndex        =   257
         Top             =   840
         Width           =   585
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   256
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Listado albaranes"
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
         Index           =   24
         Left            =   600
         TabIndex        =   246
         Top             =   240
         Width           =   4830
      End
   End
   Begin VB.Frame FrameAlbaranesClientes 
      Height          =   6375
      Left            =   5040
      TabIndex        =   142
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CheckBox chkPedidos 
         Caption         =   "Pedidos"
         Height          =   255
         Left            =   5400
         TabIndex        =   235
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelecAlbaran 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4800
         TabIndex        =   146
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   14
         Left            =   5880
         TabIndex        =   144
         Top             =   5880
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5055
         Index           =   3
         Left            =   240
         TabIndex        =   143
         Top             =   720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8916
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Nº Factura"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Albaranes(facturas) cliente"
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
         Index           =   14
         Left            =   360
         TabIndex        =   145
         Top             =   240
         Width           =   3885
      End
   End
   Begin VB.Frame FrameDeclaraAlcohol 
      Height          =   3855
      Left            =   4560
      TabIndex        =   192
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox chkVarios 
         Caption         =   "Marcar declaracion"
         Height          =   255
         HelpContextID   =   3
         Index           =   3
         Left            =   360
         TabIndex        =   202
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   194
         Text            =   "Text1"
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox cboTrimiestre 
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   193
         Text            =   "Combo1"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Fichero"
         Height          =   255
         HelpContextID   =   4
         Index           =   4
         Left            =   3000
         TabIndex        =   195
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton cmdDeclaraAlcohol 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   196
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   20
         Left            =   4320
         TabIndex        =   197
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   199
         Text            =   "Text1"
         Top             =   1320
         Width           =   4935
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Periodo "
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
         Index           =   23
         Left            =   360
         TabIndex        =   201
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo periodo liquidado"
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
         Index           =   22
         Left            =   360
         TabIndex        =   200
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Declaración alcohol AEAT"
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
         Index           =   20
         Left            =   600
         TabIndex        =   198
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FrameTelefonia 
      Height          =   3015
      Left            =   4920
      TabIndex        =   173
      Top             =   1440
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   175
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   174
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdContratoTelef 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1680
         TabIndex        =   176
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   2880
         TabIndex        =   177
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Meses duración"
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
         Index           =   17
         Left            =   360
         TabIndex        =   180
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Importe terminal"
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
         Index           =   16
         Left            =   360
         TabIndex        =   179
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Datos contrato"
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
         Index           =   18
         Left            =   120
         TabIndex        =   178
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FramePreguntaDevoluciones 
      Height          =   2175
      Left            =   3240
      TabIndex        =   167
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdDevolucion 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   172
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton optDevol 
         Caption         =   "Factura rectificativa"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   170
         Top             =   1080
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optDevol 
         Caption         =   "Albarán venta"
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   171
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   17
         Left            =   3960
         TabIndex        =   168
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Generar devolución"
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
         Index           =   17
         Left            =   600
         TabIndex        =   169
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FrameListadoComprasClientes 
      Height          =   5775
      Left            =   240
      TabIndex        =   160
      Top             =   720
      Width           =   10695
      Begin VB.CommandButton cmdTraerLineaCompraCliente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   8280
         TabIndex        =   163
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   9480
         TabIndex        =   164
         Top             =   5160
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   3615
         Index           =   4
         Left            =   120
         TabIndex        =   162
         Top             =   1440
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6376
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1095
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Factura"
            Object.Width           =   1589
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   1854
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Orig."
            Object.Width           =   972
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Dto1  "
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Dto2"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Importe"
            Object.Width           =   1781
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Nombre"
            Object.Width           =   4586
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   9
         Left            =   1320
         TabIndex        =   166
         Top             =   960
         Width           =   60
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Artículo: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   15
         Left            =   240
         TabIndex        =   165
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label9 
         Caption         =   "Compras efectuadas por el cliente cliente"
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
         Height          =   345
         Index           =   16
         Left            =   2880
         TabIndex        =   161
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame FrameMarjalChipos 
      Height          =   3135
      Left            =   2040
      TabIndex        =   147
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkCamposSocios 
         Caption         =   "Formato firma socio"
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   159
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox chkCamposSocios 
         Caption         =   "Excluir campos baja"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   150
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   157
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   149
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   148
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CommandButton cmdCamposSocio 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   151
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   15
         Left            =   4800
         TabIndex        =   152
         Top             =   2520
         Width           =   975
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   158
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   7
         Left            =   1080
         Picture         =   "frmListado5.frx":1FC48
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   6
         Left            =   1080
         Picture         =   "frmListado5.frx":1FD4A
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Socio/Cliente"
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
         TabIndex        =   156
         Top             =   840
         Width           =   1125
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   155
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label9 
         Caption         =   "Listado campos socios"
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
         Index           =   15
         Left            =   1200
         TabIndex        =   153
         Top             =   360
         Width           =   3885
      End
   End
   Begin VB.Frame FrameAlbaranesInternos 
      Height          =   3735
      Left            =   1200
      TabIndex        =   126
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdListadoAlbInt 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   131
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   9
         Left            =   4560
         TabIndex        =   130
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   129
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   128
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   137
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   133
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   127
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   4800
         TabIndex        =   132
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   3600
         TabIndex        =   141
         Top             =   2520
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   9
         Left            =   4200
         Picture         =   "frmListado5.frx":1FE4C
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   12
         Left            =   600
         TabIndex        =   140
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label lblDpto 
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
         Index           =   13
         Left            =   240
         TabIndex        =   139
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   1080
         Picture         =   "frmListado5.frx":1FED7
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   5
         Left            =   1080
         Picture         =   "frmListado5.frx":1FF62
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   138
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Listado albaranes internos"
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
         Index           =   13
         Left            =   1200
         TabIndex        =   136
         Top             =   240
         Width           =   3795
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   135
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label lblDpto 
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
         Index           =   12
         Left            =   240
         TabIndex        =   134
         Top             =   720
         Width           =   585
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   4
         Left            =   1080
         Picture         =   "frmListado5.frx":20064
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame FrameFitoCampos 
      Height          =   4695
      Left            =   4080
      TabIndex        =   102
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.OptionButton optFitoCampos 
         Caption         =   "Cliente - campos"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   111
         Top             =   3600
         Width           =   2175
      End
      Begin VB.OptionButton optFitoCampos 
         Caption         =   "Campos - Cliente"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   110
         Top             =   3600
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdFitoCampos 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   112
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   7
         Left            =   4560
         TabIndex        =   106
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   105
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Frame FrameCodtipom 
         Height          =   615
         Left            =   360
         TabIndex        =   120
         Top             =   2760
         Width           =   5535
         Begin VB.CheckBox chkCodtipom 
            Caption         =   "Servicios"
            Height          =   195
            Index           =   2
            Left            =   4200
            TabIndex        =   109
            Tag             =   "FAS"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkCodtipom 
            Caption         =   "Internas"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   108
            Tag             =   "FAI"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkCodtipom 
            Caption         =   "Ventas"
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   107
            Tag             =   "FAV"
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label frmet 
            AutoSize        =   -1  'True
            Caption         =   "Facturas"
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
            Left            =   0
            TabIndex        =   125
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   118
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   104
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   103
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   4800
         TabIndex        =   113
         Top             =   4080
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   4200
         Picture         =   "frmListado5.frx":20166
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   11
         Left            =   3600
         TabIndex        =   124
         Top             =   2280
         Width           =   465
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1200
         Picture         =   "frmListado5.frx":201F1
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         TabIndex        =   123
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   10
         Left            =   720
         TabIndex        =   122
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblIndicador 
         Caption         =   "I"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   121
         Top             =   4080
         Width           =   2895
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   119
         Top             =   1440
         Width           =   420
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   3
         Left            =   1200
         Picture         =   "frmListado5.frx":2027C
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Fitosanitarios x Campos"
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
         TabIndex        =   117
         Top             =   240
         Width           =   3510
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   2
         Left            =   1200
         Picture         =   "frmListado5.frx":2037E
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   9
         Left            =   360
         TabIndex        =   116
         Top             =   600
         Width           =   585
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   115
         Top             =   960
         Width           =   450
      End
   End
   Begin VB.Frame FrameGessocial 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame FrameFechaBaja 
         Caption         =   "Frame1"
         Height          =   735
         Left            =   960
         TabIndex        =   40
         Top             =   2040
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox txtFecha 
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   41
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   3
            Left            =   1080
            Picture         =   "frmListado5.frx":20480
            Top             =   240
            Width           =   240
         End
         Begin VB.Label lblDpto 
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
            Index           =   3
            Left            =   480
            TabIndex        =   42
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Baja"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   39
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Frame FrameGasol 
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   6135
         Begin VB.ComboBox cboEntidades 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   5775
         End
         Begin VB.Label Label1 
            Caption         =   "Colectivo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Actualizar"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   5
         Top             =   3360
         Width           =   1335
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Crear"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   2
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdGessocial 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   1
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
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
         TabIndex        =   3
         Top             =   240
         Width           =   5145
      End
   End
   Begin VB.Frame FrameSeleccionarFamilia 
      Height          =   5415
      Left            =   240
      TabIndex        =   64
      Top             =   240
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdSeleccionarFamilia 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   68
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5040
         TabIndex        =   65
         Top             =   4920
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   4095
         Index           =   0
         Left            =   240
         TabIndex        =   67
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7223
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmListado5.frx":2050B
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmListado5.frx":20655
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Seleccionar familias"
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
         Left            =   960
         TabIndex        =   66
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame FrameCambioProveedorPedido 
      Height          =   7215
      Left            =   0
      TabIndex        =   96
      Top             =   0
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton cmdCambiarProvePedido 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7200
         TabIndex        =   101
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   8280
         TabIndex        =   98
         Top             =   6720
         Width           =   975
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5775
         Index           =   2
         Left            =   120
         TabIndex        =   99
         Top             =   840
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   10186
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3775
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Dto1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Dto2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Observa"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Líneas a cambiar"
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
         TabIndex        =   100
         Top             =   600
         Width           =   1425
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   480
         Picture         =   "frmListado5.frx":2079F
         Top             =   6840
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   120
         Picture         =   "frmListado5.frx":208E9
         Top             =   6840
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Cambiar proveedor en pedido"
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
         Index           =   11
         Left            =   4560
         TabIndex        =   97
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame FrameGesocialCambioSituacion 
      Height          =   5895
      Left            =   120
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtModificable 
         Height          =   1575
         Index           =   0
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Text            =   "frmListado5.frx":20A33
         Top             =   3480
         Width           =   4695
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   1365
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   1680
         Width           =   4695
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Guardar"
         Height          =   375
         Index           =   5
         Left            =   3960
         TabIndex        =   47
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Index           =   6
         Left            =   240
         TabIndex        =   53
         Top             =   3240
         Width           =   1245
      End
      Begin VB.Label lblDpto 
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
         Index           =   5
         Left            =   240
         TabIndex        =   52
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Asociado"
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
         TabIndex        =   51
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Cambio situación asociado"
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
         Left            =   600
         TabIndex        =   48
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame FrameAguaMod 
      Height          =   3735
      Left            =   1680
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton cmdCambiarConsumo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2400
         TabIndex        =   71
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   3600
         TabIndex        =   72
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Consumo"
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   79
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Lectura modificada"
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
         Index           =   6
         Left            =   1560
         TabIndex        =   78
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Lectura actual"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   77
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Lectura anterior"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   75
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Modificar lectura facturada"
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
         TabIndex        =   73
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FrameSubRPT 
      Height          =   3015
      Left            =   1440
      TabIndex        =   54
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdSubRPT 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   57
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtModificable 
         Height          =   285
         Index           =   2
         Left            =   1320
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   2000
         Width           =   2055
      End
      Begin VB.TextBox txtModificable 
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   100
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   1440
         Width           =   4695
      End
      Begin VB.TextBox txtNoModificable 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5040
         TabIndex        =   58
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Informe"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   63
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   62
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Linea"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   61
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Informe asociado"
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
         Left            =   1200
         TabIndex        =   59
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame FrameContadoresAgua 
      Height          =   3615
      Left            =   0
      TabIndex        =   463
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   471
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   470
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   469
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtDescClie 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   468
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   467
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Contador"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   466
         Top             =   2520
         Width           =   1455
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   465
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton cmdContadorAgua 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   464
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Listado contadores agua"
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
         Left            =   1080
         TabIndex        =   475
         Top             =   360
         Width           =   3510
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmListado5.frx":20A39
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblDpto 
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
         Index           =   34
         Left            =   120
         TabIndex        =   474
         Top             =   840
         Width           =   585
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   126
         Left            =   480
         TabIndex        =   473
         Top             =   1200
         Width           =   450
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   472
         Top             =   1560
         Width           =   420
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmListado5.frx":20B3B
         Top             =   1560
         Width           =   240
      End
   End
   Begin VB.Frame FrameEliminarPresupuestos 
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdElimPresu 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   12
         Top             =   2040
         Width           =   975
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   16
         Top             =   1200
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3240
         Picture         =   "frmListado5.frx":20C3D
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Eliminar presupuestos FAZ"
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
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   3825
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmListado5.frx":20CC8
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.Frame FrameReimpresionSignotec 
      Height          =   6735
      Left            =   1440
      TabIndex        =   203
      Top             =   480
      Width           =   6735
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   21
         Left            =   5280
         TabIndex        =   206
         Top             =   6000
         Width           =   975
      End
      Begin MSComctlLib.ListView lwSigno 
         Height          =   5175
         Left            =   360
         TabIndex        =   204
         Top             =   720
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   9128
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Factura"
            Object.Width           =   1763
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Documento"
            Object.Width           =   5186
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Reimpresion albaranes/facturas firmadas"
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
         Index           =   21
         Left            =   360
         TabIndex        =   205
         Top             =   240
         Width           =   5925
      End
   End
   Begin VB.Frame FrameTrata 
      Height          =   3975
      Left            =   2760
      TabIndex        =   423
      Top             =   2880
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CheckBox chkVarios 
         Caption         =   "Desglosar tratamiento"
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
         Index           =   8
         Left            =   3360
         TabIndex        =   428
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   1440
         TabIndex        =   427
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   21
         Left            =   1440
         TabIndex        =   426
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtDesVario 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   434
         Top             =   1560
         Width           =   4455
      End
      Begin VB.TextBox txtCodigoVario 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1440
         TabIndex        =   425
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtDesVario 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   431
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox txtCodigoVario 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   424
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdTratamientosLis 
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
         Left            =   4440
         TabIndex        =   429
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   39
         Left            =   5760
         TabIndex        =   430
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "F. fin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   51
         Left            =   240
         TabIndex        =   438
         Top             =   2880
         Width           =   540
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1200
         Picture         =   "frmListado5.frx":20D53
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Listado tratamientos"
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
         Index           =   38
         Left            =   1920
         TabIndex        =   437
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "F. inicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   50
         Left            =   240
         TabIndex        =   436
         Top             =   2400
         Width           =   840
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   21
         Left            =   1200
         Picture         =   "frmListado5.frx":20DDE
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   36
         Left            =   360
         TabIndex        =   435
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image imgCodigoVario 
         Height          =   240
         Index           =   3
         Left            =   1200
         Picture         =   "frmListado5.frx":20E69
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   35
         Left            =   360
         TabIndex        =   433
         Top             =   1080
         Width           =   600
      End
      Begin VB.Image imgCodigoVario 
         Height          =   240
         Index           =   2
         Left            =   1200
         Picture         =   "frmListado5.frx":20F6B
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tratamiento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   49
         Left            =   240
         TabIndex        =   432
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   20835
      Picture         =   "frmListado5.frx":2106D
      Tag             =   "-1"
      ToolTipText     =   "Buscar almacén"
      Top             =   90
      Width           =   240
   End
End
Attribute VB_Name = "frmListado5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public OpcionListado As Integer

Public OtrosDatos As String

    '==== Opciones ====
    '=============================
    '   0.-  Listados contaqdores de agua
    '   YA NO ESTA 1.-  GESSOCIAL    Crear asociado en la seccion    'NOVIEMBRE 2014. Lo quito de aqui
    '   2.-  Eliminar presupuestos "HERBELCA" Pide las fechas para la siguiente pantalla
    
    '   3.- Asignar descuentos  actividad desde sfamia
    '   4.- Insertar en sdtofm desde sactivdto
        
    '   5.- Gessocial. Cambio situacion socio(alta-baja-situacion o essocio)
    
    '   6.- Subreports de la SCRYST
    
    '   7.- Seleccion de familias
    '        Vendran las familias en otrosdatos y cadenadesdeotrofrm mostrara las que quiera
    
    '   8.- Modificar consumo en facturas de agua
    
    '   9.- Alzira. Ajuste compras tratamientos
    
    '   10.- Seleccion PROVEEDORES.  A partir de un select.....
        
    '   11.- Cambiar proveedor del pedido despues de una simulacion
      
    '   12.- Fito santiarios por campos
    '   13.- Listado albaranes internos (perosnalizable)
    
    
    '   14.- EULER. Devolverá un ALBARAN   o un pedido
    
    '   15.- Marjal-Chipos.  Informe campos socio
    
    '   16.- Listado compras cliente desde una fecha
    '   17.- devoluciones. Preguta si pasa a albaran o a frt
    
    '   18.-  Telefonia,  Impresion contrato.  Importet terminal y meses
    '   19.-  Puntos cliente
    '   20.-  Declaracion alcohol FONTENAS
    '   21.-  Reimpresion facturas/albaranes firmados SIGNOTEC
    '   22.-  CRM:  Clientes     por accion comercial
    '   23.-  Cambiar orden lineas pedido /albaranes
    '   24.-  Listado albaranes
    '   25.-  credio y caucion
    '   26.-  Importacion COARVAl
    '   27.-  Lineas especiales albaran euler  FACTURA
    '   28.-    "       ""                     ALBARANES
    '   29.-    Taxco   Nueva entrad vehcilu.
    '   30.-  importe cierre turno traspaso ALVIC
    
    '   31.- Prevision facturacion TAXCO
    
    '   32.-  Pedido cliente generar pedido proveedor
    '   33.-  Vista previa pedidos cliente(FONTENAS).
    '   34.-  Cambiar albaranes de ALVIC a ogtro codigo de cliente(comprobando NIFs)
    
    '   35.- Cambio de datos de un cliente a otro
    '   36.- Listado pedido por dia
    
    '   37.- Listado comparativo descuentos Compra-venta
    
    '   38.- Coste linea
    '   39.- Impresion tratamientos
    
    '   40.- Actualizar precio minimo articulos desde promociones
    
    '   41.- Copiar precios compra a otros proveedor
    '   42.- '   27y 28  Lineas especiales albaran euler  PROYECTO
    '   43.- EULER Proyecto Asignar albaranes cliente
    '   44.- CESTA dese App movil
    
    '   45.-  TAXCO.   Modificar en la svenci de las facturas ALVIC 'FA1','FA2','FAD','FAB'
    
    
Private WithEvents frmCli As frmBasico2
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmAc As frmBasico2 '%=%=frmFacAgentesCom
Attribute frmAc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmB3 As frmFacActividades
Attribute frmB3.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmBasico2 'frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmP As frmBasico2
Attribute frmP.VB_VarHelpID = -1

Private WithEvents frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1

Private PuedeCerrar As Boolean

Dim miSQL As String
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private cadPDFrpt As String 'Nombre del informe a enviar por email
Private vMostrarTree As Boolean

Private PrimVez As Boolean

Private auxiliar As String  ' Para quitar proveedores serviara para guardar cuales quito
Private Colec As Collection

Dim CadArticulos As String


Private Sub cboDestinoB_KeyPress(KeyAscii As Integer)
 KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cboTipoDescuento_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cboTrimiestre_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkAlb_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkCamposSocios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkCodtipom_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub chkPedidos_Click()
    If PrimVez Then Exit Sub
    
    CargaAlbaranesFacturaClienteEuler
End Sub

Private Sub chkVarios_Click(Index As Integer)
    If Index = 9 Then
        'Fecha menor iguial en promociones
        txtFecha(24).visible = chkVarios(9).Value = 0
        Label4(38).visible = chkVarios(9).Value = 0
        imgFecha(24).visible = chkVarios(9).Value = 0
        
        imgFecha(23).Left = IIf(chkVarios(9).Value = 0, 1200, 2400)
        txtFecha(23).Left = imgFecha(23).Left + 360
        
        Label4(37).Caption = IIf(chkVarios(9).Value = 0, "Inicio", "Fec. menor o igual ")
    End If
    
End Sub

Private Sub chkVarios_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub CmdAcepCopia_Click()
Dim SQL As String

    If txtProve(2).Text = "" Then
        MsgBox "Debe introducir el proveedor origen.", vbExclamation
        PonerFoco txtProve(2)
        Exit Sub
    End If
    
    If txtDescProve(2).Text = "" Then
        MsgBox "El proveedor origen ha de existir. Reintroduzca.", vbExclamation
        PonerFoco txtProve(2)
        Exit Sub
    End If
    
    If txtProve(3).Text = "" Then
        MsgBox "Debe introducir el proveedor destino.", vbExclamation
        PonerFoco txtProve(3)
        Exit Sub
    End If
    
    If txtDescProve(3).Text = "" Then
        MsgBox "El proveedor destino ha de existir. Reintroduzca.", vbExclamation
        PonerFoco txtProve(3)
        Exit Sub
    End If
    
    If txtProve(2).Text = txtProve(3) Then
        MsgBox "El proveedor origen no ha de coincidir con el proveedor destino. Revise.", vbExclamation
        PonerFoco txtProve(2)
        Exit Sub
    End If
    
    If txtFamia(3).Text = "" Then
        SQL = "select count(*) from slispr where codprove = " & DBSet(txtProve(2), "N")
        If TotalRegistros(SQL) = 0 Then
            MsgBox "El proveedor origen no tiene artículos con precios. Revise.", vbExclamation
            Exit Sub
        End If
    Else
        SQL = "select count(*) from slispr inner join sartic on slispr.codartic = sartic.codartic where slispr.codprove = " & DBSet(txtProve(2), "N")
        SQL = SQL & " and sartic.codfamia = " & DBSet(txtFamia(3).Text, "N")
        If TotalRegistros(SQL) = 0 Then
            MsgBox "El proveedor origen no tiene artículos con precios en esa familia. Revise.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' si hay articulos del destino que los tenga el origen los mostramos
    SQL = "select count(*) from slispr where codprove = " & DBSet(txtProve(3), "N")
    SQL = SQL & " and codartic in (select slispr.codartic from slispr inner join sartic on slispr.codartic = sartic.codartic where slispr.codprove = " & DBSet(txtProve(2), "N")
    If txtFamia(3).Text <> "" Then
        SQL = SQL & " and sartic.codfamia = " & DBSet(txtFamia(3).Text, "N") & ")"
    Else
        SQL = SQL & ")"
    End If
        
    
    CadArticulos = ""
    
    If TotalRegistros(SQL) <> 0 Then
        Dim Sql2 As String
        
        Sql2 = "slispr.codprove = " & DBSet(txtProve(3), "N") & " and slispr.codartic in (select slispr.codartic from slispr inner join sartic on slispr.codartic = sartic.codartic where slispr.codprove = " & DBSet(txtProve(2), "N")
        If txtFamia(3).Text <> "" Then
            Sql2 = Sql2 & " and sartic.codfamia = " & DBSet(txtFamia(3).Text, "N") & ")"
        Else
            Sql2 = Sql2 & ")"
        End If
        
        
        Set frmMens = New frmMensajes
        frmMens.cadWhere = Sql2
        frmMens.OpcionMensaje = 28
        frmMens.Show vbModal
    
    End If
    
    If CadArticulos <> "NO" Then
        If CopiaPrecios Then
            MsgBox "Proceso finalizado.", vbInformation
            Unload Me
        End If
    End If
    
End Sub

Private Function CopiaPrecios() As Boolean
Dim SQL As String

    On Error GoTo eCopiaPrecios
    
    CopiaPrecios = False

    conn.BeginTrans
    
    ' insertamos los que no tiene
    SQL = "insert ignore into slispr (codartic,codprove,precioac,fechanue,precionu,dtopermi,cantfija,cantmini,fechaini,fechafin,preciopr,dtoperm1,dtoline1,dtoline2,precioexp,referprov,descripprov)   "
    SQL = SQL & " select slispr.codartic, " & DBSet(txtProve(3), "N") & ",slispr.precioac,slispr.fechanue,slispr.precionu,slispr.dtopermi,slispr.cantfija,slispr.cantmini,slispr.fechaini,slispr.fechafin,slispr.preciopr,slispr.dtoperm1,slispr.dtoline1,slispr.dtoline2,slispr.precioexp,slispr.referprov,slispr.descripprov "
    SQL = SQL & " from slispr inner join sartic on slispr.codartic = sartic.codartic "
    SQL = SQL & " where slispr.codprove = " & DBSet(txtProve(2), "N")
    If txtFamia(3).Text <> "" Then SQL = SQL & " and sartic.codfamia = " & DBSet(txtFamia(3).Text, "N")
    
    If CadArticulos <> "" Then SQL = SQL & " and not slispr.codartic in " & CadArticulos
    
    conn.Execute SQL
    
    ' actualizamos los que me han dicho
    If CadArticulos <> "" Then
        SQL = "update slispr dd, slispr ff set "
        SQL = SQL & " dd.precioac=ff.precioac,"
        SQL = SQL & " dd.fechanue=ff.fechanue,"
        SQL = SQL & " dd.precionu=ff.precionu,"
        SQL = SQL & " dd.dtopermi=ff.dtopermi,"
        SQL = SQL & " dd.cantfija=ff.cantfija,"
        SQL = SQL & " dd.cantmini=ff.cantmini,"
        SQL = SQL & " dd.fechaini=ff.fechaini,"
        SQL = SQL & " dd.fechafin=ff.fechafin,"
        SQL = SQL & " dd.preciopr=ff.preciopr,"
        SQL = SQL & " dd.dtoperm1=ff.dtoperm1,"
        SQL = SQL & " dd.dtoline1=ff.dtoline1,"
        SQL = SQL & " dd.dtoline2=ff.dtoline2,"
        SQL = SQL & " dd.precioexp=ff.precioexp,"
        SQL = SQL & " dd.referprov=ff.referprov,"
        SQL = SQL & " dd.descripprov=ff.descripprov "
        SQL = SQL & " where ff.codprove = " & DBSet(txtProve(2), "N")
        SQL = SQL & " and dd.codprove = " & DBSet(txtProve(3), "N")
        SQL = SQL & " and dd.codartic in " & CadArticulos
        SQL = SQL & " and dd.codartic = ff.codartic "
        
        conn.Execute SQL
    
    End If
    
    
    conn.CommitTrans
    CopiaPrecios = True
    Exit Function


eCopiaPrecios:
    conn.RollbackTrans
    MuestraError Err.Number, "Copia Precios", Err.Description
End Function


Private Sub cmdAceptarLinEspEuler_Click()
    cadParam = ""
    For numParam = 4 To 7
        
            
            If numParam < 6 Then
                If txtModificable(numParam).Text = "" Then cadParam = cadParam & "- " & RecuperaValor("Articulo|Descripcion|", CInt(numParam - 3)) & vbCrLf
            End If
            If numParam <> 6 Then
                If txtNumero(numParam + 1).Text = "" Then cadParam = cadParam & "- " & RecuperaValor("Cantidad|Precio||importe|", CInt(numParam - 3)) & vbCrLf
            End If
      
    Next
    If cadParam <> "" Then
        cadParam = "Campos obligatorios" & vbCrLf & cadParam
        MsgBox cadParam, vbExclamation
        Exit Sub
    End If
    
    'OK, luego insertamos/modificamos
    If OpcionListado = 27 Then
        'FACTURAS
        If Not lblDpto(37).visible Then
            cadPDFrpt = " codtipom= " & DBSet(RecuperaValor(OtrosDatos, 1), "T")
            cadPDFrpt = cadPDFrpt & "  AND numfactu= " & DBSet(RecuperaValor(OtrosDatos, 2), "N")
            cadPDFrpt = cadPDFrpt & " AND fecfactu = " & DBSet(RecuperaValor(OtrosDatos, 3), "F")
            cadPDFrpt = cadPDFrpt & "  AND codtipoa= " & DBSet(RecuperaValor(OtrosDatos, 4), "T")
            cadPDFrpt = cadPDFrpt & "  AND numalbar= " & DBSet(RecuperaValor(OtrosDatos, 5), "N")
            
            
            cadPDFrpt = DevuelveDesdeBD(conAri, "max(numlinea)", "slifac_eu2", cadPDFrpt & " AND 1", "1")
            cadPDFrpt = Val(cadPDFrpt) + 1
        Else
            cadPDFrpt = lblDpto(37).Tag
        End If
           
        cadParam = "REPLACE INTO slifac_eu2(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,articulo,descrarticulo,cantidad,precioar,dtoline1,importel) VALUES ("
        cadParam = cadParam & DBSet(RecuperaValor(OtrosDatos, 1), "T") & "  , " & DBSet(RecuperaValor(OtrosDatos, 2), "N")
        cadParam = cadParam & "," & DBSet(RecuperaValor(OtrosDatos, 3), "F") & "," & DBSet(RecuperaValor(OtrosDatos, 4), "T")
        cadParam = cadParam & "," & DBSet(RecuperaValor(OtrosDatos, 5), "N") & "," & cadPDFrpt
        cadParam = cadParam & "," & DBSet(txtModificable(4).Text, "T") & ",'" & DevNombreSQL(txtModificable(5).Text)
        cadParam = cadParam & "'," & DBSet(txtNumero(5).Text, "N") & "," & DBSet(txtNumero(6).Text, "N")
        cadParam = cadParam & "," & DBSet(txtNumero(7).Text, "N") & "," & DBSet(txtNumero(8).Text, "N") & ")"
    ElseIf OpcionListado = 28 Then
        'Albaranes
        If Not lblDpto(37).visible Then
            cadPDFrpt = " codtipom= " & DBSet(RecuperaValor(OtrosDatos, 1), "T")
            cadPDFrpt = cadPDFrpt & "  AND numalbar= " & DBSet(RecuperaValor(OtrosDatos, 2), "N")
            
            
            cadPDFrpt = DevuelveDesdeBD(conAri, "max(numlinea)", "slialb_eu2", cadPDFrpt & " AND 1", "1")
            cadPDFrpt = Val(cadPDFrpt) + 1
        Else
            cadPDFrpt = lblDpto(37).Tag
        End If
            
        cadParam = "REPLACE INTO slialb_eu2(codtipom,Numalbar,numlinea,articulo,descrarticulo,cantidad,precioar,dtoline1,importel) VALUES ("
        cadParam = cadParam & DBSet(RecuperaValor(OtrosDatos, 1), "T") & "  , " & DBSet(RecuperaValor(OtrosDatos, 2), "N") & "," & cadPDFrpt
        cadParam = cadParam & "," & DBSet(txtModificable(4).Text, "T") & ",'" & DevNombreSQL(txtModificable(5).Text)
        cadParam = cadParam & "'," & DBSet(txtNumero(5).Text, "N") & "," & DBSet(txtNumero(6).Text, "N")
        cadParam = cadParam & "," & DBSet(txtNumero(7).Text, "N") & "," & DBSet(txtNumero(8).Text, "N") & ")"
    
    
    Else
        'PROYECTP
        
        If Not lblDpto(37).visible Then
            cadPDFrpt = " codtipom= " & DBSet(RecuperaValor(OtrosDatos, 1), "T")
            cadPDFrpt = cadPDFrpt & "  AND numproyec= " & DBSet(RecuperaValor(OtrosDatos, 2), "N")
            
            
            cadPDFrpt = DevuelveDesdeBD(conAri, "max(numlinea)", "sproyectolin2", cadPDFrpt & " AND 1", "1")
            If Val(cadPDFrpt) < 1000 Then cadPDFrpt = "1000"
            cadPDFrpt = Val(cadPDFrpt) + 1
        Else
            cadPDFrpt = lblDpto(37).Tag
        End If
            
        cadParam = "REPLACE INTO sproyectolin2(codtipom,numproyec,numlinea,articulo,descrarticulo,cantidad,precioar,dtoline1,importel) VALUES ("
        cadParam = cadParam & DBSet(RecuperaValor(OtrosDatos, 1), "T") & "  , " & DBSet(RecuperaValor(OtrosDatos, 2), "N") & "," & cadPDFrpt
        cadParam = cadParam & "," & DBSet(txtModificable(4).Text, "T") & ",'" & DevNombreSQL(txtModificable(5).Text)
        cadParam = cadParam & "'," & DBSet(txtNumero(5).Text, "N") & "," & DBSet(txtNumero(6).Text, "N")
        cadParam = cadParam & "," & DBSet(txtNumero(7).Text, "N") & "," & DBSet(txtNumero(8).Text, "N") & ")"
        
        
        
    End If
    If ejecutar(cadParam, False) Then
        Espera 0.25
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
    
End Sub

Private Sub cmdAceptarPedFontenas_Click()

    If lw(12).ListItems.Count = 0 Then Exit Sub
    If lw(12).SelectedItem Is Nothing Then Exit Sub
    
    CadenaDesdeOtroForm = lw(12).SelectedItem.SubItems(2)
    Unload Me
End Sub

Private Sub cmdAjusteVtosFaccliALVIC_Click()
    
    
    cadParam = "('FA1','FA2','FAD','FAB')"
    If txtFecha(25).Text = "" Or txtFecha(26).Text = "" Then
        'txtFecha
        cadSelect = "Fechas obligatorias"
    Else
        If CDate(txtFecha(25).Text) < CDate(cadFormula) Then
            cadSelect = "Fecha inicio a un ajuste YA realizado (" & cadFormula & ")"
        Else
            'Comprobaremos que HAY facturas en el intervalo
            cadSelect = "fecfactu >=" & DBSet(txtFecha(25).Text, "F") & " AND fecfactu <=" & DBSet(txtFecha(26).Text, "F")
            cadSelect = cadSelect & " AND codtipom in " & cadParam & " AND 1"
            cadSelect = DevuelveDesdeBD(conAri, "count(*)", "scafac", cadSelect, "1")
            If Val(cadSelect) > 0 Then
                cadSelect = ""
            Else
                cadSelect = "NO existe facturas ALVIC para ese intervalo"
            End If
        End If
    End If
    If cadSelect <> "" Then
        MsgBox cadSelect, vbExclamation
        Exit Sub
    End If
    
    If MsgBox("¿Continuar con el proceso?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    conSubRPT = AjusteSvenciFacturasAlvic
    Screen.MousePointer = vbDefault
    lblIndicador(9).Caption = ""
    
    If conSubRPT Then Unload Me
    
End Sub

Private Sub cmdAlvic_Click()
    
    
    'Comprobaciones
    ComprobarImportesAlvic
    
    If lblDpto(42).Tag <> 0 Then
        MsgBox "Existe diferencia entre importes ajustados y traspaso" & vbCrLf & vbCrLf & lblDpto(42).Caption, vbExclamation
    
    Else
        'Ajustamos
        
        miSQL = ""
        For numParam = 1 To lw(9).ListItems.Count
            miSQL = miSQL & ", (" & vUsu.Codigo & "," & lw(9).ListItems(numParam).Text & "," & DBSet(lw(9).ListItems(numParam).Tag, "N") & ")"
        Next
        If miSQL <> "" Then
            miSQL = Mid(miSQL, 2)
            
            miSQL = "REPLACE INTO tmpscapla (codusu,codplant,cantidad ) VALUES " & miSQL
            If ejecutar(miSQL, False) Then
                CadenaDesdeOtroForm = "OK"
                Unload Me
            Else
                CadenaDesdeOtroForm = ""
            End If
        End If
    End If
End Sub

Private Sub cmdAsignarAlbarEuler_Click()
    Screen.MousePointer = vbHourglass
    numParam = CByte(Abs((ModificarAlbaranesVinculados)))
    lblIndicador(8).Caption = ""
    
    Screen.MousePointer = vbDefault
    If numParam = 1 Then
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If
        
 End Sub
 
 
 
Private Sub cmdCambiarConsumo_Click()
    If txtNumero(0).Text = "" Then Exit Sub
    
    miSQL = RecuperaValor(OtrosDatos, 3)
    If Val(miSQL) <> Val(Label1(7).Tag) Then
        
        If MsgBox("¿Seguro que desea cambiar el consumo?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Screen.MousePointer = vbHourglass
        conn.BeginTrans
        If Not HacerUpdateConsumo Then
            conn.RollbackTrans
        Else
            conn.CommitTrans
            CadenaDesdeOtroForm = "OK"
        End If
        Screen.MousePointer = vbDefault
    End If
    Unload Me
End Sub

Private Sub cmdCambiarProvePedido_Click()
    miSQL = ""
    For NumRegElim = 1 To lw(2).ListItems.Count
        If lw(2).ListItems(NumRegElim).Checked Then miSQL = miSQL & ", " & lw(2).ListItems(NumRegElim).Tag
    Next
    If miSQL = "" Then
        MsgBox "Seleccione alguna linea de articulo para cambiar de proveedor", vbExclamation
    Else
        CadenaDesdeOtroForm = miSQL
        Unload Me
    End If
End Sub

Private Sub cmdCambioCliente_Click()
    cadParam = ""
    If Me.txtCliente(15).Text = "" Then cadParam = "N"
    If Me.txtDescClie(15).Text = "" Then cadParam = "N"
    If Me.txtCliente(16).Text = "" Then cadParam = "N"
    If Me.txtDescClie(16).Text = "" Then cadParam = "N"
    
    If cadParam <> "" Then
        MsgBox "Clientes obligatorios", vbExclamation
        Exit Sub
    End If
    
    
    'Enero 2020
    'Aviso del NIF
    cadSelect = DevuelveDesdeBD(conAri, "nifclien", "sclien", "codclien", txtCliente(15).Text)
    If cadSelect = "" Then cadParam = "Cliente origen incorrecto. Falta NIF" & vbCrLf
    cadFormula = DevuelveDesdeBD(conAri, "nifclien", "sclien", "codclien", txtCliente(16).Text)
    If cadFormula = "" Then cadParam = cadParam & "Cliente destino incorrecto. Falta NIF" & vbCrLf
    If cadParam <> "" Then
        MsgBox cadParam, vbExclamation
    Else
        If cadFormula <> cadSelect Then
            cadParam = "NIF origen - destino distintos" & vbCrLf & "¿Continuar?"
            If MsgBox(cadParam, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then cadParam = ""
        End If
    End If
    If cadParam <> "" Then Exit Sub
    
    
    If MsgBox("Seguro que quiere cambiar el cliente?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
    miSQL = InputBox("Password de seguirdad", "Cambio cliente")
    If UCase(miSQL) <> "ARIADNA" Then
        MsgBox "password incorrecto", vbExclamation
        Exit Sub
    End If
    
    
    
    
    
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    
    If RealizaUpdatesCambioCliente Then
        conn.CommitTrans
        miSQL = "Origen " & Me.txtCliente(15).Text & " " & Me.txtDescClie(15).Text & vbCrLf
        miSQL = miSQL & "Destino " & Me.txtCliente(16).Text & " " & Me.txtDescClie(16).Text & vbCrLf
        
        Set LOG = New cLOG
        LOG.Insertar 40, vUsu, miSQL
        Set LOG = Nothing
        CambiReferenciaCliente
        MsgBox "Proceso realizado correctamente", vbInformation
        
        
        
    Else
        conn.RollbackTrans
    End If
    lblIndicador(5).Caption = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCamposSocio_Click()
    'a
    
    InicializarVbles True

    If Me.chkCamposSocios(0).Value = 1 Then
        miSQL = "sclienhuertos.fecbajas"
        cadSelect = "  (" & miSQL & ") is null"
        cadFormula = " isnull({" & miSQL & "})"
    End If


    If txtCliente(6).Text <> "" Or txtCliente(7).Text <> "" Then
        miSQL = " Cliente: "
        cadTitulo = "{sclienhuertos.codclien}"
        If Not PonerDesdeHasta(cadTitulo, "CLI", 6, 7, miSQL) Then Exit Sub
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "              "
        cadPDFrpt = cadPDFrpt & miSQL
    End If
    If Me.chkCamposSocios(0).Value = 1 Then cadPDFrpt = Trim(cadPDFrpt & "     " & chkCamposSocios(0).Caption)
        
    cadParam = cadParam & "DesdeHasta=""" & Trim(cadPDFrpt) & """|"
    numParam = numParam + 1
    
    Screen.MousePointer = vbHourglass
    If Not HayRegParaInforme("sclienhuertos", cadSelect, True) Then
        MsgBox "No hay datos para mostrar con estos valores", vbExclamation

    Else
    
        cadTitulo = "Listado campos socios"
        
        If chkCamposSocios(1).Value = 1 Then
            cadNomRPT = "marListadoCli.rpt"
        Else
            cadNomRPT = "marListado.rpt"
        End If
        cadPDFrpt = ""
        conSubRPT = False
        
        LlamarImprimir

    End If
    Screen.MousePointer = vbDefault
        
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 5 Then
        'GESSOCIAL
        'Cambio de situacion
        'Pone SALIR.
        CadenaDesdeOtroForm = Trim(txtModificable(0).Text)
    
    ElseIf Index = 7 Then
        CadenaDesdeOtroForm = ""
    ElseIf Index = 11 Or Index = 44 Then
        CadenaDesdeOtroForm = ""
    ElseIf Index >= 17 And Index <= 19 Then
        CadenaDesdeOtroForm = ""
    ElseIf Index = 23 Then
        If numParam = 1 Then
            If MsgBox("Descartar cambios?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    ElseIf Index = 118 Then
        CadenaDesdeOtroForm = ""
    ElseIf Index = 26 Then
        If lw(7).Tag = 1 Then
            If MsgBox("Borrar datos importacion?", vbQuestion + vbYesNoCancel) = vbYes Then CargaColumnasCoarval True
            Exit Sub
        End If
        
        
        
    ElseIf Index = 38 Then
        If Me.txtNumero(9).Text <> "" Then
            CadenaDesdeOtroForm = txtNumero(9).Text
            PuedeCerrar = True
            
        End If
    End If
    Unload Me
    
    
End Sub

Private Sub cmdCargaDtoFamiliaActiv_Click()
    miSQL = ""
    
    If Me.txtActiv(0).Text = "" Then miSQL = miSQL & "-Actividad"
    If Me.cboTipoDescuento(0).ListIndex < 0 Then miSQL = miSQL & "-Tipo descuento"
    If miSQL <> "" Then
        MsgBox "Faltan campos: " & vbCrLf & miSQL, vbExclamation
        Exit Sub
    End If
    
    If MsgBox("¿continuar con el proceso de generacion de descuentos por actividad?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub

    
    'SI SOLO SON NUEVOS.
    ' INSERT IGNORE INTO
    ' Si son nuevos y actualizar, replace
    
    If Me.chkVarios(0).Value = 1 Then
        miSQL = " INSERT IGNORE "
    Else
        miSQL = "REPLACE "
    End If
    miSQL = miSQL & " INTO sactivdtos SELECT " & Me.txtActiv(0).Text & " codactiv,codfamia,clasifica "
    miSQL = miSQL & " FROM sfamiadtos where clasifica=" & cboTipoDescuento(0).ItemData(cboTipoDescuento(0).ListIndex)
    
    
    If ejecutar(miSQL, False) Then
        
        CadenaDesdeOtroForm = txtActiv(0).Text
        Unload Me
    End If
End Sub

Private Sub cmdCesta_Click()
    
    miSQL = ""
    cadSelect = ""
    cadFormula = ""
    cadNomRPT = ""
    For NumRegElim = 1 To Me.lw(15).ListItems.Count
        If cadNomRPT = "" Then cadNomRPT = lw(15).ListItems(NumRegElim).Tag
         'Si esta seleccionado lo va a traer, y si no, lo borraa
        If Me.lw(15).ListItems(NumRegElim).Checked Then
            cadFormula = cadFormula & "X"
            
        Else
            
            miSQL = miSQL & "X"
            cadSelect = cadSelect & ", " & Mid(lw(15).ListItems(NumRegElim).Key, 2)
        End If
    Next
    
    
    cadParam = "Proceso insertar desde cesta: " & vbCrLf
    
    cadParam = cadParam & "A insertar : " & Len(cadFormula) & IIf(Len(cadFormula) = 0, "   NINGUNA !!!! ", "") & vbCrLf
    If miSQL <> "" Then cadParam = cadParam & "No insertar : " & Len(miSQL) & vbCrLf
    cadParam = cadParam & vbCrLf & "¿Continuar?"
    If MsgBox(cadParam, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    
    If cadSelect <> "" Then
        cadSelect = Mid(cadSelect, 2) 'quitamos la primera coma
        cadSelect = "DELETE FROM cestas_lineas WHERE cestaLineaId IN (" & Trim(cadSelect) & ") AND cestaID=" & cadNomRPT
        If ejecutar(cadSelect, False) Then Espera 0.5
        
    End If
    CadenaDesdeOtroForm = cadNomRPT
    Unload Me
End Sub

Private Sub cmdComprasTratamientos_Click()
    
    numParam = 0
    If Me.txtFecha(5).Text = "" Then numParam = 5
    If Me.txtFecha(4).Text = "" Then numParam = 4
    If numParam > 0 Then
        MsgBox "Campos fecha son obligatorios", vbExclamation
        PonerFoco txtFecha(CInt(numParam))
        Exit Sub
    End If
    
    
    Me.lblIndicador(1).Caption = "Comienzo proceso"
    Screen.MousePointer = vbHourglass
    InicializarVbles True
    
    
    If GeneraDatosComprasTratamientos Then
        HaPulsadoElBotonDeImprimir = False
        cadTitulo = "Ajuste compras tratamientos"
        cadParam = cadParam & "pdh1=""Fechas: " & txtFecha(4).Text & " - " & txtFecha(5).Text & """|"
        numParam = numParam + 1
        
        cadParam = cadParam & "Detalle=" & Abs(Me.chkVarios(2).Value) & "|"
        numParam = numParam + 1
        vMostrarTree = True
        cadNomRPT = "rAjuCompTra.rpt"   'cadPDFrpt & ".rpt"
        cadPDFrpt = ""
        conSubRPT = False
        
        LlamarImprimir
        
        If HaPulsadoElBotonDeImprimir Then
            miSQL = "Va a generar el apunte. Continuar?"
            If MsgBox(miSQL, vbQuestion + vbYesNoCancel) = vbYes Then
                GenerarApunteAjusteTratamientos
                Unload Me
            End If
            
       End If
    End If
    
    Me.lblIndicador(1).Caption = ""
    Screen.MousePointer = vbDefault
    
        
End Sub

Private Sub cmdContadorAgua_Click()

    InicializarVbles True

    If txtCliente(0).Text <> "" Or txtCliente(1).Text <> "" Then
        miSQL = " Cliente: "
        cadTitulo = "{aguacontadores.codclien}"
        If Not PonerDesdeHasta(cadTitulo, "CLI", 0, 1, miSQL) Then Exit Sub
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "              "
        cadPDFrpt = cadPDFrpt & miSQL
        
        cadParam = cadParam & "DesdeHasta=""" & cadPDFrpt & """|"
        numParam = numParam + 1
    End If
    
    
    Screen.MousePointer = vbHourglass
    If Not HayRegParaInforme("aguacontadores", cadSelect, True) Then
        MsgBox "No hay datos para mostrar con estos valores", vbExclamation

    Else
    
        cadTitulo = "Listado contadores agua"
        If optVarios(0).Value Then
            cadPDFrpt = "rAgua1"
        Else
            cadPDFrpt = "rAgua2"
            vMostrarTree = True
        End If
        cadNomRPT = cadPDFrpt & ".rpt"
        cadPDFrpt = ""
        conSubRPT = False
        
        LlamarImprimir

    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdContratoTelef_Click()

    CadenaDesdeOtroForm = Trim(txtNumero(1).Text) & "|" & Trim(txtNumero(2).Text) & "|"
    Unload Me
End Sub

Private Sub cmdCreditoCaucion_Click()
    If txtFecha(15).Text = "" Or txtFecha(16).Text = "" Then
        MsgBox "Fechas requeridas", vbExclamation
        Exit Sub
    
    End If
    
    Screen.MousePointer = vbHourglass
    vMostrarTree = HacerCreditoYCaucion
    Screen.MousePointer = vbDefault
    If vMostrarTree Then  'ha ido buien
         InicializarVbles True
        
        If Me.chkVarios(5).Value = 1 Then
            GenerarFicheroCreditoYCaucion
        Else
            cadTitulo = "Listado ventas credito y caucion"
        
            cadNomRPT = "rCredCaul.rpt"
            cadParam = cadParam & "pDH= ""Desde " & txtFecha(15).Text & " hasta " & txtFecha(16).Text & """|"
            numParam = numParam + 1
            conSubRPT = False
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            LlamarImprimir
        End If
    End If
End Sub

Private Sub cmdCRMClieAccion_Click()
Dim AuxF As String
Dim b As Boolean
    'scrmacciones agente codtraba tipo fechora codclien
    If Me.optVarios(6).Value Then conn.Execute "DELETE from tmpcrmclien WHERE codusu = " & vUsu.Codigo
    
    miSQL = ""
    cadParam = ""
    OtrosDatos = ""
    For NumRegElim = 1 To lw(5).ListItems.Count
        If lw(5).ListItems(NumRegElim).Checked Then
            cadParam = cadParam & "X"
            miSQL = Trim(miSQL & "                -" & Replace(lw(5).ListItems(NumRegElim).SubItems(1), """", "'") & "(" & lw(5).ListItems(NumRegElim).Text & ")")
            If (Len(cadParam) Mod 3) = 0 Then miSQL = miSQL & """ + chr(13) + """
            OtrosDatos = OtrosDatos & ", " & lw(5).ListItems(NumRegElim).Text
        End If
    Next
    If Len(cadParam) = 0 Then
        MsgBox "Seleccione alguna accion", vbExclamation
        Exit Sub
    End If
    
    If Len(cadParam) = lw(5).ListItems.Count Then auxiliar = ""         'TODAS LAS ACCIONES
    OtrosDatos = Mid(OtrosDatos, 2)
    auxiliar = miSQL
    InicializarVbles True

    
    miSQL = "Fecha: "
    If txtFecha(11).Text = "" And txtFecha(12).Text = "" Then
        miSQL = ""
    Else
        If Not PonerDesdeHasta("{scrmacciones.fechora}", "F", 11, 12, miSQL) Then Exit Sub
        If Me.optVarios(6).Value Then
            'Reestablezco las variables guardandome la fecha
            AuxF = cadSelect
            cadFormula = "": cadSelect = ""
        End If
    End If
    If Me.optVarios(5).Value Then
        miSQL = "VISITADOS         " & miSQL
    Else
        miSQL = "** NO VISITADOS **        " & miSQL
    End If
    If auxiliar <> "" Then miSQL = miSQL & """ + chr(13) + ""Acciones: " & auxiliar
    cadParam = cadParam & "pDH=""" & miSQL & """|"
    
 
    If txtAgente(0).Text <> "" Or txtAgente(1).Text <> "" Then
    
        If Me.optVarios(5).Value Then
            auxiliar = "{scrmacciones.agente}"
        Else
            auxiliar = "{sclien.codagent}"
        End If
        
        If vParamAplic.NumeroInstalacion = vbHerbelca Then auxiliar = "{sclien.visitador}"
        
        miSQL = "pDHAgente=""" & IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Visitador  ", "Agente ")
        If Not PonerDesdeHasta(auxiliar, "AGT", 0, 1, miSQL) Then Exit Sub
    End If
    
    If txtCliente(10).Text <> "" Or txtCliente(11).Text <> "" Then
        auxiliar = "{scrmacciones"
        If Me.optVarios(6).Value Then auxiliar = "{sclien"
        auxiliar = auxiliar & ".codclien}"
        miSQL = "pDHCliente="""
        If Not PonerDesdeHasta(auxiliar, "CLI", 10, 11, miSQL) Then Exit Sub
    End If
    
    If Me.optVarios(5).Value Then
        auxiliar = "scrmacciones.tipo IN (" & OtrosDatos & ")"
        If Not AnyadirAFormula(cadSelect, auxiliar) Then Exit Sub
        auxiliar = "{scrmacciones.tipo} IN [" & OtrosDatos & "]"
        If Not AnyadirAFormula(cadFormula, auxiliar) Then Exit Sub
        
    End If
   
    If Me.optVarios(6).Value Then
        Screen.MousePointer = vbHourglass
        Me.lblIndicador(3).Caption = "Leyendo acciones"
        Me.lblIndicador(3).Refresh
        miSQL = Replace(cadSelect, "{", "")
        miSQL = Replace(miSQL, "}", "")
        If miSQL <> "" Then miSQL = " AND " & miSQL
        miSQL = " SELECT " & vUsu.Codigo & ",codclien FROM sclien WHERE 1=1 " & miSQL
        miSQL = miSQL & " AND not codclien IN ( SELECT codclien FROM scrmacciones WHERE tipo IN  (" & OtrosDatos & ")"
        If AuxF <> "" Then
            AuxF = Replace(AuxF, "{", "")
            AuxF = Replace(AuxF, "}", "")
            miSQL = miSQL & " AND " & AuxF
        End If
        miSQL = miSQL & ")"
        miSQL = "INSERT INTO tmpcrmclien(codusu,codclien) " & miSQL
        conn.Execute miSQL
    
        
    
        Me.lblIndicador(3).Caption = ""
    End If
    Me.lblIndicador(3).Caption = "Registros BD"
    Me.lblIndicador(3).Refresh
    Screen.MousePointer = vbHourglass
    b = True
    If Me.optVarios(5).Value Then
        If cadSelect <> "" Then cadSelect = " AND " & cadSelect
        cadSelect = "{scrmacciones.codclien} = {sclien.codclien} " & cadSelect
        If Not HayRegParaInforme("scrmacciones,sclien", cadSelect, True) Then b = False
        
        If Me.optVarios(7).Value Then
            cadNomRPT = "rAccionesComercVisitadosFec.rpt"
        Else
            cadNomRPT = "rAccionesComercVisitados.rpt"
        End If
        
        'Coultar fecha
        cadTitulo = "0"
        If chkVarios(6).visible And Me.chkVarios(6).Value = 1 Then cadTitulo = "1"
        
        cadTitulo = "OcultarFecha=" & cadTitulo & "|"
        cadParam = cadParam & cadTitulo
        numParam = numParam + 1
        
        cadTitulo = "Visitados"
    Else
        cadSelect = "codusu = " & vUsu.Codigo
        cadFormula = "{tmpcrmclien.codusu} = " & vUsu.Codigo
        If Not HayRegParaInforme("tmpcrmclien", cadSelect, True) Then b = False
        cadTitulo = "No visitados"
        cadNomRPT = "rAccionesComercNOVisi.rpt"
   End If
    Me.lblIndicador(3).Caption = ""
    If Not b Then
        Screen.MousePointer = vbDefault
        MsgBox "No existe datos con los valores selccionados", vbExclamation
        Exit Sub
    End If
    
    
    cadPDFrpt = ""
    conSubRPT = False
    
    LlamarImprimir
End Sub

Private Sub cmdDeclaraAlcohol_Click()
Dim b As Boolean
            
    miSQL = ""
    If Me.txtNumero(4).Text = "" Or Me.cboTrimiestre(0).ListIndex < 0 Then
        miSQL = "Indique el periodo"
    Else
        If Me.chkVarios(3).Value = 1 Then
            'Ha marcado declaracion definitiva. Veremos si es la que le corresponde
            cadParam = Format(Me.txtNumero(4).Text, "000") & Format(Me.cboTrimiestre(0).ListIndex + 1, "00")
            If txtNoModificable(5).Tag <> cadParam Then miSQL = "Trimestre a liquidar: " & Mid(txtNoModificable(5).Tag, 5, 2) & " - " & Mid(txtNoModificable(5).Tag, 1, 4)
        End If
    End If
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    b = GeneraDatosDeclaraAlcohol
    Screen.MousePointer = vbDefault
    
    If b Then
        InicializarVbles False
        
        If Me.chkVarios(4).Value = 1 Then
            GenerarFicheroAlcohol
        Else
            cadTitulo = "Listado alcohol AEAT"
        
            cadNomRPT = "fonAeatAlcohol.rpt"
            conSubRPT = False
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            LlamarImprimir
        End If
        If Me.chkVarios(3).Value = 1 Then
            If MsgBox("Marcar declaracion?", vbQuestion + vbYesNo) = vbYes Then
                ejecutar auxiliar, False
                Unload Me
            End If
        End If
    End If
    auxiliar = ""
End Sub

Private Sub cmdDevolucion_Click()
    CadenaDesdeOtroForm = IIf(Me.optDevol(0).Value, "ALV", "ART")
    Unload Me
End Sub

Private Sub cmdElimPresu_Click()
    miSQL = "1=1"
    If Me.txtFecha(0).Text <> "" Then miSQL = miSQL & " AND scafac.fecfactu >=" & DBSet(Me.txtFecha(0).Text, "F")
    If Me.txtFecha(1).Text <> "" Then miSQL = miSQL & " AND scafac.fecfactu <=" & DBSet(Me.txtFecha(1).Text, "F")
    
    cadFormula = DevuelveDesdeBD(conAri, "count(*)", "scafac", miSQL & " AND codtipom ", "FAZ", "T")
    If cadFormula = "" Then cadFormula = "0"
    If Val(cadFormula) = 0 Then
        MsgBox "Ningun dato a eliminar", vbExclamation
    Else
        CadenaDesdeOtroForm = miSQL
        frmVarios3.Opcion = 5
        frmVarios3.Show vbModal
        Unload Me
        
    End If
End Sub



Private Sub cmdEstablecerAlbaranPrincipal_Click(Index As Integer)
Dim DesmarcarSolo As Boolean
    miSQL = ""
    cadFormula = ""
    If Index = 0 Then
    
        If treeAlb(0).SelectedItem Is Nothing Then Exit Sub
            
        
        If Not treeAlb(0).SelectedItem.Parent Is Nothing Then
            MsgBox "Seelccione la cabecera del albaran", vbExclamation
            Exit Sub
        End If
        
        
        If treeAlb(0).SelectedItem.Bold Then
            'El albaran que quiere poner es el mismo que estaba
            miSQL = "Desmarcar como albaran principal"
            DesmarcarSolo = True
        Else
            miSQL = "¿Desea establecer el albaran " & vbCrLf & treeAlb(0).SelectedItem.Text & vbCrLf & " como el principal del proyecto?"
            DesmarcarSolo = False
        End If
        If MsgBox(miSQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        'Preguntamos
        'Desmarcar el que habia
        For NumRegElim = 1 To treeAlb(0).Nodes.Count
            If treeAlb(0).Nodes(NumRegElim).Bold Then
                treeAlb(0).Nodes(NumRegElim).Bold = False
                treeAlb(0).Nodes(NumRegElim).BackColor = vbWhite
            End If
        Next
        
        If Not DesmarcarSolo Then
            treeAlb(0).SelectedItem.Bold = True
            treeAlb(0).SelectedItem.BackColor = vbGreen
            treeAlb(0).SelectedItem.Checked = True
            
        End If
    Else
        
        If Not treeAlb(1).SelectedItem.Parent Is Nothing Then
            MsgBox "Seelccione la cabecera del albaran", vbExclamation
            Exit Sub
        End If
                
        
        
        For NumRegElim = 1 To treeAlb(1).Nodes.Count
            If treeAlb(1).Nodes(NumRegElim).Bold Then
                'Estaba este
                treeAlb(1).Nodes(NumRegElim).Bold = False
                treeAlb(1).Nodes(NumRegElim).BackColor = vbWhite
            End If
        
        Next
        
        
        If treeAlb(1).SelectedItem Is Nothing Then
            MsgBox "Selecccione el albaran a establecer como principal", vbExclamation
        Else
            treeAlb(1).SelectedItem.Bold = True
            treeAlb(1).SelectedItem.Checked = True
            treeAlb(1).SelectedItem.BackColor = vbGreen
        End If
    End If
    
    
    
    
    
End Sub

Private Sub cmdFitoCampos_Click()

        
    
    
    InicializarVbles True
    Screen.MousePointer = vbHourglass
    If GenerarFitoCampos Then
        vMostrarTree = True
        cadTitulo = "Campos- Fitosantiarios"
        
        cadNomRPT = "rCamposFito"
        If optFitoCampos(1).Value Then
            cadNomRPT = cadNomRPT & "cli"
            cadTitulo = "Cliente - " & cadTitulo
        End If
        cadNomRPT = cadNomRPT & ".rpt"
        cadPDFrpt = cadNomRPT
        conSubRPT = False
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        LlamarImprimir
    
    End If
    Me.lblIndicador(2).Caption = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdImpFraCoarval_Click()

    If lw(7).Tag = 0 Then
        ImportacionCoarval
    Else
        If MsgBox("¿Realizar la integracion?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Set miRsAux = New ADODB.Recordset
        conn.BeginTrans
        If GeneraFraCli Then
            conn.CommitTrans
            MsgBox "Integracion finalizada con exito", vbInformation
            CargaColumnasCoarval True
        Else
            conn.RollbackTrans
        End If
        Set miRsAux = Nothing
    End If
End Sub

Private Sub ImportacionCoarval()
Dim Rc As Byte

    
    
    
    
    
    Rc = AbrirFicheroYProcesarCoarval
    
    If Rc = 2 Then Exit Sub
    

    
    Set miRsAux = New ADODB.Recordset
    If Rc = 1 Then
        'Hay errores
        CargaColumnasCoarval True
        
        miRsAux.Open "SELECT codclien,auxiliar FROM tmpcrmclien WHERE codusu = " & vUsu.Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            lw(7).ListItems.Add , , Format(miRsAux!codClien, "000")
            lw(7).ListItems(NumRegElim).SubItems(1) = miRsAux!auxiliar
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
    Else
        CargaColumnasCoarval False
        cmdImpFraCoarval.Caption = "INTEGRAR"
        
        CargaFacturasOK
        
        
    End If
    Set miRsAux = Nothing
End Sub





Private Sub cmdLinPed_Click(Index As Integer)
Dim Destino As Integer
Dim Origen As Integer
Dim C As Integer



    If Me.lw(6).SelectedItem Is Nothing Then Exit Sub
    
    If Index <= 1 Then
        If lw(6).SelectedItem.Index = 1 Then Exit Sub
        
        If Index = 0 Then
            NumRegElim = 1
        Else
            NumRegElim = lw(6).SelectedItem.Index - 1
        End If
    Else
        If lw(6).SelectedItem.Index = lw(6).ListItems(lw(6).ListItems.Count).Index Then Exit Sub
        
        If Index = 3 Then
            NumRegElim = lw(6).ListItems(lw(6).ListItems.Count).Index
        Else
            NumRegElim = lw(6).SelectedItem.Index + 1
        End If
        
        
        
        
    End If
    
    Destino = CInt(NumRegElim)
    Origen = lw(6).SelectedItem.Index
    

    
    CambiarITem Origen, Destino
    numParam = 1
End Sub

Private Sub CambiarITem(Origen As Integer, Destino As Integer)
Dim i As Integer
Dim C As String
Dim J As Integer
Dim GuardaOrigen As String
    C = ""
    If Abs(Origen - Destino) > 1 Then
        GuardaOrigen = lw(6).ListItems(Origen).Text & "|"
        For i = 1 To 7
            GuardaOrigen = GuardaOrigen & lw(6).ListItems(Origen).SubItems(i) & "|"
        Next
        
        If Origen > Destino Then
            For J = Origen - 1 To Destino Step -1
                For i = 1 To 7
                    lw(6).ListItems(J + 1).SubItems(i) = lw(6).ListItems(J).SubItems(i)
                Next
                
                lw(6).ListItems(J + 1).Text = lw(6).ListItems(J).Text
                
            Next
        
        Else
            For J = Origen To Destino - 1
                For i = 1 To 7
                    lw(6).ListItems(J).SubItems(i) = lw(6).ListItems(J + 1).SubItems(i)
                Next
                
                lw(6).ListItems(J).Text = lw(6).ListItems(J + 1).Text
                
            Next
        
        
        End If
        
        
        
        'Reestablecemos el nodo origen
        For i = 1 To 7
            lw(6).ListItems(Destino).SubItems(i) = RecuperaValor(GuardaOrigen, i + 1)
        Next
        
        lw(6).ListItems(Destino).Text = RecuperaValor(GuardaOrigen, 1)
        
    Else
        
        For i = 1 To 7
            C = lw(6).ListItems(Destino).SubItems(i)
            lw(6).ListItems(Destino).SubItems(i) = lw(6).ListItems(Origen).SubItems(i)
            lw(6).ListItems(Origen).SubItems(i) = C
        Next
        
            C = lw(6).ListItems(Destino).Text
            lw(6).ListItems(Destino).Text = lw(6).ListItems(Origen).Text
            lw(6).ListItems(Origen).Text = C
        
    
    End If
    lw(6).ListItems(Destino).Selected = True
    Set lw(6).SelectedItem = lw(6).ListItems(Destino)
    PonerFocoOBj lw(6)
End Sub



Private Sub cmdListaComparaDto_Click()
    

    
    
    InicializarVbles True
    If GeneraDatosDtoComparativo Then
    
    
        cadTitulo = "Compara descuentos venta-compra"
      '  cadParam = cadParam & "pdh1=""Fechas: " & txtFecha(4).Text & " - " & txtFecha(5).Text & """|"
      '  numParam = numParam + 1
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
       ' cadParam = cadParam & "Detalle=" & Abs(Me.chkVarios(2).Value) & "|"
       ' numParam = numParam + 1
        vMostrarTree = True
        cadNomRPT = "rDtosVtaComrpa.rpt"   'cadPDFrpt & ".rpt"
        cadPDFrpt = ""
        conSubRPT = False
        
        LlamarImprimir
        
    
    End If
    lblIndicador(7).Caption = ""
End Sub

Private Sub cmdListadoAlb_Click()
        
        
        
    If chkAlb(0).Value = 0 And chkAlb(1).Value = 0 Then
        MsgBox "Seleccione algun tipo de albaran", vbExclamation
        Exit Sub
    End If
    
    
    
    InicializarVbles True
    
    
    If Me.cboDestinoB.ListIndex <> 2 Then
        cadSelect = IIf(cboDestinoB.ListIndex = 0, "<>", "=")
        cadSelect = "(scaalb.codtipom)" & cadSelect & "'ALZ'"
        If cboDestinoB.ListIndex = 0 Then
            auxiliar = "Albaranes."
        Else
            auxiliar = "  *Alb* "
        End If
    Else
        auxiliar = "  **Ambos** "
    End If
    
    miSQL = ""
    If chkAlb(0).Value = 1 And chkAlb(1).Value = 1 Then
        miSQL = "TODO"
    Else
        If chkAlb(0).Value = 1 Then
            miSQL = "Pendiente facurar"
        Else
            miSQL = "Facturados"
        End If
    End If
    cadPDFrpt = Trim(miSQL & "  " & auxiliar)
    miSQL = ""
    If txtFecha(13).Text <> "" Or txtFecha(14).Text <> "" Then
        miSQL = " Fecha: "
        If Not PonerDesdeHasta("{scaalb.fechaalb}", "F", 13, 14, miSQL) Then Exit Sub
        cadPDFrpt = cadPDFrpt & miSQL
    End If

    If txtCliente(8).Text <> "" Or txtCliente(9).Text <> "" Then
        miSQL = " Cliente: "
        If Not PonerDesdeHasta("{scaalb.codclien}", "CLI", 8, 9, miSQL) Then Exit Sub
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "      "
        cadPDFrpt = cadPDFrpt & miSQL
        
    End If
        
    



    cadParam = cadParam & "dh=""" & cadPDFrpt & """|"
    numParam = numParam + 1
    
    
    Screen.MousePointer = vbHourglass
    conSubRPT = ListadoFacturasAlbaranes
    Screen.MousePointer = vbDefault
    
    If Not conSubRPT Then Exit Sub
    cadNomRPT = "rListadoAlbarl.rpt"
    cadPDFrpt = cadNomRPT
    conSubRPT = False
    cadTitulo = "Listado albaranes"
    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
    LlamarImprimir
    
    
End Sub

Private Sub cmdListadoAlbInt_Click()
    InicializarVbles True
    
    cadNomRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "82", "N")
    
    cadSelect = "(scaalb.codtipom)='ALI'"
    cadFormula = "{scaalb.codtipom}='ALI'"
    
    If txtFecha(8).Text <> "" Or txtFecha(9).Text <> "" Then
        miSQL = " Fecha: "
        If Not PonerDesdeHasta("{scaalb.fechaalb}", "F", 8, 9, miSQL) Then Exit Sub
        cadPDFrpt = cadPDFrpt & miSQL
    End If

    If txtCliente(4).Text <> "" Or txtCliente(5).Text <> "" Then
        miSQL = " Cliente: "
        If Not PonerDesdeHasta("{scaalb.codclien}", "CLI", 4, 5, miSQL) Then Exit Sub
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "      "
        cadPDFrpt = cadPDFrpt & miSQL
        

    End If
    

    cadParam = cadParam & "dh=""" & cadPDFrpt & """|"
    numParam = numParam + 1
    
    
    Screen.MousePointer = vbHourglass
    If HayRegParaInforme("scaalb", cadSelect, False) Then
        
    
        cadTitulo = "Listado albaranes internos"
        
        cadPDFrpt = ""
        conSubRPT = False
        
        LlamarImprimir

    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdListPedDia_Click()
    
    Screen.MousePointer = vbHourglass
    If ListadoPedidoPorDia Then
        
        InicializarVbles True
    
        'cadNomRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "82", "N")
    
        
    
        If txtFecha(20).Text <> "" Or txtFecha(19).Text <> "" Then
            miSQL = " Fecha pedido: "
            If Not PonerDesdeHasta("F", "F", 19, 20, miSQL) Then Exit Sub
            cadPDFrpt = cadPDFrpt & miSQL
        End If

        cadParam = cadParam & "DesdeHasta=""" & cadPDFrpt & """|"
        numParam = numParam + 1
    
    
        cadTitulo = "Listado pedidos por dia"
        cadPDFrpt = ""
        conSubRPT = False
        cadNomRPT = "rPedxDia.rpt"
        cadFormula = "{tmpinformes.codusu}= " & vUsu.Codigo
        LlamarImprimir
    
        
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOrdenarLineas_Click()
    If numParam = 0 Then
        MsgBox "Ningun cambio realizado", vbExclamation
        Exit Sub
    End If
    
    
    If MsgBox("¿Desea actualizar los cambios realizados?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    conn.BeginTrans
    If ReordenarLineas Then
        conn.CommitTrans
        Unload Me
    Else
        conn.RollbackTrans
    End If
    
End Sub

Private Sub cmdPedCli_A_prov_Click()
Dim J As Integer
Dim vTipoMov As CTiposMov
Dim Import As Currency


    On Error GoTo ecmdPedCli_A_prov_Click
    
    miSQL = PonerTrabajadorConectado(cadNomRPT)
    If miSQL = "" Then Err.Raise 513, , "No se puede establecer el trabajador conectado"
    cadPDFrpt = miSQL

    miSQL = ""
    cadSelect = "|"
    cadFormula = ""
    numParam = 0
    cadTitulo = ""
    Set Colec = New Collection
    For NumRegElim = 1 To lw(11).ListItems.Count
            If lw(11).ListItems(NumRegElim).Checked Then
                cadTitulo = cadTitulo & "X"
                cadParam = "|" & lw(11).ListItems(NumRegElim).SubItems(7) & "|"
                If InStr(1, cadSelect, cadParam) = 0 Then

                    cadSelect = cadSelect & lw(11).ListItems(NumRegElim).SubItems(7) & "|"
                    numParam = numParam + 1
                End If
            End If
    Next
    
    
    If numParam = 0 Then
        MsgBox "Seleccione alguna linea para crear el pedido", vbExclamation
        Exit Sub
    End If
    
    'Lo primero va a ser asignar precio coste
    miSQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
    cadSelect = Mid(cadSelect, 2)
    For J = 1 To numParam
        cadFormula = RecuperaValor(cadSelect, J)
        miSQL = ""
        For NumRegElim = 1 To lw(11).ListItems.Count
            If lw(11).ListItems(NumRegElim).Checked Then
                
                If lw(11).ListItems(NumRegElim).SubItems(7) = cadFormula Then miSQL = miSQL & ", " & Mid(lw(11).ListItems(NumRegElim).Key, 2)
                    
            End If
        Next
        If miSQL = "" Then Err.Raise 513, , "Lineas para el proveedor: " & cadFormula
        miSQL = Mid(cadFormula & Space(12), 1, 12) & miSQL   '12 ara codprove   resto lineas pedido
        Colec.Add CStr(miSQL)
    Next
    
    miSQL = "Se va a generar : " & vbCrLf & "Pedidos de proveedor: " & Colec.Count & vbCrLf & "Lineas totales a crear:  " & Len(cadTitulo)
    If chkVarios(7).Value Then miSQL = miSQL & vbCrLf & vbCrLf & " ***** Se copiaran precios desde el pedido de venta  ******"
        
    miSQL = miSQL & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(miSQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    
    
   
    '                               line    codprove nompr  codart  nomart   canti      vtaneto
    'INSERT INTO tmpinformes codusu,campo1, codigo1,nombre1,nombre2,nombre3, importe1,importeb1
    '         compra    dto1    dto2       cost calcu     marge
    '       importeb2   porcen1  porcen2 , importebd4,importeb5
    miSQL = ""
    For NumRegElim = 1 To lw(11).ListItems.Count
        If lw(11).ListItems(NumRegElim).Checked Then
            cadParam = Mid(lw(11).ListItems(NumRegElim).Key, 2)
            miSQL = miSQL & ", (" & vUsu.Codigo & "," & cadParam & "," & lw(11).ListItems(NumRegElim).SubItems(7) & "," & DBSet(lw(11).ListItems(NumRegElim).SubItems(8), "T")
            miSQL = miSQL & "," & DBSet(lw(11).ListItems(NumRegElim).Text, "T") & "," & DBSet(lw(11).ListItems(NumRegElim).SubItems(1), "T")
            miSQL = miSQL & "," & DBSet(lw(11).ListItems(NumRegElim).SubItems(2), "N")
            Import = ImporteFormateado(lw(11).ListItems(NumRegElim).SubItems(2))  'Cantidad
            If Import = 0 Then
                miSQL = miSQL & ",0,0,0,0"
            Else
                'Importel / cantidad
                Import = Round(ImporteFormateado(lw(11).ListItems(NumRegElim).SubItems(6)) / Import, 4)
                miSQL = miSQL & "," & DBSet(Import, "N")
                miSQL = miSQL & "," & DBSet(lw(11).ListItems(NumRegElim).SubItems(3), "N")
                miSQL = miSQL & "," & DBSet(lw(11).ListItems(NumRegElim).SubItems(4), "N")
                miSQL = miSQL & "," & DBSet(lw(11).ListItems(NumRegElim).SubItems(5), "N")
                
            End If
            miSQL = miSQL & ",0,0,0,0,0,0)"
        End If
    Next
    miSQL = Mid(miSQL, 2)
    miSQL = "INSERT INTO tmpinformes (codusu,campo1, codigo1,nombre1,nombre2,nombre3, importe1,importe2,importeb3,importe4,importe5,importeb1,importeb2  ,porcen1  ,porcen2,importeb4,importeb5) VALUES " & miSQL
    conn.Execute miSQL
    
    CadenaDesdeOtroForm = ""
    frmFacCosteArtVar.Show vbModal
    If CadenaDesdeOtroForm = "" Then Exit Sub
    
    
    
    
    
    
    
    Set miRsAux = New ADODB.Recordset
    
    cadParam = DevuelveDesdeBD(conAri, "codclien", "scaped", "numpedcl", OtrosDatos, "N") 'El cliente
    If cadParam = "" Then Err.Raise 513, , "Error obteniendo pedido: " & OtrosDatos
    For J = 1 To Colec.Count
        cadSelect = Colec.Item(J)
        
        miSQL = Trim(Mid(cadSelect, 13))
        miSQL = Mid(miSQL, 2)
        cadSelect = Trim(Mid(cadSelect, 1, 12))
        
        
        
        
        Set vTipoMov = New CTiposMov
        If Not vTipoMov.Leer("PEC") Then Err.Raise 513, , "Error obteniendo contadores pedido proveedor"
         vTipoMov.ConseguirContador vTipoMov.TipoMovimiento
        
        'Cabecera
        cadFormula = "INSERT INTO scappr(numpedpr,fecpedpr,codprove,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,codtraba,"
        cadFormula = cadFormula & "codtrab1,dtognral,dtoppago,codclien,observa1,observa2,observa3,NReferencia,coddirre,coddirea ,coddiref ,codforpa ) "
        cadFormula = cadFormula & " SELECT " & vTipoMov.Contador + 1 & ",now(),codprove,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprov1"
        cadFormula = cadFormula & " , " & cadPDFrpt & " codtraba ," & cadPDFrpt & "  codtrab1, 0 dtognral, 0  dtoppago"
        cadFormula = cadFormula & ", " & cadParam & " codclien"
        
        'cadFormula = cadFormula & ", 'Generado desde pedido cliente " & OtrosDatos & "' observa1 , "
        'cadFormula = cadFormula & DBSet(cadNomRPT, "T") & "  observa2 , now() observa3,"
        
        cadFormula = cadFormula & ", null , null, null,"
        
        cadTitulo = "PED-" & Right("00000000" & OtrosDatos, 8)
        cadFormula = cadFormula & DBSet(cadTitulo, "T") & " NReferencia , NULL coddirre,NULL coddirea , NULL coddiref ,codforpa "
        cadFormula = cadFormula & " FROM sprove WHERE codprove = " & Trim(cadSelect)
        conn.Execute cadFormula
        
        cadFormula = "INSERT INTO slippr(numpedpr,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,codclien,numpedV)"
        cadFormula = cadFormula & "  SELECT " & vTipoMov.Contador + 1 & ",numlinea ,1 codalmac ,codartic,nomartic,ampliaci,cantidad,"
        
        'Como minimo el que tiene
        'If Me.chkVarios(7).Value = 1 Then
            cadFormula = cadFormula & " precioar,dtoline1,dtoline2,importel"
        'Else
        '    cadFormula = cadFormula & "0 precioar, 0 dtoline1, 0  dtoline2,0 importel"
        'End If
        cadFormula = cadFormula & ", " & cadParam & " codclien" & ", " & OtrosDatos & " nupedv"
        cadFormula = cadFormula & " FROM sliped  "
        cadFormula = cadFormula & " WHERE numpedcl = " & OtrosDatos & " AND numlinea in (" & miSQL & ")"
        conn.Execute cadFormula
        
        
        
        
        
        'Actualizamos con los precios de coste de "costes articulo"
        cadFormula = "Select campo1,importeb2,porcen1,porcen2,importeb4,importe1 from tmpinformes where codusu =  " & vUsu.Codigo
        cadFormula = cadFormula & " AND importeb4 >0 AND campo1 IN (" & miSQL & ")"
        miRsAux.Open cadFormula, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
        
            'actualizamos el precoste de pedido cliente
            cadFormula = "UPDATE sliped SET precoste =" & DBSet(miRsAux!importeb4, "N")
            cadFormula = cadFormula & " WHERE numpedcl = " & OtrosDatos
            cadFormula = cadFormula & " AND  numlinea = " & miRsAux!campo1
            conn.Execute cadFormula
            
            'ACutalizamos en pedidos proveedor
            Import = Round(miRsAux!importeb4 * miRsAux!Importe1, 2) 'precio final x cantidad
            cadFormula = "UPDATE slippr SET precioar =" & DBSet(miRsAux!importeb2, "N")
            cadFormula = cadFormula & ", dtoline1=" & DBSet(miRsAux!Porcen1, "N")
            cadFormula = cadFormula & ", dtoline2=" & DBSet(miRsAux!Porcen2, "N")
            cadFormula = cadFormula & ", importel=" & DBSet(Import, "N")
            cadFormula = cadFormula & " WHERE numpedpr = " & vTipoMov.Contador + 1
            cadFormula = cadFormula & " AND numlinea  = " & miRsAux!campo1
            conn.Execute cadFormula
            
            
            
        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        
        vTipoMov.IncrementarContador vTipoMov.TipoMovimiento
        
    Next
    
    Set miRsAux = Nothing
    Set vTipoMov = Nothing
    Screen.MousePointer = vbDefault
    
    MsgBox "Proceso generado correctamente. Fecha pedido proveedor: " & Format(Now, "dd/mm/yyyy"), vbInformation
    Unload Me
    
    Exit Sub
ecmdPedCli_A_prov_Click:
    MuestraError Err.Number, , Err.Description
    Set vTipoMov = Nothing
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrecioMinimo_Click()
    Screen.MousePointer = vbHourglass
    If HacerProcesoPrecioMinimo Then Unload Me
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdPrevFraALVIC_Click()

    
    Screen.MousePointer = vbHourglass
    If chkVerificarCtas.Value = 1 Then
    
        If HacerPrevisionCuentas Then
            cadNomRPT = "rPrevAlvicCta.rpt"
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            cadTitulo = "Prevision factura ALVIC"
            LlamarImprimir
            If MsgBox("Se van a crear las cuentas. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
                
                
            End If
        End If
    Else
        If HacerPrevisionAlvic Then
            cadNomRPT = "rPrevAlvic.rpt"
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            cadTitulo = "Prevision factura ALVIC"
            LlamarImprimir
            
            
            
            'Vamos a agrupar las formas de EFECTIVO
            ' Vienen: 0 VTA CONTADO  1: Efectivo
            ' Pasaran todas a VTAS contado (0)
            auxiliar = "Select   scaalb.codforpa,count(*) Cuantos ,nomforpa from scaalb inner join sforpa on sforpa.codforpa=scaalb.codforpa WHERE  "
            auxiliar = auxiliar & cadSelect & " and tipforpa=0 group by 1 ORDER BY codforpa "
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open auxiliar, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            auxiliar = ""
            miSQL = ""
            
            NumRegElim = 0
            Set Colec = New Collection
            While Not miRsAux.EOF
                'El cero sera la forma de pago que s
                If miRsAux!codforpa = 0 Then
                    cadFormula = miRsAux!codforpa & " - " & miRsAux!nomforpa
                Else
                    auxiliar = auxiliar & miRsAux!codforpa & " - " & miRsAux!nomforpa & " (" & miRsAux!Cuantos & ")" & vbCrLf
                    
                    'Esta hay que camiarla
                    miSQL = "UPDATE scaalb SET codforpa =0 WHERE " & cadSelect & " AND codforpa = " & miRsAux!codforpa
                    Colec.Add miSQL
                    NumRegElim = NumRegElim + 1
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If NumRegElim > 0 Then
                If cadFormula = "" Then cadFormula = "Forma de pago 0"
                miSQL = "Va a poner la forma de pago: " & cadFormula & " a:" & vbCrLf & vbCrLf & auxiliar
                If MsgBox(miSQL, vbQuestion + vbYesNoCancel) = vbYes Then
                    For NumRegElim = 1 To Colec.Count
                        miSQL = Colec.Item(NumRegElim)
                        ejecutar miSQL, False
                    Next
                End If
            End If
            Set miRsAux = Nothing
            NumRegElim = 0
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdPuntosCliente_Click()
    miSQL = ""
    If Me.txtFecha(10).Text = "" Then miSQL = miSQL & "- Fecha" & vbCrLf
    If txtNumero(3).Text = "" Then miSQL = miSQL & "- Puntos" & vbCrLf
    If txtModificable(3).Text = "" Then miSQL = miSQL & "- Observaciones" & vbCrLf
    If miSQL <> "" Then
        miSQL = "Campos requeridos: " & vbCrLf & miSQL
        MsgBox miSQL, vbExclamation
        Exit Sub
    End If
    
    If CDate(Me.txtFecha(10).Text) < vParamAplic.PtosFechaIncio Then
        MsgBox "Fehca debe ser mayor igual a " & vParamAplic.PtosFechaIncio, vbExclamation
        Exit Sub
    End If
    
    miSQL = "Va a incrementar los puntos para el cliente: " & vbCrLf & RecuperaValor(OtrosDatos, 2) & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(miSQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    conn.BeginTrans
    If HacerIncrementoPuntosCliente Then
        conn.CommitTrans
        
    
        CadenaDesdeOtroForm = "OK"
        Unload Me
    Else
        conn.RollbackTrans
    End If
End Sub

Private Sub cmdSdtofmInsert_Click()
    If Me.txtFecha(2).Text = "" Then Exit Sub
    
    If Me.txtActiv(1).Text = "" And txtFamia(0).Text = "" Then Exit Sub
    
    miSQL = "Va a actualizar los descuentos familia / marca por actividad. " & vbCrLf
    If Me.txtActiv(1).Text <> "" Then miSQL = miSQL & " ACTIVIDAD: " & txtActiv(1).Text & "-" & Me.txtDescActiv(1).Text & vbCrLf
    If txtFamia(0).Text <> "" Then miSQL = miSQL & " FAMILIA: " & txtFamia(0).Text & "-" & Me.txtDescFamia(0).Text & vbCrLf
    
    miSQL = miSQL & vbCrLf & "¿Desea continuar con el proceso?"
    If MsgBox(miSQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    'OK, vamos p'alla
    Screen.MousePointer = vbHourglass
    lblIndicador(0).Caption = "Preparando BD"
    lblIndicador(0).Refresh
    conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
    
    
    'Cargamos todos los posibles descuentos
    lblIndicador(0).Caption = "Leyendo familia actividad"
    lblIndicador(0).Refresh
    
    miSQL = "insert into tmpinformes(codusu,codigo1,campo1,importe1,fecha1)"
    miSQL = miSQL & " SELECT " & vUsu.Codigo & ", sactivdtos.codactiv,sactivdtos.codfamia,dtoline1"
    miSQL = miSQL & "," & DBSet(txtFecha(2).Text, "F")
    miSQL = miSQL & " From sactivdtos, sfamiadtos WHERE  sfamiadtos.codfamia=sactivdtos.codfamia AND"
    miSQL = miSQL & " sfamiadtos.clasifica=sactivdtos.clasifica "
    If Me.txtActiv(1).Text <> "" Then miSQL = miSQL & " AND sactivdtos.codactiv = " & Me.txtActiv(1).Text
    If txtFamia(0).Text <> "" Then miSQL = miSQL & " AND sfamiadtos.codfamia = " & txtFamia(0).Text
    conn.Execute miSQL
    
    
    
    'Pequeña comprobacion
    miSQL = "CodUsu = " & vUsu.Codigo & " and not campo1 in (select codfamia from sfamia) AND 1"
    miSQL = DevuelveDesdeBD(conAri, "campo1", "tmpinformes", CStr(miSQL), 1)
    If miSQL <> "" Then
        MsgBox "La familia " & miSQL & " NO existe", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    
    'Si ha puesto solo los nuevos veo cuales tengo que borrar de la temporal
    If Me.chkVarios(1).Value = 1 Then
        Set miRsAux = New ADODB.Recordset
        lblIndicador(0).Caption = "Comprobando valores existente"
        lblIndicador(0).Refresh
        miSQL = "Select codfamia from sdtofm where  "
        miSQL = miSQL & " codclien IS NULL AND codmarca is null "
        If Me.txtActiv(1).Text <> "" Then miSQL = miSQL & " AND codactiv =" & Me.txtActiv(1).Text
        If txtFamia(0).Text <> "" Then miSQL = miSQL & " AND codfamia = " & txtFamia(0).Text
        
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        miSQL = ""
        While Not miRsAux.EOF
            lblIndicador(0).Caption = "Fam: " & miRsAux!Codfamia
            lblIndicador(0).Refresh
            miSQL = miSQL & ", " & miRsAux!Codfamia
            If Len(miSQL) > 400 Then
                miSQL = Mid(miSQL, 2)
                miSQL = vUsu.Codigo & " AND campo1 IN (" & miSQL & ")"
                miSQL = "DELETE FROM tmpinformes WHERE codusu = " & miSQL
                conn.Execute miSQL
                miSQL = ""
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If miSQL <> "" Then
            miSQL = Mid(miSQL, 2)
            miSQL = vUsu.Codigo & " AND campo1 IN (" & miSQL & ")"
            miSQL = "DELETE FROM tmpinformes WHERE codusu = " & miSQL
            conn.Execute miSQL
        End If
    Else
    
        'Borrare SEGURO de actualizar los que tienen descuento especial a 1
        Set miRsAux = New ADODB.Recordset
        lblIndicador(0).Caption = "Comprobando descuentos especiales"
        lblIndicador(0).Refresh
        miSQL = "Select codfamia from sdtofm where  "
        miSQL = miSQL & " codclien IS NULL AND codmarca is null "
        If Me.txtActiv(1).Text <> "" Then miSQL = miSQL & " AND codactiv =" & Me.txtActiv(1).Text
        If txtFamia(0).Text <> "" Then miSQL = miSQL & " AND codfamia = " & txtFamia(0).Text
        miSQL = miSQL & " AND dtoesp= 1"
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        miSQL = ""
        While Not miRsAux.EOF
            lblIndicador(0).Caption = "Esp. Familia: " & miRsAux!Codfamia
            lblIndicador(0).Refresh
            miSQL = miSQL & ", " & miRsAux!Codfamia
            If Len(miSQL) > 400 Then
                miSQL = Mid(miSQL, 2)
                miSQL = vUsu.Codigo & " AND campo1 IN (" & miSQL & ")"
                miSQL = "DELETE FROM tmpinformes WHERE codusu = " & miSQL
                conn.Execute miSQL
                miSQL = ""
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If miSQL <> "" Then
            miSQL = Mid(miSQL, 2)
            miSQL = vUsu.Codigo & " AND campo1 IN (" & miSQL & ")"
            miSQL = "DELETE FROM tmpinformes WHERE codusu = " & miSQL
            conn.Execute miSQL
        End If
    
    
        'QUIERE METERLOS TODOS
        'Borro de sdtofm con codactiv  e inserto desde tmpinformes
        lblIndicador(0).Caption = "Eliminando registros anteriores"
        lblIndicador(0).Refresh
        miSQL = "DELETE from sdtofm WHERE codclien is null and codmarca is null"
        If Me.txtActiv(1).Text <> "" Then miSQL = miSQL & " AND codactiv =" & txtActiv(1).Text

        If txtFamia(0).Text <> "" Then miSQL = miSQL & " AND codfamia = " & txtFamia(0).Text
        miSQL = miSQL & " AND dtoesp = 0"
        conn.Execute miSQL
    End If
    
    'INSERTAMOS desde tmpinformes
    lblIndicador(0).Caption = "Insertando en descuentos"
    lblIndicador(0).Refresh
    miSQL = "INSERT INTO sdtofm (codclien,codfamia,codmarca,fechadto,dtoline1,dtoline2,codactiv,dtoEsp) SELECT"
    miSQL = miSQL & " null,campo1,null,fecha1,importe1,0,codigo1,0 FROM tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
    
    
    Unload Me
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub cmdSelecAlbaran_Click()
    If lw(3).ListItems.Count = 0 Then Exit Sub
    If lw(3).SelectedItem Is Nothing Then Exit Sub
    
    With lw(3).SelectedItem
        CadenaDesdeOtroForm = chkPedidos.Value & .Text & "|" & .SubItems(1) & "|" & .SubItems(2) & "|"
    End With
    Unload Me
End Sub

Private Sub cmdSeleccionarFamilia_Click()
    
    
    miSQL = ","
    For NumRegElim = 1 To lw(0).ListItems.Count
        If lw(0).ListItems(NumRegElim).Checked Then miSQL = miSQL & lw(0).ListItems(NumRegElim).Text & ","
    Next
    If miSQL = "," Then
        MsgBox "Seleccione alguna familia", vbExclamation
    Else
        CadenaDesdeOtroForm = miSQL
        Unload Me
    End If
End Sub

Private Sub cmdSelProvee_Click()
    
    miSQL = ""
    For NumRegElim = 1 To lw(1).ListItems.Count
        If lw(1).ListItems(NumRegElim).Checked Then miSQL = miSQL & "," & lw(1).ListItems(NumRegElim).Text
    Next
    If miSQL = "" Then
        MsgBox "Seleccione alguna proveedor", vbExclamation
    Else
        CadenaDesdeOtroForm = miSQL
        
        
        'Proceso la cadena para saber guardar cuales he quiado de los que hay
        LeerGuardarSeleccionProveedoresAQuitar False
        
        Unload Me
    End If
End Sub

Private Sub cmdSubRPT_Click()
    If Trim(txtModificable(1).Text) = "" Or Trim(txtModificable(2).Text) = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = txtNoModificable(2).Text & "|" & txtModificable(1).Text & "|" & txtModificable(2).Text & "|"
    
    
    Unload Me
End Sub


Private Sub cmdTaxcoCambioCliente_Click()
    miSQL = ""
    If lw(13).ColumnHeaders.Count = 0 Then miSQL = "1"
    If txtCliente(14).Text = "" Then miSQL = "1"
    If Me.txtDescClie(14).Text = "" Then miSQL = "1"
    If miSQL <> "" Then Exit Sub
    
    cadParam = ""
    miSQL = ""
    For NumRegElim = 1 To lw(13).ListItems.Count
        
        If lw(13).ListItems(NumRegElim).Tag = 1 Then
            miSQL = miSQL & "X"
        Else
            If lw(13).ListItems(NumRegElim).Checked Then cadParam = cadParam & "X"
        End If
            
    Next
    If Len(cadParam) = 0 Then
        MsgBox "Seleccione algun dato", vbExclamation
        Exit Sub
    End If
    
    If miSQL <> "" Then
        miSQL = "Hay facturas con distinta forma de pago: " & Len(miSQL) & vbCrLf
        MsgBox miSQL, vbExclamation
    End If
    
    cadParam = "A " & Len(cadParam) & " facturas va a asignarle el cliente " & vbCrLf & vbCrLf & Me.txtCliente(14).Text & " " & Me.txtDescClie(14).Text & vbCrLf
    cadParam = cadParam & vbCrLf & "¿Continuar?"
    If MsgBox(cadParam, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    vMostrarTree = realizarCambioClienteTaxco
    Screen.MousePointer = vbDefault
    If vMostrarTree Then
        conn.CommitTrans
        Espera 0.5
        txtCliente_LostFocus 14
    Else
        conn.RollbackTrans
    End If
    
End Sub

Private Sub cmdTaxcoNuevoEntradaVehiculo_Click()
Dim vC As CCliente

    
    If lw(8).SelectedItem Is Nothing Then Exit Sub
    
    CadenaDesdeOtroForm = "OK"
    
    
    
    
    If lw(8).ListItems.Count > 1 Then
        If lw(8).ListItems(1).Tag <> lw(8).SelectedItem.Tag Then CadenaDesdeOtroForm = ""
    Else
         If lw(8).ListItems.Count = 0 Then CadenaDesdeOtroForm = ""
    End If
    
    If CadenaDesdeOtroForm <> "" Then
        cadTitulo = lw(8).SelectedItem.SubItems(2)
        Set vC = New CCliente
        If Not vC.LeerDatos(cadTitulo) Then
            CadenaDesdeOtroForm = ""
        Else
            If vC.ClienteBloqueado(0, False) Then CadenaDesdeOtroForm = ""
        End If
        Set vC = Nothing
    End If
    cadTitulo = ""
    If CadenaDesdeOtroForm = "" Then Exit Sub
    
    
    CadenaDesdeOtroForm = lw(8).SelectedItem.Tag
   
    Unload Me
    
End Sub

Private Sub cmdTraerLineaCompraCliente_Click()
    If lw(4).ListItems.Count = 0 Then Exit Sub
    If lw(4).SelectedItem Is Nothing Then Exit Sub
    If lw(4).SelectedItem.Tag = "" Then
        MsgBox "No se puedesen seleccionar cantidades negativas", vbExclamation
        Exit Sub
    End If
    
    'codtipom numfactu fecfactu codtipoa numalbar numlinea
    With lw(4).SelectedItem
    
        If Mid(.Text, 1, 1) = "F" Then
            'Es una factura
            miSQL = "codtipom = " & DBSet(.Text, "T") & " AND numfactu = " & .SubItems(1) & " AND fecfactu =" & DBSet(.SubItems(2), "F")
            miSQL = miSQL & " AND " & .Tag
            miSQL = "Select * from slifac WHERE " & miSQL
        Else
            miSQL = "Select * from slialb WHERE " & .Tag
        
        End If
    End With
    
    'Cerramos mirsaux y lo abrimos con el sql
    miRsAux.Close
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdTratamientosLis_Click()
   
    InicializarVbles True
    If txtCodigoVario(2).Text <> "" Or txtCliente(3).Text <> "" Then
        miSQL = "Tratamiento: "
        cadTitulo = "{advtrata.codtrata}"
        If Not PonerDesdeHasta(cadTitulo, "VVV", 2, 3, miSQL) Then Exit Sub
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "              "
        cadPDFrpt = cadPDFrpt & miSQL
    End If
    
    If txtFecha(21).Text <> "" Then
        miSQL = " Fecha inicio desde " & txtFecha(21).Text
        cadTitulo = " {advtrata.fechaini} >= date(" & Format(txtFecha(21).Text, "yyyy,mm,dd") & ") "
        cadFormula = cadFormula & " " & IIf(cadFormula <> "", "AND", "") & cadTitulo
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & " "
        cadPDFrpt = cadPDFrpt & miSQL
    End If
    
    If txtFecha(22).Text <> "" Then
        miSQL = " Fecha fin hasta" & txtFecha(22).Text
        cadTitulo = " {advtrata.fechafin} <= date(" & Format(txtFecha(22).Text, "yyyy,mm,dd") & ") "
        cadFormula = cadFormula & " " & IIf(cadFormula <> "", "AND", "") & cadTitulo
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & " "
        cadPDFrpt = cadPDFrpt & miSQL
    End If
    
    
    cadParam = cadParam & "DesdeHasta=""" & Trim(cadPDFrpt) & """|"
    numParam = numParam + 1
   
    cadParam = cadParam & "desglosa=" & Val(chkVarios(8).Value) & "|"
    numParam = numParam + 1
    vMostrarTree = False
    cadTitulo = "Tratamientos"
    cadNomRPT = "rTratamientos.rpt"
    cadPDFrpt = ""
    conSubRPT = True
   
   LlamarImprimir
   
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        If OpcionListado = 14 Then CargaAlbaranesFacturaClienteEuler
        If OpcionListado = 16 Then CargarFacturasVentaCliente
        If OpcionListado = 20 Then PonerFocoBtn cmdDeclaraAlcohol
        If OpcionListado = 21 Then CargaDatosReimpresion
        If OpcionListado = 26 Then cmdImpFraCoarval_Click
        If OpcionListado = 27 Or OpcionListado = 28 Or OpcionListado = 42 Then datosLineasAlbarEulerEspecial
        If OpcionListado = 29 Then PonerFoco txtModificable(6)
        If OpcionListado = 30 Then PonerImportesFormaPagoALVIC
        If OpcionListado = 32 Then CarcaLineasPedidoCliente
        If OpcionListado = 33 Then CargaPedidosConLw
        If OpcionListado = 38 Then CargaDatosPreCosteArtVario
        
        If OpcionListado = 41 Then PonerFoco txtProve(2)
        If OpcionListado = 43 Then CargaDatosAlbaranesEulerVinculacion
        
        If OpcionListado = 44 Then CargaDatosCestaUsuario
        
        If OpcionListado = 45 Then CargaFechasPropuestasCambioFechaTaxco
        
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim Indice As Byte

    Me.Icon = frmPpal.Icon
    PrimVez = True
    '++
    CargaIconosAyuda2
    
    FrameContadoresAgua.visible = False
    FrameGessocial.visible = False
    FrameEliminarPresupuestos.visible = False
    FrameDtoAsginar.visible = False
    FrameActualizaSdtoFm.visible = False
    FrameGesocialCambioSituacion.visible = False
    FrameSubRPT.visible = False
    FrameSeleccionarFamilia.visible = False
    FrameAguaMod.visible = False
    FrameComprasTratamientos.visible = False
    FrameSelProveedores.visible = False
    FrameCambioProveedorPedido.visible = False
    FrameFitoCampos.visible = False
    FrameAlbaranesInternos.visible = False
    FrameAlbaranesClientes.visible = False
    FrameMarjalChipos.visible = False
    FrameListadoComprasClientes.visible = False
    FramePreguntaDevoluciones.visible = False
    FrameTelefonia.visible = False
    FramePtosCliente.visible = False
    FrameDeclaraAlcohol.visible = False
    FrameReimpresionSignotec.visible = False
    FrameCRMClieAccion.visible = False
    FrameOrdenarLineas.visible = False
    FrameListadoAlb.visible = False
    FrameCreditoYCaucion.visible = False
    FrameCoarval.visible = False
    FrameLineasAlbaFalsoEuler.visible = False
    FrameTaxcoNuevaTaller.visible = False
    frameAlvic.visible = False
    FramePreviFacturaTaxo.visible = False
    FramePedCli_A_prov.visible = False
    FrameBusqPreviaPedFontenas.visible = False
    FrameTaxcoGasolineraCambiCli.visible = False
    FrameCambioCliente.visible = False
    FrameListPedxDia.visible = False
    FrameComparativoDtos.visible = False
    FrameCosteLin.visible = False
    FrameTrata.visible = False
    FrameACtualizaPrecioMinimo.visible = False
    FrameCopiarPrecios.visible = False
    FrameAsignarAlbaranesEuler.visible = False
    FrameTaxcoSvenciAlvic.visible = False
    
    PuedeCerrar = True
    
    Indice = OpcionListado
    Me.Caption = "Listado"
    Select Case OpcionListado
    Case 0
        PonerFrameVisible FrameContadoresAgua
        optVarios(0).Value = True
    
    Case 1
        PonerFrameVisible FrameGessocial
        optVarios(2).Value = True
        
        FrameFechaBaja.BorderStyle = 0
        
        
    Case 2
        PonerFrameVisible FrameEliminarPresupuestos
        
    Case 3
        PonerFrameVisible FrameDtoAsginar

        CargarCombo_Tabla cboTipoDescuento(0), "sfamiatipodto", "clasifica", "nombre"
        
    Case 4
        PonerFrameVisible FrameActualizaSdtoFm
        lblIndicador(0).Caption = ""
    Case 5
        PonerFrameVisible FrameGesocialCambioSituacion
        Me.txtNoModificable(0).Text = RecuperaValor(OtrosDatos, 1)
        Me.txtNoModificable(1).Text = RecuperaValor(OtrosDatos, 2)
        Me.txtModificable(0).Text = ""
        
    Case 6
        PonerFrameVisible FrameSubRPT
        '  OtrosDatos   -> NUEVO|Codigo|Descrip|RPT|
        '   si es nuevo solo importa el codigo
        conSubRPT = RecuperaValor(Me.OtrosDatos, 1) = "0" 'MODIFICAR
        txtNoModificable(2).Text = RecuperaValor(Me.OtrosDatos, 2)
        txtModificable(1).Text = RecuperaValor(Me.OtrosDatos, 3)
        txtModificable(2).Text = RecuperaValor(Me.OtrosDatos, 4)
    Case 7
        PonerFrameVisible FrameSeleccionarFamilia
        CargaFamilias
        
    Case 8
        PonerFrameVisible Me.FrameAguaMod
        Me.txtNumero(0).Text = ""
        
        'OtrosDatos:    consumo anterior|con actuaql|m3|numfact fecfact para update|
        
        txtNoModificable(3).Text = RecuperaValor(OtrosDatos, 1)
        txtNoModificable(4).Text = RecuperaValor(OtrosDatos, 2)
        Label1(7).Tag = RecuperaValor(OtrosDatos, 3)
        Label1(7).Caption = Label1(7).Tag & " m3"
        Me.Caption = "Consumo"
        
    Case 9
        PonerFrameVisible FrameComprasTratamientos
        Me.Caption = "Tratamientos"
        lblIndicador(1).Caption = ""
        
    Case 10
         
        PonerFrameVisible FrameSelProveedores
        Me.Caption = "Proveedor"
        CargaProveedores
    Case 11
        PonerFrameVisible FrameCambioProveedorPedido
        Me.Caption = "Pedido proveedor"
        CargaLineasPedidoProveedor
        
    Case 12
        PonerFrameVisible FrameFitoCampos
        miSQL = Format(DateAdd("m", -2, Now), "/mm/yyyy")
        Me.txtFecha(6).Text = "01" & miSQL
        
        lblIndicador(2).Caption = ""
    Case 13
        PonerFrameVisible FrameAlbaranesInternos
        
    Case 14
        PonerFrameVisible FrameAlbaranesClientes
        Me.chkPedidos.Value = Mid(OtrosDatos, 1, 1)
        
        OtrosDatos = Mid(OtrosDatos, 2)
    Case 15
    
        PonerFrameVisible FrameMarjalChipos
    Case 16
        PonerFrameVisible FrameListadoComprasClientes
        
    Case 17
        PonerFrameVisible FramePreguntaDevoluciones
        
    Case 18
        PonerFrameVisible FrameTelefonia
        limpiar Me
        txtNumero(2).Text = "24"
        
    Case 19
        PonerFrameVisible FramePtosCliente
        limpiar Me
        txtFecha(10).Text = Format(Now, "dd/mm/yyyy")
        lblDpto(18).Caption = RecuperaValor(OtrosDatos, 2)
        
    Case 20
        
        CargaComoboTrimestre 0
        DatosTrimestreAnterior
        
        
        PonerFrameVisible Me.FrameDeclaraAlcohol
        
        
    Case 21
        Me.Caption = "Documentos firmados"
        PonerFrameVisible Me.FrameReimpresionSignotec
        Set Me.lwSigno.SmallIcons = frmPpal.ImgListPpal
        
    Case 22
    
        FrameAccionComerOrden.BorderStyle = 0
        
        PonerFrameVisible FrameCRMClieAccion
        CargaDatosAccionesComerciales
        lblIndicador(3).Caption = "" 'indicador
        
        lblDpto(24).Caption = IIf(vParamAplic.NumeroInstalacion = vbHerbelca, "Visitador", "Agente")
        
        
        optVarios(5).Caption = "Visitados"
        optVarios(6).Caption = "NO visitados"
        
    Case 23
        
        PonerFrameVisible FrameOrdenarLineas
        
    Case 24
        PonerFrameVisible FrameListadoAlb
        txtFecha(13).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
        txtFecha(14).Text = txtFecha(13).Text
        cboDestinoB.ListIndex = 0
        
        
    Case 25
        
        PonerFrameVisible FrameCreditoYCaucion
        
        txtFecha(15).Text = "01" & Format(DateAdd("m", -1, Now), "/mm/yyyy")
        auxiliar = DiasMes(CByte(Month(DateAdd("m", -1, Now))), Year(DateAdd("m", -1, Now)))
        txtFecha(16).Text = auxiliar & Format(DateAdd("m", -1, Now), "/mm/yyyy")
        
    Case 26
        Me.Caption = "IMPORTAR"
        PonerFrameVisible FrameCoarval
        
        cmdImpFraCoarval.Caption = "Leer fich."
        lw(7).Tag = 0  'Pendiente de recibir datos
        
        
    Case 27, 28, 42
        If OpcionListado = 28 Then
            Me.Caption = "Euler albaranes"
            Indice = 27
        ElseIf OpcionListado = 42 Then
            Me.Caption = "Euler PROYECTOS"
            Indice = 27
        Else
            Me.Caption = "Euler (Facturas)"
        End If
        PonerFrameVisible FrameLineasAlbaFalsoEuler
        
    Case 29
        'TAXCO
        'Nueva entrada vehiculo
        PonerFrameVisible FrameTaxcoNuevaTaller
    Case 30
        PonerFrameVisible frameAlvic
        Label9(28).Caption = "Importes traspaso ALVIC (" & Format(CCur(OtrosDatos), FormatoImporte) & ")"
        lblDpto(42).Caption = ""
        lblDpto(42).Tag = 0
        
        
    Case 31
        'TAXCO
        
        PonerFrameVisible FramePreviFacturaTaxo
    
    
    Case 32
        PonerFrameVisible FramePedCli_A_prov
    
    
    Case 33
        PonerFrameVisible FrameBusqPreviaPedFontenas
        Set lw(12).SmallIcons = Me.imglistPed
        
    Case 34
        Me.Caption = "Ajuste cliente ALVIC"
        PonerFrameVisible FrameTaxcoGasolineraCambiCli
        
    Case 35
        
        Me.Caption = "Cambio cliente en base de datos"
        PonerFrameVisible FrameCambioCliente
        lblIndicador(5).Caption = ""
    Case 36
        Me.txtFecha(19).Text = Format(vEmpresa.FechaIni, "dd/mm/yyyy")
        PonerFrameVisible FrameListPedxDia
        lblIndicador(6).Caption = ""
    Case 37
        PonerFrameVisible FrameComparativoDtos
        lblIndicador(7).Caption = ""
        
    Case 38
        PonerFrameVisible FrameCosteLin
        PuedeCerrar = False
        
    Case 39
        PonerFrameVisible FrameTrata
        
    Case 40
        PonerFrameVisible FrameACtualizaPrecioMinimo
        optVarios(11).Value = True
    
    Case 41
        PonerFrameVisible FrameCopiarPrecios
        Me.Caption = "Proceso"
    Case 43
        PonerFrameVisible FrameAsignarAlbaranesEuler
        Me.Caption = "EULER.     Proyectos. Asignar albaranes"
        lblIndicador(8).Caption = ""
        
    Case 44
        PonerFrameVisible FrameCestaApp
        Me.Caption = "APP-Ariges.     Cestas  almacen"
        
    Case 45
        PonerFrameVisible FrameTaxcoSvenciAlvic
        Me.Caption = "ALVIC. Formas de pago"
        lblIndicador(9).Caption = ""
    End Select
    
    Me.cmdCancelar(CInt(Indice)).Cancel = True
End Sub

Private Sub CargaIconosAyuda2()
Dim i As Integer
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    
    imgActividad(0).Picture = imgBuscar(0).Picture
    imgActividad(1).Picture = imgBuscar(0).Picture
    
    imgFamilia(0).Picture = imgBuscar(0).Picture
    
    Err.Clear
End Sub



Private Sub PonerFrameVisible(ByRef F As Frame)
    F.Top = 0
    F.Left = 120
    F.visible = True
    Me.Height = F.Height + 480
    Me.Width = F.Width + 240
End Sub

Private Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadSelect = ""
    cadParam = "|"
    numParam = 0
    cadTitulo = ""
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    vMostrarTree = False
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
End Sub

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .Opcion = 3000   'VAN TODOS EN ESTE SACO
        .NombrePDF = ""
        .NombrePDF = cadPDFrpt
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .MostrarTreeDesdeFuera = vMostrarTree
        .Show vbModal
    End With
End Sub



Private Sub FrameVtaPlazosTfnoia_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not PuedeCerrar Then Cancel = 1
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    miSQL = CadenaDevuelta
End Sub

Private Sub frmB3_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    miSQL = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    
    If CadenaSeleccion <> "" Then
        If CadenaSeleccion <> "NO" Then
            CadArticulos = "(" & Mid(CadenaSeleccion, 2) & ")"
        Else
            CadArticulos = CadenaSeleccion
        End If
    End If
End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub imgActividad_Click(Index As Integer)

'    AbreBuscaGrid 1
    Set frmB3 = New frmFacActividades
    frmB3.DatosADevolverBusqueda = "0|1|"
    frmB3.DeConsulta = True
    frmB3.Show vbModal
    Set frmB3 = Nothing

    If miSQL <> "" Then
        
        txtActiv(Index).Text = RecuperaValor(miSQL, 1)
        Me.txtDescActiv(Index).Text = RecuperaValor(miSQL, 2)
       
    End If

End Sub


Private Sub AbreBuscaGrid(Cual As Byte)


    Screen.MousePointer = vbHourglass
    
    Set frmB = New frmBuscaGrid
    
    
    
    If Cual = 1 Then
        frmB.vTitulo = "Actividad"
        miSQL = "Codigo|sactiv|codactiv|N||20·"
        miSQL = miSQL & "descripcion|sactiv|nomactiv|T||45·"
        frmB.vCampos = miSQL
    
    ElseIf Cual = 2 Then
        frmB.vTitulo = "Tratamientos"
        miSQL = "Codigo|advtrata|codtrata|N||20·"
        miSQL = miSQL & "descripcion|advtrata|nomtrata|T||45·"
        frmB.vCampos = miSQL
        
    End If
    frmB.vSQL = ""
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    miSQL = ""
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault


End Sub


Private Sub imgAgente_Click(Index As Integer)

    miSQL = ""
'    Set frmAc = New frmFacAgentesCom
'    frmAc.DatosADevolverBusqueda = "0"
    If Not IsNumeric(txtAgente(Index)) Then txtAgente(Index).Text = ""
'    frmAc.Show vbModal
    Set frmAc = New frmBasico2
    AyudaAgentesComerciales frmAc, txtAgente(Index), , True
    Set frmAc = Nothing
    PonerFoco txtAgente(Index)
    If miSQL <> "" Then
        txtAgente(Index).Text = RecuperaValor(miSQL, 1)
        txtDescAge(Index).Text = RecuperaValor(miSQL, 2)
    End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim Cual As Byte
Dim Chec As Boolean


    If Index < 2 Then
        'Proveedores
        Cual = 0
        Chec = True
        If (Index Mod 2) = 0 Then Chec = False
    ElseIf Index < 4 Then
        Cual = 1
        Chec = True
        If (Index Mod 2) = 0 Then Chec = False
        
    ElseIf Index < 6 Then
        Cual = 2
        Chec = True
        If (Index Mod 2) = 0 Then Chec = False

    ElseIf Index < 8 Then
        '6 7
        Cual = 5
        Chec = Index = 6
    ElseIf Index < 10 Then
            '8 9
        Cual = 13
        Chec = (Index Mod 2) = 0
    
    ElseIf Index < 12 Then
        '10 11
        Cual = 15
        Chec = Index = 11
    
    
    'Else
'        '12 13
'        Cual = 8
'        Chec = (Index Mod 2) = 1
    
    
    
    End If
         
    For NumRegElim = 1 To lw(Cual).ListItems.Count
        lw(Cual).ListItems(NumRegElim).Checked = Chec
    Next
    'Facturas taxco cambio cliente
    If Chec And Cual = 13 Then
        For NumRegElim = 1 To lw(Cual).ListItems.Count
            If lw(Cual).ListItems(NumRegElim).Tag = 1 Then lw(Cual).ListItems(NumRegElim).Checked = False
        Next
    End If
End Sub

Private Sub imgCliente_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    miSQL = ""

            Set frmCli = New frmBasico2
            AyudaClientes frmCli, txtCliente(Index).Text
            Set frmCli = Nothing

    If miSQL <> "" Then
        Me.txtCliente(Index).Text = RecuperaValor(miSQL, 1)
        Me.txtDescClie(Index).Text = RecuperaValor(miSQL, 2)
        PonerFoco txtCliente(Index)
        If Index = 14 Then txtCliente_LostFocus Index
        If Index = 15 Then txtCliente_LostFocus Index
    End If
End Sub



Private Sub imgCodigoVario_Click(Index As Integer)
    
        
    AbreBuscaGrid 1
    
        
    If miSQL <> "" Then
        
        txtCodigoVario(Index).Text = RecuperaValor(miSQL, 1)
        Me.txtDesVario(Index).Text = RecuperaValor(miSQL, 2)
       
    End If

    
End Sub

Private Sub imgFamilia_Click(Index As Integer)
    miSQL = ""
'    Set frmMtoFamilia = New frmAlmFamiliaArticulo
'    frmMtoFamilia.DatosADevolverBusqueda = "0|1"
'    frmMtoFamilia.Show vbModal
    Set frmMtoFamilia = New frmBasico2
    AyudaFamilias frmMtoFamilia, txtFamia(Index)
    Set frmMtoFamilia = Nothing
    If miSQL <> "" Then
        txtFamia(Index).Text = RecuperaValor(miSQL, 1)
        txtDescFamia(Index).Text = RecuperaValor(miSQL, 2)
        miSQL = ""
    End If
    PonerFoco txtFamia(Index)
End Sub

Private Sub imgFecha_Click(Index As Integer)
   miSQL = ""
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(Index).Text <> "" Then
        If IsDate(txtFecha(Index).Text) Then frmC.Fecha = CDate(txtFecha(Index).Text)
   End If
   frmC.Show vbModal
   Set frmC = Nothing
   If miSQL <> "" Then txtFecha(Index).Text = Format(miSQL, "dd/mm/yyyy")
End Sub




Private Sub imgProveedor_Click(Index As Integer)
    miSQL = ""
    Set frmP = New frmBasico2
'    frmP.DatosADevolverBusqueda = "0|1"
'    frmP.Show vbModal
    AyudaProveedores frmP, txtProve(Index)
    Set frmP = Nothing
    If miSQL <> "" Then
        txtProve(Index).Text = RecuperaValor(miSQL, 1)
        Me.txtDescProve(Index).Text = RecuperaValor(miSQL, 2)
        miSQL = ""
    End If

    PonerFoco txtProve(Index)

End Sub

Private Sub lblDestinoB_Click()
    If vParamAplic.NumeroInstalacion <> vbFenollar Then Exit Sub
    
    If Me.cboDestinoB.ListCount = 1 Then
        cboDestinoB.AddItem "Presupuestos"
        cboDestinoB.AddItem "Ambos"
        cboDestinoB.ListIndex = 2
        HaMostradoCanal2_El_B = True
        
    End If
End Sub

Private Sub lw_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Index = 12 Then OrdenacionPedidosFontenas ColumnHeader.Index
    If Index = 13 Then OrdenacionTaxcoCliente ColumnHeader.Index
    
End Sub

Private Sub lw_DblClick(Index As Integer)
    Select Case Index
    Case 3
        cmdSelecAlbaran_Click
    Case 4
        cmdTraerLineaCompraCliente_Click
    
    Case 8
        cmdTaxcoNuevoEntradaVehiculo_Click
    Case 9
        CambiarImporteALvic
    Case 12
        cmdAceptarPedFontenas_Click
    End Select
End Sub

Private Sub lw_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If Index = 13 Then
        'Cambiar facturas ALVIC de cliente (dentro de un mismo NIF)
        If Item.Checked Then If Item.Tag = 1 Then Item.Checked = False
        
    End If
End Sub

Private Sub lw_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 8 Then
            cmdTaxcoNuevoEntradaVehiculo_Click
        Else
            cmdTraerLineaCompraCliente_Click
        End If
    End If
End Sub

Private Sub lwSigno_DblClick()
    If lwSigno.SelectedItem Is Nothing Then Exit Sub
    
    
    LanzaVisorMimeDocumento Me.hwnd, lwSigno.SelectedItem.Tag
    
End Sub

Private Sub optFitoCampos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub optVarios_Click(Index As Integer)
      If Index = 5 Or Index = 6 Then FrameAccionComerOrden.visible = optVarios(5).Value
      
      If Index = 7 Or Index = 8 Then chkVarios(6).visible = optVarios(7).Value
      
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
  
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
        If indD = 11 Then Subtipo = "FH"  'Acciones comerciales fechora
    Case "CLI"
        'Cliente
        Set TDes = txtCliente(indD)
        Set THas = txtCliente(indH)
        Set DesD = txtDescClie(indD)
        Set DesH = txtDescClie(indH)
        Subtipo = "N"

'
    Case "PRO"
        Set TDes = txtProve(indD)
        Set THas = txtProve(indH)
        Set DesD = txtDescProve(indD)
        Set DesH = txtDescProve(indH)
        Subtipo = "N"
        
    Case "FAM"

        Set TDes = txtFamia(indD)
        Set THas = txtFamia(indH)
        Set DesD = txtDescFamia(indD)
        Set DesH = txtDescFamia(indH)
        Subtipo = "N"
        
        
'
'    Case "ART"
'
'        Set TDes = txtArticulo(indD)
'        Set THas = txtArticulo(indH)
'        Set DesD = txtDescArticulo(indD)
'        Set DesH = txtDescArticulo(indH)
'        Subtipo = "T"
    Case "AGT"
        Set TDes = txtAgente(indD)
        Set THas = txtAgente(indH)
        Set DesD = txtDescAge(indD)
        Set DesH = txtDescAge(indH)
        Subtipo = "N"


    Case "VVV"
        Set TDes = txtCodigoVario(indD)
        Set THas = txtCodigoVario(indH)
        Set DesD = txtDesVario(indD)
        Set DesH = txtDesVario(indH)
        Subtipo = "N"
        If indD = 2 Then Subtipo = "T"
        
    End Select
    
    devuelve = CadenaDesdeHasta(TDes.Text, THas.Text, campo, Subtipo)
    If devuelve = "Error" Then
        PonerFoco TDes
        Exit Function
    End If
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






Private Sub treeAlb_NodeCheck(Index As Integer, ByVal Node As MSComctlLib.Node)
Dim N As Node
    If Node.Parent Is Nothing Then
        'Si lo marca/desmarc, hace eso para todo los nodos
        Set N = Node.Child.FirstSibling
        
    Else
        Node.Parent.Checked = Node.Checked
        Set N = Node.FirstSibling
            
    End If
    
    If Not (N Is Nothing) Then
        Do
            N.Checked = Node.Checked
            Set N = N.Next
        Loop Until N Is Nothing
    End If
    
End Sub

Private Sub txtActiv_Gotfocus(Index As Integer)
    ConseguirFoco txtActiv(Index), 3
End Sub

Private Sub txtActiv_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgActividad_Click Index
    End If
    
End Sub

Private Sub txtActiv_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpressGnral KeyAscii, 2, True
    If KeyAscii = teclaBuscar Then
        KEYBusquedaAct KeyAscii, Index 'actividad
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusquedaAct(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgActividad_Click (Indice)
End Sub

Private Sub txtActiv_LostFocus(Index As Integer)
    txtActiv(Index).Text = Trim(txtActiv(Index).Text)
    cadTitulo = ""
    miSQL = ""
    If txtActiv(Index).Text <> "" Then
        If IsNumeric(txtActiv(Index).Text) Then
            cadTitulo = DevuelveDesdeBD(conAri, "nomactiv", "sactiv", "codactiv", txtActiv(Index).Text, "N")
            
            If Index <= 1 And cadTitulo = "" Then miSQL = "No existe la actividad"
            
        Else
            miSQL = "Campo numerico"
        End If
    End If
    Me.txtDescActiv(Index).Text = cadTitulo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtActiv(Index).Text = ""
        PonerFoco txtActiv(Index)
    End If
End Sub




Private Sub txtCliente_GotFocus(Index As Integer)
    ConseguirFoco txtCliente(Index), 3
End Sub

Private Sub txtCliente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgCliente_Click Index
    End If
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCliente_LostFocus(Index As Integer)
Dim Descri As String
Dim OK As Boolean
    Descri = ""
    OK = False
    txtCliente(Index).Text = Trim(txtCliente(Index).Text)
    If txtCliente(Index).Text <> "" Then
        If Not IsNumeric(txtCliente(Index).Text) Then
            MsgBox "Campo codigo cliente debe ser numérico", vbExclamation
            PonerFoco txtCliente(Index)
        Else
            Descri = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCliente(Index).Text, "N")
            If Descri = "" Then
                MsgBox "No existe el cliente : " & txtCliente(Index).Text, vbExclamation
            Else
                OK = True
            End If
        End If
    End If
    Me.txtDescClie(Index).Text = Descri
    If Index = 14 Then
        If Not OK Then txtCliente(Index).Text = ""
        CargaAlbaranesNif_Alvic Not OK
    End If
    If Index = 15 Then
        If Not OK Then txtCliente(Index).Text = ""
        CambiReferenciaCliente
    End If
        
    
End Sub



Private Sub txtAgente_GotFocus(Index As Integer)
    ConseguirFoco txtAgente(Index), 3
End Sub

Private Sub txtAgente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgAgente_Click Index
    End If
End Sub

Private Sub txtAgente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtAgente_LostFocus(Index As Integer)
Dim Descri As String
    
    Descri = ""
    txtAgente(Index).Text = Trim(txtAgente(Index).Text)
    If txtAgente(Index).Text <> "" Then
        If Not IsNumeric(txtAgente(Index).Text) Then
            MsgBox "Campo codigo agente debe ser numérico", vbExclamation
            txtAgente(Index).Text = ""
            PonerFoco txtAgente(Index)
        Else
            Descri = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", txtAgente(Index).Text, "N")
            If Descri = "" Then
                MsgBox "No existe el agente : " & txtAgente(Index).Text, vbExclamation
            End If
        End If
    End If
    Me.txtDescAge(Index).Text = Descri
    
End Sub






Private Sub txtCodigoVario_GotFocus(Index As Integer)
    ConseguirFoco txtCodigoVario(Index), 3
End Sub

Private Sub txtCodigoVario_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then KeyCode = 0: imgCodigoVario_Click Index
    End If
End Sub

Private Sub txtCodigoVario_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCodigoVario_LostFocus(Index As Integer)
Dim Codigo  As String
Dim Error As Boolean

    Error = False
    txtCodigoVario(Index).Text = Trim(txtCodigoVario(Index).Text)
    Codigo = ""
    miSQL = ""
    
    
    If txtCodigoVario(Index).Text <> "" Then
    
        If Index < 2 Then
            'TIPO ARTICULO
    
            Codigo = DevuelveDesdeBD(conAri, "nomtipar", "stipar", "codtipar", txtCodigoVario(Index).Text, "T")
            If Codigo = "" Then miSQL = "El codigo no pertence a ningún tipo articulo"
            
        ElseIf Index < 4 Then
            Codigo = DevuelveDesdeBD(conAri, "nomtrata", "advtrata", "codtrata", txtCodigoVario(Index).Text, "T")
            If Codigo = "" Then Codigo = "NO existe  tratamiento"
        End If
    End If
    Me.txtDesVario(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Error Then
            txtCodigoVario(Index).Text = ""
            PonerFoco txtCodigoVario(Index)
        End If
    End If
End Sub




'******************************************************************************
'******************************************************************************
'
'   GESOCIAL
'



Private Sub txtFamia_GotFocus(Index As Integer)
    ConseguirFoco txtFamia(Index), 3
End Sub

Private Sub txtFamia_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgFamilia_Click Index
    End If
End Sub

Private Sub txtFamia_KeyPress(Index As Integer, KeyAscii As Integer)
'     KEYpressGnral KeyAscii, 2, True
    If KeyAscii = teclaBuscar Then
        KEYBusquedaFam KeyAscii, Index 'familia
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusquedaFam(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFamilia_Click (Indice)
End Sub


Private Sub txtFamia_LostFocus(Index As Integer)
Dim Codigo  As String
    txtFamia(Index).Text = Trim(txtFamia(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtFamia(Index).Text <> "" Then
        If IsNumeric(txtFamia(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtFamia(Index).Text, "N")
            If Codigo = "" Then
                MsgBox "El codigo no pertence a ninguna familia", vbExclamation
                If Index = 0 Then
                    txtFamia(Index).Text = ""
                    txtDescFamia(Index).Text = ""
                End If
            End If
        Else
            miSQL = "Campo numerico"
        End If
    End If
     
    Me.txtDescFamia(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        txtFamia(Index).Text = ""
        PonerFoco txtFamia(Index)
    End If
End Sub


Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpressGnral KeyAscii, 2, True
    If KeyAscii = teclaBuscar Then
        KEYFecha KeyAscii, Index
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            txtFecha(Index).Text = ""
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub



Private Sub txtModificable_KeyPress(Index As Integer, KeyAscii As Integer)

    
        If Index <> 5 Then KEYpressGnral KeyAscii, 2, True
    
    
    
    
End Sub



Private Sub CargaFamilias()
    Set miRsAux = New ADODB.Recordset
    
    miSQL = "Select codfamia,nomfamia from sfamia where codfamia in (" & OtrosDatos & ") ORDER BY codfamia"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    lw(0).ListItems.Clear
    While Not miRsAux.EOF
        lw(0).ListItems.Add , , miRsAux!Codfamia
        NumRegElim = NumRegElim + 1
        lw(0).ListItems(NumRegElim).SubItems(1) = miRsAux!nomfamia
        lw(0).ListItems(NumRegElim).Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub





Private Sub CargaProveedores()
    Set miRsAux = New ADODB.Recordset
    
    LeerGuardarSeleccionProveedoresAQuitar True
    
    
    miRsAux.Open OtrosDatos, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    miSQL = ""
    While Not miRsAux.EOF
        miSQL = miSQL & ", " & miRsAux!Codigo1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    miSQL = Mid(miSQL, 2)
    
    miSQL = "Select codprove,nomprove from sprove where codprove in (" & miSQL & ") ORDER BY codprove"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    lw(1).ListItems.Clear
    While Not miRsAux.EOF
        lw(1).ListItems.Add , , miRsAux!Codprove
        NumRegElim = NumRegElim + 1
        lw(1).ListItems(NumRegElim).SubItems(1) = miRsAux!nomprove
        cadTitulo = "|" & miRsAux!Codprove & "|"
        If InStr(1, auxiliar, cadTitulo) = 0 Then
            lw(1).ListItems(NumRegElim).Checked = True
        Else
            lw(1).ListItems(NumRegElim).Checked = False
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub

Private Sub txtModificable_LostFocus(Index As Integer)
     'Matricula TAXCo
      If Index = 6 Then LeerMatriculaTaxco
End Sub

Private Sub txtNumero_GotFocus(Index As Integer)
    
     ConseguirFoco txtNumero(Index), 3
    
End Sub

Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumero_LostFocus(Index As Integer)
Dim J As Long
Dim Impor As Currency

 txtNumero(Index).Text = Trim(txtNumero(Index).Text)
    miSQL = ""
    If txtNumero(Index).Text <> "" Then
        
        Select Case Index
        Case 0
            If Not PonerFormatoEntero(txtNumero(Index)) Then
                txtNumero(Index).Text = ""
                
            Else
                'Calulamos la diferencia
                
                J = Val(Trim(Mid(Me.txtNoModificable(3).Text, 11))) 'quitamos la fecha y queda la lectura
                J = Val(txtNumero(Index).Text) - J
                If J < 0 Then
                    MsgBox "Consumo negativo", vbExclamation
                    J = 0
                Else
                    miSQL = txtNumero(Index).Text
                    
                End If
            End If
        Case 1, 3, 5, 8, 9
        
            If Not PonerFormatoDecimal(txtNumero(Index), 3) Then txtNumero(Index).Text = ""
        Case 2
            If Not PonerFormatoEntero(txtNumero(Index)) Then txtNumero(Index).Text = ""
        
        Case 4
            
            
            If Not PonerFormatoEntero(txtNumero(Index)) Then txtNumero(Index).Text = ""
        
        
        Case 6
            If Not PonerFormatoDecimal(txtNumero(Index), 2) Then txtNumero(Index).Text = ""
        Case 7
            If Not PonerFormatoDecimal(txtNumero(Index), 4) Then txtNumero(Index).Text = ""
        
            
        End Select
    Else
        
    End If
    If Index = 0 Then
        txtNumero(0).Text = miSQL
        If J = 0 Then J = RecuperaValor(OtrosDatos, 3)
        Label1(7).Tag = J
        Label1(7).Caption = Label1(7).Tag & " m3"
        If txtNumero(0).Text = "" Then PonerFoco txtNumero(0)
        
    Else
        If Index >= 5 And Index <= 7 Then
            If Index <> 8 Then 'que no sea el importe
                If txtNumero(6).Text <> "" And txtNumero(5).Text <> "" Then
                    Impor = ImporteFormateado(txtNumero(6).Text) * ImporteFormateado(txtNumero(5).Text)
                    If txtNumero(7).Text <> "" Then Impor = Impor * ((100 - ImporteFormateado(txtNumero(7).Text)) / 100)
                    txtNumero(8).Text = Format(Round(Impor, 2), FormatoImporte)
                End If
            End If
        End If
    End If
End Sub




Private Function HacerUpdateConsumo() As Boolean
Dim Rc As ADODB.Recordset
Dim L As Integer
Dim Consumo As Integer
Dim ini As Integer
Dim fin As Integer
Dim Meses As Integer
Dim Normal As Boolean   'Versus industrial
Dim Tramo As Integer
Dim SeFactura As Boolean
    HacerUpdateConsumo = False
    On Error GoTo eHacerUpdateConsumo
    
    
    cadSelect = ""
    cadFormula = RecuperaValor(OtrosDatos, 4)  'WHERE de numfavctu fecfactu...
    Set Rc = New ADODB.Recordset
    
    
    
    
    'Las modificaciones del consumo se reflejan en
    ' Linea 1-4     CONSUMO AGUA por bloques
    ' Linea 6-9     ALCANTARILLADO
    ' Linea 20 Consumo para la impresion en report (importe=0)
    ' Linea 21   Canon cuota consumo
    
    'Ademas habra que podificar la ultima lectura de contador
    
    miSQL = RecuperaValor(OtrosDatos, 4)
    miSQL = "Select  cantidad from slifac WHERE " & miSQL
    miSQL = miSQL & " AND numlinea in (5,10,12,25)"   'lleva el periodo
    Rc.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Rc.EOF Then Err.Raise 513, , "Imposible obtener periodo(EOF)"
    ini = -1
    Normal = True 'Todo bien
    Do
        Meses = Rc!cantidad
        If ini < 0 Then
            ini = Rc!cantidad
        Else
            If ini <> Meses Then Normal = False
        End If
        Rc.MoveNext
    Loop Until Rc.EOF
    Rc.Close
    
    If Not Normal Then Err.Raise 513, , "Imposible obtener periodo(cant. distinta lineas 5,10,12,25)"
    
    
    'Si tiene cada uno de los bloques entonces actualizaremos
    'OtrosDatos:    consumo anterior|con actuaql|m3|numfact fecfact para update|
    
    'Consumo para report
    Consumo = Val(Label1(7).Tag)
    
    'Indica el consumo
    miSQL = "UPDATE slifac set cantidad = " & Consumo & ",numbultos=" & Consumo & " WHERE " & cadFormula
    miSQL = miSQL & " AND numlinea=20"
    conn.Execute miSQL
    'Canon Generalita sobre consumo
    miSQL = "Select  numlinea,nomartic,ampliaci,precioar from slifac WHERE " & cadFormula
    miSQL = miSQL & " AND numlinea =21"
    Rc.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not Rc.EOF Then
        miSQL = "UPDATE slifac set cantidad = " & Consumo & ",numbultos=" & Consumo
        miSQL = miSQL & ", importel=round(precioar * " & Consumo & ",2)"
        miSQL = miSQL & " WHERE " & cadFormula
        miSQL = miSQL & " AND numlinea=21"
        conn.Execute miSQL
    End If
    Rc.Close
    
    'Consumo por tramos
    miSQL = "Select  numlinea,nomartic,ampliaci,precioar from slifac WHERE " & cadFormula
    miSQL = miSQL & " AND numlinea between 1 and 4"
    Rc.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    L = 0
    While Not Rc.EOF
        L = L + 1
        Rc.MoveNext
    Wend
    
    
    
    cadSelect = "Leyendo tramos factura"
    If L > 0 Then
        'Si no tiene lineas es que no se factura
        Normal = L = 3   'Si no es industrial
    
        Rc.MoveFirst
        ini = 0
        
    
        'OK vamos a ver los tramos
        
        cadSelect = Rc!Ampliaci
        'Tres tramos
        L = InStr(1, cadSelect, "m3")
        cadSelect = Trim(Mid(cadSelect, 1, L - 1))
        L = InStrRev(cadSelect, " ")
        cadSelect = Trim(Mid(cadSelect, L))
        L = Val(cadSelect)
        If L = 0 Then Err.Raise 513, , "Imposible obtener tramo(I)"
                    
        If Normal Then
            fin = Meses * L
        Else
            fin = L
        End If
        Tramo = fin
        If Consumo <= fin Then
            fin = Consumo
            Consumo = 0
        Else
            Consumo = Consumo - fin
        End If
        miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
        miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
        miSQL = miSQL & " WHERE " & cadFormula
        miSQL = miSQL & " AND numlinea=1"
        conn.Execute miSQL
        


    
        'Segunda linea (para ambos Normal e industrial
        If Consumo = 0 Then
            fin = 0
            
        Else
            
            
            If Normal Then
                cadSelect = "Tramo II"
                'Solo para NORMAL vemos el segundo nivel de consumo
                Rc.MoveNext
                cadSelect = Rc!Ampliaci
                'Tres tramos
                L = InStr(1, cadSelect, "m3")
                cadSelect = Trim(Mid(cadSelect, 1, L - 1))
                L = InStrRev(cadSelect, " ")
                cadSelect = Trim(Mid(cadSelect, L))
                L = Val(cadSelect)
                If L = 0 Then Err.Raise 513, , "Imposible obtener tramo(II)"
                    
                fin = Meses * L
          
                fin = fin - Tramo
                If fin <= 0 Then Err.Raise 513, , "Obtener tramo(II). Valor 0"
                
                If Consumo >= fin Then
                    Consumo = Consumo - fin
                Else
                    fin = Consumo
                    Consumo = 0
                End If
            Else
                'Indutria. DOS tramos
                fin = Consumo   'lo que quede va al tramo 2
                Consumo = 0
            End If
                  
        End If
        
        
        
        
        
        miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
        miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
        miSQL = miSQL & " WHERE " & cadFormula
        miSQL = miSQL & " AND numlinea=2"
        conn.Execute miSQL
    
        'Tercera linea SOLO domestico
        If Normal Then
            If Consumo = 0 Then
                fin = 0
            Else
                fin = Consumo
            End If
            miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
            miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
            miSQL = miSQL & " WHERE " & cadFormula
            miSQL = miSQL & " AND numlinea=3"
            conn.Execute miSQL
        End If
    End If
    Rc.Close
    
    
    
    '****************************************************************************************************
    'Alcantarillado
    miSQL = "Select  numlinea,nomartic,ampliaci,precioar from slifac WHERE " & cadFormula
    miSQL = miSQL & " AND numlinea between 6 and 8"
    Rc.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    L = 0
    While Not Rc.EOF
        L = L + 1
        Rc.MoveNext
    Wend
    
    If L > 0 Then
        'Se le han facturado alncatarillado
    
        Rc.MoveFirst
        ini = 0
        Consumo = Val(Label1(7).Tag)
    
        'OK vamos a ver los tramos
        cadSelect = Rc!Ampliaci
        'Tres tramos
        L = InStr(1, cadSelect, "m3")
        cadSelect = Trim(Mid(cadSelect, 1, L - 1))
        L = InStrRev(cadSelect, " ")
        cadSelect = Trim(Mid(cadSelect, L))
        L = Val(cadSelect)
        If L = 0 Then Err.Raise 513, , "Imposible obtener tramo alcant(I)"
            
        fin = Meses * L
        Tramo = fin
        If Consumo <= fin Then
            fin = Consumo
            Consumo = 0
        Else
            Consumo = Consumo - fin
        End If
        miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
        miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
        miSQL = miSQL & " WHERE " & cadFormula
        miSQL = miSQL & " AND numlinea=6"
        conn.Execute miSQL
        


    
        'Segunda linea (para ambos Normal e industrial
        If Consumo = 0 Then
            fin = 0
            
        Else
            
            
           
            cadSelect = "Tramo II alcantarillado"
            
            Rc.MoveNext
            cadSelect = Rc!Ampliaci
            'Tres tramos
            L = InStr(1, cadSelect, "m3")
            cadSelect = Trim(Mid(cadSelect, 1, L - 1))
            L = InStrRev(cadSelect, " ")
            cadSelect = Trim(Mid(cadSelect, L))
            L = Val(cadSelect)
            If L = 0 Then Err.Raise 513, , "Imposible obtener tramo(II) alcantarillado"
                
            fin = Meses * L
            
            fin = fin - Tramo
            If fin <= 0 Then Err.Raise 513, , "Obtener tramo(II) Alcantarillado. Valor <=0"
            
            If Consumo >= fin Then
                Consumo = Consumo - fin
            Else
                fin = Consumo
                Consumo = 0
            End If
     
            
            
            miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
            miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
            miSQL = miSQL & " WHERE " & cadFormula
            miSQL = miSQL & " AND numlinea=7"
            conn.Execute miSQL
        
            'Tercera linea
            
            If Consumo = 0 Then
                fin = 0
            Else
                fin = Consumo
            End If
            miSQL = "UPDATE slifac set cantidad = " & fin & ",numbultos=" & fin
            miSQL = miSQL & ", importel=round(precioar * " & fin & ",2)"
            miSQL = miSQL & " WHERE " & cadFormula
            miSQL = miSQL & " AND numlinea=8"
            conn.Execute miSQL

        
        
        End If
    End If
    Rc.Close
    
    'Vemos el numero de conador
    miSQL = "Select referenc,fecfactu from scafac1 WHERE " & cadFormula
    Rc.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadPDFrpt = Rc!referenc
    cadNomRPT = Format(Rc!FecFactu, FormatoFecha)
    Rc.Close
    
    'Updates que faltan.
    'Scafac1 donde pondra lecutar actual
    Consumo = Val(Label1(7).Tag)
    miSQL = Mid(Me.txtNoModificable(4).Text, 1, 10) & " " & Me.txtNumero(0).Text
    miSQL = "UPDATE scafac1 set observa2 = '" & miSQL & "'"
    miSQL = miSQL & " WHERE " & cadFormula
    conn.Execute miSQL
    
    
    miSQL = "UPDATE aguahcolecturas set lec_actual = " & Me.txtNumero(0).Text
    miSQL = miSQL & " WHERE contador = " & DBSet(cadPDFrpt, "T") & " AND fecha_factura ='" & cadNomRPT & "'"
    conn.Execute miSQL
    
    miSQL = "UPDATE aguacontadores  set lec_anterior = " & Me.txtNumero(0).Text
    miSQL = miSQL & " WHERE contador = " & DBSet(cadPDFrpt, "T")
    conn.Execute miSQL



    Set LOG = New cLOG
    miSQL = Mid(Me.txtNoModificable(4).Text, 11) & " / " & Me.txtNumero(0).Text
    
    miSQL = "FRA MOD[CONSUMO] " & miSQL & vbCrLf & cadFormula
    miSQL = Replace(miSQL, "codtipom", "")
    miSQL = Replace(miSQL, "numfactu", "")
    miSQL = Replace(miSQL, "fecfactu", "")
    miSQL = Replace(miSQL, " and =", "")
    miSQL = Replace(miSQL, "'", " ")
    LOG.Insertar 8, vUsu, miSQL
    Set LOG = Nothing
    
    
    
    HacerUpdateConsumo = True
    
eHacerUpdateConsumo:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description & vbCrLf & cadSelect
    Set Rc = Nothing
End Function



Private Function GeneraDatosComprasTratamientos() As Boolean
Dim RN As ADODB.Recordset
Dim total As Currency

    On Error GoTo eGeneraDatosComprasTratamientos
    
    Set miRsAux = New ADODB.Recordset
    Set RN = New ADODB.Recordset
    
    Me.lblIndicador(1).Caption = "Obteniendo ventas articulo"
    Me.lblIndicador(1).Refresh
    miSQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
'''''''    miSQL = "select " & vUsu.Codigo & ", @rownum:=@rownum+1 AS rownum ,slifac.codartic,codfamia,sartic.nomartic,"
'''''''    miSQL = miSQL & " sum(cantidad*sartic.preciouc) from slifac,sartic,(SELECT @rownum:=0) r"
'''''''    miSQL = miSQL & " where slifac.codartic=sartic.codartic AND (codtipom,numfactu,fecfactu) IN"
'''''''    miSQL = miSQL & " (Select codtipom,numfactu,fecfactu from scafac1 where fecfactu between "
'''''''    miSQL = miSQL & DBSet(txtFecha(4).Text, "F") & " AND " & DBSet(txtFecha(5).Text, "F")
'''''''
'''''''    miSQL = miSQL & " and referenc like 'Parte:%') group by codartic"
'''''''    'miSQL = miSQL & " and referenc like '%') group by codartic"

'    miSQL = "select " & vUsu.Codigo & ", @rownum:=@rownum+1 AS rownum ,slifac.codartic,cliAbono,codfamia,sartic.nomartic,"
'    miSQL = miSQL & " sum(cantidad*sartic.preciouc) from scafac,scafac1,slifac,sartic,sclien  WHERE scafac1.Codtipom = "
'    miSQL = miSQL & " scafac.Codtipom And scafac1.NumFactu = scafac.NumFactu And scafac1.FecFactu = scafac.FecFactu"
'    miSQL = miSQL & " AND scafac1.Codtipom = slifac.Codtipom And scafac1.NumFactu = slifac.NumFactu And scafac1.FecFactu = slifac.FecFactu"
'    miSQL = miSQL & " and scafac1.numalbar=slifac.numalbar and scafac1.codtipoa=slifac.codtipoa and slifac.codartic=sartic.codartic"
'    miSQL = miSQL & " and scafac.fecfactu between " & DBSet(txtFecha(4).Text, "F") & " AND " & DBSet(txtFecha(5).Text, "F")
'    miSQL = miSQL & " and referenc like 'Parte:%' and scafac.codclien=sclien.codclien"
'    miSQL = miSQL & " group by slifac.codartic,cliAbono"

    
    miSQL = "select slifac.codartic,sum(cantidad) cuantos from scafac1,slifac WHERE "
    miSQL = miSQL & " scafac1.Codtipom = slifac.Codtipom And scafac1.NumFactu = slifac.NumFactu And scafac1.FecFactu = slifac.FecFactu"
    miSQL = miSQL & " and scafac1.numalbar=slifac.numalbar and scafac1.codtipoa=slifac.codtipoa "
    miSQL = miSQL & " and scafac1.fecfactu between " & DBSet(txtFecha(4).Text, "F") & " AND " & DBSet(txtFecha(5).Text, "F")
    miSQL = miSQL & " and referenc like 'Parte:%' "
    miSQL = miSQL & " group by slifac.codartic"
    miRsAux.Open miSQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Ningun dato ", vbExclamation
        GoTo eGeneraDatosComprasTratamientos
    End If
    
    
    miSQL = ""
    While Not miRsAux.EOF
        miSQL = miSQL & ", " & DBSet(miRsAux!codArtic, "T")
        miRsAux.MoveNext
    Wend
    
    miRsAux.MoveFirst
    
    Me.lblIndicador(1).Caption = "Abriendo articulos"
    Me.lblIndicador(1).Refresh

    
    miSQL = Mid(miSQL, 2)
    miSQL = "Select codartic,preciouc,codfamia,nomartic from sartic where codartic IN (" & miSQL & ")"
    RN.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    miSQL = ""
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Me.lblIndicador(1).Caption = "Art: " & miRsAux!codArtic
        Me.lblIndicador(1).Refresh
        
        
        'tmpinformes(codusu,codigo1,nombre1,campo2,campo1,nombre2,importe1) " & miSQL
        
        miSQL = miSQL & ", (" & vUsu.Codigo & "," & NumRegElim & "," & DBSet(miRsAux!codArtic, "T") & ",0,"
    
        RN.Find "codartic = " & DBSet(miRsAux!codArtic, "T"), , adSearchForward, 1
        'NO PUEDE SER EOF
        If Not IsNull(miRsAux!Cuantos) Then
            total = DBLet(RN!precioUC, "N")
            total = Round(total * miRsAux!Cuantos, 2)
        Else
            total = 0
        End If
        
        
        miSQL = miSQL & RN!Codfamia & "," & DBSet(RN!NomArtic, "T") & "," & DBSet(total, "N") & ")"
        
        
        If Len(miSQL) > 10000 Then
            miSQL = Mid(miSQL, 2)
            miSQL = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,campo2,campo1,nombre2,importe1) VALUES " & miSQL
            conn.Execute miSQL
            miSQL = ""
        End If
        
        miRsAux.MoveNext
    Wend
    
    If miSQL <> "" Then
        miSQL = Mid(miSQL, 2)
        miSQL = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,campo2,campo1,nombre2,importe1) VALUES " & miSQL
        conn.Execute miSQL
    End If
    
    Me.lblIndicador(1).Caption = "Cerrando"
    Me.lblIndicador(1).Refresh
    miRsAux.Close
    RN.Close
    Espera 0.5
    
    'miSQL = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.Codigo))
    If NumRegElim > 0 Then
        GeneraDatosComprasTratamientos = True
    Else
        MsgBox "No existen datos entre las fechas", vbExclamation
    End If
    
    
    
    
    
    
eGeneraDatosComprasTratamientos:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        GeneraDatosComprasTratamientos = False
    End If
    Set miRsAux = Nothing
    Set RN = Nothing
End Function



Private Sub GenerarApunteAjusteTratamientos()
Dim Mc As Contadores
Dim ImporteTotal As Currency

    On Error GoTo eGenerarApunteAjusteTratamientos

    ResultadoFechaContaOK = EsFechaOKConta(Now, True)
    If ResultadoFechaContaOK > 0 Then
        If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
        Exit Sub
    End If
    
    Set miRsAux = New ADODB.Recordset
    numParam = 0
    If vEmpresa.FechaFin <= CDate(Now) Then numParam = 1
    
    
    Set Mc = New Contadores
    cadPDFrpt = "Obteniendo contador asientos"
    If Mc.ConseguirContador("0", numParam = 0, True) = 1 Then Err.Raise 513, , cadPDFrpt
    

    

    miRsAux.Open "Select * from advparametros", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Ni puede ser eof ni puede no tener valor(NULL)
    cadPDFrpt = miRsAux!DiarioAjustes
    numParam = miRsAux!CoceptoAjustes
    miRsAux.Close
    
    
    'Cabecera
    '--------------------------------------
    Me.lblIndicador(1).Caption = "Apunte"
    Me.lblIndicador(1).Refresh
    
    miSQL = "Ajuste compras tratamientos(Ariges). Periodo : " & txtFecha(4).Text & " - " & txtFecha(5).Text
    miSQL = miSQL & vbCrLf & "Realizado por " & vUsu.Nombre & " el " & Now
    miSQL = "," & DBSet(miSQL, "T", "S")
    miSQL = ") VALUES (" & cadPDFrpt & ",'" & Format(Now, FormatoFecha) & "'," & Mc.Contador & miSQL
    If vParamAplic.ContabilidadNueva Then
        'feccreacion,usucreacion,desdeaplicacion
        miSQL = miSQL & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'Ariges'"
    End If
    miSQL = miSQL & ")"
    If vParamAplic.ContabilidadNueva Then
        miSQL = "INSERT INTO hcabapu(numdiari,fechaent,numasien,obsdiari,feccreacion,usucreacion,desdeaplicacion" & miSQL
    Else
        miSQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, obsdiari" & miSQL
    End If
    ConnConta.Execute miSQL
    
    cadTitulo = "linapu"
    If vParamAplic.ContabilidadNueva Then cadTitulo = "hlinapu"
    cadTitulo = "INSERT INTO " & cadTitulo & "(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr,idcontab,punteada,traspasado) VALUES "
    
    'linapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr,idcontab,punteada,traspasado)
    cadFormula = ", (" & cadPDFrpt & ",'" & Format(Now, FormatoFecha) & "'," & Mc.Contador & ","
    
    'apunte al HABER de las cuentas de la familia
    'y un unico apunte al DEBE por el total a la nueva cuenta generica de compras para tratamientos.

    'miSQL = "Select campo1,campo2,nomfamia, ctaventa,ctavtaser,ctavent1,ctavtaseralt,sum(importe1) as impor "
    miSQL = "Select campo1,nomfamia,ctacompr ,ctacomprser,sum(importe1) as impor "
    miSQL = miSQL & " from tmpinformes,sfamia where tmpinformes.campo1=sfamia.codfamia AND codusu = " & vUsu.Codigo & " group by campo1"

    
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    cadSelect = ""
    ImporteTotal = 0
    
    While Not miRsAux.EOF
        Me.lblIndicador(1).Caption = "Linea: " & NumRegElim
        Me.lblIndicador(1).Refresh
    
        numParam = 0
        If DBLet(miRsAux!Impor, "N") <> 0 Then
'            'Si utiliza cuentas o cuenta alternativa  y es distinta la de servicios de la normal
'            If miRsAux!campo2 = 1 Then
'                'ALTERNATIVA
'                If miRsAux!ctavent1 <> miRsAux!ctavtaseralt Then
'                    cadNomRPT = miRsAux!ctavtaseralt
'                    cadPDFrpt = miRsAux!ctavent1
'                    numParam = 1
'                End If
'            Else
'                'Cuentas normales, vaos, las no alternativas
'                If miRsAux!ctaventa <> miRsAux!ctavtaser Then
'                    cadNomRPT = miRsAux!ctavtaser
'                    cadPDFrpt = miRsAux!ctaventa
'                    numParam = 1
'                End If
'
'            End If
            cadNomRPT = miRsAux!ctacompr
            cadPDFrpt = DBLet(miRsAux!ctacomprser, "T")
            If cadPDFrpt <> "" Then
                If cadNomRPT <> cadPDFrpt Then numParam = 1
            End If

        End If
    
        If numParam = 1 Then
    
            NumRegElim = NumRegElim + 1
            
            ImporteTotal = ImporteTotal + miRsAux!Impor
            
            'linliapu,codmacta,numdocum,codconce,
            miSQL = NumRegElim & ",'" & cadNomRPT & "','Fam:" & Format(miRsAux!campo1, "0000") & "'," & numParam
            
            'ampconce,timporteD,timporteH,
            miSQL = miSQL & "," & DBSet(miRsAux!nomfamia, "T") & ",NULL," & DBSet(miRsAux!Impor, "N") & ","
            
            'ctacontr,idcontab,punteada,traspasado)
            miSQL = miSQL & DBSet(cadPDFrpt, "T") & ",'CONTAB',0,0)"
            miSQL = cadFormula & miSQL
            cadSelect = cadSelect & miSQL
            
            'Contra apunte
            NumRegElim = NumRegElim + 1
            'linliapu,codmacta,numdocum,codconce,
            miSQL = NumRegElim & ",'" & cadPDFrpt & "','Fam:" & Format(miRsAux!campo1, "0000") & "'," & numParam
            
            'ampconce,timporteD,timporteH,
            miSQL = miSQL & "," & DBSet(miRsAux!nomfamia, "T") & "," & DBSet(miRsAux!Impor, "N") & ",NULL,"
            
            'ctacontr,idcontab,punteada,traspasado)
            miSQL = miSQL & DBSet(cadNomRPT, "T") & ",'CONTAB',0,0)"
            miSQL = cadFormula & miSQL
            cadSelect = cadSelect & miSQL
            
            
            
            If Len(cadSelect) > 20000 Then
                miSQL = Mid(cadSelect, 2)
                miSQL = cadTitulo & miSQL
                ConnConta.Execute miSQL
                cadSelect = ""
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If cadSelect <> "" Then
            miSQL = Mid(cadSelect, 2)
            miSQL = cadTitulo & miSQL
            ConnConta.Execute miSQL
    End If
    
    
    
    
    If NumRegElim > 0 Then
        'Debe ser >0 SIEMPRE
        If vParamAplic.ContabilidadNueva Then
            miSQL = "Asiento generado."
        Else
            miSQL = "El asiento esta en la introduccion de apuntes."
        End If
        miSQL = miSQL & vbCrLf & vbCrLf & "Fecha: " & Format(Now, "dd/mm/yyyy") & vbCrLf & "Número: " & Mc.Contador
        MsgBox miSQL, vbInformation
        Me.lblIndicador(1).Caption = "Proceso finalizado"
    Else
        miSQL = "DELETE FROM cabapu where numasien=" & Mc.Contador & " and fechaent =" & DBSet(Now, "F")
        ConnConta.Execute miSQL
        
    End If
        
    Set miRsAux = Nothing
    
eGenerarApunteAjusteTratamientos:
    If Err.Number <> 0 Then
        MuestraError Err.Number
        
    End If
    Set miRsAux = Nothing
End Sub



'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'  MODIFICAR PROVEEDOR EN PEDIDO DESPUES DE UNA SIMULACION
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
Private Sub CargaLineasPedidoProveedor()
    Set miRsAux = New ADODB.Recordset
    
    miSQL = "select tmpslipreu.*,artvario from tmpslipreu,sartic where tmpslipreu.codartic=sartic.codartic and codusu = " & vUsu.Codigo & " ORDER BY numlinea"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    lw(2).ListItems.Clear
    While Not miRsAux.EOF
        lw(2).ListItems.Add , , miRsAux!codArtic
        NumRegElim = NumRegElim + 1
        lw(2).ListItems(NumRegElim).SubItems(1) = miRsAux!NomArtic
        lw(2).ListItems(NumRegElim).Checked = True
        
        
        If miRsAux!codAlmac = 1 Then
            miSQL = "Descuentos"
        ElseIf miRsAux!codAlmac = 2 Then
            miSQL = "Precio"
        Else
            If miRsAux!artvario = 1 Then
                miSQL = "Art. varios"
            Else
                miSQL = "Igual que original"
            End If
        End If
        lw(2).ListItems(NumRegElim).SubItems(6) = miSQL
        lw(2).ListItems(NumRegElim).Tag = miRsAux!numlinea
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'  FITO CAMPOS
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Private Function GenerarFitoCampos() As Boolean
Dim RN As ADODB.Recordset

    On Error GoTo eGenerarFitoCampos
    GenerarFitoCampos = False
    
    Set miRsAux = New ADODB.Recordset
    Set RN = New ADODB.Recordset
    
    Me.lblIndicador(2).Caption = "Obteniendo ventas articulo"
    Me.lblIndicador(2).Refresh
    miSQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
    'Tipo factura
    cadPDFrpt = ""
    cadTitulo = ""
    For numParam = 0 To 2
        If chkCodtipom(numParam).Value = 1 Then
            cadPDFrpt = cadPDFrpt & ", '" & Me.chkCodtipom(numParam).Tag & "'"
            cadTitulo = cadTitulo & "- " & Me.chkCodtipom(numParam).Caption
        End If
    Next numParam
    
    If cadPDFrpt <> "" Then
        cadPDFrpt = " AND slifaccampos.codtipom IN (" & Mid(cadPDFrpt, 2) & ")"
        cadTitulo = "Facturas: " & Mid(cadTitulo, 2)
    End If
    numParam = 0
    

    miSQL = " slifaccampos.codtipom = scafac1.codtipom"
    miSQL = miSQL & " and slifaccampos.numfactu = scafac1.numfactu and slifaccampos.fecfactu = scafac1.fecfactu"
    miSQL = miSQL & " and slifaccampos.codtipoa = scafac1.codtipoa and slifaccampos.numalbar = scafac1.numalbar"
    miSQL = miSQL & " and scafac.codtipom = scafac1.codtipom and scafac.numfactu = scafac1.numfactu"
    miSQL = miSQL & "  and scafac.fecfactu = scafac1.fecfactu"
    
    
    If cadPDFrpt <> "" Then miSQL = miSQL & cadPDFrpt
  

    
    If txtFecha(6).Text <> "" Or txtFecha(7).Text <> "" Then
        cadNomRPT = " Fecha: "
        If Not PonerDesdeHasta("{scafac.fecfactu}", "F", 6, 7, cadNomRPT) Then Exit Function
        If cadTitulo <> "" Then cadTitulo = cadTitulo & "              "
        cadTitulo = cadTitulo & cadNomRPT
        
        
    End If
    ' """ + chr(13) + """
    If txtCliente(2).Text <> "" Or txtCliente(3).Text <> "" Then
        cadPDFrpt = " Cliente: "
        If Not PonerDesdeHasta("{scafac.codclien}", "CLI", 2, 3, cadNomRPT) Then Exit Function
        If Len(cadTitulo) > 60 Then
            cadTitulo = cadTitulo & """ + chr(13) + """
        Else
            cadTitulo = cadTitulo & String(10, " ")
        End If
        cadTitulo = Trim(cadTitulo & cadNomRPT)
        
    End If
    cadParam = cadParam & "DesdeHasta=""" & cadTitulo & """|"
    numParam = numParam + 1
    
    
    
       
    
    cadTitulo = " FROM slifaccampos ,scafac1,scafac WHERE " & miSQL
    If cadSelect <> "" Then
        cadSelect = Replace(cadSelect, "{", "")
        cadSelect = Replace(cadSelect, "}", "")
        cadSelect = " AND " & cadSelect
    End If
    cadTitulo = cadTitulo & cadSelect
    'Me lo guarda:
    CadenaDesdeOtroForm = cadTitulo
    
    
    cadTitulo = "Select slifaccampos.*,codclien,fechaalb " & cadTitulo & " ORDER BY codclien"
     
    miRsAux.Open cadTitulo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    cadTitulo = ""
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        If NumRegElim > 65535 Then NumRegElim = 1
        
        '                   codclien  campo secuenc fra     nomvari  partida  campook cant   artfito fecfac  fecalb
        'tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,nombre3,porcen1,importe1,obser,fecha1 ,fecha2,)
        
        miSQL = ", (" & vUsu.Codigo & "," & miRsAux!codClien & "," & miRsAux!codCampo & "," & NumRegElim & ",'"
        miSQL = miSQL & Mid(miRsAux!codtipom & "   ", 1, 3) & Format(miRsAux!Numfactu, "000000") & "',"
        'La tengo guardada en las lineas del campo
        If IsNull(miRsAux!nomvarie) Or IsNull(miRsAux!nompartida) Then
            'Lo cojo de ARIAGRO
            cadNomRPT = "NULL,NULL,1"
        Else
            cadNomRPT = DBSet(miRsAux!nomvarie, "T") & "," & DBSet(miRsAux!nompartida, "T") & ",0"
        End If
        miSQL = miSQL & cadNomRPT & ",0,NULL,"
        miSQL = miSQL & DBSet(miRsAux!FecFactu, "F") & "," & DBSet(miRsAux!FechaAlb, "F") & ")"
        
        cadTitulo = cadTitulo & miSQL
        
        If Len(cadTitulo) > 30000 Then
            cadTitulo = Mid(cadTitulo, 2)
            miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,nombre3,porcen1,importe1,obser,fecha1 ,fecha2) VALUES  " & cadTitulo
            conn.Execute miSQL
            cadTitulo = ""
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    If cadTitulo <> "" Then
        cadTitulo = Mid(cadTitulo, 2)
        miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,nombre3,porcen1,importe1,obser,fecha1 ,fecha2) VALUES  " & cadTitulo
        conn.Execute miSQL
        cadTitulo = ""
    End If
     
    If NumRegElim = 0 Then
        MsgBox "Ningun dato generado", vbExclamation
        GoTo eGenerarFitoCampos
    End If
     
     
     
    'Ajustamos los campos que no tenemos NOmvarie nompartida
    miSQL = "Select distinct(campo1) from tmpinformes where codusu =" & vUsu.Codigo & " AND porcen1<>0 ORDER BY campo1"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadTitulo = ""
    miSQL = ""
    NumRegElim = 0
    While Not miRsAux.EOF
        miSQL = miSQL & ", " & miRsAux!campo1
        NumRegElim = NumRegElim + 1
        miRsAux.MoveNext
        If NumRegElim > 30 Then
            cadTitulo = cadTitulo & Mid(miSQL, 2) & "|"
            miSQL = ""
            NumRegElim = 0
        End If
    Wend
    miRsAux.Close

    If miSQL <> "" Then cadTitulo = cadTitulo & Mid(miSQL, 2) & "|"
    
    cadNomRPT = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,rcampos.codsitua"
    cadNomRPT = cadNomRPT & " from ((@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
    cadNomRPT = cadNomRPT & " inner join @#variedades on rcampos.codvarie = variedades.codvarie)"
    cadNomRPT = cadNomRPT & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
    
    cadNomRPT = Replace(cadNomRPT, "@#", vParamAplic.Ariagro & ".") & " AND rcampos.codcampo IN ("
    
    
    
    Do
        NumRegElim = InStr(1, cadTitulo, "|")
        If NumRegElim > 0 Then
            Me.lblIndicador(2).Caption = "Campos " & Len(cadTitulo)
            Me.lblIndicador(2).Refresh
    
            miSQL = Mid(cadTitulo, 1, NumRegElim - 1)
            cadTitulo = Mid(cadTitulo, NumRegElim + 1)
            
            miSQL = cadNomRPT & miSQL & ")"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            miSQL = "UPDATE tmpinformes set nombre2=@@ , nombre3=## , porcen1=0 WHERE codusu = " & vUsu.Codigo & " AND campo1 = "
            While Not miRsAux.EOF
                cadPDFrpt = Replace(miSQL, "@@", DBSet(miRsAux!nomparti, "T"))
                cadPDFrpt = Replace(cadPDFrpt, "##", DBSet(miRsAux!nomvarie, "T"))
                cadPDFrpt = cadPDFrpt & miRsAux!codCampo
                conn.Execute cadPDFrpt
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        Else
            cadTitulo = ""
        End If
       
    Loop Until cadTitulo = ""
    
    
    Me.lblIndicador(2).Caption = "Lineas fitosanitarios"
    Me.lblIndicador(2).Refresh
    
    miSQL = Replace(CadenaDesdeOtroForm, "slifaccampos", "slifac")
    miSQL = miSQL & " and numlote<>''"
    miSQL = "Select slifac.*   " & miSQL
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    cadTitulo = ""
    While Not miRsAux.EOF
        cadPDFrpt = Mid(miRsAux!codtipom & "   ", 1, 3) & Format(miRsAux!Numfactu, "000000")
        If cadTitulo <> cadPDFrpt Then
            'Vemos si hay que updatear
            If cadTitulo <> "" Then
                Me.lblIndicador(2).Caption = "Fra: " & cadTitulo
                Me.lblIndicador(2).Refresh
                miSQL = "UPDATE tmpinformes set obser = " & DBSet(cadNomRPT, "T")
                miSQL = miSQL & " WHERE codusu = " & vUsu.Codigo & " AND nombre1= '" & cadTitulo & "'"
                conn.Execute miSQL
            End If
            cadTitulo = cadPDFrpt
            cadNomRPT = ""
        End If
        
        vMostrarTree = True
        If cadNomRPT <> "" Then
            If Len(cadNomRPT) > 240 Then
            
                vMostrarTree = False
            Else
                cadNomRPT = cadNomRPT & vbCrLf
            End If
        
        End If
        If vMostrarTree Then cadNomRPT = cadNomRPT & miRsAux!NomArtic & "(" & miRsAux!cantidad & ")"


        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    GenerarFitoCampos = True
    
    
    
    
eGenerarFitoCampos:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        GenerarFitoCampos = False
    End If
    Set miRsAux = Nothing
    Set RN = Nothing
End Function




Private Sub LeerGuardarSeleccionProveedoresAQuitar(Leer As Boolean)
Dim NF As Integer

    On Error GoTo eLeerGuardarSeleccionProveedoresAQuitar
    
    miSQL = App.Path & "\provquit.dat"
    If Leer Then
        auxiliar = ""
        If Dir(miSQL, vbArchive) <> "" Then
            NF = FreeFile
            Open miSQL For Input As #NF
            Line Input #NF, auxiliar
            Close #NF
            

        End If
        If auxiliar = "" Then auxiliar = "|"
    Else
    
        '1.- En auxiliar tenemos los que QUITA
        '2.- Comprobaremos que de los que quita de normal, no lo ha vuelto a seleccionar
        '3.- Comprobaremos que de los que ha quitado estan en auxiliar
        
        
        For NumRegElim = 1 To lw(1).ListItems.Count
            If Not lw(1).ListItems(NumRegElim).Checked Then
                cadSelect = "|" & lw(1).ListItems(NumRegElim).Text & "|"
                If InStr(1, auxiliar, cadSelect) > 0 Then
                    'YA esta
                Else
                    auxiliar = auxiliar & lw(1).ListItems(NumRegElim).Text & "|"
                End If
            End If
        Next
        
        For NumRegElim = 1 To lw(1).ListItems.Count
            If lw(1).ListItems(NumRegElim).Checked Then
                'Vemos si de los que no seelccionaba lo ha vuelto a marcar
                cadSelect = "|" & lw(1).ListItems(NumRegElim).Text & "|"
                NF = InStr(1, auxiliar, cadSelect)
                If NF > 0 Then
                    
                    cadTitulo = Mid(auxiliar, 1, NF)    'Dejo el pipe con esta
                    NF = InStr(NF + 1, auxiliar, "|")
                    If NF > 0 Then
                        cadPDFrpt = Mid(auxiliar, NF + 1) 'quito el pipe
                    Else
                        cadPDFrpt = ""
                    End If
                    auxiliar = cadTitulo & cadPDFrpt
                End If
            End If
        Next
        
    
    
    
    
    
            NF = FreeFile
            Open miSQL For Output As #NF
            Print #NF, auxiliar
            Close #NF
    

    End If


    Exit Sub
eLeerGuardarSeleccionProveedoresAQuitar:
    MuestraError Err.Number, Err.Description
    auxiliar = ""
End Sub




Private Sub CargaAlbaranesFacturaClienteEuler()
    On Error GoTo eCargaAlbaranesFacturaClienteEuler
    
    lw(3).ListItems.Clear
    If Me.chkPedidos.Value = 1 Then
        miSQL = " select 'PED',scaped.numpedcl,fecpedcl,sum(importel),-1 from scaped,sliped "
        miSQL = miSQL & " where  scaped.numpedcl = sliped.numpedcl"
        miSQL = miSQL & " and codclien=" & OtrosDatos & " group by 1,2"
    Else
        miSQL = " select scaalb.codtipom,scaalb.numalbar,fechaalb,sum(importel),-1 from scaalb,slialb"
        miSQL = miSQL & " where scaalb.codtipom = slialb.codtipom And scaalb.NumAlbar = slialb.NumAlbar"
        miSQL = miSQL & " and codclien=" & OtrosDatos & " group by 1,2"
        miSQL = miSQL & " Union"
        miSQL = miSQL & " select scafac1.codtipoa,numalbar,fechaalb,brutofac,scafac.NumFactu  from scafac, scafac1 where"
        miSQL = miSQL & " scafac.codtipom = scafac1.codtipom And scafac.NumFactu = scafac1.NumFactu And "
        miSQL = miSQL & " scafac.FecFactu = scafac1.FecFactu and codclien=" & OtrosDatos
    End If
    miSQL = miSQL & " order by 1,3 desc,2"
        
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        
        lw(3).ListItems.Add , , miRsAux.Fields(0)
        NumRegElim = NumRegElim + 1
        lw(3).ListItems(NumRegElim).SubItems(1) = Format(miRsAux.Fields(1), "0000000")
        lw(3).ListItems(NumRegElim).SubItems(2) = Format(miRsAux.Fields(2), "dd/mm/yyyy")
        lw(3).ListItems(NumRegElim).SubItems(3) = Format(miRsAux.Fields(3), FormatoImporte)
        
        'lw(3).ListItems(NumRegElim).SubItems(4) = IIf(miRsAux.Fields(4) = 1, "Si", " ")
        If miRsAux.Fields(4) = -1 Then
            lw(3).ListItems(NumRegElim).SubItems(4) = " "
        Else
            lw(3).ListItems(NumRegElim).SubItems(4) = Format(miRsAux.Fields(4), "000000")
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    CadenaDesdeOtroForm = ""
eCargaAlbaranesFacturaClienteEuler:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Sub




Private Sub CargarFacturasVentaCliente()
Dim EnAlbaranes As Boolean
   
    'El recordset lo hemos abierto en form albaranes
    'Para que si no hay resultados, NO abra el este form para no mostrar res
    'Set miRsAux = New ADODB.Recordset
    'miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Si es articulo de varios
    Label4(9).Caption = ""
    
    'Articulo de varios
    cadParam = "artvario"
    miSQL = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", RecuperaValor(Me.OtrosDatos, 1), "T", cadParam)
    Label4(9).Caption = miRsAux!codArtic & " - " & miSQL
    conSubRPT = cadParam = "1"
    
    
    '----------------------------------------------------------------------------
    '----------------------------------------------------------------------------
    ' Primero las facturas,y para herbelca vamos tambien a albaranes
    '----------------------------------------------------------------------------
    '----------------------------------------------------------------------------
    NumRegElim = 0
    lw(4).ListItems.Clear
    
    
    
    
    If InStr(1, miRsAux.Source, "slifac") > 0 Then
    
        
        While Not miRsAux.EOF
            lw(4).ListItems.Add , , miRsAux!codtipom
            NumRegElim = NumRegElim + 1
            lw(4).ListItems(NumRegElim).SubItems(1) = Format(miRsAux!Numfactu, "00000")
            lw(4).ListItems(NumRegElim).SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
            
            lw(4).ListItems(NumRegElim).SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
            lw(4).ListItems(NumRegElim).SubItems(4) = DBLet(miRsAux!origpre, "T") & " "
            
            
            lw(4).ListItems(NumRegElim).SubItems(5) = Format(miRsAux!cantidad, FormatoCantidad)
            miSQL = " "
            If miRsAux!dtoline1 > 0 Then miSQL = Format(miRsAux!dtoline1, FormatoDescuento)
            lw(4).ListItems(NumRegElim).SubItems(6) = miSQL
            miSQL = " "
            If miRsAux!dtoline2 > 0 Then miSQL = Format(miRsAux!dtoline2, FormatoDescuento)
            lw(4).ListItems(NumRegElim).SubItems(7) = miSQL
        
            lw(4).ListItems(NumRegElim).SubItems(8) = Format(miRsAux!ImporteL, FormatoImporte)
            
            
            'Si es articulo de varios lo especifico
            miSQL = " "
            If conSubRPT Then miSQL = miSQL & miRsAux!NomArtic
            lw(4).ListItems(NumRegElim).SubItems(9) = miSQL
            
            'codtipom numfactu fecfactu codtipoa numalbar numlinea
            miSQL = "codtipoa = " & DBSet(miRsAux!Codtipoa, "T") & " AND numalbar = " & miRsAux!Numalbar & " AND numlinea =" & miRsAux!numlinea
            lw(4).ListItems(NumRegElim).Tag = miSQL
            
            'Si es negativo:
            If miRsAux!cantidad < 0 Then
                lw(4).ListItems(NumRegElim).ForeColor = vbRed
                For numParam = 1 To lw(4).ColumnHeaders.Count - 1
                    lw(4).ListItems(NumRegElim).ListSubItems(numParam).ForeColor = vbRed
                Next
                lw(4).ListItems(NumRegElim).Tag = ""   'Sera cantidad negativa. Para que cuando lo seleccionen, no lo pueda devovler
            End If
            miRsAux.MoveNext
        Wend
                
                
        If vParamAplic.NumeroInstalacion = 2 Then
            'Vamos a abrir albaranes
            miRsAux.Close
                    
            numParam = InStr(1, miRsAux.Source, " codclien =")
            cadFormula = Mid(miRsAux.Source, numParam + 11)
            numParam = InStr(1, LCase(cadFormula), " and ")
            cadFormula = Mid(cadFormula, 1, numParam)
            
            miSQL = " Select slialb.*,fechaalb from slialb,scaalb     where  scaalb.codtipom=slialb.codtipom and"
            miSQL = miSQL & " scaalb.NumAlbar = slialb.NumAlbar  AND codclien = " & cadFormula
            miSQL = miSQL & " AND codartic<>" & DBSet(vParamAplic.ArtReciclado, "T")
            miSQL = miSQL & " AND scaalb.codtipom <>'ALZ' "    'para quitar los que no sean albaranes
            miSQL = miSQL & " AND codartic = " & DBSet(RecuperaValor(Me.OtrosDatos, 1), "T")
            miSQL = miSQL & " ORDER BY fechaalb,numlinea"
            miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            EnAlbaranes = True
        End If
    Else
        EnAlbaranes = True
    End If
    
    If EnAlbaranes Then
        
        
        While Not miRsAux.EOF
            lw(4).ListItems.Add , , miRsAux!codtipom
            NumRegElim = NumRegElim + 1
            lw(4).ListItems(NumRegElim).SubItems(1) = Format(miRsAux!Numalbar, "00000")
            lw(4).ListItems(NumRegElim).SubItems(2) = Format(miRsAux!FechaAlb, "dd/mm/yyyy")
            
            lw(4).ListItems(NumRegElim).SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
            lw(4).ListItems(NumRegElim).SubItems(4) = DBLet(miRsAux!origpre, "T") & " "
            
            
            lw(4).ListItems(NumRegElim).SubItems(5) = Format(miRsAux!cantidad, FormatoCantidad)
            miSQL = " "
            If miRsAux!dtoline1 > 0 Then miSQL = Format(miRsAux!dtoline1, FormatoDescuento)
            lw(4).ListItems(NumRegElim).SubItems(6) = miSQL
            miSQL = " "
            If miRsAux!dtoline2 > 0 Then miSQL = Format(miRsAux!dtoline2, FormatoDescuento)
            lw(4).ListItems(NumRegElim).SubItems(7) = miSQL
        
            lw(4).ListItems(NumRegElim).SubItems(8) = Format(miRsAux!ImporteL, FormatoImporte)
            
            
            'Si es articulo de varios lo especifico
            miSQL = " "
            If conSubRPT Then miSQL = miSQL & miRsAux!NomArtic
            lw(4).ListItems(NumRegElim).SubItems(9) = miSQL
            
            'codtipom numfactu fecfactu codtipoa numalbar numlinea
            miSQL = "codtipom = " & DBSet(miRsAux!codtipom, "T") & " AND numalbar = " & miRsAux!Numalbar & " AND numlinea =" & miRsAux!numlinea
            lw(4).ListItems(NumRegElim).Tag = miSQL
            
            'Si es negativo:
            If miRsAux!cantidad < 0 Then
                lw(4).ListItems(NumRegElim).ForeColor = vbRed
                For numParam = 1 To lw(4).ColumnHeaders.Count - 1
                    lw(4).ListItems(NumRegElim).ListSubItems(numParam).ForeColor = vbRed
                Next
                lw(4).ListItems(NumRegElim).Tag = ""   'Sera cantidad negativa. Para que cuando lo seleccionen, no lo pueda devovler
            End If
            miRsAux.MoveNext
        Wend
        
   
    End If
    Set lw(4).SelectedItem = lw(4).ListItems(1)
    DoEvents
    lw(4).SetFocus
    
    
    
    numParam = 0
End Sub


Private Function HacerIncrementoPuntosCliente() As Boolean
    On Error GoTo eHacerIncrementoPuntosCliente
    HacerIncrementoPuntosCliente = False
    miSQL = DBSet(Me.txtNumero(3).Text, "N")
    cadFormula = RecuperaValor(OtrosDatos, 1)
    cadPDFrpt = "UPDATE sclien set puntos=" & miSQL & "+ coalesce(puntos,0) WHERE codclien=" & cadFormula
    conn.Execute cadPDFrpt
    
    cadPDFrpt = DevuelveDesdeBD(conAri, "max(numero)", "smovalpuntos", "codclien", cadFormula)
    If cadPDFrpt = "" Then cadPDFrpt = "0"
    cadPDFrpt = CStr(Val(cadPDFrpt) + 1)
    
    cadSelect = "INSERT INTO smovalpuntos(codclien,numero,codtipom,numalbar,fechaalb,concepto,puntos,fecMov,observaciones) VALUES ("
    cadSelect = cadSelect & cadFormula & "," & cadPDFrpt & ",'',0," & DBSet(txtFecha(10).Text, "F") & ",2,"
    cadSelect = cadSelect & miSQL & ",now()," & DBSet(Me.txtModificable(3).Text, "T") & ")"
    conn.Execute cadSelect
    HacerIncrementoPuntosCliente = True
    Exit Function
eHacerIncrementoPuntosCliente:
    MuestraError Err.Number, Err.Description
End Function


Private Sub CargaComoboTrimestre(Indice As Integer)
    cboTrimiestre(Indice).Clear
    cboTrimiestre(Indice).AddItem "1er trimestre"
    cboTrimiestre(Indice).AddItem "2º trimestre"
    cboTrimiestre(Indice).AddItem "3er trimestre"
    cboTrimiestre(Indice).AddItem "4º trimestre"
End Sub

Private Sub DatosTrimestreAnterior()


    Set miRsAux = New ADODB.Recordset
    miSQL = "Select * from declarafontenas ORDER BY trimestre DESC"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'NO puede ser EOF
    '-------------------------
    NumRegElim = miRsAux!trimestre
    'Voy a calcular el proximo a declarar
    numParam = CByte(Right(CStr(NumRegElim), 2))
    NumRegElim = CInt(Left(CStr(NumRegElim), 4))
    
    miSQL = Me.cboTrimiestre(0).List(numParam - 1) & " de " & NumRegElim
    miSQL = miSQL & "   Saldo: " & Format(miRsAux!saldotrimestre, FormatoCantidad)
    
    If numParam = 4 Then
    
        numParam = 1 'El ultimo fue el cuarto
        NumRegElim = NumRegElim + 1
    Else
        numParam = numParam + 1
    End If
    txtNumero(4).Text = NumRegElim
    Me.cboTrimiestre(0).ListIndex = numParam - 1
    txtNoModificable(5).Text = miSQL
    txtNoModificable(5).Tag = NumRegElim & Format(numParam, "00")
    miRsAux.Close
    
    Set miRsAux = Nothing

End Sub


Private Function GeneraDatosDeclaraAlcohol() As Boolean
Dim RS As ADODB.Recordset
Dim fin As Boolean
Dim Entradas As Currency
Dim Salidas As Currency
Dim SQLNumlinea As String

    On Error GoTo eGeneraDatosDeclaraAlcohol
    GeneraDatosDeclaraAlcohol = False
    
    cadTitulo = "" 'llevara las fechas para elñ rpt
    
    NumRegElim = Me.cboTrimiestre(0).ListIndex * 3 'que trimestre
    NumRegElim = NumRegElim + 1
    cadSelect = "'" & Me.txtNumero(4).Text & "-" & Format(NumRegElim, "00") & "-01'"
    cadTitulo = "01/" & Format(NumRegElim, "00") & "/" & Me.txtNumero(4).Text
    NumRegElim = NumRegElim + 2
    cadFormula = DiasMes(CByte(NumRegElim), CInt(txtNumero(4).Text))
    cadTitulo = cadTitulo & cadFormula & "/" & Format(NumRegElim, "00") & "/" & Me.txtNumero(4).Text
    cadFormula = "'" & Me.txtNumero(4).Text & "-" & Format(NumRegElim, "00") & "-" & cadFormula & "'"
    
    'Para cuando vaya a buscar NUMLOTE a las facturas, para acotar la busqueda
    cadPDFrpt = "between " & cadSelect & " AND " & cadFormula
    
    cadSelect = "select * from smoval where fechamov  between " & cadSelect & " AND " & cadFormula
    'Deberiamos ponerlo en parametro
    cadSelect = cadSelect & " AND codartic = '0ALC99AL' "
    
    
    'Guardamos en esta varible para las fechas, ya que cadPDFrpt lo utilizamos aqui bajo
    cadFormula = cadPDFrpt
    
    
    'Cargamos dos RS , entradas y salidas
    Set miRsAux = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    
    conn.Execute "DELETE from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute "DELETE from tmpprevision2 where codusu = " & vUsu.Codigo

    
    NumRegElim = 0
    'Entradas
    miSQL = cadSelect & " AND detamovi = 'ALC' ORDER BY fechamov"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    'miSQL = cadSelect & " AND detamovi = 'PRO' And tipomovi=0  ORDER BY fechamov"
    
    miSQL = cadSelect & " AND detamovi IN ('PRO','PRE') And tipomovi=0  ORDER BY fechamov"
    
        RS.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    NumRegElim = 0
    Entradas = 0
    Salidas = 0
    fin = miRsAux.EOF And RS.EOF   'Los dos vacios
    'insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`campo2`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`importe2`,`importe3`,`importe4`,`importe5`,`porcen1`,`porcen2`,`importeb1`,`importeb2`,`importeb3`,`importeb4`,`importeb5`,`fecha1`,`fecha2`,`obser`) values ( '0','1','12',NULL,'lote1','loteventa','descrip','1000','2',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'20010101','201012012','12')
    cadPDFrpt = "" 'El insert into
    
    
    While Not fin
        
        'Comun
        NumRegElim = NumRegElim + 1
        '`codusu`,`codigo1`,

        cadParam = ", (" & vUsu.Codigo & "," & NumRegElim & ","
        
        'Compras
        'campo1 nombre1 importe1   codprove,lote,cantidad,fecha
        cadNomRPT = "NULL,NULL,NULL,NULL"
        If Not miRsAux.EOF Then
        
        
            'select numlotes from slialp where codartic='0ALC99AL' and numalbar ='Da0001' and codprove=46066
            OtrosDatos = "codartic='0ALC99AL' and numalbar =" & DBSet(miRsAux!Document, "T") & " AND codprove "
            miSQL = DevuelveDesdeBD(conAri, "numlotes", "slialp", OtrosDatos, miRsAux!codigope)
            If miSQL = "" Then
                OtrosDatos = "slifpc.codprove=scafpa.codprove and slifpc.numfactu=scafpa.numfactu and"
                OtrosDatos = OtrosDatos & " slifpc.fecfactu=scafpa.fecfactu and slifpc.numalbar=scafpa.numalbar and"
                OtrosDatos = OtrosDatos & " scafpa.fechaalb " & cadFormula
                OtrosDatos = OtrosDatos & " AND slifpc.codartic='0ALC99AL' and scafpa.numalbar =" & DBSet(miRsAux!Document, "T") & " AND scafpa.codprove "
                miSQL = DevuelveDesdeBD(conAri, "numlotes", "slifpc,scafpa ", OtrosDatos, miRsAux!codigope)
                
            End If
        
            cadNomRPT = miRsAux!codigope & "," & DBSet(miSQL, "T", "S") & "," & DBSet(miRsAux!cantidad, "N") & "," & DBSet(miRsAux!FechaMov, "F")
            Entradas = Entradas + miRsAux!cantidad
            miRsAux.MoveNext
        End If
        cadParam = cadParam & cadNomRPT & ","
        
        
        
        
        'Produccion-SALIDAS
        ' formato  lote   art fab   cantidad fecha   'En produccion será formato: Que se ha producido  Art fab: articulo fabricado
        'nombre2,nombre3,observa,importe2,fecha2
       
        cadNomRPT = "NULL,NULL,NULL,NULL,NULL"
        If Not RS.EOF Then
    
            If RS!detamovi = "PRE" Then
                SQLNumlinea = "if (round(cantidad,2) -" & DBSet(RS!cantidad, "N") & "=0,0,1)"
                OtrosDatos = ""
                OtrosDatos = "slienvpr2.codartic=sartic.codartic and sartic.codfamia=sfamia.codfamia " & OtrosDatos & " and codarti2='0ALC99AL' and codigo"
                miSQL = DevuelveDesdeBD(conAri, "concat(numlote,'|',nomfamia,'|')", "slienvpr2,sartic,sfamia", OtrosDatos, RS!Document & " ORDER BY 2 ASC", , SQLNumlinea)
                If miSQL = "" Then
                    MsgBox "No se encuentra el movimiento: " & RS!FechaMov & " " & RS!codArtic & " ENV: " & RS!Document
                Else
                    OtrosDatos = DevuelveDesdeBD(conAri, "cantidad", "slienvpr", "codigo", RS!Document)  'Canitdad producida
                    Salidas = Salidas + RS!cantidad
                    cadNomRPT = "'" & OtrosDatos & "L'," & DBSet(RecuperaValor(miSQL, 1), "T") & "," & DBSet(RecuperaValor(miSQL, 2), "T")
                    cadNomRPT = cadNomRPT & "," & DBSet(RS!cantidad, "N") & "," & DBSet(RS!FechaMov, "F")
                End If
            
            Else
                SQLNumlinea = "if (round(cantidad,2) -" & DBSet(RS!cantidad, "N") & "=0,0,1)"
                
                OtrosDatos = ""
               
                OtrosDatos = "sliordpr2.codartic=sartic.codartic and sartic.codfamia=sfamia.codfamia and codarti2='0ALC99AL' " & OtrosDatos & " and codigo"
                miSQL = DevuelveDesdeBD(conAri, "concat(numlote,'|',nomfamia,'|')", "sliordpr2,sartic,sfamia", OtrosDatos, RS!Document & " ORDER BY 2 ASC", , SQLNumlinea)
                If miSQL = "" Then
                    MsgBox "No se encuentra el movimiento: " & RS!FechaMov & " " & RS!codArtic & " Prod: " & RS!Document, vbExclamation
                    
                Else
                    OtrosDatos = DevuelveDesdeBD(conAri, "cantidad", "sliordpr", "codigo", RS!Document)  'Cantidad producida
                    Salidas = Salidas + RS!cantidad
                    cadNomRPT = "'" & OtrosDatos & "L'," & DBSet(RecuperaValor(miSQL, 1), "T") & "," & DBSet(RecuperaValor(miSQL, 2), "T")
                    cadNomRPT = cadNomRPT & "," & DBSet(RS!cantidad, "N") & "," & DBSet(RS!FechaMov, "F")
                    
                End If
            
            End If
            RS.MoveNext
        End If
        
        
        cadParam = cadParam & cadNomRPT & ")"   'UNA LINEA
        cadPDFrpt = cadPDFrpt & cadParam
        
        
        If (NumRegElim Mod 30) = 0 Then
            
            cadPDFrpt = Mid(cadPDFrpt, 2) 'quitamos la primera coma
            cadParam = "INSERT INTO tmpinformes (codusu,codigo1,campo1 ,nombre1 ,importe1,fecha1,nombre2,nombre3,obser,importe2,fecha2) VALUES "
            cadParam = cadParam & cadPDFrpt
            conn.Execute cadParam
            cadPDFrpt = ""
        End If
        
        
        
        If miRsAux.EOF And RS.EOF Then fin = True
    
    
    Wend
    miRsAux.Close
    RS.Close
    If cadPDFrpt <> "" Then
        cadPDFrpt = Mid(cadPDFrpt, 2) 'quitamos la primera coma
        cadParam = "INSERT INTO tmpinformes (codusu,codigo1,campo1 ,nombre1 ,importe1,fecha1,nombre2,nombre3,obser,importe2,fecha2) VALUES "
        cadParam = cadParam & cadPDFrpt
        conn.Execute cadParam
    End If
    
    auxiliar = "INSERT INTO declarafontenas(trimestre,fechamov,usuario,pc,saldo_anterior,entradas,salidas,saldotrimestre) VALUES ("
    auxiliar = auxiliar & Format(Me.txtNumero(4).Text, "0000") & Format(cboTrimiestre(0).ListIndex + 1, "00") & ","
    auxiliar = auxiliar & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(vUsu.PC, "T") & ","
    
    cadParam = "INSERT INTO tmpprevision2(codusu,tipo,descripcion,importe1,importe2,importe3,importe4) VALUES (" & vUsu.Codigo & ",1,"
    
    cadParam = cadParam & "'" & cadTitulo & Me.cboTrimiestre(0).Text & "/" & txtNumero(4).Text & "',"
    ' miSQL = "Select * from declarafontenas ORDER BY trimestre DESC"
    cadPDFrpt = DevuelveDesdeBD(conAri, "saldotrimestre", "declarafontenas", "1", " 1 ORDER BY trimestre DESC")
    If cadPDFrpt = "" Then cadPDFrpt = "0"
    
    cadParam = cadParam & DBSet(cadPDFrpt, "N") & "," & DBSet(Entradas, "N") & "," & DBSet(Salidas, "N") & ","
    auxiliar = auxiliar & DBSet(cadPDFrpt, "N") & "," & DBSet(Entradas, "N") & "," & DBSet(Salidas, "N") & ","
    
    Entradas = Entradas - Salidas + CCur(cadPDFrpt)
    cadParam = cadParam & DBSet(Entradas, "N") & ")"
    auxiliar = auxiliar & DBSet(Entradas, "N") & ")"
    
    conn.Execute cadParam
    
    
    
    
    If NumRegElim = 0 Then
        MsgBox "No se han generado datos", vbExclamation
    Else
        GeneraDatosDeclaraAlcohol = True
    End If
eGeneraDatosDeclaraAlcohol:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set RS = Nothing
    
End Function



Private Sub GenerarFicheroAlcohol()
Dim NF As Integer

    On Error GoTo eGenerarFicheroAlcohol
    miSQL = App.Path & "\fonten.csv"
    If Dir(miSQL, vbArchive) <> "" Then Kill miSQL
    
    NF = FreeFile
    Open miSQL For Output As #NF
    
    miSQL = "Periodo;" & Me.cboTrimiestre(0).Text & ";;"
    Print #NF, miSQL
    Print #NF, ";;;;"
    Print #NF, ";;;;"
    Print #NF, ";;;;"
    Print #NF, ";;ENTRADAS;;;;;;SALIDAS"
    Print #NF, ";;;;"
    Print #NF, "Fecha;Proveedor;Cantidad;Lote;;FECHA;Fabricacion;NºLote;Descripcion;Cantidad;"
    Print #NF, ";;;;"
    miSQL = "Select tmpinformes.*,nomprove from tmpinformes left join sprove on codprove=campo1 where codusu = " & vUsu.Codigo & " ORDER by codigo1"
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If IsNull(miRsAux!fecha1) Then
            miSQL = ";;;;;"
        Else
            miSQL = miRsAux!fecha1 & ";" & Replace(DBLet(miRsAux!nomprove, "T"), ";", "") & ";" & miRsAux!Importe1 & ";;;"
        End If
        
        If IsNull(miRsAux!fecha2) Then
            miSQL = miSQL & ";;;;;"
        Else
            miSQL = miSQL & miRsAux!fecha2 & ";" & Replace(DBLet(miRsAux!nombre2, "T"), ";", "") & ";"
            miSQL = miSQL & Replace(DBLet(miRsAux!nombre3, "T"), ";", "") & ";" & Replace(DBLet(miRsAux!obser, "T"), ";", "") & ";"
            miSQL = miSQL & miRsAux!Importe2 & ";"
        End If
        
        Print #NF, miSQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Print #NF, ";;;;"
    Print #NF, ";;;;"
    Print #NF, ";;;;"
    miSQL = "Select * from tmpprevision2 where codusu = " & vUsu.Codigo
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'No es eof
    Print #NF, ";;Total;" & miRsAux!Importe2 & ";;;;;Total;" & miRsAux!Importe3
    
    Print #NF, ";;;;"
    Print #NF, ";;;;"
    Print #NF, ";;;;"
    Print #NF, ";;;;"
    Print #NF, "Anterior;Entradas;Salidas;Saldo;;;;;;;"
    Print #NF, miRsAux!Importe1 & ";" & miRsAux!Importe2 & ";" & miRsAux!Importe3 & ";" & miRsAux!importe4 & ";;;;;;"
    Print #NF, ";;;;"
    Print #NF, ";;;;"
    Print #NF, ";;;;"
    Close #NF
    cd1.CancelError = True
    cd1.FileName = ""
    cd1.Filter = "*.csv|*.csv"
    cd1.ShowSave
    If cd1.FileTitle <> "" Then
        FileCopy App.Path & "\fonten.csv", cd1.FileName
        MsgBox "Fichero creado con exito: " & cd1.FileName, vbInformation
    End If
    
    Exit Sub
eGenerarFicheroAlcohol:
    If Err.Number <> 32755 Then MuestraError Err.Number, , Err.Description
    Err.Clear
End Sub






'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'
'   Reimpresion, facturas albaranes signotec
'
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Private Sub CargaDatosReimpresion()
    
    
    numParam = 0
    'Primero la factura... si tiene
    If Mid(CadenaDesdeOtroForm, 1, 1) = "@" Then
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
        NumRegElim = InStr(1, CadenaDesdeOtroForm, "@")
        
        cadParam = Mid(CadenaDesdeOtroForm, 1, NumRegElim - 1)
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, NumRegElim + 1)
        
        'Nombre , path
        NumRegElim = InStr(1, cadParam, "#")
        cadSelect = Mid(cadParam, 1, NumRegElim - 1)
        cadParam = Mid(cadParam, NumRegElim + 1)
        numParam = numParam + 1
        lwSigno.ListItems.Add , , "Si"
        lwSigno.ListItems(numParam).SubItems(1) = Mid(cadSelect, 1, 10)
        lwSigno.ListItems(numParam).SubItems(2) = Mid(cadSelect, 11)
        lwSigno.ListItems(numParam).Tag = cadParam
        lwSigno.ListItems(numParam).SmallIcon = 8
    End If
    
    'Despues albaranes , si tiene
    While CadenaDesdeOtroForm <> ""
        
        
        NumRegElim = InStr(1, CadenaDesdeOtroForm, "@")
        
        cadParam = Mid(CadenaDesdeOtroForm, 1, NumRegElim - 1)
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, NumRegElim + 1)
    
    
         'Nombre , path
        NumRegElim = InStr(1, cadParam, "#")
        cadSelect = Mid(cadParam, 1, NumRegElim - 1)
        cadParam = Mid(cadParam, NumRegElim + 1)
        numParam = numParam + 1
        lwSigno.ListItems.Add , , " "
        lwSigno.ListItems(numParam).SubItems(1) = Mid(cadSelect, 1, 10)
        lwSigno.ListItems(numParam).SubItems(2) = Mid(cadSelect, 11)
        lwSigno.ListItems(numParam).Tag = cadParam
        lwSigno.ListItems(numParam).SmallIcon = 7
    
    
        
    Wend
    
End Sub




'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'
'   CLIENTE    Acciones comerciales
'
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Private Sub CargaDatosAccionesComerciales()
    
    miSQL = "Select  codigo  , denominacion  from scrmtipo where codigo>=21 ORDER BY  2  "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic
    NumRegElim = 0
    While Not miRsAux.EOF
        
        
        NumRegElim = NumRegElim + 1
        lw(5).ListItems.Add , , Format(miRsAux!Codigo, "000")
        lw(5).ListItems(NumRegElim).SubItems(1) = miRsAux!denominacion
        lw(5).ListItems(NumRegElim).Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub





Private Sub CargaDatosPedidoClienteAlbaran()
    Dim IT
    Dim EsAlbaran As Boolean
    Me.Caption = "Ordenar lineas "
    If Mid(OtrosDatos, 1, 1) = "A" Then
        Me.Caption = Me.Caption & "albaran"
        miSQL = "Select  * from slialb where codtipom= '" & Mid(OtrosDatos, 1, 3) & "' AND numalbar=" & Mid(OtrosDatos, 4) & " ORDER BY ordenlin,numlinea"
        EsAlbaran = True
    Else
        Me.Caption = Me.Caption & "pedido"
        miSQL = "Select  * from sliped where numpedcl =   " & OtrosDatos & " ORDER BY numlinea"
        EsAlbaran = False
    End If
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic
    NumRegElim = 1
    While Not miRsAux.EOF
        
        
        Set IT = lw(6).ListItems.Add()
      
        IT.Text = CStr(Format(miRsAux!numlinea, "000"))
        
        IT.SubItems(1) = CStr(miRsAux!codArtic)
        IT.SubItems(2) = CStr(miRsAux!NomArtic)
        IT.SubItems(3) = Format(miRsAux!cantidad, FormatoCantidad)
        If Not EsAlbaran Then IT.SubItems(4) = Format(miRsAux!solicitadas, FormatoCantidad)
        IT.SubItems(5) = Format(miRsAux!precioar, FormatoPrecio)
        IT.SubItems(6) = Format(miRsAux!ImporteL, FormatoCantidad)
        IT.SubItems(7) = Format(NumRegElim, "00000")
        
        NumRegElim = NumRegElim + 1
        
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If EsAlbaran Then
        lw(6).ColumnHeaders(4).Text = "Cantidad"
        lw(6).ColumnHeaders(5).Width = 0
        lw(6).ColumnHeaders(4).Width = 1200
        lw(6).ColumnHeaders(6).Width = 1200
        lw(6).ColumnHeaders(7).Width = 1800
    End If
End Sub






Private Function ListadoFacturasAlbaranes() As Boolean
    
    On Error GoTo eListadoFacturasAlbaranes
    
    ListadoFacturasAlbaranes = False
    
    miSQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
    cadSelect = Replace(cadSelect, "{", "")
    cadSelect = Replace(cadSelect, "}", "")
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 0
    
    'Albaranes
    If Me.chkAlb(0).Value Then
        
        miSQL = "select scaalb.codtipom,scaalb.numalbar,fechaalb,codclien,nomclien,sum(importel) import from scaalb inner join slialb on scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar"
        miSQL = miSQL & " Where " & cadSelect
        miSQL = miSQL & " group by 1,2 ORDER BY 1,2,3"
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        auxiliar = ""
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            
            'tmpinformes   codusu codigo1   nombre1 campo1  fecha1      nombre2    importe1   nombre3 fecha2
        
            miSQL = ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & miRsAux!codtipom & "'," & miRsAux!Numalbar & "," & DBSet(miRsAux!FechaAlb, "F")
            miSQL = miSQL & "," & DBSet(Format(miRsAux!codClien, "00000") & "    " & miRsAux!NomClien, "T") & "," & DBSet(miRsAux!Import, "N") & ",null,null)"
            
            auxiliar = auxiliar & miSQL
            If Len(auxiliar) > 2000 Then InsertaEnTmpLstFactu
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        InsertaEnTmpLstFactu
    End If
    
    'Facturas
    If Me.chkAlb(1).Value Then
        
        miSQL = "select scafac1.codtipoa,scafac1.numalbar,fechaalb,codclien,nomclien,sum(importel) import,scafac1.codtipom,scafac1.numfactu,scafac1.fecfactu from"
        miSQL = miSQL & " scafac inner join scafac1 on scafac.codtipom=scafac1.codtipom and scafac.numfactu=scafac1.numfactu and scafac.fecfactu=scafac1.fecfactu"
        miSQL = miSQL & " inner join slifac on scafac.codtipom=slifac.codtipom and scafac.numfactu=slifac.numfactu and scafac.fecfactu=slifac.fecfactu"
        miSQL = miSQL & " and scafac1.numalbar=slifac.numalbar and scafac1.codtipoa=slifac.codtipoa"
        
        auxiliar = Replace(cadSelect, "scaalb", "scafac1")
        auxiliar = Replace(auxiliar, "codtipom", "codtipoa")
        miSQL = miSQL & " WHERE  " & auxiliar
        '(scafac1.codtipoa)<>'ALZ' AND (((scafac1.fechaalb >= '2018-01-01') and (scafac1.fechaalb <= '2019-03-24')))
        miSQL = miSQL & " group by 1,2 "

        
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        auxiliar = ""
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            
            'tmpinformes   codusu codigo1   campo1 nombre1 fecha1      nombre2    importe1   nombre3 fecha2
        
            miSQL = ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & miRsAux!Codtipoa & "'," & miRsAux!Numalbar & "," & DBSet(miRsAux!FechaAlb, "F")
            miSQL = miSQL & "," & DBSet(Format(miRsAux!codClien, "00000") & "   " & miRsAux!NomClien, "T") & "," & DBSet(miRsAux!Import, "N")
            miSQL = miSQL & "," & DBSet(miRsAux!codtipom & Format(miRsAux!Numfactu, "000000"), "T") & "," & DBSet(miRsAux!FecFactu, "F") & ")"
            
            auxiliar = auxiliar & miSQL
            If Len(auxiliar) > 6000 Then InsertaEnTmpLstFactu
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        InsertaEnTmpLstFactu
    End If
    
    
    If NumRegElim = 0 Then
        MsgBox "Ningun dato generado", vbInformation
    Else
        ListadoFacturasAlbaranes = True
    End If
eListadoFacturasAlbaranes:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    auxiliar = ""
End Function


Private Sub InsertaEnTmpLstFactu()
    If auxiliar = "" Then Exit Sub
    auxiliar = Mid(auxiliar, 2)
    miSQL = "INSERT INTO tmpinformes(codusu ,codigo1 ,nombre1  ,campo1  ,fecha1      ,nombre2    ,importe1   ,nombre3 ,fecha2) VALUES "
    miSQL = miSQL & auxiliar
    conn.Execute miSQL
    auxiliar = ""
End Sub



Private Function ReordenarLineas() As Boolean
    
    On Error GoTo eReordenarLineas
    ReordenarLineas = False
    
    If Mid(OtrosDatos, 1, 1) = "A" Then
    
        For NumRegElim = 1 To lw(6).ListItems.Count
            miSQL = "UPDATE slialb set ordenlin=" & NumRegElim & " WHERE "
            miSQL = miSQL & " codtipom = '" & Trim(Mid(OtrosDatos, 1, 3)) & "' and numalbar=" & Mid(OtrosDatos, 4) & " And numlinea = " & lw(6).ListItems(NumRegElim).Text
            
            conn.Execute miSQL
           
        Next
    
    
    Else
        miSQL = "UPDATE sliped SET numlinea = numlinea +32000 WHERE numpedcl=  " & OtrosDatos
        conn.Execute miSQL
        For NumRegElim = 1 To lw(6).ListItems.Count
            miSQL = "UPDATE sliped set numlinea=" & NumRegElim & " WHERE numpedcl=  " & OtrosDatos & " AND numlinea=" & lw(6).ListItems(NumRegElim).Text + 32000
            conn.Execute miSQL
        Next
    End If
    
    ReordenarLineas = True
    

eReordenarLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
End Function



Private Function HacerCreditoYCaucion() As Boolean
Dim Importe As Currency
    On Error GoTo eHacerCreditoYCaucion
    HacerCreditoYCaucion = False
    
    
    conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo

'
'
'
    auxiliar = "select coalesce(baseimp1,0) +coalesce(baseimp2,0) +coalesce(baseimp3,0)base , coalesce(imporiv1,0)+coalesce(imporiv2,0) +coalesce(imporiv3,0)  iva"
    auxiliar = auxiliar & " , numfactu,letraser,fecfactu,scafac.nomclien,sclien.tipocredito,numgrupo"
    auxiliar = auxiliar & " from scafac ,sclien,stipom  where stipom.codtipom=scafac.codtipom and scafac.codclien=sclien.codclien and"
    auxiliar = auxiliar & " stipom.codtipom='FAV' and fecfactu between "
    auxiliar = auxiliar & DBSet(txtFecha(15).Text, "F") & " AND " & DBSet(txtFecha(16).Text, "F")
    auxiliar = auxiliar & " order by stipom.codtipom,numfactu"

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open auxiliar, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    
    miSQL = "insert into tmpinformes(codusu,nombre1,codigo1, nombre2 ,  fecha1,nombre3,importe1,importe2,importe3,importe4,importe5,importeb1,importeb2,importeb3,importeb4,importeb5,obser)  VALUES  "
    auxiliar = ""
    While Not miRsAux.EOF
        NumRegElim = 1
        
        'codusu,nombre1  base e iva
        auxiliar = auxiliar & ", (" & vUsu.Codigo & ",'" & Mid(miRsAux!Base & Space(20), 1, 20) & Mid(miRsAux!IVA & Space(20), 1, 20) & "'"
            
        ',codigo1, nombre2 ,  fecha1
        auxiliar = auxiliar & "," & miRsAux!Numfactu & "," & DBSet(miRsAux!LetraSer, "T") & "," & DBSet(miRsAux!FecFactu, "F")
        
        'nombre3
        'Tipocredito y nugrupo
        auxiliar = auxiliar & ", '" & Mid(miRsAux!tipocredito & Space(10), 1, 10) & Mid(miRsAux!numGrupo & Space(10), 1, 10) & "',"
        
       
        'CREDITO
        '    30,        60  ,90,        120,     150,   180,       part,    contadao,  nada,    OP
        ',importe1,importe2,importe3,importe4,importe5,importeb1,importeb2,importeb3,importeb4,importeb5) "
        Importe = miRsAux!Base + miRsAux!IVA
        If miRsAux!tipocredito = "N" Then
            'NADA
            If miRsAux!numGrupo = "NADA" Then
                cadNomRPT = "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & DBSet(Importe, "N") & ",null"
            Else
                'Contado
                cadNomRPT = "NULL,NULL,NULL,NULL,NULL,NULL,NULL," & DBSet(Importe, "N") & ",NULL,null"
            End If
        ElseIf miRsAux!tipocredito = "OP" Then
            'organismos publicos
            cadNomRPT = "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & DBSet(Importe, "N")
            
        ElseIf miRsAux!tipocredito = "B" Then
            'Contado
            cadNomRPT = "NULL,NULL,NULL,NULL,NULL,NULL,NULL," & DBSet(Importe, "N") & ",NULL,null"
            
        
        ElseIf miRsAux!tipocredito = "X" Then
            
            If miRsAux!numGrupo = "NADA" Or miRsAux!numGrupo = "NORM" Then
                'CONTADO
                cadNomRPT = "NULL,NULL,NULL,NULL,NULL,NULL,NULL," & DBSet(Importe, "N") & ",NULL,null"
            Else
                ''    30,        60  ,90,        120,     150,   180,       part,    contadao,  nada,    OP
                Select Case miRsAux!numGrupo
                Case 60
                    cadNomRPT = "NULL," & DBSet(Importe, "N") & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,null"
                Case 90
                    cadNomRPT = "NULL,NULL," & DBSet(Importe, "N") & ",NULL,NULL,NULL,NULL,NULL,NULL,null"
                Case 120
                    cadNomRPT = "NULL,NULL,NULL," & DBSet(Importe, "N") & ",NULL,NULL,NULL,NULL,NULL,NULL"
                Case 150
                    cadNomRPT = "NULL,NULL,NULL,NULL," & DBSet(Importe, "N") & ",NULL,NULL,NULL,NULL,null"
                Case 180
                    cadNomRPT = "NULL,NULL,NULL,NULL,NULL," & DBSet(Importe, "N") & ",NULL,NULL,NULL,NULL"
                Case Else
                    cadNomRPT = DBSet(Importe, "N") & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,null"
                    If vUsu.Codigo Mod 1000 = 0 Then
                        If miRsAux!numGrupo <> "30" Then MsgBox miRsAux!NomClien & " " & miRsAux!numGrupo & " " & miRsAux!Numfactu
                    End If
                End Select
                
            End If
        Else
            'mirsaux!tipocredito="O"
            'contado
            cadNomRPT = "NULL,NULL,NULL,NULL,NULL,NULL,NULL," & DBSet(Importe, "N") & ",NULL,null"
        
        End If
            
        auxiliar = auxiliar & cadNomRPT & "," & DBSet(miRsAux!NomClien, "T") & ")"
        auxiliar = miSQL & Mid(auxiliar, 2)
        conn.Execute auxiliar
        auxiliar = ""
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If NumRegElim > 0 Then
        HacerCreditoYCaucion = True
    Else
        MsgBox "Ningun dato generado", vbExclamation
    End If
'
'


eHacerCreditoYCaucion:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Function

Private Sub GenerarFicheroCreditoYCaucion()
Dim NF As Integer

    On Error GoTo eGenerarFicheroCreditoYCaucion
    miSQL = App.Path & "\credicau.csv"
    If Dir(miSQL, vbArchive) <> "" Then Kill miSQL
    
    NF = FreeFile
    Open miSQL For Output As #NF
    
    
    miSQL = "Select codigo1,nombre1,nombre2,nombre3,fecha1,obser,importe1,importe2,importe3,importe4,importe5,importeb1,importeb2,importeb3,importeb4,importeb5"
    miSQL = miSQL & " from tmpinformes where codusu = " & vUsu.Codigo & " ORDER by codigo1"
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        auxiliar = Trim(Mid(miRsAux!nombre1, 1, 20))
        miSQL = auxiliar & ";"
        auxiliar = Trim(Mid(miRsAux!nombre1, 21, 20))
        miSQL = miSQL & auxiliar & ";"
        'tipocredito
        auxiliar = Trim(Mid(miRsAux!nombre3, 1, 10))
        miSQL = miSQL & auxiliar & ";"
        auxiliar = Replace(Trim(miRsAux!obser), ";", "")
        miSQL = miSQL & auxiliar & ";"
        auxiliar = miRsAux!Codigo1 & "/" & miRsAux!nombre2
        miSQL = miSQL & auxiliar & ";"
        miSQL = miSQL & miRsAux!fecha1 & ";"
        
        For NumRegElim = 6 To 15
            If IsNull(miRsAux.Fields(NumRegElim)) Then
                auxiliar = ""
            Else
                auxiliar = miRsAux.Fields(NumRegElim)
            End If
            miSQL = miSQL & auxiliar & ";"
        Next
        Print #NF, miSQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Close #NF
    cd1.CancelError = True
    cd1.FileName = ""
    cd1.Filter = "*.csv|*.csv"
    cd1.ShowSave
    If cd1.FileTitle <> "" Then
        FileCopy App.Path & "\credicau.csv", cd1.FileName
        MsgBox "Fichero creado con exito: " & cd1.FileName, vbInformation
    End If
    
    Exit Sub
eGenerarFicheroCreditoYCaucion:
    If Err.Number <> 32755 Then MuestraError Err.Number, , Err.Description
    Err.Clear
End Sub



Private Function AbrirFicheroYProcesarCoarval() As Byte
    On Error GoTo eAbrirFicheroYProcesarCoarval
    AbrirFicheroYProcesarCoarval = 2
    cd1.CancelError = True
    cd1.FileName = ""
    cd1.Filter = "*.csv|*.csv"
    cd1.ShowOpen
    If cd1.FileTitle <> "" Then
        
        
        
    End If
    
    AbrirFicheroYProcesarCoarval = ProcesaFicheroClientesCOARVAL(cd1.FileName, lblIndicador(4))
        
    
    lblIndicador(4).Caption = ""
    Exit Function
eAbrirFicheroYProcesarCoarval:
    If Err.Number <> 32755 Then MuestraError Err.Number, , Err.Description
    Err.Clear
End Function


Private Sub CargaColumnasCoarval(Errores As Boolean)
    lw(7).ListItems.Clear
    lw(7).ColumnHeaders.Clear
    If Errores Then
        lw(7).ColumnHeaders.Add , , "Err", 800
        lw(7).ColumnHeaders.Add , , "Descripción", 9250
        
        cmdImpFraCoarval.Caption = "Leer fich."
        lw(7).Tag = 0
        lblIndicador(4).Caption = ""
    Else
        lw(7).ColumnHeaders.Add , , "Serie", 800
        lw(7).ColumnHeaders.Add , , "Factura", 1300
        lw(7).ColumnHeaders.Add , , "Fecha", 1300
        lw(7).ColumnHeaders.Add , , "Cod.", 1100
        lw(7).ColumnHeaders.Add , , "Cliente", 3850
        lw(7).ColumnHeaders.Add , , "Base", 1300
        lw(7).ColumnHeaders(6).Alignment = lvwColumnRight
        lw(7).ColumnHeaders.Add , , "Total", 1450
        lw(7).ColumnHeaders(7).Alignment = lvwColumnRight
        lw(7).Tag = 1
    End If
End Sub



Private Sub CargaFacturasOK()
        miSQL = "select * FROM tmpintegracoarval where codusu  = " & vUsu.Codigo & " GROUP BY numserie,numfactu   ORDER BY numserie,numfactu  "
        miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            
            
            
            lw(7).ListItems.Add , , miRsAux!numSerie
            lw(7).ListItems(NumRegElim).SubItems(1) = Format(miRsAux!Numfactu, "000000")
            lw(7).ListItems(NumRegElim).SubItems(2) = Format(miRsAux!fechaalt, "dd/mm/yyyy")
            lw(7).ListItems(NumRegElim).SubItems(3) = Format(miRsAux!codClien, "000000")
            lw(7).ListItems(NumRegElim).SubItems(4) = miRsAux!NomClien
            lw(7).ListItems(NumRegElim).SubItems(5) = Format(miRsAux!Base, FormatoImporte)
            lw(7).ListItems(NumRegElim).SubItems(6) = Format(miRsAux!total, FormatoImporte)
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
End Sub


Private Function GeneraFraCli() As Boolean
Dim CodTraba As Long
Dim RF As ADODB.Recordset
Dim cad As String
Dim RIvas As ADODB.Recordset
 Dim IVA As Byte
 
 
    On Error GoTo eGeneraFraCli
    GeneraFraCli = False
    Set RF = New ADODB.Recordset
    Set RIvas = New ADODB.Recordset
    
    miSQL = PonerTrabajadorConectado(cad)
    If miSQL = "" Then Err.Raise 513, , "No se puede establecer el trabajador conectado"
    CodTraba = Val(miSQL)
    
    
    miSQL = DevuelveDesdeBD(conAri, "TiposIVA", "sparamcoarval", "1", "1", "T")   'porcentajes a tratar
    If miSQL = "" Then Err.Raise 513, , "parametros coarval. %Ivas"
    cadFormula = miSQL 'TiposIVA
    
    
    miSQL = DevuelveDesdeBD(conAri, "CodigoIVA", "sparamcoarval", "1", "1", "T")
    If miSQL = "" Then Err.Raise 513, , "parametros coarval. Ivas"
    cadSelect = miSQL 'CodigoIVA
    cad = Mid(miSQL, 2)
    
    miSQL = DevuelveDesdeBD(conAri, "CodigoIVARecargo", "sparamcoarval", "1", "1", "T")
    If miSQL = "" Then Err.Raise 513, , "parametros coarval. Ivas recargo"
    cadNomRPT = miSQL 'CodigoIVARecargo
    cad = cad & Mid(miSQL, 2)
    cad = Mid(cad, 1, Len(cad) - 1) 'quito el ulitmo pipe
    cad = Replace(cad, "|", ",")
    miSQL = "Select * from tiposiva WHERE codigiva in (" & cad & ") ORDER BY porceiva"
    
    RIvas.Open miSQL, ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    
    miSQL = "select * FROM tmpintegracoarval where codusu  = " & vUsu.Codigo & " GROUP BY numserie,numfactu   ORDER BY numserie,numfactu  "
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF

          'SCAFAC
                cad = "INSERT INTO scafac(codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,coddirec,"
                cad = cad & " codagent,codforpa,dtoppago,dtognral,brutofac,impdtopp,impdtogr,"
                cad = cad & " baseimp1,codigiv1,porciva1,imporiv1,porciva1re,imporiv1re,"
                cad = cad & " baseimp2,codigiv2,porciva2,imporiv2,porciva2re,imporiv2re,"
                cad = cad & " baseimp3,codigiv3,porciva3,imporiv3,porciva3re,imporiv3re,"
                cad = cad & "  TotalFac,intconta) VALUES ("
                    
                'numserie,numfactu,fechaalt,base,total,base_sr,iva_sr,re_sr,total_sr,base_red,iva_red,re_red,total_red,base_norm,iva_norm,re_nor,total_nor
                
                'codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,
                cad = cad & DBSet(miRsAux!numSerie, "T") & "," & miRsAux!Numfactu & "," & DBSet(miRsAux!fechaalt, "F") & "," & DBSet(miRsAux!codClien, "N") & ","
                cad = cad & DBSet(miRsAux!NomClien, "T") & "," & DBSet(miRsAux!domclien, "T") & "," & DBSet(miRsAux!codpobla, "T") & ","
                cad = cad & DBSet(miRsAux!pobclien, "T") & "," & DBSet(miRsAux!proclien, "T") & "," & DBSet(miRsAux!nifClien, "T") & ",NULL,"
                
                
                'codagent,codforpa,dtoppago,dtognral,brutofac,impdtopp,impdtogr
                miSQL = DevuelveDesdeBD(conAri, "codforpa", "sforpa", "nomforpa", miRsAux!ForPa, "T")
                If miSQL = "" Then Err.Raise 513, , "Error en forma de pago"
                cad = cad & vParamAplic.PorDefecto_Agente & "," & miSQL & ",0,0," & DBSet(Round2(miRsAux!Base, 2), "N") & ",0,0,"
                
                'BASE 1
                IVA = 0
                If Not IsNull(miRsAux!base_norm) Then
                    IVA = IVA + 1
                    miSQL = CadenaImportesIVA_Coarval(RIvas, 3)
                    cad = cad & miSQL
                End If
                
                If Not IsNull(miRsAux!base_red) Then
                    IVA = IVA + 1
                    miSQL = CadenaImportesIVA_Coarval(RIvas, 2)
                    cad = cad & miSQL
                End If
                
                If Not IsNull(miRsAux!base_sr) Then
                    IVA = IVA + 1
                    miSQL = CadenaImportesIVA_Coarval(RIvas, 1)
                    cad = cad & miSQL
                End If
                
                For numParam = IVA + 1 To 3
                    miSQL = "null,null,null,null,null,null,"
                    cad = cad & miSQL
                Next
                
                cad = cad & DBSet(miRsAux!total, "N") & ",1)"
                
                
                conn.Execute cad
               'SCAFAC1
                'vParamAplic.PorDefecto_Agente
                cad = "INSERT INTO scafac1(codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,codenvio,codtraba,codtrab1,codtrab2) VALUES ("
                cad = cad & DBSet(miRsAux!numSerie, "T") & "," & miRsAux!Numfactu & "," & DBSet(miRsAux!fechaalt, "F") & "," & DBSet(miRsAux!numSerie, "T") & ","
                cad = cad & miRsAux!Numfactu & "," & DBSet(miRsAux!fechaalt, "F") & "," & vParamAplic.PorDefecto_Envio & ","
                cad = cad & CodTraba & "," & CodTraba & "," & CodTraba & ")"
                conn.Execute cad
                
                'SVENCI
                cad = "INSERT INTO svenci(codtipom,numfactu,fecfactu,ordefect,fecefect,impefect) VALUES ("
                cad = cad & DBSet(miRsAux!numSerie, "T") & "," & miRsAux!Numfactu & "," & DBSet(miRsAux!fechaalt, "F") & ",1,"
                cad = cad & DBSet(miRsAux!fechaalt, "F") & "," & DBSet(miRsAux!total, "N") & ")"
                conn.Execute cad
                 
                'SLIFAC
                cad = "select * FROM tmpintegracoarval where codusu  = " & vUsu.Codigo & " and numserie=" & DBSet(miRsAux!numSerie, "T")
                cad = cad & " AND numfactu =" & miRsAux!Numfactu
                RF.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                IVA = 0
                miSQL = ""
                While Not RF.EOF
                    IVA = IVA + 1
                    'codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic"
                    ',cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre"
                    cad = DBSet(miRsAux!numSerie, "T") & "," & miRsAux!Numfactu & "," & DBSet(miRsAux!fechaalt, "F") & "," & DBSet(miRsAux!numSerie, "T") & ","
                    cad = cad & miRsAux!Numfactu & "," & IVA & ",1,"
                    'codartic nomartic ',cantidad,
                    cad = cad & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!NomArtic, "F") & "," & DBSet(miRsAux!cantidad, "N") & ","
                    'numbultos,precioar,dtoline1,dtoline2,importel,origpre"
                    cad = cad & DBSet(Int(miRsAux!cantidad), "N") & "," & DBSet(miRsAux!precioar, "N") & "," & DBSet(miRsAux!dtoline1, "N", "N") & ",0,"
                    cad = cad & DBSet(miRsAux!ImporteL, "N") & ",'T')"
                    miSQL = miSQL & ", (" & cad
                    RF.MoveNext
                Wend
                RF.Close
                If miSQL <> "" Then
                    miSQL = Mid(miSQL, 2)
                    cad = "insert into slifac(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic"
                    cad = cad & ",cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre) VALUES " & miSQL
                    conn.Execute cad
                End If
                        
                
                
                 
                miRsAux.MoveNext
    Wend
    miRsAux.Close
    GeneraFraCli = True
    Exit Function
eGeneraFraCli:
        MuestraError Err.Number, Err.Description
        Set RIvas = Nothing
        Set RF = Nothing
End Function



'3 NOrmal    2  Reducido   1 Super reducido
Private Function CadenaImportesIVA_Coarval(ByRef RsDeIVAs As ADODB.Recordset, Tipo As Byte) As String
Dim C As String
Dim ConRecargo As Boolean
Dim Indice As Currency
    'cadSelect    CodigoIVA
    ' cadNomRPT  CodigoIVARecargo
    ' cadFormula  % de ivas
    'PorcenIVA
    
    ConRecargo = False
    If Tipo = 3 Then
        If DBLet(miRsAux!re_nor, "N") <> 0 Then ConRecargo = True
        
    ElseIf Tipo = 2 Then
        If DBLet(miRsAux!re_red, "N") <> 0 Then ConRecargo = True
    Else
        If DBLet(miRsAux!re_sr, "N") <> 0 Then ConRecargo = True
    End If
    
    If ConRecargo Then
        C = RecuperaValor(cadNomRPT, CInt(Tipo) + 1) 'El primero es un pipe
    Else
        C = RecuperaValor(cadSelect, CInt(Tipo) + 1) 'El primero es un pipe
    End If
    RsDeIVAs.Find "codigiva= " & C, , adSearchForward, 1
    If RsDeIVAs.EOF Then
        C = "IVA " & Tipo & " -- Codigo " & C
        Err.Raise 513, , C
    End If
    '    baseimp1,codigiv1,porciva1,imporiv1,porciva1re,imporiv1re,"
    Select Case Tipo
    Case 1
        'base_sr  iva_sr   re_sr    total_sr
            C = DBSet(Round2(miRsAux!base_sr, 2), "N") & "," & RsDeIVAs!Codigiva & "," & DBSet(RsDeIVAs!PorceIVA, "N") & "," & DBSet(Round2(miRsAux!iva_sr, 2), "N") & ","
            If ConRecargo Then
                C = C & DBSet(RsDeIVAs!porcerec, "N") & "," & DBSet(miRsAux!iva_sr, "N")
            Else
                C = C & "null,null"
            End If
    Case 2
            'base_red iva_red  re_red   total_red
            C = DBSet(Round2(miRsAux!base_red, 2), "N") & "," & RsDeIVAs!Codigiva & "," & DBSet(RsDeIVAs!PorceIVA, "N") & "," & DBSet(Round2(miRsAux!iva_red, 2), "N") & ","
            If ConRecargo Then
                C = C & DBSet(RsDeIVAs!porcerec, "N") & "," & DBSet(miRsAux!iva_sr, "N")
            Else
                C = C & "null,null"
            End If
    Case 3
            'base_norm iva_norm  re_nor  total_nor
            C = DBSet(Round2(miRsAux!base_norm, 2), "N") & "," & RsDeIVAs!Codigiva & "," & DBSet(RsDeIVAs!PorceIVA, "N") & "," & DBSet(Round2(miRsAux!iva_norm, 2), "N") & ","
            If ConRecargo Then
                C = C & DBSet(RsDeIVAs!porcerec, "N") & "," & DBSet(miRsAux!iva_sr, "N")
            Else
                C = C & "null,null"
            End If
            
    End Select
    
    
    CadenaImportesIVA_Coarval = C & ","
    
End Function




Private Sub datosLineasAlbarEulerEspecial()
    Set miRsAux = New ADODB.Recordset
    lblDpto(37).visible = False
    lblDpto(37).Tag = 0
    Me.cmdAceptarLinEspEuler.Enabled = True
    If CadenaDesdeOtroForm <> "" Then
        Set miRsAux = New ADODB.Recordset
        
        
        If OpcionListado = 27 Then
            'cadPDFrpt = "codtipom numfactu fecfactu codtipoa numalbar"
            cadPDFrpt = " codtipom= " & DBSet(RecuperaValor(OtrosDatos, 1), "T")
            cadPDFrpt = cadPDFrpt & "  AND numfactu= " & DBSet(RecuperaValor(OtrosDatos, 2), "N")
            cadPDFrpt = cadPDFrpt & " AND fecfactu = " & DBSet(RecuperaValor(OtrosDatos, 3), "F")
            cadPDFrpt = cadPDFrpt & "  AND codtipoa= " & DBSet(RecuperaValor(OtrosDatos, 4), "T")
            cadPDFrpt = cadPDFrpt & "  AND numalbar= " & DBSet(RecuperaValor(OtrosDatos, 5), "N")
            cadPDFrpt = cadPDFrpt & " AND numlinea= " & CadenaDesdeOtroForm
            
            cadPDFrpt = "Select articulo ,descrarticulo ,cantidad ,precioar,dtoline1 ,importel,numlinea from slifac_eu2 WHERE " & cadPDFrpt
            
        ElseIf OpcionListado = 42 Then
            'Proeyectos
            
            cadPDFrpt = "  codtipom= " & DBSet(RecuperaValor(OtrosDatos, 1), "T")
            cadPDFrpt = cadPDFrpt & "  AND numproyec= " & DBSet(RecuperaValor(OtrosDatos, 2), "T")
            cadPDFrpt = cadPDFrpt & " AND numlinea= " & CadenaDesdeOtroForm
            
            cadPDFrpt = "Select articulo ,descrarticulo ,cantidad ,precioar,dtoline1 ,importel,numlinea from sproyectolin2 WHERE " & cadPDFrpt

            
            
        Else
            'En albaranes
            
             cadPDFrpt = "  codtipom= " & DBSet(RecuperaValor(OtrosDatos, 1), "T")
            cadPDFrpt = cadPDFrpt & "  AND numalbar= " & RecuperaValor(OtrosDatos, 2)
            cadPDFrpt = cadPDFrpt & " AND numlinea= " & CadenaDesdeOtroForm
            
            cadPDFrpt = "Select articulo ,descrarticulo ,cantidad ,precioar,dtoline1 ,importel,numlinea from slialb_eu2 WHERE " & cadPDFrpt
        End If
        miRsAux.Open cadPDFrpt, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            MsgBox "Linea no encontrada", vbExclamation
            Me.cmdAceptarLinEspEuler.Enabled = False
        Else
            txtModificable(4).Text = miRsAux!Articulo
            txtModificable(5).Text = miRsAux!descrarticulo
            txtNumero(5).Text = Format(miRsAux!cantidad, FormatoCantidad)
            txtNumero(6).Text = Format(miRsAux!precioar, FormatoPrecio)
            txtNumero(7).Text = Format(miRsAux!dtoline1, FormatoImporte)
            txtNumero(8).Text = Format(miRsAux!ImporteL, FormatoImporte)
            lblDpto(37).Tag = miRsAux!numlinea
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        CadenaDesdeOtroForm = ""
        lblDpto(37).visible = True
    Else
        txtModificable(4).Text = ""
        txtModificable(5).Text = ""
        txtNumero(5).Text = ""
        txtNumero(6).Text = ""
        txtNumero(7).Text = ""
        txtNumero(8).Text = ""
        
    End If
    
    
  
    PonerFoco txtModificable(4)
    
End Sub

Private Sub LeerMatriculaTaxco()
    Dim IT As ListItem
    Me.lw(8).ListItems.Clear
    If txtModificable(6).Text = "" Then Exit Sub

    
    Set miRsAux = New ADODB.Recordset
    cadParam = ""
    cadFormula = ""

      
    
    'Por si hay en albaranes
    miSQL = "SELECT scaalb.fechaalb,scaalb.numalbar,codclien,nomclien ,numrepar ,Observaciones  from scaalb,scaalb_eu e where scaalb.numalbar=e.numalbar and "
    miSQL = miSQL & " scaalb.codtipom=e.codtipom and e.codtipom='ALO'and bombamarca like '" & txtModificable(6).Text & "'"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        
            NumRegElim = NumRegElim + 1
            
            Set IT = lw(8).ListItems.Add(, , Format(miRsAux!FechaAlb, "dd/mm/yyyy"))
            
            IT.ToolTipText = "ALBARAN"
            
            
            IT.SubItems(1) = miRsAux!Numalbar
            
            
            IT.SubItems(2) = Format(miRsAux!codClien, "000000")
            IT.ListSubItems(2).ToolTipText = miRsAux!NomClien
            IT.SubItems(3) = miRsAux!numrepar
            IT.SubItems(4) = DBLet(miRsAux!Observaciones, "T") & " "
            IT.ListSubItems(4).ToolTipText = DBLet(miRsAux!Observaciones, "T")
            IT.SubItems(5) = Format(miRsAux!FechaAlb, "yyyymmdd") & Format(miRsAux!Numalbar, "00000") & "z"
            
            
            IT.Tag = "  scaalb.codtipom='ALO' AND scaalb.numalbar =" & miRsAux!Numalbar
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Por si hay en facturas
    miSQL = "SELECT t.fechaalb,e.numalbar,codclien,nomclien ,numrepar,Observaciones , s.numfactu,s.fecfactu,s.codtipom,e.codtipoa FROM   "
    miSQL = miSQL & " scafac s,scafac1 t,scafac_eu e WHERE"
    miSQL = miSQL & " s.codtipom = T.codtipom And s.numfactu = T.numfactu And s.FecFactu = T.FecFactu"
    miSQL = miSQL & " and s.codtipom = E.codtipom And s.codtipom = E.codtipom And s.FecFactu = E.FecFactu  and e.codtipoa =t.codtipoa and e.numalbar=t.numalbar"
    miSQL = miSQL & " AND e.codtipoa='ALO'   and bombamarca like '" & txtModificable(6).Text & "'"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF

            NumRegElim = NumRegElim + 1
            Set IT = lw(8).ListItems.Add(, , Format(miRsAux!FechaAlb, "dd/mm/yyyy"))
            IT.SubItems(1) = miRsAux!Numalbar
            IT.SubItems(2) = Format(miRsAux!codClien, "000000")
            IT.ListSubItems(2).ToolTipText = miRsAux!NomClien
            IT.SubItems(3) = DBLet(miRsAux!numrepar, "N")
            IT.SubItems(4) = DBLet(miRsAux!Observaciones, "T")
            IT.ListSubItems(4).ToolTipText = DBLet(miRsAux!Observaciones, "T")

            IT.SubItems(5) = Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!Numalbar, "00000") & "f"
    
            IT.Tag = "  scafac.codtipom=" & DBSet(miRsAux!codtipom, "T") & " AND scafac.numfactu=" & miRsAux!Numfactu & " AND scafac.fecfactu =" & DBSet(miRsAux!FecFactu, "F")
            IT.Tag = IT.Tag & " ## AND  scafac.codtipoa=" & DBSet(miRsAux!Codtipoa, "T") & " AND scafac.numalbar=" & miRsAux!Numalbar
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    
    
    
    'Desde la tabla traspada por TAXCO
    miSQL = "select smatriculataller.*,nomclien from smatriculataller inner join sclien on smatriculataller.codclien=sclien.codclien"
    miSQL = miSQL & " WHERE matricula like '" & txtModificable(6).Text & "'"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF

            NumRegElim = NumRegElim + 1
            Set IT = lw(8).ListItems.Add(, , Format("01/01/2001", "dd/mm/yyyy"))
            IT.SubItems(1) = "00000" 'miRsAux!Numalbar
            IT.SubItems(2) = Format(miRsAux!codClien, "000000")
            IT.ListSubItems(2).ToolTipText = miRsAux!NomClien
            IT.SubItems(3) = " "  'DBLet(miRsAux!numrepar, "N")
            IT.SubItems(4) = DBLet(miRsAux!marca_modelo, "T")
            IT.ListSubItems(4).ToolTipText = DBLet(miRsAux!marca_modelo, "T")

            IT.SubItems(5) = "0000000000000" & Format(NumRegElim, "0000")
    
            'IT.Tag = "  scafac.codtipom=" & DBSet(miRsAux!codtipom, "T") & " AND scafac.numfactu=" & miRsAux!Numfactu & " AND scafac.fecfactu =" & DBSet(miRsAux!FecFactu, "F")
            'IT.Tag = IT.Tag & " ## AND  scafac.codtipoa=" & DBSet(miRsAux!codtipoa, "T") & " AND scafac.numalbar=" & miRsAux!Numalbar
    
            IT.Tag = "smatriculataller.matricula = " & DBSet(miRsAux!Matricula, "T")
            IT.ForeColor = vbBlue
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    If NumRegElim > 1 Then
        Set lw(8).SelectedItem = lw(8).ListItems(1)
    Else
        If NumRegElim = 0 Then PonerFoco txtModificable(6)
    End If
    cadParam = ""
    cadFormula = ""

End Sub



'
Private Sub PonerImportesFormaPagoALVIC()
    Dim IT As ListItem
    Me.lw(9).ListItems.Clear
    

    Set miRsAux = New ADODB.Recordset
    
      
    
    'Por si hay en albaranes
    miSQL = "select codforpa,nomforpa,cantidad from tmpscapla ,sforpa where codusu=" & vUsu.Codigo & " and codplant=codforpa ORDER BY codforpa"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        
            NumRegElim = NumRegElim + 1
            
            
            
            Set IT = lw(9).ListItems.Add(, , Format(miRsAux!codforpa, "0000"))
            
            If miRsAux!codforpa = 2 Then
                'Credito. NO suma
                IT.SubItems(1) = miRsAux!nomforpa & "  " & Format(miRsAux!cantidad, FormatoImporte)
                IT.SubItems(2) = " "
                 IT.Tag = 0
                'El importe lo resto
                OtrosDatos = CCur(OtrosDatos) - miRsAux!cantidad
                Label9(28).Caption = "Importes traspaso ALVIC (" & Format(CCur(OtrosDatos), FormatoImporte) & ")"
            Else
                IT.SubItems(1) = miRsAux!nomforpa
                IT.SubItems(2) = Format(miRsAux!cantidad, FormatoImporte)
                IT.Tag = miRsAux!cantidad
            End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    'informacion adicional
    Me.lw(10).ListItems.Clear
    miSQL = "select mid(numalbaran,1,1) tipo , if(numfactura is null,'Albaran','Factura') facturado, count(*) cuanto,min(numalbaran) minimo,max(numalbaran) maximo"
    miSQL = miSQL & "  from tmpgasolimport where codusu =" & vUsu.Codigo & " group by 1,2 order by 1 asc , 2 asc"
    miRsAux.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    miSQL = ""
    cadFormula = ""
    While Not miRsAux.EOF
        If miRsAux!Tipo <> cadFormula Then
            If cadFormula <> "" Then
                AnyadeItemObservaTaxco
            
            End If
            
            'Es otro tipo
            cadFormula = miRsAux!Tipo
            
            miSQL = miRsAux!Tipo & "|"
           
        End If
        miSQL = miSQL & miRsAux!Facturado & "*" & miRsAux!cuanto & "*" & miRsAux!Minimo & "*" & miRsAux!Maximo & "*|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If cadFormula <> "" Then AnyadeItemObservaTaxco
    
    
    For numParam = 1 To 3
        cadFormula = RecuperaValor("GASOIL|GASOLINA|FICHA|", CInt(numParam))
        miSQL = "select CodigoProducto,min(codigo) linea,count(*) veces from tmpgasolimport where codusu =" & vUsu.Codigo
        miSQL = miSQL & " and  CodigoProducto like '%" & cadFormula & "%'"
        miSQL = miSQL & " and NOT CodigoProducto  in (select artculoAlvic from sarticalvic ) group by 1"
        
        miRsAux.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        miSQL = ""
        cadFormula = ""
        While Not miRsAux.EOF
            Set IT = lw(10).ListItems.Add(, , miRsAux!CodigoProducto)
            cadFormula = "Linea: " & miRsAux!linea & "    Veces: " & miRsAux!Veces
            IT.SubItems(1) = cadFormula
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    Next
    
    
    
    
    
    
    
    
    
    
    Set miRsAux = Nothing
    numParam = 0
    cadFormula = ""
    cadSelect = ""
    cmdAlvic.Enabled = NumRegElim > 0
End Sub


Private Sub AnyadeItemObservaTaxco()
Dim IT As ListItem

    
    
    cadSelect = RecuperaValor(miSQL, 2)
    cadSelect = Replace(cadSelect, "*", "|")
    If cadSelect <> "" Then
        cadNomRPT = "Tipo " & RecuperaValor(miSQL, 1)
        cadNomRPT = cadNomRPT & " - " & RecuperaValor(cadSelect, 1)
        
        Set IT = lw(10).ListItems.Add(, , cadNomRPT)
        
        cadNomRPT = "Nº " & Right("00000" & RecuperaValor(cadSelect, 2), 5) & "   Min: " & RecuperaValor(cadSelect, 3) & "    Max: " & RecuperaValor(cadSelect, 4)
        
        
        IT.SubItems(1) = cadNomRPT
    End If

    cadSelect = RecuperaValor(miSQL, 3)
    cadSelect = Replace(cadSelect, "*", "|")
    If cadSelect <> "" Then
        cadNomRPT = "Tipo " & RecuperaValor(miSQL, 1)
        cadNomRPT = cadNomRPT & " - " & RecuperaValor(cadSelect, 1)
        
        Set IT = lw(10).ListItems.Add(, , cadNomRPT)
        
        cadNomRPT = "Nº " & Right("00000" & RecuperaValor(cadSelect, 2), 5) '& "   Min: " & RecuperaValor(cadSelect, 3) & "    Max: " & RecuperaValor(cadSelect, 4)
        
        
        IT.SubItems(1) = cadNomRPT
    End If

End Sub



Private Sub CambiarImporteALvic()
Dim Importe As Currency
    If lw(9).SelectedItem Is Nothing Then Exit Sub
    If Val(lw(9).SelectedItem.Text) = 2 Then Exit Sub 'Credito NO se modifica
    miSQL = lw(9).SelectedItem.SubItems(1)
    Importe = lw(9).SelectedItem.Tag + lblDpto(42).Tag
    miSQL = InputBox(miSQL, "Ajustar importe", Importe)
    Importe = 0
    If miSQL <> "" Then
        cadFormula = ""
        If Not IsNumeric(miSQL) Then
            cadFormula = "Campo no numerico"
        Else
            Importe = CCur(TransformaPuntosComas(miSQL))
            
        End If
        If cadFormula <> "" Then
            MsgBox cadFormula, vbExclamation
        Else
            lw(9).SelectedItem.SubItems(2) = Format(Importe, FormatoImporte)
            lw(9).SelectedItem.Tag = Importe
            ComprobarImportesAlvic
        End If
    End If
End Sub


Private Sub ComprobarImportesAlvic()
Dim Importe As Currency

    Importe = 0
    For numParam = 1 To lw(9).ListItems.Count
        Importe = Importe + lw(9).ListItems(numParam).Tag
    Next
    Importe = CCur(OtrosDatos) - Importe
    If Importe <> 0 Then
        lblDpto(42).Caption = "Diferencia: " & Format(Importe, FormatoImporte)
        lblDpto(42).Tag = Importe
    Else
        lblDpto(42).Caption = ""
        lblDpto(42).Tag = 0
    End If
End Sub






Private Function HacerPrevisionAlvic() As Boolean
Dim Importe As Currency

    On Error GoTo eHacerPrevisionAlvic
    HacerPrevisionAlvic = False

    'Al hacer la prevision la primera vez vemos si existen las cuentas 4301.X para los clientes de credito
    


    'vAciamos
    auxiliar = "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
    conn.Execute auxiliar
    
    'SIEMPRE ES FACTURACION COLECTIVA
    auxiliar = IIf(Me.optVarios(9).Value, "ALD", "ALB")
    auxiliar = "UPDATE scaalb SET tipofact=0 WHERE codtipom='" & auxiliar & "' AND tipofact=1"
    ejecutar auxiliar, False
    
    
    InicializarVbles True
    auxiliar = " = '" & IIf(Me.optVarios(9).Value, "ALD", "ALB") & "'"
    
    cadSelect = "(scaalb.codtipom)" & auxiliar
    cadFormula = "{scaalb.codtipom}" & auxiliar
    cadPDFrpt = IIf(Me.optVarios(9).Value, "Gasolinera  ", "Tienda    ")
    If txtFecha(17).Text <> "" Or txtFecha(18).Text <> "" Then
        miSQL = " Fecha: "
        If Not PonerDesdeHasta("{scaalb.fechaalb}", "F", 17, 18, miSQL) Then Exit Function
        cadPDFrpt = cadPDFrpt & miSQL
    End If

    If txtCliente(12).Text <> "" Or txtCliente(13).Text <> "" Then
        miSQL = " Cliente: "
        If Not PonerDesdeHasta("{scaalb.codclien}", "CLI", 12, 13, miSQL) Then Exit Function
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "      "
        cadPDFrpt = cadPDFrpt & miSQL
        

    End If
    

    cadParam = cadParam & "pDH=""" & cadPDFrpt & """|"
    numParam = numParam + 1
    
    
    Screen.MousePointer = vbHourglass
    If Not HayRegParaInforme("scaalb", cadSelect, False) Then Exit Function
            
    
    
    
    
    
    
    
    
    
    
    auxiliar = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,nombre3,importe1,importe2)"
    auxiliar = auxiliar & " select " & vUsu.Codigo & " ,codclien,codforpa,codartic,nomclien,nomartic,sum(cantidad),sum(importel)"
    auxiliar = auxiliar & " from  scaalb,slialb where  scaalb.codtipom=slialb.codtipom and "
    auxiliar = auxiliar & " scaalb.numalbar=slialb.numalbar and " & cadSelect
    auxiliar = auxiliar & " group by codclien,codforpa,codartic"
    conn.Execute auxiliar



    Set miRsAux = New ADODB.Recordset
    auxiliar = "select distinct nombre1 from tmpinformes WHERE codusu =" & vUsu.Codigo
    miRsAux.Open auxiliar, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        auxiliar = miRsAux!nombre1
        auxiliar = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", auxiliar, "T")
        
        auxiliar = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", auxiliar)
        If auxiliar = "" Then Err.Raise 513, , "IVA incoerrecto" & miRsAux!Codigiva

        
        Importe = CCur(auxiliar)
        auxiliar = " importe3= " & DBSet(Importe, "N")
        auxiliar = auxiliar & ", importe4= round(importe2 * " & DBSet((Importe / 100), "N") & ",2) "
        auxiliar = "UPDATE tmpinformes SET " & auxiliar & " WHERE codusu =" & vUsu.Codigo
        auxiliar = auxiliar & " AND nombre1 = " & DBSet(miRsAux!nombre1, "T")
        conn.Execute auxiliar
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Espera 0.5
    auxiliar = "UPDATE tmpinformes SET importe5=importe2+importe4 WHERE codusu =" & vUsu.Codigo
    conn.Execute auxiliar

    HacerPrevisionAlvic = True
eHacerPrevisionAlvic:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Function





Private Function HacerPrevisionCuentas() As Boolean
Dim Importe As Currency

    On Error GoTo eHacerPrevisionCuentas
    HacerPrevisionCuentas = False

    'Al hacer la prevision la primera vez vemos si existen las cuentas 4301.X para los clientes de credito
    

    Set miRsAux = New ADODB.Recordset

    'vAciamos
    auxiliar = "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
    conn.Execute auxiliar
    
    'SIEMPRE ES FACTURACION COLECTIVA
    
    auxiliar = "Select scaalb.codclien,sclien.nomclien,codmacta from scaalb,sclien where scaalb.codclien=sclien.codclien"
    auxiliar = auxiliar & " AND  scaalb.codforpa=2 AND codtipom = '" & IIf(Me.optVarios(9).Value, "ALD", "ALB")
    auxiliar = auxiliar & "' GROUP BY codclien ORDER BY codclien"
    miRsAux.Open auxiliar, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    auxiliar = ""
    miSQL = ""
    While Not miRsAux.EOF
        'codusu,codigo1,nombre1,nombre2
        miSQL = miRsAux!Codmacta
        miSQL = Mid(miSQL, 1, 3) & "1" & Mid(miSQL, 5)
        auxiliar = auxiliar & ", (" & vUsu.Codigo & "," & miRsAux!codClien & "," & DBSet(miRsAux!NomClien, "T")
        auxiliar = auxiliar & "," & DBSet(miSQL, "T") & ")"
        
        
        
        If Len(auxiliar) > 5000 Then
            auxiliar = Mid(auxiliar, 2)
            conn.Execute "INSERT INTO tmpinformes(codusu,codigo1,nombre1,nombre2) VALUES " & auxiliar
            auxiliar = ""
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Len(auxiliar) > 0 Then
            auxiliar = Mid(auxiliar, 2)
            conn.Execute "INSERT INTO tmpinformes(codusu,codigo1,nombre1,nombre2) VALUES " & auxiliar
    End If


    auxiliar = "delete  from tmpinformes where codusu = " & vUsu.Codigo
    auxiliar = auxiliar & " and nombre2  IN (select codmacta from ariconta" & vParamAplic.NumeroConta & ".cuentas where apudirec='S' and codmacta like '4301%')"
    conn.Execute auxiliar
    Espera 0.5
    
    auxiliar = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", CStr(vUsu.Codigo), "T")
    
    If Val(auxiliar) > 0 Then
        HacerPrevisionCuentas = True
    Else
        MsgBox "Comprobacion finalizada", vbInformation
    End If
eHacerPrevisionCuentas:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Function









'***************************************************************************************************************************************
'***************************************************************************************************************************************
'   Crear pedido proveedor desde pedidos clientes
'
'***************************************************************************************************************************************

Private Function CarcaLineasPedidoCliente() As Boolean


    On Error GoTo eCarcaLineasPedidoCliente
    CarcaLineasPedidoCliente = False

    'Al hacer la prevision la primera vez vemos si existen las cuentas 4301.X para los clientes de credito
    Set miRsAux = New ADODB.Recordset
    miSQL = " select sartic.codprove,nomprove,sliped.codartic,sliped.nomartic,ampliaci,precioar,dtoline1,dtoline2,importel,numlinea,cantidad"
    miSQL = miSQL & " from sliped left join sartic on sliped.codartic=sartic.codartic"
    miSQL = miSQL & " left join sprove on sartic.codprove=sprove.codprove"
    miSQL = miSQL & " Where artvario = 1"
    miSQL = miSQL & " and numpedcl=" & OtrosDatos
    miSQL = miSQL & " order by NUMLINEA"
    lw(11).ListItems.Clear
    NumRegElim = 0
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
        NumRegElim = NumRegElim + 1
        lw(11).ListItems.Add , "L" & Format(miRsAux!numlinea, "00000"), miRsAux!codArtic
        lw(11).ListItems(NumRegElim).SubItems(1) = CStr(miRsAux!NomArtic)
        lw(11).ListItems(NumRegElim).ToolTipText = DBLet(miRsAux!Ampliaci, "T")
        lw(11).ListItems(NumRegElim).ListSubItems(1).ToolTipText = DBLet(miRsAux!Ampliaci, "T")
        lw(11).ListItems(NumRegElim).SubItems(2) = Format(miRsAux!cantidad, FormatoCantidad)
        lw(11).ListItems(NumRegElim).SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
        lw(11).ListItems(NumRegElim).SubItems(4) = Format(miRsAux!dtoline1, FormatoCantidad)
        lw(11).ListItems(NumRegElim).SubItems(5) = Format(miRsAux!dtoline2, FormatoCantidad)
        lw(11).ListItems(NumRegElim).SubItems(6) = Format(miRsAux!ImporteL, FormatoCantidad)
        lw(11).ListItems(NumRegElim).SubItems(7) = Format(miRsAux!Codprove, "0000")
        lw(11).ListItems(NumRegElim).SubItems(8) = DBLet(miRsAux!nomprove, "T") & " "
        lw(11).ListItems(NumRegElim).Checked = True
        
       
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
'
'    For NumRegElim = 1 To lw(11).ColumnHeaders.Count
'        Debug.Print NumRegElim & " " & lw(11).ColumnHeaders(NumRegElim).Text & " : " & lw(11).ColumnHeaders(NumRegElim).Width
'    Next
eCarcaLineasPedidoCliente:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Function






'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************
'
'       Fontenas. Vista previa pedidos
'
'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************

Private Sub CargaPedidosConLw()
Dim IT As ListItem
    

    'Al hacer la prevision la primera vez vemos si existen las cuentas 4301.X para los clientes de credito
    Set miRsAux = New ADODB.Recordset
    lw(12).ListItems.Clear
    Set lw(12).SmallIcons = Me.imglistPed
    
    NumRegElim = 0
    miRsAux.Open OtrosDatos, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        miSQL = "P"
        numParam = 2
        If DBLet(miRsAux!Estado, "N") = 0 Then
            miSQL = " "
            numParam = 1
        Else
            If miRsAux!Estado = 2 Then miSQL = "*": numParam = 3
        End If
        NumRegElim = NumRegElim + 1
        cadParam = "C" & Format(miRsAux!NumPedcl, "00000")
        Set IT = lw(12).ListItems.Add(, CStr(cadParam))
        
        With IT
            .Text = miSQL
            .SubItems(1) = DBLet(miRsAux!Descripcion, "T") & " "
            .SubItems(2) = Format(miRsAux!NumPedcl, "000000")
            .SubItems(3) = Format(miRsAux!fecpedcl, "dd/mm/yyyy")
            .SubItems(4) = Format(miRsAux!codClien, "000000")
            .SubItems(5) = CStr(miRsAux!NomClien)
            .SubItems(6) = Format(miRsAux!fecpedcl, "yyyymmdd") & Format(miRsAux!NumPedcl, "000000")
            .SubItems(7) = "L" & miSQL & Format(miRsAux!NumPedcl, "000000")
            .SubItems(8) = Format(DBLet(miRsAux!prioridad, "T"), "000") & Format(miRsAux!NumPedcl, "000000")
            .SmallIcon = numParam
        End With
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    CadenaDesdeOtroForm = ""
End Sub



Private Sub OrdenacionPedidosFontenas(columna As Integer)

'    Caption = columna & "    :- " & lw(12).SortKey
'    For NumRegElim = 1 To lw(12).ColumnHeaders.Count
'        lw(12).ColumnHeaders(NumRegElim).Width = 1800
'    Next
    columna = columna - 1
    If columna = 1 Then columna = 8
    If columna = 3 Then columna = 6
    If columna = 0 Then columna = 7
    If lw(12).SortKey = columna Then
        If lw(12).SortOrder = lvwAscending Then
            lw(12).SortOrder = lvwDescending
        Else
            lw(12).SortOrder = lvwAscending
        End If
    Else
        lw(12).SortOrder = lvwAscending
        lw(12).SortKey = columna
    End If
       
End Sub


Private Sub CargaAlbaranesNif_Alvic(SoloLimpiar As Boolean)
Dim IT
    
    lw(13).ListItems.Clear
    If SoloLimpiar Then Exit Sub
    'Set lw(12).SmallIcons = Me.imglistPed
        
    'Para el cliente puesto, vamos a leer el nif y la forma de pago
    ' nota: hay un devuelvedesdebd2 que devuleve 2 +1 campos
    
    cadParam = "codforpa"
    auxiliar = "codtipom IN ('ALB','ALD') AND codclien "
    auxiliar = DevuelveDesdeBD(conAri, "nifclien", "scaalb", auxiliar, txtCliente(14).Text, "N", cadParam)
    If auxiliar = "" Then
        MsgBox "Ninguna albaran pendiente de facturar para ese cliente", vbExclamation
        Exit Sub
    End If
    Me.txtDescClie(14).Text = Me.txtDescClie(14).Text & " (" & auxiliar & ")            "
    
    miSQL = "SELECT scaalb.codtipom,scaalb.numalbar,fechaalb,codclien,nomclien,scaalb.codforpa,nomforpa,codartic,nomartic,precoste from scaalb,slialb,sforpa  where"
    miSQL = miSQL & " scaalb.codtipom = slialb.codtipom And scaalb.Numalbar = slialb.Numalbar And scaalb.codforpa = sforpa.codforpa"
    miSQL = miSQL & " and scaalb.codtipom IN ('ALB','ALD') and nifclien=" & DBSet(auxiliar, "T") & " AND codclien<>" & txtCliente(14).Text
    'Y los articulos estan en la tabla de traspaso ALVIC
    miSQL = miSQL & " AND codartic in (select codartic from sarticalvic)"
    
    miSQL = miSQL & " order by fechaalb,codtipom,numalbar"
    
    
    
    
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 0
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    miSQL = ""
    While Not miRsAux.EOF
        
        cadSelect = miRsAux!codtipom & Format(miRsAux!Numalbar, "000000")
        If cadSelect <> miSQL Then
            'Ok, insertamos
            miSQL = cadSelect
        Else
            
            'NO insertamos
            cadSelect = ""
        End If
        
        If cadSelect <> "" Then
            Set IT = lw(13).ListItems.Add(, cadSelect)
            
            With IT
                .Text = miRsAux!codtipom
                .SubItems(1) = Format(miRsAux!Numalbar, "000000")
                .SubItems(2) = Format(miRsAux!FechaAlb, "yyyy-mm-dd")
                .SubItems(3) = Format(miRsAux!codClien, "000000")
                .ListSubItems(3).ToolTipText = miRsAux!NomClien
                .SubItems(4) = Format(miRsAux!codforpa, "00")
                If Val(miRsAux!codforpa) = Val(cadParam) Then
                    .Tag = 0
                Else
                    
                    '.Tag = 1
                    .Tag = 0
                    
                    '.Bold = True
                    .ListSubItems(3).ForeColor = vbRed
                    '.ForeColor = vbRed
                End If
                .ListSubItems(4).ToolTipText = miRsAux!nomforpa
                .SubItems(5) = miRsAux!codArtic
                .SubItems(6) = miRsAux!NomArtic
                .SubItems(7) = Right(Space(12) & Format(miRsAux!precoste, FormatoImporte), 12)
                
            End With
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    cadSelect = ""
    Set miRsAux = Nothing


End Sub

Private Sub OrdenacionTaxcoCliente(queColumna As Integer)
'    For NumRegElim = 1 To lw(13).ColumnHeaders.Count
'        Debug.Print lw(13).ColumnHeaders(NumRegElim).Text & ": " & lw(13).ColumnHeaders(NumRegElim).Width
'    Next

    queColumna = queColumna - 1
    If queColumna = lw(13).SortKey Then
        If lw(13).SortOrder = lvwAscending Then
            lw(13).SortOrder = lvwDescending
        Else
            lw(13).SortOrder = lvwAscending
        End If
    Else
      
        lw(13).SortOrder = lvwAscending
        lw(13).SortKey = queColumna
    End If
End Sub


Private Function realizarCambioClienteTaxco() As Boolean

On Error GoTo eRealizarCambioClienteTaxco

    realizarCambioClienteTaxco = False

    'De momento SOLO CAMBIO codclien
    cadParam = ""

    For NumRegElim = 1 To lw(13).ListItems.Count
        
        If lw(13).ListItems(NumRegElim).Tag = 0 Then
            If lw(13).ListItems(NumRegElim).Checked Then
                'Para el LOG
                cadParam = cadParam & vbCrLf & lw(13).ListItems(NumRegElim).Text & lw(13).ListItems(NumRegElim).SubItems(1) & " " & Format(lw(13).ListItems(NumRegElim).SubItems(2), "dd/mm/yyyy") & " " & lw(13).ListItems(NumRegElim).SubItems(3) & " " & lw(13).ListItems(NumRegElim).SubItems(5)
                miSQL = "UPDATE scaalb set codclien = " & Me.txtCliente(14).Text & " WHERE codtipom ='" & lw(13).ListItems(NumRegElim).Text & "' AND numalbar ="
                miSQL = miSQL & lw(13).ListItems(NumRegElim).SubItems(1)
                conn.Execute miSQL
            End If
        End If
            
    Next
    
    realizarCambioClienteTaxco = True
    miSQL = Trim(txtCliente(14).Text & " - " & Me.txtDescClie(14).Text) & vbCrLf
    cadParam = "[Cambio albaran]" & vbCrLf & miSQL & cadParam
    Set LOG = New cLOG
    LOG.Insertar 39, vUsu, cadParam
    Set LOG = Nothing


eRealizarCambioClienteTaxco:
    
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
        
    

End Function


Private Sub CambiReferenciaCliente()

    Screen.MousePointer = vbHourglass
    cmdCambioCliente.Enabled = False
    
    CambiReferenciaCliente2
       
    lblIndicador(5).Caption = ""
    Screen.MousePointer = vbDefault
End Sub


Private Sub CambiReferenciaCliente2()
Dim IT
    
    On Error GoTo eCambiReferenciaCliente
    lw(14).ListItems.Clear
    If Me.txtCliente(15).Text = "" Then Exit Sub
    If Me.txtDescClie(15).Text = "" Then Exit Sub
    
    'Set lw(12).SmallIcons = Me.imglistPed
    
    miSQL = "sclienmani|sdirec|sdirenvio|aguacontadores|"
    cadSelect = "Carnet manipulador|Departamentos / diecciones|Direccciones envio|Contador de agua|"
    'Veremos si puede traspasra
    For NumRegElim = 1 To 4
       
        auxiliar = RecuperaValor(miSQL, CInt(NumRegElim))
        lblIndicador(5).Caption = auxiliar
        lblIndicador(5).Refresh
        If NumRegElim = 4 Then If vParamAplic.AguasPotables <> "" Then auxiliar = ""
        
        
        If auxiliar <> "" Then
            auxiliar = DevuelveDesdeBD(conAri, "count(*)", auxiliar, "codclien", txtCliente(15).Text, "T")
            If auxiliar = "" Then auxiliar = "0"
            
            If Val(auxiliar) > 0 Then
                Set IT = lw(14).ListItems.Add(, "T" & NumRegElim)
                IT.Text = RecuperaValor(cadSelect, CInt(NumRegElim))
                IT.SubItems(1) = Format(Val(auxiliar), "000000")
            End If
        End If
        
    Next
        
    If lw(14).ListItems.Count > 0 Then
        MsgBox "El programa no puede cambiar los datos de un cliente a otro", vbExclamation
        Exit Sub
    End If
            
           
    'Si que puede  |
    'Cargaremos todos los datos
    cadSelect = "advpartes| sactuaobra |scaalb |scaavi |scafac |scafre  |"
    cadSelect = cadSelect & "scaman| scamana |scaped |scapre |scarep |schalb|schped| "
    cadSelect = cadSelect & "schpre| schrep |scrmobsclien  |sdtofm|"
    cadSelect = cadSelect & "sgaste| spree1 |sprees |sserie   |sserlin |sclientfno|"
    cadSelect = cadSelect & "sclienrenting |scliendp|"  'hasta aqui 25
    numParam = 25
    
    
    cadParam = "partes adv| actuaciones obra |albaranes|avisos|facturas|frecuencias |"
    cadParam = cadParam & "mantenimientos| Mtos anulados |pedidos|ofertas |Reparaciones |Hco albarnes|Hco pedidos| "
    cadParam = cadParam & "Hco ofertas| Hco reparaciones |Crm |Dtos familia marca|"
    cadParam = cadParam & "Gastos tecnico| Hoc precios especiales |Precios especiales |Nº Serie|Linas nº serie|Telefonos|"
    cadParam = cadParam & "Renting |Dpto contactos|"  'hasta aqui 25
    
    
    For NumRegElim = 1 To numParam
       
       
        auxiliar = RecuperaValor(cadSelect, CInt(NumRegElim))
        lblIndicador(5).Caption = auxiliar
        lblIndicador(5).Refresh
        
        
        
        auxiliar = DevuelveDesdeBD(conAri, "count(*)", auxiliar, "codclien", txtCliente(15).Text, "T")
        If auxiliar = "" Then auxiliar = "0"
        
        If Val(auxiliar) > 0 Then
            Set IT = lw(14).ListItems.Add(, "T" & NumRegElim)
            
            IT.Text = RecuperaValor(cadParam, CInt(NumRegElim))
            IT.SubItems(1) = Format(Val(auxiliar), "000000")
        End If
    
        
    Next
        
    ' Veremos si
        
    
    If lw(14).ListItems.Count > 0 Then cmdCambioCliente.Enabled = True
    
    
    
    'Cobros pendientes
    lblIndicador(5).Caption = "Cobros pendientes"
    lblIndicador(5).Refresh
    
    
    
    auxiliar = DevuelveDesdeBD(conAri, "codmacta", "sclien", "codclien", txtCliente(15).Text, "T")
    If auxiliar = "" Then Err.Raise 513, , "Cuenta contable cliente vacia"
    
    miSQL = "situacion = 0 AND codmacta"
    auxiliar = DevuelveDesdeBD(conAri, "count(*)", "ariconta" & vParamAplic.NumeroConta & ".cobros", miSQL, auxiliar, "T")
    If Val(auxiliar) > 0 Then
        Set IT = lw(14).ListItems.Add(, "C1")
        IT.Text = "Cobros pendientes"
        IT.SubItems(1) = Format(Val(auxiliar), "000000")
    End If
    
    Exit Sub
eCambiReferenciaCliente:
    MuestraError Err.Number, , Err.Description
    cmdCambioCliente.Enabled = False
    lblIndicador(5).Caption = ""
End Sub



Private Function RealizaUpdatesCambioCliente() As Boolean

    On Error GoTo eRealizaUpdatesCambioCliente
    RealizaUpdatesCambioCliente = False
    cadSelect = "advpartes| sactuaobra |scaalb |scaavi |scafac |scafre  |"
    cadSelect = cadSelect & "scaman| scamana |scaped |scapre |scarep |schalb|schped| "
    cadSelect = cadSelect & "schpre| schrep |scrmobsclien  |sdtofm|"
    cadSelect = cadSelect & "sgaste| spree1 |sprees |sserie   |sserlin |sclientfno|"
    cadSelect = cadSelect & "sclienrenting |scliendp|"  'hasta aqui 25
    numParam = 25
    
    
    cadParam = "partes adv| actuaciones obra |albaranes|avisos|facturas|frecuencias |"
    cadParam = cadParam & "mantenimientos| Mtos anulados |pedidos|ofertas |Reparaciones |Hco albarnes|Hco pedidos| "
    cadParam = cadParam & "Hco ofertas| Hco reparaciones |Crm |Dtos familia marca|"
    cadParam = cadParam & "Gastos tecnico| Hoc precios especiales |Precios especiales |Nº Serie|Linas nº serie|Telefonos|"
    cadParam = cadParam & "Renting |Dpto contactos|"  'hasta aqui 25
                                'Abril 2021     QUitamos taximetros
    
    For NumRegElim = 1 To numParam
       
       
        auxiliar = RecuperaValor(cadSelect, CInt(NumRegElim))
        lblIndicador(5).Caption = auxiliar
        lblIndicador(5).Refresh
        
        
        miSQL = "UPDATE " & auxiliar & " SET codclien =" & txtCliente(16).Text
        miSQL = miSQL & " WHERE codclien =" & txtCliente(15).Text
        
        conn.Execute miSQL
        
    Next
    
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        'Modificaremos en fac_elec
        cadSelect = DevuelveDesdeBD(conAri, "i_d", "facelec2_ariadna.cliente", "codclien_ariges", txtCliente(15).Text, "N")
        If cadSelect = "" Then
            'MsgBox "No se encuentra el cliente en el sistema de facturación electrónica. ", vbExclamation
        Else
            cadParam = "UPDATE facelec2_ariadna.cliente SET  codclien_ariges = " & txtCliente(16).Text
            cadParam = cadParam & " WHERE i_d =" & cadSelect
            If Not ejecutar(cadParam, False) Then MsgBox "Proceso correcto. Falta facturacion electronica. Avise soporte técnico", vbExclamation
        End If
    End If
    
    RealizaUpdatesCambioCliente = True
    Exit Function
eRealizaUpdatesCambioCliente:
    MuestraError Err.Number, , Err.Description
End Function



Private Function ListadoPedidoPorDia() As Boolean
    
    On Error GoTo eListadoPedidoPorDia
    ListadoPedidoPorDia = False

    Set miRsAux = New ADODB.Recordset
    
    lblIndicador(6).Caption = "Prepara datos"
    lblIndicador(6).Refresh
    '
    cadSelect = " TRUE "
    If txtFecha(19).Text <> "" Then cadSelect = cadSelect & " AND fecpedcl >= " & DBSet(txtFecha(19).Text, "F")
    If txtFecha(20).Text <> "" Then cadSelect = cadSelect & " AND fecpedcl <= " & DBSet(txtFecha(20).Text, "F")
    
    'vAciamos
    auxiliar = "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
    conn.Execute auxiliar
    
    
    NumRegElim = 0
    cadNomRPT = "INSERT INTO tmpinformes(codusu,codigo1,campo1,fecha1,importe1,nombre1,nombre2,nombre3,fecha2,fecha3) VALUES "
    For numParam = 1 To 3
        cadTitulo = RecuperaValor("Pedidos|Albaranes|Facturas|", CInt(numParam))
        lblIndicador(6).Caption = cadTitulo
        lblIndicador(6).Refresh
        
        If numParam = 1 Then
            auxiliar = "select numpedcl,fecpedcl,codclien,nomclien,'' alb,'' fra,null fechaalb,null fecfac from scaped WHERE " & cadSelect
        ElseIf numParam = 2 Then
            auxiliar = "select numpedcl,fecpedcl,codclien,nomclien,concat(codtipom,Lpad(numalbar ,7,""0"")) alb,'' fra,fechaalb ,null fecfac from scaalb where " & cadSelect
        Else
            auxiliar = "select numpedcl,fecpedcl,codclien,nomclien,concat(codtipoa,Lpad(numalbar ,7,""0"")) alb"
            auxiliar = auxiliar & " ,concat(scafac.codtipom,Lpad(scafac.numfactu ,7,""0"")) fra, fechaalb,scafac.fecfactu fecfac  FROM scafac,scafac1  "
            auxiliar = auxiliar & " where scafac.codtipom=scafac1.codtipom and scafac.numfactu=scafac1.numfactu and scafac.fecfactu=scafac1.fecfactu"
            auxiliar = auxiliar & " AND " & cadSelect
        End If
        
        
        miRsAux.Open auxiliar, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cadPDFrpt = ""
        While Not miRsAux.EOF
            NumRegElim = NumRegElim + 1
            'codusu,codigo1,campo1,fecha1,importe1,nombre1,nombre2,nombre3,fecha2,fecha3
            miSQL = ", (" & vUsu.Codigo & "," & NumRegElim & "," & miRsAux!NumPedcl & "," & DBSet(miRsAux!fecpedcl, "F") & ","
            miSQL = miSQL & miRsAux!codClien & "," & DBSet(miRsAux!NomClien, "T") & ","
            'Albara, factura
            miSQL = miSQL & DBSet(miRsAux!alb, "T", "S") & "," & DBSet(miRsAux!fra, "T", "S") & ","
            'fec Albara, fec factura
            miSQL = miSQL & DBSet(miRsAux!FechaAlb, "F", "S") & "," & DBSet(miRsAux!fecFac, "F", "S") & ")"
            cadPDFrpt = cadPDFrpt & miSQL
            
            If Len(cadPDFrpt) > 5000 Then
                
                lblIndicador(6).Caption = cadTitulo & " " & NumRegElim
                lblIndicador(6).Refresh
                cadPDFrpt = Mid(cadPDFrpt, 2)
                cadPDFrpt = cadNomRPT & cadPDFrpt
                conn.Execute cadPDFrpt
                cadPDFrpt = ""
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If Len(cadPDFrpt) > 0 Then
            cadPDFrpt = Mid(cadPDFrpt, 2)
            cadPDFrpt = cadNomRPT & cadPDFrpt
            conn.Execute cadPDFrpt
        End If
    Next
    
    
    If NumRegElim > 0 Then ListadoPedidoPorDia = True
    
eListadoPedidoPorDia:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    lblIndicador(6).Caption = ""
    Set miRsAux = Nothing
End Function


Private Sub txtProve_GotFocus(Index As Integer)
    ConseguirFoco txtProve(Index), 3
End Sub

Private Sub txtProve_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgProveedor_Click Index
    End If
End Sub

Private Sub txtProve_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub



Private Sub txtProve_LostFocus(Index As Integer)
Dim Descri As String

    Descri = ""
    txtProve(Index).Text = Trim(txtProve(Index).Text)
    If txtProve(Index).Text <> "" Then
        If PonerFormatoEntero(txtProve(Index)) Then
            Descri = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtProve(Index).Text, "N")
            If Descri = "" Then MsgBox "No existe el proveedor : " & txtProve(Index).Text, vbExclamation
        End If
    End If
    txtDescProve(Index).Text = Descri
    If Descri = "" And txtProve(Index).Text <> "" Then
        txtProve(Index).Text = ""
        PonerFoco txtProve(Index)
    End If
End Sub





Private Function GeneraDatosDtoComparativo() As Boolean
    
    On Error GoTo eGeneraDatosDtoComparativo
    GeneraDatosDtoComparativo = False
    Screen.MousePointer = vbHourglass
    lblIndicador(7).Caption = "Preparando datos"
    lblIndicador(7).Refresh
    
    conn.Execute "DELETE from tmpinformes WHERE codusu =" & vUsu.Codigo
    
    cadSelect = "sfamia.codfamia >=0 "
    If txtProve(0).Text <> "" Or txtProve(1).Text <> "" Then
        miSQL = " Proveedor: "
        cadTitulo = "{#elprov#.codprove}"
        If Not PonerDesdeHasta(cadTitulo, "PRO", 0, 1, miSQL) Then Exit Function
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & "              "
        cadPDFrpt = cadPDFrpt & miSQL
        
        
    End If
    
    
    If txtFamia(1).Text <> "" Or txtFamia(2).Text <> "" Then
        miSQL = " Familia: "
        cadTitulo = "{sfamia.codfamia}"
        If Not PonerDesdeHasta(cadTitulo, "FAM", 1, 2, miSQL) Then Exit Function
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & """ + chr(13) + """
        cadPDFrpt = cadPDFrpt & miSQL
            
    End If
    
    
    'ACtividad
    auxiliar = ""
    If txtActiv(2).Text <> "" Or txtActiv(3).Text <> "" Then
        miSQL = ""
        If txtActiv(2).Text <> "" Then
            miSQL = " desde " & txtActiv(2).Text & " " & Me.txtDescActiv(2).Text
            auxiliar = " AND sdtofm.codactiv>=" & txtActiv(2).Text
        End If
        
        If txtActiv(3).Text <> "" Then
            miSQL = miSQL & " hasta " & txtActiv(3).Text & " " & Me.txtDescActiv(3).Text
            auxiliar = auxiliar & " AND sdtofm.codactiv<=" & txtActiv(3).Text
        End If
        If cadPDFrpt <> "" Then cadPDFrpt = cadPDFrpt & """ + chr(13) + """
        miSQL = " Actividad: " & miSQL
        cadPDFrpt = cadPDFrpt & miSQL
        
        
        
    End If
    
    
    
    
    
   


    
    'Dto venta
    lblIndicador(7).Caption = "Dto venta activ"
    lblIndicador(7).Refresh
    Set miRsAux = New ADODB.Recordset
    
    cadTitulo = Replace(cadSelect, "{", "")
    cadSelect = Replace(cadTitulo, "}", "")
    miSQL = Replace(cadSelect, "#elprov#", "sfamia")
    
   
    miSQL = " FROM sdtofm,sfamia,(SELECT @rownum:=0) r where sdtofm.codfamia=sfamia.codfamia AND " & miSQL
    miSQL = miSQL & " AND codactiv >=0" & auxiliar  'axuliar lleva la actividad
    miSQL = "Select " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum,codprove,sfamia.codfamia,concat('A',codactiv),null,dtoline1,dtoline2,0,0,0,0,0 " & miSQL
    miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,importe1,importe2,importeb1,importeb2,importeb3,importeb4,importeb5) " & miSQL
    conn.Execute miSQL
    
    Espera 0.1
    
    
    
    miSQL = "Select nombre1 from tmpinformes where codusu =" & vUsu.Codigo & " GROUP by nombre1"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    miSQL = ""
    While Not miRsAux.EOF
        miSQL = miSQL & ", " & Mid(miRsAux!nombre1, 2)
        miRsAux.MoveNext
        NumRegElim = NumRegElim + 1
    Wend
    miRsAux.Close
    If NumRegElim > 1 Then
        AnchoLogin = ""
        frmMensajes.OpcionMensaje = 27
        frmMensajes.cadWhere = miSQL
        frmMensajes.Show vbModal
        
        If CadenaDesdeOtroForm = "NO" Then Exit Function
        
        If CadenaDesdeOtroForm <> "IG" Then
            'Hay que borarar las actividades que ns dicen
            lblIndicador(7).Caption = "Ajuste actividad"
            lblIndicador(7).Refresh
            CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2) 'quitamos la primera coma
            miSQL = "(" & CadenaDesdeOtroForm & ")"
            miSQL = "DELETE from tmpinformes where codusu = " & vUsu.Codigo & " AND nombre1 in " & miSQL
            conn.Execute miSQL
            
            
            cadPDFrpt = cadPDFrpt & " [ACTIVIDADES]: " & AnchoLogin
            
            
        End If
        AnchoLogin = ""
    End If
    
    cadParam = cadParam & "DesdeHasta=""" & cadPDFrpt & """|"
    numParam = numParam + 1
 
    
    'Updateamos con la actividad
    Set miRsAux = New ADODB.Recordset
    lblIndicador(7).Caption = "Dto venta activ"
    lblIndicador(7).Refresh
    NumRegElim = 0
    miSQL = "Select nombre1,max(codigo1) Maximo from tmpinformes where codusu =" & vUsu.Codigo & " GROUP by nombre1"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
    
            lblIndicador(7).Caption = "Dto venta act " & miRsAux!nombre1
            lblIndicador(7).Refresh
    
            If NumRegElim < miRsAux!Maximo Then NumRegElim = miRsAux!Maximo
            miSQL = Mid(miRsAux!nombre1, 2)
            miSQL = DevuelveDesdeBD(conAri, "nomactiv", "sactiv", "codactiv", miSQL)
            If miSQL = "" Then
                miSQL = "ERROR Acti: " & miRsAux!nombre1
            Else
                miSQL = Right("00000" & Mid(miRsAux!nombre1, 2), 5) & " " & miSQL
            End If
            
            miSQL = "UPDATE tmpinformes set nombre1=null, nombre2=" & DBSet(miSQL, "T")
            miSQL = miSQL & " WHERE nombre1='" & miRsAux!nombre1 & "' AND codusu =" & vUsu.Codigo
            conn.Execute miSQL
            miRsAux.MoveNext
            
    Wend
    miRsAux.Close
    
    'Si no pone actividad añade clientes
    If txtActiv(2).Text = "" And txtActiv(3).Text = "" Then
         lblIndicador(7).Caption = "Dto venta cliente"
         lblIndicador(7).Refresh
         NumRegElim = 0
         miSQL = Replace(cadSelect, "#elprov#", "sfamia")
         
        
         miSQL = " FROM sdtofm,sfamia,(SELECT @rownum:=0) r where sdtofm.codfamia=sfamia.codfamia AND " & miSQL
         miSQL = miSQL & " AND codclien >=0"
         miSQL = "Select " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum,codprove,sfamia.codfamia,concat('C',codclien),'',dtoline1,dtoline2,0,0,0,0,0 " & miSQL
         miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,importe1,importe2,importeb1,importeb2,importeb3,importeb4,importeb5) " & miSQL
         conn.Execute miSQL
         
        Espera 0.1
    
    End If
    
    'Updateamos el cliente
    
    miSQL = "Select nombre1,max(codigo1) Maximo from tmpinformes where nombre1<>'' and codusu =" & vUsu.Codigo & " GROUP by nombre1"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
            lblIndicador(7).Caption = "Dto venta cli " & miRsAux!nombre1
            lblIndicador(7).Refresh
            
            If NumRegElim < miRsAux!Maximo Then NumRegElim = miRsAux!Maximo
           
            miSQL = Mid(miRsAux!nombre1, 2)
            miSQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", miSQL)
            If miSQL = "" Then
                miSQL = "ERROR cli: " & miRsAux!nombre1
            Else
                miSQL = Right("00000" & Mid(miRsAux!nombre1, 2), 5) & " " & miSQL
            End If
            miSQL = "UPDATE tmpinformes set nombre1=" & DBSet(miSQL, "T")
            miSQL = miSQL & " WHERE nombre1='" & miRsAux!nombre1 & "' AND codusu =" & vUsu.Codigo
            conn.Execute miSQL
            miRsAux.MoveNext
            
    Wend
    miRsAux.Close
    
    
    
    
    lblIndicador(7).Caption = "Dto proveedor familia "
    lblIndicador(7).Refresh
    
'    miSQL = "Select campo1,campo2 from tmpinformes where codusu =" & vUsu.Codigo & " GROUP by campo1,campo2"
'    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    cad = ""
'    While Not miRsAux.EOF
'        ad = ad & ", (" & miRsAux!campo1 & "," & miRsAux!campo2 & ")"
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'
    
    
    
    
    'select * from sdtomp where codfamia >=0
    miSQL = Replace(cadSelect, "#elprov#", "sdtomp")
    
    
    miSQL = " FROM sdtomp,sfamia where sdtomp.codfamia=sfamia.codfamia AND " & miSQL
    miSQL = " coalesce(rap1,0) rap1,coalesce(rap2,0) rap2, coalesce(dtosincargo,0) dtosincargo " & miSQL
    miSQL = "Select sdtomp.codprove,sfamia.codfamia,coalesce(dtoline1,0) dtoline1,coalesce(dtoline2,0) dtoline2," & miSQL
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
   ' miSQL = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,importeb1,importeb2,importeb3,importeb4,importeb5) " & miSQL
    
    
    
    While Not miRsAux.EOF
        
        lblIndicador(7).Caption = "Dto proveedor familia " & miRsAux!Codprove & " / " & miRsAux!Codfamia
        lblIndicador(7).Refresh
        miSQL = "UPDATE tmpinformes   SET importeb1 = " & DBSet(miRsAux!dtoline1, "N", "N")
        miSQL = miSQL & ", importeb2 = " & DBSet(miRsAux!dtoline2, "N", "N") & ", importeb3 = " & DBSet(miRsAux!Rap1, "N", "N")
        miSQL = miSQL & ", importeb4 = " & DBSet(miRsAux!Rap2, "N", "N") & ", importeb5 = " & DBSet(miRsAux!dtosincargo, "N", "N")
        miSQL = miSQL & " WHERE codusu =" & vUsu.Codigo & " AND campo1 =" & miRsAux!Codprove & " AND campo2=" & miRsAux!Codfamia
        conn.Execute miSQL
        miRsAux.MoveNext
            
    Wend
    miRsAux.Close
    
    lblIndicador(7).Caption = "Familia"
    lblIndicador(7).Refresh
    
    miSQL = "Select campo2 from tmpinformes where codusu =" & vUsu.Codigo & " GROUP by campo2"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
            lblIndicador(7).Caption = "Fam " & miRsAux!campo2
            lblIndicador(7).Refresh
            
        
            
            
            miSQL = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", miRsAux!campo2)
            If miSQL = "" Then miSQL = "ERROR fam: " & miRsAux!campo2
            
            miSQL = "UPDATE tmpinformes set nombre3=" & DBSet(miSQL, "T")
            miSQL = miSQL & " WHERE campo2=" & miRsAux!campo2 & " AND codusu =" & vUsu.Codigo
            conn.Execute miSQL
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Hacemos el caculo
    lblIndicador(7).Caption = "Calculo 0"
    lblIndicador(7).Refresh
    'dtos. de venta, dtos. compra, rappel y % s/cargo y que calcule el beneficio. (100 - dto venta =       100 - dto compras - rappel - s/cargo =
    miSQL = "UPDATE tmpinformes   SET importe3 = 100 - (importe1 + importe2)"
    miSQL = miSQL & " , importe4=100"
    miSQL = miSQL & " WHERE codusu =" & vUsu.Codigo
    conn.Execute miSQL
    
    'Herbelca. Los descuentos NO son "aditivos"
    'importe4 = 100 - (importeb1+importeb2+importeb3+importeb4+importeb5)"
    'luego la suma va sobre resto cada vez
    Dim J As Integer
    cadTitulo = "importeb1|importeb2|importeb3|importeb4|importeb5|"
    For J = 1 To 5
        lblIndicador(7).Caption = "Calculo " & J
        lblIndicador(7).Refresh
        miSQL = "UPDATE tmpinformes   SET importe4=importe4 * ((100 - " & RecuperaValor(cadTitulo, J) & ")/100)"
        miSQL = miSQL & " WHERE codusu =" & vUsu.Codigo
        miSQL = miSQL & " AND " & RecuperaValor(cadTitulo, J) & " > 0"
        conn.Execute miSQL
        
         '   if  not isnull({sdtomp.dtoline1}) then     Resultado := Resultado * ((100 - {sdtomp.dtoline1})/100)  ;
         '   if  not  isnull({sdtomp.dtoline2}) then     Resultado := Resultado * ((100 - {sdtomp.dtoline2})/100) ;
         '   if  not  isnull({sdtomp.rap1}) then     Resultado := Resultado * ((100 - {sdtomp.rap1})/100) ;
         '   if  not  isnull({sdtomp.rap2}) then     Resultado := Resultado * ((100 - {sdtomp.rap2})/100) ;
    
    
    Next J
    
    GeneraDatosDtoComparativo = (NumRegElim > 0)
    
    
eGeneraDatosDtoComparativo:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    lblIndicador(7).Caption = ""
    Screen.MousePointer = vbDefault
End Function


Private Sub CargaDatosPreCosteArtVario()
    
        
    Me.txtNumero(9).Text = CadenaDesdeOtroForm
    ' Mid(txtAux(1).Text & Space(16), 1, 16) & Mid(txtAux(2).Text & Space(50), 1, 50) & Text2(16).Text
    txtNoModificable(6).Text = Trim(Mid(Me.OtrosDatos, 1, 16))
    txtNoModificable(7).Text = Trim(Mid(OtrosDatos, 17, 50))
    txtNoModificable(9).Text = Trim(Mid(OtrosDatos, 67, 14))
    CadenaDesdeOtroForm = Trim(Mid(OtrosDatos, 81))
    txtNoModificable(8).Text = CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    PonerFoco txtNumero(9)
End Sub



Private Function HacerProcesoPrecioMinimo() As Boolean
    
    HacerProcesoPrecioMinimo = False
    
    
    cadSelect = ""
    If txtFecha(23).Text = "" Then
        cadSelect = "Fecha obligatoria"
    Else
        If Me.chkVarios(9).Value = 0 Then
            If txtFecha(24).Text = "" Then cadSelect = "Fechas promocion obligatorias"
        End If
    End If
    
    If cadSelect <> "" Then
        MsgBox cadSelect, vbExclamation
        Exit Function
    End If
    
    
    
 
    
    If Me.chkVarios(9).Value = 0 Then
        cadSelect = "fechaini =" & DBSet(txtFecha(23).Text, "F") & " AND fechafin=" & DBSet(txtFecha(24).Text, "F")
        cadTitulo = "Fechas promocion. Inicio: " & txtFecha(23).Text & " Fin: " & txtFecha(24).Text
    Else
        cadSelect = "fechafin <=" & DBSet(txtFecha(23).Text, "F")
        cadTitulo = "Fechas promocion menor igual: " & txtFecha(23).Text
    End If
    
    cadParam = DevuelveDesdeBD(conAri, "count(*)", "spromo", cadSelect & " AND 1", "1")
    If Val(cadParam) = 0 Then
        MsgBox "Ningun registro a actualizar", vbExclamation
        Exit Function
    End If
    NumRegElim = Val(cadParam)
    If Me.optVarios(13).Value Then
    
        If MsgBox("Proceso irreversible. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    
    
        cadTitulo = "[PROMO] Eliminar promocion" & vbCrLf & "Lineas:  " & NumRegElim & vbCrLf & cadTitulo
        CadenaDesdeOtroForm = "Va a borrar " & NumRegElim & " registros de promociones  "
        
        cadFormula = "DELETE FROM spromo where " & cadSelect
        
    Else
        cadTitulo = "[PMV]" & IIf(Me.optVarios(11).Value, "borrar", "establecer") & " precio minimo venta en articulos" & vbCrLf & "Lineas:  " & NumRegElim & vbCrLf & cadTitulo
        CadenaDesdeOtroForm = "Va a " & IIf(Me.optVarios(12).Value, "borrar", "establecer") & " el precio minimo de venta de : " & NumRegElim & " articulos"
        
        cadFormula = 0
        If Me.optVarios(11).Value Then cadFormula = " spromo.precioac"
        cadFormula = "UPDATE sartic,spromo  set preciominvta =" & cadFormula
        cadFormula = cadFormula & " WHERE sartic.codartic=spromo.codartic and " & cadSelect
        
    End If
    
    cadParam = cadParam & vbCrLf & vbCrLf & "¿SEGURO que desea continuar?"
    If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    

    
    
    
    Screen.MousePointer = vbHourglass
    If ejecutar(cadFormula, False) Then
    
        Set LOG = New cLOG
        LOG.Insertar 29, vUsu, cadTitulo
        Set LOG = Nothing
        
        
        MsgBox "Proceso realizado con éxito", vbInformation
        
        HacerProcesoPrecioMinimo = True
    End If

        

    NumRegElim = 0
    
    
    
End Function





Private Sub CargaDatosAlbaranesEulerVinculacion()

Dim Valora As String
Dim N As Node
Dim IT

    
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    cadFormula = ""
    Me.treeAlb(0).Nodes.Clear
    Me.treeAlb(1).Nodes.Clear
    
    lblIndicador(8).Caption = "Leyendo ppal"
    lblIndicador(8).Refresh
    miSQL = "select codtipoa,numalbar from sproyectolin where "
    miSQL = miSQL & " codtipom='" & RecuperaValor(OtrosDatos, 2) & "' AND numproyec=" & RecuperaValor(OtrosDatos, 3) & " AND ppal =1 "
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then cadFormula = miRsAux!Codtipoa & Format(miRsAux!Numalbar, "000000")
    miRsAux.Close
    
    
    
    lblIndicador(8).Caption = "Leyendo vinculados"
    lblIndicador(8).Refresh
    miSQL = "select scaalb.codtipom,scaalb.numalbar,fechaalb,referenc,codartic,nomartic,cantidad,importel,numlinea,numpedcl "
    miSQL = miSQL & " from scaalb left join slialb on scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar"
    miSQL = miSQL & " WHERE codClien = " & RecuperaValor(OtrosDatos, 1) & " AND  (scaalb.codtipom,scaalb.numalbar) IN"
    miSQL = miSQL & "          (select codtipoa,numalbar from sproyectolin where "
    '                       codtipom,numproyec
    miSQL = miSQL & " codtipom='" & RecuperaValor(OtrosDatos, 2) & "' AND numproyec=" & RecuperaValor(OtrosDatos, 3) & ")"
    miSQL = miSQL & " ORDER BY slialb.codtipom,slialb.numalbar,numlinea"
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Valora = ""
    While Not miRsAux.EOF
        miSQL = miRsAux!codtipom & Format(miRsAux!Numalbar, "000000")
        If Valora <> miSQL Then
            'Insertamos el padre
            Valora = CStr(miSQL)
            miSQL = Mid(miSQL & "        ", 1, 15) & Format(miRsAux!FechaAlb, "dd/mm/yyyy") & "    Ref.: " & DBLet(miRsAux!referenc, "T")
            Set N = treeAlb(0).Nodes.Add(, , Valora, miSQL)
            N.Tag = "('" & RecuperaValor(OtrosDatos, 2) & "'," & RecuperaValor(OtrosDatos, 3) & ",'" & miRsAux!codtipom & "'," & miRsAux!Numalbar & ")"
            N.Checked = True
            If treeAlb(0).Nodes.Count < 2 Then N.EnsureVisible
            
            
            If Valora = cadFormula Then
                N.Bold = True
                N.BackColor = vbGreen
                N.Text = Replace(N.Text, " Ref.: ", " *Ref.:")
            End If
        End If
        miSQL = Mid(miRsAux!codArtic & Space(20), 1, 20) & Mid(miRsAux!NomArtic & Space(40), 1, 40)
        miSQL = miSQL & Right(Space(7) & Format(miRsAux!cantidad, FormatoCantidad), 7)
        miSQL = miSQL & Right(Space(14) & Format(miRsAux!ImporteL, FormatoCantidad), 14)
        Set N = treeAlb(0).Nodes.Add(Valora, tvwChild, , miSQL)
        N.Checked = True
        
        miRsAux.MoveNext

    Wend
    miRsAux.Close


    'Pendientes
    lblIndicador(8).Caption = "Leyendo vinculados"
    lblIndicador(8).Refresh
    miSQL = "select slialb.codtipom,slialb.numalbar,fechaalb,referenc,codartic,nomartic,cantidad,importel,numlinea "
    miSQL = miSQL & " from scaalb inner join slialb on scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar"
    miSQL = miSQL & " WHERE codClien = " & RecuperaValor(OtrosDatos, 1) & " AND  not (slialb.codtipom,slialb.numalbar) in"
    'SQL = SQL & " (select codtipoa,numalbar,numlinea from sproyectolin WHERE " & DevWHERE & ")"
    miSQL = miSQL & " (select codtipoa,numalbar from sproyectolin )"

    miSQL = miSQL & " ORDER BY slialb.codtipom,slialb.numalbar,numlinea"

    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Valora = ""
    While Not miRsAux.EOF
        miSQL = miRsAux!codtipom & Format(miRsAux!Numalbar, "000000")
        If Valora <> miSQL Then
            'Insertamos el padre
            Valora = CStr(miSQL)
            miSQL = Mid(miSQL & "        ", 1, 15) & Format(miRsAux!FechaAlb, "dd/mm/yyyy") & "    Ref: " & DBLet(miRsAux!referenc, "T")
            Set N = treeAlb(1).Nodes.Add(, , Valora, miSQL)
            N.Tag = "('" & miRsAux!codtipom & "'," & miRsAux!Numalbar & ")"
            If treeAlb(1).Nodes.Count < 2 Then N.EnsureVisible
        End If
        miSQL = Mid(miRsAux!codArtic & Space(20), 1, 20) & Mid(miRsAux!NomArtic & Space(40), 1, 40)
        miSQL = miSQL & Right(Space(7) & Format(miRsAux!cantidad, FormatoCantidad), 7)
        miSQL = miSQL & Right(Space(14) & Format(miRsAux!ImporteL, FormatoCantidad), 14)
        Set N = treeAlb(1).Nodes.Add(Valora, tvwChild, , miSQL)
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    lblIndicador(8).Caption = ""
    
    Me.cmdEstablecerAlbaranPrincipal(0).visible = cadFormula <> ""
    Me.cmdEstablecerAlbaranPrincipal(1).visible = cadFormula = ""
    
    cadFormula = ""
    Set miRsAux = Nothing
End Sub

Private Function ModificarAlbaranesVinculados() As Boolean
Dim QuedaalgunAlbaran As Boolean
Dim HahcechoPreg As Boolean
Dim HaCambiado As Boolean
Dim BorrarEnlaceAlbaranPpal As Boolean

On Error GoTo eModificarAlbaranesVinculados

    ModificarAlbaranesVinculados = False



    'Comprobacion
    HaCambiado = False
    BorrarEnlaceAlbaranPpal = False
    cadTitulo = "" 'albaran ppal
    If Me.cmdEstablecerAlbaranPrincipal(0).visible Then
        'Tiene albaran principal o tenia
        'Vemos que esta marcado
        miSQL = ""
        cadSelect = ""
        For NumRegElim = 1 To Me.treeAlb(0).Nodes.Count
            If treeAlb(0).Nodes(NumRegElim).Parent Is Nothing Then
                If treeAlb(0).Nodes(NumRegElim).Bold Then
                    cadTitulo = treeAlb(0).Nodes(NumRegElim)  'albaran principal
                    If InStr(1, treeAlb(0).Nodes(NumRegElim).Text, " *Ref.:") > 0 Then
                        'Es el mismo NODO que el que entro
                        cadSelect = "OK"
                    Else
                        miSQL = "Ha cambiado de albarán principal"
                        cadTitulo = treeAlb(0).Nodes(NumRegElim).Tag
                    End If
                    If Not treeAlb(0).Nodes(NumRegElim).Checked Then
                        MsgBox "Albarán principal NO seleccionado", vbExclamation
                        Exit Function
                    End If
                End If
            End If
        Next
        
        If cadSelect = "" Then
            'NO hay ningun albarán principal seleccionado
            If miSQL = "" Then miSQL = "No hay ningun albaran principal seleccionado"
            miSQL = miSQL & "   ¿Continuar?"
            If MsgBox(miSQL, vbExclamation + vbYesNoCancel) <> vbYes Then Exit Function
            BorrarEnlaceAlbaranPpal = True
            HaCambiado = True
        End If
    Else
        
        
        For NumRegElim = 1 To treeAlb(1).Nodes.Count
            If treeAlb(1).Nodes(NumRegElim).Parent Is Nothing Then
                If treeAlb(1).Nodes(NumRegElim).Bold And Not treeAlb(1).Nodes(NumRegElim).Checked Then
                    MsgBox "Albarán principal NO seleccionado", vbExclamation
                    Exit Function
                End If
            End If
        Next
    End If
    








    lblIndicador(8).Caption = "Actualizando"
    lblIndicador(8).Refresh
    cadSelect = ""  'quitar marca sactuaobra
    cadParam = ""   'Ponermarca sactuaobra
    miSQL = ""
    QuedaalgunAlbaran = False
    For NumRegElim = 1 To Me.treeAlb(0).Nodes.Count
        If treeAlb(0).Nodes(NumRegElim).Parent Is Nothing Then
            'Si ha quitado la marca
             If Not treeAlb(0).Nodes(NumRegElim).Checked Then
                
                'Si ha quitado la marca
                miSQL = miSQL & ", " & treeAlb(0).Nodes(NumRegElim).Tag
                                                
                cadFormula = "'" & Mid(treeAlb(0).Nodes(NumRegElim).Text, 1, 3) & "'," & Mid(treeAlb(0).Nodes(NumRegElim).Text, 4, 7)
                cadSelect = cadSelect & ", (" & cadFormula & ")"
            Else
                QuedaalgunAlbaran = True
            End If

        End If
    Next
    If miSQL <> "" Then
        If Not HahcechoPreg Then
            If MsgBox("¿Desea realizar las modificaciones?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
            HahcechoPreg = True
        End If

        miSQL = Mid(miSQL, 2)
        miSQL = "DELETE FROM sproyectolin WHERE (codtipom,numproyec,codtipoa,numalbar) IN (" & miSQL & ")"
        conn.Execute miSQL
        HaCambiado = True
    End If



    miSQL = ""
    cadNomRPT = ""
    For NumRegElim = 1 To treeAlb(1).Nodes.Count
        If treeAlb(1).Nodes(NumRegElim).Parent Is Nothing Then
            If treeAlb(1).Nodes(NumRegElim).Checked Then
                'Ppal
                If treeAlb(1).Nodes(NumRegElim).Bold Then cadNomRPT = treeAlb(1).Nodes(NumRegElim).Tag
                'Si ha quitado la marca
                miSQL = miSQL & ", " & treeAlb(1).Nodes(NumRegElim).Tag
                
                cadFormula = "'" & Mid(treeAlb(1).Nodes(NumRegElim).Text, 1, 3) & "'," & Mid(treeAlb(1).Nodes(NumRegElim).Text, 4, 7)
                cadParam = cadParam & ", (" & cadFormula & ")"
    
                QuedaalgunAlbaran = True 'queda pq se van a insertar
            End If
        End If
    Next
    If miSQL <> "" Then
        If Not HahcechoPreg Then
            If MsgBox("¿Desea realizar las modificaciones?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
            HahcechoPreg = True
        End If

        
        miSQL = Mid(miSQL, 2)
        miSQL = " WHERE (codtipom,numalbar) IN (" & miSQL & ")"
        miSQL = "SELECT '" & RecuperaValor(OtrosDatos, 2) & "'," & RecuperaValor(OtrosDatos, 3) & ",codtipom,numalbar FROM slialb" & miSQL
        miSQL = "INSERT IGNORE INTO sproyectolin(codtipom,numproyec,codtipoa,numalbar)  " & miSQL
        conn.Execute miSQL
        
        HaCambiado = True
    End If

    If HaCambiado Then
        
        If cadSelect <> "" Then
            cadSelect = Mid(cadSelect, 2)
            'Ha quitado albaranes. Les quito la marca
            miSQL = "UPDATE scaalb set actuacion =NULL WHERE (codtipom,numalbar) IN (" & cadSelect & ")"
            If Not ejecutar(miSQL, False) Then MsgBox "Error quitando marca proyecto(actuacion) en albaranes: " & cadSelect, vbExclamation
        End If
        If cadParam <> "" Then
            'Primero creamos la actuacion (si no existe)
            cadParam = Mid(cadParam, 2)
            cadTitulo = "[" & RecuperaValor(Me.OtrosDatos, 2) & RecuperaValor(Me.OtrosDatos, 3) & "]"
            cadSelect = "INSERT IGNORE INTO sactuaobra(codclien,coddirec,actuacion,fechaini,observa)"
            cadSelect = cadSelect & " SELECT codclien,null coddirec ," & DBSet(cadTitulo, "T") & " actuacion ,fecproyec fechaini,"
            cadSelect = cadSelect & " coalesce(referenc,concat('Proyecto ',numproyec)) from sproyecto"
            cadSelect = cadSelect & " WHERE codtipom = '" & RecuperaValor(Me.OtrosDatos, 2)
            cadSelect = cadSelect & "' AND  numproyec= " & RecuperaValor(Me.OtrosDatos, 3)
            
            ejecutar cadSelect, False
            
            miSQL = "UPDATE scaalb set actuacion =" & DBSet(cadTitulo, "T") & " WHERE (codtipom,numalbar) IN (" & cadParam & ")"
            If Not ejecutar(miSQL, False) Then MsgBox "Error poniendo marca proyecto(actuacion) en albaranes: " & cadParam, vbExclamation
            
            
        End If
        
        
        'Ha cambiado cosas, y NO hay ninugn albaran seleccionado: BORRO
        If Not QuedaalgunAlbaran Then
            cadSelect = "DELETE FROM sactuaobra WHERE codclien =" & RecuperaValor(Me.OtrosDatos, 1) & " AND coddirec is NULL"
            cadSelect = cadSelect & " AND actuacion=" & DBSet(cadTitulo, "T")
            If Not ejecutar(cadSelect, False) Then MsgBox "Error eliminando  actuacion(proyecto): " & cadTitulo, vbExclamation
        End If
        
        
        
        'Si hay ppal, lo marco como tal
        If cadNomRPT <> "" Then
            miSQL = " WHERE codtipom='" & RecuperaValor(OtrosDatos, 2) & "' AND numproyec =" & RecuperaValor(OtrosDatos, 3)
            cadSelect = "UPDATE sproyectolin SET ppal=0 " & miSQL
            conn.Execute cadSelect
            
            cadSelect = "UPDATE sproyectolin SET ppal=1 " & miSQL
            cadSelect = cadSelect & " AND (codtipoa,numalbar) IN (" & cadNomRPT & ")"
            conn.Execute cadSelect
            
            If Me.cmdEstablecerAlbaranPrincipal(1).visible Then
                'Esta agregando el nuevo ALBARAN , que adenmas es PPAL
                cadSelect = "INSERT INTO sproyectolin2(numproyec,codtipom,numlinea,articulo,descrarticulo,cantidad,precioar,dtoline1,importel)"
                cadSelect = cadSelect & " SELECT " & RecuperaValor(OtrosDatos, 3) & " numpro, '" & RecuperaValor(OtrosDatos, 2) & "' tipom"
                cadSelect = cadSelect & " ,numlinea,codartic,coalesce(nomartic, if(ampliaci is null, '',coalesce( char(13),ampliaci))) descr"
                cadSelect = cadSelect & " ,cantidad ,precioar,dtoline1 + dtoline2,importel  FROM slialb WHERE "
                cadSelect = cadSelect & " (codtipom,numalbar) IN (" & cadNomRPT & ")"
                conn.Execute cadSelect
                
            End If
        Else
            'NO hay ppal.
            'Si habia lo desmarco
            If BorrarEnlaceAlbaranPpal Then
                miSQL = " WHERE codtipom='" & RecuperaValor(OtrosDatos, 2) & "' AND numproyec =" & RecuperaValor(OtrosDatos, 3)
                cadSelect = "UPDATE sproyectolin SET ppal=0 " & miSQL
                conn.Execute cadSelect
            End If
        End If
        
        'FALTAR añadir las entradas ppal como pidieron reunion Bittor y Jose
        If BorrarEnlaceAlbaranPpal Then
             miSQL = " DELETE FROM sproyectolin2 WHERE numlinea<1000 and codtipom='" & RecuperaValor(OtrosDatos, 2) & "' AND numproyec =" & RecuperaValor(OtrosDatos, 3)
            conn.Execute miSQL
        End If
        
        
    End If
    ModificarAlbaranesVinculados = True
    
    Exit Function
eModificarAlbaranesVinculados:
    MuestraError Err.Number, , "Consulte soporte técnico" & vbCrLf & Err.Description
End Function




Private Sub CargaDatosCestaUsuario()
  Dim Item As ListItem
 On Error GoTo eCargaDatosCestaUsuario
        
        'El RS viene cargado desde el PEDIDO CLIENTE
        'Solo hay que recorrer y mosotrar
        
        'select cestaLineaId,cestas_lineas.cestaId,numlinea,cestas_lineas.codartic,nomartic,cantidad,codusu,fecha,codclien  from cestas inner join
        'cestas_lineas on cestas_lineas.cestaId =cestas.cestaId
        'left join sartic on cestas_lineas.codartic=sartic.codartic
        
        
        While Not miRsAux.EOF
            If Me.lw(15).ListItems.Count = 0 Then
                'Es el primero
                Label9(43).Caption = Label9(43).Caption & " Cli: " & miRsAux!codClien & CadenaDesdeOtroForm
                
            End If
            Set Item = Me.lw(15).ListItems.Add(, "K" & miRsAux!cestaLineaId)
            Item.Text = Format(miRsAux!numlinea, "0000")
            Item.SubItems(1) = miRsAux!codArtic
            Item.SubItems(2) = miRsAux!NomArtic
            Item.SubItems(3) = Format(miRsAux!cantidad, FormatoCantidad)
            Item.SubItems(4) = Format(miRsAux!CanStock, FormatoCantidad)
            Item.Tag = miRsAux!cestaId
            Item.Checked = True
            If DBLet(miRsAux!CanStock, "N") <= 0 Then
                Item.ListSubItems(4).ForeColor = vbRed
                Item.ListSubItems(4).Bold = True
            End If
            
            miRsAux.MoveNext
        Wend
        
        CadenaDesdeOtroForm = ""
eCargaDatosCestaUsuario:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
   
End Sub



Private Sub CargaFechasPropuestasCambioFechaTaxco()
    
    miSQL = "select * from slog where accion=29 and descripcion like '[ALVIC_FP]%' ORDER BY fecha desc"
    
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then
        miSQL = "Imposible situar fecha inicio."
        cadFormula = ""
    Else
        
        cadFormula = DBLet(miRsAux!Descripcion, "T")
        NumRegElim = InStr(1, cadFormula, "Fin:")
        If NumRegElim = 0 Then
            cadFormula = ""
        Else
            cadFormula = Trim(Mid(cadFormula, NumRegElim + 4))
        End If
        If cadFormula = "" Then
            miSQL = "NO localiza FFin en : " & miRsAux!Descripcion
        Else
            If Not IsDate(cadFormula) Then
                miSQL = "No es campo fecha valido"
                cadFormula = ""
            Else
                cadFormula = Format(DateAdd("d", 1, CDate(cadFormula)), "dd/mm/yyyy")
            End If
        End If
    End If
    miRsAux.Close
    
    If cadFormula = "" Then
        Me.cmdAjusteVtosFaccliALVIC.Enabled = False
        MsgBox miSQL, vbExclamation
    Else
        txtFecha(25).Text = cadFormula
        BloquearTxt txtFecha(25), True
        Me.cmdAjusteVtosFaccliALVIC.Enabled = True
    End If
    Set miRsAux = Nothing
    
End Sub



Private Function AjusteSvenciFacturasAlvic() As Boolean

On Error GoTo eAjusteSvenciFacturasAlvic

    
    lblIndicador(9).Caption = "Prepara datos"
    lblIndicador(9).Refresh
    
    miSQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
    
    
    lblIndicador(9).Caption = "Comprobar forma pago"
    lblIndicador(9).Refresh
    
    'El SELECT srive para varios
    cadTitulo = " scafac.fecfactu between " & DBSet(txtFecha(25).Text, "F") & "  AND " & DBSet(txtFecha(26).Text, "F")
    cadTitulo = cadTitulo & " AND scafac.codforpa <>2 and scafac.codtipom in ('FA1','FA2','FAD','FAB')"
    
    'Comprobacion que todoas las formas de pago estan en seforpa
    miSQL = "select distinct substring(observa3,1,3)"
    miSQL = miSQL & " from  scafac left join scafac1 on scafac1.codtipom = scafac .codtipom and scafac1.numfactu= "
    miSQL = miSQL & " scafac.numfactu and scafac1.fecfactu = scafac .fecfactu"
    miSQL = miSQL & " left join slifac on scafac1.codtipom = slifac.codtipom and scafac1.numfactu= slifac.numfactu and scafac1.fecfactu = slifac.fecfactu and"
    miSQL = miSQL & " scafac1.Codtipoa = slifac.Codtipoa And scafac1.Numalbar = slifac.Numalbar WHERE "
    miSQL = miSQL & cadTitulo
    miSQL = miSQL & " group by scafac1.codtipom ,scafac1.numfactu,scafac1.fecfactu ,observa3"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cadPDFrpt = DBLet(miRsAux.Fields(0), "T")
        If cadPDFrpt = "" Then cadPDFrpt = "-1"
        miSQL = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", cadPDFrpt)
        
        If miSQL = "" Then Err.Raise 513, , "Forma de pago no existe: " & cadPDFrpt
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Si llega aqui, esta bien
    lblIndicador(9).Caption = "Eliminar en svenci"
    lblIndicador(9).Refresh
    
    miSQL = "DELETE FROM svenci WHERE (codtipom,numfactu,fecfactu) IN "
    miSQL = miSQL & " (SELECT codtipom,numfactu,fecfactu FROM scafac WHERE " & cadTitulo & " )"
    conn.Execute miSQL
    
    
    lblIndicador(9).Caption = "Crear en svenci"
    lblIndicador(9).Refresh
    
    miSQL = " INSERT INTO svenci (codtipom,numfactu,fecfactu,ordefect,fecefect,impefect,textoauxiliar)"
    miSQL = miSQL & " select scafac1.codtipom ,scafac1.numfactu,scafac1.fecfactu ,substring(observa3,1,3)"
    miSQL = miSQL & " ,scafac1.fecfactu ,sum(precoste) totales , trim(substring(observa3,4))"
    miSQL = miSQL & " from  scafac left join "
    miSQL = miSQL & " scafac1 on scafac1.codtipom = scafac .codtipom and scafac1.numfactu= scafac .numfactu and scafac1.fecfactu = scafac .fecfactu"
    miSQL = miSQL & " left join slifac on scafac1.codtipom = slifac.codtipom and scafac1.numfactu= slifac.numfactu and scafac1.fecfactu = slifac.fecfactu and"
    miSQL = miSQL & " scafac1.Codtipoa = slifac.Codtipoa And scafac1.Numalbar = slifac.Numalbar WHERE "
    miSQL = miSQL & cadTitulo
    miSQL = miSQL & " group by scafac1.codtipom ,scafac1.numfactu,scafac1.fecfactu ,observa3"
    conn.Execute miSQL
    
    
    
    
    lblIndicador(9).Caption = "Centimos"
    lblIndicador(9).Refresh
    

    miSQL = "INSERT INTO tmpinformes(CodUsu , nombre1, Codigo1, fecha1,campo1, Importe1, Importe2, Importe3) "
    miSQL = miSQL & " select " & vUsu.Codigo & " , scafac.codtipom , scafac.numfactu,scafac.fecfactu ,min(ordefect) efectoajuste,"
    miSQL = miSQL & " totalfac,sum(impefect) efectos,totalfac-sum(impefect) ajuste"
    miSQL = miSQL & " From scafac, svenci"
    miSQL = miSQL & " Where scafac.codtipom = svenci.codtipom And scafac.Numfactu = svenci.Numfactu And scafac.FecFactu = svenci.FecFactu AND "
    miSQL = miSQL & cadTitulo
    miSQL = miSQL & " group by 1,2,3 having totalfac-sum(impefect)<>0"
    conn.Execute miSQL

    
    lblIndicador(9).Caption = "Ajuste"
    lblIndicador(9).Refresh
    Espera 0.25
    
    cadTitulo = ""
    miSQL = "Select * from tmpinformes where codusu =" & vUsu.Codigo
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        lblIndicador(9).Caption = "Ajuste " & miRsAux!Codigo1 & " - " & miRsAux!Importe3
        lblIndicador(9).Refresh
        miSQL = "UPDATE svenci set impefect = impefect + " & DBSet(miRsAux!Importe3, "N")
        miSQL = miSQL & " WHERE codtipom =" & DBSet(miRsAux!nombre1, "T")
        miSQL = miSQL & " AND numfactu =" & DBSet(miRsAux!Codigo1, "N")
        miSQL = miSQL & " AND fecfactu =" & DBSet(miRsAux!fecha1, "F")
        miSQL = miSQL & " AND ordefect =" & DBSet(miRsAux!campo1, "N")
        conn.Execute miSQL
    
        If Abs(miRsAux!Importe3) > 10 Then
            cadTitulo = cadTitulo & miRsAux!nombre1 & "  " & Format(miRsAux!Codigo1, "000000")
            cadTitulo = cadTitulo & "  " & Format(miRsAux!fecha1, "dd/mm/yyyy") & Right(Space(15) & miRsAux!Importe3, 15) & vbCrLf
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If cadTitulo <> "" Then MsgBox "AJUSTES no normalizados" & vbCrLf & cadTitulo, vbExclamation
    
    
    'Metemos en el log
    
                '  LOG de acciones.                      5: Facturas compras
    Set LOG = New cLOG
    miSQL = "[ALVIC_FP] "
    If cadTitulo <> "" Then miSQL = miSQL & vbCrLf & vbCrLf & cadTitulo
    miSQL = miSQL & vbCrLf & "Inicio: " & Me.txtFecha(25).Text & "   Fin: " & txtFecha(26).Text
    LOG.Insertar 29, vUsu, miSQL
    Set LOG = Nothing
    '-----------------------
    
    
    
    AjusteSvenciFacturasAlvic = True
    

eAjusteSvenciFacturasAlvic:
    If Err.Number <> 0 Then MuestraError Err.Number, lblIndicador(9).Caption, Err.Description
    Set miRsAux = Nothing
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub
