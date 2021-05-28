VERSION 5.00
Begin VB.Form frmInformesNew2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "L"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   14025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraCambPrecTar 
      Height          =   6735
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   8640
      Begin VB.OptionButton optSituaArt 
         Caption         =   "Obsoletos"
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
         Left            =   2700
         TabIndex        =   52
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdCambiPrecio 
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
         Left            =   6030
         TabIndex        =   12
         Top             =   6120
         Width           =   1065
      End
      Begin VB.OptionButton optSituaArt 
         Caption         =   "Caducados"
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
         Left            =   6270
         TabIndex        =   11
         Top             =   5160
         Width           =   1455
      End
      Begin VB.OptionButton optSituaArt 
         Caption         =   "Bloqueados"
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
         Left            =   4440
         TabIndex        =   10
         Top             =   5160
         Width           =   1470
      End
      Begin VB.TextBox txtDescTarifa 
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   1680
         Width           =   5565
      End
      Begin VB.TextBox txtTarifa 
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
         Left            =   1605
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1680
         Width           =   1110
      End
      Begin VB.TextBox txtDescArticulo 
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
         Index           =   9
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   4605
         Width           =   4395
      End
      Begin VB.TextBox txtArticulo 
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
         Left            =   1605
         MaxLength       =   16
         TabIndex        =   8
         Top             =   4605
         Width           =   2175
      End
      Begin VB.OptionButton optSituaArt 
         Caption         =   "Todos"
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
         Left            =   1305
         TabIndex        =   9
         Top             =   5160
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.TextBox txtDescArticulo 
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   4200
         Width           =   4395
      End
      Begin VB.TextBox txtArticulo 
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
         Left            =   1605
         MaxLength       =   16
         TabIndex        =   7
         Top             =   4200
         Width           =   2175
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
         Index           =   38
         Left            =   1605
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1200
         Width           =   1350
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
         Index           =   34
         Left            =   7200
         TabIndex        =   13
         Top             =   6120
         Width           =   1065
      End
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
         Index           =   7
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   3555
         Width           =   5565
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
         Index           =   7
         Left            =   1605
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3555
         Width           =   1050
      End
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
         Index           =   6
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   3150
         Width           =   5565
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
         Index           =   6
         Left            =   1605
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3150
         Width           =   1050
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
         Index           =   20
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   2400
         Width           =   5565
      End
      Begin VB.TextBox txtCodProve 
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
         Left            =   1605
         TabIndex        =   4
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Image imgTarifa 
         Height          =   240
         Index           =   0
         Left            =   1350
         Picture         =   "frmInformesNew2.frx":0000
         Tag             =   "-1"
         ToolTipText     =   "Buscar tarifa"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa"
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
         Index           =   43
         Left            =   360
         TabIndex        =   28
         Top             =   1635
         Width           =   585
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   9
         Left            =   1320
         ToolTipText     =   "Buscar artículo"
         Top             =   4605
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
         Index           =   111
         Left            =   585
         TabIndex        =   26
         Top             =   4605
         Width           =   645
      End
      Begin VB.Image imgArticulo 
         Height          =   240
         Index           =   8
         Left            =   1320
         ToolTipText     =   "Buscar artículo"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Index           =   29
         Left            =   360
         TabIndex        =   24
         Top             =   3915
         Width           =   750
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
         Index           =   110
         Left            =   585
         TabIndex        =   23
         Top             =   4200
         Width           =   645
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   38
         Left            =   1320
         Picture         =   "frmInformesNew2.frx":0A02
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  cambio"
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
         Index           =   42
         Left            =   360
         TabIndex        =   21
         Top             =   885
         Width           =   1470
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   7
         Left            =   1320
         ToolTipText     =   "Buscar familia"
         Top             =   3555
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
         Index           =   109
         Left            =   585
         TabIndex        =   20
         Top             =   3555
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Index           =   28
         Left            =   360
         TabIndex        =   18
         Top             =   2865
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
         Index           =   108
         Left            =   585
         TabIndex        =   17
         Top             =   3150
         Width           =   645
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   6
         Left            =   1320
         ToolTipText     =   "Buscar familia"
         Top             =   3150
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Index           =   27
         Left            =   360
         TabIndex        =   15
         Top             =   2130
         Width           =   1005
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   20
         Left            =   1320
         ToolTipText     =   "Buscar proveedor"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label lblTitulo 
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
         Height          =   360
         Index           =   31
         Left            =   360
         TabIndex        =   1
         Top             =   315
         Width           =   7980
      End
   End
   Begin VB.Frame FramePromociones 
      Height          =   6015
      Left            =   135
      TabIndex        =   29
      Top             =   135
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton cmdACtualizaPromo 
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
         Left            =   4050
         TabIndex        =   38
         Top             =   5400
         Width           =   1065
      End
      Begin VB.CommandButton cmdCambioPromo 
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
         Left            =   4050
         TabIndex        =   37
         Top             =   5400
         Width           =   1065
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
         Index           =   13
         Left            =   1635
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   4005
         Width           =   1095
      End
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
         Index           =   13
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "Text5"
         Top             =   4005
         Width           =   3630
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
         Index           =   12
         Left            =   1635
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   3600
         Width           =   1095
      End
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
         Index           =   12
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text5"
         Top             =   3600
         Width           =   3630
      End
      Begin VB.TextBox txtCodProve 
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
         Index           =   23
         Left            =   1635
         TabIndex        =   34
         Top             =   2880
         Width           =   1095
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
         Index           =   23
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text5"
         Top             =   2880
         Width           =   3630
      End
      Begin VB.TextBox txtTarifa 
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
         Left            =   1635
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtDescTarifa 
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
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Text5"
         Top             =   2280
         Width           =   3765
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
         Index           =   42
         Left            =   1635
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1680
         Width           =   1350
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
         Index           =   41
         Left            =   1635
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1320
         Width           =   1350
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
         Index           =   38
         Left            =   5250
         TabIndex        =   39
         Top             =   5400
         Width           =   1065
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   13
         Left            =   1350
         Top             =   4005
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
         Index           =   134
         Left            =   480
         TabIndex        =   51
         Top             =   3600
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Index           =   38
         Left            =   255
         TabIndex        =   50
         Top             =   3315
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
         Index           =   133
         Left            =   480
         TabIndex        =   49
         Top             =   3960
         Width           =   645
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   12
         Left            =   1350
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgProveedor 
         Height          =   240
         Index           =   23
         Left            =   1350
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Left            =   255
         TabIndex        =   46
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Image imgTarifa 
         Height          =   240
         Index           =   1
         Left            =   1350
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa"
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
         Index           =   48
         Left            =   255
         TabIndex        =   44
         Top             =   2280
         Width           =   585
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   42
         Left            =   1350
         Picture         =   "frmInformesNew2.frx":0A8D
         Top             =   1680
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
         Index           =   132
         Left            =   480
         TabIndex        =   42
         Top             =   1320
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
         Index           =   131
         Left            =   480
         TabIndex        =   41
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  promoción"
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
         Index           =   47
         Left            =   255
         TabIndex        =   40
         Top             =   960
         Width           =   1800
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   41
         Left            =   1350
         Picture         =   "frmInformesNew2.frx":0B18
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label lblTitulo 
         Caption         =   "l"
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
         Height          =   360
         Index           =   35
         Left            =   255
         TabIndex        =   30
         Top             =   360
         Width           =   6075
      End
   End
End
Attribute VB_Name = "frmInformesNew2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Integer
    '1      .- Listado reparaciones efectuadas
    '2      .- Reparaciones tecnico
    
    '3      .- Revision carcteres multibase
    '4      .- Listado recargas telefonia movil
    '5      .- Facturacion de recargas
    
    '6      .- Listado de TRAZA por codprove en ventas.   ENERO 2008
    
    '       LIQUIDACION PROVEEDORES. Socios tipo TERRASANA
    '7      .- Cambio precio articulos
    '8      .- Generar facturas
    '9      .- Imprimir facturas proveedores (socios)
    '10     .-   "      ALBARANES   "           "
    
    
    '13     .- Generacion y facturacion de tickets agrupados
    '14     .- Listado del punto anterior
    
    '15     .- Listado trazabilidad albaranes
        
    '16     .- Ventas x agentes
    
    '17     .- Listado trabajadores . NO HACE DESDE HASTA
    
    '18     .- Cambio de proveedor en albaranes. Solicita el codprove
    
    '19     .- Cerrar aviso. Datos para crear albaran
    
    '20     .- Listado plantillas ofertas
    
    
    '21     .- Seleccionar otras ofertas del cliente
    '22     .- Listado llamadas
    
    '23     .- Listado situacion albaranes
    '24     .- Modificar expediente y legal en frecuencias
    
    '25     .- Datos para facturacion de cliente
    
    '26     .- Impresion pedidos por zona
        
    '27     .- REIMPRESION DE ALBARANES
    '28     .- Copiar precios compra-venta y venta-compra
    
    '30     .- Reparaciones en garantia de proveedor
    
    '31     .- Calculo del riesgo
    '32     .- Propuesta de pedido
    
    '33     .- Rp descuentos proveedor
    
    '34     .- Modificacion precios. Para pedir d/h tarifa, provee, familia tipo...
    '35     .- Dtos por activi marca fam
    
    '36     .- Resumen ventas agente
    
    '37     .- Beneficios por agente
    
    '38     .- Genera promociones.  D/H varios y lanzar frm
    '39     .-   actuali promo.   Comparte FRAME
    
    '40     .- Beneficios por proveedor. Igual que el de agente(COMPARTE FRAME)
    '41     .- Beneficio por cliente
    
                    
    '42     .- Listado control de albaranes  05/MAYO/2014
    '43     .-   "        "     "" facturados 15 Mayo
    
    '44     .- Cambio de precios lista proveedor  (igual que el de clientes)
    
    '45     .- Cambia departamento en frecencias (TEINSA)
    '46     .- Informe productividad
    '47     .- Cliente potencial. Pasar a cliente
    
    '48     .- beneficio Marca, Agente, Proveedor

    '49     .- ventas marca-familia
    '50     .- Compras marca-familia
    
    '51     .- Copiar pedido
    '52     .- Copiar albaran
    
    '53     .- Listado costes EULER
    '54     .- Listado lotes (fontenas)
    
    
Private IndiceImg As Integer
Private WithEvents frmCli As frmBasico2
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmBasico2
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmPr As frmBasico2
Attribute frmPr.VB_VarHelpID = -1
Private WithEvents frmBaPr As frmBasico2 'frmFacBancosPropios
Attribute frmBaPr.VB_VarHelpID = -1
Private WithEvents frmFP As frmBasico2 'frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmAg As frmBasico2 '%=%=frmFacAgentesCom
Attribute frmAg.VB_VarHelpID = -1
Private WithEvents frmZ As frmFacZonas
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmEn As frmFacFormasEnvio
Attribute frmEn.VB_VarHelpID = -1
Private WithEvents frmRut As frmFacRutas
Attribute frmRut.VB_VarHelpID = -1

Private PrimeraVez As Boolean




'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports

'nuevo Febrero 2010
Private cadPDFrpt As String 'Nombre del informe a enviar por email
Private vImprimedirecto As Boolean '
Private vMultiInforme As Integer
'-----------------------------------





'Variables comunes a todos os botones aceptar
Dim cadFrom As String
Dim campo As String, devuelve As String
Dim Codigo  As String
Dim ImpTot As Currency
Dim ImpTeo As Currency
Dim miSQL As String

Private Cadena_frmB As String
Private cadImpresion As String  'Facturacion


Private Sub cmdACtualizaPromo_Click()
    'Vamos a actualizar los precios promo
    'de siguiente a actual
    campo = ""
    If Me.txtTarifa(1).Text = "" Then campo = campo & vbCrLf & "-Tarifa"
    If campo <> "" Then
        campo = "Campos obligatorios: " & campo
        MsgBox campo, vbExclamation
        PonerFoco txtFecha(41)
        Exit Sub
    End If
    
    InicializarVbles
    
    cadSelect = "{spromo.codlista} = " & txtTarifa(1).Text
    'Proveedor  18 19
    If txtCodProve(23).Text <> "" Or txtCodProve(23).Text <> "" Then
        'devuelve = "pDHProve=""Proveedor: "
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 23, 23, devuelve) Then Exit Sub
        
    End If
    
    
    'Familia  4 5
    If txtFamia(12).Text <> "" Or txtFamia(13).Text <> "" Then
        'devuelve = "pDHFamilia=""Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 12, 13, devuelve) Then Exit Sub
        
    End If
    
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
  
   'Insert into
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    NumRegElim = 0
    campo = "select count(*) from "
    campo = campo & " spromo,sartic "
    campo = campo & " WHERE spromo.codartic=sartic.codartic AND " & cadSelect
    campo = campo & " AND not  fechain1  is null"
    miRsAux.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    If NumRegElim = 0 Then
        MsgBox "No hay valores con esos parametros", vbInformation
    Else
        campo = "Se van a actualizar " & NumRegElim & " registro(s). ¿Continuar?"
        If MsgBox(campo, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
    End If
    
    
    If NumRegElim > 0 Then
        campo = "select spromo.* from "
        campo = campo & " spromo,sartic "
        campo = campo & " WHERE spromo.codartic=sartic.codartic AND " & cadSelect
        campo = campo & " AND not  fechain1  is null" 'Vale, actualizo
        miRsAux.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            campo = "UPDATE spromo set "
            campo = campo & " fechaini = " & DBSet(miRsAux!fechain1, "F") & " , fechafin = " & DBSet(miRsAux!fechafi1, "F")
            campo = campo & " , precioac = " & DBSet(miRsAux!precionu, "N") & " , precioa1 = " & DBSet(miRsAux!precion1, "N") & ""
            campo = campo & " ,fechain1 = null , fechafi1 = null , precionu = null , precion1 = null"
            campo = campo & " WHERE spromo.codartic= " & DBSet(miRsAux!codArtic, "T") & " AND codlista = " & Me.txtTarifa(1).Text

             conn.Execute campo
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    Set miRsAux = Nothing
    
    Screen.MousePointer = vbDefault
    
    If NumRegElim > 0 Then Unload Me
    
End Sub


Private Sub cmdCambioPromo_Click()
    campo = ""
    If Me.txtFecha(41).Text = "" Then campo = campo & vbCrLf & "-Fecha inicio promocion"
    If Me.txtFecha(42).Text = "" Then campo = campo & vbCrLf & "-Fecha fin promocion"
    If Me.txtTarifa(1).Text = "" Then campo = campo & vbCrLf & "-Tarifa"
    If Me.txtCodProve(23).Text = "" Then campo = campo & vbCrLf & "-Proveedor"
    If campo <> "" Then
        campo = "Campos obligatorios: " & campo
        MsgBox campo, vbExclamation
        PonerFoco txtFecha(41)
        Exit Sub
    End If
    
    InicializarVbles
    
    cadSelect = "{spromo.codlista} = " & txtTarifa(1).Text
    'Proveedor  18 19
    If txtCodProve(23).Text <> "" Or txtCodProve(23).Text <> "" Then
        'devuelve = "pDHProve=""Proveedor: "
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 23, 23, devuelve) Then Exit Sub
        
    End If
    
    
    'Familia  4 5
    If txtFamia(12).Text <> "" Or txtFamia(13).Text <> "" Then
        'devuelve = "pDHFamilia=""Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 12, 13, devuelve) Then Exit Sub
        
    End If
    
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "spromo,sartic "
    campo = campo & " WHERE spromo.codartic=sartic.codartic AND " & cadSelect
    
    
    Set miRsAux = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    
    'Compruebo que no hay datos con precionuevo
    devuelve = "Select count(*) from " & campo & " AND not fechain1 is null"
    NumRegElim = 0
    miRsAux.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Existen promociones sin actualizar", vbExclamation
        campo = ""
        NumRegElim = 0
    End If
   'Insert into
    If campo <> "" Then CargaTmpCambioPromo
    Screen.MousePointer = vbDefault
    
    If NumRegElim = 0 Then Exit Sub
    
    
    frmAlmCambPromo.Show vbModal
    Unload Me
End Sub



Private Sub CargaTmpCambioPromo()
On Error GoTo ECargaTmpCambioPromo

    conn.Execute "DELETE FROM tmpinformes where codusu = " & vUsu.Codigo

    campo = "select precioac,precioa1,spromo.codartic,nomartic FROM " & campo
    miRsAux.Open campo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    campo = ""
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        'tmpinformes(codusu,codigo1,nombre1,nombre2,importe1,importe2,importe3,importe4)
        campo = campo & ", (" & vUsu.Codigo & "," & NumRegElim & "," & DBSet(miRsAux!codArtic, "T") & ","
        campo = campo & DBSet(miRsAux!NomArtic, "T") & "," & DBSet(miRsAux!precioac, "N") & ","
        campo = campo & DBSet(miRsAux!precioa1, "N", "N") & ",0,0)"
        
        If (NumRegElim Mod 50) = 0 Then
            campo = Mid(campo, 2)
            campo = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,nombre2,importeb1,importeb2,importeb3,importeb4) VALUES " & campo
            conn.Execute campo
            campo = ""
        End If
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If campo <> "" Then
        campo = Mid(campo, 2)
        campo = "INSERT INTO tmpinformes(codusu,codigo1,nombre1,nombre2,importeb1,importeb2,importeb3,importeb4) VALUES " & campo
        conn.Execute campo
    End If


    'En fecha1 y fecha2 pongo las fechas nuevas de promocion
    campo = "UpDATE tmpinformes set fecha1=" & DBSet(txtFecha(41).Text, "F") & ",  fecha2=" & DBSet(txtFecha(42).Text, "F")
    campo = campo & " ,campo1 = " & txtTarifa(1).Text
    campo = campo & " WHERE codusu = " & vUsu.Codigo
    conn.Execute campo

ECargaTmpCambioPromo:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "CargaTmp"
    Else
        If NumRegElim = 0 Then MsgBox "No hay registros con estos valores", vbExclamation
    End If
    
End Sub


Private Sub cmdCambiPrecio_Click()
    
    'Trocito comun
    campo = ""
    If Me.txtFecha(38).Text = "" Then campo = campo & vbCrLf & "-Fecha"
    If Opcion = 34 And Me.txtTarifa(0).Text = "" Then campo = campo & vbCrLf & "-Tarifa"
    If Me.txtCodProve(20).Text = "" Then campo = campo & vbCrLf & "-Proveedor"
    If campo <> "" Then
        campo = "Campos obligatorios: " & campo
        MsgBox campo, vbExclamation
        Exit Sub
    End If
    InicializarVbles



    If Opcion = 34 Then
        CambioPreciosTarifasVenta
    Else
        CambioPreciosTarifasCOMPRAS
    End If
End Sub
    
    
Private Sub CambioPreciosTarifasVenta()
Dim OkDobleComprobacion As Boolean
Dim SelectPrees As String

    cadSelect = "{slista.codlista} = " & txtTarifa(0).Text
    SelectPrees = ""
    If txtCodProve(20).Text <> "" Or txtCodProve(20).Text <> "" Then
        'devuelve = "pDHProve=""Proveedor: "
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 20, 20, devuelve) Then Exit Sub
        SelectPrees = SelectPrees & " AND sartic.codprove =" & txtCodProve(20).Text
    End If
    
    
    'Familia  4 5
    If txtFamia(6).Text <> "" Or txtFamia(7).Text <> "" Then
        'devuelve = "pDHFamilia=""Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 6, 7, devuelve) Then Exit Sub
        If txtFamia(6).Text <> "" Then SelectPrees = SelectPrees & " AND sartic.codfamia >=" & txtFamia(6).Text
        If txtFamia(7).Text <> "" Then SelectPrees = SelectPrees & " AND sartic.codfamia <=" & txtFamia(7).Text
        
    End If
    
    
    'Marca 2 3
    
    If txtArticulo(8).Text <> "" Or txtArticulo(9).Text <> "" Then
        devuelve = " "
        campo = "{sartic.codartic}"
        If Not PonerDesdeHasta(campo, "ART", 8, 9, devuelve) Then Exit Sub
        If txtArticulo(8).Text <> "" Then SelectPrees = SelectPrees & " AND sartic.codfamia >=" & DBSet(txtArticulo(8).Text, "T")
        If txtArticulo(9).Text <> "" Then SelectPrees = SelectPrees & " AND sartic.codfamia <=" & DBSet(txtArticulo(9).Text, "T")
        
    End If
    
    If Me.optSituaArt(0).Value Then
        'TODOS. No hacemos nada
        campo = ""
    Else
        
        If Me.optSituaArt(1).Value Then
            campo = "1"
        
        ElseIf Me.optSituaArt(2).Value Then
            campo = "2"
        Else
            campo = "3"
        End If
        campo = " AND {sartic.codstatu} = " & campo
        SelectPrees = SelectPrees & " AND sartic.codstatu = " & campo
    End If
    cadSelect = cadSelect & campo
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "slista,sartic "
    campo = campo & " WHERE slista.codartic=sartic.codartic AND " & cadSelect
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay registros con estos valores", vbExclamation
        Exit Sub
    End If
    
    
    
    'Ahora comprobare que ninugno tiene fechanue
    OkDobleComprobacion = True
    Codigo = ""
    campo = "slista,sartic "
    campo = campo & " WHERE slista.codartic=sartic.codartic AND " & cadSelect
    campo = campo & " AND not fechanue is null"
    If HayRegParaInforme(campo, "", True) Then
        Codigo = "Hay registros sin actualizar en lista de precios"
        OkDobleComprobacion = False
    End If
    
    
    '2017. Si hay actualizacion precios especiales tambien comprobaremos esto
    If vParamAplic.ActualizaPrecioEspecial Then
               
        campo = "sprees,sartic "
        campo = campo & " WHERE sprees.codartic=sartic.codartic AND "
        campo = campo & " not fechanue is null " & SelectPrees
        If HayRegParaInforme(campo, "", True) Then
            Codigo = Codigo & vbCrLf & "Hay registros sin actualizar en precios especiales"
            OkDobleComprobacion = False
        End If
    End If
    
    If Not OkDobleComprobacion Then
        MsgBox Codigo, vbExclamation
        Exit Sub
    End If
    
    
    'Ponemos el tmpprecioac a cero
    campo = "UPDATE slista,sartic set tmpprecioac=0 WHERE "
    campo = campo & " slista.codartic=sartic.codartic AND " & cadSelect
    ejecutar campo, False
    
    frmAlmCambPrec.vFecha = CDate(txtFecha(38).Text)
    frmAlmCambPrec.parSelSQL = cadSelect
    frmAlmCambPrec.Ventas = True
    frmAlmCambPrec.Show vbModal
End Sub


Private Sub CambioPreciosTarifasCOMPRAS()
    


    campo = "{slispr.codprove}"
    If Not PonerDesdeHasta(campo, "PRO", 20, 20, devuelve) Then Exit Sub
    

    
    
    'Familia  4 5
    If txtFamia(6).Text <> "" Or txtFamia(7).Text <> "" Then
        'devuelve = "pDHFamilia=""Familia: "
        campo = "{sartic.codfamia}"
        If Not PonerDesdeHasta(campo, "FAM", 6, 7, devuelve) Then Exit Sub
        
    End If
    
    
    'Marca 2 3
    
    If txtArticulo(8).Text <> "" Or txtArticulo(9).Text <> "" Then
        devuelve = " "
        campo = "{sartic.codartic}"
        If Not PonerDesdeHasta(campo, "ART", 8, 9, devuelve) Then Exit Sub
        
    End If
    
    If Me.optSituaArt(0).Value Then
        'TODOS. No hacemos nada
        campo = ""
    Else
        
        If Me.optSituaArt(1).Value Then
            campo = "1"
        
        ElseIf Me.optSituaArt(2).Value Then
            campo = "2"
        Else
            campo = "3"
        End If
        campo = " AND {sartic.codstatu} = " & campo
    End If
    cadSelect = cadSelect & campo
    cadSelect = QuitarCaracterACadena(cadSelect, "{")
    cadSelect = QuitarCaracterACadena(cadSelect, "}")
    
    campo = "slispr,sartic "
    campo = campo & " WHERE slispr.codartic=sartic.codartic AND " & cadSelect
    
    If Not HayRegParaInforme(campo, "", True) Then
        MsgBox "No hay registros con estos valores", vbExclamation
        Exit Sub
    End If
    
    'Borro la temporal
    campo = "DELETE FROM tmpslipreu WHERE codusu = " & vUsu.Codigo
    conn.Execute campo
    
    'Ahora comprobare que ninugno tiene fechanue
    campo = "slispr,sartic "
    campo = campo & " WHERE slispr.codartic=sartic.codartic AND " & cadSelect
    campo = campo & " AND not fechanue is null"
    If HayRegParaInforme(campo, "", True) Then
        MsgBox "Hay registros sin actualizar cambio precio", vbExclamation
        Exit Sub
    End If
    
    
    
    
    'Ponemos el tmpprecioac a cero
    campo = "INSERT INTO tmpslipreu(codusu,numofert,numlinea,codartic,nomartic,ampliaci,precioar) "
    campo = campo & "select " & vUsu.Codigo & ",slispr.codprove,@rownum:=@rownum+1 AS rownum ,slispr.codartic,nomartic,format(precioac,4),0.00  "
    campo = campo & " FROM slispr,sartic ,(SELECT @rownum:=0) r WHERE "
    campo = campo & " slispr.codartic=sartic.codartic AND " & cadSelect
    If ejecutar(campo, False) Then
        
        'Tema estetico.  Ampliaci es texto y lelva un importe. Reemplzamos la coma po
        campo = "UPDATE tmpslipreu SET ampliaci=replace(ampliaci,'.',',') WHERE codusu = " & vUsu.Codigo
        conn.Execute campo
        
        frmAlmCambPrec.vFecha = CDate(txtFecha(38).Text)
        frmAlmCambPrec.parSelSQL = "tmpslipreu.codusu = " & vUsu.Codigo
        frmAlmCambPrec.Ventas = False
        
        frmAlmCambPrec.Show vbModal
    End If
End Sub




Private Sub cmdCancel_Click(Index As Integer)
    'Si estamos en calculo de riesgo, cancelar  puede parar el proceso para salir
    If Index = 31 Then
        If Opcion = 0 Then
            'Le ha dado a cancelar.
            If MsgBox("¿Desea parar el proceso?", vbQuestion + vbYesNo) = vbYes Then Opcion = 31
                
            Exit Sub
        End If
    End If
    Unload Me
End Sub





Private Sub LlamarImprimir(PongoNombrePDF As Boolean, Optional NumeroDeCopias As Integer)
    If NumeroDeCopias = 0 Then NumeroDeCopias = 1
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .NumeroCopias = NumeroDeCopias
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2000 + Opcion   '2000 mas la opcion de entrada
        .NombrePDF = ""
        If PongoNombrePDF Then .NombrePDF = cadPDFrpt
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub



Private Function HacerListadogarantiaProveedor() As Boolean
Dim RT As ADODB.Recordset
Dim RAlb As ADODB.Recordset
Dim Donde As Byte
Dim C As Integer
Dim SQ As String


On Error GoTo EHacerListadogarantiaProveedor
    HacerListadogarantiaProveedor = False
    InicializarVbles
    cadTitulo = ""
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    'Comporobar si hay registros
    SQ = ""
    If txtFecha(36).Text <> "" Or txtFecha(37).Text <> "" Then
        campo = "{schrep.fechaalb}"
        If Not PonerDesdeHasta(campo, "F", 36, 37, "") Then Exit Function
        If txtFecha(36).Text <> "" Then cadTitulo = cadTitulo & " desde " & txtFecha(36).Text
        If txtFecha(37).Text <> "" Then cadTitulo = cadTitulo & " hasta " & txtFecha(37).Text
        If cadTitulo <> "" Then cadTitulo = "Fechas: " & cadTitulo
        SQ = cadSelect
    End If
    
    If txtCodProve(15).Text <> "" Or txtCodProve(16).Text <> "" Then
        
        campo = "{sartic.codprove}"
        If Not PonerDesdeHasta(campo, "PRO", 15, 16, devuelve) Then Exit Function
        cadTitulo = cadTitulo & devuelve
    End If
    
    If cadSelect <> "" Then
        cadSelect = Replace(cadSelect, "{", "(")
        cadSelect = Replace(cadSelect, "}", ")")
    End If
    
    If SQ <> "" Then
        SQ = Replace(SQ, "{", "(")
        SQ = Replace(SQ, "}", ")")
    End If
    
    
    
    Label3(94).Caption = "Preparando datos"
    Label3(94).Refresh
    miSQL = "DELETE from tmpnlotes where codusu = " & vUsu.Codigo
    conn.Execute miSQL
    miSQL = "DELETE from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute miSQL
    
    
    cadFormula = "insert into tmpinformes (codusu,codigo1,importe1,importe2,nombre1,nombre2,campo1,importeb1) VALUES (" & vUsu.Codigo & ","
    cadFrom = "insert into tmpnlotes (codusu,codprove,numalbar,fechaalb,numlinea,nomartic) values (" & vUsu.Codigo & ","
   
    Label3(94).Caption = "Leyendo reparaciones"
    Label3(94).Refresh
    
    miSQL = "select schrep.numalbar, schrep.fechaalb,sartic.codprove,1,sserie.codartic,schrep.nomartic,sserie.numserie"
    miSQL = miSQL & " from schrep ,sserie,sartic where schrep.numserie=sserie.numserie"
    miSQL = miSQL & " and schrep.codartic=sserie.codartic and  sartic.codartic=sserie.codartic "
   ' miSQL = miSQL & " and ((schrep.fechaalb >= '2001-11-01') "
   ' miSQL = miSQL & " and (schrep.fechaalb <= '2011-01-01')) and sartic.codprove>1"
    If cadSelect <> "" Then miSQL = miSQL & " AND " & cadSelect
    miSQL = miSQL & " and not fechavta  is null"
    miSQL = miSQL & " and schrep.fechaalb<DATE_ADD(fechavta, interval garantia day) group by 1,2"
   
    
    
    
    
    
    
    
    
    'HAremos la busqueda
    '-------------------------------------------
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    miSQL = ""
    cadTitulo = ""
    If miRsAux.EOF Then
        'MsgBox "Ningun dato generado", vbExclamation
        NumRegElim = 0
        miRsAux.Close
        GoTo EHacerListadogarantiaProveedor
    End If
    
    
    
    
    Label3(94).Caption = "Leyendo facturas"
    Label3(94).Refresh
    
    miSQL = "select slifac.numalbar,scafac1.fechaalb,sum(importel),sum(cantidad*preciove) from scafac1,slifac,sartic  where"
    miSQL = miSQL & " scafac1.codtipom=slifac.codtipom and scafac1.numfactu=slifac.numfactu and"
    miSQL = miSQL & " scafac1.FecFactu = slifac.FecFactu And scafac1.NumAlbar = slifac.NumAlbar And scafac1.codtipoa = slifac.codtipoa"
    miSQL = miSQL & " and slifac.codartic=sartic.codartic"
    If SQ <> "" Then
        Codigo = Replace(SQ, "schrep", "scafac1")
        miSQL = miSQL & " AND " & Codigo
    End If
    miSQL = miSQL & " group  by 1,2 order by 1,2"
    Set RT = New ADODB.Recordset
    RT.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    
    
    
    'Los albaranaes
    Label3(94).Caption = "Leyendo albaranes"
    Label3(94).Refresh
    
    miSQL = "select slialb.numalbar,scaalb.fechaalb,sum(importel),sum(cantidad*preciove) from scaalb,slialb,sartic  where"
    miSQL = miSQL & " slialb.codtipom=scaalb.codtipom AND slialb.numalbar=scaalb.numalbar AND scaalb.codtipom='ALR' "
    miSQL = miSQL & " and slialb.codartic=sartic.codartic  "
    If SQ <> "" Then
        Codigo = Replace(SQ, "schrep", "scaalb")
        miSQL = miSQL & " AND " & Codigo
    End If
    miSQL = miSQL & " group  by 1,2 order by 1,2"
    Set RAlb = New ADODB.Recordset
    RAlb.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    NumRegElim = 0
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Label3(94).Caption = "Recorriendo datos " & NumRegElim
        Label3(94).Refresh
        

        'primero veo si esta en
        'Debug.Print miRsAux!FechaAlb & "  " & miRsAux!NumAlbar
        Donde = 0
        If SituarRsGarantiaProveedorAlbaran(RAlb, miRsAux!Numalbar, miRsAux!FechaAlb) Then
            ImpTot = DBLet(RAlb.Fields(2), "N")
            ImpTeo = DBLet(RAlb.Fields(3), "N")
            Donde = 1
        Else
            If SituarRsGarantiaProveedor(RT, miRsAux!Numalbar, miRsAux!FechaAlb) Then
                ImpTot = DBLet(RT.Fields(2), "N")
                ImpTeo = DBLet(RT.Fields(3), "N")
                Donde = 2
            Else
                ImpTot = 0
                ImpTeo = 0
            End If
        End If
        
        'INSERTAMOS
        'en tmpinformes
        Codigo = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & DBLet(miRsAux!numSerie, "T") & "'," & Donde & "," & miRsAux!Codprove & ")"
        Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
        Codigo = cadFormula & Codigo
        conn.Execute Codigo
        

        Codigo = miRsAux!Numalbar & ",'" & Format(miRsAux!FechaAlb, FormatoFecha) & "',1,'')"
        Codigo = NumRegElim & "," & Codigo
        Codigo = cadFrom & Codigo
        conn.Execute Codigo
                
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    
    
    'Ultima fase, actualizar nombre proveedor
    Label3(94).Caption = "Leyendo proveedores"
    Label3(94).Refresh
    Codigo = "Select importeb1 from tmpinformes where codusu = " & vUsu.Codigo & " GROUP  BY  importeb1"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Codigo = CStr(CLng(miRsAux!importeb1))
        Label3(94).Caption = "Proveedor: " & Codigo
        Label3(94).Refresh
        
        miSQL = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", Codigo, "N")
        Codigo = "UPDATE tmpinformes set nombre3=" & DBSet(miSQL, "T") & " WHERE codusu = " & vUsu.Codigo
        Codigo = Codigo & " AND importeb1 = " & DBSet(miRsAux!importeb1, "N")
        conn.Execute Codigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
EHacerListadogarantiaProveedor:
    If Err.Number <> 0 Then
        MuestraError Err.Number
        NumRegElim = 0
    End If
    Set miRsAux = Nothing
    Set RT = Nothing
    Set RAlb = Nothing
End Function



Private Function SituarRsGarantiaProveedor(ByRef R As ADODB.Recordset, NA As Long, FA As Date) As Boolean
Dim Salir As Boolean
    SituarRsGarantiaProveedor = False
    If R.EOF Then
        Exit Function
    Else
        Salir = False
        While Not Salir
            'Mismo numero de albaran
            If R!Numalbar = NA Then
                If R!FechaAlb = FA Then
                    'OK es este. NO lo muevas
                    SituarRsGarantiaProveedor = True
                    Salir = True
                End If
            Else
                If R!Numalbar > NA Then
                    'El numero de albaran es mayor que el que pedimos. Nos salimos sin mover
                    Salir = True
                Else
                    'Es menor. Que se mueva
                End If
            End If
            
            If Not Salir Then
                R.MoveNext
                If R.EOF Then Salir = True
            End If
        Wend
    End If
End Function



Private Function SituarRsGarantiaProveedorAlbaran(ByRef R As ADODB.Recordset, NA As Long, FA As Date) As Boolean
Dim Salir As Boolean
    SituarRsGarantiaProveedorAlbaran = False
    If R.EOF Then
        Exit Function
    Else
        Salir = False
        While Not Salir
            'Mismo numero de albaran
            If R!Numalbar = NA Then
                If R!FechaAlb = FA Then
                    'OK es este. NO lo muevas
                    SituarRsGarantiaProveedorAlbaran = True
                    Salir = True
                End If
            Else
                If R!Numalbar > NA Then
                    'El numero de albaran es mayor que el que pedimos. Nos salimos sin mover
                    Salir = True
                Else
                    'Es menor. Que se mueva
                End If
            End If
            
            If Not Salir Then
                R.MoveNext
                If R.EOF Then Salir = True
               
            End If
        Wend
    End If
End Function


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 34
            PonerFoco txtFecha(38)
        Case 38
            PonerFoco txtFecha(41)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargaIconosNew()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    For Each Ima In Me.imgProveedor
        Ima.Picture = imgTarifa(0).Picture
    Next
    For Each Ima In Me.imgFamilia
        Ima.Picture = imgTarifa(0).Picture
    Next
    For Each Ima In Me.imgArticulo
        Ima.Picture = imgTarifa(0).Picture
    Next
    
    Err.Clear
End Sub





Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim IndiceCancel As Integer

    Me.Icon = frmPpal.Icon
    PrimeraVez = True
    
    CargaIconosNew
    
    limpiar Me
    
    FramePromociones.visible = False
    FraCambPrecTar.visible = False
    
    Caption = "Listado"
    IndiceCancel = Opcion
    
    Select Case Opcion
    Case 34, 44
        'cambio de precios
        '34.  De ventas
        '44: compras
        PonerFrameVisible FraCambPrecTar, H, W
        
        
        Me.lblTitulo(31).Caption = "Cambio precios tarifas "
        If Opcion = 34 Then
            lblTitulo(31).Caption = lblTitulo(31).Caption & "(VENTAS)"
        Else
            lblTitulo(31).Caption = lblTitulo(31).Caption & "(COMPRAS)"
        End If
        
        
        'Solo ventas
        lblDpto(43).visible = Opcion = 34
        imgTarifa(0).visible = Opcion = 34
        txtTarifa(0).visible = Opcion = 34
        txtDescTarifa(0).visible = Opcion = 34
        
        
        IndiceCancel = 34
    
    Case 38, 39
        If Opcion = 38 Then
             Me.lblTitulo(35).Caption = "Cambio precios promociones"
             lblDpto(47).Caption = "Nueva fecha promoción"
        Else
            Me.lblTitulo(35).Caption = "Actualizar precios promociones"
            lblDpto(47).Caption = "Fecha promoción"
            IndiceCancel = 38
        End If
        cmdACtualizaPromo.visible = Opcion = 39
        Me.cmdCambioPromo.visible = Opcion = 38
        PonerFrameVisible FramePromociones, H, W
    End Select
    
    Me.Height = H + 150
    Me.Width = W
    Me.cmdCancel(IndiceCancel).Cancel = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set miRsAux = Nothing
End Sub


Private Sub frmAg_DatoSeleccionado(CadenaSeleccion As String)
    Cadena_frmB = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Cadena_frmB = CadenaDevuelta
    
End Sub


Private Sub frmC_Selec(vFecha As Date)
    txtFecha(IndiceImg).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmEn_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub


Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescArticulo(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmPr_DatoSeleccionado(CadenaSeleccion As String)
    txtCodProve(IndiceImg) = RecuperaValor(CadenaSeleccion, 1)
    txtDescProve(IndiceImg) = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmRut_DatoSeleccionado(CadenaSeleccion As String)
    campo = CadenaSeleccion
End Sub




Private Sub imgArticulo_Click(Index As Integer)
    IndiceImg = Index
    Set frmMtoArticulos = New frmBasico2
    'frmMtoArticulos.DatosADevolverBusqueda3 = "@1@" 'Abrimos en Modo Busqueda
'    frmMtoArticulos.DesdeTPV = False
'    frmMtoArticulos.Show vbModal
    AyudaArticulos frmMtoArticulos, txtArticulo(IndiceImg)
    Set frmMtoArticulos = Nothing
End Sub





Private Sub imgFamilia_Click(Index As Integer)
    Cadena_frmB = ""
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vTitulo = "Familia"
    campo = "Codigo|sfamia|Codfamia|N||20·"
    campo = campo & "descripcion|sfamia|nomfamia|T||45·"
    frmB.vCampos = campo
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    frmB.vTabla = "sfamia"
    frmB.vSQL = ""
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
    If Cadena_frmB <> "" Then
        
        Me.txtFamia(Index).Text = RecuperaValor(Cadena_frmB, 1)
        Me.txtDescFamia(Index).Text = RecuperaValor(Cadena_frmB, 2)
        If Index = 2 Then
            PonerFoco txtFamia(3)
        End If
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
   IndiceImg = Index
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(IndiceImg).Text <> "" Then
        If IsDate(txtFecha(IndiceImg).Text) Then frmC.Fecha = CDate(txtFecha(IndiceImg).Text)
    End If
   frmC.Show vbModal
   Set frmC = Nothing
    
End Sub




Private Sub imgProveedor_Click(Index As Integer)
    IndiceImg = Index
'    Set frmPr = New frmComProveedores
'    frmPr.DatosADevolverBusqueda = "0|1|"
'    frmPr.Show vbModal

    Set frmPr = New frmBasico2
    AyudaProveedores frmPr, txtCodProve(IndiceImg)
    Set frmPr = Nothing

    PonerFoco txtCodProve(IndiceImg)
End Sub


Private Sub imgTarifa_Click(Index As Integer)
            Cadena_frmB = ""
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vTitulo = "Tarifas"
            campo = "Codigo|starif|codlista|N||20·"
            campo = campo & "Nombre|startif|nomlista|T||40·"
            frmB.vCampos = campo
            frmB.vCargaFrame = False
            frmB.vDevuelve = "0|1|"
            frmB.vselElem = 1
            frmB.vConexionGrid = 1  'ODBC Ariges
            frmB.vTabla = "starif"
            frmB.vSQL = ""
            frmB.Show vbModal
            Set frmB = Nothing
            Screen.MousePointer = vbDefault
            If Cadena_frmB <> "" Then
                Me.txtTarifa(Index).Text = RecuperaValor(Cadena_frmB, 1)
                Me.txtDescTarifa(Index).Text = RecuperaValor(Cadena_frmB, 2)
            End If
End Sub


Private Sub optSituaArt_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub


Private Sub txtArticulo_GotFocus(Index As Integer)
    ConseguirFoco txtArticulo(Index), 3
End Sub

Private Sub txtArticulo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgArticulo_Click Index
    End If
End Sub

Private Sub txtArticulo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
    
End Sub


Private Sub txtArticulo_LostFocus(Index As Integer)
Dim T As String
    
    txtArticulo(Index).Text = Trim(txtArticulo(Index).Text)
    If txtArticulo(Index).Text = "" Then
        'EN blanco
        txtDescArticulo(Index).Text = ""
        Exit Sub
    End If
    
    
    T = "codartic"
    Codigo = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtArticulo(Index).Text, "T", T)
    If Codigo = "" Then
        MsgBox "No existe el artículo : " & txtArticulo(Index).Text, vbExclamation
    Else
        txtArticulo(Index).Text = T
    End If
    Me.txtDescArticulo(Index).Text = Codigo
    Codigo = ""
End Sub


Private Sub txtCodProve_GotFocus(Index As Integer)
    ConseguirFoco txtCodProve(Index), 3
End Sub

Private Sub txtCodProve_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) > 0 Then
        'Caption = KeyCode
        If KeyCode = 65 Then imgProveedor_Click Index
    End If
End Sub

Private Sub txtCodProve_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCodProve_LostFocus(Index As Integer)
    txtCodProve(Index).Text = Trim(txtCodProve(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtCodProve(Index).Text <> "" Then
        If IsNumeric(txtCodProve(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtCodProve(Index).Text, "N")
            If Codigo = "" Then
                If Index = 20 Or Index = 12 Then
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
End Sub


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
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFamia_LostFocus(Index As Integer)
    txtFamia(Index).Text = Trim(txtFamia(Index).Text)
    Codigo = ""
    miSQL = ""
    If txtFamia(Index).Text <> "" Then
        If IsNumeric(txtFamia(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtFamia(Index).Text, "N")
            If Codigo = "" Then MsgBox "El codigo no pertence a ningun familia", vbExclamation
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


Private Sub txtTarifa_GotFocus(Index As Integer)
    ConseguirFoco txtTarifa(Index), 3
End Sub

Private Sub txtTarifa_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTarifa_LostFocus(Index As Integer)

    txtTarifa(Index).Text = Trim(txtTarifa(Index).Text)
    Codigo = ""
    miSQL = ""

    If txtTarifa(Index).Text <> "" Then
        If IsNumeric(txtTarifa(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomlista", "starif", "codlista", txtTarifa(Index).Text, "N")
            If Codigo = "" Then miSQL = "El código no pertence a ninguna tarifa"
        Else
            miSQL = "Campo numerico"
        End If
    End If
    
    Me.txtDescTarifa(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Index = 0 Then
            txtTarifa(Index).Text = ""
            PonerFoco txtTarifa(Index)
        End If
    End If
End Sub



Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
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

'Dado un FRAME lo pone a true y lo situa en x:120 y:0 y devuelve lo que debe medir el form
Private Sub PonerFrameVisible(ByRef F As Frame, ByRef CH As Integer, CW As Integer)
    F.Top = 0
    F.Left = 120
    F.visible = True
    CH = F.Height + 420
    CW = F.Width + 240
End Sub





Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
    vMultiInforme = 0
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
        If indD = 27 Or indH = 28 Then Subtipo = "FH"
        
        
    Case "PRO"
        Set TDes = txtCodProve(indD)
        Set THas = txtCodProve(indH)
        Set DesD = txtDescProve(indD)
        Set DesH = txtDescProve(indH)
        Subtipo = "N"
 
    Case "ART"

        Set TDes = txtArticulo(indD)
        Set THas = txtArticulo(indH)
        Set DesD = txtDescArticulo(indD)
        Set DesH = txtDescArticulo(indH)
        Subtipo = "T"
      
 
        
    Case "FAM"
        'FAMILIA
         
        Set TDes = Me.txtFamia(indD)
        Set THas = txtFamia(indH)
        Subtipo = "N"
        Set DesD = txtDescFamia(indD)
        Set DesH = txtDescFamia(indH)
    


    End Select
    
    devuelve = CadenaDesdeHasta(TDes.Text, THas.Text, campo, Subtipo)
    If devuelve = "Error" Then Exit Function
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

'Para las reparaciones. Carga el importe real y teorico.
Private Sub CargaImporteRealReparaciones()
'Dim ImpTot As Currency
'Dim ImpTeo As Currency
Dim miSQL As String
Dim RT As ADODB.Recordset

    'A partir de la reparacion , mirare en los albaranes, y de los albaranes ver el coste real de la reparacion y el teorico
    Set miRsAux = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    
    'Meto el select para las
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If
    

 
    
    'Montamos el select al reves
    
    
'    codigo = "s.codtipom=l.codtipom and s.codtipoa=l.codtipoa " & codigo
'    codigo = "s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND " & codigo
'    codigo = "s.codtipom=l.codtipom AND " & codigo
'    codigo = "sartic.codartic = l.codartic AND " & codigo
'    'codigo = "select l.*,s.fechaalb,preciove,h.numrepar,h.fecrepar from  schrep h,slifac l,scafac1 s,sartic where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb  AND " & codigo
'    codigo = "select l.*,s.fechaalb,preciove,h.numrepar,h.fecrepar from  schrep h,slifac l,scafac1 s,sartic where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb AND h.codtipom= s.codtipoa AND " & codigo
'    codigo = codigo & " ORDER BY s.numalbar ,s.fechaalb"
    
    'AHORA   FEBRERO 2014
    Label3(158).Caption = "Selecionar reparaciones"
    Label3(158).Refresh
    miSQL = "select numrepar,fecrepar,codtipom,numalbar,fechaalb from schrep h where " & Mid(Codigo, 5)

    
    
    
    
    
    
    
    
    
    
    
    'EL ORDEN
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    NumRegElim = 1
    miSQL = ""
    cadTitulo = ""
    While Not miRsAux.EOF
        Label3(158).Caption = miRsAux!numrepar & " " & miRsAux!fecrepar
        Label3(158).Refresh
        
        
        miSQL = "Select sum(importel) tot,sum(round(preciove*cantidad,2)) *100 teo,sum(coalesce(l.preciouc,0.000)*cantidad) coste "
        miSQL = miSQL & " FROM slifac l,scafac1 s,sartic"
        
        miSQL = miSQL & " WHERE s.codtipom=l.codtipom and s.codtipoa=l.codtipoa "
        miSQL = miSQL & " AND s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND "
        miSQL = miSQL & " s.numalbar=l.numalbar AND sartic.codartic = l.codartic "
        miSQL = miSQL & " AND s.codtipoa=" & DBSet(miRsAux!codtipom, "T")
        miSQL = miSQL & " AND s.numalbar=" & miRsAux!Numalbar
        miSQL = miSQL & " AND s.fechaalb=" & DBSet(miRsAux!FechaAlb, "F")
        
        RT.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT!Tot) Or Not IsNull(RT!teo) Then
                NumRegElim = NumRegElim + 1
                miSQL = ", (" & vUsu.Codigo & "," & miRsAux!numrepar & ",'" & Format(miRsAux!fecrepar, FormatoFecha) & "',0,0,"
                miSQL = miSQL & NumRegElim & "," & Val(DBLet(RT!teo, "N")) & "," & TransformaComasPuntos(CStr(RT!Tot))
                miSQL = miSQL & ",'" & DBLet(RT!coste, "N") & "')"
                cadTitulo = cadTitulo & miSQL
            End If
        Else
           ' S top
        End If
        RT.Close
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close


    If cadTitulo <> "" Then
        'El ultimo
        cadTitulo = Mid(cadTitulo, 2)
        
        miSQL = "insert into tmpnlotes (codusu,codprove,fechaalb,numalbar,nomartic,numlinea,codartic,cantidad,numlotes) VALUES " & cadTitulo
        conn.Execute miSQL
        cadTitulo = ""
    End If

   
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If
    
    Label3(158).Caption = "Garantias"
    Label3(158).Refresh
    miSQL = "select numrepar,fecrepar,tieneman,h.codclien,m.nummante,"
    miSQL = miSQL & " mes01act,mes02act,mes03act,mes04act,mes05act,mes06act,mes07act,mes08act,mes09act,mes10act,mes11act,mes12act"
    miSQL = miSQL & " from schrep h,sserie s left join scaman m  on s.nummante=m.nummante and s.codclien=m.codclien"
    miSQL = miSQL & " where h.numserie=s.numserie and s.codartic=h.codartic "
    If Codigo <> "" Then miSQL = miSQL & Codigo
    
    'EL ORDEN
    IndiceImg = 12
    If txtFecha(1).Text <> "" Then IndiceImg = Month(CDate(txtFecha(1).Text))
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadTitulo = "" 'Para saber los mantenimientos que ya hemos sumado
    While Not miRsAux.EOF
        Label3(158).Caption = "Garantia " & miRsAux!numrepar & " " & miRsAux!fecrepar
        Label3(158).Refresh
        ImpTot = 0
        If miRsAux!TieneMan = 1 Then
            If IsNull(miRsAux!nummante) Then
                
                miSQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", miRsAux!codClien)
                miSQL = "Cliente: " & miRsAux!codClien & "   " & miSQL & vbCrLf
                miSQL = "Error grave." & vbCrLf & "Con mantenimiento y sin numero de mantenimiento" & vbCrLf & miSQL
                miSQL = miSQL & "Reparacion: " & miRsAux!numrepar & " " & miRsAux!fecrepar & vbCrLf
                'MsgBox miSQL, vbExclamation
            End If
            miSQL = Format(miRsAux!codClien, "000000") & DBLet(miRsAux!nummante, "T") & "|"
                
            If InStr(1, cadTitulo, miSQL) = 0 Then
                cadTitulo = cadTitulo & miSQL
                '--------------------------------------------------------------------
                'OK, TIENE MANTENIMIENTO
                'Ire recorriendo los importes desde mes01act hasta el mes hasta
                'Si la fecha es fin es nada, entonces hare tooodos
                For NumRegElim = 1 To IndiceImg
                    If Not IsNull(miRsAux.Fields(NumRegElim + 4)) Then ImpTot = ImpTot + miRsAux.Fields(NumRegElim + 4)
                Next
            Else
               ' St op
            End If
        End If
        If ImpTot <> 0 Then
            'UPDATEAMOS LA tmp
            miSQL = "UPDATE tmpnlotes set nomartic=" & CLng(ImpTot * 100) & " WHERE codusu = " & vUsu.Codigo
            miSQL = miSQL & " AND codprove = " & miRsAux!numrepar & " AND fechaalb = '" & Format(miRsAux!fecrepar, FormatoFecha) & "' AND numalbar =0"
            conn.Execute miSQL
        End If
        '--------------
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Set RT = Nothing
    
End Sub
    


Private Sub EstadisticaReparacionTecnicoNueva()
Dim RT As ADODB.Recordset
Dim OptimizarSelect As String
Dim RAlb As ADODB.Recordset
Dim C As Long
Dim EnAlbaranes As Boolean

    Label3(63).Caption = "Obteniendo reg. albaranes"
    Label3(63).Refresh
    

    'Preparamos las temporales
    Codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    Codigo = "DELETE FROM tmpnlotes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    
    
    
    'LOS INSERTS PARA LAS TABLAS temporales                                         numserie
    cadFormula = "insert into tmpinformes (codusu,codigo1,importe1,importe2,nombre1,nombre2) VALUES (" & vUsu.Codigo & ","
    cadFrom = "insert into tmpnlotes (codusu,codprove,numalbar,fechaalb,numlinea,nomartic) values (" & vUsu.Codigo & ","
    
    
    'Optimizacion del select
    If cadSelect <> "" Then
        'Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " WHERE " & cadSelect
    Else
        Codigo = ""
    End If
    

    Codigo = "Select distinct(codtipom) from schrep  " & Codigo
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    OptimizarSelect = ""
    While Not miRsAux.EOF
        Codigo = DBLet(miRsAux!codtipom, "T")
        If Codigo <> "" Then OptimizarSelect = OptimizarSelect & " OR scafac1.codtipoa = '" & Codigo & "'"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If OptimizarSelect <> "" Then
        OptimizarSelect = Mid(OptimizarSelect, 4) 'quito el preimer or
        OptimizarSelect = "(" & OptimizarSelect & ")"
    End If
    
    'Cargo el RS con todos los datos de los albarnes
    miSQL = "select scafac1.numalbar,scafac1.fechaalb,scafac1.codtipoa,sum(importel),sum(cantidad*preciove)"
    miSQL = miSQL & " from scafac1,slifac,sartic where  scafac1.codtipom =slifac.codtipom  and scafac1.numfactu  =slifac.numfactu"
    miSQL = miSQL & " and scafac1.fecfactu  =slifac.fecfactu"
    miSQL = miSQL & " and scafac1.codtipoa  =slifac.codtipoa  and scafac1.numalbar  =slifac.numalbar and sartic.codartic=slifac.codartic"
    
    
    cadNomRPT = ""
    If txtFecha(2).Text <> "" Then cadNomRPT = cadNomRPT & " AND fechaalb >='" & Format(txtFecha(2).Text, FormatoFecha) & "'"
    If txtFecha(3).Text <> "" Then cadNomRPT = cadNomRPT & " AND fechaalb <='" & Format(txtFecha(3).Text, FormatoFecha) & "'"
    If OptimizarSelect <> "" Then cadNomRPT = cadNomRPT & " AND " & OptimizarSelect
    

    miSQL = miSQL & cadNomRPT
    miSQL = miSQL & " group by scafac1.numalbar,scafac1.fechaalb,scafac1.codtipoa order by codtipoa,numalbar,fechaalb"

    'Cargamos las sumas en facturas
    miRsAux.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    'Cargamos las sumas en albaranaes  ############
    miSQL = "Select scaalb.numalbar,scaalb.fechaalb,scaalb.codtipom,sum(importel),sum(cantidad*preciove)"
    miSQL = miSQL & " from scaalb,slialb,sartic where  scaalb.codtipom =slialb.codtipom  and scaalb.numalbar  =slialb.numalbar"
    miSQL = miSQL & " and sartic.codartic=slialb.codartic"
    cadNomRPT = Replace(cadNomRPT, " scafac1.codtipoa", "scaalb.codtipom")
    miSQL = miSQL & cadNomRPT
    miSQL = miSQL & " group by numalbar,fechaalb,codtipom order by codtipom,numalbar,fechaalb"

    Set RAlb = New ADODB.Recordset
    RAlb.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    
    
    
    
    
    
    
    
    'Cargamos el rS de la reparaciones
    If cadSelect <> "" Then
        'Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " WHERE " & cadSelect
    Else
        Codigo = ""
    End If
    

    Codigo = " from schrep  " & Codigo
    
    Set RT = New ADODB.Recordset
    RT.Open "select count(*)" & Codigo, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    C = DBLet(RT.Fields(0), "N")
    RT.Close
    
    RT.Open "select * " & Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    While Not RT.EOF
        NumRegElim = NumRegElim + 1
        
        Label3(63).Caption = "Rep: " & RT!numrepar & "  (" & NumRegElim & "/" & C & ")"
        Label3(63).Refresh
        
        If IsNull(RT!codtipom) Or IsNull(RT!Numalbar) Or IsNull(RT!FechaAlb) Then
            ImpTeo = 0
            ImpTot = 0
        Else
            
            PonerIMportesAlbaranes RAlb, RT!codtipom, RT!Numalbar, RT!FechaAlb, ImpTot, ImpTeo, EnAlbaranes
        End If
        
 

            
                'INSERTAMOS
                'en tmpinformes
                Codigo = "'" & DevNombreSQL(RT!NomArtic) & "','" & DBLet(RT!numSerie, "T") & "')"
                Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
                Codigo = cadFormula & Codigo
                conn.Execute Codigo
                
                
                'en tmpnlotes
                '       numprove
                Codigo = DBLet(RT!nummante, "T")
                If Codigo <> "" Then
                    Codigo = UCase(Codigo)
                    'Debug.Print Codigo
                    If Codigo = "S/MTO" Or Codigo = "SIN ESPC." Then
                        Codigo = "0"
                    Else
                        Codigo = "1"
                    End If
                Else
                    Codigo = "0"
                End If
                Codigo = Abs(EnAlbaranes) & ",'" & Format(RT!fecrepar, FormatoFecha) & "'," & Codigo & ",'" & Trim(DevNombreSQL(RT!NomClien)) & "')"
                Codigo = NumRegElim & "," & Codigo
                Codigo = cadFrom & Codigo
                conn.Execute Codigo
                
                 
      
        
        RT.MoveNext
    Wend
    RT.Close
    RAlb.Close
    miRsAux.Close
  Set RT = Nothing
  Set RAlb = Nothing
 

    
    


End Sub



Private Sub PonerIMportesAlbaranes(ByRef RAlba As ADODB.Recordset, codtipom As String, alb As Long, fech As Date, ByRef Impt As Currency, ByRef ImpTeor As Currency, ByRef EnAlbaranes As Boolean)
Dim fin As Boolean
Dim Esta As Boolean

    Impt = 0
    ImpTeor = 0
    
    'Comprobamos en albaranes primero
    EnAlbaranes = True
    Esta = False
    If Not RAlba.EOF Then
        
        fin = False
        While Not fin
            If RAlba!codtipom = codtipom Then
                If RAlba!Numalbar = alb Then
                    If RAlba!FechaAlb = fech Then
                        'AQUI ESTA
                        fin = True
                        Esta = True
                        Impt = RAlba.Fields(3)
                        ImpTeor = RAlba.Fields(4)
                    End If
                    
                Else
                    'SI ha sobrepasado YA no esta
                    If RAlba!Numalbar > alb Then fin = True
                End If
            End If
            RAlba.MoveNext
            If RAlba.EOF Then fin = True
        Wend
    
        RAlba.MoveFirst
        If Esta Then Exit Sub  'Ya lo hemos encontrado
    End If
    
    
    EnAlbaranes = False
    If miRsAux.EOF Then Exit Sub
    fin = False
    While Not fin
        If miRsAux!Codtipoa = codtipom Then
            If miRsAux!Numalbar = alb Then
                If miRsAux!FechaAlb = fech Then
                    'AQUI ESTA
                    fin = True
                    Impt = miRsAux.Fields(3)
                    ImpTeor = miRsAux.Fields(4)
                End If
            Else
                If miRsAux!Numalbar > alb Then miRsAux.MoveLast
            End If
        End If
        miRsAux.MoveNext
        If miRsAux.EOF Then fin = True
    Wend
    miRsAux.MoveFirst
End Sub

Private Sub EstadisticaReparacionTecnico()
    'Preparamos las temporales
    Codigo = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    Codigo = "DELETE FROM tmpnlotes WHERE codusu = " & vUsu.Codigo
    conn.Execute Codigo
    
    

    
    'LOS INSERTS PARA LAS TABLAS temporales                                         numserie
    cadFormula = "insert into tmpinformes (codusu,codigo1,importe1,importe2,nombre1,nombre2) VALUES (" & vUsu.Codigo & ","
    cadFrom = "insert into tmpnlotes (codusu,codprove,numalbar,fechaalb,numlinea,nomartic) values (" & vUsu.Codigo & ","
    
    'Montamos el select al reves
    'PARA LAS FACTURAS
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If
    Codigo = " s.codtipom=l.codtipom and s.codtipoa=l.codtipoa " & Codigo
    Codigo = " s.fecfactu = l.fecfactu and s.numfactu=l.numfactu AND " & Codigo
    Codigo = " s.codtipom=l.codtipom AND " & Codigo
    Codigo = " sartic.codartic = l.codartic AND " & Codigo
    Codigo = " h.numserie=sserie.numserie AND h.codclien=sserie.codclien AND " & Codigo
    Codigo = " sclien.codclien = h.codclien AND " & Codigo
    Codigo = " where  h.numalbar = s.numalbar and h.fechaalb=s.fechaalb AND " & Codigo
    'Las tablas
    Codigo = " from schrep h,slifac l,scafac1 s,sclien , sserie,sartic" & Codigo
    Codigo = "select l.*,s.fechaalb,preciove,h.fecrepar,h.nomclien,tieneman,h.nomartic,h.numserie " & Codigo
    'EL ORDEN
    Codigo = Codigo & " ORDER BY s.numalbar ,s.fechaalb"
    miRsAux.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    miSQL = ""
    While Not miRsAux.EOF
    

    
        If Codigo <> CStr(miRsAux!Numalbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If Codigo <> "" Then
                NumRegElim = NumRegElim + 1
                'INSERTAMOS
                'en tmpinformes
                Codigo = RecuperaValor(miSQL, 1)
                Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
                Codigo = cadFormula & Codigo
                conn.Execute Codigo
                
                
                'en tmpnlotes
                '       numprove
                Codigo = RecuperaValor(miSQL, 2)
                Codigo = NumRegElim & "," & Codigo
                Codigo = cadFrom & Codigo
                conn.Execute Codigo
                
                
            End If
            'Meto dos datos enpipados
            miSQL = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & miRsAux!numSerie & "')"
            miSQL = miSQL & "|"
            miSQL = miSQL & "0,'" & Format(miRsAux!fecrepar, FormatoFecha) & "'," & miRsAux!TieneMan & ",'" & DevNombreSQL(miRsAux!NomClien) & "')|"
            Codigo = miRsAux!Numalbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!PrecioVe * miRsAux!cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        NumRegElim = NumRegElim + 1
        
        'en tmpinformes
        Codigo = RecuperaValor(miSQL, 1)
        Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
        Codigo = cadFormula & Codigo
        conn.Execute Codigo
        
        
        'en tmpnlotes
        '       numprove
        Codigo = RecuperaValor(miSQL, 2)
        Codigo = NumRegElim & "," & Codigo
        Codigo = cadFrom & Codigo
        conn.Execute Codigo
        
        
    End If
    

    
    'AHORA HAGO EL INSERT PARA LOS ALBARANES QUE NO HAN SIDO FACTURADOS
    'PARA LOS ALBARANES
    If cadSelect <> "" Then
        Codigo = Replace(cadSelect, "schrep.", "h.")
        Codigo = " AND " & Codigo
    Else
        Codigo = ""
    End If

    miSQL = "select l.*,preciove,tieneman,h.fechaalb,h.numserie,h.nomclien,fecrepar"
    miSQL = miSQL & " from schrep h,scaalb c,slialb l,sartic a,sserie s "
    miSQL = miSQL & " WHERE h.codtipom=c.codtipom and h.numalbar=c.numalbar and h.fechaalb=c.fechaalb and"
    miSQL = miSQL & " l.numalbar=c.numalbar and l.codtipom=c.codtipom and l.codartic=a.codartic"
    miSQL = miSQL & " and h.numserie=s.numserie and h.codclien =s.codclien" & Codigo
    
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Codigo = ""
    miSQL = ""
    While Not miRsAux.EOF
        If Codigo <> CStr(miRsAux!Numalbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)) Then
            If Codigo <> "" Then
                NumRegElim = NumRegElim + 1
                'INSERTAMOS
                'en tmpinformes
                Codigo = RecuperaValor(miSQL, 1)
                Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
                Codigo = cadFormula & Codigo
                conn.Execute Codigo
                
                
                'en tmpnlotes
                '       numprove
                Codigo = RecuperaValor(miSQL, 2)
                Codigo = NumRegElim & "," & Codigo
                Codigo = cadFrom & Codigo
                conn.Execute Codigo
                
            End If
            'Meto dos datos enpipados
            miSQL = "'" & DevNombreSQL(miRsAux!NomArtic) & "','" & miRsAux!numSerie & "')"
            miSQL = miSQL & "|"
            miSQL = miSQL & "1,'" & Format(miRsAux!fecrepar, FormatoFecha) & "'," & miRsAux!TieneMan & ",'" & DevNombreSQL(miRsAux!NomClien) & "')|"
            Codigo = miRsAux!Numalbar & "|" & Format(miRsAux!FechaAlb, FormatoFecha)
            ImpTot = 0
            ImpTeo = 0
        End If
        ImpTeo = ImpTeo + (miRsAux!PrecioVe * miRsAux!cantidad)
        ImpTot = miRsAux!ImporteL + ImpTot
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If miSQL <> "" Then
        NumRegElim = NumRegElim + 1
        
        'en tmpinformes
        Codigo = RecuperaValor(miSQL, 1)
        Codigo = NumRegElim & "," & TransformaComasPuntos(CStr(ImpTot)) & "," & TransformaComasPuntos(CStr(ImpTeo)) & "," & Codigo
        Codigo = cadFormula & Codigo
        conn.Execute Codigo
        
        
        'en tmpnlotes
        '       numprove
        Codigo = RecuperaValor(miSQL, 2)
        Codigo = NumRegElim & "," & Codigo
        Codigo = cadFrom & Codigo
        conn.Execute Codigo
        
        
    End If
    


End Sub





'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Facturacion de recargas de telefonia
'
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Function ObtenerContadorAlbaran(NumAlb As String) As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean

    On Error GoTo ErrConAlb

    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer("ALV") Then
        Do
            NumAlb = vTipoMov.ConseguirContador("ALV")
            vTipoMov.IncrementarContador ("ALV")
            miSQL = "select count(*) from scaalb where codtipom='ALV' and numalbar=" & NumAlb
            Existe = (RegistrosAListar(miSQL) > 0)
        Loop Until Existe = False
        ObtenerContadorAlbaran = True
    Else
        ObtenerContadorAlbaran = False
    End If
    Set vTipoMov = Nothing
    Exit Function
    
ErrConAlb:
    ObtenerContadorAlbaran = False
    MuestraError Err.Number, "Obtener contador albaran", Err.Description
End Function




'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Facturacion y contabilizacion de tickets
'       ========================================
'
'
'
'
'       Cuando esta la marca de contabilizar tickets agrupados, lo que haremos sera
'       a partir de los FTI crear las facturas agrupadas con el contador FTG "EN LA CONTABILIDAD"
'       en el ariges, en scafac, no creo ninguna factura
'       O bien una diaria o una mensual (dependera del parametro)
'
'
'       Insertaremos en una tabla los tckets que entran en cada factura
'----------------------------------------------------------------------
'----------------------------------------------------------------------


'------------------------------------------------------------------
'------------------------------------------------------------------
'
'       Informe de trazabilidad
'       ========================================
'
'
'
'
'       A partir del desde /hasta mostraremos el informe que tiene la asociacion
'       entre albaranes de compra / venta
'
'
'       Hay datos tanto en albaranes como en facturas, con lo cual insertare sobre tmp
'
'
'----------------------------------------------------------------------
'----------------------------------------------------------------------




'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'
'INforme de pedido de proveedores. Despues podra generar un pedido desde aqui
'
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------




Private Function ListadoCostesEuler() As Boolean
Dim Ins As Boolean
Dim RT As ADODB.Recordset
Dim fin As Boolean
Dim Fin2 As Boolean
Dim J As Integer

    On Error GoTo eListadoCostesEuler
    ListadoCostesEuler = False
    Label3(207).Caption = "Leyendo BD" 'indicador
    Label3(207).Refresh
    Set miRsAux = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    cadSelect = Replace(cadSelect, "{", "")
    cadSelect = Replace(cadSelect, "}", "")
        
    miSQL = "DELETE from tmpinformes where codusu =" & vUsu.Codigo
    conn.Execute miSQL
    
    
    miSQL = "insert into tmpinformes(codusu,codigo1,campo1,nombre1,fecha1,campo2,nombre2,importe1,importe2,importe3,importe4,nombre3)"
    miSQL = miSQL & " select " & vUsu.Codigo & ", @rownum:=@rownum+1 AS rownum  , scafac.numfactu,"
    miSQL = miSQL & " concat(scafac.codtipom ,right(concat('000000',scafac.numfactu),8)),scafac.fecfactu,scafac.codclien,scafac.nomclien,"
    miSQL = miSQL & " brutofac,0,0,0, codccost"
    miSQL = miSQL & " from (SELECT @rownum:=0) r,scafac ,scafac1 ,straba,sclien WHERE"
    
    miSQL = miSQL & " scafac.NumFactu = scafac1.NumFactu And  scafac.FecFactu =scafac1.FecFactu And scafac.codtipom = scafac1.codtipom    "
    miSQL = miSQL & " AND scafac1.codtraba =straba.codtraba  "
    miSQL = miSQL & " AND scafac.codclien =sclien.codclien AND "
    miSQL = miSQL & cadSelect
    miSQL = miSQL & " group by scafac.codtipom,scafac.numfactu,scafac.fecfactu"
    conn.Execute miSQL
    
    
    Espera 0.1
    Label3(207).Caption = "Leyendo costes" 'indicador
    Label3(207).Refresh
    
    miSQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadFormula = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    J = 0
    fin = False
    Do
        Label3(207).Caption = J + 1 & " de " & cadFormula
        Label3(207).Refresh
    
        miSQL = "SELECT codusu,codigo1,nombre1,fecha1 from tmpinformes where codusu = " & vUsu.Codigo & " ORDER BY codigo1 LIMIT " & J & ",10"
        miRsAux.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            fin = True
        Else
            J = J + 10
            miSQL = ""
            While Not miRsAux.EOF
                miSQL = miSQL & ", ('" & Mid(miRsAux!nombre1, 1, 3) & "'," & Mid(miRsAux!nombre1, 4) & "," & DBSet(miRsAux!fecha1, "F") & ")"
                
    
    
                miRsAux.MoveNext
            Wend
            miRsAux.MoveFirst
            
            miSQL = Mid(miSQL, 2) 'quitamos la primera coma
            Codigo = "SELECT codtipom,numfactu,fecfactu,sum(if(tipo=0,cantidad*precioar,0)),sum(if(tipo<>0,cantidad*precioar,0)) from slifac_eu"
            Codigo = Codigo & " where (codtipom,numfactu,fecfactu) in (" & miSQL & ") Group by 1,2,3"
                
            RT.Open Codigo, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RT.EOF
                miSQL = RT!codtipom & Right("00000000" & RT!Numfactu, 8)
                miRsAux.MoveFirst
                
                Fin2 = False
                
                While Not Fin2
                    If miRsAux.EOF Then
                        Fin2 = True
                    Else
                        If miRsAux!nombre1 = miSQL Then
                            If miRsAux!fecha1 = RT!FecFactu Then
                                Fin2 = True
                                Codigo = "UPDATE tmpinformes set importe2=" & DBSet(RT.Fields(3), "N")
                                Codigo = Codigo & ",importe3=" & DBSet(RT.Fields(4), "N")
                                Codigo = Codigo & " WHERE codusu =" & vUsu.Codigo & " AND codigo1 =" & miRsAux!Codigo1
                                conn.Execute Codigo
                            End If
                        End If
                        miRsAux.MoveNext
                    End If
                Wend
                RT.MoveNext
            Wend
            RT.Close
            
        End If
        miRsAux.Close
    
    
    Loop Until fin

    Label3(207).Caption = "Centros de coste" 'indicador
    Label3(207).Refresh
    
    NumRegElim = 0
    miSQL = "select nombre3 from tmpinformes where codusu = " & vUsu.Codigo & " GROUP BY 1"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If DBLet(miRsAux!nombre3) = "" Then
            miSQL = "UPDATE tmpinformes set nombre3 = 'C.Coste NULO'"
            miSQL = miSQL & " WHERE coalesce(nombre3,'') = '' AND codusu = " & vUsu.Codigo
            conn.Execute miSQL
        Else
            miSQL = DevuelveDesdeBD(conConta, "nomccost", IIf(vParamAplic.ContabilidadNueva, "ccoste", "cabccost"), "codccost", miRsAux!nombre3, "T")
            If miSQL = "" Then miSQL = "No encontrado:" & miRsAux!nombre3
            miSQL = "UPDATE tmpinformes set nombre3 = " & DBSet(miSQL, "T")
            miSQL = miSQL & " WHERE nombre3 = " & DBSet(miRsAux!nombre3, "T") & " AND codusu = " & vUsu.Codigo
            conn.Execute miSQL
        End If
        NumRegElim = 1 'Para saber si hay datos
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Diferencias
    Label3(207).Caption = "Diferencias" 'indicador
    Label3(207).Refresh
    
    Codigo = "update tmpinformes set importe4=importe1 -(importe2 + importe3)"
    Codigo = Codigo & " Where (Importe2 + Importe3 <> 0) And CodUsu = " & vUsu.Codigo
    conn.Execute Codigo

    Codigo = "update tmpinformes set importe4=NULL"
    Codigo = Codigo & " Where (Importe2 + Importe3 = 0) And CodUsu =  " & vUsu.Codigo
    conn.Execute Codigo
    
    If NumRegElim = 1 Then
        ListadoCostesEuler = True
    Else
        MsgBox "No existen datos ", vbExclamation
    End If
    
    
eListadoCostesEuler:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set RT = Nothing
    Label3(207).Caption = ""
End Function

'    Dim campo As String, devuelve As String
'Dim Codigo  As String
'Dim ImpTot As Currency
'Dim ImpTeo As Currency
'Dim miSQL As String
    


