VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacModAlbFac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar datos factura / albar�n"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   9480
      TabIndex        =   17
      Top             =   6660
      Width           =   1065
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   780
      Left            =   9090
      TabIndex        =   49
      Top             =   1035
      Width           =   2670
      Begin VB.TextBox Text1 
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
         Index           =   16
         Left            =   525
         TabIndex        =   50
         Tag             =   "T|N|S|||scafac|totalfac||N|"
         Text            =   "12345678911234567899"
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Factura"
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
         Index           =   2
         Left            =   120
         TabIndex        =   53
         Top             =   0
         Width           =   1365
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   1545
      MaxLength       =   80
      TabIndex        =   16
      Tag             =   "Observaci�n 5|T|S|||scafac1|observa5||N|"
      Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
      Top             =   6045
      Width           =   10200
   End
   Begin VB.TextBox Text1 
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
      Index           =   30
      Left            =   1545
      MaxLength       =   80
      TabIndex        =   15
      Tag             =   "Observaci�n 4|T|S|||scafac1|observa4||N|"
      Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
      Top             =   5640
      Width           =   10200
   End
   Begin VB.TextBox Text1 
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
      Index           =   29
      Left            =   1545
      MaxLength       =   80
      TabIndex        =   14
      Tag             =   "Observaci�n 3|T|S|||scafac1|observa3||N|"
      Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
      Top             =   5235
      Width           =   10200
   End
   Begin VB.TextBox Text1 
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
      Index           =   28
      Left            =   1545
      MaxLength       =   80
      TabIndex        =   13
      Tag             =   "Observaci�n 2|T|S|||scafac1|observa2||N|"
      Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
      Top             =   4830
      Width           =   10200
   End
   Begin VB.TextBox Text1 
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
      Index           =   27
      Left            =   1545
      MaxLength       =   80
      TabIndex        =   12
      Tag             =   "Observaci�n 1|T|S|||scafac1|observa1||N|"
      Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
      Top             =   4425
      Width           =   10200
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   10680
      TabIndex        =   18
      Top             =   6660
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   240
      TabIndex        =   43
      Top             =   1050
      Width           =   8760
      Begin VB.TextBox Text1 
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
         Left            =   3615
         TabIndex        =   47
         Tag             =   "Nombre|T|S|||scafac|nomclien||N|"
         Text            =   "12345678911234567899"
         Top             =   285
         Width           =   4785
      End
      Begin VB.TextBox Text1 
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
         Left            =   1290
         TabIndex        =   44
         Tag             =   "tel�fono Cliente|T|S|||scafac|codclien|0000|N|"
         Text            =   "12345678911234567899"
         Top             =   285
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   1
         Left            =   240
         TabIndex        =   52
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   11
         Left            =   2805
         TabIndex        =   46
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
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
         Index           =   10
         Left            =   555
         TabIndex        =   45
         Top             =   285
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
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
      Height          =   855
      Left            =   240
      TabIndex        =   32
      Top             =   120
      Width           =   11580
      Begin VB.TextBox Text1 
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
         Left            =   9765
         TabIndex        =   42
         Tag             =   "Cod|F|N|||scafac1|fechaalb|dd/mm/yyyy|S|"
         Text            =   "Text1 7"
         Top             =   330
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Index           =   4
         Left            =   7530
         TabIndex        =   40
         Tag             =   "Cod|N|N|||scafac1|numalbar|00000|S|"
         Text            =   "Text1 7"
         Top             =   330
         Width           =   1110
      End
      Begin VB.TextBox Text1 
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
         Index           =   3
         Left            =   6930
         TabIndex        =   38
         Tag             =   "Cod|T|N|||scafac1|codtipoa||S|"
         Text            =   "Text1 7"
         Top             =   330
         Width           =   570
      End
      Begin VB.TextBox Text1 
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
         Index           =   2
         Left            =   4230
         TabIndex        =   37
         Tag             =   "F|F|N|||scafac|fecfactu|dd/mm/yyyy|S|"
         Text            =   "Text1 7"
         Top             =   330
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   1935
         TabIndex        =   35
         Tag             =   "Cod|N|N|||scafac|numfactu|00000|S|"
         Text            =   "Text1 7"
         Top             =   330
         Width           =   1110
      End
      Begin VB.TextBox Text1 
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
         Left            =   1335
         TabIndex        =   33
         Tag             =   "Cod|T|N|||scafac|codtipom||S|"
         Text            =   "Text1 7"
         Top             =   330
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura/Albar�n"
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
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Top             =   0
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha alb."
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
         Left            =   9015
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Albar�n"
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
         Index           =   6
         Left            =   6090
         TabIndex        =   39
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
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
         Left            =   3480
         TabIndex        =   36
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Factura"
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
         Left            =   555
         TabIndex        =   34
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   1545
      MaxLength       =   35
      TabIndex        =   1
      Tag             =   "Domicilio|T|N|||scafac|domclien||N|"
      Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
      Top             =   2340
      Width           =   4710
   End
   Begin VB.TextBox Text2 
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
      Left            =   8475
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   2340
      Width           =   3285
   End
   Begin VB.TextBox Text1 
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
      Index           =   14
      Left            =   7890
      MaxLength       =   4
      TabIndex        =   7
      Tag             =   "Cod. Agente|N|N|0|9999|scafac|codagent|0000|N|"
      Text            =   "Text1"
      Top             =   2340
      Width           =   585
   End
   Begin VB.TextBox Text1 
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
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   0
      Tag             =   "tel�fono Cliente|T|S|||scafac|telclien||N|"
      Text            =   "12345678911234567899"
      Top             =   1935
      Width           =   2010
   End
   Begin VB.TextBox Text1 
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
      Left            =   2445
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Poblaci�n|T|N|||scafac|pobclien||N|"
      Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
      Top             =   2730
      Width           =   3810
   End
   Begin VB.TextBox Text1 
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
      Left            =   1545
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "CPostal|T|N|||scafac|codpobla||N|"
      Text            =   "Text15"
      Top             =   2730
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Left            =   1545
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Provincia|T|N|||scafac|proclien||N|"
      Text            =   "Text1 Text1 Text1 Text1 Text22"
      Top             =   3135
      Width           =   4695
   End
   Begin VB.TextBox Text1 
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
      Left            =   7890
      MaxLength       =   3
      TabIndex        =   6
      Tag             =   "Direccion/Dpto.|N|S|0|999|scafac|coddirec|000|N|"
      Text            =   "Text1"
      Top             =   1935
      Width           =   585
   End
   Begin VB.TextBox Text1 
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
      Left            =   8475
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   19
      Tag             =   "Direccion/Dpto.|T|S|||scafac|nomdirec||N|"
      Text            =   "Text1"
      Top             =   1935
      Width           =   3285
   End
   Begin VB.TextBox Text1 
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
      Index           =   26
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   5
      Tag             =   "Refere. Cliente|T|S|||scafac1|referenc|||"
      Text            =   "Text1 Text1 Text1 Te"
      Top             =   3570
      Width           =   1770
   End
   Begin VB.TextBox Text1 
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
      Index           =   18
      Left            =   6870
      MaxLength       =   4
      TabIndex        =   8
      Tag             =   "Banco|N|S|0|9999|scafac|codbanco|0000|N|"
      Text            =   "Text1 7"
      Top             =   3540
      Width           =   690
   End
   Begin VB.TextBox Text1 
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
      Index           =   19
      Left            =   7875
      MaxLength       =   4
      TabIndex        =   9
      Tag             =   "Sucursal|N|S|0|9999|scafac|codsucur|0000|N|"
      Text            =   "Text1 7"
      Top             =   3540
      Width           =   690
   End
   Begin VB.TextBox Text1 
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
      Index           =   20
      Left            =   8955
      MaxLength       =   2
      TabIndex        =   10
      Tag             =   "Digito Control|T|S|||scafac|digcontr|00|N|"
      Text            =   "Text1 7"
      Top             =   3540
      Width           =   450
   End
   Begin VB.TextBox Text1 
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
      Index           =   21
      Left            =   9795
      MaxLength       =   10
      TabIndex        =   11
      Tag             =   "Cuenta Bancaria|T|S|||scafac|cuentaba|0000000000|N|"
      Text            =   "Text1 7"
      Top             =   3540
      Width           =   1620
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10170
      Top             =   2745
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
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
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   7605
      Picture         =   "frmFacModAlbFac.frx":0000
      Tag             =   "-1"
      ToolTipText     =   "Buscar direc./dpto"
      Top             =   1980
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
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
      Index           =   12
      Left            =   330
      TabIndex        =   48
      Top             =   4155
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
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
      Index           =   7
      Left            =   330
      TabIndex        =   31
      Top             =   2340
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Agente"
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
      Index           =   34
      Left            =   6435
      TabIndex        =   30
      Top             =   2340
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "Tel�fono"
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
      Index           =   19
      Left            =   330
      TabIndex        =   29
      Top             =   1935
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Poblaci�n"
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
      Index           =   16
      Left            =   330
      TabIndex        =   28
      Top             =   2730
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
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
      Index           =   17
      Left            =   330
      TabIndex        =   27
      Top             =   3135
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Direc./Dpto"
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
      Left            =   6435
      TabIndex        =   26
      Top             =   1935
      Width           =   1170
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   1260
      ToolTipText     =   "Buscar poblaci�n"
      Top             =   2745
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Ref. Cliente"
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
      Index           =   13
      Left            =   330
      TabIndex        =   25
      Top             =   3570
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Banco"
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
      Left            =   6870
      TabIndex        =   24
      Top             =   3300
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Sucursal"
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
      Index           =   4
      Left            =   7875
      TabIndex        =   23
      Top             =   3300
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "DC"
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
      Left            =   8955
      TabIndex        =   22
      Top             =   3300
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta"
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
      Left            =   9795
      TabIndex        =   21
      Top             =   3300
      Width           =   840
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   7605
      ToolTipText     =   "Buscar agente"
      Top             =   2340
      Width           =   240
   End
End
Attribute VB_Name = "frmFacModAlbFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Where As String

Private WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2 '%=%=frmFacAgentesCom
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmDptoEnvio As frmFacCliEnvDpto
Attribute frmDptoEnvio.VB_VarHelpID = -1

Dim kCampo  As Integer


Private Sub cmdAceptar_Click()
Dim cT As cTag
Dim T As TextBox
    Set cT = New cTag
    
    Where = ""
    For Each T In Text1
        T.Text = Trim(T.Text)
        cT.Cargar T
        If cT.Vacio = "N" Then
            If T.Text = "" Then
                Where = Where & vbCrLf & "-" & cT.Nombre
                kCampo = T.Index
            End If
        End If
    Next
    
    If Text1(12).Text = "" Xor Text1(13).Text = "" Then
        Where = Where & vbCrLf & "- Departamento incorrecto"
        kCampo = 12
    End If
    
    
    If Where <> "" Then
        Where = "Campos obligados: " & vbCrLf & Where
        MsgBox Where, vbExclamation
        If kCampo > 0 Then PonerFoco Text1(kCampo)
        Exit Sub
    End If
    
    If MsgBox("Desea actualizar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    'Vamos a actualizar. No hacemos transaacciones ya que las modificaciones son a datos "NO" relevantes
    Where = "UPDATE scafac1 SET referenc = " & DBSet(Text1(26).Text, "T", "S")
    For kCampo = 27 To 31
        Where = Where & ", observa" & kCampo - 26 & " = " & DBSet(Text1(kCampo).Text, "T", "S")
    Next
    Where = Where & " WHERE scafac1.numfactu=" & DBSet(Data1.Recordset!Numfactu, "N") & " AND scafac1.codtipom =" & DBSet(Data1.Recordset!codtipom, "T")
    Where = Where & " AND scafac1.fecfactu=" & DBSet(Data1.Recordset!FecFactu, "F")
    Where = Where & " AND  scafac1.numalbar=" & DBSet(Data1.Recordset!Numalbar, "N") & " AND scafac1.codtipoa =" & DBSet(Data1.Recordset!Codtipoa, "T")
    conn.Execute Where
    
    'domclien codpobla pobclien proclien coddirec  nomdirec codagent codbanco codsucur  digcontr cuentaba
    Where = ""
    For NumRegElim = 7 To 21
        If NumRegElim <= 14 Or NumRegElim >= 18 Then
            Text1(NumRegElim).Text = Trim(Text1(NumRegElim).Text)
            cT.Cargar Text1(NumRegElim)
            Where = Where & ", " & cT.columna & " = " & DBSet(Text1(NumRegElim).Text, cT.TipoDato, cT.Vacio)
        End If
    Next
    Where = Mid(Where, 2) 'primiera coma
    Where = "UPDATE scafac SET " & Where
    Where = Where & " WHERE scafac.numfactu=" & DBSet(Data1.Recordset!Numfactu, "N") & " AND scafac.codtipom =" & DBSet(Data1.Recordset!codtipom, "T")
    Where = Where & " AND scafac.fecfactu=" & DBSet(Data1.Recordset!FecFactu, "F")
    conn.Execute Where

    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If kCampo < 0 Then
        PonerCamposForma Me, Data1
        Text1_LostFocus 14
        kCampo = 0
    End If
    Screen.MousePointer = vbDefault
End Sub
 
Private Sub Form_Load()
Dim SQL As String
On Error GoTo EPonerCampos
    
     'Icono de busqueda
    For kCampo = 1 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = imgBuscar(0).Picture
    Next kCampo

    kCampo = -1
    Me.Icon = frmPpal.Icon
    limpiar Me

    
    SQL = "select * from scafac,scafac1 where scafac.numfactu = scafac1.numfactu and scafac.codtipom = scafac1.codtipom and scafac.fecfactu = scafac1.fecfactu "
    SQL = SQL & " AND " & Where
    Data1.ConnectionString = conn
    Data1.RecordSource = SQL
    Data1.Refresh
    
   
    
    Exit Sub
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub
    





Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Where = CadenaSeleccion
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
    Where = CadenaSeleccion
End Sub

Private Sub frmDptoEnvio_DatoSeleccionado(CadenaSeleccion As String)
    Where = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    
    Where = ""
        Select Case Index
        Case 2 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            kCampo = 9
            If Where <> "" Then
                  Text1(kCampo).Text = RecuperaValor(Where, 1)
                  Text1(kCampo + 1).Text = RecuperaValor(Where, 2)
                  Text1(kCampo + 2).Text = ""
            End If
                      
            
        
        Case 0 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
            kCampo = 12
            Set frmDptoEnvio = New frmFacCliEnvDpto
            frmDptoEnvio.DireccionesEnvio = False
            If Text1(kCampo).Text <> "" Then
                frmDptoEnvio.VerDatoDpto = CInt(Text1(kCampo).Text)
            Else
                frmDptoEnvio.VerDatoDpto = -1
            End If
            frmDptoEnvio.codClien = CLng(Text1(6).Text)
            frmDptoEnvio.NomClien = Text1(15).Text
            frmDptoEnvio.Show vbModal
            Set frmDptoEnvio = Nothing
            If Where <> "" Then
                  Text1(kCampo).Text = RecuperaValor(Where, 1)
                  Text1(kCampo + 1).Text = RecuperaValor(Where, 2)
             End If
             
        Case 1 'Agente
            kCampo = 14
'            Set frmA = New frmFacAgentesCom
'            frmA.DatosADevolverBusqueda = "0"
'            frmA.Show vbModal
            Set frmA = New frmBasico2
            AyudaAgentesComerciales frmA, Text1(kCampo), , True
            Set frmA = Nothing

            If Where <> "" Then
                  Text1(kCampo).Text = RecuperaValor(Where, 1)
                  Text2(kCampo).Text = RecuperaValor(Where, 2)
            End If
        End Select
        
        If Where <> "" Then
            PonerFoco Text1(kCampo)
            Where = ""
        End If
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), 4
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpressGnral KeyAscii, 4, False
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 9: KEYBusqueda KeyAscii, 2 'poblacion
            Case 12: KEYBusqueda KeyAscii, 0 'direc/dpto
            Case 14: KEYBusqueda KeyAscii, 1 'agente
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 4, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
        
    If Not PerderFocoGnral(Text1(Index), 4) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 9 'Cod. Postal
             If Text1(Index).Locked Then Exit Sub
             If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
                Exit Sub
             End If
            
             Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
             Text1(Index + 2).Text = devuelve
        
       
        
        Case 12 'Cod. Direc
            devuelve = ""
            If PonerFormatoEntero(Text1(Index)) Then
                devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(6).Text, "N", , "coddirec", Text1(12).Text, "N")
                If devuelve = "" Then
                    MsgBox "No existe la direccion/departamento: " & Text1(Index).Text, vbInformation
                    
                End If
            End If
            Text1(Index + 1).Text = devuelve
            If devuelve = "" Then
                If Text1(Index).Text <> "" Then PonerFoco Text1(Index)
                Text1(Index).Text = ""
                
            End If
            
        Case 14 'Cod. Agente
            devuelve = ""
            If PonerFormatoEntero(Text1(Index)) Then
                devuelve = PonerNombreDeCod(Text1(Index), conAri, "sagent", "nomagent")
                'If devuelve = "" Then MsgBox "NO existe el agente: " & Text1(Index).Text, vbExclamation
            End If
            Text2(Index).Text = devuelve
            If devuelve = "" Then
                If Text1(Index).Text <> "" Then PonerFoco Text1(Index)
                Text1(Index).Text = ""
            End If
        
            
        Case 18 To 21 'banco, sucursal
            PonerFormatoEntero Text1(Index)
        
    End Select
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
      KEYpressGnral KeyAscii, 4, False
End Sub
