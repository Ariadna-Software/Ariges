VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacTPVEnt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punto de Venta"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   15675
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacTPVEnt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   15675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFito 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   43
      Top             =   8280
      Width           =   10455
      Begin VB.TextBox txtFito2 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Text            =   "frmFacTPVEnt.frx":000C
         Top             =   120
         Width           =   10095
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   42
      Tag             =   "precio art para albaran"
      Top             =   720
      Visible         =   0   'False
      Width           =   430
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   11280
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   41
      Text            =   "Fam"
      Top             =   4920
      Visible         =   0   'False
      Width           =   430
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   11280
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   40
      Tag             =   "P|T|S|||sliven|dto2|||"
      Text            =   "Dto"
      Top             =   5760
      Visible         =   0   'False
      Width           =   430
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   11280
      MaxLength       =   12
      TabIndex        =   39
      Tag             =   "PorcentajeIVA|N|N|||sliven|dto2|||"
      Text            =   "%IVA"
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   11040
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   38
      Tag             =   "precioarticulo|N|N|||sliven|dto2|||"
      Text            =   "dto2"
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   11040
      MaxLength       =   12
      TabIndex        =   37
      Tag             =   "Dto2|N|N|||sliven|dto2|||"
      Text            =   "dto2"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   10680
      MaxLength       =   12
      TabIndex        =   36
      Tag             =   "Dto1|N|S|||sliven|codigiva||N|"
      Text            =   "dto1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   27
      Top             =   2880
      Width           =   15360
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
         Height          =   315
         Index           =   2
         Left            =   1560
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   720
         Width           =   735
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
         Height          =   315
         Index           =   2
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   720
         Width           =   2820
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
         Index           =   1
         Left            =   9360
         MaxLength       =   80
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   240
         Width           =   5895
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   240
         Width           =   4605
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
         Left            =   1500
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label label1 
         Caption         =   "Dpto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   555
      End
      Begin VB.Image imgBuscar 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1080
         ToolTipText     =   "Buscar artículo"
         Top             =   720
         Width           =   360
      End
      Begin VB.Label label1 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   10
         Left            =   7560
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1100
         ToolTipText     =   "Buscar artículo"
         Top             =   220
         Width           =   360
      End
      Begin VB.Label label1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   900
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   14040
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   12600
      TabIndex        =   26
      Top             =   8400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   11040
      MaxLength       =   12
      TabIndex        =   25
      Tag             =   "Tipo iva|N|N|0||sliven|codigiva|0|N|"
      Text            =   "tipoiva"
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame FrameTot 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   7920
      TabIndex        =   19
      Top             =   960
      Width           =   7650
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   630
         Index           =   8
         Left            =   4755
         TabIndex        =   24
         Top             =   105
         Width           =   2775
      End
      Begin VB.Label label1 
         BackColor       =   &H00F5F5F5&
         Caption         =   "TOTAL LINEA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   100
         Width           =   4720
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   840
         Index           =   6
         Left            =   3000
         TabIndex        =   21
         Top             =   840
         Width           =   4545
      End
      Begin VB.Label label1 
         BackColor       =   &H00F5F5F5&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.Frame FrameCab 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   4335
      Begin VB.ComboBox cboNumVenta 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frmFacTPVEnt.frx":0012
         Left            =   1920
         List            =   "frmFacTPVEnt.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label label1 
         Caption         =   "Nº Venta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotVentas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   345
         Index           =   4
         Left            =   2880
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label label1 
         Caption         =   "Ventas Abiertas:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label label1 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   840
      MaxLength       =   15
      TabIndex        =   2
      Tag             =   "Cod.EAN|T|S|||slitick|codigoea|||"
      Text            =   "Artic EAN"
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   10080
      MaxLength       =   12
      TabIndex        =   8
      Tag             =   "Precio art.|N|N|0|999999.0000|slitick|precioar|###,##0.0000|N|"
      Text            =   "Precio ar."
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   240
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "Linea|N|N|||slitick|numlinea|0|S|"
      Text            =   "line"
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1920
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Cod. Artículo|T|N|||slitick|codartic|||"
      Text            =   "Artic Artic Artic5"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   6840
      MaxLength       =   16
      TabIndex        =   5
      Tag             =   "Cantidad|N|N|||slitick|cantidad|#,###,###,##0.00||"
      Text            =   "1,234,567,891.25"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   315
      Index           =   4
      Left            =   8160
      MaxLength       =   12
      TabIndex        =   6
      Tag             =   "Precio|N|N|0|999999.0000|slitick|precioiv|###,##0.0000|N|"
      Text            =   "123,456.7879"
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   9120
      MaxLength       =   16
      TabIndex        =   7
      Tag             =   "Importe|N|N|||slitick|importel|#,###,###,##0.00|N|"
      Text            =   "Importe"
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   315
      Index           =   2
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   4
      Tag             =   "Nombre artículo|T|N|||slitick|nomartic|||"
      Text            =   "nomArtic"
      Top             =   6600
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3240
      TabIndex        =   12
      ToolTipText     =   "Buscar artículo"
      Top             =   6600
      Visible         =   0   'False
      Width           =   195
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
      Left            =   11040
      TabIndex        =   9
      Top             =   8400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9720
      Top             =   9000
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   1164
      ButtonWidth     =   1746
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Venta  F4"
            Object.ToolTipText     =   "Nueva Venta"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Borrar  F6"
            Object.ToolTipText     =   "Eliminar Venta"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Total  F5"
            Object.ToolTipText     =   "Total Venta"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Termi."
            Object.ToolTipText     =   "Ventas de otros terminales"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar F2"
            Object.ToolTipText     =   "Buscar Artículo"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Linea  F7"
            Object.ToolTipText     =   "Eliminar linea"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Modifi F8"
            Object.ToolTipText     =   "Modificar linea"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Revisar F9"
            Object.ToolTipText     =   "Revisar ventas dia"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Lotes F11"
            Object.ToolTipText     =   "Lote fitosanitario"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cestas"
            Object.ToolTipText     =   "Leer cestas"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   9240
      Top             =   8640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   12600
      TabIndex        =   10
      Top             =   8400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacTPVEnt.frx":0016
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   1650
      Left            =   4680
      Top             =   1080
      Width           =   2955
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nueva Venta"
         HelpContextID   =   2
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar Venta"
         HelpContextID   =   2
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnTotal 
         Caption         =   "&Total Venta"
         HelpContextID   =   2
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnTraerVenta 
         Caption         =   "&Otros termi."
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineasBuscar 
         Caption         =   "&Buscar Artículo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnLineasElim 
         Caption         =   "Eliminar &Linea"
         HelpContextID   =   2
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnModificarLinea 
         Caption         =   "&Modificar linea"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnRevisarVentasDia 
         Caption         =   "Revisión ventas día"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnLotefitosanitarios 
         Caption         =   "Nº lote fitosanitario"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnCestas 
         Caption         =   "Cestas"
      End
      Begin VB.Menu mnbarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacTPVEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public NomrePC_conectado As String
                              

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents FrmArt As frmAlmArticu2   'Form Articulos LLEVA BUSQUEDA LOTES
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2   'Form clientes (busquedas)
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmA As frmBasico2
Attribute frmA.VB_VarHelpID = -1


'Pantalla para traer la venta de otro terminal
Private WithEvents frmTraerVen As frmFacTPVTraerVen
Attribute frmTraerVen.VB_VarHelpID = -1


Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------


Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas


Dim PrimeraVez As Boolean


'Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

'Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String
'WHERE para seleccionar una venta de otro terminal
Private CadSelVenta As String

Private Ordenacion As String 'Para el ORDER BY de la consulta
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private TieneAPP_Cestas As Boolean


Dim Sql As String
Dim i As Integer


Dim NumTermi As Byte

Dim CodTraba As String 'trabajador conectado
Dim NomTraba As String
Dim codAlmac As String 'almacen por defecto del trabajador



Private Sub cboNumVenta_Click()
   Modo = 2
   ModificaLineas = 0
   
   If PosicionarData Then
        CargaGrid Me.DataGrid1, Me.Data2, True
        
        If Me.Data1.Recordset.AbsolutePosition > 0 Then
            If Not Data1.Recordset.EOF Then
                'poner el total
                Me.Label1(6).Caption = Format(DevuelveDesdeBDNew(conAri, "scaven", "imptotal", "numtermi", CStr(Data1.Recordset!NumTermi), "N", , "numventa", CStr(Data1.Recordset!NumVenta), "N", "fecventa", CStr(Data1.Recordset!fecventa), "F"), FormatoImporte)
                'poner el cliente
                Me.Text1(0).Text = Data1.Recordset!codClien
                Me.Text1(0).Text = Format(Text1(0).Text, "000000")
                Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                'poner las observaciones de la venta
                Me.Text1(1).Text = DBLet(Data1.Recordset!observa1, "T")
                'poner direc./ departamento
                If Not IsNull(Data1.Recordset!CodDirec) Then
                    Text1(2).Text = DBLet(Data1.Recordset!CodDirec, "N")
                    Me.Text1(2).Text = Format(Text1(2).Text, "000")
                    Text2(2).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "coddirec", Text1(2).Text, "N", , "codclien", Text1(0).Text, "N")
                Else
                    Text1(2).Text = ""
                    Text2(2).Text = ""
                End If
'                Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "sdirec", "nomdirec", "coddirec", "cliente", "N")
            End If
                
            If vParamTPV.HayVisor Then EnviarVisorPuerto Label1(7).Caption, Label1(8).Caption, Label1(5).Caption, Label1(6).Caption
                
'                If Data2.Recordset.EditMode = adEditNone Then
'                    If Me.Data2.Recordset.AbsolutePosition <> Me.Data2.Recordset.RecordCount Then
                    If Data2.Recordset.RecordCount > DataGrid1.VisibleRows Then
                        Data2.Recordset.MoveLast
                        If DataGrid1.Row < 0 Then DataGrid1.Row = DataGrid1.VisibleRows - 1
                    End If
'                End If
                
                BotonAnyadirLinea
        End If
    End If
End Sub











Private Sub cmdAceptar_Click()
    'Hago el update
    conn.BeginTrans
    If HacerUpdate Then
        conn.CommitTrans
        CommitConexion
        CargaGrid2 DataGrid1, Data2
        'Me pongo a insertar otra
        BotonAnyadirLinea
    Else
        conn.RollbackTrans
    End If
End Sub

Private Sub cmdAux_Click(Index As Integer)
'    Select Case Index
'        Case 1 'Busqueda de Cod. Artic
'            Set frmArt = New frmAlmArticulos
'            frmArt.DatosADevolverBusqueda = "1" 'Poner en Modo busqueda
'            frmArt.Show vbModal
'            Set frmArt = Nothing
'            txtAux_LostFocus (1)
'    End Select
'    PonerFoco txtAux(Index)
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Ventas (scaven)
' y los registros correspondientes de las tablas de lineas (sliven)
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    Cad = "Va a eliminar la venta:         " & vbCrLf & vbCrLf
    Cad = Cad & "     Nº venta:  " & Me.Data1.Recordset!NumVenta & vbCrLf
    Cad = Cad & "     Fecha:  " & Format(Me.Data1.Recordset!fecventa, "dd-mm-yyyy")
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "
       
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton3) <> vbYes Then Exit Sub
    
    
    Cad = vbCrLf & vbCrLf & "Va a eliminar la venta:         " & vbCrLf & vbCrLf
    Cad = Cad & "     Nº venta:  " & Me.Data1.Recordset!NumVenta & vbCrLf
    Cad = Cad & "     Fecha:  " & Format(Me.Data1.Recordset!fecventa, "dd-mm-yyyy")
    Cad = Cad & vbCrLf & vbCrLf & Space(15) & "¿ S E G U R O ? " & vbCrLf & vbCrLf
    Cad = String(30, "*") & Cad & String(30, "*")
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
    
    
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition

        If Not Eliminar() Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            PosicionarDataTrasEliminar
            'Volvemos a cargar las ventas que quedan ahora
            PonerVentasAbiertas
            cargaComboVentas
            If Not Data1.Recordset.EOF Then
                PosicionarComboVentas (Me.Data1.Recordset!NumVenta)
            Else
                'Limpiar los totales
                ReiniciarVisor
                cmdCancelar_Click
                Text1(0).Text = ""
                Text2(0).Text = ""
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Venta.", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea de la venta. Tabla: sliven

    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.BOF Then Exit Sub
    If DBLet(Data2.Recordset!codArtic, "T") = "" Then Exit Sub
    Sql = "¿Seguro que desea eliminar la línea de venta?     " & vbCrLf
    Sql = Sql & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    Sql = Sql & vbCrLf & "Importe:  " & Data2.Recordset!ImporteL
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        If EliminarLinea Then
'            ModificaLineas = 0
            CargaGrid2 DataGrid1, Data2
             
            If Not SituarDataTrasEliminar(Data2, NumRegElim, True) Then
                If Data2.Recordset.RecordCount = 0 Then
                'Elimine el ultimo registro
                    ReiniciarVisor
                    BotonAnyadirLinea
                End If
            End If
            If DataGrid1.Enabled Then DataGrid1.SetFocus
        End If
'        CancelaADODC
    End If

EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar linea Venta.", Err.Description
End Sub



Private Sub cmdCancelar_Click()
    On Error Resume Next
    Select Case Modo
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            FrameFito.visible = True
            If Not Data2.Recordset.EOF Then Me.Data2.Recordset.MoveLast
            If ModificaLineas = 1 Then
                ModificaLineas = 0
                Me.DataGrid1.Enabled = True
            End If
            Me.Label1(6).Caption = Format(ObtenerImporteTotal(False), FormatoImporte)
            PonerFocoGrid DataGrid1
'            If DataGrid1.Enabled Then DataGrid1.SetFocus
'            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            Me.cmdAceptar.visible = False
            Me.cmdCancelar.visible = False
           
    End Select
    PonerModo 2
    
    If Err.Number <> 0 Then MuestraError Err.Number, "", Err.Description
End Sub


Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Select Case KeyCode
'        Case 37 'Flecha izquierda

'        Case 38 'Flecha Arriba

        Case 40 'flecha Abajo
            'Pasar a la siguiente linea
            'si estamos en la ultima linea añadimos nueva linea
            If Me.Data2.Recordset.AbsolutePosition = Me.Data2.Recordset.RecordCount Then
                KeyCode = 0
                BotonAnyadirLinea
            End If
    End Select

    If Err.Number <> 0 Then MuestraError Err.Number, "", Err.Description
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then 'ESC
        cmdCancelar_Click
        Unload Me
    End If
End Sub



Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Data2.Recordset.EOF Or Me.Data2.Recordset.BOF Then
        Me.Label1(7).Caption = ""
        Me.Label1(8).Caption = ""
        txtFito2(0).Text = ""
    Else
        'Actualizar el total de linea
        If (ModificaLineas = 0) Or (ModificaLineas = 1 And Data2.Recordset.AbsolutePosition = Data2.Recordset.RecordCount And Me.Data2.Recordset!NomArtic <> "") Then
            Me.Label1(7).Caption = Mid(Me.Data2.Recordset!NomArtic, 1, 19)
            Me.Label1(8).Caption = Format(Me.Data2.Recordset!ImporteL, FormatoImporte)
            
            CargaLotesArticulos
            
        End If
    End If
    If vParamTPV.HayVisor Then
        EnviarVisorPuerto Label1(7).Caption, Label1(8).Caption, "TOTAL", Label1(6).Caption
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If vParamTPV Is Nothing Then Unload Me
    If CodTraba = "" Then Unload Me
    
    If PrimeraVez Then
        PrimeraVez = False
        'Cargo en tmpinformes los IVAS para poder mostrarlos en el frmarticulo2
        CargarIVas
        Me.cboNumVenta.ListIndex = Me.cboNumVenta.ListCount - 1
    End If
End Sub

Private Sub CargarIVas()
    On Error GoTo EC
    
    conn.Execute "DELETE frOM tmpinformes WHERE codusu = " & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from tiposiva", ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    'insert into `tmpinformes` (`codusu`,`codigo1`,`porcen1`)
    CadSelVenta = ""
    While Not miRsAux.EOF
        CadSelVenta = CadSelVenta & ", (" & vUsu.Codigo & "," & miRsAux!Codigiva & "," & TransformaComasPuntos(CStr(miRsAux!PorceIVA)) & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If CadSelVenta <> "" Then
        CadSelVenta = Mid(CadSelVenta, 2)
        CadSelVenta = "insert into `tmpinformes` (`codusu`,`codigo1`,`porcen1`) VALUES " & CadSelVenta
        conn.Execute CadSelVenta
    End If
    
    
    
EC:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    CadSelVenta = ""
End Sub

Private Sub Form_Load()


    'Icono del formulario
    Me.Icon = frmPpal.Icon

    Me.imgBuscar(0).Picture = frmPpal.ImgListPpal.ListImages(17).Picture
    Me.imgBuscar(1).Picture = frmPpal.ImgListPpal.ListImages(17).Picture


    

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.ImageListTPV
        .Buttons(2).Image = 1   'Nueva Venta
        .Buttons(4).Image = 2   'Eliminar Venta
        .Buttons(6).Image = 3   'TOTALIZAR venta
        .Buttons(8).Image = 7   'Traer venta de otro terminal

        .Buttons(12).Image = 4   'buscar articulo
        .Buttons(14).Image = 5   'Eliminar linea venta
        
        .Buttons(16).Image = 9   'Revisar ventas dia
        .Buttons(18).Image = 8   'Revisar ventas dia
        
        
        .Buttons(21).Image = 10  'Lotes fito
        .Buttons(21).visible = vParamAplic.ManipuladorFitosanitarios2
        Me.mnLotefitosanitarios.visible = vParamAplic.ManipuladorFitosanitarios2
        
        .Buttons(23).Image = 11
        .Buttons(23).Enabled = False
        .Buttons(26).Image = 6  'Salir
    End With


    txtFito2(0).visible = vParamAplic.ManipuladorFitosanitarios2 Or vParamAplic.Ariagro <> ""

    PrimeraVez = True
    LimpiarCampos   'Limpia los campos TextBox

   
    'Terminal con el que trabajaremos, leemos el nombre del ordenador
'    SQL = ComputerName
    Sql = Me.NomrePC_conectado
    Sql = DevuelveDesdeBDNew(conAri, "spatpvt", "numtermi", "nombrepc", Sql, "T")
'    MsgBox "Nº de terminal para " & Me.NomrePC_conectado & ": " & SQL, vbInformation
    If Not IsNumeric(Sql) Then
        MsgBox "No se ha configurado el terminal de venta." & vbCrLf & "Debe configurar primero los parámetros del TPV.", vbExclamation
        Exit Sub
    End If
    NumTermi = CInt(Sql)
   
    
    'Leemos los parametros del TPV
    Set vParamTPV = New CParamTPV
    If vParamTPV.Leer() = 1 Then
        Sql = "No se han podido cargar los Parámetros generales del TPV." & vbCrLf
        MsgBox Sql & "Debe configurar la aplicación.", vbExclamation
        Set vParamTPV = Nothing
        Unload Me
    ElseIf vParamTPV.Leer2(CStr(NumTermi)) = 1 Then
        Sql = "No se han podido cargar los Parámetros del terminal TPV." & vbCrLf
        MsgBox Sql & "Debe configurar la aplicación.", vbExclamation
        Set vParamTPV = Nothing
        Unload Me
    End If
    

    'Poner el trabajador que esta conectado
    CodTraba = PonerTrabajadorConectado(NomTraba)
    If CodTraba = "" Then
        Sql = "No se ha encontrado ningún trabajador con ese login." & vbCrLf
        Sql = Sql & "Compruebe que el trabajador tiene asignado su login."
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    
    'Almacen por defecto el del trabajador
    If CodTraba <> "" Then
        codAlmac = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", CodTraba, "N")
        If codAlmac = "0" Or codAlmac = "" Then codAlmac = DevuelveDesdeBDNew(conAri, "salmpr", "min(codalmac)", "", "")
    Else
        codAlmac = DevuelveDesdeBDNew(conAri, "salmpr", "min(codalmac)", "", "")
    End If
    


    'Compruebo si tiene mas de una forma de pago
    '
    Set miRsAux = New ADODB.Recordset
    Sql = "SELECT count(*) from sforpa"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") = 1 Then Sql = ""
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If Sql = "" Then vParamTPV.FormaPagoUnica = True


    NombreTabla = "scaven" 'tabla cabecera de la venta
    NomTablaLineas = "sliven" 'Tabla lineas de venta
    Ordenacion = " ORDER BY numtermi, numventa, fecventa "
    CadSelVenta = "" 'selecciona la venta de otro terminal

    'Poner los grid sin apuntar a nada
    LimpiarDataGrids

    'Cargamos el Data1 con las cabeceras
    Data1.ConnectionString = conn

    'Inicializar Cabecera de Caja
    Me.Label1(2).Caption = Format(Now, "dd-mm-yy  hh:mm")


    'Abrir Visor en puerto serie
    AbrirVisorPuerto

    'Iniciar Totales y los envia al visor del puerto
    ReiniciarVisor

    'Comprobar si hay ventas abiertas
    PonerVentasAbiertas

    'Ver si tiene nuestra APP de cestas
    EstableceCestas
    
    If TieneAPP_Cestas Then
        Toolbar1.Buttons(23).Enabled = True
        Me.mnCestas.Enabled = True
    Else
            Me.mnCestas.Enabled = False
    End If
    
    
    If CInt(Me.Label1(4).Caption) > 0 Then
        cargaCabeceras True
    Else
        cargaCabeceras False
    End If

    cargaComboVentas
        
End Sub



Private Sub PonerVentasAbiertas()
    Sql = "SELECT count(*) FROM " & NombreTabla & " WHERE numtermi=" & NumTermi
    If CadSelVenta <> "" Then Sql = Sql & " OR " & CadSelVenta
    Me.Label1(4).Caption = CStr(NumRegistros(Sql)) 'nº ventas abiertas
End Sub


Private Sub cargaComboVentas()
Dim RS As ADODB.Recordset
Dim N As Integer
    On Error GoTo ECargaCombo
    
    Me.cboNumVenta.visible = False
    Me.cboNumVenta.Clear

    Sql = "SELECT * FROM " & NombreTabla & " WHERE numtermi=" & NumTermi
    If CadSelVenta <> "" Then Sql = Sql & " OR " & CadSelVenta
    Sql = Sql & " ORDER BY numventa"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    i = 0
    Sql = "                      "
    While Not RS.EOF
        N = 15 - Len(CStr(RS!NumVenta))
            
        Me.cboNumVenta.AddItem Right(Sql & RS!NumVenta, N)
        cboNumVenta.ItemData(cboNumVenta.NewIndex) = i
        Me.cboNumVenta.RightToLeft = True
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    Me.cboNumVenta.visible = True
    
ECargaCombo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos del combo.", Err.Description
End Sub



Private Sub cargaCabeceras(enlaza As Boolean)

    'ASignamos un SQL al DATA1
    If Not enlaza Then
        CadenaConsulta = "Select * from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " where numventa=-1"
    Else
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE numtermi=" & NumTermi
        If CadSelVenta <> "" Then CadenaConsulta = CadenaConsulta & " OR " & CadSelVenta
        CadenaConsulta = CadenaConsulta & " ORDER BY numtermi,numventa,fecventa "
    End If
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
End Sub



Private Sub LimpiarCampos()
    On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Error1
    
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        'NO SE PUEDEN QUEDAR VENTAS ABIERTAS
        
        Sql = DevuelveDesdeBD(conAri, "count(*)", "sliven", "numtermi", CStr(NumTermi))
        If Val(Sql) > 0 Then
            MsgBox "Tiene ventas abiertas sin realizar tickets", vbCritical
            Cancel = 1
            
        Else
            Sql = "DELETE FROM scaven WHERE numtermi=" & NumTermi & " AND NOT (numtermi,numventa,fecventa) "
            Sql = Sql & " IN (select numtermi,numventa,fecventa from sliven )"
            ejecutar Sql, True
        End If
            
    End If
    
    'cerramos el puerto serie del visor
    If Cancel = 0 Then
        If Not vParamTPV Is Nothing Then
            If vParamTPV.HayVisor Then
                If Me.MSComm1.PortOpen Then
                    Me.MSComm1.Output = Mid(vEmpresa.nomempre & Space(40), 1, 40)
                    Me.MSComm1.PortOpen = False
                End If
            End If
            
            Set vParamTPV = Nothing
        End If
    End If
Error1:
    If Err.Number Then Err.Clear
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    DatoDevueltoArticulo CadenaSeleccion
End Sub
Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    DatoDevueltoArticulo CadenaSeleccion
End Sub
Private Sub DatoDevueltoArticulo(CadenaSeleccion As String)
    If vParamAplic.NumeroInstalacion = vbAlzira Then
        'Solo lo lleva ALZIRA
        If Mid(CadenaSeleccion, 1, 18) = "##LOTESAGRUPADOS##" Then
            'Ha seleccionado un LOTE NAVIDAD
            CadenaSeleccion = Mid(CadenaSeleccion, 19) 'llevará empipado los idLote y UDS
            InsertaLineasArticulosAgrupados
            CadenaSeleccion = ""
            Exit Sub
            'Cargamos grid
        End If
    End If
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1)
    HaDevueltoDatos = True

End Sub



Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'clientes (busquedas)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod clien
    Text1(0).Text = Format(Text1(0).Text, "000000")
    
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom clien
End Sub



Private Sub frmTraerVen_CargarVenta(cadSel As String, nVen As Long)
 'Traer venta de otro terminal
    CadSelVenta = cadSel
    
    If nVen < 0 Then
        'He mamdado a buscar las ventas del dia
        'Nven:      -1 Imprimir tick
        '           -2 Ver lineas ticket
        CadSelVenta = Abs(nVen) & "|" & cadSel & "|"
    
    Else
        
        PonerVentasAbiertas
        
        cargaCabeceras True
        cargaComboVentas
        
        
        PosicionarComboDes Me.cboNumVenta, CStr(nVen)

    End If

End Sub


Private Sub imgBuscar_Click(Index As Integer)

    On Error GoTo ErrImg
    
    If Data1.Recordset.EOF Then Exit Sub
    
    Select Case Index
        Case 0 'clientes
            If Text1(0).Text <> "" And Me.Text2(0).Text = "" Then Text1(0).Text = ""
            Set frmCli = New frmBasico2
            AyudaClientes frmCli, Text1(0)
            Set frmCli = Nothing
            PonerFoco Text1(0)
            
'            If CLng(Data1.Recordset!CodClien) <> CLng(Text1(0).Text) Then
'             'si se ha cambiado el cliente actualizar la cabecera venta (scaven)
'                ModificarVenta
'            End If
            
'           txtAux_LostFocus (1)

        Case 1 'Departamento
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(0).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                MandaBusquedaPrevia_dpto " codclien= " & Val(Text1(0).Text)
'                Indice = 12
             End If
    End Select
    Exit Sub
    
ErrImg:
    MuestraError Err.Number, "Busqueda de clientes", Err.Description
End Sub

Private Sub mnCestas_Click()
    'Leer cestas
    If Modo < 2 Then Exit Sub
    
    
    'Veremos si  el cliente tiene una cesta abierta
    Sql = "cestas inner join cestas_lineas on cestas.cestaId= cestas_lineas.cestaId "
    Sql = DevuelveDesdeBD(conAri, "count(*)", Sql, "codclien", Text1(0).Text)
    If Val(Sql) = 0 Then
        MsgBox "El cliente no tiene ninguna cesta abierta", vbExclamation
        Exit Sub
    End If
    CadenaDesdeOtroForm = ""
    frmFacTPV_Cesta.Text1(0).Text = Text1(0).Text
    frmFacTPV_Cesta.Text2(0).Text = Text2(0).Text
    frmFacTPV_Cesta.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        
        If InsertaLineasArticulosCesta Then
            CargaGrid2 DataGrid1, Data2
            BotonAnyadirLinea
            
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnEliminar_Click()
'Elimina una venta
    BotonEliminar
End Sub



Private Sub mnLineasBuscar_Click()
Dim B As Boolean

    If Data1.Recordset.EOF Then Exit Sub
'    If ModificaLineas = 0 Then Exit Sub

    B = (Screen.ActiveControl.Name = "Text1")
    If B Then B = (Screen.ActiveControl.Index = 0) Or (Screen.ActiveControl.Index = 2)
    
    If B Then
        If (Screen.ActiveControl.Index = 0) Then
            imgBuscar_Click (0)
        ElseIf (Screen.ActiveControl.Index = 2) Then
            imgBuscar_Click (1)
        End If
    Else
        
        
        
        HaDevueltoDatos = False
        
        If vParamAplic.NumeroInstalacion = vbAlzira Then
            Set FrmArt = New frmAlmArticu2
            FrmArt.DesdeTPV = True
            FrmArt.DatosADevolverBusqueda = "0"
            FrmArt.parAlmacen = codAlmac
            FrmArt.Show vbModal
            Set FrmArt = Nothing
        Else
        
            Set frmA = New frmBasico2
            AyudaArticulos frmA, txtAux(2), , codAlmac, True
            Set frmA = Nothing
        End If
        
        
        
        If HaDevueltoDatos = True Then
            PonerFoco txtAux(2)
            txtAux_LostFocus (2)
        End If
    End If
End Sub

Private Sub mnLineasElim_Click()
'Elimina una linea de venta
    BotonEliminarLinea
End Sub

Private Sub mnLotefitosanitarios_Click()
    If Modo <> 2 Then Exit Sub
    If ModificaLineas <> 0 Then Exit Sub
    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.BOF Then Exit Sub
    If DBLet(Data2.Recordset!numLote, "T") = "" Then Exit Sub
    ModificaLote
End Sub

Private Sub mnModificarLinea_Click()
    BotonModifcarLinea
End Sub

Private Sub mnNuevo_Click()
'Inicia una nueva venta
    BotonAnyadir
End Sub


Private Sub mnRevisarVentasDia_Click()
    RevisarEntradasDia
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub



'Private Sub MandaBusquedaPrevia(cadB As String)
''Carga el formulario frmBuscaGrid con los valores correspondientes
'Dim cad As String
'Dim tabla As String
'Dim Titulo As String
'Dim devuelve As String
'
'    'Llamamos a al form
'    '##A mano
'    cad = ""
'    'tabla = "sartic INNER JOIN salmac ON sartic.codartic=salmac.codartic "
'
'
'     'SELECT `sartic`.`codartic`, `salmac`.`codalmac`, `slista`.`precioac`
'    tabla = "(`sartic` `sartic` LEFT OUTER JOIN `salmac` `salmac` ON `sartic`.`codartic`=`salmac`.`codartic`)"
'    tabla = tabla & " LEFT OUTER JOIN `slista` `slista` ON `sartic`.`codartic`=`slista`.`codartic`"
'
'
'
'    Titulo = "Artículos"
'    devuelve = "0|1|2|"
'    cad = cad & "Cod. Artic.|sartic|codartic|T||22·"
'    cad = cad & "Des. Artic.|sartic|nomartic|T||57·"
''    cad = cad & "Precio|sartic|preciove|N|###,##0.0000|15·"
'    cad = cad & "Stock|salmac|canstock|N|#,###,###,##0.00|8·"
'    cad = cad & "Precio|slista|precioac|N|#,###,###,##0.00|12·"
'
'    If cadB = "" Then
'        cadB = " codalmac = " & codAlmac
'    Else
'        cadB = cadB & " AND codalmac=" & codAlmac
'    End If
'    'La tarifa a buscar SIEMPRE es la de parametros
'    cadB = cadB & " AND ( codlista = " & vParamAplic.CodTarifa & " OR codlista is null)"
'
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = devuelve
''        frmB.vDevuelve = devuelve
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 1
'        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
''        If HaDevueltoDatos Then
''''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''''                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            PonerFoco Text1(kCampo)
'        'End If
'        If ModificaLineas = 1 Then
'            If txtAux(2).Text <> "" Then
'                PonerFoco txtAux(2)
'                txtAux_LostFocus (2)
'            Else
'                PonerFoco txtAux(1)
'            End If
'        End If
'    End If
'    Screen.MousePointer = vbDefault
'End Sub



Private Sub MandaBusquedaPrevia_dpto(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
'para busquedas de direc./depatamentos del cliente
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim devuelve As String
Dim Desc As String

    'Llamamos a al form
    '##A mano
    Cad = ""
    
    If vParamAplic.HayDeparNuevo = 1 Then
        Titulo = "Dptos Cliente: "
        Desc = "Dpto."
    ElseIf vParamAplic.HayDeparNuevo = 0 Then
        Titulo = "Direc. Cliente: "
        Desc = "Direc."
    Else
        Titulo = "Obras Cliente: "
        Desc = "Obra"
    End If
    Titulo = Titulo & Text1(0).Text & " - " & Text2(0).Text
    Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|15·"
    Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||65·"
    tabla = "sdirec"
    devuelve = "0|1|"
    
    
    
   
'    Tabla = "sartic INNER JOIN salmac ON sartic.codartic=salmac.codartic "
'    Titulo = "Artículos"
'    devuelve = "0|1|2|"
'    cad = cad & "Cod. Artic.|sartic|codartic|T||25·"
'    cad = cad & "Des. Artic.|sartic|nomartic|T||60·"
'    cad = cad & "Stock|salmac|canstock|N|#,###,###,##0.00|15·"
    
'    If cadB = "" Then
'        cadB = " codalmac = " & codAlmac
'    Else
'        cadB = cadB & " AND codalmac=" & codAlmac
'    End If
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = devuelve
'        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        '#
        frmB.Label1.FontSize = 11
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
        
'        If ModificaLineas = 1 Then
'            If txtAux(2).Text <> "" Then
'                PonerFoco txtAux(2)
'                txtAux_LostFocus (2)
'            Else
'                PonerFoco txtAux(1)
'            End If
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
'            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
'        PonerModo 2
'        PonerCampos
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
    
    Modo = Kmodo
    
    
    B = (Modo = 3) Or (Modo = 4)
   
    'Si no es modo lineas Boquear los TxtAux
    For i = 1 To 5
        BloquearTxt txtAux(i), Not B
        txtAux(i).visible = B
    Next i
    txtAux(12).visible = txtAux(5).visible
    txtAux2(2).visible = B
    'El Importe siempre bloqueadoç
    BloquearTxt txtAux(5), True
    
    DataGrid1.Enabled = (Modo = 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    If Modo = 2 Then
        'Me garantizo que los
        'botones NO se vean
        Me.cmdAceptar.visible = False
        Me.cmdCancelar.visible = False
    End If
    
    
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
'    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    Exit Sub
    
EPonerModo:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
    B = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function
    
    
    
    
    
    DatosOk = B
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea(ByRef ArticuloConLotes As Boolean) As Boolean
'Antes de insertar una linea de venta comprueba que los datos son OK
Dim B As Boolean
Dim CadenaInsertTmpLotes As String
Dim CuantosLotesDistintos As Byte
Dim NumerloLote As String
Dim CantidadEnTotal As Currency


    On Error GoTo EDatosOkLinea

    B = CompForm(Me, 3)
    
    
    If B Then
        If txtAux2(2).Text = "" And txtAux(2).Text <> "" Then
            MsgBox "Codigo artículo incorrecto", vbExclamation
            DatosOkLinea = False
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    
    
    '------------------------------------------------------
    '------------------------------------------------------
    '------------------------------------------------------
    'Numero de LOTE fitosanitarios
    ArticuloConLotes = False
    
    Sql = DevuelveDesdeBD(conAri, "numserie", "sartic", "codartic", txtAux(2).Text, "T")
    If Sql <> "" Then
        ArticuloConLotes = True
        'Los que no lleven el nuevo controlo sigue como antes
        If Not vParamAplic.ManipuladorFitosanitarios2 Then
            'OK. Salimos YA
            DatosOkLinea = True
            Exit Function
        End If
        
        
        
        CadenaInsertTmpLotes = ""
        
        
        
        
        
        Sql = "select numlotes,fecentra,Codartic,canentra - vendida"
        Sql = Sql & "  disponible from slotes where "
        Sql = Sql & " codartic=" & DBSet(txtAux(2).Text, "T") & " and canentra - vendida  >0 order by fecentra "
      
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        
        NumerloLote = ""
        
        If Not miRsAux.EOF Then
            CuantosLotesDistintos = 1
            NumerloLote = miRsAux!numlotes
            CantidadEnTotal = 0
            
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                If miRsAux!numlotes <> NumerloLote Then
                    'Otro lote. No controlaremos nada
                    CuantosLotesDistintos = CuantosLotesDistintos + 1
                Else
                    CantidadEnTotal = CantidadEnTotal + miRsAux!disponible
                End If
                'insert into tmpnlotes(codusu,numlinea,fechaalb,codprove,cantidad,numlotes)
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!fecentra, "F")
                'CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible * 100, "N")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ",0," & DBSet(miRsAux!numlotes, "T") & ")"
                Sql = miRsAux!disponible 'por si solo hay uno
               
                miRsAux.MoveNext
               
            Wend
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    Else
        'Datos OK linea. =True
        'Nos salimos, pq como no lleva lotes no compruebo nada
        DatosOkLinea = True
        Exit Function
        
    End If
        
        
    
    'Si hay mas de uno mostraremos cual y cuanto puede coger
    If NumRegElim = 0 Then
        MsgBox "Ningun lote disponible para el artículo", vbExclamation
        B = False
    Else
        'Este sera el numero de lote asignadao
        'Con lo cual, que haremos, pondremos
        
        If NumRegElim = 1 Then
            If CCur(Sql) < ImporteFormateado(txtAux(3).Text) Then
                MsgBox "Cantidad en el lote insuficiente:" & Sql, vbExclamation
                B = False
            Else
                'Donde va la cantidad asignada en el SQL es en : ,0,'
                'Luego reeplazo esto por la cantidad del albaran
                Sql = TransformaComasPuntos(CStr(ImporteFormateado(txtAux(3).Text)))
                CadenaInsertTmpLotes = Replace(CadenaInsertTmpLotes, ",0,'", "," & Sql & ",'")
                
            End If
        Else
            'Hay mas de un LOTE - Fecha entrada
            'Veremos si por lo menos es el mismo lote
            'Si es el mismo lote reasignaremos las cantidades
            If CuantosLotesDistintos = 1 Then
                'Hay mas de un lote pero de dsitintas fechas de entrada
                'Veremos i hay suficiente  o no
                If CantidadEnTotal < ImporteFormateado(txtAux(3).Text) Then
                    MsgBox "Cantidad en el lote insuficiente:" & CantidadEnTotal & "(+)", vbExclamation
                    B = False
                Else
                    'Hay suficiente en este LOTE. Volvemos a abri , PARA este lote y volvemos a cargar el SQL
                    Sql = "select numlotes,fecentra,Codartic,canentra - vendida"
                    Sql = Sql & "  disponible from slotes where "
                    Sql = Sql & " codartic=" & DBSet(txtAux(2).Text, "T") & " and canentra - vendida  >0"
                    Sql = Sql & " AND numlotes= " & DBSet(NumerloLote, "T") & " order by fecentra "
                    Set miRsAux = New ADODB.Recordset
                    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    CadenaInsertTmpLotes = "" 'Vamos a volver a cargar el SQL de insert in lotes
                    CantidadEnTotal = ImporteFormateado(txtAux(3).Text)
                    NumRegElim = 0
                    While Not miRsAux.EOF
                        NumRegElim = NumRegElim + 1
                         CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
                         CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!fecentra, "F")
                         'CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T")
                         CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible * 100, "N")
                         
                         If CantidadEnTotal > miRsAux!disponible Then
                             CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible, "N")
                             CantidadEnTotal = CantidadEnTotal - miRsAux!disponible
                             CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!numlotes, "T") & ")"
                             miRsAux.MoveNext
                         Else
                            CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(CantidadEnTotal, "N")
                            CantidadEnTotal = 0
                            CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!numlotes, "T") & ")"
                            
                            While Not miRsAux.EOF
                                miRsAux.MoveNext
                            Wend
                         End If
                         'Vuevlo a dejar numregelim por lo menos a DOS, para que no vea que es lote unico
                         NumRegElim = 2
                    Wend
                    miRsAux.Close
                    Set miRsAux = Nothing
                         
                End If 'Cantidad suficoente
            End If  'mas de un lote
        End If
    
        'Mas de un  lote disponible
        Screen.MousePointer = vbHourglass
        
        
        If B Then
            conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo
            Espera 0.3
            CadenaInsertTmpLotes = Mid(CadenaInsertTmpLotes, 2)
            CadenaInsertTmpLotes = "insert into tmpnlotes(codusu,codartic,numlinea,fechaalb,codprove,cantidad,numlotes) VALUES " & CadenaInsertTmpLotes
            conn.Execute CadenaInsertTmpLotes
            
            
            
            If NumRegElim = 1 Then
                CadenaDesdeOtroForm = "OK"
                Espera 0.3
            Else
                If CuantosLotesDistintos = 1 Then
                    CadenaDesdeOtroForm = "OK"
                    Espera 0.3
                Else
                    CadenaDesdeOtroForm = ""
                    frmFacTPVLotes.TotalLineas = ImporteFormateado(txtAux(3).Text)
                    frmFacTPVLotes.NombreArticulo = txtAux2(2).Text
                    frmFacTPVLotes.Show vbModal
                End If
            End If
            If CadenaDesdeOtroForm <> "OK" Then
                'Ha cancelado el proceso
                conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo
                Espera 0.3
            End If
        End If
        Screen.MousePointer = vbDefault
    End If   'Numregeleim0
    
    

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    DatosOkLinea = B
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub mnTotal_Click()
'Dim cCli As CCliente
Dim Sql As String
Dim bModif As Boolean
Dim ArticulosFito As Boolean


    'comprobar q el campo de cliente tiene valor
    If Trim(Text1(0).Text) = "" Then
        MsgBox "El campo Cliente debe tener valor.", vbInformation
        Exit Sub
    End If

    If txtAux(4).Enabled Then
        'Cuando es articulo de varios obtenia totales pero SIN insertar en sliven
        'Al pulsar F5 se salia sin haber insrtado en sliven
        Sql = " numventa=" & Data1.Recordset!NumVenta & " and fecventa=" & DBSet(Data1.Recordset!fecventa, "F") & " AND numtermi"
        
        Sql = DevuelveDesdeBD(conAri, "sum(importel)", "sliven", Sql, CStr(Data1.Recordset!NumTermi), "N")
        If Sql = "" Then Sql = "0"
        Sql = Format(Sql, FormatoImporte)
        If Sql <> Label1(6).Caption Then
            MsgBox "No ha confirmado la ultima linea que estaba insertando", vbExclamation
            PonerFoco txtAux(4)
            Exit Sub
        End If
    End If
    '-- Si el cliente esta bloqueado no permitimos generar ticket/Albaran/Factura
    '-- y no abrimos la pantalla de totales
    '-- si el cliente se ha modificado y no tiene la misma tarifa q para el q se crearon las lineas tampoco
    
    'obtenemos el cliente actual guardado en la cabecera de la venta
    Sql = ""
    If Not Data1.Recordset.EOF Then Sql = Data1.Recordset!codClien
    If Not ClienteOK(Text1(0).Text, Sql, bModif) Then
        Text1(0).Text = Sql
        Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
        PonerFoco Text1(0)
        Exit Sub
    End If
'    If bModif Then ModificarVenta
    
    
    '-- si no hay lineas en la venta no permitimos generar ticket/Albaran/Factura
    '-- y no abrimos la pantalla de totales
    Sql = ObtenerWhereCP(False)
    Sql = Replace(Sql, NombreTabla, NomTablaLineas)
    Sql = "SELECT COUNT(*) FROM " & NomTablaLineas & " WHERE " & Sql
    If Not RegistrosAListar(Sql) > 0 Then
        MsgBox "La venta debe tener líneas para totalizar.", vbExclamation
        Exit Sub
    End If
    
    If vParamTPV.ProhibirTiketsNegativos Then
        If ImporteFormateado(Me.Label1(6).Caption) < 0 Then
            MsgBox "Importe total ticket negativo", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    'Si tiene fitosanitarios, habra que hacer mas controles
    ArticulosFito = False
    If vParamAplic.ManipuladorFitosanitarios2 Then
        If Not ControlesFitosantiarios(ArticulosFito) Then Exit Sub
    End If
    
    
    
    cmdCancelar_Click
    
    Me.Label1(2).Caption = Format(Now, "dd-mm-yy  hh:mm")
'    If CDate(Data1.Recordset!fecventa) <> CDate(Format(Now, "dd/mm/yyyy")) Then
        'Actualizar la fecha de la venta
        ModificarVenta
'    End If
    Volver_A_Cargar_Datos = False  'Nos diara si hemos cerrado la venta o no
    frmFacTPVTotal.ImporteInicial = Me.Label1(6).Caption
    frmFacTPVTotal.cadSel = ObtenerWhereCP(False)
    frmFacTPVTotal.LLevaArticulosFitosanitarios = ArticulosFito
    Load frmFacTPVTotal
    frmFacTPVTotal.Show vbModal

    'SEPTIEMBRE 2015.
    'QUitamos  If frmFacTPVTotal.cadSel = "1" Then 'Se genero correctamente el documento(ticket,alb,factu)
    'Si se ha cerarado la venta, ticket albaran o factura
    'Ponemos a true la variabale global multiproposito Volver_A_Cargar_Datos
    
    If Volver_A_Cargar_Datos Then 'Se genero correctamente el documento(ticket,alb,factu)
        'Refrescar los datos
        'Volvemos a cargar las ventas que quedan ahora
        PonerVentasAbiertas
        cargaCabeceras (True)
        cargaComboVentas
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        ReiniciarVisor
        LimpiarCampos
        Me.cboNumVenta.ListIndex = Me.cboNumVenta.ListCount - 1
    Else
        'Por si acaso a tocado algo
        Data1.Refresh
        PosicionarData
    
    End If
End Sub



Private Sub mnTraerVenta_Click()
'Traer venta de otro terminal
    CadSelVenta = ""
    Set frmTraerVen = New frmFacTPVTraerVen
    frmTraerVen.parNumTermi = NumTermi
    frmTraerVen.Show vbModal
    Set frmTraerVen = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cCli As CCliente
Dim devuelve As String
Dim bModif As Boolean

    devuelve = ""
    
    Select Case Index
        Case 0 'cod CLIENTE
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
                
                If Text2(Index).Text = "" Then
                    'no existe el cliente y no sale del campo
                    PonerFoco Text1(Index)
                Else
                    If Not (Me.Data1.Recordset.EOF And Me.Data1.Recordset.BOF) Then '(RAFA/ALZIRA 31082006) Al poner el foco directamente en el cliente este trozo falla cuando se pierde el foco sin que haya ninguna venta
                        devuelve = Me.Data1.Recordset!codClien
                    End If
                        
                    If ClienteOK(Text1(Index).Text, devuelve, bModif, True) Then
                        If bModif Then ModificarVenta
                    Else
                        Text1(Index).Text = devuelve
                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                        Text1(Index).Text = Format(Text1(Index).Text, "000000")
                        PonerFoco Text1(Index)
                    End If
'                    Set cCli = New CCliente
'                    If cCli.LeerDatos(Text1(Index).Text) Then
'                        'no permitimos ventas a clientes bloqueados
'                        If cCli.ClienteBloqueado Then
'                            PonerFoco Text1(Index)
'                            Set cCli = Nothing
'                            Exit Sub
'                        End If
'
'                        If Not (Me.Data1.Recordset.EOF And Me.Data1.Recordset.BOF) Then '(RAFA/ALZIRA 31082006) Al poner el foco directamente en el cliente este trozo falla cuando se pierde el foco sin que haya ninguna venta
'                            If CLng(Me.Data1.Recordset!CodClien) <> CLng(Text1(0).Text) Then 'OPP
'                                '--- Laura: 11/04/2007
'                                '--- comprobar q la tarifa del nuevo cliente es la misma q la del cliente q
'                                '--- habia antes siempre y cuando haya lineas de precios ya q si no no estariamos
'                                '--- aplicando la tarifa correcta al cliente
'                                If Not Me.Data2.Recordset.EOF Then 'si hay lineas
'                                    'obtener la tarifa del cliente actual
'                                    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", CStr(Me.Data1.Recordset!CodClien), "N")
'                                    If devuelve <> CStr(cCli.Tarifa) Then
'                                        devuelve = "No se puede seleccionar el cliente " & Text1(0).Text & " "
'                                        devuelve = devuelve & "ya que tiene distinta tarifa de precios." & vbCrLf
'                                        devuelve = devuelve & "Seleccione un cliente de la misma tarifa o elimine la venta."
'                                        MsgBox devuelve, vbExclamation, "Comprobar cliente"
'                                        Text1(Index).Text = Me.Data1.Recordset!CodClien
'                                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
'                                        Text1(Index).Text = Format(Text1(Index).Text, "000000")
'                                        PonerFoco Text1(Index)
'                                        Set cCli = Nothing
'                                        Exit Sub
'                                    End If
'                                End If
'                                ModificarVenta
'                            End If
'                        End If
                    
                        'mostrar las observaciones del cliente
'                        cCli.MostrarObservaciones
'                    End If
'                    Set cCli = Nothing
                End If
            Else
                'si el formato no es numerico no lo aceptamos
                Text2(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            
        Case 1 'OBSERVACIONES
            '---- Laura: 06/10/2006
            'Poner el foco en la linea
            If Screen.ActiveControl.Name <> "Text1" Then
                PonerFoco txtAux(1)
                ModificarVenta
            End If
            
        Case 2 'DIREC./DPTO
            If PonerFormatoEntero(Text1(Index)) Then
                Text1(Index).Text = Format(Text1(Index).Text, "000")
                
                'Comprobar que el cliente seleccionada tiene esa direccion
                If PonerDptoEnCliente Then
                    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                    devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(0).Text, "N", , "coddirec", Text1(2).Text, "N")
                    If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
                Else
                    Text1(Index).Text = ""
                    Text2(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
    End Select
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 'Nueva venta
            mnNuevo_Click
            
        Case 4  'Borrar Venta
            mnEliminar_Click
            
        Case 6  'Total Venta
            mnTotal_Click
            
        Case 8 'traer venta de otro terminal
            mnTraerVenta_Click
            
        Case 12 'Buscar articulo
            mnLineasBuscar_Click
            
        Case 14  'Eliminar Linea
            mnLineasElim_Click
                   
        Case 16
            BotonModifcarLinea
                   
        Case 18
            'Revision entradas dia
            mnRevisarVentasDia_Click
        
        Case 21
            mnLotefitosanitarios_Click
        
        Case 23
            mnCestas_Click
        
        Case 26    'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If KeyAscii = 27 Then cmdCancelar_Click
    If cerrar Then Unload Me
End Sub

  

Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)

    On Error GoTo ECargaGrid

'    b = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vData.LockType = adLockOptimistic
    
    CargaGrid2 vDataGrid, vData
    DataGrid1.RowHeight = 360
    
    If Data2.Recordset.RecordCount = 0 Then DataGrid1_RowColChange 1, 1
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String

    On Error GoTo ECargaGrid2

    vDataGrid.ScrollBars = dbgNone
    vData.Refresh


    tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux(1)|T|Cod. EAN|2050|;S|txtAux(2)|T|Código|1850|;S|cmdAux(1)|B||0|;S|txtAux2(2)|T|Desc. Artículo|5650|;"
    tots = tots & "S|txtAux(3)|T|Cantidad|1150|;"
'    If vParamTPV.ModoCalculo = 1 Then
        tots = tots & "S|txtAux(4)|T|Precio|1600|;"
'    Else
'        tots = tots & "S|txtAux(6)|T|Precio|1200|;"
'    End If
    tots = tots & "S|txtAux(5)|T|Importe|1700|;N||||0|;N||||0|;S|txtAux(12)|T|Dto|650|;N||||0|;"
       
    arregla tots, DataGrid1, Me, 360
    
    'cantidad
    Me.DataGrid1.Columns(8).Alignment = dbgRight
'    Me.DataGrid1.Columns(8).NumberFormat = FormatoCantidad & " "
    
    'Precio
    Me.DataGrid1.Columns(9).Alignment = dbgRight
'    Me.DataGrid1.Columns(9).NumberFormat = FormatoPrecio
    
    'Importe
    Me.DataGrid1.Columns(10).Alignment = dbgRight
'    Me.DataGrid1.Columns(10).NumberFormat = FormatoImporte
   
   

   
   
    vDataGrid.ScrollBars = dbgAutomatic
    vDataGrid.Enabled = (Modo <> 3)
    
    Exit Sub
    
ECargaGrid2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        If Index = 5 Then
        End If
    Else
        If Index = 1 And (KeyCode = 37 Or KeyCode = 38) Then Exit Sub
        If Index = 4 And KeyCode = 39 Then Exit Sub
        KEYdownLineas KeyCode
    End If
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
Dim Cant As String

    If KeyAscii = vbKeyTab Then
        If Index = 4 Then
        End If
    End If
    If KeyAscii = 42 Then '*
        Cant = txtAux(Index).Text
        If EsNumerico(Cant) Then
            txtAux(3).Text = Cant
            PonerFormatoDecimal txtAux(3), 1
        End If
        KeyAscii = 0
        txtAux(Index).Text = ""
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 1 'Cod. ARTICULO EAN
            PonerArticuloEan (txtAux(Index).Text)
                
        Case 2 'Cod. Articulo
           PonerArticuloCod (txtAux(Index).Text), False
        
        Case 3 'CANTIDAD
          If PonerFormatoDecimal(txtAux(Index), 1) Then
                'txtAux(5).Text = ObtenerImporteLin2
                'PonerFormatoDecimal txtAux(5), 1
                FijarImportes False
          End If
            
        Case 4 'Precio
             If txtAux(Index).Text <> "" And txtAux(Index).Locked = False Then
                 'Tipo 2: Decimal(10,4)
                If PonerFormatoDecimal(txtAux(Index), 2) Then
                
                    
                    'si es articulo de varios y he modificado el precio del articulo
                    'el precio sin IVA del articulo habra que recalcularlo
                    txtAux(10).Text = ObtenerPrecioSinIVAvarios(txtAux(2).Text, txtAux(Index).Text)
                    
                     'txtAux(5).Text = ObtenerImporteLin2
                    'PonerFormatoDecimal txtAux(5), 1
                    FijarImportes True
                    
                    
                    
                    If Screen.ActiveControl.Name <> "txtAux" Then
                        GuardarLinea
                    ElseIf Screen.ActiveControl.Index > 4 Then
                        GuardarLinea
                    End If
                    'GuardarLinea
                End If
            End If
    End Select
    
'     If (Index = 3 Or Index = 4) Then  'Cant., Precio
'        If txtAux(1).Text = "" Then Exit Sub
''        txtAux(5).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
'        PonerFormatoDecimal txtAux(5), 1
'    End If
End Sub




Private Function Eliminar() As Boolean
Dim B As Boolean
Dim MenError As String
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    vWhere = ObtenerWhereCP(False)
    
    MenError = "Eliminando tablas de venta."
    
    MenError = "Eliminando tablas de venta."
    'La de los campos
    Sql = "DELETE FROM " & NomTablaLineas & "2 WHERE " & Replace(vWhere, NombreTabla, NomTablaLineas & "2")
    conn.Execute Sql
    
    Sql = "DELETE FROM " & NomTablaLineas & " WHERE " & Replace(vWhere, NombreTabla, NomTablaLineas)
    conn.Execute Sql
            
    Sql = "DELETE FROM slivenlotes WHERE " & Replace(vWhere, NombreTabla, "slivenlotes")
    conn.Execute Sql
            
            
    Sql = "DELETE FROM " & NombreTabla & " WHERE " & vWhere
    conn.Execute Sql
    
            
    'Devolvemos contador, si no estamos actualizando
    B = vParamTPV.DevolverContador(NumTermi, Data1.Recordset!NumVenta)
        
FinEliminar:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, MenError, Err.Description
    End If
    If Not B Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function



Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function PosicionarData() As Boolean
Dim vWhere As String

    PosicionarData = False
    
    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
'        vWhere = "numtermi= " & NumTermi & " and "
        vWhere = "numventa= " & Trim(Me.cboNumVenta.List(Me.cboNumVenta.ListIndex))
         If SituarDataMULTI(Data1, vWhere, "") Then
            PosicionarData = True
        Else
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
'        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
'        PonerCadenaBusqueda
    End If
End Function


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    If Me.Data1.Recordset.EOF Then Exit Function
    
    Sql = NombreTabla & ".numtermi= " & Data1.Recordset!NumTermi & " and "
    Sql = Sql & NombreTabla & ".numventa= " & Data1.Recordset!NumVenta & " and " 'Trim(Me.cboNumVenta.List(Me.cboNumVenta.ListIndex))
    Sql = Sql & NombreTabla & ".fecventa= " & DBSet(Data1.Recordset!fecventa, "F")
    
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    Sql = "SELECT numtermi, numventa,fecventa,horventa, numlinea, codigoea,codartic, nomartic, cantidad, precioiv, importel,precioar,codigiva "
    '
    Sql = Sql & " ,if((dto1+dto2)>0,""*"","""") as dto,numlote"
    
    
    
    Sql = Sql & " FROM " & NomTablaLineas
    
    If enlaza Then
        Sql = Sql & " WHERE numtermi=" & Data1.Recordset!NumTermi & " and numventa=" & Data1.Recordset!NumVenta 'Trim(Me.cboNumVenta.List(Me.cboNumVenta.ListIndex))
        Sql = Sql & " AND fecventa=" & DBSet(Data1.Recordset!fecventa, "F")
    Else
        Sql = Sql & " WHERE numventa = -1 "
    End If
    Sql = Sql & Ordenacion & ",numlinea"  '" Order by codtipom, numalbar, numlinea"
    MontaSQLCarga = Sql
End Function



   
Private Function EliminarLinea() As Boolean
Dim B As Boolean

    EliminarLinea = False
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    Sql = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
    

    
    On Error GoTo EEliminarLinea
    conn.BeginTrans
    
    conn.Execute Sql 'Eliminar linea


    If vParamAplic.ManipuladorFitosanitarios2 Then
        Sql = Replace(Sql, "sliven", "slivenlotes")
        conn.Execute Sql 'Eliminar linea fitos
    End If


    Me.Label1(6).Caption = Format(ObtenerImporteTotal(True, B), FormatoImporte)

    B = Not B
    
EEliminarLinea:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea Albaran " & vbCrLf & Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    EliminarLinea = B
End Function




Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
    Else
        LimpiarDataGrids
    End If
End Sub


Private Sub PonerArticuloEan(codArt As String)
    codArt = Trim(codArt)
    If codArt <> "" Then
        '---- [23/09/2009] LAURA : Añadir lineas de Cod. EAN y quitar de la cabecera
'        SQL = DevuelveDesdeBDNew(conAri, "sartic", "codartic", "codigoea", codArt, "T")
        Sql = DevuelveDesdeBDNew(conAri, "sarti3", "codartic", "codigoea", codArt, "T")
        '----
        
        If Sql <> "" Then
            PonerArticuloCod (Sql), False
        Else
            MsgBox "No existe un artículo asociado al cód. EAN: " & codArt, vbInformation
            'Si el campo de codigo no esta vacio
            txtAux(1).Text = ""
            If txtAux(2).Text <> "" Then PonerFoco txtAux(2)
            
        End If
    End If
End Sub



Private Sub PonerArticuloCod(codArt As String, InsertoDesdeCestas As Boolean)
Dim vArtic As CArticulo
Dim Dto1 As String
Dim Dto2 As String
Dim nCantidad As String
Dim nUnids As String
Dim Bloq As Boolean

    On Error GoTo EArticulo
    
    codArt = Trim(codArt)
    
    If codArt <> "" Then
        Set vArtic = New CArticulo
        If vArtic.Existe(codArt) Then
            If vArtic.LeerDatos(codArt) Then
            
            
                vArtic.MostrarStatusArtic Bloq
            
                If Bloq Then
                
                    'DE LEER EL ARTICULO. Ha dado error
                    If txtAux(1).Text <> "" Then PonerFoco txtAux(1)
                    If txtAux(2).Text <> "" Then PonerFoco txtAux(2)
                    txtAux(1).Text = ""
                    txtAux(2).Text = ""
                    txtAux2(2).Text = ""
                    
                Else
                    txtAux(2).Text = vArtic.Codigo
                    txtAux2(2).Text = UCase(vArtic.Nombre)
                    If vArtic.TextoVentas <> "" Then vArtic.MostrarTextoVen
                    Me.Label1(7).Caption = vArtic.Nombre
                    If txtAux(3).Text = "" Then 'cantidad
                        txtAux(3).Text = "1"
                        PonerFormatoDecimal txtAux(3), 1
                    End If
                    
                    'Fijar tipo IVA de aritculo
                    txtAux(7).Text = vArtic.TipoIVA
                    txtAux(11).Text = vArtic.ObtenerPorceIVA
                    
                    nCantidad = txtAux(3).Text
                    'Precio sin dtos
                    txtAux(10).Text = vArtic.ObtPrecioParaCliente2(Text1(0).Text, nCantidad, CStr(Data1.Recordset!fecventa), Dto1, Dto2)
                    
                    'Precio para albaran
                    txtAux(14).Text = txtAux(10).Text
                    
                                    
                    If CCur(nCantidad) <> CCur(txtAux(3).Text) Then
                    'se puede vender por cajas y se insertan 2 lineas
                        nUnids = CCur(txtAux(3).Text) - CCur(nCantidad)
                        txtAux(3).Text = nCantidad
                        PonerFormatoDecimal txtAux(3), 1
                    Else
                        nUnids = ""
                    End If
                    
                    
                    'Estos dos NO son visibles
                    txtAux(8).Text = Dto1
                    txtAux(9).Text = Dto2
                    If Val(Dto1) + Val(Dto2) > 0 Then txtAux(12).Text = "*"
                                
                    ' ---- [21/10/2009] [LAURA]
                    txtAux(13).Text = vArtic.Familia 'para obtener centro de coste si necesario
                    ' ----
                    
                    'Fijar importes
                    FijarImportes False
                    
    
                    
    
                    
    
                    Me.Label1(8).Caption = txtAux(6).Text
                    
                    'recalculamos el importe total de la venta
                    Me.Label1(6).Caption = Format(ObtenerImporteTotal(False) + txtAux(5).Text, FormatoImporte)
                    
                   'Mostramos por el visor
                   EnviarVisorPuerto Me.Label1(7).Caption, Label1(8).Caption, "TOTAL", Label1(6).Caption
                    
                    
                   'si es de varios el precio se puede modificar y lo desbloqueamos
                   txtAux(4).Enabled = (vArtic.EsDeVarios = 1)
                   txtAux(4).Locked = Not (vArtic.EsDeVarios = 1)
                   If vArtic.EsDeVarios <> 0 Then
                        PonerFoco txtAux(4)
                   Else
                        If Not InsertoDesdeCestas Then
                            GuardarLinea
                        Else
                            nUnids = ""
                        End If
                        If nUnids <> "" Then
                            txtAux(3).Text = nUnids
                            PonerFormatoDecimal txtAux(3), 1
                            PonerArticuloCod codArt, False
                        End If
                   End If
                End If
            Else 'de leer
                MsgBox "No  se pudo leer el artículo", vbInformation
            End If
            
        Else
            'DE LEER EL ARTICULO. Ha dado error
            If txtAux(1).Text <> "" Then
                PonerFoco txtAux(1)
            Else
                txtAux2(2).Text = ""
            End If
        End If
        Set vArtic = Nothing
    Else
        txtAux2(2).Text = ""
    End If

EArticulo:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Poner datos del Artículo.", Err.Description
        Set vArtic = Nothing
    End If
End Sub




Private Sub LLamaLineas(alto As Single, xModo As Byte)
'Pone posicion TOP y LEFT de los controles en el form
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el alto
    For i = 0 To txtAux.Count - 1
        If ModificaLineas = 1 Then txtAux(i).Text = ""
        txtAux(i).Top = alto
        txtAux(i).Height = Me.DataGrid1.RowHeight - 10
    Next i
    If ModificaLineas = 1 Then txtAux2(2).Text = ""
    txtAux2(2).Top = alto
    txtAux2(2).Height = Me.DataGrid1.RowHeight - 10
    txtAux2(2).Enabled = False
    
    'El precio siempre bloqueado a no ser que sea articulo de varios
    txtAux(4).Locked = True
    txtAux(4).Enabled = False
    
    'El importe siempre bloqueado
    txtAux(5).Locked = True
    txtAux(5).Enabled = False
    
    
    ' ---- [21/10/2009] [LAURA] : se pone en el arreglatots
'    'Pongo el del DTO ajustadito al precio
'    txtAux(12).Left = txtAux(5).Left + txtAux(5).Width + 10
'    txtAux(12).Width = DataGrid1.Columns(13).Width - 20
End Sub


Private Sub BotonAnyadir()
    On Error GoTo EVenta
                
    ReiniciarVisor
    If InsertarVenta Then
        PonerVentasAbiertas
    
        'Cargar Data 1 con las ventas abiertas (cabeceras)
        cargaCabeceras True
        
        'Cargar las ventas que hay abiertas en el combo
        cargaComboVentas
    
        'Nos situamos en la ultima venta
        Me.cboNumVenta.ListIndex = Me.cboNumVenta.ListCount - 1
    
        'Ponemos el importe total de la venta
'        If Not Data1.Recordset.EOF Then
'            SQL = DevuelveDesdeBDNew(conAri, "scaven", "imptotal", "numtermi", CStr(NumTermi), "N", , "numventa", CStr(Data1.Recordset!numventa), "N", "fecventa", CStr(Data1.Recordset!fecventa), "F")
'            Me.Label1(6).Caption = Format(CCur(SQL), FormatoImporte)
'        End If
        PonerModo 3
        DoEvents
        If vParamTPV.Rapida Then
            'Entrada rapidad  DAVID
            PonerFoco txtAux(1) ' Nos situamos automáticamente en el campo de la linea
        Else
            PonerFoco Text1(0) ' (RAFA/ALZIRA 31082006) Nos situamos automáticamente en el campo de cliente
        End If
    End If
    
EVenta:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Venta.", Err.Description
    End If
End Sub




Private Sub BotonAnyadirLinea()
'Añade una nueva linea de Venta
Dim anc As Single

    On Error GoTo EAnyadirLinea
        
    ModificaLineas = 1
    AnyadirLinea DataGrid1, Data2
    FrameFito.visible = False
    
    If DataGrid1.Row < 0 Then
        anc = ObtenerAlto(DataGrid1, 50)
    Else
        anc = ObtenerAlto(DataGrid1, 20)
    End If
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(1)
    Exit Sub

EAnyadirLinea:
    MuestraError Err.Number, "Añadir linea.", Err.Description
End Sub




Private Function ObtenerImporteLin() As Currency
    'MODIFICADO.      Antes:  ObtenerImporteLin2(d1 As String, d2 As String
    'David  14  Mayo 08
    '
    'Antes pasabamos el descuento. Ahora debe estar en el campo txtaux 8 y 9
    If vParamTPV.Redondea2 Then
        'Redondea a 2 DECIMALES
         ObtenerImporteLin = CalcularImporte(txtAux(3).Text, txtAux(10).Text, txtAux(8).Text, txtAux(9).Text, vParamAplic.TipoDtos)
        
    Else
       'a 4
       'Este es el k hacen todos menos castelduc
       ObtenerImporteLin = CalcularImporte4(txtAux(3).Text, txtAux(10).Text, txtAux(8).Text, txtAux(9).Text, vParamAplic.TipoDtos)
    End If
End Function



Private Function FijarImportes(HaCambiadoElPrecio As Boolean)


            'Cambio
            'El parametro CalculaIVAsobrePVP indicara si
            ' el iva de la linea lo vamops a calcular sobre
            ' el pvp , es decir para que los decimales en pantalla se ajusten
            'primero calculamos el PVP. Luego fijamos el precioar
            

    If vParamTPV.CalculaIVAsobrePVP Then
        FijarImportesSobrePVP HaCambiadoElPrecio
    Else
        'Normal. Todos menos castelduc
        FijarImportes2 HaCambiadoElPrecio
        
    End If
        
End Function


Private Function FijarImportes2(HaCambiadoElPrecio As Boolean)
Dim IvaA2dec As Currency



            'Calculo el PVP unitario
            Sql = CalcularDto(txtAux(6).Text, CStr(txtAux(11).Text))
            
            txtAux(6).Text = ObtenerImporteLin
            
            
            If vParamTPV.Redondea2 Then
                PonerFormatoDecimal txtAux(6), 1
            Else
                PonerFormatoDecimal txtAux(6), 2   '4 digitos
            End If
            
            
            
            Sql = CalcularDto(txtAux(6).Text, CStr(txtAux(11).Text))
            
            
            If vParamTPV.Redondea2 Then
                'A dos decimales. El IVA lo sacara a partir del PVP
                IvaA2dec = CCur(ComprobarCero(Sql))
                IvaA2dec = Round2(IvaA2dec, 2)
                txtAux(5).Text = Round2(CCur(ComprobarCero(txtAux(6).Text)) + IvaA2dec, 2)
                PonerFormatoDecimal txtAux(5), 1
                
            Else
                'Como estaba
                'El IVA
                txtAux(5).Text = Round(CCur(ComprobarCero(txtAux(6).Text)) + CCur(ComprobarCero(Sql)), 4)
                PonerFormatoDecimal txtAux(5), 5
            End If
            
            If Not HaCambiadoElPrecio Then
            
                If vParamTPV.Redondea2 Then
                    'A dos decimales
                    txtAux(4).Text = Round2(CCur(txtAux(5).Text) / CCur(txtAux(3).Text), 2)
                    PonerFormatoDecimal txtAux(4), 1
                
                Else
                    'Como estaba
                    txtAux(4).Text = Round2(CCur(txtAux(5).Text) / CCur(txtAux(3).Text), 4)
                    PonerFormatoDecimal txtAux(4), 2
                End If
            End If
            
                
           
                
End Function

Private Sub PosicionarComboVentas(Venta As Long)

    For i = 0 To Me.cboNumVenta.ListCount - 1
        If CLng(Trim(Me.cboNumVenta.List(i))) = Venta Then
            Me.cboNumVenta.ListIndex = i
            Exit For
        End If
    Next i
    
End Sub


Private Function ObtenerImporteTotal(Optional actualiza As Boolean, Optional Error As Boolean) As Currency
'Suma el total de las lineas y actualiza la tabla scaven con el valor correcto
Dim RS As ADODB.Recordset
Dim total As Currency
        
    On Error GoTo ETotal
'    If Data2.Recordset.EOF Then Exit Function
    
'    SQL = Me.Data2.RecordSource
'    If SQL <> "" Then
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        total = 0
'        While Not RS.EOF
'            total = total + RS!ImporteL
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'    End If

    total = 0
    Sql = "select sum(importel) FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        total = DBLet(RS.Fields(0).Value, "D")
    End If
    
    RS.Close
    Set RS = Nothing
    
    ObtenerImporteTotal = total
        
    If actualiza Then
        Sql = "UPDATE " & NombreTabla & " SET imptotal=" & DBSet(total, "N")
        Sql = Sql & ObtenerWhereCP(True)
        conn.Execute Sql
    End If
    
ETotal:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Calculando Importe Total.", Err.Description
        Error = True
    Else
        Error = False
    End If
End Function


Private Function EstaFormularioAbierto() As Boolean
Dim Formulario As Form
    EstaFormularioAbierto = False
    For Each Formulario In Forms
        If (UCase(Formulario.Name) = UCase("frmFacTPVLotes")) Then
            EstaFormularioAbierto = True
            Exit For
        End If
    Next
    
    
End Function

Private Function InsertarLinea() As Boolean
'Inserta una linea de venta en la tabla: sliven
Dim B As Boolean
Dim ccoste As String
Dim LlevaArticuloLotes As Boolean
  
    If EstaFormularioAbierto Then Exit Function
    If Not DatosOkLinea(LlevaArticuloLotes) Then Exit Function
    
    On Error GoTo EInsLinea
    conn.BeginTrans
    
    'Anyadimos nuevo registro y rellenamos los campos clave (ocultos)
    'El campo de numventa no tiene valor es que aun se ha insertado la linea
    If Me.txtAux(0).Text = "" Then
        'numero de linea de la venta
        Sql = "numtermi=" & Data1.Recordset!NumTermi & " and numventa=" & Data1.Recordset!NumVenta & " and fecventa=" & DBSet(Data1.Recordset!fecventa, "F")
        Sql = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", Sql)
        txtAux(0).Text = Sql
        
        'El campo numlote lo guardaremos para saber si el articulo lleva lotes  o no, NO para indicar el nº d lote
        Sql = "INSERT INTO sliven(numtermi,numventa,fecventa,horventa,numlinea,codigoea,codartic,nomartic,cantidad,precioiv,importel,precioar,codigiva,dto1,dto2,implineareal,impartalb,codccost,numlote) "
        Sql = Sql & " VALUES (" & Data1.Recordset!NumTermi & "," & Data1.Recordset!NumVenta & "," & DBSet(Data1.Recordset!fecventa, "F") & ","
        '& DBSet(Data1.Recordset!fecventa & " " & Format(Now, "hh:mm:ss"), "FH") & ","
        Sql = Sql & DBSet(Now, "FH") & ","
        Sql = Sql & txtAux(0).Text & "," & DBSet(txtAux(1).Text, "T") & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T") & ","
        
        'DAVID.
        '   Antes pasaba el 6. Ahora paso el 10.  EL precio articulo esta en el 10
        '                  cantidad                        precioiv                             importe                             precioar
        Sql = Sql & DBSet(txtAux(3).Text, "N") & "," & DBSet(txtAux(4).Text, "N") & "," & DBSet(txtAux(5).Text, "N") & "," & DBSet(txtAux(10).Text, "N") & "," & txtAux(7).Text
        
        
        'Nuevo###       David
        'Para llevar so tiene descuentos en linea los guarda implineareal
        '                             dto1                              dto2                                 imp linea real
        Sql = Sql & "," & DBSet(txtAux(8).Text, "N", "N") & "," & DBSet(txtAux(9).Text, "N", "N") & "," & DBSet(txtAux(6).Text, "N", "N") & ","
        
        
        'Enero 2010
        'Importe para ajustar la linea dl importe del articulo. Por el tema de los decimales
        Sql = Sql & DBSet(txtAux(14).Text, "N", "N") & ","
        
        
        
        ' ---- [21/10/2009] [LAURA] : añadir centro de coste
        If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 1 Then 'por familia
            '- obtener centro de coste de la familia del articulo
            ccoste = DevuelveDesdeBDNew(conAri, "sfamia", "codccost", "codfamia", txtAux(13).Text, "N")
            Sql = Sql & DBSet(ccoste, "T", "S")
        Else
            Sql = Sql & ValorNulo
        End If
    
    
        If LlevaArticuloLotes Then
        
            If vParamAplic.ManipuladorFitosanitarios2 Then
                Sql = Sql & ",'*'"
            Else
                'ANtiguo. Para los que no se acojan a fitsonatiarios
                ccoste = DBSet(txtAux(2).Text, "T") & " ORDER BY fecentra desc"
                ccoste = DevuelveDesdeBD(conAri, "numlotes", "slotes", "codartic", ccoste, "N")
                Sql = Sql & "," & DBSet(ccoste, "T")
            End If
        Else
            Sql = Sql & ",NULL"
        End If
        Sql = Sql & ")"
        conn.Execute Sql
        
        
        'Desde tmpnlotes llevaremos los lotes de los registros fitosanitarios
        If LlevaArticuloLotes Then
            If vParamAplic.ManipuladorFitosanitarios2 Then
                Sql = "INSERT INTO slivenlotes(numtermi,numventa,fecventa,horventa,numlinea,sublinea,cantidad,numlote,fecentra,codartic)"
                Sql = Sql & " SELECT " & Data1.Recordset!NumTermi & "," & Data1.Recordset!NumVenta & "," & DBSet(Data1.Recordset!fecventa, "F")
                Sql = Sql & ",now()," & txtAux(0).Text
                Sql = Sql & " , numlinea , Cantidad, numlotes,fechaalb,codartic "
                Sql = Sql & " FROM tmpnlotes  WHERE codusu = " & vUsu.Codigo & " and cantidad <>0 "
    
                conn.Execute Sql
            End If
        End If
        
    End If
    
    Me.Label1(6).Caption = Format(ObtenerImporteTotal(True), FormatoImporte)
    B = True

EInsLinea:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Insertando linea.", Err.Description
    End If
    If B Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    InsertarLinea = B
End Function


Private Sub ReiniciarVisor()
    Me.Label1(6).Caption = "0,00"
    Me.Label1(7).Caption = "Precio"
    Me.Label1(8).Caption = "0,00"
    If vParamTPV.HayVisor Then
        EnviarVisorPuerto Label1(7).Caption, Label1(8).Caption, Label1(5).Caption, Label1(6).Caption
    End If
End Sub


Public Sub EnviarVisorPuerto(lin1 As String, imp1 As String, lin2 As String, imp2 As String)
Dim Buffer1 As Variant
Dim Buffer2 As Variant

    If Me.MSComm1.PortOpen Then
        If Len(imp1) >= 8 Then
            Buffer1 = Left(Mid(lin1, 1, 11) & Space(20), 11) & Right(Space(20) & Mid(imp1, 1, 9), 9)
        Else
            Buffer1 = Left(Mid(lin1, 1, 13) & Space(20), 13) & Right(Space(20) & Mid(imp1, 1, 7), 7)
        End If
        Buffer2 = Left(Mid(lin2, 1, 7) & Space(20), 7) & Right(Space(20) & Mid(imp2, 1, 13), 13)
        Me.MSComm1.Output = Buffer1 & Buffer2
    End If
End Sub



Private Function InsertarVenta() As Boolean
'Inserta en tabla temporal la cabecera de una nueva venta
Dim cont As Long
Dim B As Boolean
Dim MenError As String

    On Error GoTo EInsVenta
    
    conn.BeginTrans
    
    cont = vParamTPV.ConseguirContador(NumTermi)
    If Not IsNumeric(cont) Then
        MenError = "No se ha podido obtener nº de contador."
        B = False
    Else
        'Insertamos la cabecera de venta
        MenError = "Insertando en tabla de ventas."
        Sql = "INSERT INTO " & NombreTabla & " (numtermi,numventa,fecventa,horventa,codtraba,imcambio,imptotal,codclien) VALUES "
        Sql = Sql & "(" & NumTermi & "," & cont & "," & DBSet(Now, "F") & ", " & DBSet(Now, "FH") & "," & CodTraba & "," & ValorNulo & ",0,"
        'Poner el cliente que hay por defecto en los parametros
        If vParamTPV.Cliente = "" Then
            Sql = Sql & "0)"
        Else
            Sql = Sql & vParamTPV.Cliente & ")"
        End If
        
        conn.Execute Sql
        
        
        MenError = "Incrementando contador de venta."
        B = vParamTPV.IncrementarContador(NumTermi)
    End If
    
EInsVenta:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error insertando Venta." & vbCrLf & MenError, Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    InsertarVenta = B
End Function


Private Function ModificarVenta() As Boolean
'Modificar el cliente de una venta
Dim Sql As String
    On Error GoTo ErrModVen
    
    Sql = "UPDATE scaven SET "
    If Trim(Text1(0).Text) <> "" Then Sql = Sql & " codclien=" & DBSet(Text1(0).Text, "N") & ","
    Sql = Sql & "coddirec=" & DBSet(Text1(2).Text, "N", "S")

    'Si el cliente NO es de varios, fuerzo
    If Val(Text1(0).Text) <> vParamTPV.Cliente Then Sql = Sql & ",nifvarios=NULL"
        
    'Fuerzo un null
    Sql = Sql & ",ManipuladorNumCarnet=NULL,ManipuladorFecCaducidad =NULL ,ManipuladorNombre =  NULL , TipoCarnet = NULL"
    
    
    Sql = Sql & ", observa1=" & DBSet(Text1(1).Text, "T")
            
    Sql = Sql & " WHERE numtermi=" & Me.Data1.Recordset!NumTermi & " AND numventa=" & Me.Data1.Recordset!NumVenta
    Sql = Sql & " AND fecventa=" & DBSet(Me.Data1.Recordset!fecventa, "F")
    conn.Execute Sql
    
    Data1.Refresh
    PosicionarData
    
    Exit Function
    
ErrModVen:
    ModificarVenta = False
    MuestraError Err.Number, "Modificar cabecera de la venta.", Err.Description
End Function


Private Sub txtAux2_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux2(Index)
End Sub



Private Sub GuardarLinea()
Dim PrimeraLin As Boolean
    
    
    
    
    If Modo = 4 Then
        'ESTOY MODIFICANDO
        'No hago nada
        
    Else
        
        If Data2.Recordset.EOF = True Then PrimeraLin = True
         If InsertarLinea Then
            DataGrid1.AllowAddNew = False
            FrameFito.visible = True
    '        PonerModo 2
            
            If PrimeraLin Then
                CargaGrid DataGrid1, Data2, True
            Else
                CargaGrid2 DataGrid1, Data2
            End If
            
            If ModificaLineas = 1 Then 'Insertar
                BotonAnyadirLinea
            Else
                PonerModo 2
            End If
    
            
        Else
            txtAux(0).Text = "" 'Limpiamos el num linea para volver a insertar tras corregir datos
        End If
    End If
End Sub



Private Function AbrirVisorPuerto() As Boolean
On Error GoTo EAbrirVisor

    'Establecemos el puerto de comunicaciones
    If vParamTPV.HayVisor Then
        Me.MSComm1.CommPort = vParamTPV.NumPuerto
        
        ' 9600 baudios, sin paridad, 8 bits de datos y 1 bit de parada.
        Me.MSComm1.Settings = vParamTPV.VelociPuerto & ",N,8,1"
        
        ' Indicar al control que lea todo el búfer al usar Input.
        MSComm1.InputLen = 0

        'Abrimos el puerto
        Me.MSComm1.PortOpen = True
                
        Me.MSComm1.Output = Mid(vEmpresa.nomempre & Space(40), 1, 40)
    End If
    
EAbrirVisor:
    If Err.Number <> 0 Then MuestraError Err.Number, "Abrir visor.", Err.Description
End Function




Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(0).Text
    
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(2).Text, NomDpto) Then
        Text2(2).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function



Private Function ClienteOK(newCli As String, oldCli As String, modifica As Boolean, Optional mostrarObs As Boolean) As Boolean
'(IN) newCli: cliente nuevo q queremos poner
'(IN) oldCli: cliente guardado actualmente si existe
Dim cCli As CCliente
Dim devuelve As String

    On Error GoTo ErrCliOK
    ClienteOK = False
    
    If newCli <> "" Then newCli = CStr(Val(newCli))
    Set cCli = New CCliente
    If cCli.LeerDatos(newCli) Then
        '-- Si el cliente esta bloqueado no permitimos este cliente para la venta
        If cCli.ClienteBloqueado(2, False) Then
            Set cCli = Nothing
            Exit Function
        End If
        
        If vParamAplic.NumeroInstalacion <> vbTaxco Then
            If cCli.NIF = vParam.CifEmpresa Then
                MsgBox "Debe realizar albaranes internos.", vbExclamation
                Set cCli = Nothing
                Exit Function
            End If
        End If
        
        '-- si se ha modificado el cliente y si hay lineas de articulos:
        '   comprobar q el nuevo cliente tiene la misma tarifa q el cliente anterior
        '   sino no permitimos el nuevo cliente para la venta
        If (oldCli <> "") And (newCli <> oldCli) Then
            If Not Me.Data2.Recordset.EOF Then 'si hay lineas
                'obtener la tarifa del cliente actual
                devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", oldCli, "N")
                If devuelve <> CStr(cCli.Tarifa) Then
                    devuelve = "No se puede seleccionar el cliente " & newCli & " "
                    devuelve = devuelve & "ya que tiene distinta tarifa de precios." & vbCrLf
                    devuelve = devuelve & "Seleccione un cliente de la misma tarifa o elimine la venta."
                    MsgBox devuelve, vbExclamation, "Comprobar cliente"
                    Set cCli = Nothing
                    Exit Function
                Else
                    modifica = True
                End If
            Else
                modifica = True
            End If
        End If
        ClienteOK = True
        
        'mostrar las observaciones del cliente
        If mostrarObs Then cCli.MostrarObservaciones
    End If
    
    Set cCli = Nothing
    Exit Function
    
ErrCliOK:
    MuestraError Err.Number, "Comprobar cliente correcto.", Err.Description
End Function



Private Sub RevisarEntradasDia()

    CadSelVenta = ""
    Set frmTraerVen = New frmFacTPVTraerVen
    frmTraerVen.parNumTermi = -1   'El menos 1 siginifac que me mostrara las ventas del dia
    frmTraerVen.Show vbModal
    Set frmTraerVen = Nothing
    If CadSelVenta <> "" Then
        
        i = RecuperaValor(CadSelVenta, 1)

        If i = 1 Then
        
            Sql = RecuperaValor(CadSelVenta, 2)
            CadSelVenta = RecuperaValor(CadSelVenta, 3)
        
            'IMPRESION
            If Sql = "FTI" Then
                 ImprimirTicketDirecto CadSelVenta, Now
            Else
                'Es una factura de venta
                ImprimirFAV
            End If
            
        Else
            Sql = RecuperaValor(CadSelVenta, 5)   'Tipo movimiento albaran
            'Ponemos el numero de albaran, NO el numero de factura
            CadSelVenta = RecuperaValor(CadSelVenta, 4)
            'Ver lineas detalles
            '----------------------------------
            With frmFacHcoFacturas2
                .DesdeFichaCliente = False
                 .hcoCodMovim = CadSelVenta
                 .hcoCodTipoM = Sql
                 .hcoFechaMov = Now
                 .Show vbModal
            End With
            
            
            
            
            
        End If
    End If
    CadSelVenta = ""
    Screen.MousePointer = vbDefault
End Sub


Private Sub ImprimirFAV()
Dim cadParam As String
Dim numParam As Byte
Dim nomDocu As String

'    '===================================================
'    '============ PARAMETROS ===========================
'    If FormatoFacturaTPV Then
'        indRPT = 18 'FACTURAS TPV
'    Else
'        indRPT = 12 'Facturas Clientes
'    End If
    If Not PonerParamRPT2(18, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        Exit Sub
    End If

    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    frmImprimir.NombrePDF = pPdfRpt
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme

    'EN sql tengo el numero de factura
    Sql = "({scafac.numfactu} = " & CadSelVenta & ")"
    
    'Tipo documento
    Sql = Sql & " AND ({scafac.codtipom}='FAV') "
    
    'fecha factu
    Sql = Sql & " AND ({scafac.fecfactu} =  Date(" & Year(Now) & "," & Month(Now) & "," & Day(Now) & ")" & ")"

    CadSelVenta = Sql
    'If Not HayRegParaInforme("scafac", CadSelVenta) Then Exit Sub

    With frmImprimir
            .FormulaSeleccion = Sql
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 53
            .Titulo = ""
            .Show vbModal
    End With
End Sub



Private Function FijarImportesSobrePVP(HaCambiadoElPrecio As Boolean)
Dim IvaA2dec As Currency
Dim V2 As Currency


            If txtAux(10).Text = "" Then txtAux(10).Text = "0"
            If txtAux(14).Text = "" Then txtAux(14).Text = "0"
            'Lo primero que hago es obtener el PVP de una unidad
            Sql = CalcularDto(txtAux(10).Text, CStr(txtAux(11).Text))
            IvaA2dec = CCur(ComprobarCero(Sql))
            
            IvaA2dec = CCur(txtAux(10).Text) + IvaA2dec
            
            'Columna precio
            txtAux(14).Text = IvaA2dec 'me lo guardo por problemas de decimales
            If vParamTPV.Redondea2 Then
                'A dos decimales.
                IvaA2dec = Round2(IvaA2dec, 2)
                txtAux(4).Text = IvaA2dec
                PonerFormatoDecimal txtAux(4), 1
            Else
                PonerFormatoDecimal txtAux(4), 2
                txtAux(4).Text = IvaA2dec
                IvaA2dec = Round2(IvaA2dec, 3)
            End If
            IvaA2dec = ObtenerImporteLinPVP(CStr(IvaA2dec), vParamTPV.Redondea2)

            txtAux(5).Text = IvaA2dec
        
            
            PonerFormatoDecimal txtAux(5), 1   'importe a dos decimales siempre

            

          
            
            IvaA2dec = CCur(txtAux(5).Text)   'Precio con IVA
            'Calcualremos el precio sin IVA
            If txtAux(11).Text = "" Then txtAux(11).Text = "0"
            V2 = (CCur(txtAux(11).Text) / 100) + 1
            V2 = (IvaA2dec / V2)
            
            
            
            
            
    
            'nos falta fijar el precio linea
           IvaA2dec = Round2(V2 / CCur(txtAux(3).Text), 4)

       
            

            
            'If V2 <> IvaA2dec Then txtAux(10).Text = V2
            'Este lo obviamos y cuando pasamos a albaran cojeremos el del 14
            txtAux(14).Text = IvaA2dec
     
            'Piongo el importe linea correcto. Lo fuerzo a 4 decimales para ajustar mejor
            V2 = ObtenerImporteLinPVP(CStr(IvaA2dec), False)
            txtAux(6).Text = V2
           
                
End Function



Private Function ObtenerImporteLinPVP(TextoImporteConIva As String, A2Decimales As Boolean) As Currency
    'MODIFICADO.      Antes:  ObtenerImporteLin2(d1 As String, d2 As String
    'David  18 Enero 2010
    '
    'Calcula el precio la linea, a partir del PVP
    If A2Decimales Then
        'Redondea a 2 DECIMALES
         ObtenerImporteLinPVP = CalcularImporte(txtAux(3).Text, TextoImporteConIva, txtAux(8).Text, txtAux(9).Text, vParamAplic.TipoDtos)
        
    Else
       'a 4
       'Este es el k hacen todos menos castelduc
       ObtenerImporteLinPVP = CalcularImporte4(txtAux(3).Text, TextoImporteConIva, txtAux(8).Text, txtAux(9).Text, vParamAplic.TipoDtos)
    End If
End Function



Private Sub BotonModifcarLinea()
Dim anc As Integer
Dim T As cTag
Dim EnBD As Boolean
Dim Cad As String
 
    If Me.Data1.Recordset.EOF Then Exit Sub
    If Me.Data2.Recordset.EOF Then Exit Sub

    If Modo = 3 Then
        'Estoy insertando. Deberia seleccionar la linea a modificar
        MsgBox "Seleccione la linea a modificar", vbExclamation
        Exit Sub
    End If
    
    
    If DataGrid1.Row < 0 Then
        anc = ObtenerAlto(DataGrid1, 50)
    Else
        anc = ObtenerAlto(DataGrid1, 20)
    End If
    ModificaLineas = 2
    LLamaLineas CSng(anc), 4
    'Relleno los campos
    
    'numtermi, numventa,fecventa,horventa, numlinea, codigoea,codartic, nomartic, cantidad, precioiv, importel,precioar,codigiva "
    'NUEVO
    'SQL = SQL & " ,if((dto1+dto2)>0,""*"","""") as dto"
    
    Set T = New cTag
    For anc = 0 To txtAux.Count - 1
        EnBD = False
        If T.Cargar(txtAux(anc)) Then
            If T.columna <> "" Then EnBD = True
                
        Else
            'S top
        End If
        If T.columna = "dto2" Then EnBD = False
        If EnBD Then
            Cad = DBLet(Data2.Recordset.Fields(T.columna))
            If T.Formato <> "" Then Cad = Format(Cad, T.Formato)
            txtAux(anc).Text = Cad
        Else
            txtAux(anc).Text = ""
        End If
        
    Next anc
    txtAux2(2).Text = Data2.Recordset!NomArtic
    Set T = Nothing
    
    Me.cmdAceptar.visible = True
    Me.cmdCancelar.visible = True
    PonerArticuloCod txtAux(2), False
    PonerFoco txtAux(3)
End Sub



    


Private Function HacerUpdate() As Boolean

        On Error GoTo EHacerUpdate
        HacerUpdate = False
        
    
       'codigoea,codartic,nomartic,cantidad,precioiv,
       'importel,precioar,codigiva,dto1,dto2,implineareal,impartalb,codccost
       '  añadir centro de coste
        If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 1 Then 'por familia
            '- obtener centro de coste de la familia del articulo
            Sql = DevuelveDesdeBDNew(conAri, "sfamia", "codccost", "codfamia", txtAux(13).Text, "N")
            Sql = DBSet(Sql, "T", "S")
        Else
            Sql = ValorNulo
        End If
        
       Sql = "UPDATE sliven set codccost = " & Sql
       Sql = Sql & ", codigoea = " & DBSet(txtAux(1).Text, "T")
       Sql = Sql & ", codartic = " & DBSet(txtAux(2).Text, "T")
       Sql = Sql & ", nomartic = " & DBSet(txtAux2(2).Text, "T")
       
        
        'DAVID.
        '   Antes pasaba el 6. Ahora paso el 10.  EL precio articulo esta en el 10
       Sql = Sql & ", cantidad = " & DBSet(txtAux(3).Text, "N")
       Sql = Sql & ", precioiv = " & DBSet(txtAux(4).Text, "N")
       Sql = Sql & ", importel = " & DBSet(txtAux(5).Text, "N")
       Sql = Sql & ", precioar = " & DBSet(txtAux(10).Text, "N")
       Sql = Sql & ", codigiva = " & txtAux(7).Text
        '
        '                             dto1                              dto2
        Sql = Sql & ", dto1 = " & DBSet(txtAux(8).Text, "N", "N")
        Sql = Sql & ", dto2 = " & DBSet(txtAux(9).Text, "N", "N")
        'imp linea real
        Sql = Sql & ", implineareal = " & DBSet(txtAux(6).Text, "N", "N")
        
        
        'Enero 2010
        'Importe para ajustar la linea dl importe del articulo. Por el tema de los decimales
        
        Sql = Sql & ", impartalb = " & DBSet(txtAux(14).Text, "N", "N")
        
        
        
        
        
        Sql = Sql & " WHERE numtermi=" & Data1.Recordset!NumTermi & " and numventa=" & Data1.Recordset!NumVenta & " and fecventa=" & DBSet(Data1.Recordset!fecventa, "F")
        Sql = Sql & " AND numlinea = " & Data2.Recordset!numlinea
        conn.Execute Sql
        HacerUpdate = True
        Exit Function
EHacerUpdate:
        MuestraError Err.Number, "Updatenado lineas"
       
End Function

Private Sub CargaLotesArticulos()
Dim C As String
    
    txtFito2(0).Text = ""
    If Not vParamAplic.ManipuladorFitosanitarios2 Then
        txtFito2(0).Text = Trim(DBLet(Data2.Recordset!numLote, "T"))
        If txtFito2(0).Text <> "" Then
            txtFito2(0).Text = "LOTE:  " & txtFito2(0).Text
        
            C = DevuelveDesdeBD(conAri, "concat(numserie,""|"",if(fecvigen is null,'',fecvigen),""|"")", "sartic", "codartic", CStr(Data2.Recordset!codArtic), "T")
            If C <> "" Then
             
                txtFito2(0).Text = txtFito2(0).Text & "    Nº Serie: " & RecuperaValor(C, 1)
                txtFito2(0).Text = txtFito2(0).Text & "   Fecha vigencia: " & Format(RecuperaValor(C, 2), "dd/mm/yyyy")
            End If
        
        End If
        Exit Sub
    End If
    
    'Solo manipuladores fitosanitarios
    If DBLet(Data2.Recordset!numLote, "T") = "" Then Exit Sub
    
    


    
    Set miRsAux = New ADODB.Recordset
    C = Mid(Data2.RecordSource, InStr(1, Data2.RecordSource, " WHERE "))
    C = Mid(C, 1, InStr(1, C, " ORDER BY")) & " AND numlinea =" & Data2.Recordset!numlinea
    C = "Select numlote,cantidad,fecentra from slivenlotes  " & C & "  order by sublinea"
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    C = ""
    While Not miRsAux.EOF

        If NumRegElim = 0 Then
            C = C & "LOTE:  "
        Else
            C = C & "        "
        End If
        C = C & Mid(miRsAux!numLote & Space(20), 1, 20) & "   "
        If NumRegElim = 0 Then
            C = C & "F.ENTR: "
        Else
            C = C & Space(8)
        End If
        C = C & Format(miRsAux!fecentra, "dd/mm/yyyy") & "   "
        If NumRegElim = 0 Then
            C = C & "CANTIDAD: "
        Else
            C = C & Space(10)
        End If
        C = C & Format(miRsAux!cantidad, FormatoCantidad)
        
        miRsAux.MoveNext
        If NumRegElim = 0 Then
            If Not miRsAux.EOF Then C = C & "   +++"
        End If
        
        
        C = C & vbCrLf
        NumRegElim = NumRegElim + 1
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    txtFito2(0).Text = C

    
    

End Sub



Private Sub ModificaLote()
Dim CadenaInsertTmpLotes As String
Dim J As Integer
Dim LotesArticulos
Dim IncioBusqueda As Integer
Dim fin As Boolean
        Set miRsAux = New ADODB.Recordset
          
        CadenaInsertTmpLotes = Mid(Data2.RecordSource, InStr(1, Data2.RecordSource, " WHERE "))
        CadenaInsertTmpLotes = Mid(CadenaInsertTmpLotes, 1, InStr(1, CadenaInsertTmpLotes, " ORDER BY")) & " AND numlinea =" & Data2.Recordset!numlinea
        CadenaInsertTmpLotes = "Select numlote,cantidad,fecentra from slivenlotes  " & CadenaInsertTmpLotes & "  order by sublinea"
        miRsAux.Open CadenaInsertTmpLotes, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        LotesArticulos = "|"
        While Not miRsAux.EOF
            LotesArticulos = LotesArticulos & miRsAux!numLote & "#@#" & Format(miRsAux!fecentra, "dd/mm/yyyy") & Mid(miRsAux!cantidad & Space(10), 1, 10) & "|"
          
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
            CadenaInsertTmpLotes = ""
            Sql = "select codartic,numlotes,fecentra,canentra-vendida disponible from slotes where "
            Sql = Sql & " codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " and canentra-vendida >0 order by fecentra "
            
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                'insert into tmpnlotes(codusu,numlinea,fechaalb,codprove,cantidad,numlotes)
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & ", (" & vUsu.Codigo & "," & DBSet(miRsAux!codArtic, "T") & "," & NumRegElim
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!fecentra, "F")
                'CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(txtAux(2).Text, "T") & "," & DBSet(txtAux2(2).Text, "T")
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!disponible * 100, "N") & ","
                                
                Sql = "|" & miRsAux!numlotes & "#@#"
                fin = False
                IncioBusqueda = 1
                
                While Not fin
                    
                     
                    J = InStr(IncioBusqueda, LotesArticulos, Sql)
                    If J > 0 Then
                        J = J + Len(Sql)
                        Sql = Mid(LotesArticulos, J, 10)
                        
                        If Sql = Format(miRsAux!fecentra, "dd/mm/yyyy") Then
                            'OK, esta es la linea
                            Sql = Trim(Mid(LotesArticulos, J + 10, 10))
                            fin = True
                        Else
                            Sql = "|" & miRsAux!numlotes & "#@#"   'Vuelve a la busqueda
                            IncioBusqueda = InStr(J, LotesArticulos, "|")
                        End If
                    Else
                        Sql = "0"
                        fin = True
                    End If
                Wend
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & Sql
                
                CadenaInsertTmpLotes = CadenaInsertTmpLotes & "," & DBSet(miRsAux!numlotes, "T") & ")"
                
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
                
                
            'Si hay mas de uno mostraremos cual y cuanto puede coger
            If NumRegElim = 0 Then
                MsgBox "Ningun lote disponible para el artículo", vbExclamation
              
            Else
              
                'Mas de un  lote disponible
                Screen.MousePointer = vbHourglass
                
                conn.Execute "DELETE FROM tmpnlotes where codusu =" & vUsu.Codigo
                Espera 0.3
                CadenaInsertTmpLotes = Mid(CadenaInsertTmpLotes, 2)
                CadenaInsertTmpLotes = "insert into tmpnlotes(codusu,codartic,numlinea,fechaalb,codprove,cantidad,numlotes) VALUES " & CadenaInsertTmpLotes
                conn.Execute CadenaInsertTmpLotes
                
                
                
              
                    CadenaDesdeOtroForm = ""
                    frmFacTPVLotes.TotalLineas = Data2.Recordset!cantidad
                    frmFacTPVLotes.NombreArticulo = Data2.Recordset!NomArtic
                    frmFacTPVLotes.Show vbModal
              
                    If CadenaDesdeOtroForm = "OK" Then
                        Sql = ObtenerWhereCP(True)
                        Sql = Replace(Sql, "scaven", "slivenlotes")
                        Sql = "DELETE FROM slivenlotes " & Sql
                        Sql = Sql & " AND numlinea = " & Data2.Recordset!numlinea
                        conn.Execute Sql
                        Espera 0.2
                        Sql = "INSERT INTO slivenlotes(numtermi,numventa,fecventa,horventa,numlinea,sublinea,cantidad,numlote,fecentra,codartic)"
                        Sql = Sql & " SELECT " & Data1.Recordset!NumTermi & "," & Data1.Recordset!NumVenta & "," & DBSet(Data1.Recordset!fecventa, "F")
                        Sql = Sql & ",now()," & Data2.Recordset!numlinea
                        Sql = Sql & " , numlinea , Cantidad, numlotes,fechaalb,codartic FROM tmpnlotes  WHERE codusu = " & vUsu.Codigo & " and cantidad >0 "
            
                        conn.Execute Sql
                        Espera 0.3
                        CargaLotesArticulos
                        
                    End If
              
            End If

End Sub



Private Function ControlesFitosantiarios(ByRef TieneArticulosFito As Boolean) As Boolean
Dim Aux As String
Dim CADENA As String
Dim MenError As String
Dim linea As Integer

  
  
        ControlesFitosantiarios = False
  
        'Veremos si hay algun dato en la venta que lleva fitosanitarios
        Set miRsAux = New ADODB.Recordset
        
        MenError = ""
        
        Sql = ObtenerWhereCP(True)
        Sql = Replace(Sql, "scaven", "sliven")
        Sql = Sql & " and not isnull(numserie) and trim(numserie)<>''"
        Sql = " inner join sartic on sliven.codartic=sartic.codartic" & Sql
        Sql = "SELECT numlinea,sliven.codartic,sum(cantidad) lasuma FROM sliven" & Sql & " GROUP BY 1"
            
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Aux = ""
        linea = 0
        
        
        While Not miRsAux.EOF
            If Not IsNull(miRsAux!numlinea) Then
                Sql = Format(miRsAux!numlinea, "0000000") & Mid(miRsAux!codArtic & Space(16), 1, 16) & DBLet(miRsAux!lasuma, "N") & "|"
                Aux = Aux & Sql
                TieneArticulosFito = True
            End If
            miRsAux.MoveNext

        Wend
        miRsAux.Close
      
        While Aux <> ""
            i = InStr(1, Aux, "|")
            'Tiene que ser >0
            CADENA = Mid(Aux, 1, i - 1)
            Aux = Mid(Aux, i + 1)
                
        
            'codartic,cantidad
            Sql = ObtenerWhereCP(True)
            Sql = Replace(Sql, "scaven", "slivenlotes")
            Sql = "select numlinea,sum(cantidad) lasuma FROM slivenlotes " & Sql
            Sql = Sql & " AND numlinea = " & Mid(CADENA, 1, 7) & " GROUP BY 1"
            
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then

                If CCur(Trim(Mid(CADENA, 24))) <> miRsAux!lasuma Then
                    Sql = "Linea: " & Val(Mid(CADENA, 1, 7)) & "   " & Mid(CADENA, 8, 16) & "@#@" & "  - Canti. <> asiganda"
                    CADENA = Trim(Mid(CADENA, 8, 16))
                    CADENA = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", CADENA, "T")
                    Sql = Replace(Sql, "@#@", CADENA)
                    MenError = Sql & vbCrLf & MenError
                End If
            Else
                Sql = "Linea: " & Val(Mid(CADENA, 1, 7)) & "   " & Mid(CADENA, 8, 16) & "@#@" & "  -Sin asignar"
                CADENA = Trim(Mid(CADENA, 8, 16))
                CADENA = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", CADENA, "T")
                Sql = Replace(Sql, "@#@", CADENA)
                MenError = MenError & vbCrLf & Sql
            End If
            
            miRsAux.Close
            
            
            
            
        Wend
        Set miRsAux = Nothing
        
        If MenError <> "" Then
            MenError = "Error Lotes fitosantiarios" & vbCrLf & String(45, "=") & vbCrLf & MenError
            MsgBox MenError, vbExclamation
            Exit Function
        Else
            ControlesFitosantiarios = True
        End If
        
        

        
        
        
   

End Function




Private Sub InsertaLineasArticulosAgrupados()
Dim Uds As Integer
Dim Aux As String

    On Error GoTo eInsertaCajasNavidad

    Set miRsAux = New ADODB.Recordset
        
    Sql = "numtermi=" & Data1.Recordset!NumTermi & " and numventa=" & Data1.Recordset!NumVenta & " and fecventa=" & DBSet(Data1.Recordset!fecventa, "F")
    Sql = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", Sql)
    Sql = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", Sql)
    NumRegElim = Val(Sql)
    
    Sql = "select * from sarticAgrupado  where idcaja=" & RecuperaValor(CadenaDesdeOtroForm, 1)
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then Err.Raise 513, , "Error leyendo caja: " & RecuperaValor(CadenaDesdeOtroForm, 1)
    Uds = CInt(RecuperaValor(CadenaDesdeOtroForm, 2))
    
    
    '(numtermi,numventa,fecventa,horventa,numlinea,codigoea,codartic,nomartic,cantidad,precioiv,importel,precioar,codigiva,dto1,dto2,implineareal,impartalb,codccost,numlote) "
    Sql = ""
    CargaLineArticuloLotes 0, Uds
    If Sql <> "" Then Aux = Aux & Sql
        
    Sql = ""
    CargaLineArticuloLotes 1, Uds
    If Sql <> "" Then Aux = Aux & Sql
    
    Sql = ""
    CargaLineArticuloLotes 2, Uds
    If Sql <> "" Then Aux = Aux & Sql
        
    Aux = Mid(Aux, 2)
    Sql = "INSERT INTO sliven(numtermi,numventa,fecventa,horventa,numlinea,codigoea,codartic,nomartic,cantidad,precioiv,importel,precioar,codigiva,dto1,dto2,implineareal,impartalb,codccost,numlote) VALUES  " & Aux
    conn.Execute Sql
    
     Set miRsAux = Nothing
   'Cargamos grid
   CargaGrid2 DataGrid1, Data2
   'Me pongo a insertar otra
   BotonAnyadirLinea
Exit Sub
eInsertaCajasNavidad:
    
    MuestraError Err.Number
    Set miRsAux = Nothing
    
End Sub

'mirsaux esta cargado
Private Sub CargaLineArticuloLotes(QueArticulo As Byte, Cajas As Integer)
Dim Articulo As String
Dim Importe As Currency
Dim Codigiva As String
Dim Porciva As Currency
Dim Precio As Currency

        ' 0: 21%     1: 10%      2:4%
        Sql = ""
        Articulo = ""
        If QueArticulo = 0 Then
            Articulo = DBLet(miRsAux!artIVA_Normal, "T")
            If Articulo <> "" Then Importe = miRsAux!baseIVA_Normal
        ElseIf QueArticulo = 1 Then
            Articulo = DBLet(miRsAux!artIVA_Reducido, "T")
            If Articulo <> "" Then Importe = miRsAux!baseIVA_Reducido
        ElseIf QueArticulo = 2 Then
            Articulo = DBLet(miRsAux!artIVA_SuperReducido, "T")
            If Articulo <> "" Then Importe = miRsAux!baseIVA_superReducido
        End If
        
        If Articulo = "" Then Exit Sub
        
        
        
        
        Codigiva = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", Articulo, "T")
        
        Porciva = CCur(DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", Codigiva))
        
        Sql = ", (" & Data1.Recordset!NumTermi & "," & Data1.Recordset!NumVenta & "," & DBSet(Data1.Recordset!fecventa, "F") & ","
        Sql = Sql & DBSet(Now, "FH") & "," & NumRegElim
        'codigoea,codartic,nomartic,cantidad,
        Sql = Sql & ",null," & DBSet(Articulo, "T") & "," & DBSet(miRsAux!texto & " Ref(" & miRsAux!Referencia & ")", "T") & "," & DBSet(Cajas, "T") & ","
        
        'precioiv,importel,precioar,codigiva,"
        
        Precio = (Porciva * Importe) / 100
        'If Cajas = 1 Then Precio = Round(Precio, 2)
        
        
        Sql = Sql & DBSet((Precio + Importe), "N") & ","
        Sql = Sql & DBSet((Precio + Importe) * Cajas, "N") & "," & DBSet(Importe, "N") & "," & DBSet(Codigiva, "N")
        
        'dto1 dto2        implineareal,impartalb
        
        Sql = Sql & ",0,0," & DBSet(Importe * Cajas, "N") & "," & DBSet(Precio, "N") & ","
        






'        If vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 1 Then 'por familia
'            '- obtener centro de coste de la familia del articulo
'            ccoste = DevuelveDesdeBDNew(conAri, "sfamia", "codccost", "codfamia", txtAux(13).Text, "N")
'            SQL = SQL & DBSet(ccoste, "T", "S")
'        Else
            Sql = Sql & ValorNulo
        'End If
    
        Sql = Sql & ",NULL)"
        NumRegElim = NumRegElim + 1
End Sub



Private Sub EstableceCestas()
On Error GoTo eEstableceCestas
    TieneAPP_Cestas = False
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from cestas where false", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TieneAPP_Cestas = True
    miRsAux.Close
    
eEstableceCestas:
    If Err.Number <> 0 Then
        Err.Clear
        conn.Errors.Clear
    End If

    Set miRsAux = Nothing
    

End Sub






Private Function InsertaLineasArticulosCesta() As Boolean
Dim i As Integer
Dim Aux As String
Dim RS As ADODB.Recordset

    On Error GoTo eInsertaCajasNavidad
    InsertaLineasArticulosCesta = False
    Set RS = New ADODB.Recordset
        
    
    Sql = " select cestas.*,cestas_lineas.*, nomartic,codigiva from cestas inner join cestas_lineas on cestas.cestaId= cestas_lineas.cestaId "
    Sql = Sql & " inner join sartic on cestas_lineas.codartic=sartic.codartic where cestas.codclien=" & Text1(0).Text & " order by numlinea"
       
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then Err.Raise 513, , "Error leyendo cesta cliente " & Text1(0).Text
    
    While Not RS.EOF
    
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next
        txtAux2(2).Text = ""

        
        txtAux(3).Text = RS!cantidad
        If Not PonerFormatoDecimal(txtAux(3), 1) Then Err.Raise 513, , "Poner cantidad"
        
        txtAux(2).Text = RS!codArtic
        PonerArticuloCod (txtAux(2).Text), True
        
        If txtAux(2).Text = "" Then Err.Raise 513, , "Articulo  bloqueado" & miRsAux!codArtic
            
        If Not InsertarLinea Then Err.Raise 513, , "Error insertando liena"
           
        
        
        
        
       
    
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
   
   'Borramos la cestas
   
   Sql = " delete from cestas_lineas where cestaId IN (Select cestaid from cestas where codclien=" & Text1(0).Text & ")"
    conn.Execute Sql
   Sql = " delete from cestas where codclien=" & Text1(0).Text
   conn.Execute Sql
   
   
   InsertaLineasArticulosCesta = True
   
   
Exit Function
eInsertaCajasNavidad:
    
    MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    
End Function

