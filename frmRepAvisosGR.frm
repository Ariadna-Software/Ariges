VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepAvisos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avisos de clientes"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   13395
   Icon            =   "frmRepAvisosGR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   48
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   49
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3825
      TabIndex        =   46
      Top             =   90
      Width           =   2010
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   47
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cambiar visitado"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cerrar Aviso"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Entrada Equipo"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5940
      TabIndex        =   44
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   45
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
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
      Left            =   11700
      TabIndex        =   43
      Top             =   225
      Width           =   1575
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
      Left            =   12195
      TabIndex        =   17
      Top             =   8010
      Width           =   1065
   End
   Begin VB.Frame FrameAveria 
      Caption         =   " Datos Averia "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6210
      Left            =   6720
      TabIndex        =   34
      Top             =   1710
      Width           =   6525
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
         Index           =   13
         Left            =   2820
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   960
         Width           =   3495
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
         Index           =   13
         Left            =   1995
         MaxLength       =   30
         TabIndex        =   14
         Tag             =   "T�cnico|N|S|0|9999|scaavi|codtecni|0000|N|"
         Text            =   "Text1"
         Top             =   960
         Width           =   810
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
         Height          =   4170
         Index           =   3
         Left            =   120
         MaxLength       =   800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Tag             =   "Observaciones|T|S|||scaavi|observac||N|"
         Top             =   1920
         Width           =   6195
      End
      Begin VB.ComboBox cboSituacion 
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
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Situaci�n|N|N|||scaavi|situacio||N|"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1710
         ToolTipText     =   "Buscar trabajador"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Para el t�cnico"
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
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de averia detectada"
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
         Index           =   45
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Situaci�n"
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
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame FrameCliente 
      Caption         =   " Datos Cliente "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6225
      Left            =   120
      TabIndex        =   25
      Top             =   1710
      Width           =   6495
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
         Height          =   1515
         Index           =   14
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   12
         Tag             =   "ObsIn.|T|S|||scaavi|observacrm|||"
         Text            =   "frmRepAvisosGR.frx":000C
         Top             =   4560
         Width           =   6000
      End
      Begin VB.CheckBox chkVisitado 
         Alignment       =   1  'Right Justify
         Caption         =   "VISITADO"
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
         Left            =   4950
         TabIndex        =   40
         Tag             =   "V|N|N|||scaavi|visitado|||"
         Top             =   225
         Width           =   1335
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
         Index           =   5
         Left            =   1125
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Nombre Cliente|T|N|||scaavi|nomclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   645
         Width           =   5115
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
         Left            =   240
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Cod. Cliente|N|N|0|999999|scaavi|codclien|000000|N|"
         Text            =   "Text1"
         Top             =   645
         Width           =   810
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
         Left            =   240
         MaxLength       =   35
         TabIndex        =   8
         Tag             =   "Domicilio|T|N|||scaavi|domclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   2580
         Width           =   6000
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
         Left            =   225
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "NIF Cliente|T|N|||scaavi|nifclien||N|"
         Text            =   "123456789"
         Top             =   1935
         Width           =   1770
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
         Left            =   4230
         MaxLength       =   20
         TabIndex        =   7
         Tag             =   "tel�fono Cliente|T|S|||scaavi|telclien||N|"
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
         Left            =   1125
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "Poblaci�n|T|N|||scaavi|pobclien||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
         Top             =   3270
         Width           =   5115
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "CPostal|T|N|||scaavi|codpobla||N|"
         Text            =   "Text15"
         Top             =   3270
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
         Left            =   240
         MaxLength       =   30
         TabIndex        =   11
         Tag             =   "Provincia|T|N|||scaavi|proclien||N|"
         Text            =   "Text1 Text1 Text1 Text1 Text22"
         Top             =   3960
         Width           =   6000
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
         Left            =   240
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Direccion/Dpto.|N|S|0|9999|scaavi|coddirec|000|N|"
         Text            =   "Text1"
         Top             =   1290
         Width           =   810
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
         Index           =   12
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   1290
         Width           =   5115
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
         Index           =   4
         Left            =   240
         TabIndex        =   42
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   41
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         TabIndex        =   33
         Top             =   360
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1245
         ToolTipText     =   "Buscar cliente"
         Top             =   360
         Width           =   240
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
         Index           =   8
         Left            =   240
         TabIndex        =   32
         Top             =   2295
         Width           =   960
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   690
         ToolTipText     =   "Buscar cliente varios"
         Top             =   1695
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   1650
         Width           =   495
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
         Index           =   7
         Left            =   4230
         TabIndex        =   30
         Top             =   1650
         Width           =   960
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
         Index           =   9
         Left            =   240
         TabIndex        =   29
         Top             =   2985
         Width           =   960
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
         Index           =   11
         Left            =   240
         TabIndex        =   28
         Top             =   3675
         Width           =   1005
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   12
         Left            =   810
         ToolTipText     =   "Buscar direc./dpto"
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Dpto"
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
         TabIndex        =   27
         Top             =   1005
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Enabled         =   0   'False
         Height          =   240
         Index           =   9
         Left            =   1245
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   2985
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   870
      Width           =   13125
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
         Left            =   8025
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Operador|N|N|0|9999|scaavi|codtraba|0000|N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   810
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
         Index           =   2
         Left            =   8850
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   240
         Width           =   4035
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
         Left            =   4185
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Aviso|F|N|||scaavi|fechaavi|dd/mm/yyyy|N|"
         Top             =   240
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
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
         Left            =   1185
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "N� Aviso|N|S|0||scaavi|numaviso|0000000|S|"
         Text            =   "Text1 7"
         Top             =   250
         Width           =   1050
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   7740
         Picture         =   "frmRepAvisosGR.frx":002B
         Tag             =   "-1"
         ToolTipText     =   "Buscar trabajador"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Operador"
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
         Index           =   21
         Left            =   6705
         TabIndex        =   23
         Top             =   255
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha aviso"
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
         Index           =   14
         Left            =   2670
         TabIndex        =   22
         Top             =   255
         Width           =   1185
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3885
         Picture         =   "frmRepAvisosGR.frx":0A2D
         ToolTipText     =   "Buscar fecha"
         Top             =   255
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "N� Aviso"
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
         Index           =   50
         Left            =   240
         TabIndex        =   21
         Top             =   255
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   135
      TabIndex        =   18
      Top             =   7965
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   195
         TabIndex        =   19
         Top             =   135
         Width           =   1755
      End
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
      Left            =   10995
      TabIndex        =   16
      Top             =   8010
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4080
      Top             =   7440
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
      Left            =   12240
      TabIndex        =   37
      Top             =   8010
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmRepAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

                              

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBasico2 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmB1 As frmBasico2 'Form para busquedas
Attribute frmB1.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

Private WithEvents frmC As frmBasico2 'Form Mto Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCV As frmBasico2 'frmFacClientesV  'Form Mto Clientes Varios
Attribute frmCV.VB_VarHelpID = -1
Private WithEvents frmT As frmBasico2 'frmAdmTrabajadores  'Form Mto Trabajadores
Attribute frmT.VB_VarHelpID = -1


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


'Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas


'Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

'Dim PrimeraVez As Boolean


Dim EsDeVarios As Boolean
Private CodTipoMov As String

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String 'Para el ORDER BY de la consulta
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Private EsCabecera2 As Boolean
Private HaCambiadoCP As Boolean
Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


Private Sub cboSituacion_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub cboSituacion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
                
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificarCabAlbaran Then
                    TerminaBloquear
                    PosicionarData
                    
                    
                    'Ahora mandaremos el email
                    Me.Refresh
                    DoEvents
                    Screen.MousePointer = vbHourglass
                    EnviarEmail
                    Screen.MousePointer = vbDefault

                    
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ModificarCabAlbaran() As Boolean
Dim b As Boolean
Dim SQL As String

    On Error GoTo EModificaAlb

    conn.BeginTrans
    
    'Si es cliente de varios actualizar datos cliente en tabla:sclvar
    b = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    
    If b Then
        b = ModificaDesdeFormulario(Me, 1)

'        If b Then
            'comprobar si se ha cambiado el cliente
            'o si se ha cambiado la fecha del albaran
'            If (CInt(Me.Data1.Recordset!CodClien) <> CInt(Text1(4).Text)) Or (CDate(Data1.Recordset!FechaAlb) <> CDate(Text1(1).Text)) Then
'                'si hay numeros de serie en ese albaran, actualizamos el cliente
'                'al nuevo cliente
'                SQL = "UPDATE sserie SET codclien=" & DBSet(Text1(4).Text, "N") & ","
'                SQL = SQL & " fechavta=" & DBSet(Text1(1).Text, "F")
'                SQL = SQL & " WHERE codtipom='" & CodTipoMov & "'" & " AND numalbar=" & Data1.Recordset!NumAlbar & " and fechavta=" & DBSet(Data1.Recordset!FechaAlb, "F")
'                Conn.Execute SQL
'
'                'Modificar el cliente en la smoval
'                SQL = "UPDATE smoval SET codigope=" & DBSet(Text1(4).Text, "N") & ","
'                SQL = SQL & " fechamov=" & DBSet(Text1(1).Text, "F")
'                SQL = SQL & ", horamovi= concat(" & DBSet(Text1(1).Text, "F") & ",hour(horamovi),':',minute(horamovi),':',second(horamovi))"
'                SQL = SQL & " WHERE detamovi='" & CodTipoMov & "'" & " AND document=" & DBSet(CStr(Data1.Recordset!NumAlbar), "T") & " and fechamov=" & DBSet(Data1.Recordset!FechaAlb, "F")
'                Conn.Execute SQL
'            End If
'        End If
    End If
    
EModificaAlb:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    ModificarCabAlbaran = b
End Function




Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
'            LimpiarDataGrids
            PonerModo 0
            PonerFoco Text1(0)
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
    End Select
End Sub


Private Sub BotonAnyadir()
'A�adir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
Dim cad As String
Dim RS As ADODB.Recordset

    LimpiarCampos 'Vac�a los TextBox
    'Poner los grid sin apuntar a nada
'    LimpiarDataGrids
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    NomTraba = ""
    'Poner el nombre del trabajador que esta conectado
    Text1(2).Text = PonerTrabajadorConectado(NomTraba)
    Text2(2).Text = NomTraba

    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    cboSituacion.ListIndex = 0
    PonerFoco Text1(1)
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
'        LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
    
End Sub


Private Sub BotonVerTodos()
Dim cad As String
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        EsCabecera2 = True
'        cad = " codtipom='" & CodTipoMov & "'"
        cad = ""
        MandaBusquedaPrevia cad
    Else
        LimpiarCampos
'        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4

    PonerFoco Text1(1)
   
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsClienteVarios(Text1(4).Text)
    BloquearDatosCliente (DeVarios)
End Sub

Private Function PuedeRealizarAccion(EsEliminar As Boolean) As Boolean
Dim Rc As Byte

    PuedeRealizarAccion = False
    
    If Not (Me.Data1.Recordset Is Nothing) Then
        
        If Not Data1.Recordset.EOF Then
        
            If cboSituacion.ListIndex > 0 Then
                If cboSituacion.ListIndex = 3 Then
                
                    'Para eliminar dejo k borre las cerrdas
                    If Not EsEliminar Then CadenaDesdeOtroForm = "Esta cerrada"
                Else
                     If cboSituacion.ListIndex = 1 Then
                        'Ya esta en reparacion. Creo que no debo dejar de pasar al formulario
                        CadenaDesdeOtroForm = "Ya esta en reparaci�n"
                    Else
                        CadenaDesdeOtroForm = ""  'DEJO PASAR
                    End If
                End If
            Else
                CadenaDesdeOtroForm = ""
            End If
            
        End If
    Else
        CadenaDesdeOtroForm = "No hay datos seleccionados"
    End If
    If CadenaDesdeOtroForm <> "" Then
                                'Dejare seguir
        'If vUsu.Nivel = 0 Then
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & "�Continuar?"
            NumRegElim = vbQuestion + vbYesNo
        'Else
        '    NumRegElim = vbExclamation
        'End If
        CadenaDesdeOtroForm = CStr(Abs(MsgBox(CadenaDesdeOtroForm, NumRegElim) = vbYes))
        If CadenaDesdeOtroForm = "1" Then PuedeRealizarAccion = True
        CadenaDesdeOtroForm = ""
        NumRegElim = 0
    Else
        PuedeRealizarAccion = True
    End If
    
End Function


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Mantenimientos (scaman)
' y los registros correspondientes de las tablas de lineas (sliman y slima1)
Dim cad As String
Dim NumAlbElim As Long

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    If Not PuedeRealizarAccion(True) Then Exit Sub
    
    cad = "Cabecera de Avisos." & vbCrLf
    cad = cad & "------------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Aviso:            "
    cad = cad & vbCrLf & "N�:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")
    cad = cad & vbCrLf & vbCrLf & " �Desea Eliminarlo? "
      
    Screen.MousePointer = vbHourglass
       
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumAlbElim = Data1.Recordset.Fields(0).Value
        
        If Not Eliminar(NumAlbElim) Then
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            PosicionarDataTrasEliminar
        End If
        
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub



Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
'    If hcoCodMovim <> "" And Not Data1.Recordset.EOF And Modo <> 5 Then PonerCadenaBusqueda
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    Me.imgBuscar(4).Picture = Me.imgBuscar(2).Picture
    Me.imgBuscar(6).Picture = Me.imgBuscar(2).Picture
    Me.imgBuscar(9).Picture = Me.imgBuscar(2).Picture
    Me.imgBuscar(12).Picture = Me.imgBuscar(2).Picture
    Me.imgBuscar(13).Picture = Me.imgBuscar(2).Picture


    'ICONITOS DE LA BARRA
'    btnAnyadir = 5
'    btnPrimero = 17
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Bot�n Buscar
'        .Buttons(2).Image = 2   'Bot�n Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
''        .Buttons(10).Image = 10 'Mto Lineas Ofertas
''        .Buttons(11).Image = 33 'N� Serie si lineas con articulos de control N� serie
''        .Buttons(12).Image = 26 'GEnerar factura
'
'        .Buttons(9).Image = 26  'Cambiar visitado
'        .Buttons(10).Image = 42
'        .Buttons(11).Image = 27 'Imprimir Pedido
'        .Buttons(12).Image = 16 'Imprimir Pedido
'        .Buttons(14).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
'    End With
      
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 37 ' cambiar visitado
        .Buttons(2).Image = 38 ' cerrar aviso
        .Buttons(3).Image = 39 ' entrada equipo
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
      
    LimpiarCampos   'Limpia los campos TextBox
    CargarComboSituacion
    
    VieneDeBuscar = False
    CodTipoMov = "AVI" 'Avisos de averias de clientes
      
    'Comprobar si es Departamento o Direccion
    Me.Label1(1).Caption = DevuelveTextoDepto(True)
    If vParamAplic.TieneCRM Then
        Label1(4).Caption = "Observaciones CRM"
    Else
        Label1(4).Caption = "Observaciones internas"
    End If
    
    
    
    '## A mano
    NombreTabla = "scaavi"
    Ordenacion = " ORDER BY fechaavi,numaviso "
 
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where numaviso=-1"
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
'    PrimeraVez = True
    
    PonerModo 0
End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cboSituacion.ListIndex = -1
    Me.chkVisitado.Value = 0
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Llama desde Prismatico Direcciones/Departamentos
        Text1(12).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
        Text2(12).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB1_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
        cadB = Aux
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Text1(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000000")
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim Indice As Byte
Dim devuelve As String

    Indice = 9
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve) 'Poblacion
    'provincia
    Text1(Indice + 2).Text = devuelve
End Sub


Private Sub frmCV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes Varios
Dim Indice As Byte

    Indice = 6
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosClienteVario (Text1(Indice).Text)
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag) + 1
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim Indice As Byte

    Indice = Val(Me.imgBuscar(2).Tag)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 4 'Cod. Cliente
            PonerFoco Text1(4)
            Set frmC = New frmBasico2
            AyudaClientes frmC, Text1(4)
            Set frmC = Nothing
            
        Case 6 'NIF para cliente de Varios
            Set frmCV = New frmBasico2
            AyudaClientesV frmCV, Text1(Index)
            Set frmCV = Nothing
            
        Case 12 'Cod. Direc.
             'Mostrar las Direc. o Dptos del cliente seleccionado
             If Trim(Text1(4).Text) = "" Then
                MsgBox "Debe seleccionar un cliente.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
             Else
                EsCabecera2 = False
                MandaBusquedaPrevia " codclien= " & Val(Text1(4).Text)
             End If
             
        Case 2, 13 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
            Me.imgBuscar(2).Tag = Index

            Set frmT = New frmBasico2
            AyudaTrabajadores frmT, Text1(Index)
            Set frmT = Nothing
            
        Case 9 'Cod. Postal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            VieneDeBuscar = True
    End Select
    PonerFoco Text1(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   Indice = Index + 1
   Me.imgFecha(0).Tag = Index
   
    PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
'    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
'    Else   'Eliminar Albaran
         BotonEliminar
'    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Aviso
    BotonImprimir 408, False '408: Informe de Aviso de averia
End Sub


Private Sub mnModificar_Click()
    'Modificar albaran
    If Not PuedeRealizarAccion(False) Then Exit Sub
        
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub


Private Sub mnNuevo_Click()
    'A�adir Cabecera de Pedidos
    BotonAnyadir
End Sub


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub




'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
   
    If Index <> 3 Then ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
If Not Text1(Index).MultiLine Then KEYdown KeyCode       'Con las flechas cuando
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Index <> 3 And Index <> 14) Or (Index = 3 And Modo = 1) Or (Index = 14 And Modo = 1) Then
        If KeyAscii = teclaBuscar Then
            Select Case Index
                Case 1: KEYFecha KeyAscii, 0 'fecha
                Case 2: KEYBusqueda KeyAscii, 2 'trabajador
                Case 4: KEYBusqueda KeyAscii, 4 'cliente
                Case 6: KEYBusqueda KeyAscii, 6 'cliente de varios
                Case 9: KEYBusqueda KeyAscii, 9 'poblacion
                Case 12: KEYBusqueda KeyAscii, 12 'direccion/dpto
                Case 13: KEYBusqueda KeyAscii, 13 'tecnico trabajador
            End Select
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
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
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha Aviso
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
        Case 2, 13 'Cod Operador
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Cod. Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Modo=1 Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
                Else
                    PonerDatosCliente (Text1(Index).Text)
                End If
                If Not Text1(4).Locked Then
                    PonerFoco Text1(5)
                Else
                    PonerFoco Text1(12)
                End If
            Else
                LimpiarDatosCliente
            End If
            
        Case 6 'NIF
            If Not EsDeVarios Then Exit Sub
            'si no se ha modificado el nif del cliente no hacer nada (Modo 4=Modificar)
            If (Modo = 4) Then
                If (Text1(6).Text = Data1.Recordset!nifClien) Then Exit Sub
            End If
            PonerDatosClienteVario (Text1(Index).Text)
                     
        Case 9 'Cod. Postal
             If Text1(Index).Locked Then Exit Sub
             If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
                Exit Sub
             End If
             If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                 Text1(Index + 2).Text = devuelve
             End If
             VieneDeBuscar = False
            
        Case 12 'Cod. Direc
            If Text1(Index).Text = "" Then
                'Text1(Index + 1).Text = ""
                Text2(12).Text = ""
                Exit Sub
            End If
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            
            'Comprobar que el cliente seleccionada tiene esa direccion
            If PonerDptoEnCliente Then
                'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
                If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
            Else
                PonerFoco Text1(Index)
            End If
            
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    
    If chkVistaPrevia = 1 Then
        EsCabecera2 = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String

    'Llamamos a al form
    '##A mano
    cad = ""
    If EsCabecera2 Then
'        cad = cad & ParaGrid(Text1(0), 17, "N� Aviso")
'        cad = cad & ParaGrid(Text1(1), 15, "Fecha")
'        cad = cad & ParaGrid(Text1(4), 15, "Cliente")
'        cad = cad & ParaGrid(Text1(5), 53, "Nombre Cliente")
'        tabla = NombreTabla
'        Titulo = "Avisos"
'        devuelve = "0|1|"

        Set frmB1 = New frmBasico2
        AyudaAvisos frmB1
        Set frmB1 = Nothing

        Exit Sub

    Else
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
        Titulo = Titulo & Text1(4).Text & " - " & Text1(5).Text
'        cad = cad & "Cod. " & Desc & "|sdirec|coddirec|N||15�"
'        cad = cad & "Desc. " & Desc & "|sdirec|nomdirec|T||35�"
'        tabla = "sdirec"
'        devuelve = "0|1|"

        Set frmB = New frmBasico2
        AyudaMantenimientosAux frmB, Titulo, Desc, Text1(0), cadB
        Set frmB = Nothing


    End If
           
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
'        frmB.vselElem = 0
'        frmB.vConexionGrid = conAri  'Conexi�n a BD: Ariges
'        If Not EsCabecera2 Then frmB.Label1.FontSize = 11
''        frmB.vBuscaPrevia = chkVistaPrevia
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
'    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(kCampo)
            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    
    Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "straba", "nomtraba", "codtraba")
    Text2(12).Text = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", Text1(4).Text, "N", , "coddirec", Text1(12).Text, "N")
    Text2(13).Text = PonerNombreDeCod(Text1(13), conAri, "straba", "nomtraba", "codtraba")

    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
'Dim i As Byte
Dim NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1

        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    'campo n� aviso es un contador y siempre bloqueado salvo al buscar
    BloquearTxt Text1(0), (Modo <> 1), True
    'campo fecha aviso es clave primaria
    BloquearTxt Text1(1), (Modo <> 1 And Modo <> 3)
    
    
    'El nombre del dpto no lo modificamos      Lo quito yo, er david
    'BloquearTxt Text1(13), (Modo <> 1)
    b = False
    If Modo = 1 Then
        b = True
    Else
        If (Modo = 3) Then
            b = True
        Else
            If (Modo = 4) Then
                If Me.cboSituacion.ListIndex = 0 Then
                    b = True
                Else
                    'Antes Feb. 10. Solo podia admin
                    b = True   'vUsu.Nivel = 0 'solo el admin
                End If
            End If
        End If
    End If
    
    Me.cboSituacion.Enabled = b
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
'    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(0).Enabled = b And Modo <> 4
'    Next i
    
    BloquearImg imgBuscar(2), Not b
    BloquearImg imgBuscar(4), Not b
    BloquearImg imgBuscar(6), Not b
    BloquearImg imgBuscar(9), Not b
    BloquearImg imgBuscar(12), Not b
    BloquearImg imgBuscar(13), Not b
    
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    
    
    If Text2(2).Text = "" Or Text2(13).Text = "" Then
        MsgBox "Faltan datos: tomador aviso/ t�cnico del aviso", vbExclamation
        Exit Function
    End If
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    'campo Amliacion Linea y ENTER
    If Index = 16 And KeyAscii = 13 Then PonerFocoBtn Me.cmdAceptar
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 8 'Imprimir Albaran
            mnImprimir_Click
    End Select
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Function Eliminar(NumAlbElim As Long) As Boolean
Dim SQL As String
Dim b As Boolean
Dim vTipoMov As CTiposMov

    On Error GoTo FinEliminar

    conn.BeginTrans
    SQL = ObtenerWhereCP(True)
    
    SQL = "DELETE FROM " & NombreTabla & " " & SQL
    conn.Execute SQL
            
    'Devolvemos contador, si no estamos actualizando
    Set vTipoMov = New CTiposMov
    b = CBool(vTipoMov.DevolverContador(CodTipoMov, NumAlbElim))
    Set vTipoMov = Nothing
        
FinEliminar:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Eliminando Aviso de aver�a.", Err.Description
    End If
    If Not b Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
    End If
    Eliminar = b
End Function



Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
'         If SituarDataGral(Data1, Text1(30).Text, "T", Text1(0).Text, "N", Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String

    On Error Resume Next
    
    SQL = " " & NombreTabla & ".numaviso= " & Val(Text1(0).Text)
'    If EsHistorico Then
    SQL = SQL & " AND " & NombreTabla & ".fechaavi=" & DBSet(Text1(1).Text, "F")
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
''--------------------------------------------------------------------
'' MontaSQlCarga:
''   Bas�ndose en la informaci�n proporcionada por el vector de campos
''   crea un SQl para ejecutar una consulta sobre la base de datos que los
''   devuelva.
'' Si ENLAZA -> Enlaza con el data1
''           -> Si no lo cargamos sin enlazar a ningun campo
''--------------------------------------------------------------------
'Dim SQL As String
'
'    SQL = "SELECT codtipom, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, origpre, dtoline1, dtoline2, importel "
'    SQL = SQL & " FROM " & NomTablaLineas
'    If enlaza Then
'        SQL = SQL & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
''        If EsHistorico Then SQL = SQL & " and fechaalb='" & Format(Text1(1).Text, FormatoFecha) & "'"
'    Else
'        SQL = SQL & " WHERE numalbar = -1"
'    End If
'    SQL = SQL & " Order by codtipom, numalbar, numlinea"
'    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean

        b = (Modo = 2) 'Or (Modo = 5 And ModificaLineas = 0))
        'Insertar
        Toolbar1.Buttons(1).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(2).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(3).Enabled = b
        Me.mnEliminar.Enabled = b
            
        b = (Modo = 2)
        'Mantenimiento lineas
        Toolbar5.Buttons(1).Enabled = b
        Toolbar5.Buttons(2).Enabled = b
        Toolbar5.Buttons(3).Enabled = b
        
        'Imprimir
        Toolbar1.Buttons(8).Enabled = (Modo = 2)
        Me.mnImprimir.Enabled = (Modo = 2)
        
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub


Private Function InsertarAviso(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        MenError = DevuelveDesdeBDNew(conAri, NombreTabla, "numaviso", "numaviso", Text1(0).Text, "N", , "fechaavi", Text1(1).Text, "F")
        If MenError <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    MenError = ""
    
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Avisos (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del cliente si es de varios
    If EsDeVarios Then
        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
        MenError = "Modificando datos cliente varios"
        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
    End If
           
    If bol Then
        MenError = "Error al actualizar el contador del movimiento."
        vTipoMov.IncrementarContador (CodTipoMov)
    End If
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Aviso." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarAviso = True
            
            
            
            
        Else
            conn.RollbackTrans
            InsertarAviso = False
        End If
End Function


Private Sub LimpiarDatosCliente()
Dim I As Byte

    For I = 4 To 12
        Text1(I).Text = ""
    Next I
    Text2(12).Text = ""
End Sub
    


Private Sub BotonImprimir(OpcionListado As Integer, EnvioMail As Boolean)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Aviso para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 16
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub


    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N� de Aviso
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'N� Aviso
        devuelve = "{" & NombreTabla & ".numaviso}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
        
        'El campo fecha tambien es clave primaria
        'para Crystal
        devuelve = Text1(1).Text
        devuelve = "{" & NombreTabla & ".fechaavi}=Date(" & Year(devuelve) & "," & Month(devuelve) & "," & Day(devuelve) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'para MySQL
        devuelve = "{" & NombreTabla & ".fechaavi}='" & Format(Text1(1).Text, FormatoFecha) & "'"
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    End If


    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
    
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .SeleccionaRPTCodigo = pRptvMultiInforme
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = EnvioMail
            .Opcion = OpcionListado
            .Titulo = "Avisos de averias."
            .NombreRPT = nomDocu
            .NombrePDF = pPdfRpt
            .ConSubInforme = True
            .Show vbModal
    End With
End Sub



Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarAviso(SQL, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ahora mandaremos el email
                Me.Refresh
                DoEvents
                Screen.MousePointer = vbHourglass
                EnviarEmail
                Screen.MousePointer = vbDefault
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
        PonerCampos
    Else
        LimpiarCampos
'        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Sub PonerDatosCliente(codClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If codClien = "" Then
        LimpiarDatosCliente
        Exit Sub
    End If

    Set vCliente = New CCliente
    
    'si se ha modificado el cliente volver a cargar los datos
    If vCliente.Existe(codClien) Then
        If vCliente.LeerDatos(codClien) Then
            'si el cliente esta bloqueado salimos
            If vCliente.ClienteBloqueado(2, False) Then
                LimpiarDatosCliente
                Set vCliente = Nothing
                Exit Sub
            End If
            
'            EsDeVarios = vCliente.EsClienteVarios(Text1(4).Text)
            EsDeVarios = vCliente.DeVarios
            BloquearDatosCliente (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el cliente no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!codClien) Then
                    Set vCliente = Nothing
                    Exit Sub
                End If
            End If
            
'            If (Not EsDeVarios) Or (EsDeVarios And modo = 3) Then
            Text1(4).Text = vCliente.Codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vCliente.Nombre  'Nom clien
                Text1(8).Text = vCliente.Domicilio
                Text1(9).Text = vCliente.CPostal
                Text1(10).Text = vCliente.Poblacion
                Text1(11).Text = vCliente.Provincia
                Text1(6).Text = vCliente.NIF
                Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            End If

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
            End If
                           
                           
            'Me cargo lo que habia en departamentos
            Text2(12).Text = ""
            Text1(12).Text = ""
            'Comprobar si el cliente tiene cobros pendientes
'            ComprobarCobrosCliente CodClien, Text1(1).Text
        End If
    Else
        LimpiarDatosCliente
    End If
    Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub


Private Sub PonerDatosClienteVario(nifClien As String)
Dim vCliente As CCliente
Dim b As Boolean
   
    If nifClien = "" Then Exit Sub
   
    Set vCliente = New CCliente
    b = vCliente.LeerDatosCliVario(nifClien)
    If b Then Text1(5).Text = vCliente.Nombre         'Nom clien
    Text1(8).Text = vCliente.Domicilio
    Text1(9).Text = vCliente.CPostal
    Text1(10).Text = vCliente.Poblacion
    Text1(11).Text = vCliente.Provincia
    Text1(7).Text = DBLet(vCliente.TfnoClien, "T")
            
    Set vCliente = Nothing
End Sub


Private Sub BloquearDatosCliente(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(9).visible = bol
        Me.imgBuscar(9).Enabled = bol
        Me.imgBuscar(6).Enabled = bol
        Me.imgBuscar(6).visible = bol
        
        For I = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
    End If
End Sub


Private Function ActualizarClienteVarios(clien As String, NIF As String) As Boolean
Dim vCliente As CCliente

    On Error GoTo EActualizarCV

    ActualizarClienteVarios = False
    
    Set vCliente = New CCliente
    If EsClienteVarios(clien) Then
        vCliente.NIF = NIF
        vCliente.Nombre = Text1(5).Text
        vCliente.Domicilio = Text1(8).Text
        vCliente.CPostal = Text1(9).Text
        vCliente.Poblacion = Text1(10).Text
        vCliente.Provincia = Text1(11).Text
        vCliente.TfnoClien = Text1(7).Text
        vCliente.ActualizarClienteV (NIF)
    End If
    Set vCliente = Nothing
    
    ActualizarClienteVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarClienteVarios = False
    Else
        ActualizarClienteVarios = True
    End If
End Function


Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.Codigo = Text1(4).Text
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(12).Text, NomDpto) Then
        Text2(12).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function




Private Sub CargarComboSituacion()
'### Combo Tipo Facturaci�n
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Abierta, 1-En Reparacion, 2-Pendiente, 3-Cerrado

    Me.cboSituacion.Clear
    cboSituacion.AddItem "Abierta"
    cboSituacion.ItemData(cboSituacion.NewIndex) = 0

    cboSituacion.AddItem "En reparaci�n"
    cboSituacion.ItemData(cboSituacion.NewIndex) = 1
    
    cboSituacion.AddItem "Pendiente"
    cboSituacion.ItemData(cboSituacion.NewIndex) = 2
    
    cboSituacion.AddItem "Cerrado"
    cboSituacion.ItemData(cboSituacion.NewIndex) = 3
End Sub


Private Sub CambiarSituacionVisitado()
    'Esto es a mano, a pi�on
    On Error GoTo EC
    Screen.MousePointer = vbHourglass
    NumRegElim = 1
    If Me.chkVisitado.Value = 1 Then NumRegElim = 0
    Me.chkVisitado.Value = NumRegElim
    CadenaDesdeOtroForm = "UPDATE scaavi SET visitado = " & Me.chkVisitado.Value
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " WHERE numaviso =" & Val(Text1(0).Text)
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND  fechaavi = '" & Format(Text1(1).Text, FormatoFecha) & "'"
    conn.Execute CadenaDesdeOtroForm
    
    PosicionarData
    
    Screen.MousePointer = vbDefault
    Exit Sub
EC:
    MuestraError Err.Number
    Screen.MousePointer = vbDefault
End Sub



Private Sub EnviarEmail()
Dim Des As String

    On Error GoTo EEnvio
    If Dir(App.Path & "\Docum.pdf") <> "" Then Kill App.Path & "\Docum.pdf"
        

    'Obtengo el mail del TOMADOR
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "straba", "maitraba", "codtraba", Text1(2).Text, "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "El operador que toma el aviso no tiene e-mail", vbExclamation
        Exit Sub
    End If
    Des = CadenaDesdeOtroForm
                       
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "straba", "maitraba", "codtraba", Text1(13).Text, "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "El t�cnico no tiene e-mail", vbExclamation
        Exit Sub
    End If
    Des = Des & "|" & CadenaDesdeOtroForm & "|" 'TOOODO por no crear mas variables
                       
    If MsgBox("      �Desea enviar el e-mail?      ", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                       
        BotonImprimir 408, True
        'Si esta creado es que lo ha ecxportado a pdf bien
        If Dir(App.Path & "\Docum.pdf") <> "" Then
                       
                       
        'Llamaremos a enviar mail con los datos que me de la gana... vamos digo yo
        'Nombre para|email para|Asunto|Mensaje|mailtomador|nombretomador|
        frmEMail.Opcion = 3
        frmEMail.DatosEnvio = Text2(13).Text & "|" & RecuperaValor(Des, 2) & "|[ARIGES]: Aviso de " & Text1(5).Text & "|"
        'Peque�o texto para el mensaje
        CadenaDesdeOtroForm = "Tomado por : " & Text2(2).Text & vbCrLf & vbCrLf & vbCrLf & "Cliente: " & Text1(5).Text & vbCrLf
        For NumRegElim = 6 To 7
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & Label1(NumRegElim).Caption & ": " & Text1(NumRegElim) & vbCrLf
        Next NumRegElim
        
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Label1(45).Caption & ": " & Text1(3).Text & vbCrLf
        NumRegElim = 0
        frmEMail.DatosEnvio = frmEMail.DatosEnvio & CadenaDesdeOtroForm & "|"
        'Datos del enviante del mail
        frmEMail.DatosEnvio = frmEMail.DatosEnvio & RecuperaValor(Des, 1) & "|" & Text2(2).Text & "|"
        CadenaDesdeOtroForm = ""

        frmEMail.Show vbModal
    Else
        MsgBox "Documento PDF no encontrado", vbExclamation
    End If
    Exit Sub
EEnvio:
    MuestraError Err.Number, "Enviar mail"

End Sub

Private Sub LanzarReparaciones()


    CadenaDesdeOtroForm = "No data selected"
    If Not (Me.Data1.Recordset Is Nothing) Then
        
        If Not Data1.Recordset.EOF Then
        
            If cboSituacion.ListIndex > 0 Then
                'Ya esta en reparacion. Creo que no debo dejar de pasar al formulario
                If cboSituacion.ListIndex = 3 Then
                    CadenaDesdeOtroForm = "Aviso cerrado"
                Else
                    CadenaDesdeOtroForm = "Ya esta en reparaci�n"
                    If cboSituacion.ListIndex = 2 Then CadenaDesdeOtroForm = ""
                End If
                
            Else
                EsDeVarios = EsClienteVarios(CStr(Data1.Recordset!codClien))
                If EsDeVarios Then
                    CadenaDesdeOtroForm = "Cliente varios no se le pueden asignar articulos con numero de serie"
                Else
                    CadenaDesdeOtroForm = ""
                End If
                
            End If
            
        End If
            
    End If
    If CadenaDesdeOtroForm <> "" Then
        MsgBox CadenaDesdeOtroForm, vbExclamation
        CadenaDesdeOtroForm = ""
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
        '       Codigo y fecha
        CadenaDesdeOtroForm = Val(Text1(0).Text) & "|" & Text1(1).Text & "|"
        '       codcli, nomcli (ya que para varios se puede modificar
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(4).Text & "|" & Text1(5).Text & "|"
        '       Departamento     Desc DPTO
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(12).Text & "|" & Text2(12).Text & "|"
        '  NIF     TELEFONO
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(6).Text & "|" & Text1(7).Text & "|"
        '       Domicilio    codpobla
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(8).Text & "|" & Text1(9).Text & "|"
        '       descpobla     provincia
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(10).Text & "|" & Text1(11).Text & "|"
        frmRepEntReparacionesGR.EntradaEquipo = CadenaDesdeOtroForm
        frmRepEntReparacionesGR.ControlRep = False
        frmRepEntReparacionesGR.EsHistorico = False
        frmRepEntReparacionesGR.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            DoEvents
            'Ha metido la reparacion. Ahora pongo el campo del combo a EN reparacion
            CadenaDesdeOtroForm = "UPDATE scaavi SET situacio = 1"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & " WHERE numaviso =" & Val(Text1(0).Text)
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND  fechaavi = '" & Format(Text1(1).Text, FormatoFecha) & "'"
            conn.Execute CadenaDesdeOtroForm
            PosicionarData
            CadenaDesdeOtroForm = ""
            'Ahora pongo el combo de situacion en 1
            Me.cboSituacion.ListIndex = 1  'situacio=1
            
        End If
    Screen.MousePointer = vbDefault
End Sub





Private Sub CerrarAviso()
Dim Albaran As Long
 CadenaDesdeOtroForm = "No data selected"
    If Not (Me.Data1.Recordset Is Nothing) Then
        
        If Not Data1.Recordset.EOF Then
            If cboSituacion.ListIndex = 3 Then
                CadenaDesdeOtroForm = "Aviso cerrado"
            ElseIf cboSituacion.ListIndex > 0 Then
                CadenaDesdeOtroForm = "Ya esta en reparaci�n"
                'Ya esta en reparacion. Creo que no debo dejar de pasar al formulario
                If cboSituacion.ListIndex = 2 Then CadenaDesdeOtroForm = ""
                
            Else
                CadenaDesdeOtroForm = ""
                For NumRegElim = 5 To 11
                    If Text1(NumRegElim).Text = "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & RecuperaValor(Text1(NumRegElim).Tag, 1) & vbCrLf
                Next NumRegElim
                If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = "Campos cliente obligatorios: " & vbCrLf & CadenaDesdeOtroForm
                
                If Text1(2).Text = "" Or Text1(13).Text = "" Then
                    If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & vbCrLf
                    CadenaDesdeOtroForm = vbCrLf & CadenaDesdeOtroForm & "Campos trabajadores son obligatorios para cerrar el aviso" & vbCrLf
                End If
            End If
            
        End If
            
    End If
    If CadenaDesdeOtroForm <> "" Then
        MsgBox CadenaDesdeOtroForm, vbExclamation
        CadenaDesdeOtroForm = ""
        Exit Sub
    End If
    
    
    'Voy a ver si el departmanto EXISTE
    If Text1(12).Text <> "" Then
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", "codclien = " & Text1(4).Text & " AND coddirec ", Text1(12).Text, "N")
        If CadenaDesdeOtroForm = "" Then
            MsgBox "No existe el " & Label1(1).Caption & " para el cliente: " & Text1(4).Text, vbExclamation
            
            Exit Sub
        End If
        CadenaDesdeOtroForm = ""
    End If
    
    
    
    
    CadenaDesdeOtroForm = ""
    frmListado2.Opcion = 19
    frmListado2.Show vbModal
    
    
        
        
        If CadenaDesdeOtroForm <> "" Then
            Screen.MousePointer = vbHourglass
            conn.BeginTrans
            Albaran = GenerarAlbaran2
            If Albaran > 0 Then
                conn.CommitTrans
            
                'Ha metido la reparacion. Ahora pongo el campo del combo a EN reparacion
                CadenaDesdeOtroForm = "UPDATE scaavi SET situacio = 3" 'cerrada
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " WHERE numaviso =" & Val(Text1(0).Text)
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND  fechaavi = '" & Format(Text1(1).Text, FormatoFecha) & "'"
                conn.Execute CadenaDesdeOtroForm
                PosicionarData
                'Ahora pongo el combo de situacion en 3
                Me.cboSituacion.ListIndex = 3  'Cerrada
                CadenaDesdeOtroForm = ""
                
                'AHora lanzo abrir formulario de entrada de reparaciones
                'Y que se ponga en modo insertar linea
                If vParamAplic.TipoFormularioClientes = 0 Then
                    frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
                    frmFacEntAlbaranes2.AlbAvisoGenerado = Albaran
                    frmFacEntAlbaranes2.hcoCodTipoM = "ALR"
                    frmFacEntAlbaranes2.EsHistorico = False
                    frmFacEntAlbaranes2.Show vbModal
                End If
                
                
                
            Else
                conn.RollbackTrans
            End If
        End If
    Screen.MousePointer = vbDefault

    
    
End Sub



Private Function GenerarAlbaran2() As Long
Dim NumAlb As Long
Dim cad As String
Dim vTipoMov As CTiposMov
Dim Cli As CCliente

    On Error GoTo EGenerarAlbaran
    
    GenerarAlbaran2 = 0
    NumAlb = 0
    Set Cli = New CCliente
    If Not Cli.LeerDatos(Text1(4).Text) Then
        Set Cli = Nothing
        Exit Function
    End If
    
    Set vTipoMov = New CTiposMov
    NumAlb = vTipoMov.ConseguirContador("ALR")
    Do
        cad = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", vTipoMov.TipoMovimiento, "T", , "numalbar", CStr(NumAlb), "N")
        If cad <> "" Then
            'Ya existe el contador incrementarlo
            HaDevueltoDatos = True
            vTipoMov.IncrementarContador (vTipoMov.TipoMovimiento)
            NumAlb = vTipoMov.ConseguirContador(vTipoMov.TipoMovimiento)
        Else
            HaDevueltoDatos = False
        End If
    Loop Until Not HaDevueltoDatos
    
    
    'Voy a insertar el albaran """""A MANO """"""
    'codtipom,numalbar,fechaalb,factursn,
    cad = "'ALR'," & NumAlb & ",'" & Format(RecuperaValor(CadenaDesdeOtroForm, 1), FormatoFecha) & "',0,"
    'CodClien , nomclien, domclien, codpobla,
    
    cad = cad & Cli.Codigo & "," & DBSet(Text1(5).Text, "T") & "," & DBSet(Text1(8).Text, "T") & "," & DBSet(Text1(9).Text, "T") & ","
    
    'pobclien, proclien, nifClien, telclien,
    cad = cad & DBSet(Text1(10).Text, "T") & "," & DBSet(Text1(11).Text, "T") & "," & DBSet(Text1(6).Text, "T") & "," & DBSet(Text1(7).Text, "T") & ","
    'CodDirec , nomdirec, referenc,
    cad = cad & DBSet(Text1(12).Text, "T") & "," & DBSet(Text2(12).Text, "T") & ",NULL,"
    'facturkm , cantidkm,
    cad = cad & "0,NULL,"
    'CodTraba , codtrab1, codtrab2,
    cad = cad & Text1(2).Text & "," & Text1(13).Text & "," & Text1(13).Text & ","
    'codagent , codforpa, CodEnvio, DtoPPago, DtoGnral, tipofact,
    cad = cad & Cli.Agente & "," & Cli.ForPago & "," & Cli.FEnvio & "," & DBSet(Cli.DtoPPago, "N", "N") & "," & DBSet(Cli.DtoGnral, "N", "N") & "," & Cli.TipoFactu & ","
    'observa01 , observa02, observa03, observa04, observa05,
    For kCampo = 2 To 6
        cad = cad & DBSet(Trim(RecuperaValor(CadenaDesdeOtroForm, kCampo)), "T", "S") & ","
    Next kCampo
    
    'numpedcl, fecpedcl NumOfert , fecofert, , FecEntre, sementre,
    cad = cad & Text1(0).Text & ",'" & Format(Text1(1).Text, FormatoFecha) & "',NULL,NULL,NULL,NULL,"
    'codtipmf , NumFactu, FecFactu, EsTicket, NumTermi, NumVenta,
    cad = cad & "NULL,NULL,NULL,0,NULL,NULL,"
    
    'Aportacion , pesoalba, portes
    cad = cad & "0,0,0"
    
    'INSER INTO
    CadenaDesdeOtroForm = "codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec,nomdirec,referenc,facturkm,cantidkm,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,observa01,observa02,observa03,observa04,observa05,numpedcl,fecpedcl,numofert,fecofert,fecentre,sementre,codtipmf,numfactu,fecfactu,esticket,numtermi,numventa,aportacion,pesoalba,portes"
    cad = "INSERT INTO scaalb(" & CadenaDesdeOtroForm & ") VALUES (" & cad & ")"
    conn.Execute cad
    
    GenerarAlbaran2 = NumAlb
    
    Exit Function
EGenerarAlbaran:
    MuestraError Err.Number, Err.Description
End Function

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'cambiar visitado
            CambiarSituacionVisitado
        Case 2  'Cerrar AVISO
            CerrarAviso
        Case 3
            LanzarReparaciones
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub