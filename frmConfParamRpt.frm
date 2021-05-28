VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfParamRpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Documentos"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   Icon            =   "frmConfParamRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Informes Disponibles"
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
      Height          =   3390
      Left            =   225
      TabIndex        =   30
      Top             =   5220
      Width           =   10455
      Begin VB.Frame FrameToolAux0 
         Height          =   645
         Left            =   105
         TabIndex        =   31
         Top             =   270
         Width           =   1500
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   150
            TabIndex        =   32
            Top             =   180
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
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
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   90
         TabIndex        =   33
         Top             =   990
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3836
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
            Text            =   "Linea"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   9772
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Documento"
            Object.Width           =   6156
         EndProperty
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   360
         Picture         =   "frmConfParamRpt.frx":000C
         ToolTipText     =   "Nueva "
         Top             =   540
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   2
         Left            =   1080
         Picture         =   "frmConfParamRpt.frx":0A0E
         ToolTipText     =   "Eliminar "
         Top             =   540
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmConfParamRpt.frx":1410
         ToolTipText     =   "Modificar"
         Top             =   540
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   270
      TabIndex        =   28
      Top             =   135
      Width           =   3075
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   29
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
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
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3480
      TabIndex        =   26
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   210
         TabIndex        =   27
         Top             =   210
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
               Object.ToolTipText     =   "Último"
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
      Left            =   9045
      TabIndex        =   25
      Top             =   360
      Width           =   1530
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
      Left            =   1665
      MaxLength       =   100
      TabIndex        =   6
      Tag             =   "Fichero envio mail|T|N|||scryst|pdfrpt|||"
      Text            =   "Text1"
      Top             =   2400
      Width           =   6795
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Impresion directa"
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
      Left            =   8595
      TabIndex        =   1
      Tag             =   "Impr  directa|N|N|||scryst|imprimedirecto|||"
      Top             =   1980
      Width           =   2025
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
      Left            =   300
      MaxLength       =   140
      TabIndex        =   8
      Tag             =   "Linea pie 2|T|S|||scryst|lineapi2|||"
      Text            =   "Text1"
      Top             =   3495
      Width           =   10245
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
      Left            =   300
      MaxLength       =   140
      TabIndex        =   11
      Tag             =   "Linea pie 5|T|S|||scryst|lineapi5|||"
      Text            =   "Text1"
      Top             =   4665
      Width           =   10245
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
      Left            =   300
      MaxLength       =   140
      TabIndex        =   10
      Tag             =   "Linea pie 4|T|S|||scryst|lineapi4|||"
      Text            =   "Text1"
      Top             =   4275
      Width           =   10245
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
      Left            =   300
      MaxLength       =   140
      TabIndex        =   9
      Tag             =   "Linea pie 3|T|S|||scryst|lineapi3|||"
      Text            =   "Text1"
      Top             =   3885
      Width           =   10245
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
      Left            =   9615
      TabIndex        =   13
      Top             =   8805
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   225
      TabIndex        =   20
      Top             =   8730
      Width           =   2640
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
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   165
         Width           =   2280
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
      Index           =   5
      Left            =   300
      MaxLength       =   140
      TabIndex        =   7
      Tag             =   "Linea pie 1|T|S|||scryst|lineapi1|||"
      Text            =   "Text1"
      Top             =   3105
      Width           =   10245
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
      Left            =   9990
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "Revisión ISO|N|S|0|99|scryst|codigrev|00||"
      Text            =   "Te"
      Top             =   1425
      Width           =   570
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
      Index           =   3
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Código ISO|T|S|||scryst|codigiso|||"
      Text            =   "Text1"
      Top             =   1470
      Width           =   1650
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
      Index           =   2
      Left            =   1665
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "Fichero rpt|T|N|||scryst|documrpt|||"
      Text            =   "Text1"
      Top             =   1935
      Width           =   6810
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
      Left            =   1665
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Descripción|T|N|||scryst|nomcryst|||"
      Text            =   "Text1"
      Top             =   1485
      Width           =   3810
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
      Left            =   8445
      TabIndex        =   12
      Top             =   8805
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   9630
      TabIndex        =   14
      Top             =   8820
      Width           =   1065
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
      Left            =   1180
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Código Documento|N|N|||scryst|codcryst||S|"
      Text            =   "Text"
      Top             =   1035
      Width           =   810
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6435
      Top             =   225
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informes disponibles"
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
      Left            =   2250
      TabIndex        =   24
      Top             =   5355
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero envio mail"
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
      Left            =   300
      TabIndex        =   23
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Líneas para el  Pie del Informe"
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
      Left            =   300
      TabIndex        =   22
      Top             =   2820
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Revisión ISO"
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
      Left            =   8625
      TabIndex        =   19
      Top             =   1470
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Código ISO"
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
      Left            =   5670
      TabIndex        =   18
      Top             =   1485
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero rpt"
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
      Left            =   300
      TabIndex        =   17
      Top             =   1965
      Width           =   1125
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   16
      Top             =   1485
      Width           =   1170
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   15
      Top             =   1035
      Width           =   705
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private WithEvents frmB As frmBasico2 'BuscaGrid
Attribute frmB.VB_VarHelpID = -1
Dim HaDevueltoDatos  As Boolean
Dim CadB As String

Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
'4: Modificar


Private Sub Check1_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim vParamRpt As CParamRpt 'Clase Parametros para Reports
Dim cad As String, Indicador As String
Dim actualiza As Boolean

    If DatosOk Then
        'Modifica datos en la Tabla: scryst
'        I = ModificaDesdeFormulario(Me)
        'Actualizar campos de la clase
            Set vParamRpt = New CParamRpt
            vParamRpt.Codigo = Text1(0).Text
            vParamRpt.Descripcion = Text1(1).Text
            vParamRpt.Documento = Text1(2).Text
            vParamRpt.PDFrpt = Text1(10).Text
            vParamRpt.CodigoISO = Text1(3).Text
            If Trim(Text1(4).Text) <> "" Then
                vParamRpt.CodigoRevision = CInt(Text1(4).Text)
            Else
                vParamRpt.CodigoRevision = -1
            End If
            vParamRpt.LineaPie1 = Text1(5).Text
            vParamRpt.LineaPie2 = Text1(6).Text
            vParamRpt.LineaPie3 = Text1(7).Text
            vParamRpt.LineaPie4 = Text1(8).Text
            vParamRpt.LineaPie5 = Text1(9).Text
            vParamRpt.ImprimeDirecto = Check1.Value
        If Modo = 3 Then 'INSERTAR
            actualiza = vParamRpt.Insertar
        ElseIf Modo = 4 Then 'MODIFICAR
            actualiza = vParamRpt.Modificar(Text1(0).Text)
            TerminaBloquear
        End If
        Set vParamRpt = Nothing
        If actualiza = 0 Then 'Inserta o Modifica
            cad = "codcryst=" & Text1(0).Text
            If SituarData(Data1, cad, Indicador) Then
                PonerModo 2
                Me.lblIndicador.Caption = Indicador
            End If
        End If
        PonerFocoBtn Me.cmdSalir
    End If
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 3 'Insertar
            LimpiarCampos
            PonerModo 0
            'PonerFoco Text1(0)
        Case 4 'Modificar
            TerminaBloquear
            If Data1.Recordset.EOF Then
                PonerModo 0
                LimpiarCampos
            Else
                PonerCampos
                PonerModo 2
                Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            End If
    End Select
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'PonerCadenaBusqueda
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
'    With Me.Toolbar1
'        btnPrimero = 11
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1
'        .Buttons(2).Image = 2
'        .Buttons(4).Image = 3   'Anyadir
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(8).Image = 15  'Salir
'        .Buttons(btnPrimero).Image = 6 'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
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
    
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    
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
    
    NombreTabla = "scryst"
    Ordenacion = " ORDER BY codcryst"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    PonerModo 0
    PonerFoco Text1(0)
    limpiar Me
    Me.chkVistaPrevia.Value = 1
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Me.Toolbar1.Buttons(2).Enabled = False
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
    End If
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Image3_Click(Index As Integer)

    If Modo <> 2 Then Exit Sub
    
    If Index > 0 Then
        If ListView1.ListItems.Count = 0 Then Exit Sub
        If ListView1.SelectedItem Is Nothing Then Exit Sub
    End If
    
    
    If Index = 2 Then
        'Eliminar
        CadenaDesdeOtroForm = "Va a eliminar el Informe : " & vbCrLf & "Descripcion: " & ListView1.SelectedItem.SubItems(1)
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & "Informe(RPT) : " & ListView1.SelectedItem.SubItems(2) & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbYes Then
              CadenaConsulta = "DELETE FROM scryst2 WHERE codcryst = " & Data1.Recordset!codcryst
              CadenaConsulta = CadenaConsulta & " AND linea =" & ListView1.SelectedItem.Text
              If ejecutar(CadenaConsulta, False) Then ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
    Else
        'NUEVO - MODIFICAR
        If Index = 0 Then
            CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "max(linea)", "scryst2", "codcryst", CStr(Data1.Recordset!codcryst))
            If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "0"
            CadenaDesdeOtroForm = CStr(Val(CadenaDesdeOtroForm) + 1)
            CadenaConsulta = "1|" & CadenaDesdeOtroForm & "|||"
        Else
            CadenaConsulta = "0|" & ListView1.SelectedItem.Text & "|" & ListView1.SelectedItem.SubItems(1) & "|" & ListView1.SelectedItem.SubItems(2) & "|"
        End If
        
        
        CadenaDesdeOtroForm = ""
        frmListado5.OpcionListado = 6
        frmListado5.OtrosDatos = CadenaConsulta
         frmListado5.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            If Index = 0 Then
                CadenaConsulta = DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "T") & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T") & ")"
                CadenaConsulta = Data1.Recordset!codcryst & "," & RecuperaValor(CadenaDesdeOtroForm, 1) & "," & CadenaConsulta
                CadenaConsulta = "INSERT INTO scryst2(codcryst,linea,descriprp,nomcryst) VALUES (" & CadenaConsulta
                ejecutar CadenaConsulta, False
                ListView1.ListItems.Clear
                CargaSubRpt
            Else
                CadenaConsulta = "UPDATE scryst2 SET descriprp = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 2), "T")
                CadenaConsulta = CadenaConsulta & ",nomcryst = " & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T")
                CadenaConsulta = CadenaConsulta & " WHERE codcryst = " & Data1.Recordset!codcryst
                CadenaConsulta = CadenaConsulta & " AND linea =" & RecuperaValor(CadenaDesdeOtroForm, 1)
                ejecutar CadenaConsulta, False
                
                ListView1.SelectedItem.SubItems(1) = RecuperaValor(CadenaDesdeOtroForm, 2)
                ListView1.SelectedItem.SubItems(2) = RecuperaValor(CadenaDesdeOtroForm, 3)
            End If
        
        
        End If
    End If
    
End Sub

Private Sub mnModificar_Click()
    If Text1(0).Text = "" Then Exit Sub
     If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
     Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub Text1_LostFocus(Index As Integer)
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    If Index = 4 Then 'cod. ISO
        If Text1(Index).Text = "" Then Exit Sub
        If Not PonerFormatoEntero(Text1(Index)) Then
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
        End If
    End If
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  ' insertar
            Image3_Click (0)
        Case 2 ' modificar
            Image3_Click (1)
        Case 3 ' eliminar
            Image3_Click (2)
        Case Else
    End Select
   
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Anyadir
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        
        Case 5
            'BUSCAR
            BotonBuscar
            
        Case 6
            'Ver todos
            BotonVerTodos
        
    End Select
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        '### A mano
        '-------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(0).Text = ""
                PonerFoco Text1(0)
            End If
    End If
End Sub

Private Sub HacerBusqueda()
Dim cad As String
Dim CadB As String
CadB = ObtenerBusqueda(Me, False)

If chkVistaPrevia = 1 Then
    MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
End If
End Sub


Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub



Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3
    
    Text1(0).Text = SugerirCodigoSiguienteStr("scryst", "codcryst")
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    PonerModo 4
    
    'Bloquear el código que es clave primaria
    BloquearTxt Text1(0), True
    'Si no es root o administradar no Mofificar la descripcion del documento
    If (vUsu.Nivel <> 0 And vUsu.Nivel <> 1) Then BloquearTxt Text1(1), True
    
    PonerFoco Text1(1)
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me, 1)
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    
    Me.ListView1.ListItems.Clear
    
    If Val(Data1.Recordset!LlevaMulitInformes) = 1 Then
        Image3(0).visible = True
        CargaSubRpt
    Else
        Image3(0).visible = False
    End If
    Image3(1).visible = Image3(0).visible
    Image3(2).visible = Image3(0).visible


    PonerModoOpcionesMenu

    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.ListView1.ListItems.Clear
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte
   
    Modo = Kmodo
        
    '----------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
   
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    b = (Kmodo = 2) Or (Kmodo = 0)
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    PonerBotonCabecera Not b
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
   
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    BloquearChecks Me, Modo
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
Dim i As Integer

    b = (Modo = 0) Or (Modo = 2)
    Me.Toolbar1.Buttons(1).Enabled = b 'Insertar
    Me.mnNuevo.Enabled = b
    b = (Modo = 2)
    Me.Toolbar1.Buttons(2).Enabled = b 'Modificar
    Me.mnModificar.Enabled = b
    
    Me.Toolbar1.Buttons(3).Enabled = False 'eliminar
    
    
'    If Val(Data1.Recordset!LlevaMulitInformes) = 1 Then
'        Image3(0).visible = True
'        CargaSubRpt
'    Else
'        Image3(0).visible = False
'    End If
'    Image3(1).visible = Image3(0).visible
'    Image3(2).visible = Image3(0).visible
    
    b = (Modo = 2) And Val(Data1.Recordset!LlevaMulitInformes) = 1
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        ToolAux(i).Buttons(2).Enabled = b
        ToolAux(i).Buttons(3).Enabled = b
    Next i
    
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String

    'Llamamos a al form
    Screen.MousePointer = vbHourglass
'    cad = ParaGrid(Text1(0), 10, "Código")
'    cad = cad & ParaGrid(Text1(1), 30, "Descripción")
'    cad = cad & ParaGrid(Text1(2), 60, "Fichero")
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "scryst"
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|"
'        frmB.vTitulo = "Tipos documento"
'        frmB.vselElem = 1
'        frmB.vConexionGrid = conAri
'        frmB.vCargaFrame = False
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(0)
'        End If
    
    Set frmB = New frmBasico2
    AyudaMantenimientoReports frmB, Text1(0)
    Set frmB = Nothing
    
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaSubRpt()
Dim IT As ListItem

    Set miRsAux = New ADODB.Recordset
    CadenaConsulta = "Select * from scryst2 where codcryst =" & Data1.Recordset!codcryst & " ORDER BY linea"
    miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add
        IT.Text = miRsAux!linea
        IT.SubItems(1) = miRsAux!descriprp
        IT.SubItems(2) = miRsAux!nomcryst
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Set miRsAux = Nothing
    
    
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
