VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacFacAsignar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control albaranes facturados"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17910
   ClipControls    =   0   'False
   Icon            =   "frmFacFacAsignar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   17910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
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
      Left            =   11280
      TabIndex        =   12
      ToolTipText     =   "Buscar envío"
      Top             =   5400
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
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
      Index           =   2
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   6
      Tag             =   "Tipalb|T|S|||scafac1|notasportes|||"
      Text            =   "matr"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3780
      TabIndex        =   24
      Top             =   90
      Width           =   1020
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   25
         Top             =   180
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificación otros datos albarán"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   150
      TabIndex        =   22
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   23
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
      Height          =   315
      Left            =   14310
      TabIndex        =   21
      Top             =   270
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmFacFacAsignar.frx":000C
      Left            =   3600
      List            =   "frmFacFacAsignar.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "Archi|N|N|||scafac1|docarchiv|||"
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Left            =   8520
      MaxLength       =   10
      TabIndex        =   14
      Tag             =   "Alb|N|N|||scafac1|numalbar|0000|S|"
      Text            =   "nunmalb"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   13
      Tag             =   "Tipalb|T|N|||scafac1|codtipoa||S|"
      Text            =   "tipo"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
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
      Left            =   12120
      TabIndex        =   11
      ToolTipText     =   "Buscar fecha"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Left            =   10920
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Fecha|F|S|||scafac1|fecenvio|dd/mm/yyyy||"
      Text            =   "fecha"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux2 
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
      Index           =   1
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Index           =   4
      Left            =   7620
      MaxLength       =   16
      TabIndex        =   5
      Tag             =   "Envio|N|N|||scafac1|codenvio|0||"
      Text            =   "envio"
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
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
      Left            =   8640
      TabIndex        =   27
      ToolTipText     =   "Buscar envío"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6480
      MaxLength       =   16
      TabIndex        =   4
      Tag             =   "Fecha|F|S|||scafac1|fechaalb|dd/mm/yy||"
      Text            =   "Fecha alb"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||scafac1|fecfactu|dd/mm/yy|S|"
      Text            =   "fecha"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "numalbar|N|N|||scafac1|numfactu|0000000|S|"
      Text            =   "numlote"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   150
      TabIndex        =   18
      Top             =   9525
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
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
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "Buscar fecha"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
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
      Index           =   0
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Left            =   480
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "codtipom|T|N|||scafac1|codtipom||S|"
      Text            =   "codartic codarti"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
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
      Left            =   13575
      TabIndex        =   9
      Top             =   9690
      Width           =   1065
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
      Left            =   14775
      TabIndex        =   10
      Top             =   9690
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5040
      Top             =   5280
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacFacAsignar.frx":0022
      Height          =   8520
      Left            =   150
      TabIndex        =   15
      Top             =   885
      Width           =   17625
      _ExtentX        =   31089
      _ExtentY        =   15028
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
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
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
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
      Left            =   240
      TabIndex        =   16
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
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
      Begin VB.Menu mnModAdv 
         Caption         =   "Mo&dificar + datos"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacFacAsignar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmB1 As frmFacFormasEnvio 'formas de envio
Attribute frmB1.VB_VarHelpID = -1
Private WithEvents frmFl As frmFlotas
Attribute frmFl.VB_VarHelpID = -1



Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer

Dim EsBusqueda As Boolean

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
'            If DatosOk Then
'                If InsertarDesdeForm(Me) Then
'                    EsBusqueda = True
'                    CadenaBusqueda = " WHERE  slotes.codartic=" & DBSet(txtAux(0).Text, "T") & " AND numlotes=" & DBSet(txtAux(1).Text, "T")
'                    CargaGrid True
'                    BotonAnyadir
'                End If
'            End If
        
        Case 4 'MODIFICAR
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                 If ModificaDesdeFormulario(Me, 3) Then
                     TerminaBloquear
                     NumReg = Data1.Recordset.AbsolutePosition
                     PonerModo 2
                     CancelaADODC Me.Data1
                     CargaGrid True
                     LLamaLineas 10
                     SituarDataPosicion Data1, NumReg, Indicador
                 End If
                 lblIndicador.Caption = Indicador
                 PonerFocoGrid DataGrid1
             End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    HaDevueltoDatos = ""
    Select Case Index
        Case 1
'            MandaBusquedaPrevia2 'Index = 1
            Set frmB1 = New frmFacFormasEnvio
            frmB1.DeConsulta = True
            frmB1.DatosADevolverBusqueda = "0|"
            frmB1.Show vbModal
            Set frmB1 = Nothing
            
        Case 0, 3 'fecha entrada
            Set frmF = New frmCal
            frmF.Fecha = Now
            If txtAux(Index + 2).Text <> "" Then frmF.Fecha = CDate(txtAux(Index + 2).Text)
            Screen.MousePointer = vbDefault
            frmF.Show vbModal
            Set frmF = Nothing
            If HaDevueltoDatos <> "" Then
                txtAux(Index + 2).Text = HaDevueltoDatos
                PonerFoco txtAux(Index + 2)
            End If
            
            
        Case 2
            'Flotas
            Set frmFl = New frmFlotas
            frmFl.DatosADevolverBusqueda = "0|1|"
            frmFl.Show vbModal
            Set frmFl = Nothing
            If HaDevueltoDatos <> "" Then
                txtAux(8).Text = HaDevueltoDatos
                PonerFoco txtAux(Index + 2)
            End If
    End Select
End Sub


Private Sub cmdCancelar_Click()
Dim Indicador As String

    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
            EsBusqueda = False
           
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            DataGrid1.Enabled = True
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                PonerModo 2
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else
                PonerModo 0
            End If
            LLamaLineas 10
            
        Case 4  'Modificar
            TerminaBloquear
            NumRegElim = Data1.Recordset.AbsolutePosition
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            PonerModo 2
            LLamaLineas 10
            DataGrid1.Enabled = True
            SituarDataPosicion Data1, NumRegElim, Indicador
            DeseleccionaGrid Me.DataGrid1
            lblIndicador.Caption = Indicador
            PonerFocoGrid DataGrid1
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub










Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'ICONOS de La toolbar
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'Añadir
'        .Buttons(6).Image = 4 'Modificar
'        .Buttons(7).Image = 13 'Modificar + datos
'        .Buttons(10).Image = 16 'Imprimir
'        .Buttons(11).Image = 15 'Salir
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
        .Buttons(1).Image = 36 ' modificar más datos
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
   
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    

    Ordenacion = " ORDER BY 1,2,3 "
    CadenaConsulta = MontaSQLCarga(False)
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
   
        PonerModo 0
 
'    CargaGrid (Modo = 2 Or Modo = 0)
    CargaGrid False
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String
    
    On Error GoTo ECarga
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, False
    
    tots = "S|txtAux(0)|T|Tipo|500|;S|txtAux(1)|T|Factura|1050|;S|txtAux(2)|T|Fecha|1100|;S|cmdAux(0)|B||0|;"
    tots = tots & "S|txtAux2(0)|T|Nombre cliente|3850|;S|txtAux(6)|T|Tipo|650|;S|txtAux(7)|T|Albarán|1050|"
    tots = tots & ";S|txtAux(3)|T|Fec.Albarán|1100|;S|txtAux(4)|T|Código|700|;S|cmdAux(1)|B||0|;"
    tots = tots & "S|txtAux2(1)|T|Envio|2160|;"
    'tots = tots & "S|cmdAux(3)|B||0|;S|txtAux(5)|T|Fec.Envio|1400|;S|Combo1|C|Doc.|600|;"
    tots = tots & "S|cmdAux(2)|B||0|;S|txtAux(8)|T|Cod.|900|;S|txtAux2(2)|T|Vehiculo|1900|;"
    tots = tots & "S|cmdAux(3)|B||0|;S|txtAux(5)|T|Fec.Envio|1400|;S|Combo1|C|Doc.|600|;"
    
    arregla tots, DataGrid1, Me, 350

'    'dtos alineados a la dcha
'    DataGrid1.Columns(6).Alignment = dbgCenter

    DataGrid1.ScrollBars = dbgVertical
    
    
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim B As Boolean

    DeseleccionaGrid Me.DataGrid1
    B = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To txtAux.Count - 1
        If jj < 3 Then
            txtAux2(jj).Height = Me.DataGrid1.RowHeight
            txtAux2(jj).Top = alto
            txtAux2(jj).visible = B
        End If
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = B
    Next jj

    Me.Combo1.visible = B
    Me.Combo1.Top = alto
    
    
    For jj = 0 To Me.cmdAux.Count - 1
        
            Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
            Me.cmdAux(jj).Top = alto
            Me.cmdAux(jj).visible = B

    Next jj
End Sub



'Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
''Mantenimiento de Articulos
'    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
'    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    HaDevueltoDatos = CadenaDevuelta
End Sub

Private Sub frmB1_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtAux(4).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    HaDevueltoDatos = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFl_DatoSeleccionado(CadenaSeleccion As String)
    
    HaDevueltoDatos = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub



Private Sub mnModAdv_Click()
    If Modo <> 2 Then Exit Sub
    
    If vUsu.Nivel > 1 Then
        MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    HaDevueltoDatos = ""
    HaDevueltoDatos = HaDevueltoDatos & "scafac.numfactu=" & DBSet(Data1.Recordset!Numfactu, "N") & " AND scafac.codtipom =" & DBSet(Data1.Recordset!codtipom, "T")
    HaDevueltoDatos = HaDevueltoDatos & " AND scafac.fecfactu=" & DBSet(Data1.Recordset!FecFactu, "F")
    HaDevueltoDatos = HaDevueltoDatos & " AND  scafac1.numalbar=" & DBSet(Data1.Recordset!Numalbar, "N") & " AND scafac1.codtipoa =" & DBSet(Data1.Recordset!Codtipoa, "T")
    frmFacModAlbFac.Where = HaDevueltoDatos
    frmFacModAlbFac.Show vbModal
    
    HaDevueltoDatos = ""
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub



Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 5 'Busqueda
            mnBuscar_Click
        Case 6 'Ver Todos
            mnVerTodos_Click
        Case 1 'Nuevo
            'mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3 ' eliminar
        
            
        Case 8 'Imprimir
            If Modo = 2 Or Modo = 0 Then
                frmListado2.Opcion = 43
                frmListado2.Show vbModal
            End If
    End Select
End Sub




Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)

                      
    If Kmodo = 1 Then 'Modo Buscar
        PonerFoco txtAux(0)
    End If
                                 
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    BloquearTxt txtAux(2), (Modo = 4)
    BloquearTxt txtAux(6), (Modo = 4)
    BloquearTxt txtAux(7), (Modo = 4)
    BloquearTxt txtAux(3), (Modo = 4)   'todos bloqueados al modificad
    
    
    Me.cmdAux(0).Enabled = (Modo <> 4)
                   
    '-----------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu  'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
      PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

    'Insertar
    Toolbar1.Buttons(1).Enabled = False
    Me.mnNuevo.Enabled = False

    
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    
    'Eliminar
    Toolbar1.Buttons(3).Enabled = False
    
    'Modificar avanzada
    Toolbar5.Buttons(1).Enabled = B
    Me.mnModAdv.Enabled = B
    
    B = (Modo >= 3 Or Modo = 1)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(6).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B

End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
    
    SQL = "select scafac.codtipom,scafac.numfactu,scafac.fecfactu,nomclien,codtipoa,numalbar,fechaalb,"
    SQL = SQL & "scafac1.codenvio,nomenvio,scafac1.notasportes,nomflota,fecenvio"
    SQL = SQL & ",if(docarchiv=1,""SI"","""") docarchiv "
    SQL = SQL & " FROM scafac inner join scafac1"
    SQL = SQL & " on scafac.codtipom = scafac1.codtipom And scafac.NumFactu = scafac1.NumFactu And scafac.FecFactu = scafac1.FecFactu"
    SQL = SQL & " inner join senvio  on scafac1.codenvio=senvio.codenvio"
    SQL = SQL & " left join sflotas on scafac1.notasportes =sflotas.codflota"
    SQL = SQL & " where true "
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
    Else
        SQL = SQL & " AND  false"
    End If
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        anc = ObtenerAlto(Me.DataGrid1, 30)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            txtAux(kCampo).Text = ""
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    On Error Resume Next

    EsBusqueda = False
    LimpiarCampos
    
    CadenaConsulta = MontaSQLCarga(True)
    PonerCadenaBusqueda
    PonerFocoGrid DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    
 
    For i = 0 To 2
        txtAux(i).Text = DBLet(DataGrid1.Columns(i).Value, "T")
    Next i

    txtAux2(0).Text = DBLet(DataGrid1.Columns(3).Value, "T")
    
    txtAux(6).Text = DBLet(Me.DataGrid1.Columns(4).Value, "T")
    txtAux(7).Text = DBLet(Me.DataGrid1.Columns(5).Value, "T")
    txtAux(3).Text = DBLet(Me.DataGrid1.Columns(6).Value, "T")
    txtAux(4).Text = DBLet(Me.DataGrid1.Columns(7).Value, "T")
    txtAux2(1).Text = DBLet(DataGrid1.Columns(8).Value, "T")
    
    txtAux(5).Text = DBLet(Data1.Recordset!FecEnvio, "F")
    txtAux(8).Text = DBLet(Data1.Recordset!notasportes, "T")
    
    txtAux2(2).Text = DBLet(Data1.Recordset!nomflota, "T")
        
    If UCase(DBLet(DataGrid1.Columns(12).Value, "T")) = "SI" Then
        Combo1.ListIndex = 0
    Else
        Combo1.ListIndex = 1
    End If
    

    
    DataGrid1.Enabled = False
    PonerFoco txtAux(4)
End Sub





Private Function DatosOk() As Boolean
Dim B As Boolean


    On Error GoTo ErrDatosOK

    DatosOk = False
    B = CompForm(Me, 3)
    If Not B Then Exit Function
    
    If txtAux(5).Text <> "" Then
        'Fecha envio no puede ser anterior a fecha albaran
        If CDate(txtAux(5).Text) < CDate(txtAux(3).Text) Then
            MsgBox "Fecha envio no puede ser anterior a fecha del albarán", vbExclamation
            PonerFoco txtAux(5)
            B = False
        End If
    End If

    If txtAux(8).Text <> "" Then
        HaDevueltoDatos = DevuelveDesdeBD(conAri, "nomflota", "sflotas", "codflota", txtAux(8).Text, "T")
        If HaDevueltoDatos = "" Then
            MsgBox "Error vehiculo. " & txtAux(6).Text, vbExclamation
            PonerFoco txtAux(8)
            B = False
        End If
        HaDevueltoDatos = ""
    End If

    
    DatosOk = B
    Exit Function
    
ErrDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos OK.", Err.Description
End Function



Private Sub MandaBusquedaPrevia2()
'Private Sub MandaBusquedaPrevia2(Envio As Boolean)
''Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
'Dim Tabla As String
'Dim Titulo As String
'
'    'Llamamos a al form
'    cad = ""
'    'Estamos en Modo de Cabeceras
'    'Registro de la tabla de cabeceras: slista
        'Cod Diag.|tabla|columna|tipo|formato|10·
        'If Envio Then
            Cad = "Codigo|senvio|codenvio|N||20·"
            Cad = Cad & "Decripcion|senvio|nomenvio|T||60·"
        'Else
        '    cad = "Codigo|szonas|codzonas|N||20·"
        '    cad = cad & "Decripcion|szonas|nomzonas|T||60·"
        'End If
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        
        'frmB.vTabla = tabla
        frmB.vSQL = ""
        
        
        '###A mano
        frmB.vDevuelve = "0|1|"
        'If Envio Then
            frmB.vTitulo = "Forma de envio"
            frmB.vTabla = "senvio"
        'Else
        '    frmB.vTitulo = "ZONAS"
        '    frmB.vTabla = "szonas"
        'End If
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri       'Conexión a BD: Ariges
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos <> "" Then

                txtAux(4).Text = RecuperaValor(HaDevueltoDatos, 1)
                txtAux2(1).Text = RecuperaValor(HaDevueltoDatos, 2)
        End If
    

End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    
    If cadB = "" Then Exit Sub
    
 '   If chkVistaPrevia = 1 Then
 '       MandaBusquedaPrevia cadB
 '   ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaBusqueda = " AND " & cadB
        CadenaConsulta = MontaSQLCarga(True)
        'CadenaBusqueda = " AND " & cadB
        PonerCadenaBusqueda
 '   End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        Exit Sub
    Else
        PonerModo 2
        PonerCampos
    End If
    LLamaLineas 10
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Mod avanzada
            mnModAdv_Click
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If Index > 0 Then PonerFoco txtAux(Index - 1)
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
                If Index = 3 Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    SendKeys "{tab}"
                End If
    End Select
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'   KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 0 'fecha
            Case 4: KEYBusqueda KeyAscii, 1 'envio
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Cad As String

    On Error Resume Next
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 4
            Cad = ""
            If txtAux(Index).Text <> "" Then
                If Not IsNumeric(txtAux(Index).Text) Then
                    MsgBox "Campo numerico", vbExclamation
                    txtAux(Index).Text = ""
                Else
                    If Index = 4 Then
                        Cad = DevuelveDesdeBD(conAri, "nomenvio", "senvio", "codenvio", txtAux(Index).Text)
                  '  Else
                  '      cad = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", txtAux(Index).Text)
                    End If
                    If Cad = "" Then MsgBox "No existe el valor en la BD: " & txtAux(Index).Text, vbExclamation
                End If
                If Cad = "" And txtAux(Index).Text <> "" Then
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
                      
            End If
            txtAux2(1).Text = Cad
        Case 2, 3, 5 'fecha
              PonerFormatoFecha txtAux(Index)
              
            
    Case 8
            Cad = ""
            If txtAux(Index).Text <> "" Then
                Cad = DevuelveDesdeBD(conAri, "nomflota", "sflotas", "codflota", txtAux(Index).Text, "T")
                If Cad = "" Then MsgBox "No existe el valor en la BD: " & txtAux(Index).Text, vbExclamation
                
                If Cad = "" And txtAux(Index).Text <> "" Then
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
                
            End If
            txtAux2(2).Text = Cad
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub

