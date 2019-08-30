VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmADVTratamientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tratamientos"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   9585
   Icon            =   "frmADVTratamientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAuxP 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   3840
      MaxLength       =   16
      TabIndex        =   8
      Text            =   "lin"
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAuxP 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   4680
      MaxLength       =   16
      TabIndex        =   9
      Text            =   "plag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAuxP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   2
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   29
      Text            =   "d.plag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdAuxP 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   5640
      TabIndex        =   28
      ToolTipText     =   "Buscar artículo"
      Top             =   4200
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   6120
      MaxLength       =   16
      TabIndex        =   7
      Tag             =   "1|N|S|0||t|c|##,##0.000|S|"
      Text            =   "cantidad"
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Fecha fin|F|S|||advtrata|fechafin|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Nombre|T|S|||advtrata|nomtrata|||"
      Text            =   "Tex"
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   2160
      TabIndex        =   23
      ToolTipText     =   "Buscar artículo"
      Top             =   5280
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   4920
      MaxLength       =   16
      TabIndex        =   6
      Tag             =   "1|N|S|0||t|c|##,##0.000|S|"
      Text            =   "cantidad"
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   16
      Text            =   "nombre artic"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   5
      Text            =   "codartic"
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   7875
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8400
      TabIndex        =   11
      Top             =   7875
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8355
      TabIndex        =   22
      Top             =   7875
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   7710
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   180
         Width           =   2595
      End
   End
   Begin VB.TextBox Text1 
      Height          =   2715
      Index           =   2
      Left            =   240
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   4
      Tag             =   "Observaciones|T|S|||advtrata|observac||N|"
      Text            =   "frmADVTratamientos.frx":000C
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha inicio|F|S|||advtrata|fechaini|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lineas"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Plagas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   19
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7800
      Top             =   480
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   240
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "ID|T|N|||advtrata|codtrata||S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   7920
      Top             =   480
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmADVTratamientos.frx":0012
      Height          =   2760
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4868
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmADVTratamientos.frx":0027
      Height          =   2760
      Left            =   4680
      TabIndex        =   27
      Top             =   1560
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4868
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   690
      Left            =   8160
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1217
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
      Caption         =   ""
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
      Caption         =   "Descripción"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   26
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Artículos"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   3
      Left            =   6960
      Picture         =   "frmADVTratamientos.frx":003C
      ToolTipText     =   "Buscar fecha"
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Fin"
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   24
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   1
      Left            =   5760
      Picture         =   "frmADVTratamientos.frx":00C7
      ToolTipText     =   "Buscar fecha"
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Inicio"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   17
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Id"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   550
      Width           =   1095
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
      Left            =   360
      TabIndex        =   14
      Top             =   7920
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmADVTratamientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As Boolean
Public Event DatoSeleccionado(CadenaSeleccion As String)


'--------------------------------------------------------------------------

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents FrmArt As frmAlmArticu2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmPl As frmAlmPlagas
Attribute frmPl.VB_VarHelpID = -1

Dim NombreTabla As String
Dim NomTablaLineas As String
Dim Ordenacion As String

Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas


Dim CadenaConsulta As String
Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe

Private HaDevueltoDatos As Boolean



Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        Text1(kCampo).BackColor = vbWhite
        cadSeleccion = ""
        HacerBusqueda
        
    Case 3 'INSERTAR
        If DatosOk(True) Then
            If InsertarDesdeForm(Me) Then
                Data1.Refresh
                 PosicionarData True
                 If Not Me.Data1.Recordset.EOF Then
                    BotonLineas False
                    mnNuevo_Click
                 End If
            End If
        End If
    Case 4 'MODIFICAR
        If DatosOk(True) Then
             If ModificaDesdeFormulario(Me, 1) Then
                 TerminaBloquear
                 PosicionarData False
                 
             End If
         End If
            
    Case 5, 6 'articuls y plagas
        If InsertarModificarLinea Then
            'Reestablecemos los campos
            'y ponemos el grid
            If ModificaLineas = 2 Then TerminaBloquear
            If Modo = 5 Then
                DataGrid1.AllowAddNew = False
            Else
                DataGrid2.AllowAddNew = False
            End If
            
            CargaGrids Modo - 4, True
            
            If ModificaLineas = 1 Then 'Insertar
                ModificaLineas = 0
                BotonAnyadirLineas
            ElseIf ModificaLineas = 2 Then 'Modificar
                If Modo = 5 Then
                    Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                Else
                    Data3.Recordset.Find (Data3.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                End If
                ModificaLineas = 0
                PonerBotonCabecera True
                Me.lblIndicador.Caption = ""
                If Modo = 5 Then
                    CargaTxtAux False, False
                    DataGrid1.Enabled = True
                    PonerFocoGrid Me.DataGrid1
                Else
                    CargaTxtPlagas False, False
                    DataGrid2.Enabled = True
                    PonerFocoGrid Me.DataGrid2
                    
                End If
            End If
        End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click()
    Set FrmArt = New frmAlmArticu2
    'frmArt.DatosADevolverBusqueda = "@1@" 'Poner en Modo busqueda
    FrmArt.DesdeTPV = False
    FrmArt.Show vbModal
    Set FrmArt = Nothing
    PonerFoco txtAux(0)
End Sub

Private Sub cmdAuxP_Click()
    CadenaConsulta = ""
    Set frmPl = New frmAlmPlagas
    frmPl.DatosADevolverBusqueda = "0|1|"
    frmPl.Show vbModal
    Set frmPl = Nothing
    If CadenaConsulta <> "" Then
        Me.txtAuxP(1).Text = RecuperaValor(CadenaConsulta, 1)
        Me.txtAuxP(2).Text = RecuperaValor(CadenaConsulta, 2)
        CadenaConsulta = ""
        PonerFoco txtAuxP(1)
    End If
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Mantenimiento Lineas traspasos
            CargaTxtAux False, False
            DataGrid1.Enabled = True
            DataGrid1.AllowAddNew = False
            If Not ModificaLineas = 2 Then 'Modificar
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
             PonerBotonCabecera True
            DataGrid1.Refresh
            PonerFocoBtn Me.cmdRegresar
        Case 6 '
            CargaTxtPlagas False, False
            DataGrid2.Enabled = True
            DataGrid2.AllowAddNew = False
            If Not ModificaLineas = 2 Then 'Modificar
                If Not Data3.Recordset.EOF Then Data3.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            DataGrid2.Refresh
            PonerFocoBtn Me.cmdRegresar
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera

    'Quitar lineas y volver a la cabecera
    If Modo >= 5 Then 'modo 5: Mantenimiento Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        Me.cmdRegresar.visible = Me.DatosADevolverBusqueda
        Me.cmdCancelar.Cancel = True
        Me.cmdRegresar.Caption = "Regresar"
    Else
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        
        RaiseEvent DatoSeleccionado(Data1.Recordset.Fields(0) & "|" & Data1.Recordset.Fields(1) & "|")
        Unload Me
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Form_activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
    btnAnyadir = 5 'Posicion del boton Añadir en la toolbar1
    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(9).Image = 10 'Mantenimiento Líneas
        .Buttons(10).Image = 48 'plagas
        .Buttons(12).Image = 16 'Imprimir
        .Buttons(13).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
       
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)


    cadSeleccion = ""
    NombreTabla = "advtrata"
    NomTablaLineas = "advtrata_lineas" '
    Ordenacion = " ORDER BY codtrata"
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " WHERE codtrata = -1"

    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    

    CargaGrids 0, False '(Modo = 2) 'False

    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

'0.Los 2     1.- Articulos      2.-Plagas
Private Sub CargaGrids(Cual As Byte, Enlazando As Boolean)
    If Cual <> 2 Then CargaGridArticulos Enlazando
    If Cual <> 1 Then CargaGridPlagas Enlazando
End Sub

Private Sub CargaGridArticulos(enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(1, enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, False
      
    DataGrid1.Columns(0).visible = False 'Cod. trasp
    DataGrid1.Columns(1).visible = False 'Numlinea
    
    I = 2
    'Cod. Artículo
    DataGrid1.Columns(I).Caption = "Cod. Articulo"
    DataGrid1.Columns(I).Width = 2000
    
    'Nombre Artículo
    I = I + 1
    DataGrid1.Columns(I).Caption = "Nombre Articulo"
    DataGrid1.Columns(I).Width = 3500
    
    'Cantidad
    I = I + 1
    DataGrid1.Columns(I).Caption = "Dosis hab."
    DataGrid1.Columns(I).Width = 1500
    DataGrid1.Columns(I).Alignment = dbgRight
    DataGrid1.Columns(I).NumberFormat = FormatoImporte & "0"  '3 decimales
    
    'Observaciones
    I = I + 1
    DataGrid1.Columns(I).Caption = "Cantidad"
    DataGrid1.Columns(I).Width = 1400
    DataGrid1.Columns(I).Alignment = dbgRight
    DataGrid1.Columns(I).NumberFormat = FormatoImporte & "0"  '3 decimales
       
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
    Next I
       
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub CargaGridPlagas(enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(2, enlaza)
    CargaGridGnral DataGrid2, Me.Data3, SQL, False
      
    DataGrid2.Columns(0).visible = False 'Cod. trasp
    I = 1
    '
    DataGrid2.Columns(I).Caption = "Orden"
    DataGrid2.Columns(I).Width = 770
    
    I = 2
    'Cod. Artículo
    DataGrid2.Columns(I).Caption = "Codigo"
    DataGrid2.Columns(I).Width = 770
    
    'Nombre Artículo
    I = I + 1
    DataGrid2.Columns(I).Caption = "Nombre"
    DataGrid2.Columns(I).Width = 2500
       
    For I = 0 To DataGrid2.Columns.Count - 1
        DataGrid2.Columns(I).AllowSizing = False
    Next I
       
    DataGrid2.Enabled = b
    DataGrid2.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid2.Tag, Err.Description
End Sub



'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim I As Byte
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = 290
        Next I
        Me.cmdAux.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
            Next I
        End If
        
        If ModificaLineas = 1 Then 'Insertar
            For I = 0 To txtAux.Count - 1
'                If i <> 1 Then txtAux(i).Locked = False
                'LAURA 19/10/2006
                If I <> 1 Then BloquearTxt txtAux(I), False
            Next I
            cmdAux.Enabled = True
        ElseIf ModificaLineas = 2 Then
            'Poner valor a los txtAux
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = DataGrid1.Columns(I + 2).Text
            Next I
            BloquearTxt txtAux(0), True
            cmdAux.Enabled = False
            BloquearTxt txtAux(2), False
            BloquearTxt txtAux(3), False
        End If
        
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        
        
        'Fijamos altura y posición Top
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        Me.cmdAux.Top = alto
        Me.cmdAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        txtAux(0).Left = DataGrid1.Left + 340 'codartic
        txtAux(0).Width = DataGrid1.Columns(2).Width - 200
        cmdAux.Left = txtAux(0).Left + txtAux(0).Width
        txtAux(1).Left = cmdAux.Left + cmdAux.Width + 10 'Nom artic
        txtAux(1).Width = DataGrid1.Columns(3).Width - 25
        For I = 2 To txtAux.Count - 1 'Cantidad y Observacion
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 25
            txtAux(I).Width = DataGrid1.Columns(I + 2).Width - 35
        Next I
    End If

    'Los ponemos Visibles o No
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = visible
    Next I
    cmdAux.visible = visible
End Sub




Private Sub CargaTxtPlagas(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim I As Byte
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAuxP.Count - 1
            txtAuxP(I).Top = 290
        Next I
        Me.cmdAuxP.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid2
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For I = 0 To txtAuxP.Count - 1
                txtAuxP(I).Text = ""
            Next I
        End If
        
        If ModificaLineas = 1 Then 'Insertar
            For I = 0 To txtAuxP.Count - 1
                If I <> 2 Then BloquearTxt txtAuxP(I), False
            Next I
            cmdAuxP.Enabled = True
        ElseIf ModificaLineas = 2 Then
            'Poner valor a los txtAux
            For I = 0 To txtAuxP.Count - 1
                txtAuxP(I).Text = DataGrid2.Columns(I + 1).Text
            Next I
            BloquearTxt txtAuxP(0), True
            
            
        End If
        
        If DataGrid2.Row < 0 Then
            alto = DataGrid2.Top + 220
        Else
            alto = DataGrid2.Top + DataGrid2.RowTop(DataGrid2.Row) + 10
        End If
        
        
        'Fijamos altura y posición Top
        For I = 0 To txtAuxP.Count - 1
            txtAuxP(I).Top = alto
            txtAuxP(I).Height = DataGrid2.RowHeight
        Next I
        Me.cmdAuxP.Top = alto
        Me.cmdAuxP.Height = DataGrid2.RowHeight
        
        'Fijamos anchura y posicion Left
        txtAuxP(0).Left = DataGrid2.Left + 340 'codartic
        txtAuxP(0).Width = DataGrid2.Columns(1).Width - 25
        txtAuxP(1).Left = txtAuxP(0).Left + txtAuxP(0).Width + 25
        txtAuxP(1).Width = DataGrid2.Columns(2).Width - 35
        cmdAuxP.Left = txtAuxP(1).Left + txtAuxP(1).Width - 25
        txtAuxP(2).Left = cmdAuxP.Left + cmdAuxP.Width + 10 'Nom artic
        txtAuxP(2).Width = DataGrid2.Columns(3).Width - 130
        
        
        
    End If

    'Los ponemos Visibles o No
    For I = 0 To txtAuxP.Count - 1
        txtAuxP(I).visible = visible
    Next I
    cmdAuxP.visible = visible
End Sub




''Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'''Almacenes Propios
''Dim indice As Byte
''    indice = CByte(Me.imgBuscar(0).Tag)
''    Text1(indice + 2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
''    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2)
''End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
       
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda

    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Text1(CInt(imgFecha(1).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub




Private Sub frmPl_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(index As Integer)
Dim Aux As String

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
'    imgBuscar(2).Tag = Index
    
    Select Case index
'        Case 2
'            Aux = CadenaConsulta
'            CadenaConsulta = ""
'            Set frmPl = New frmAlmPlagas
'            frmPl.DatosADevolverBusqueda = "0|1|"
'            frmPl.Show vbModal
'            Set frmPl = Nothing
'            If CadenaConsulta <> "" Then
'                Text1(2).Text = RecuperaValor(CadenaConsulta, 1)
'                Text2(0).Text = RecuperaValor(CadenaConsulta, 2)
'            End If
'            CadenaConsulta = Aux
    End Select
    
    PonerFoco Text1(index)
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFecha_Click(index As Integer)
Dim Indice As Byte

   Screen.MousePointer = vbHourglass
   imgFecha(1).Tag = index
   Set frmF = New frmCal
   frmF.Fecha = Now
   

   
   PonerFormatoFecha Text1(index)
   If Text1(index).Text <> "" Then frmF.Fecha = CDate(Text1(index).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(index)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo >= 5 Then   'Eliminar lineas Traspaso Almacenes
        BotonEliminarLinea
    Else 'Eliminar Cabecera Traspaso Almacenes
        BotonEliminar
    End If
End Sub

Private Sub mnModificar_Click()
    If Modo >= 5 Then  'Modificar lineas Traspaso Almacenes
        If BLOQUEADesdeFormulario(Me) Then BotonModificarLinea
    Else 'Modificar Cabecera Traspaso Almacenes
        If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo >= 5 Then  'Añadir lineas Traspaso Almacenes
        BotonAnyadirLineas
    
    Else 'Añadir Cabecera Traspaso Almacenes
        BotonAnyadir
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_GotFocus(index As Integer)
    kCampo = index
    If index <> 5 Then ConseguirFoco Text1(index), Modo
End Sub


Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And index = 5 And Modo = 1 Then
        PonerFocoBtn cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub Text1_LostFocus(index As Integer)
    
    If Not PerderFocoGnral(Text1(index), Modo) Then Exit Sub
    
    Select Case index
        Case 0 'Codigo
            
        Case 1, 3 'Fecha
            If Text1(index).Text <> "" And Modo <> 1 Then PonerFormatoFecha Text1(index)
        Case 2
            'Observaciones
            If Text1(index).Text <> "" Then Text1(index).Text = QuitarCaracterEnter(Text1(index).Text)
    End Select
End Sub


Private Sub txtAux_GotFocus(index As Integer)
    ConseguirFocoLin txtAux(index)
End Sub

Private Sub TxtAux_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If index = 3 And KeyCode = 40 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtAux_KeyPress(index As Integer, KeyAscii As Integer)
    If index = 3 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub txtAux_LostFocus(index As Integer)
Dim devuelve As String
Dim CA As CArticulo
    'Quitar espacios en blanco por los lados
    txtAux(index).Text = Trim(txtAux(index).Text)
    
    Select Case index
        Case 0 'Cod. Articulo
            If txtAux(index).Text = "" Then
                txtAux(index + 1).Text = ""
            ElseIf ModificaLineas = 1 Then 'Insertando linea
                'Comprobamos si ya existe una linea con el artículo, solo si estamos insertando (ModificaLineas=1)
                'conAri: conexion a BD Ariges
                devuelve = "codtrata=" & DBSet(Text1(0).Text, "T") & " AND codartic "
                devuelve = DevuelveDesdeBD(conAri, "codartic", NomTablaLineas, devuelve, txtAux(0).Text, "T")
                If devuelve <> "" Then
                    devuelve = "Ya hay una línea con ese Artículo: " & vbCrLf
                    devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
                    MsgBox devuelve, vbExclamation
                     txtAux(index).Text = ""
                    PonerFoco txtAux(index)
                Else
                    devuelve = ""
                    Set CA = New CArticulo
                    If Not CA.LeerDatos(txtAux(0).Text) Then
                        MsgBox "No existe el articulo", vbExclamation
                        
                    Else
                        If CA.Status > 0 Then
                            CA.MostrarStatusArtic True
                        Else
                            If Not CA.EnInventario(1) Then devuelve = CA.Nombre
                        End If
                    End If
                    Set CA = Nothing
                    txtAux(1).Text = devuelve
                    If txtAux(1).Text = "" Then
                        txtAux(index).Text = ""
                        PonerFoco txtAux(index)
                    End If
                End If
            End If
            
        Case 2, 3 'Cantidad (Comprobamos formato como si fuera un Importe)
            'Formato tipo 1: Decimal(12,2)
            If txtAux(index) <> "" Then PonerFormatoDecimal txtAux(index), 8
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
            
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
        Case 9, 10 'Mantenimiento Lineas
            BotonLineas Button.index = 10
            
        Case 12 'Imprimir
            BotonImprimir
            
        Case 13  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo

    'Modo 2. Hay datos y estamos visualizandolos
    '-------------------------------------------
    b = (Kmodo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
              
    'Como el campo 0 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), (Modo <> 1 And Modo <> 3)
    
    'Modo 1:Busqueda / Modo 3: Insertar / Modo 4: Modificar
    '-------------------------------------------------------
    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 1 To 3
        If I <> 2 Then Me.imgFecha(I).Enabled = b
    Next I
    
    
    
    
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    cmdRegresar.visible = False
    If Modo = 2 Then
        '
        If Me.DatosADevolverBusqueda Then
            If Not Me.Data1.Recordset.EOF Then cmdRegresar.visible = True

        End If
    End If
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    
    
        Toolbar1.Buttons(10).visible = True
  
    
    
         b = (Modo = 2) Or (Modo >= 5)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(7).Enabled = b
        Me.mnEliminar.Enabled = b
        
        '--------------------------------
        b = (Modo = 2)
        'Lineas Traspaso Almacenes
        Toolbar1.Buttons(9).Enabled = b
        'Actualizar
        Toolbar1.Buttons(10).Enabled = b
        'Imprimir
        Toolbar1.Buttons(12).Enabled = b
            
        '-------------------------------
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'VerTodos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b

End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox

    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(index As Integer)
'Botones(Flechas) de Desplazamiento de Registros de la Toolbar

    Select Case Modo
        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
            If Data2.Recordset.EOF Then Exit Sub
            DesplazamientoData Data2, index
        Case Else 'Datos de Cabecera
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, index
            PonerCampos
    End Select
End Sub

'Cual: 1 El articulos    2. Plagas
Private Function MontaSQLCarga(Cual As Byte, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim tabla As String
On Error GoTo EMontaSQL
 
    If Cual = 1 Then
        tabla = NomTablaLineas
    
    
        SQL = "SELECT " & tabla & ".codtrata, "
        SQL = SQL & tabla & ".numlinea, " & tabla & ".codartic, Articulos.nomartic, dosishab," & tabla & ".cantidad "
        SQL = SQL & " FROM ((" & tabla & " LEFT JOIN sartic AS Articulos ON " & tabla & ".codartic ="
        SQL = SQL & " Articulos.codartic))"
    
    Else
        
        tabla = "advtrataPlagas"
        SQL = "select codtrata,numlinea,codplaga,nombrepl  from " & tabla & " INNER JOIN splagas ON "
        SQL = SQL & tabla & ".codplaga = splagas.codigopl"
    
    End If
    If enlaza Then
        SQL = SQL & ObtenerWhereCP(True)  '" WHERE codtrasp = " & Data1.Recordset!codtrasp
    Else
        SQL = SQL & " WHERE codtrata = -1"
    End If
    SQL = SQL & " ORDER BY " & tabla & ".numlinea"
    MontaSQLCarga = SQL
    
EMontaSQL:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrids 0, False

        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
        PonerFoco Text1(0)
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonLineas(Plagas As Boolean)
On Error GoTo ErrorLineas

    Screen.MousePointer = vbHourglass
    If Plagas Then
        NumRegElim = 6
    Else
        NumRegElim = 5
    End If
    PonerModo CByte(NumRegElim)
    ModificaLineas = 0
    PonerBotonCabecera True
    CargaGrids CByte(NumRegElim - 4), True
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lineas"
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonAnyadir()
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
        
    'Ponemos el grid lineas Traspaso enlazando a ningun sitio
    CargaGrids 0, False
    
    

    
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLineas()
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    
    ModificaLineas = 1
    
    vWhere = ObtenerWhereCP(False)
    If Modo = 5 Then
        cmdAceptar.Tag = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    Else
        cmdAceptar.Tag = SugerirCodigoSiguienteStr("advtrataPlagas", "numlinea", vWhere)
    End If
    
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    'Situamos el grid al final
    If Modo = 5 Then
        AnyadirLinea DataGrid1, Data2
    
        CargaTxtAux True, True
        PonerFoco txtAux(0)
    Else
        AnyadirLinea DataGrid2, Data3
        CargaTxtPlagas True, True
        Me.txtAuxP(0).Text = cmdAceptar.Tag
        PonerFoco txtAuxP(0)
        
    End If
    

End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    'Como el campo 0 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), True
    PonerFoco Text1(1)
End Sub




Private Sub BotonModificarLinea()
Dim I As Integer

    If Modo = 5 Then
        If Data2.Recordset.EOF Then Exit Sub
        If Data2.Recordset.RecordCount < 1 Then Exit Sub
    Else
        If Data3.Recordset.EOF Then Exit Sub
        If Data3.Recordset.RecordCount < 1 Then Exit Sub
    End If
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub
    
    ModificaLineas = 2 'Modificar

    Screen.MousePointer = vbHourglass
    
    PonerBotonCabecera False
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If Modo = 5 Then
        If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
            I = DataGrid1.Bookmark - DataGrid1.FirstRow
            DataGrid1.Scroll 0, I
            DataGrid1.Refresh
        End If
        
        cmdAceptar.Tag = Data2.Recordset!numlinea
    
        CargaTxtAux True, False
        PonerFoco txtAux(2) 'Poner el foco
        Me.DataGrid1.Enabled = False
    Else
        If DataGrid2.Bookmark < DataGrid2.FirstRow Or DataGrid2.Bookmark > (DataGrid2.FirstRow + DataGrid2.VisibleRows - 1) Then
            I = DataGrid2.Bookmark - DataGrid2.FirstRow
            DataGrid2.Scroll 0, I
            DataGrid2.Refresh
        End If
        
        cmdAceptar.Tag = Data3.Recordset!numlinea
    
        CargaTxtPlagas True, False
        PonerFoco txtAuxP(1) 'Poner el foco
        Me.DataGrid2.Enabled = False
    
    End If
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    SQL = DevuelveDesdeBD(conAri, "count(*)", "advpartes", "codtrata", CStr(Data1.Recordset!codtrata), "T")
    If SQL = "" Then SQL = "0"
    If Val(SQL) > 0 Then
        MsgBox "Partes relacionados(" & SQL & ") con este tratamiento", vbExclamation
        Exit Sub
    End If
    
    
    
    
    SQL = "Va a eliminar el tratamiento:" & vbCrLf
    SQL = SQL & "------------------------------------------" & vbCrLf & vbCrLf
    SQL = SQL & vbCrLf & "ID   : " & Text1(0).Text
    SQL = SQL & vbCrLf & "Nombre  : " & CStr(Data1.Recordset.Fields(1))
    SQL = SQL & vbCrLf & vbCrLf & " ¿Desea continuar ? "
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
'
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        NumRegElim = Data1.Recordset.Fields(0)
'        vTipoMov.DevolverContador CodTipoMov, NumRegElim
'        Set vTipoMov = Nothing
    
        NumRegElim = Data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            CargaGrids 0, False
            PonerModo 0
        End If
    End If
     
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
    
    conn.BeginTrans
    SQL = ObtenerWhereCP(True)  '" WHERE  codtrasp=" & Data1.Recordset!codtrasp
    
    'Lineas
    conn.Execute "Delete  from " & NomTablaLineas & SQL
    'Lineas tratamientos
    conn.Execute "Delete  from advtrataPlagas" & SQL
    
    'Cabeceras
    conn.Execute "Delete  from " & NombreTabla & SQL
                      

FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub BotonEliminarLinea()
Dim SQL As String
On Error GoTo Error2
    'Ciertas comprobaciones
    If Modo = 6 Then
        If Data3.Recordset.EOF Then Exit Sub
    Else
        If Data2.Recordset.EOF Then Exit Sub
    End If

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    '### a mano
    SQL = "Seguro que desea eliminar la línea @1:"
    SQL = SQL & "del Artículo" & vbCrLf & "Código: @2"
    SQL = SQL & vbCrLf & "Descripción: @3"
    If Modo = 5 Then
        SQL = Replace(SQL, "@1", "del Artículo")
        SQL = Replace(SQL, "@2", Data2.Recordset!codArtic)
        SQL = Replace(SQL, "@3", Data2.Recordset.Fields(3))
    Else
        SQL = Replace(SQL, "@1", "la plaga")
        SQL = Replace(SQL, "@2", Data3.Recordset!codplaga)
        SQL = Replace(SQL, "@3", Data3.Recordset!nombrepl)
    End If
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
    
        'Hay que eliminar
        If Modo = 5 Then
            SQL = NomTablaLineas
            kCampo = Data2.Recordset!numlinea
        Else
            SQL = "advtrataPlagas"
            kCampo = Data3.Recordset!numlinea
        End If
        SQL = "Delete from " & SQL & ObtenerWhereCP(True)
        SQL = SQL & " and numlinea=" & kCampo
        
        conn.Execute SQL
        
        If Modo = 5 Then
            CancelaADODC Me.Data2
            CargaGrids 1, True
            CancelaADODC Me.Data2
        Else
            CancelaADODC Me.Data3
            CargaGrids 2, True
            CancelaADODC Me.Data3
        End If
    End If
    ModificaLineas = 0
Error2:
        Screen.MousePointer = vbDefault
        ModificaLineas = 0
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Línea ", Err.Description
End Sub


Private Function DatosOk(Optional cabecera As Boolean) As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function

   
    
    DatosOk = b
End Function

Private Function DatosOkLineaPlag() As Boolean
    DatosOkLineaPlag = True
    For NumRegElim = 0 To Me.txtAuxP.Count - 2  'El nombre de plaga NO lo miro
        txtAuxP(NumRegElim).Text = Trim(txtAuxP(NumRegElim).Text)
        If txtAuxP(NumRegElim).Text = "" Then
            MsgBox "Campo obligado", vbExclamation
            DatosOkLineaPlag = False
        Else
            If Not IsNumeric(txtAuxP(NumRegElim)) Then
                MsgBox "Campo numerico", vbExclamation
                DatosOkLineaPlag = False
            End If
        End If
        If Not DatosOkLineaPlag Then
            PonerFoco txtAuxP(NumRegElim)
            Exit For
        End If
    Next
    
    
End Function



Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim EsDeDosis As Boolean
Dim cad As String

    DatosOkLinea = False
    b = True
        
    If txtAux(0).Text = "" Then
        MsgBox "El campo Cod. Artículo no puede ser nulo", vbExclamation
        b = False
        PonerFoco txtAux(0)
    Else
        'Veamos si es articulo de dosis
        EsDeDosis = False
        cad = " sunida.codunida=sartic.codunida AND codartic"
        cad = DevuelveDesdeBD(conAri, "estrabajo", "sunida,sartic", cad, txtAux(0).Text, "T")
        EsDeDosis = cad = "1"
    End If
        
    'Comprobamos el campo Cantidad
    If txtAux(2).Text <> "" Then

        If Not IsNumeric(txtAux(2).Text) Then
            MsgBox "El campo dosis debe ser numérico", vbExclamation
            b = False
        End If
        
        If b Then
            If Not EsDeDosis Then
                MsgBox "Articulo no es de dosis", vbExclamation
                b = False
            End If
        End If
    End If
    
    
    
    
    If Not b Then
        PonerFoco txtAux(2)
        Exit Function
    End If
    
    If txtAux(3).Text <> "" Then

        If Not IsNumeric(txtAux(3).Text) Then
            MsgBox "El campo Cantidad debe ser numérico", vbExclamation
            b = False
        End If
    End If
    
    If Not b Then
        PonerFoco txtAux(2)
        Exit Function
    End If
    
    
    If EsDeDosis Then
        If txtAux(2).Text = "" Then
            'Sept 2012.  Martin llama para que dejemos pasar. Avisamos pero dejamos pasar
            If MsgBox("Deberia indicar las dosis. Continuar igualmente?", vbQuestion + vbYesNo) = vbNo Then b = False

        Else
'            If ImporteFormateado(txtAux(2).Text) >= 1000 Then
'                MsgBox "Dosis menor que 1000", vbExclamation
'                b = False
'            End If
        End If
    End If
    
    If b Then
         If Not (txtAux(2).Text = "" Xor txtAux(3).Text = "") Then
            MsgBox "Introduza dosis o cantidad", vbExclamation
            b = False
            
        End If
    End If
    
    'b = ComprobarStock(txtAux(0).Text, txtAux(1).Text, txtAux(2).Text, CodTipoMov)
    'b = ComprobarStock(txtAux(0).Text, Text1(2).Text, txtAux(2).Text, CodTipoMov)
         
    DatosOkLinea = b
End Function


Private Sub PonerBotonCabecera(b As Boolean)
On Error Resume Next
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.cmdRegresar.Cancel = True
        If Modo = 5 Then
            Me.lblIndicador.Caption = "Lin. artículos"
        Else
            Me.lblIndicador.Caption = "Lin. plagas"
        End If
    Else
        Me.lblIndicador.Caption = ""
        Me.cmdCancelar.Cancel = True
    End If
     'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Function InsertarModificarLinea() As Boolean

On Error GoTo EInsertarModificarLinea

    If Modo = 5 Then
        InsertarModificarLinea = InsertarModificarLineaArt
    Else
        InsertarModificarLinea = InsertarModificarLineaPla
    End If
    CadenaConsulta = ""
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Lineas " & vbCrLf & Err.Description
    InsertarModificarLinea = False
End Function

Private Function InsertarModificarLineaArt() As Boolean
Dim SQL As String

    
    
    SQL = ""
    
    'Si no ha puesto dosis, la cantidad PUEDE ser cero
    'Para ello si txtaux(2)[DOSIS] esta vacio pondre que no es NULL cantidad
    txtAux(2).Text = Trim(txtAux(2).Text)
    If txtAux(2).Text = "" Then
        CadenaConsulta = "N"
    Else
        CadenaConsulta = "S"
    End If
    
    InsertarModificarLineaArt = False
    Select Case ModificaLineas
    Case 1 'Insertar
        If DatosOkLinea() Then 'INSERTAR
            SQL = "INSERT INTO advtrata_lineas (codtrata,numlinea,codartic,dosishab,cantidad) "
            SQL = SQL & " VALUES (" & DBSet(Text1(0).Text, "T") & ", "
            SQL = SQL & cmdAceptar.Tag & ", "
            SQL = SQL & DBSet(txtAux(0).Text, "T") & ", "
            SQL = SQL & DBSet(txtAux(2).Text, "N", "S") & ","
            SQL = SQL & DBSet(txtAux(3).Text, "N", CadenaConsulta) & ") "
        Else
'            PonerFoco txtAux(3)
        End If
    Case 2 'Modificar
        If DatosOkLinea() Then
            SQL = "UPDATE advtrata_lineas Set dosishab = " & DBSet(txtAux(2).Text, "N", "S")
            SQL = SQL & ", cantidad = " & DBSet(txtAux(3).Text, "N", CadenaConsulta)
            SQL = SQL & ObtenerWhereCP(True) & " AND " '" WHERE codtrasp =" & Val(Text1(0).Text) & " AND "
            SQL = SQL & " numlinea =" & cmdAceptar.Tag
        End If
    End Select
    
    If ejecutar(SQL, False) Then InsertarModificarLineaArt = True
    

End Function


Private Function InsertarModificarLineaPla() As Boolean
Dim SQL As String

        
   
    
    InsertarModificarLineaPla = False
    Select Case ModificaLineas
    Case 1 'Insertar
        If DatosOkLineaPlag() Then 'INSERTAR
            SQL = "INSERT INTO advtrataPlagas (codtrata,numlinea,codplaga) "
            SQL = SQL & " VALUES (" & DBSet(Text1(0).Text, "T") & ", "
            SQL = SQL & txtAuxP(0).Text & ", "
            SQL = SQL & txtAuxP(1).Text & ")"
        Else
'            PonerFoco txtAux(3)
        End If
    Case 2 'Modificar
        If DatosOkLineaPlag() Then
            SQL = "UPDATE advtrataPlagas Set codplaga = " & DBSet(txtAuxP(1).Text, "N", "N")
            SQL = SQL & ObtenerWhereCP(True) & " AND " '
            SQL = SQL & " numlinea =" & Data3.Recordset!numlinea
        End If
    End Select
    CadenaConsulta = ""
    
    If ejecutar(SQL, False) Then InsertarModificarLineaPla = True

End Function


Private Sub MandaBusquedaPrevia(CadB As String)


    
       

    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vCampos = "Código|advtrata|codtrata|T||30·Denominacion|advtrata|nomtrata|T||70·"
    frmB.vTabla = NombreTabla
    frmB.vSQL = ""
    HaDevueltoDatos = False
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vTitulo = "Tratamientos"
    frmB.vselElem = 0
    frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
    '#
    frmB.Show vbModal
    Set frmB = Nothing

    Screen.MousePointer = vbDefault
End Sub

Private Sub HacerBusqueda()
Dim CadB As String
    
    CadB = ObtenerBusqueda(Me, False)
    cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & Ordenacion
            PonerCadenaBusqueda
        Else
'            CadenaConsulta = "select * from " & NombreTabla & Ordenacion
'            PonerCadenaBusqueda
            MsgBox "Introducir criterios de búsqueda", vbExclamation
            PonerFoco Text1(0)
        End If
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
        Me.DataGrid1.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    If Err.Number <> 0 Then MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
   ' Text2(0).Text = PonerNombreDeCod(Text1(2), conAri, "splagas", "nombrepl")
    CargaGrids 0, True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub






Private Function InsertarCabeceraHistorico() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarCab

    SQL = "SELECT codtrasp,fechatra,almaorig,almadest,codtraba,observa1 from advtrata "
    SQL = SQL & ObtenerWhereCP(True)
    SQL = SQL & " AND fechatra='" & Format(Data1.Recordset!fechatra, "yyyy-mm-dd") & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        SQL = "INSERT INTO schtra (codtrasp, fechatra,hormovim,almaorig,almadest,codtraba,observa1) "
        SQL = SQL & " VALUES (" & RS.Fields(0).Value & ", '" & Format(RS.Fields(1).Value, "yyyy-mm-dd") & "', '"
        SQL = SQL & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & RS.Fields(2).Value & ", " & RS.Fields(3).Value & ", "
        SQL = SQL & RS.Fields(4).Value & ", " & DBSet(RS.Fields(5).Value, "T") & ")"
    End If
    RS.Close
    Set RS = Nothing
    
    conn.Execute SQL
    
EInsertarCab:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabeceraHistorico = False
    Else
        InsertarCabeceraHistorico = True
    End If
End Function


Private Function InsertarLineasHistorico(MenError As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarLineas

    SQL = "SELECT codtrasp, numlinea, codartic, cantidad, observa2 from slitra "
    SQL = SQL & ObtenerWhereCP(True)
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    RS.MoveFirst
    While Not RS.EOF
        SQL = "INSERT INTO slhtra (codtrasp, fechamov, numlinea, codartic, cantidad, observa2)"
        SQL = SQL & " VALUES (" & RS.Fields(0).Value & ", '" & Format(Data1.Recordset!fechatra, FormatoFecha) & "', "
        SQL = SQL & RS.Fields(1).Value & ", " & DBSet(RS.Fields(2).Value, "T") & ", "
        SQL = SQL & DBSet(RS.Fields(3).Value, "N") & ", " & DBSet(RS.Fields(4).Value, "T") & ")"
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
EInsertarLineas:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
        RS.Close
        Set RS = Nothing
        MenError = Err.Number & ": " & Err.Description
    Else
        MenError = ""
        InsertarLineasHistorico = True
    End If
End Function




Private Function BorrarTraspaso(MenError As String) As Boolean
Dim SQL As String

    BorrarTraspaso = False
    
    'Borramos las lineas
    SQL = "Delete from "
    SQL = SQL & "slitra"
    SQL = SQL & " WHERE codtrasp = " & Data1.Recordset!codtrasp
    conn.Execute SQL
    
    'La cabecera
    SQL = "Delete from "
    SQL = SQL & "advtrata"
    SQL = SQL & " WHERE codtrasp =" & Data1.Recordset!codtrasp
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        BorrarTraspaso = False
        MenError = Err.Number & ": " & Err.Description
    Else
        BorrarTraspaso = True
        MenError = ""
    End If
End Function





Private Sub BotonImprimir()

    
  
    '    AbrirListado (7) '7: Informe Traspaso de Almacen
        

End Sub


Private Sub BotonImprimirHco()
Dim indRPT As Byte
Dim cadParam As String
Dim cad As String
Dim numParam As Byte
Dim nomDocu As String

    cadParam = "|"
    numParam = 0
    If Not PonerParamEmpresa(cadParam, numParam) Then Exit Sub
    
    indRPT = 2 '2: Historico Traspaso de Almacen
    If PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        With frmImprimir
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SeleccionaRPTCodigo = pRptvMultiInforme
            .NombreRPT = nomDocu
            .NombrePDF = pPdfRpt
            .EnvioEMail = False
            .Opcion = 7
            .Titulo = "Hist. Traspaso Alm."
            If cadSeleccion <> "" Then
                .FormulaSeleccion = cadSeleccion
            Else
                'Se Llama desde dobleclick en frmAlmMovimArticulos
                'o estamos en Historico
                cad = "{schtra.codtrasp}= " & Data1.Recordset!codtrasp
                cad = cad & " and {schtra.fechatra}= Date(" & Year(Data1.Recordset!fechatra) & "," & Month(Data1.Recordset!fechatra) & "," & Day(Data1.Recordset!fechatra) & ")" & ""
                .FormulaSeleccion = cad
            End If
            .Show vbModal
        End With
    End If
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function ObtenerWhereCP(conWhere As Boolean) As String
On Error Resume Next
    ObtenerWhereCP = " codtrata= " & DBSet(Text1(0).Text, "T")
    If conWhere Then ObtenerWhereCP = " WHERE " & ObtenerWhereCP
    
End Function


Private Sub PosicionarData(VieneDeInsertar As Boolean)
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String
Dim b As Boolean
    
    b = Data1.Recordset.EOF
    If Not b And VieneDeInsertar Then b = True

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             PonerCampos
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    End If
End Sub




Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrids 0, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtAuxP_GotFocus(index As Integer)
    ConseguirFocoLin txtAuxP(index)
End Sub

Private Sub txtAuxP_KeyPress(index As Integer, KeyAscii As Integer)
    If index = 1 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAuxP_LostFocus(index As Integer)

    txtAuxP(index).Text = Trim(txtAuxP(index).Text)
    If txtAuxP(index).Text = "" Then
        If index = 1 Then txtAuxP(2).Text = ""
        Exit Sub
    End If
    
    Select Case index
        Case 0 'Cod. Articulo
            If Not IsNumeric(txtAuxP(index).Text) Then
                MsgBox "Campo numerico", vbExclamation
                txtAuxP(index).Text = ""
                PonerFoco txtAuxP(index)
            End If
        Case 1
            CadenaConsulta = ""
            If Not IsNumeric(txtAuxP(index).Text) Then
                MsgBox "Campo numerico", vbExclamation
            Else
                CadenaConsulta = DevuelveDesdeBD(conAri, "nombrepl", "splagas", "codigopl", txtAuxP(1).Text)
                If CadenaConsulta = "" Then MsgBox "No existe la plaga"
            End If
            Me.txtAuxP(2).Text = CadenaConsulta
            If CadenaConsulta = "" Then
                Me.txtAuxP(1).Text = ""
                PonerFoco txtAuxP(1)
            End If
    End Select
End Sub

