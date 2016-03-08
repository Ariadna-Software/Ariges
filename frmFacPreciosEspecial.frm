VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacPreciosEspecial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Precios Especiales"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8520
   ClipControls    =   0   'False
   Icon            =   "frmFacPreciosEspecial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1000
      Left            =   480
      TabIndex        =   28
      Top             =   410
      Width           =   7695
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Cod. Cliente|N|N|0|999999|sprees|codclien|000000|S|"
         Text            =   "Text1"
         Top             =   200
         Width           =   750
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   16
         TabIndex        =   1
         Tag             =   "Cod. Artículo|T|N|||sprees|codartic||S|"
         Text            =   "Text1"
         Top             =   580
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   2800
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   580
         Width           =   4605
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   200
         Width           =   3645
      End
      Begin VB.CheckBox chkPermiteDto 
         Caption         =   "Permite Descuento"
         Height          =   195
         Left            =   5880
         TabIndex        =   29
         Tag             =   "Permite Descuento|N|N|||sprees|dtopermi||N|"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   200
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Artículo"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   580
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   900
         ToolTipText     =   "Buscar cliente"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   900
         ToolTipText     =   "Buscar artículo"
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.Frame FrameNuevos 
      Caption         =   "Valores Nuevos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1815
      Left            =   3600
      TabIndex        =   23
      Top             =   1440
      Width           =   4575
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   2910
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Fecha Cambio|F|S|||sprees|fechanue|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   1860
         MaxLength       =   5
         TabIndex        =   8
         Tag             =   "Descuento Especial Nuevo|N|S|0|99.90|sprees|dtoespe1|#0.00|N|"
         Text            =   "Text1"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   7
         Tag             =   "Precio Caja Nuevo|N|S|0|999999.0000|sprees|precion1|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   6
         Tag             =   "Precio Nuevo|N|S|0|999999.0000|sprees|precionu|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Cambio"
         Height          =   255
         Left            =   2910
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   4005
         Picture         =   "frmFacPreciosEspecial.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Dto. Especial"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Precio Caja"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Precio"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FrameActuales 
      Caption         =   "Valores Actuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1815
      Left            =   480
      TabIndex        =   19
      Top             =   1440
      Width           =   2895
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   1875
         MaxLength       =   5
         TabIndex        =   5
         Tag             =   "Comision|N|S|0|99.90|sprees|comision|#0.00|N|"
         Text            =   "Text1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   1875
         MaxLength       =   5
         TabIndex        =   4
         Tag             =   "Descuento Especial|N|S|0|99.90|sprees|dtoespec|#0.00|N|"
         Text            =   "Text1"
         Top             =   1040
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1275
         MaxLength       =   12
         TabIndex        =   3
         Tag             =   "Precio Caja Actual|N|S|0|999999.0000|sprees|precioa1|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   640
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1275
         MaxLength       =   12
         TabIndex        =   2
         Tag             =   "Precio Actual|N|N|0|999999.0000|sprees|precioac|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Comisión"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Dto. Especial"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1040
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Caja"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   640
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Precio"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   6210
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7155
      TabIndex        =   11
      Top             =   6210
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7155
      TabIndex        =   12
      Top             =   6210
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   6045
      Width           =   2655
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
         TabIndex        =   18
         Top             =   180
         Width           =   2115
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6480
         TabIndex        =   16
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacPreciosEspecial.frx":0097
      Height          =   2535
      Left            =   480
      TabIndex        =   13
      Top             =   3480
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3240
      Top             =   6240
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   4560
      Top             =   6240
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
      TabIndex        =   15
      Top             =   8220
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
Attribute VB_Name = "frmFacPreciosEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public CadenaSituarData As String

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes 'Form Mantenimiento Clientes
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticu2  'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String
Dim PrimeraVez As Boolean
Private HaDevueltoDatos As Boolean


Private Sub chkPermiteDto_Click()
    If Modo = 3 Or Modo = 4 Then 'Insertar o Modificar
        If Me.chkPermiteDto.Value = 1 Then
            Me.Text1(4).Text = ""
            BloquearTxt Text1(4), True
        Else
            BloquearTxt Text1(4), False
        End If
    End If
End Sub

Private Sub chkPermiteDto_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkPermiteDto_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkPermiteDto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PonerFoco Text1(2) 'ENTER
End Sub


Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        HacerBusqueda
    Case 3 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then PosicionarData
                'Ponemos la cadena consulta
        End If
        
    Case 4 'MODIFICAR
        If DatosOk Then
             If ModificaDesdeFormulario(Me, 1) Then
                 TerminaBloquear
                 PosicionarData
             End If
         End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

'Private Sub cmdRegresar_Click()
''Este es el boton Cabecera
'Dim cad As String
'Dim Indicador As String
'
'    'Quitar lineas y volver a la cabecera
'    If Modo = 5 Then 'modo 5: Lineas Articulos x Almacen
'        DataGrid1.ClearFields
'        cad = "(codmovim=" & Val(Text1(0).Text) & ")"
'        If SituarData(Data1, cad, Indicador) Then
'            PonerModo 2
'            lblIndicador.Caption = Indicador
'            Me.Toolbar1.Buttons(9).Enabled = True
'            Me.Toolbar1.Buttons(10).Enabled = True
'        End If
'    End If
'End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If CadenaSituarData <> "" Then
        CadenaSituarData = ""
        PonerModo 2
        PonerCampos
        
    End If
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Next kCampo
    
    
    'La toolbar
    btnPrimero = 14 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 15 'Salir
        .Buttons(14).Image = 6 'Primero
        .Buttons(15).Image = 7 'Anterior
        .Buttons(16).Image = 8 'Siguiente
        .Buttons(17).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    PrimeraVez = True
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "sprees" 'Tabla Precios Especiales de Articulos
    Ordenacion = " ORDER BY codclien, codartic"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE "
    
    If CadenaSituarData = "" Then
        CadenaConsulta = CadenaConsulta & " codartic = -1" 'No recupera datos
    Else
        CadenaConsulta = CadenaConsulta & " codartic=" & RecuperaValor(CadenaSituarData, 1) & " AND codclien = " & RecuperaValor(CadenaSituarData, 2)
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If CadenaSituarData = "" Then
        PonerModo 0
        CargaGrid (Modo = 2)
    Else
       ' PonerModo 2
       ' PonerCampos
    End If
    
    'Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim Inicio As Byte
Dim SQL As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
    
    DataGrid1.Columns(0).visible = False 'Cod. Cliente
    DataGrid1.Columns(1).visible = False 'Cod. Articulo
    Inicio = 2
    
    'Numero Linea
    DataGrid1.Columns(Inicio).Caption = "Num. Linea"
    DataGrid1.Columns(Inicio).Width = 1100
    
    'Fecha Cambio
    DataGrid1.Columns(Inicio + 1).Caption = "Fecha Cambio"
    DataGrid1.Columns(Inicio + 1).Width = 1300
    
    'Precio Unidad
    DataGrid1.Columns(Inicio + 2).Caption = "Precio Unidad"
    DataGrid1.Columns(Inicio + 2).Width = 1600
    DataGrid1.Columns(Inicio + 2).Alignment = dbgRight
    DataGrid1.Columns(Inicio + 2).NumberFormat = FormatoPrecio
    
    'Precio Caja
    DataGrid1.Columns(Inicio + 3).Caption = "Precio Caja"
    DataGrid1.Columns(Inicio + 3).Width = 1600
    DataGrid1.Columns(Inicio + 3).Alignment = dbgRight
    DataGrid1.Columns(Inicio + 3).NumberFormat = FormatoPrecio
    
    
    'Descuento Especial
    DataGrid1.Columns(Inicio + 4).Caption = "Dto. Especial"
    DataGrid1.Columns(Inicio + 4).Width = 1400
    DataGrid1.Columns(Inicio + 4).Alignment = dbgRight
    DataGrid1.Columns(Inicio + 4).NumberFormat = FormatoDescuento
    
    DataGrid1.ScrollBars = dbgAutomatic
    
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
    Next i
    DataGrid1.Enabled = b
    PrimeraVez = False
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Articulos
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        'Estamos en Cabecera
        'Recupera todo el registro de Tarifas de Precios
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim Indice As Byte
    Indice = 8
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Clientes
    Text1(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0  'Cod. Cliente
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
        Case 1 'Codigo Articulo
            Set frmA = New frmAlmArticu2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Abre en modo busqueda
            frmA.DesdeTPV = False
            frmA.Show vbModal
            Set frmA = Nothing
    End Select
    PonerFoco Text1(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass

   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Indice = 8
   
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
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
     BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Dim cadkey As Integer
    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo, cadkey
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
        KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim campo As String
Dim tabla As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    Select Case Index
        Case 0 'Codigo Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 1 'Codigo Articulo
            campo = "nomartic"
            tabla = "sartic"
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, tabla, campo)
        
        Case 2, 3, 5, 6 'Precios Actuales y Nuevos
            'Formato tipo 2: Decimal(10,4)
            If Text1(Index).Text <> "" And Modo <> 1 Then PonerFormatoDecimal Text1(Index), 2
        
        Case 4, 7, 9
            'Formato tipo 4: Decimal(4,2)
            If Text1(Index).Text <> "" And Modo <> 1 Then PonerFormatoDecimal Text1(Index), 4
        
        Case 8 'Fecha Cambio
            If Modo <> 1 And Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            PonerFocoBtn Me.cmdAceptar
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
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
        Case 10 'Imprimir
                AbrirListado (30) '30: Informe Precios Especiales de Articulos
        Case 11  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
    If KeyAscii = 27 And Modo = 1 Then cmdCancelar_Click 'busqueda
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    
    '===========================================
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
          
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
     
    BloquearChecks Me, Modo
    
           
    '==============================
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b And Modo <> 4 'Si modificar no activado pq son claves ajenas
    Next i
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = b
    Next i
    
    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    '===============================
    b = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkPermiteDto.Value = 0
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
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
Dim tabla As String
    
    tabla = "spree1"
    SQL = "SELECT * FROM " & tabla
    If enlaza Then
        SQL = SQL & " WHERE codclien=" & Data1.Recordset!codclien & " AND codartic=" & DBSet(Data1.Recordset!codArtic, "T")
    Else
        SQL = SQL & " WHERE codartic = -1"
    End If
    SQL = SQL & " ORDER BY " & tabla & ".numlinea desc"
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

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
'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Vacía los TextBox
    
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    
    'Para que si no se ha cargado el Data1 inicialmente, tenga valor cuando situamos el Data
'    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
'    Data1.RecordSource = CadenaConsulta
           
    'Ponemos el grid de lineas enlazando a ningun sitio
    CargaGrid False
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    PonerFoco Text1(2)
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    SQL = "Precios Especiales." & vbCrLf
    SQL = SQL & "--------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Eliminar El Precio Especial:"
    SQL = SQL & vbCrLf & "Cod. Clien. : " & Text1(0).Text
    SQL = SQL & vbCrLf & "Cod. Artic. : " & Text1(1).Text
    
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar ? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        'DataGrid1.Enabled = False
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
    
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            'MsgBox Err.Number & " : " & Err.Description, vbExclamation
            MuestraError Err.Number, "Eliminar Precio Especial", Err.Description
            Data1.Recordset.CancelUpdate
        End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        conn.BeginTrans
        SQL = " WHERE codclien=" & Val(Data1.Recordset!codclien)
        SQL = SQL & " AND codartic=" & DBSet(Data1.Recordset!codArtic, "T")
        
        'Lineas
        conn.Execute "Delete  from spree1 " & SQL
        
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


Private Function DatosOk() As Boolean
Dim b As Boolean
On Error Resume Next

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    'Comprobar que si hay valores nuevos, la fecha de cambio no es nulo
    If Not EsVacio(Text1(5)) Or Not EsVacio(Text1(6)) Or Not EsVacio(Text1(7)) Then
        b = (Not EsVacio(Text1(8)))
    End If
    If Not b Then
        MsgBox "La Fecha de Cambio debe tener valor.", vbInformation
        Exit Function
    End If
    
    'Comprobar que si no hay valores nuevos no haya fecha de Cambio
    If EsVacio(Text1(5)) And EsVacio(Text1(6)) And EsVacio(Text1(7)) Then
        b = (EsVacio(Text1(8)))
    End If
    If Not b Then
        MsgBox "No hay valores para la fecha de cambio", vbInformation
        Exit Function
    End If
    
    DatosOk = True
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    cad = cad & ParaGrid(Text1(0), 10, "Cliente")
    cad = cad & "Nombre Cliente|sclien|nomclien|T||36·"
    cad = cad & ParaGrid(Text1(1), 15, "Cod. Artic")
    cad = cad & "Desc. Artic|sartic|nomartic|T||38·"
    
    tabla = "(" & NombreTabla & " LEFT JOIN sclien ON " & NombreTabla & ".codclien=sclien.codclien" & ")"
    tabla = tabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic"
    
    Titulo = "Precios Especiales"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|2|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
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
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    'Poner el nombre del cod. cliente
    Text2(0).Text = PonerNombreDeCod(Text1(0), 1, "sclien", "nomclien")
    'Poner el nombre del cod. Articulo
    Text2(1).Text = PonerNombreDeCod(Text1(1), 1, "sartic", "nomartic")
    
    'Si los campos de precios nuevos son cero mostrar cadena vacia
    If Text1(5).Text <> "" Then
        If Text1(5).Text = 0 Then Text1(5).Text = ""
    End If
    If Text1(6).Text <> "" Then
        If Text1(6).Text = 0 Then Text1(6).Text = ""
    End If
    
    BloquearChecks Me, Modo
    
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub BotonActualizar()
'Actualizar Precios Especiales
Dim SQL As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún Precio Especial para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
   
    SQL = "Actualización Precios Especiales de Artículos." & vbCrLf
    SQL = SQL & "---------------------------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Actualizar el Precio Especial para:"
    SQL = SQL & vbCrLf & " Cod. Clien. :  " & CStr(Format(Data1.Recordset.Fields(0), "000000"))
    SQL = SQL & vbCrLf & " Cod. Artic. :  " & Data1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & vbCrLf & " ¿Desea continuar ? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then
        Exit Sub
    End If
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    If ActualizarPreEspecial Then
        SituarDataTrasEliminar Data1, NumRegElim
    End If
End Sub


Private Function ActualizarPreEspecial() As Boolean
'Actualiza los Precios Especiales insertando los precios actuales con la fecha de cambio en el hostórico
' y modificando el la tabla de precios especiales pasando los valores nuevos a ser los actuales.
Dim Donde As String
Dim SQL As String
Dim bol As Boolean
On Error GoTo EActualizarPreEspecial
    
   
    'Aqui empieza transaccion
    conn.BeginTrans
    bol = ActualizarElPrecio(Donde)

EActualizarPreEspecial:
        If Err.Number <> 0 Then
            SQL = "Actualizar Precio Especial." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
'            If OpcionActualizar = 1 Then
                MuestraError Err.Number, SQL, Err.Description
'            Else
'                SQL = Donde & " -> " & Err.Description
'                SQL = Mid(SQL, 1, 200)
'                InsertaError SQL
'            End If
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            ActualizarPreEspecial = True
        Else
            conn.RollbackTrans
            ActualizarPreEspecial = False
        End If
End Function


Private Function ActualizarElPrecio(ByRef ADonde As String) As Boolean

    ActualizarElPrecio = False
    
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Precios Especiales"
    If Not InsertarLineasHistorico Then Exit Function
'    IncrementarProgres 2
    
    
    'Modificamos en cabeceras de Tarifas
    ADonde = "Modificando datos en cabecera de Precios Especiales"
    If Not ModificarCabecera Then Exit Function
'    IncrementarProgres 2
    ActualizarElPrecio = True
End Function


Private Function ModificarCabecera() As Boolean
'Modifica la tabla de cabeceras de Tarifas
Dim SQL As String

    On Error Resume Next

    SQL = "UPDATE " & NombreTabla & " SET precioac=precionu, precioa1=precion1, dtoespec=dtoespe1, fechanue=null, precionu=0, precion1=0"
    SQL = SQL & " WHERE codclien=" & Data1.Recordset!codclien & " AND codartic=" & DBSet(Data1.Recordset!codArtic, "T")
    conn.Execute SQL
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        ModificarCabecera = False
    Else
        ModificarCabecera = True
    End If
End Function


Private Function InsertarLineasHistorico() As Boolean
Dim SQL As String
Dim NumF As String
On Error Resume Next

    'Obtenemos la siguiente numero de linea de tarifa
    SQL = "codclien=" & Data1.Recordset!codclien & " AND codartic=" & DBSet(Data1.Recordset!codArtic, "T")
    NumF = SugerirCodigoSiguienteStr("spree1", "numlinea", SQL)

    SQL = "INSERT INTO spree1 (codclien, codartic, numlinea, fechanue, precioac, precioa1, dtoespec)"
    SQL = SQL & " VALUES (" & Data1.Recordset.Fields(0).Value & ", " & DBSet(Data1.Recordset.Fields(1).Value, "T") & ", "
    SQL = SQL & NumF & ", " & DBSet(Text1(4).Text, "F") & ", "
    SQL = SQL & DBSet(Data1.Recordset!precioac, "N") & ", " & DBSet(Data1.Recordset!precioa1, "N") & ", "
    SQL = SQL & DBSet(Data1.Recordset!dtoespec, "N") & ") "
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function


Private Sub BotonImprimir()
        frmListado.NumCod = Text1(0).Text
        AbrirListado (8) '8: Informe Movimientos Almacen
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(codclien=" & Text1(0).Text & " AND codartic=" & DBSet(Text1(1).Text, "T") & ")"
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub
