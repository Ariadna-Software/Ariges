VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmHcoInven 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hist�rico Inventario"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   ClipControls    =   0   'False
   Icon            =   "frmAlmHcoInven.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   7335
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1095
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "Cod. Articulo|T|N|||shinve|codartic||S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   795
         Picture         =   "frmAlmHcoInven.frx":000C
         ToolTipText     =   "Buscar art�culo"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Art�culo"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   0
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   15
      Text            =   "nom"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   6
      Tag             =   "Existencia|N|N|||shinve|existenc|#,###,###,##0.00|N|"
      Text            =   "cantidad"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   4440
      MaxLength       =   16
      TabIndex        =   5
      Tag             =   "Hora|H|N|||shinve|horainve|hh:mm:ss|N|"
      Text            =   "hora"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   5640
      Width           =   2505
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   3360
      MaxLength       =   11
      TabIndex        =   4
      Tag             =   "Fecha|F|N|||shinve|fechainv|dd/mm/yyyy|N|"
      Text            =   "fecha"
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "Cod. Almacen|N|N|0|999|shinve|codalmac|000|S|"
      Text            =   "codalmac"
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   960
      TabIndex        =   12
      ToolTipText     =   "Buscar almacen"
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   5790
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6795
      TabIndex        =   2
      Top             =   5790
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6795
      TabIndex        =   11
      Top             =   5790
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
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
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2760
      Top             =   5520
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
      Bindings        =   "frmAlmHcoInven.frx":010E
      Height          =   4170
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7355
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   2880
      Top             =   5880
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
      TabIndex        =   9
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmAlmHcoInven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
'Private WithEvents frmF As frmCal 'Calendario de Fechas
Private WithEvents frmA As frmAlmAlPropios 'Almacen Origen/Destino
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmArtic As frmBasico2  'Articulos
Attribute frmArtic.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte

Dim kCampo As Integer
Dim PrimeraVez As Boolean
Dim btnPrimero As Byte 'Variable que indica el N� del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String
Dim CadenaBusqueda As String 'Cadena para la consulta de de busqueda en Grid
'Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe

Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim i As Integer
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'Busqueda
            HacerBusqueda
        Case 4 'Modificar
            If DatosOk Then
'                If ModificaDesdeFormulario(Me, 3) Then
                If ModificarLinea Then
                      TerminaBloquear
                      i = data1.Recordset.Fields(0)
'                      LLamaLineas Modo, 0
                      PonerModo 2
                      CancelaADODC Me.Data2
                      
                      data1.Recordset.Find (data1.Recordset.Fields(0).Name & " =" & i)
                      CargaGrid True
                  End If
                  DataGrid1.SetFocus
            End If
    End Select
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Function DatosOk() As Boolean
Dim b As Boolean
    
    b = CompForm(Me, 3)
    If Not b Then Exit Function
       
    DatosOk = b
End Function


Private Sub Imprimir()
'Dim cad As String
'Dim numParam As Byte
'
'    'Resto parametros
'    cad = ""
'    cad = cad & "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
'    numParam = 1
'
'    With frmImprimir
'        .NombreRPT = "rAlmMovim.rpt"
'        .OtrosParametros = cad
'        .NumeroParametros = numParam
'        .FormulaSeleccion = cadSeleccion
'        '.SoloImprimir = True
'        .Opcion = 9
'        .Titulo = ""
'        .Show vbModal
'    End With
End Sub


Private Sub cmdAux_Click()
'Abre Formulario de Mantenimiento de Almacenes Propios
    Set frmA = New frmAlmAlPropios
    frmA.DatosADevolverBusqueda = "0"
    frmA.Show vbModal
    Set frmA = Nothing
    PonerFoco txtAux(0)
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Busqueda
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
        Case 4 'Modificar
            PonerModo 2
            LLamaLineas 10
    End Select

ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
   
    'ICONOS de La toolbar
    btnPrimero = 11 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 4 'Modificar
        
        .Buttons(8).Image = 16 'Imprimir
        .Buttons(9).Image = 15 'Salir
        
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
    NombreTabla = "shinve"
    Ordenacion = " ORDER BY codartic, codalmac, fechainv desc, horainve "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codartic = -1"
    data1.RecordSource = CadenaConsulta
    data1.Refresh
    PonerCampos
    PonerModo 0
    
    CargaGrid (Modo = 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim tots As String
Dim SQL As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
     
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, PrimeraVez
    
    'SELECT shinve.codartic, shinve.codalmac, salmpr.nomalmac, shinve.fechainv, shinve.horainve,existenc
    tots = "N||||0|;S|txtAux(0)|T|Alm.|800|;S|cmdAux|B||0|;S|txtAux2(0)|T|Nom. Alm.|2500|;S|txtAux(1)|T|Fecha|1150|;"
    tots = tots & "S|txtAux(2)|T|Hora|1050|;S|txtAux(3)|T|Existencia|1200|;"
    
    arregla tots, DataGrid1, Me
    DataGrid1.Columns(5).Alignment = dbgRight
    
    DataGrid1.ScrollBars = dbgAutomatic

    DataGrid1.Enabled = b
    If Modo = 2 Then DataGrid1.Enabled = True
    PrimeraVez = False
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacen Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArtic_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim codArtic As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
            cadB = ""
            cadB = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadenaConsulta = "select codartic from " & NombreTabla & " WHERE " & cadB & " GROUP BY codartic " & Ordenacion
            PonerCadenaBusqueda
            
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    'Codigo Articulos
    If Index = 0 Then
        Set frmArtic = New frmBasico2
        'frmArtic.DatosADevolverBusqueda3 = "@1@" 'Abrimos en Modo Busqueda
'        frmArtic.DesdeTPV = False
'        frmArtic.Show vbModal
        AyudaArticulos frmArtic, Text1(0)
        Set frmArtic = Nothing
    End If
    PonerFoco Text1(0)
    Screen.MousePointer = vbDefault
End Sub




Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)

    If Text1(Index).BackColor = vbYellow Then Text1(Index).BackColor = vbWhite

    If Trim(Text1(Index).Text) = "" Then
        Text2(Index).Text = ""
        Exit Sub
    ElseIf (Modo = 1 And IsNumeric(Text1(Index))) Then
        Text2(0).Text = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic")
    End If
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    If Modo = 1 Then
        ConseguirFoco txtAux(Index), Modo
    End If
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 3 And KeyCode = 40 Then
        PonerFocoBtn Me.cmdAceptar
   Else
        KEYdown KeyCode
   End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 And KeyAscii = 12 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String 'Para mensajes

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'cod. almacen
            If txtAux(Index).Text = "" Then
             
            Else
                devuelve = DevuelveDesdeBD(conAri, "nomalmac", "salmpr", "codalmac", txtAux(Index).Text, "N")
'                Text2(1).Text = SQL
                If devuelve = "" Then 'No existe
                    devuelve = "No existe el Almacen" & vbCrLf
                    devuelve = devuelve & "C�digo: " & txtAux(Index).Text
                    MsgBox devuelve, vbExclamation
                    PonerFoco txtAux(Index)
                Else
                    txtAux(Index).Text = Format(txtAux(Index).Text, "000")
                End If
            End If
            
        Case 1 'Fecha Movimiento
             If txtAux(Index).Text <> "" Then PonerFormatoFecha txtAux(Index)
        Case 3
            If txtAux(Index).Text <> "" Then
                If Not PonerFormatoDecimal(txtAux(Index), 1) Then
                    PonerFoco txtAux(Index)
'                Else
'                    PonerFocoBtn Me.cmdAceptar
                End If
'            Else
'                  PonerFocoBtn Me.cmdAceptar
            End If
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            BotonBuscar
        Case 2 'Ver Todos
            BotonVerTodos
        Case 5 'Modificar
            If BLOQUEADesdeFormulario(Me) Then BotonModificar
        Case 8 'Imprimir
'            Imprimir
        Case 9  'Salir
            Unload Me
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    
    PonerIndicador Me.lblIndicador, Modo
    
    NumReg = 1
    If Not data1.Recordset.EOF Then
        If data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg

   'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    b = (Modo = 0) Or (Modo = 2)
    PonerBotonCabecera b
              
    b = Modo <> 0 And Modo <> 2
  
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2) Or (Modo = 0)
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
'    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = b
'    Me.mnVerTodos.Enabled = b
    
    b = (Modo = 2)

    'Modificar
    Toolbar1.Buttons(5).Enabled = b
'    Me.mnModificar.Enabled = b

    'Imprimir
    Toolbar1.Buttons(8).Enabled = False
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData data1, Index
    PonerCampos
    CargaGrid True
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String

    SQL = "SELECT shinve.codartic, shinve.codalmac, salmpr.nomalmac, shinve.fechainv, shinve.horainve,existenc "
    SQL = SQL & " FROM (shinve INNER JOIN salmpr on shinve.codalmac=salmpr.codalmac)"
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
'            If Data1.Recordset.RecordCount > 1 Then
            'Si devuelve + de 1 registro en el DataGrid poner la info del primer articulo
                SQL = SQL & CadenaBusqueda & " AND codartic=" & DBSet(Text1(0).Text, "T")
'            Else
'                SQL = SQL & CadenaBusqueda
'            End If
        Else
            SQL = SQL & " WHERE codartic = " & DBSet(Text1(0).Text, "T")
        End If
    Else
        SQL = SQL & " WHERE codartic = '-1'"
    End If
    SQL = SQL & " " & Ordenacion & " DESC "
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
Dim anc As Single
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
            
        anc = ObtenerAlto(Me.DataGrid1)
        LLamaLineas anc
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonModificar()
Dim anc As Single
Dim i As Integer
    
    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    PonerModo 4
    
    anc = ObtenerAlto(Me.DataGrid1)

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(1).Text
    txtAux(1).Text = DataGrid1.Columns(3).Text
    txtAux(2).Text = DataGrid1.Columns(4).Text
    txtAux(3).Text = DataGrid1.Columns(5).Text
    LLamaLineas anc
   
   'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonVerTodos()
'Ver todos
    EsBusqueda = False
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
        CargaGrid True
    Else
        CadenaConsulta = "Select codartic from " & NombreTabla & " group by codartic " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerBotonCabecera(b As Boolean)

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then Me.lblIndicador.Caption = ""
    
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
'    cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            'Cadena para el Data1
            CadenaConsulta = "select codartic from " & NombreTabla & " WHERE " & cadB & " GROUP BY codartic " & Ordenacion
            'Cadena para el Datagrid y el Data2
            CadenaBusqueda = " WHERE " & cadB 'Para cargar la consulta del CargaGrid
        Else
            'obtener todos los articulos
            CadenaConsulta = "select codartic from " & NombreTabla & " GROUP BY codartic " & Ordenacion
            CadenaBusqueda = ""
        End If
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim i As Byte
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    data1.RecordSource = CadenaConsulta

    data1.Refresh
    If data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de b�squeda", vbInformation
        Screen.MousePointer = vbDefault
        PonerFoco Text1(0)
        'Limpiar los Campos Auxiliares
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
        Exit Sub
    Else
        PonerModo 2
        LLamaLineas 10
        PonerCampos
        CargaGrid True
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

    If data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, data1
    Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sartic", "nomartic")
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
            
    cad = cad & "Articulo|shinve|codartic|T||25�Denominacion|sartic|nomartic|T||70�"
    tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ") "
'        tabla = tabla & " GROUP BY shinve.codartic "
    'tabla = "sartic"
    Titulo = "Hist�rico Inventario"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
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
            PonerFoco Text1(kCampo)
'            Toolbar1.Buttons(5).Enabled = True 'Imprimir
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single)
Dim jj As Byte
Dim ini As Byte
Dim b As Boolean

    DeseleccionaGrid Me.DataGrid1
    
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar Lineas
    
    If Modo = 4 Then 'modificar
        ini = 1
    Else
        ini = 0
    End If
    
    For jj = ini To txtAux.Count - 1
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
    Next jj

    b = (Modo = 1)
    Me.cmdAux.Height = DataGrid1.RowHeight
    Me.cmdAux.Top = alto
    Me.cmdAux.visible = b
End Sub


Private Function ModificarLinea() As Boolean
Dim SQL As String
On Error GoTo EModificar

    ModificarLinea = False
    SQL = "UPDATE " & NombreTabla & " SET fechainv=" & DBSet(txtAux(1).Text, "F")
    SQL = SQL & ", horainve='" & Format(txtAux(1).Text & " " & txtAux(2).Text, "yyyy-mm-dd hh:mm:ss") & "'"
    SQL = SQL & ", existenc=" & DBSet(txtAux(3).Text, "N")
    SQL = SQL & " WHERE codartic=" & DBSet(Text1(0).Text, "T") & " AND codalmac=" & Me.Data2.Recordset.Fields(1).Value
    conn.Execute SQL
    ModificarLinea = True
EModificar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Linea", Err.Description
End Function
