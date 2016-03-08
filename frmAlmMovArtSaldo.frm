VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmMovArtSaldo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos Articulos desde inventario"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12960
   ClipControls    =   0   'False
   Icon            =   "frmAlmMovArtSaldo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdActualizStock 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   7440
      Width           =   1455
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   5415
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9551
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hora"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Detalle"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cli/Pro/Tra"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nombre"
         Object.Width           =   6667
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Entrada"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Salida"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   12735
      Begin VB.CheckBox Check1 
         Caption         =   "Control stock"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4080
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   690
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   10920
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   690
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   690
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   9240
         MaxLength       =   16
         TabIndex        =   13
         Tag             =   "Cod. alma|N|N|||smoval|codalmac||N|"
         Text            =   "Text1"
         Top             =   210
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   210
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   10
         Tag             =   "Cod. Articulo|T|N|||smoval|codartic||N|"
         Text            =   "Text1"
         Top             =   210
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   210
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   12360
         Picture         =   "frmAlmMovArtSaldo.frx":000C
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   12360
         Picture         =   "frmAlmMovArtSaldo.frx":1A7E
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "STOCK inventario"
         Height          =   255
         Index           =   4
         Left            =   6120
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Stock ACTUAL"
         Height          =   255
         Index           =   3
         Left            =   9720
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha inventario"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   8880
         Picture         =   "frmAlmMovArtSaldo.frx":34F0
         ToolTipText     =   "Buscar artículo"
         Top             =   247
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmAlmMovArtSaldo.frx":35F2
         ToolTipText     =   "Buscar artículo"
         Top             =   247
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Artículo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Almacén"
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   7320
      Width           =   2505
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "BUSQUEDA"
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10440
      TabIndex        =   0
      Top             =   7440
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11640
      TabIndex        =   1
      Top             =   7440
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11640
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8280
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
      TabIndex        =   3
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmAlmMovArtSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmAlPropios 'Almacen Origen/Destino
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmArtic As frmAlmArticu2  'Articulos
Attribute frmArtic.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte

Dim kCampo As Integer
Dim PrimeraVez As Boolean
Dim btnPrimero As Byte 'Variable que indica el Nº del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String
Dim CadenaBusqueda As String 'Cadena para la consulta de de busqueda en Grid
Dim cadSeleccion2 As String 'Cadena de seleccion para FormulaSelection del Informe
'---- Laura: 27/09/2006
'cadena para la SQL de los totales de cantida e importe por articulo mostrado
'Dim cadSelGrid As String


Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Private HaDevueltoDatos As Boolean


Dim vStock As Currency
Dim RS As ADODB.Recordset

'------------------------------------------
Dim CadClie As String   '|codigo·nombre|
Dim CadProve As String
Dim cadTraba As String

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    If Modo = 1 Then HacerBusqueda
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Imprimir()
Dim Cad As String

    
    If Data1.Recordset.EOF Then Exit Sub
    
    'Resto parametros
    Cad = Trim(Text1(2).Text)
    If Cad = "" Then Cad = " ----"
    Cad = "|pDHArticulo=""Almacen: " & Text1(1).Text & " " & Text2(1).Text & "    Fecha inventario:  " & Cad
    If Image1(1).visible Then Cad = Cad & "     *** ERROR ***  Ficha articulo: " & Text1(4).Text
    Cad = Cad & """|"
    Cad = Cad & "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    If Text1(5).Text = "" Then
        cadTraba = "0"
    Else
        cadTraba = TransformaComasPuntos(ImporteFormateado(Text1(5).Text))
    End If
    Cad = Cad & "|Incial=" & cadTraba & "|"

            
    With frmImprimir
        .NombreRpt = "rAlmMovimInven.rpt"
        .OtrosParametros = Cad
        .NumeroParametros = 3
        
        Cad = "({smoval.codAlmac} = " & Data1.Recordset!codAlmac & ") AND ({smoval.codartic}=""" & DevNombreSQL(Data1.Recordset!codArtic) & """)"
        .FormulaSeleccion = Cad
        .EnvioEMail = False
        .opcion = 9
        .Titulo = "Informe Movimientos Articulos"
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub





Private Sub cmdActualizStock_Click()
    If Modo <> 2 Then Exit Sub
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub

    'Esta bien. No actualio nada
    If Me.Image1(0).visible Then Exit Sub


    'If Format(cmdActualizStock.Tag, FormatoCantidad) = Text1(5).Text Then Exit Sub

    If MsgBox("¿Coninuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    CadenaBusqueda = "UPDATE salmac set canstock =  " & DBSet(Me.cmdActualizStock.Tag, "N")
    CadenaBusqueda = CadenaBusqueda & " WHERE codartic = " & DBSet(Text1(0).Text, "T")
    CadenaBusqueda = CadenaBusqueda & " AND codalmac = " & Text1(1).Text
    conn.Execute CadenaBusqueda
    
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    CadenaBusqueda = "Articulo:  " & Text1(0).Text & " " & Text2(0).Text & vbCrLf
    CadenaBusqueda = CadenaBusqueda & "Almacen:  " & Text1(1).Text & " " & Text2(1).Text & vbCrLf
    CadenaBusqueda = CadenaBusqueda & "Inventario:  " & Text1(2).Text & "   Uds: " & Text1(5).Text & vbCrLf
    CadenaBusqueda = CadenaBusqueda & "Stock actual:  " & Text1(4).Text & "    Stock Real: " & Format(cmdActualizStock.Tag, FormatoCantidad)
    LOG.Insertar 33, vUsu, CadenaBusqueda
    Set LOG = Nothing
    '-------------------
    
    
    
    PonerCampos
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

   If Modo = 1 Then       'Buscar
        LimpiarCampos
        If Data1.Recordset Is Nothing Then PrimeraVez = True
        PonerModo 0
        PrimeraVez = False
       
    End If
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco Text1(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
   
    'ICONOS de La toolbar
    btnPrimero = 8 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 16 'Imprimir
        .Buttons(6).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    
    PrimeraVez = True
    
    NombreTabla = "smoval"
    Ordenacion = " ORDER BY codartic," & NombreTabla & ".codalmac, fechamov desc, horamovi "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
        
    Data1.CursorType = adOpenDynamic
    Data1.ConnectionString = conn
    CadenaConsulta = "Select codartic,codalmac from " & NombreTabla & " WHERE codartic = -1"
    Data1.RecordSource = CadenaConsulta
    'Data1.Refresh
    LimpiarCampos
    Modo = 0
    BotonBuscar
    
    cmdActualizStock.visible = False
    'If vUsu.Codigo Mod 1000 = 0 Then cmdActualizStock.visible = True
    If vUsu.Nivel = 0 Then cmdActualizStock.visible = True
    Screen.MousePointer = vbDefault
End Sub




Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacen Propios
    Text1(1).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArtic_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass

        cadB = ""
        cadB = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = cadB & " AND " & ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        'CadenaConsulta = "select codartic from " & NombreTabla & " WHERE " & cadB & " GROUP BY codartic " & Ordenacion
        CadenaConsulta = "select codartic,codalmac from " & NombreTabla & " WHERE " & cadB & " GROUP BY codartic,codalmac " & Ordenacion
        PonerCadenaBusqueda
        
               
        
        'cadb= Replace(cadSeleccion, ")", "}")
        cadSeleccion2 = "{smoval.codartic} = """ & RecuperaValor(CadenaDevuelta, 1)
        cadSeleccion2 = cadSeleccion2 & """ AND {smoval.codalmac} = " & RecuperaValor(CadenaDevuelta, 2)
    
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    'Codigo Articulos
    If Index = 0 Then
        Set frmArtic = New frmAlmArticu2
        'frmArtic.DatosADevolverBusqueda3 = "@1@" 'Abrimos en Modo Busqueda
        frmArtic.DesdeTPV = False
        frmArtic.Show vbModal
        Set frmArtic = Nothing
    Else
        Set frmA = New frmAlmAlPropios
        frmA.DatosADevolverBusqueda = "0"
        frmA.Show vbModal
        Set frmA = Nothing
    End If
    PonerFoco Text1(Index)
    Screen.MousePointer = vbDefault
End Sub











Private Sub lw1_DblClick()
'Abrir el formulario del Mantenimiento del que viene el Movimiento
'Se busca en histórico o en Form
Dim SQL As String
Dim Documento As String
    
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub

    Screen.MousePointer = vbHourglass
    Documento = lw1.SelectedItem.Tag


    Select Case lw1.SelectedItem.SubItems(2)
        Case "TRA" 'traspaso de almacenes
            'Traspaso de Almacen
            With frmAlmTraspaso
                .EsHistorico = True
                .hcoCodMovim = Documento
                .hcoFechaMovim = lw1.SelectedItem.Text
                .Show vbModal
            End With

        Case "REG" 'Movimientos de Almacen
                    'Movimientos de Almacen
            With frmAlmMovimientos
                .EsHistorico = True
                .hcoCodMovim = Documento
                .hcoFechaMovim = lw1.SelectedItem.Text
                .Show vbModal
            End With

        Case "ALV", "ART", "ALM", "ALZ", "ALI", "ALS", "ALO", "ALE", "ALR"
                                'ALV:Albaran de Venta (a clientes)
                                'ART: Albaran rectificativo
                                'ALM: ALbaran Mostrador
                                'ALZ: Albaranes "B"
                                'ALI: Albaranes INTERNOS
            'comprobar si el Albaran esta facturado o no
            'si no esta facturado abrir el formulario de Entrada de Albaranes: frmFacEntAlbaranes
            'si esta ya facturado abrir el histórico de facturas: frmFacHcoFacturas


            If vParamAplic.NumeroInstalacion = 2 Then
                If Val(vUsu.AlmacenPorDefecto) <> vParamAplic.AlmacenB Then
                    If lw1.SelectedItem.SubItems(2) = "ALZ" Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
            End If




            'consultamos si existe el albaran en la tabla de albaranes: scaalb
            SQL = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", lw1.SelectedItem.SubItems(2), "T", , "numalbar", Documento, "N")
            If SQL <> "" Then 'existe el Albaran
                If vParamAplic.TipoFormularioClientes = 0 Then
                         With frmFacEntAlbaranes2
                            If EsNumerico(Documento) Then
                                .hcoCodMovim = Format(Documento, "0000000")
                            Else
                                .hcoCodMovim = Documento
                            End If
                            .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                            .RecuperarFactu = False
                            .Show vbModal
                        End With

                Else
                    'FORMULARIO SAIL
                         With frmFacEntAlbSAIL
                            If EsNumerico(Documento) Then
                                .hcoCodMovim = Format(Documento, "0000000")
                            Else
                                .hcoCodMovim = Documento
                            End If
                            .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                            .RecuperarFactu = False
                            .Show vbModal
                        End With
                End If

            Else 'No existe en albaran, abrir Historico Factura
                With frmFacHcoFacturas2
                    .DesdeFichaCliente = False
                    If EsNumerico(Documento) Then
                        .hcoCodMovim = Format(Documento, "0000000")
                    Else
                        .hcoCodMovim = Documento
                    End If
                    .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                    .hcoFechaMov = lw1.SelectedItem.Text

                    .Show vbModal
                End With
            End If

        Case "ALR" 'Albaran de Reparacion (a clientes)
                If vParamAplic.TipoFormularioClientes = 0 Then
                     With frmFacEntAlbaranes2
                        If EsNumerico(Documento) Then
                            .hcoCodMovim = Format(Documento, "0000000")
                        Else
                            .hcoCodMovim = Documento
                        End If
                        .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                        .RecuperarFactu = False
                        .Show vbModal
                    End With
                End If
        Case "ALC" 'Albaran de Compra (a Proveedores)
            'comprobar si el Albaran esta facturado o no
            'si no esta facturado abrir el formulario de Entrada de Albaranes: frmComEntAlbaranes
            'si esta ya facturado abrir el histórico de facturas: frmComHcoFacturas

            'consultamos si existe el albaran en la tabla de albaranes: scaalp
            SQL = DevuelveDesdeBDNew(conAri, "scaalp", "numalbar", "codprove", lw1.SelectedItem.SubItems(3), "N", , "numalbar", Documento, "T", "fechaalb", lw1.SelectedItem.Text, "F")
            If SQL <> "" Then 'existe el Albaran
                If vParamAplic.TipoFormularioClientes = 0 Then
                    With frmComEntAlbaranes
                        .hcoCodMovim = Documento
                        .hcoFechaMovim = lw1.SelectedItem.Text
                        .hcoCodProve = lw1.SelectedItem.SubItems(3) 'aqui es el proveedor
                        .EsHistorico = False
                        .Show vbModal
                    End With
                Else
                    'SAIL
                    With frmComEntAlbaranSA
                        .hcoCodMovim = Documento
                        .hcoFechaMovim = lw1.SelectedItem.Text
                        .hcoCodProve = lw1.SelectedItem.SubItems(3) 'aqui es el proveedor
                        .EsHistorico = False
                        .Show vbModal
                    End With
                End If
            Else
                'alb pasados a hco
                SQL = DevuelveDesdeBDNew(conAri, "schalp", "numalbar", "codprove", lw1.SelectedItem.SubItems(3), "N", , "numalbar", Documento, "T", "fechaalb", lw1.SelectedItem.Text, "F")
                If SQL <> "" Then 'existe el Albaran
                    If vParamAplic.TipoFormularioClientes = 0 Then
                        With frmComEntAlbaranes
                            .hcoCodMovim = Documento
                            .hcoFechaMovim = lw1.SelectedItem.Text
                            .hcoCodProve = lw1.SelectedItem.SubItems(3) 'aqui es el proveedor
                            .EsHistorico = True
                            .Show vbModal
                        End With
                    Else
                        'SAIL
                        With frmComEntAlbaranSA
                            .hcoCodMovim = Documento
                            .hcoFechaMovim = lw1.SelectedItem.Text
                            .hcoCodProve = lw1.SelectedItem.SubItems(3) 'aqui es el proveedor
                            .EsHistorico = True
                            .Show vbModal
                        End With
                    End If
                Else
                    'No existe en albaran, abrir Historico Factura
                    If vParamAplic.TipoFormularioClientes = 0 Then
                        With frmComHcoFacturas2
                            .hcoCodMovim = Documento
                            .hcoFechaMovim = lw1.SelectedItem.Text
                            .hcoCodProve = lw1.SelectedItem.SubItems(3) 'aqui es el proveedor
                            .Show vbModal
                        End With
                    Else
                        'SAIL
                        
                        
                    End If
                
                End If
            End If


        '**********************************
        'Laura: modificado 11/09/06
'        Case "FTI" 'Factura Ticket de venta
        Case "ATI" 'Albaran Ticket de venta
        '**********************************
            'Abrir el historico de facturas
             With frmFacHcoFacturas2
                .DesdeFichaCliente = False
                If EsNumerico(Documento) Then
                    .hcoCodMovim = Format(Documento, "0000000")
                Else
                    .hcoCodMovim = Documento
                End If
                .hcoCodTipoM = lw1.SelectedItem.SubItems(2)
                .hcoFechaMov = lw1.SelectedItem.Text
                .Show vbModal
            End With
    End Select

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

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    If Trim(Text1(Index).Text) = "" Then
        If Index < 2 Then Text2(Index).Text = ""
        Exit Sub
    ElseIf (Modo = 1) Then 'Busqueda
'        If index = 0 Then
'            Text2(0).Text = PonerNombreDeCod(Text1(index), conAri, "sartic", "nomartic")
'        Else
'            If PonerFormatoEntero(Text1(index)) Then
'                Text2(1).Text = PonerNombreDeCod(Text1(index), conAri, "salmpr", "nomalmac")
'            Else
'                Text2(1).Text = ""
'            End If
'        End If
        
    End If
End Sub







Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            BotonBuscar
        Case 2 'Ver Todos
            BotonVerTodos
        Case 5 'Imprimir
            Imprimir
        Case 6  'Salir
            Unload Me
        Case 8 To 11 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte
    
    
    lblIndicador.Caption = "Poner modo"
    lblIndicador.Refresh
    Modo = Kmodo
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    NumReg = 1
    If Not Data1.Recordset Is Nothing Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
   'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    b = Modo <> 1
    lblIndicador.Caption = "Bloq txt"
    lblIndicador.Refresh
    BloquearTxt Text1(0), b
    BloquearTxt Text1(1), b
    'BloquearText1 Me, Modo
    
    
    lblIndicador.Caption = "Select case"
    lblIndicador.Refresh
    Select Case Kmodo
    Case 0    'Modo Inicial
        Toolbar1.Buttons(5).Enabled = False 'Imprimir
        PonerBotonCabecera True
    Case 1 'Modo Buscar
        lblIndicador.Caption = "BUSQUEDA"
        Toolbar1.Buttons(5).Enabled = False 'Imprimir
        PonerBotonCabecera False
        PonerFoco Text1(0)
        
    Case 2    'Preparamos para que pueda Modificar
        PonerBotonCabecera True
    End Select
           
    b = Modo <> 0 And Modo <> 2
  
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I

    lblIndicador.Caption = "Poner long. campos"
    lblIndicador.Refresh
    'PonerLongCampos   'Lo acabo de comentar  03/11/2010     En ejecucion se queda colgado en este punto ¿Pq?  No lo se

    b = (Kmodo >= 3) Or Modo = 1
    Toolbar1.Buttons(1).Enabled = Not b
    Toolbar1.Buttons(2).Enabled = Not b
    lblIndicador.Caption = ""
    lblIndicador.Refresh
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub



Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lw1.ListItems.Clear
    Image1(0).visible = False
    Image1(1).visible = False
    Me.Check1.Value = 0
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    
'    CalcularTotales
End Sub


'Private Function MontaSQLCarga(enlaza As Boolean) As String
''--------------------------------------------------------------------
'' MontaSQlCarga:
''   Basándose en la información proporcionada por el vector de campos
''   crea un SQl para ejecutar una consulta sobre la base de datos que los
''   devuelva.
'' Si ENLAZA -> Enlaza con el data1
''           -> Si no lo cargamos sin enlazar a ningun campo
''--------------------------------------------------------------------
'Dim SQL As String
'Dim selSQL As String
'Dim cadBuscar2 As String
'Dim I As Integer
'
'    cadSelGrid = ""
'
'    selSQL = "SELECT smoval.codartic, smoval.codalmac, nomalmac, fechamov, horamovi, if(smoval.tipomovi=0,""S"",""E"") as tipomovi, detamovi, "
'    selSQL = selSQL & "cantidad, impormov, codigope, letraser, document, numlinea "
'
'    SQL = " FROM (smoval LEFT OUTER JOIN salmpr on smoval.codalmac=salmpr.codalmac)"
'    If enlaza Then
'        If EsBusqueda And CadenaBusqueda <> "" Then
'            'LAura: 29/09/06
''            If Data1.Recordset.RecordCount > 1 Then
'            'Si devuelve + de 1 registro en el DataGrid poner la info del primer articulo
'                'quitar codartic de la cadena busqueda
''                i = InStr(CadenaBusqueda, "(smoval.codartic")
''                If i > 0 Then
''
''                End If
'
'                SQL = SQL & CadenaBusqueda & " AND codartic=" & DBSet(Text1(0).Text, "T")
''            Else
''                SQL = SQL & CadenaBusqueda
''            End If
'        Else
'            SQL = SQL & " WHERE codartic = " & DBSet(Text1(0).Text, "T")
'        End If
'    Else
'        SQL = SQL & " WHERE codartic = '-1'"
'    End If
'    SQL = SQL & " " & Ordenacion & " DESC "
'    '---- Laura: 27/09/2006
'    cadSelGrid = SQL
'    SQL = selSQL & SQL
'    '----
'    MontaSQLCarga = SQL
'End Function


Private Sub BotonBuscar()
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
        Me.lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        PonerFoco Text1(0)
        Text1(1).Text = vUsu.AlmacenPorDefecto
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
    EsBusqueda = False
'    LimpiarCampos
'    'Ponemos el grid lineasfacturas enlazando a ningun sitio
'    CargaGrid False
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
       
    Else
        CadenaConsulta = "Select codartic,codalmac from " & NombreTabla & " group by codartic,codalmac " & Ordenacion
        PonerCadenaBusqueda
        Toolbar1.Buttons(5).Enabled = True 'Imprimir
    End If
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
Dim bol As Boolean

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then Me.lblIndicador.Caption = ""
    
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
Dim cadB2 As String

    cadB = ObtenerBusqueda(Me, False)
    cadSeleccion2 = ObtenerBusqueda(Me, True) 'Para la consulta de report




    'If vParamAplic.AlmacenB > 1 Then
    If vParamAplic.NumeroInstalacion = 2 Then
        'HERBELCA
        If vUsu.CodigoAgente > 0 Then
            'Es solo un agente. Solo puede ver sus movimientos
            If vUsu.AlmacenPorDefecto > 0 Then
                If cadB <> "" Then cadB = cadB & " AND "
                If cadSeleccion2 <> "" Then cadSeleccion2 = cadSeleccion2 & " AND "
                    
                cadB = cadB & " smoval.codalmac = " & vUsu.AlmacenPorDefecto
                cadSeleccion2 = cadSeleccion2 & " {smoval.codalmac} = " & vUsu.AlmacenPorDefecto
            End If
        End If
    End If









        If cadB <> "" Then
            'Cadena para el Data1
            CadenaConsulta = "select codartic,codalmac from " & NombreTabla & " WHERE " & cadB & " GROUP BY codartic,codalmac " & Ordenacion
            'Cadena para el Datagrid y el Data2
            'el codartic no se incluye en la cadB de las lineas pq siempre
            'se muestran las de un codartic concreto
            'Text1(0).Text = ""
            'cadB2 = ObtenerBusqueda(Me, False)
'            CadenaBusqueda = ""
            'If cadB2 <> "" Then 'Para cargar la consulta del CargaGrid
            '    CadenaBusqueda = " WHERE " & cadB2
            'Else
            '    CadenaBusqueda = ""
            'End If

        Else
            'obtener todos los articulos
            CadenaConsulta = "select codartic,codalmac from " & NombreTabla & " GROUP BY codartic,codalmac " & Ordenacion
            CadenaBusqueda = ""
        End If
        PonerCadenaBusqueda
'    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim I As Byte
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    lblIndicador.Caption = "Obt SQL"
    lblIndicador.Refresh
    Data1.RecordSource = CadenaConsulta


    lblIndicador.Caption = "Refresh"
    lblIndicador.Refresh
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de búsqueda", vbInformation
        Screen.MousePointer = vbDefault
        PonerFoco Text1(0)
      
        Exit Sub
    Else
        PonerModo 2
        Toolbar1.Buttons(5).Enabled = True 'Imprimir
     
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
Dim I As Integer
Dim Aux As String

On Error GoTo EPonerCampos
 
    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Aux = "ctrstock"
    'Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sartic", "nomartic")
    Text2(0).Text = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", Text1(0).Text, "T", Aux)
    Me.Check1.Value = 0
    If Aux = "1" Then Me.Check1.Value = 1
    
    'De salmac
    Aux = "Select nomalmac,canstock ,stockinv ,fechainv ,horainve from salmac,salmpr where salmac.codAlmac = salmpr.codAlmac "
    Aux = Aux & " AND salmac.codAlmac = " & Data1.Recordset!codAlmac
    Aux = Aux & " AND codartic =" & DBSet(Data1.Recordset!codArtic, "T")
    Set RS = New ADODB.Recordset
    RS.Open Aux, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Aux = "||||"
    vStock = 0
    Text2(1).Text = ""
    If Not RS.EOF Then
        Text2(1).Text = RS!nomalmac
    
        Aux = ""
        If Not IsNull(RS!FechaINV) Then Aux = Aux & Format(RS!FechaINV, "dd/mm/yyyy")

        Aux = Aux & "|"
        If Not IsNull(RS!horainve) Then Aux = Aux & Format(RS!horainve, "hh:mm:ss")
        Aux = Aux & "|"
        
        Aux = Aux & Format(RS!CanStock, FormatoCantidad) & "|"
        Aux = Aux & Format(DBLet(RS!Stockinv, "N"), FormatoCantidad) & "|"
        vStock = DBLet(RS!Stockinv, "N")
    End If
    RS.Close
    For I = 1 To 4
        Text1(I + 1).Text = RecuperaValor(Aux, I)
    Next I
    
    
    'AHora pongo los datos del list viesw
    Me.Image1(0).visible = False
    Me.Image1(1).visible = False
    
    CargaListView
    
    
    
    
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
    
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
    Set RS = Nothing
End Sub



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
            
    Cad = Cad & "Código|smoval|codartic|T||18·Denominacion|sartic|nomartic|T||70·Alm.|smoval|codalmac|T||7·"
    Tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ") "
    Tabla = Tabla & " GROUP BY smoval.codartic,smoval.codalmac "
    Titulo = "Movimientos de Articulos"

           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|2|"
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
            Toolbar1.Buttons(5).Enabled = True 'Imprimir
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Function PonerNombreCliente(Codigo As Long, movim As String) As String
'Devuelve el nombre del Trabajador/Cliente/Proveedor para ponerlo en la caja de texto text2 en la parte inferior del form
Dim Nombre As String
'
    'CadClie
    'CadProve
    'cadTraba

    Select Case movim
        Case "TRA", "REG", "DFI", "ALI", "ALS"
            If Not EstaEnCadenas(Codigo, 1, Nombre) Then
                'Obtener nombre de la tabla de trabajadores
                Nombre = DevuelveDesdeBDNew(conAri, "straba", "nomtraba", "codtraba", CStr(Codigo), "N")
                AnyadirCadena Codigo, 1, Nombre
            End If
        Case "ALV", "ALR", "ALM", "ART", "FAV", "FTI", "ATI", "ALZ", "ALO", "ALE"
            If Not EstaEnCadenas(Codigo, 2, Nombre) Then
                'Obtener nombre de la tabla de Clientes
                Nombre = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", CStr(Codigo), "N")
                AnyadirCadena Codigo, 2, Nombre
            End If
            'Label2.Caption = "Cliente"
        Case "ALC"
            'Obtener el nombre de la tabla de Proveedores
            If Not EstaEnCadenas(Codigo, 3, Nombre) Then
                Nombre = DevuelveDesdeBDNew(conAri, "sprove", "nomprove", "codprove", CStr(Codigo), "N")
                AnyadirCadena Codigo, 3, Nombre
            End If
            'Label2.Caption = "Proveedor"
    End Select
    PonerNombreCliente = Nombre
End Function

Private Function EstaEnCadenas(Codigo As Long, TipoRef As Byte, ByRef Nombre As String) As Boolean
Dim J As Long
Dim I As Long
Dim Aux As String
'        cadTraba
'    CadProve
'    cadTraba
    Aux = "|" & Codigo & "·"
    If TipoRef = 2 Then
        J = InStr(1, CadClie, Aux)
        
    ElseIf TipoRef = 1 Then
        J = InStr(1, cadTraba, Aux)
    Else
        J = InStr(1, CadProve, Aux)
    End If
    
    'J = 0
    
    If J = 0 Then Exit Function
    
    J = J + Len(Aux)
    If TipoRef = 2 Then
        I = InStr(J, CadClie, "|")
        Aux = Mid(CadClie, J, I - J)
    ElseIf TipoRef = 1 Then
        I = InStr(J, cadTraba, "|")
        Aux = Mid(cadTraba, J, I - J)
    Else
        
        I = InStr(J, CadProve, "|")
        Aux = Mid(CadProve, J, I - J)
    End If
    Nombre = Aux
    EstaEnCadenas = True
End Function


Private Function AnyadirCadena(Codigo As Long, TipoRef As Byte, ByRef Nombre As String) As Boolean
    If TipoRef = 2 Then
        CadClie = CadClie & Codigo & "·" & Nombre & "|"
    ElseIf TipoRef = 1 Then
        cadTraba = cadTraba & Codigo & "·" & Nombre & "|"
    Else
        CadProve = CadProve & Codigo & "·" & Nombre & "|"
    End If
    
End Function



'Private Sub CargaListView()
'Dim t1
'Dim Tt
'Dim i As Integer
'
'    For i = 1 To 3
'        RS.Open "Select * from smoval", conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        RS.Close
'        t1 = Timer
'        CargaListView2
'        t1 = Timer - t1
'        Tt = Tt + t1
'    Next
'    Caption = Tt
'End Sub

Private Sub CargaListView()
Dim Cantidad As Currency
Dim Aux As String
Dim IT As ListItem

    lw1.ListItems.Clear
    CadClie = "|"
    CadProve = "|"
    cadTraba = "|"
    Aux = "SELECT smoval.codartic, smoval.codalmac, fechamov, horamovi, tipomovi, detamovi, "
    Aux = Aux & " cantidad,  codigope, letraser, document, numlinea "
    Aux = Aux & " FROM  smoval WHERE codartic =" & DBSet(Data1.Recordset!codArtic, "T")
    Aux = Aux & " AND codalmac =" & DBSet(Data1.Recordset!codAlmac, "N")
    
    'Si lleva fehca inv
    If Text1(2).Text <> "" Then
        Aux = Aux & " AND fechamov > " & DBSet(Text1(2).Text, "F")
    End If
    
    
    Aux = Aux & " order by Fechamov , horamovi "
     RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = Format(RS!FechaMov, "dd/mm/yyyy")
        IT.SubItems(1) = Format(RS!horamovi, "hh:mm:ss")
        IT.SubItems(2) = RS!detamovi
        IT.SubItems(3) = RS!codigope
        
        'If It.SubItems(2) = "ALR" And It.SubItems(3) = "752" Then
        
        'smoval.tipomovi=0,""S""
        '   0: SALIDA
        '   1: ENTRADA
        Cantidad = RS!Cantidad
        If RS!tipomovi = 1 Then
            IT.SubItems(5) = Format(Cantidad, FormatoCantidad)
            IT.SubItems(6) = " "
            
        Else
            IT.SubItems(5) = " "
            IT.SubItems(6) = Format(Cantidad, FormatoCantidad)
            Cantidad = -Cantidad
        End If
        vStock = vStock + Cantidad
        IT.SubItems(7) = Format(vStock, FormatoCantidad)
        
       ' If Me.chkCargaNombres.Value = 1 Then
            Aux = PonerNombreCliente(RS!codigope, RS!detamovi)
            If Aux = "" Then Aux = "Error leyendo desde BD"
            IT.SubItems(4) = Aux
       ' End If
       
       
       
       IT.Tag = DBLet(RS!document)
       RS.MoveNext
        
        
    
    Wend
    RS.Close
    
    'Si es el mismo importe k el stock
    Me.cmdActualizStock.Tag = vStock  'me to el stock aqui
    CadClie = Format(vStock, FormatoCantidad)
    Me.Image1(0).visible = CadClie = Text1(4).Text
    Me.Image1(1).visible = Not Me.Image1(0).visible
    
    CadClie = "":    CadProve = "":    cadTraba = ""   'liberar espacio
End Sub

