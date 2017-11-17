VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmMovPuntos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos albaranes / puntos"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   ClipControls    =   0   'False
   Icon            =   "frmAlmMovPuntos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdActualizStock 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   7320
      Width           =   1455
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   5655
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9975
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Numero"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "puntos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   10335
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   210
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   16
         TabIndex        =   9
         Tag             =   "Cod. cliente|N|N|||smovalpuntos|codclien||N|"
         Text            =   "Text1"
         Top             =   210
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   9960
         Picture         =   "frmAlmMovPuntos.frx":000C
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   9960
         Picture         =   "frmAlmMovPuntos.frx":1A7E
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Puntos"
         Height          =   255
         Index           =   3
         Left            =   7320
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmAlmMovPuntos.frx":34F0
         ToolTipText     =   "Buscar artículo"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   7200
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
      Left            =   7920
      TabIndex        =   0
      Top             =   7320
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   7320
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   7320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
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
Attribute VB_Name = "frmAlmMovPuntos"
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


Private Modo As Byte

Dim kCampo As Integer
Dim PrimeraVez As Boolean
Dim btnPrimero As Byte 'Variable que indica el Nº del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String
Dim CadenaBusqueda As String 'Cadena para la consulta de de busqueda en Grid
'---- Laura: 27/09/2006
'cadena para la SQL de los totales de cantida e importe por articulo mostrado
'Dim cadSelGrid As String


Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Private HaDevueltoDatos As Boolean


Dim vStock As Currency
Dim Rs As ADODB.Recordset



Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    If Modo = 1 Then HacerBusqueda
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Imprimir()
Dim cad As String

    
    If Data1.Recordset.EOF Then Exit Sub



            
    With frmImprimir
        .NombreRPT = "rPuntosCliente.rpt"
        .OtrosParametros = cad
        .NumeroParametros = 0
        
        cad = "({sclien.codclien} = " & Data1.Recordset!codClien & ")"
        .FormulaSeleccion = cad
        .EnvioEMail = False
        .Opcion = 9
        .Titulo = "Informe puntos"
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
    MsgBox "No hago nada"

    Exit Sub
        
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
    

    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
        
    Data1.CursorType = adOpenDynamic
    Data1.ConnectionString = conn
    CadenaConsulta = "Select codclien from smovalpuntos WHERE false"
    Data1.RecordSource = CadenaConsulta
    'Data1.Refresh
    LimpiarCampos
    Modo = 0
    BotonBuscar
    
    cmdActualizStock.visible = False
    'If vUsu.Codigo Mod 1000 = 0 Then cmdActualizStock.visible = True
    If vUsu.Login = "root" Then cmdActualizStock.visible = True
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
         
        CadenaConsulta = "select distinct codclien from smovalpuntos WHERE " & cadB & " ORDER BY codclien"
        PonerCadenaBusqueda
        
               
    
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


    Select Case lw1.SelectedItem.SubItems(1)

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
                If Val(vUsu.AlmacenPorDefecto2) <> vParamAplic.AlmacenB Then
                    If lw1.SelectedItem.SubItems(2) = "ALZ" Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
            End If




            'consultamos si existe el albaran en la tabla de albaranes: scaalb
            Documento = lw1.SelectedItem.SubItems(3)
            SQL = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", lw1.SelectedItem.SubItems(1), "T", , "numalbar", Documento, "N")
            If SQL <> "" Then 'existe el Albaran
                If vParamAplic.TipoFormularioClientes = 0 Then
                         With frmFacEntAlbaranes2
                            If EsNumerico(Documento) Then
                                .hcoCodMovim = Format(Documento, "0000000")
                            Else
                                .hcoCodMovim = Documento
                            End If
                            .hcoCodTipoM = lw1.SelectedItem.SubItems(1)
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
                        .Show vbModal
                    End With
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
Dim i As Byte
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
  
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i

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
        CadenaConsulta = "Select distinct codclien from smovalpuntos ORDER BY 1"
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

        If cadB <> "" Then
            'Cadena para el Data1
            CadenaConsulta = "select distinct codclien from smovalpuntos WHERE " & cadB & " "
            

        Else
            'obtener todos los articulos
            CadenaConsulta = "select distinct codclien from smovalpuntos  ORDER BY 1"
            CadenaBusqueda = ""
        End If
        PonerCadenaBusqueda
'    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim i As Byte
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    lblIndicador.Caption = "Obt SQL"
    lblIndicador.Refresh
    Data1.RecordSource = CadenaConsulta


    lblIndicador.Caption = "Refresh"
    lblIndicador.Refresh
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla puntos(smovalpuntos)  para ese criterio de búsqueda", vbInformation
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
Dim i As Integer
Dim Aux As String

On Error GoTo EPonerCampos
 
    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Aux = "puntos"
    'Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sartic", "nomartic")
    Text2(0).Text = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(0).Text, "T", Aux)

    vStock = 0
    If Aux = "" Then Aux = "0"
    vStock = CCur(Aux)
    Text1(1).Text = Format(Aux, FormatoImporte)
    'De salmac
    
    Set Rs = New ADODB.Recordset
    
    
    
    'AHora pongo los datos del list viesw
    Me.Image1(0).visible = False
    Me.Image1(1).visible = False
    
    CargaListView
    
    
    
    
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
    
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
    Set Rs = Nothing
End Sub



Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
            
    cad = cad & "Código|sclien|codclien|T||18·Nombre|sclien|nomclien|T||60·Puntos|sclien|puntos|T||17·"
    tabla = "(smovalpuntos LEFT JOIN sclien ON smovalpuntos.codclien=sclien.codclien" & ") "
    tabla = tabla & " GROUP BY smovalpuntos.codclien"
    Titulo = "Movimientos de puntos"

           
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



Private Sub CargaListView()
Dim cantidad As Currency
Dim Aux As String
Dim IT As ListItem
Dim Total As Currency

    lw1.ListItems.Clear
    Aux = "Select smovalpuntos.*,nomtipom from smovalpuntos left join stipom on smovalpuntos.codtipom=stipom.codtipom where "
    Aux = Aux & "  codclien =" & CStr(Data1.Recordset!codClien)
    Aux = Aux & " order by Fechaalb , fecmov "
    Rs.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Total = 0
    While Not Rs.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = Format(Rs!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(1) = IIf(IsNull(Rs!codtipom), " ", Rs!codtipom)
        If IsNull(Rs!nomtipom) Then
            IT.SubItems(2) = " "
        Else
            IT.SubItems(2) = Rs!nomtipom
        End If
        IT.SubItems(3) = Format(Rs!NumAlbar, "00000")
        
        cantidad = Rs!Puntos
        
        IT.SubItems(4) = Format(cantidad, FormatoCantidad)
        Total = Total + cantidad
        IT.SubItems(5) = Format(Total, FormatoCantidad)
        
  
        'IT.Tag = DBLet(Rs!document)
        Rs.MoveNext
        
    Wend
    Rs.Close
    
    'Si es el mismo importe k el stock
    Me.cmdActualizStock.Tag = Total
    
    Me.Image1(0).visible = Total = vStock
    Me.Image1(1).visible = Not Me.Image1(0).visible
    
End Sub

