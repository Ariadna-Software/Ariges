VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmNumLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Números de Lote"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   12315
   ClipControls    =   0   'False
   Icon            =   "frmAlmNumLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   9480
      MaxLength       =   16
      TabIndex        =   4
      Tag             =   "Vendida|N|N|||slotes|vendida|#,###,###,##0.00|N|"
      Text            =   "Vendida"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   8160
      TabIndex        =   16
      ToolTipText     =   "Buscar artículo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   8280
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Cantidad|N|N|||slotes|canentra|#,###,###,##0.00|N|"
      Text            =   "cantidad"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha entrada|F|N|||slotes|fecentra||S|"
      Text            =   "fecha"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   5640
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "Num. Lotes|T|N|||slotes|numlotes||S|"
      Text            =   "numlote"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   5160
      Width           =   2535
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
         TabIndex        =   15
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   13
      ToolTipText     =   "Buscar artículo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   480
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Cod. Artículo|T|N|||slotes|codartic||S|"
      Text            =   "codartic codarti"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9960
      TabIndex        =   5
      Top             =   5325
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11115
      TabIndex        =   6
      Top             =   5325
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11115
      TabIndex        =   7
      Top             =   5325
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12315
      _ExtentX        =   21722
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
            Object.ToolTipText     =   "Comprobacion"
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
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   8160
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
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
      Bindings        =   "frmAlmNumLote.frx":000C
      Height          =   4425
      Left            =   240
      TabIndex        =   8
      Top             =   705
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7805
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
      TabIndex        =   10
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
   Begin VB.Menu mnFiltro 
      Caption         =   "Filtro"
      Begin VB.Menu mnFiltro1 
         Caption         =   "Sin filtro"
         Index           =   0
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Disponible"
         Index           =   2
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Ultimo mes"
         Index           =   3
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Ultimos 6 meses"
         Index           =   4
      End
      Begin VB.Menu mnFiltro1 
         Caption         =   "Ultimo año"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmAlmNumLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticu2 'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer

Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As Boolean

Dim KK As Integer


Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long

    
    On Error GoTo Error1
    
    
    'FALTA#####
    ' LOG. si cambian cantidad vendida
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    EsBusqueda = True
                    CadenaBusqueda = " WHERE  slotes.codartic=" & DBSet(txtAux(0).Text, "T") & " AND numlotes=" & DBSet(txtAux(1).Text, "T")
                    CargaGrid True
                    BotonAnyadir
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                 If ModificaDesdeFormulario(Me, 3) Then
                     'LOG
                     If vParamAplic.ManipuladorFitosanitarios2 Then
                        
                        Indicador = "Articulo: " & Data1.Recordset!codArtic & " -- " & Data1.Recordset!NomArtic
                        Indicador = Indicador & vbCrLf & "Anterior: " & Data1.Recordset!vendida
                        Indicador = Indicador & "     ACtual: " & Me.txtAux(4).Text
                        
                        Set LOG = New cLOG
                        LOG.Insertar 35, vUsu, Indicador
                        Set LOG = Nothing
                     End If
                 
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

    Select Case Index
        Case 0 'Cod Articulo
            Set frmA = New frmAlmArticu2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Poner Modo Busqueda
            frmA.DesdeTPV = False
            frmA.Show vbModal
            Set frmA = Nothing
            PonerFoco txtAux(0)
        
        Case 1 'fecha entrada
            Set frmF = New frmCal
            frmF.Fecha = Now
            If txtAux(2).Text <> "" Then frmF.Fecha = CDate(txtAux(2).Text)
            Screen.MousePointer = vbDefault
            frmF.Show vbModal
            Set frmF = Nothing
            PonerFoco txtAux(2)
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
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Data1.Recordset.Fields(0) & "|"
    Cad = Cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
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
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(8).visible = vUsu.Login = "root"
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 15 'Salir
    End With

    LimpiarCampos   'Limpia los campos TextBox
   
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "slotes" 'Tabla numeros de lotes
    Ordenacion = " ORDER BY codartic,fecentra,numlotes "
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE numlotes = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
'    CargaGrid (Modo = 2 Or Modo = 0)
    CargaGrid False
    
    KK = ByteValueLeer(Me.Name)
    mnFiltro.Tag = KK
    For kCampo = 0 To Me.mnFiltro1.Count - 1
        If KK = kCampo Then
            mnFiltro1(kCampo).Checked = True
            Exit For
        End If
    Next
    
    
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String
    
    On Error GoTo ECarga
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, False
    
    tots = "S|txtAux(0)|T|Artic.|1600|;S|cmdAux(0)|B||0|;S|txtAux2(0)|T|Nombre artículo|4100|;S|txtAux(1)|T|Nº Lote|1700|;S|txtAux(2)|T|Fecha ent.|1200|;S|cmdAux(1)|B||0|;S|txtAux(3)|T|Cantidad|1200|;S|txtAux(4)|T|Vendida|1200|;"
    arregla tots, DataGrid1, Me


'    'dtos alineados a la dcha
    DataGrid1.Columns(3).Alignment = dbgCenter
'    DataGrid1.Columns(6).Alignment = dbgCenter

    DataGrid1.ScrollBars = dbgAutomatic
    
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
Dim b As Boolean

    DeseleccionaGrid Me.DataGrid1
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To txtAux.Count - 1
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
    Next jj
    txtAux2(0).Height = Me.DataGrid1.RowHeight
    txtAux2(0).Top = alto
    txtAux2(0).visible = b
    
    
    For jj = 0 To Me.cmdAux.Count - 1
        Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(jj).Top = alto
        Me.cmdAux(jj).visible = b
    Next jj
End Sub



Private Sub Form_Unload(Cancel As Integer)
    KK = 0
    For kCampo = 0 To Me.mnFiltro1.Count - 1
        If mnFiltro1(kCampo).Checked Then
            KK = kCampo
            Exit For
        End If
    Next
    If mnFiltro.Tag <> KK Then ByteValueGuardar Me.Name, CByte(KK)
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2)
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
        Aux = ValorDevueltoFormGrid(txtAux(0), CadenaDevuelta, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(txtAux(1), CadenaDevuelta, 2)
        cadB = cadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtAux(2).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnFiltro1_Click(Index As Integer)
    For KK = 0 To Me.mnFiltro1.Count - 1
        If Index <> 1 Then
            If KK = Index Then
                Me.mnFiltro1(KK).Checked = True
            Else
                Me.mnFiltro1(KK).Checked = False
            End If
        End If
    Next
End Sub

Private Sub mnModificar_Click()
    BotonModificar
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
            
        Case 8
            If vUsu.Login <> "root" Then Exit Sub
            HacerComprobacion
        
        Case 10 'Imprimir
'            BotonImprimir
            
        Case 11  'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
     'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
                      
    If Kmodo = 1 Then 'Modo Buscar
        PonerFoco txtAux(0)
    End If
                                 
    BloquearTxt txtAux(0), (Modo = 4)
    Me.cmdAux(0).Enabled = (Modo <> 4)
                      
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b

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
Dim b As Boolean

    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    b = ((Modo >= 3))
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
    
    SQL = "SELECT " & NombreTabla & ".codartic, " & " Articulos.nomartic, numlotes, fecentra, canentra, vendida "
    SQL = SQL & " FROM " & NombreTabla & " LEFT OUTER JOIN sartic AS Articulos ON " & NombreTabla & ".codartic ="
    SQL = SQL & " Articulos.codartic"
    
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
    Else
        SQL = SQL & " WHERE numlotes = -1"
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
        anc = ObtenerAlto(Me.DataGrid1, 10)
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

Private Function ValorFiltro() As String
    ValorFiltro = ""
    If Me.mnFiltro1(2).Checked Then
        ValorFiltro = " canentra - vendida >0 "
    ElseIf Me.mnFiltro1(3).Checked Then
        MsgBox "Falt"
    End If
End Function

Private Sub BotonVerTodos()
    On Error Resume Next

    EsBusqueda = False
    LimpiarCampos
    
    CadenaConsulta = ValorFiltro
    If CadenaConsulta <> "" Then CadenaConsulta = " WHERE " & CadenaConsulta
    
    
    CadenaConsulta = "Select * from " & NombreTabla & CadenaConsulta
    CadenaConsulta = CadenaConsulta & Ordenacion
    PonerCadenaBusqueda
    PonerFocoGrid DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
    
    'fecha de entrada
    txtAux(2).Text = Format(Now, "dd/mm/yyyy")
    
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim I As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    
    'codartic
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "T")
    txtAux2(0).Text = DBLet(DataGrid1.Columns(1).Value, "T")
    
    'numlote
    txtAux(1).Text = DBLet(Me.DataGrid1.Columns(2).Value, "T")
    'fecha entrada
    txtAux(2).Text = DBLet(DataGrid1.Columns(3).Value, "F")
    FormateaCampo txtAux(2)
    
    'cantidad
    txtAux(3).Text = DBLet(DataGrid1.Columns(4).Value, "N")
    FormateaCampo txtAux(3)
    
    'cantidad
    txtAux(4).Text = DBLet(DataGrid1.Columns(5).Value, "N")
    FormateaCampo txtAux(4)
    
    

    DataGrid1.Enabled = False
    PonerFoco txtAux(1)
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar
    
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Function
    
    If Not PuedeRealizarLaAccion Then Exit Function
    
    SQL = "¿Seguro que desea eliminar el Nº de lote?" & vbCrLf
    SQL = SQL & vbCrLf & "Nº Lote: " & Data1.Recordset.Fields(2).Value
    SQL = SQL & vbCrLf & "Artículo: " & Data1.Recordset.Fields(1).Value
            
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Data1.Recordset.AbsolutePosition
        SQL = "Delete from " & NombreTabla & " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data1.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data1.Recordset!fecentra, "F")
        
        conn.Execute SQL
        CancelaADODC Me.Data1
        CargaGrid True
        CancelaADODC Me.Data1
        SituarDataPosicion Me.Data1, NumRegElim, SQL
    End If
        
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Nº de lote", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cart As CArticulo

    On Error GoTo ErrDatosOK

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    
    'comprobar que el articulo insertado tiene control de numero de serie
    If Modo = 3 Then
        Set cart = New CArticulo
        If cart.LeerDatos(txtAux(0).Text) Then
            If Not cart.TieneNumLote Then
                b = False
                MsgBox "El artículo no tiene control de nº de lote.", vbInformation
            End If
        End If
        Set cart = Nothing
    End If
    
    
    If b Then
        If ImporteFormateado(txtAux(4).Text) > ImporteFormateado(txtAux(3).Text) Then
            MsgBox "Cantidad vendida no puede ser mayor a cantidad entrada", vbExclamation
            PonerFoco txtAux(4)
            b = False
        End If
    End If
    
    If Modo = 4 Then
        If b Then b = PuedeRealizarLaAccion
    End If
    
    DatosOk = b
    Exit Function
    
ErrDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos OK.", Err.Description
End Function



Private Sub MandaBusquedaPrevia(cadB As String)
''Carga el formulario frmBuscaGrid con los valores correspondientes
'Dim cad As String
'Dim Tabla As String
'Dim Titulo As String
'
'    'Llamamos a al form
'    cad = ""
'    'Estamos en Modo de Cabeceras
'    'Registro de la tabla de cabeceras: slista
'    cad = cad & ParaGrid(txtAux(0), 20, "Código")
'    cad = cad & ParaGrid(txtAux(1), 80, "Descripción")
'    Tabla = NombreTabla
'    Titulo = "Tipos de Contrato"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = Tabla
'        frmB.vSQL = cadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = Titulo
'        frmB.vselElem = 1
'        frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
''        Else   'de ha devuelto datos, es decir NO ha devuelto datos
''            If Modo = 5 Then
''                PonerFoco txtAux(0)
''            Else
'                PonerFoco txtAux(kCampo)
''            End If
'        End If
'    End If
'    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
Dim Aux As String
    
    cadB = ObtenerBusqueda(Me, False)
    Aux = ValorFiltro
    If Aux <> "" Then Aux = " AND " & Aux
    
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
        CadenaBusqueda = " WHERE " & cadB
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
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


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
   KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    On Error Resume Next
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'Cod. Articulo
            txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), conAri, "sartic", "nomartic")
            If txtAux2(0).Text = "" And txtAux(0).Text <> "" Then PonerFoco txtAux(0)
            
        Case 2 'fecha
            PonerFormatoFecha txtAux(Index)
            
        Case 3, 4 'cantidad entrada
            PonerFormatoDecimal txtAux(Index), 1
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function PuedeRealizarLaAccion() As Boolean
Dim SQL As String

    PuedeRealizarLaAccion = True
    If Not vParamAplic.ManipuladorFitosanitarios2 Then Exit Function
            
            
    SQL = "OK"
    If Modo = 4 Then
        SQL = ""
        'Si cambia cantidad vendida o nº lote entonces comprobaremos
        If Data1.Recordset!numlotes <> Me.txtAux(1).Text Then
            SQL = "OK"
        ElseIf Data1.Recordset!vendida <> ImporteFormateado(Me.txtAux(4).Text) Then
            'SQL = "OK"
        End If
    Else
        SQL = "OK"
    End If
        
    PuedeRealizarLaAccion = False
    If SQL <> "" Then
        SQL = "numlote =" & DBSet(Data1.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(Data1.Recordset!fecentra, "F")
        SQL = SQL & " AND codartic = " & DBSet(Data1.Recordset!codArtic, "T") & " AND 1"
        SQL = DevuelveDesdeBD(conAri, "count(*)", "slialblotes", SQL, "1")
        If Val(SQL) > 0 Then
            MsgBox "El lote ya ha sido vendido", vbExclamation
            Exit Function
        End If
        
        SQL = "numlote =" & DBSet(Data1.Recordset!numlotes, "T") & " AND fecentra = " & DBSet(Data1.Recordset!fecentra, "F")
        SQL = SQL & " AND codartic = " & DBSet(Data1.Recordset!codArtic, "T") & " AND 1"
        SQL = DevuelveDesdeBD(conAri, "count(*)", "slivenlotes", SQL, "1")
        If Val(SQL) > 0 Then
            MsgBox "El lote esta siendo vendido", vbExclamation
            Exit Function
        End If
    End If
    PuedeRealizarLaAccion = True
    
End Function




Private Sub HacerComprobacion()
Dim Cad As String
Dim cantidad As Currency
Dim ColArtic As Collection
Dim I As Integer
Dim EnElLote As Currency
Dim NF As Integer
Dim UpdateaSlotes As Boolean


    conn.Execute "DELETE FROM tmpstockfec WHERE codusu =" & vUsu.codigo
    
    lblIndicador.Caption = "Ajuste devoluciones"
    lblIndicador.Refresh
    Cad = "UPDATE slotes set vendida=canentra where canentra < 0"
    conn.Execute Cad
    Espera 0.5
    
    
    
    lblIndicador.Caption = "Leyendo lotes"
    lblIndicador.Refresh
    Cad = "insert into tmpstockfec(codusu,codartic,codalmac,stock)"
    Cad = Cad & " select " & vUsu.codigo & ",codartic,1,sum(canentra-vendida) from slotes group by codartic"
    conn.Execute Cad
     
     
    
      
     
    lblIndicador.Caption = "Leyendo salmac"
    lblIndicador.Refresh
    Espera 0.5
    
    Set miRsAux = New ADODB.Recordset
    Cad = "Select codartic, sum(canstock) canti from salmac where codartic in (select codartic from tmpstockfec where codusu =" & vUsu.codigo & ") GROUP BY codartic"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColArtic = New Collection
    NF = FreeFile
    Open App.Path & "\ajslot.txt" For Output As #NF
    While Not miRsAux.EOF
        
        Cad = DevuelveDesdeBD(conAri, "stock", "tmpstockfec", "codusu = " & vUsu.codigo & " AND codartic", miRsAux!codArtic, "T")
        cantidad = CCur(Cad)
        
        If miRsAux!canti = cantidad Then
            Cad = "DELETE FROM tmpstockfec WHERE codusu =" & vUsu.codigo & " AND codartic = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute Cad
        Else
        
            cantidad = miRsAux!canti
        
            Cad = miRsAux!codArtic & "|" & cantidad & "|"
            ColArtic.Add Cad
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    For I = 1 To ColArtic.Count
        lblIndicador.Caption = I & " de " & ColArtic.Count
        lblIndicador.Refresh
        
        
        Cad = RecuperaValor(ColArtic.item(I), 2)
        cantidad = CCur(Cad)   'canstock
        
        
        Cad = RecuperaValor(ColArtic.item(I), 1)
        Debug.Print Cad
        
        
        If cantidad < 0 Then
            'STOCK NEGATIVO
            
            'Cantidad negativa
            Print #NF, RecuperaValor(ColArtic.item(I), 1) & "::" & RecuperaValor(ColArtic.item(I), 2) & "::" & cantidad & "      NEGATIVO"
            
        Else
            Cad = "Select * from slotes WHERE  codartic=" & DBSet(Cad, "T") & " AND canentra >0 order by fecentra desc"
            miRsAux.Open Cad, conn, adOpenKeyset, adLockReadOnly, adCmdText
            Cad = ""
            While Not miRsAux.EOF
               
               
                If miRsAux!codArtic = "2020903032012" Then Stop
               
                EnElLote = 0
                UpdateaSlotes = False
                If cantidad < 0 Then
                    'Significa que hay mas en lote s que en stock. Nos fiamos del stock
                    
                    
                    
                    
                Else
                    
                        If miRsAux!canentra > cantidad Then
                            'Perfecto. TOOOOdos va a este lote
                            EnElLote = miRsAux!canentra - cantidad
                            cantidad = 0
                            UpdateaSlotes = True
                        Else
                            
                            EnElLote = 0
                            cantidad = cantidad - miRsAux!canentra
                            UpdateaSlotes = True
                        End If
                                 
                End If
                
                
                If UpdateaSlotes Then
                    Cad = "UPDATE slotes set vendida=" & DBSet(EnElLote, "N") & " WHERE codartic =" & DBSet(miRsAux!codArtic, "T")
                    Cad = Cad & " AND numlotes=" & DBSet(miRsAux!numlotes, "T") & " AND fecentra=" & DBSet(miRsAux!fecentra, "F")
                    conn.Execute Cad
                End If
                miRsAux.MoveNext
                If cantidad = 0 Then
                    'HA IDO TODO BIEN
                    While Not miRsAux.EOF
                        Cad = "UPDATE slotes set vendida=canentra WHERE codartic =" & DBSet(miRsAux!codArtic, "T")
                        Cad = Cad & " AND numlotes=" & DBSet(miRsAux!numlotes, "T") & " AND fecentra=" & DBSet(miRsAux!fecentra, "F")
                        conn.Execute Cad
                        miRsAux.MoveNext
                    Wend
                End If
            Wend
            miRsAux.Close
            
            If cantidad <> 0 Then
                Print #NF, RecuperaValor(ColArtic.item(I), 1) & "::" & RecuperaValor(ColArtic.item(I), 2) & "::" & cantidad
            Else
                If Cad = "" Then Print #NF, RecuperaValor(ColArtic.item(I), 1) & "::" & RecuperaValor(ColArtic.item(I), 2) & " 0STOCK"
            End If
        
        End If
    Next
    Close #NF
    
    lblIndicador.Caption = ""
    Set miRsAux = Nothing
    Set ColArtic = Nothing
End Sub
