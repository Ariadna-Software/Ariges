VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComDtosFamMarca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dtos. Familia/Marca Proveedor"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   13905
   ClipControls    =   0   'False
   Icon            =   "frmComDtosFamMarca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   12120
      MaxLength       =   5
      TabIndex        =   8
      Tag             =   "Sin cargo|N|S|0|99.90|sdtomp|dtosincargo|#0.00|N|"
      Text            =   "Sinca"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   11760
      MaxLength       =   5
      TabIndex        =   7
      Tag             =   "Rappel 2|N|S|0|99.90|sdtomp|rap2|#0.00|N|"
      Text            =   "Dto 2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   11160
      MaxLength       =   5
      TabIndex        =   6
      Tag             =   "Rappel 1|N|S|0|99.90|sdtomp|rap1|#0.00|N|"
      Text            =   "rapp1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   5400
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   22
      ToolTipText     =   "Buscar proveedor"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   3840
      TabIndex        =   21
      ToolTipText     =   "Buscar familia artículo"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   2
      Left            =   6480
      TabIndex        =   20
      ToolTipText     =   "Buscar marca"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   3
      Left            =   9720
      TabIndex        =   19
      ToolTipText     =   "Buscar fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   360
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Cod. Proveedor|N|N|0|999999|sdtomp|codprove|000000|S|"
      Text            =   "provee"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   9960
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Descuento 1|N|N|0|99.90|sdtomp|dtoline1|#0.00|N|"
      Text            =   "Dto 1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   10560
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "Descuento 2|N|N|0|99.90|sdtomp|dtoline2|#0.00|N|"
      Text            =   "Dto 2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   8760
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Fecha Aplicación|F|N|||sdtomp|fechadto|dd/mm/yyyy|N|"
      Text            =   "fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   6000
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Cod. Marca|N|S|0|9999|sdtomp|codmarca|0000|S|"
      Text            =   "codm"
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtaux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   18
      Text            =   "text 2 text 2 text 2 text"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10560
      TabIndex        =   9
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11715
      TabIndex        =   10
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11715
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   16
      Text            =   "Text2 Text2 Text2 Text2 Text2 Text2 Text"
      Top             =   4920
      Width           =   4005
   End
   Begin VB.TextBox txtaux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "Cod. Familia|N|N|0|9999|sdtomp|codfamia|0000|S|"
      Text            =   "cod."
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
            Object.ToolTipText     =   "Generar descuentos"
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
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5040
      Top             =   5400
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
      Bindings        =   "frmComDtosFamMarca.frx":000C
      Height          =   4005
      Left            =   120
      TabIndex        =   25
      Top             =   720
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   7064
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
   Begin VB.Label Label3 
      Caption         =   "Desc. Proveedor"
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
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
      TabIndex        =   13
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
Attribute VB_Name = "frmComDtosFamMarca"
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

Private WithEvents frmP As frmComProveedores  'Form Mantenimiento Proveedores
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmFam As frmAlmFamiliaArticulo  'Form Mantenimiento Familias Articulos
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmM As frmAlmMarcas  'Form Mantenimiento Marcas
Attribute frmM.VB_VarHelpID = -1



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



Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
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
        If DatosOk Then
            If InsertarDesdeForm(Me) Then
                CargaGrid True
                BotonAnyadir
            End If
        End If
    Case 4 'MODIFICAR
        If DatosOk And BLOQUEADesdeFormulario(Me) Then
             If ModificaAqui Then
                 TerminaBloquear
                 NumReg = Data1.Recordset.AbsolutePosition
                 PonerModo 2
                 CancelaADODC Me.Data1
                 CargaGrid True
                 LLamaLineas 10
                 SituarDataPosicion Data1, NumReg, Indicador
             End If
             lblIndicador.Caption = Indicador
             DataGrid1.SetFocus
         End If
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Cod Proveedor
            Set frmP = New frmComProveedores
            frmP.DatosADevolverBusqueda = "0"
            frmP.Show vbModal
            Set frmP = Nothing
        Case 1 'Cod Familia
            Set frmFam = New frmAlmFamiliaArticulo
            frmFam.DatosADevolverBusqueda = "0"
            frmFam.Show vbModal
            Set frmFam = Nothing
        Case 2 'Cod Marca
            Set frmM = New frmAlmMarcas
            frmM.DatosADevolverBusqueda = "0"
            frmM.Show vbModal
            Set frmM = Nothing
        Case 3 'Fecha aplicacion
            Screen.MousePointer = vbHourglass
            Set frmF = New frmCal
            frmF.Fecha = Now
            If txtAux(Index).Text <> "" Then frmF.Fecha = CDate(txtAux(Index).Text)
            Screen.MousePointer = vbDefault
            frmF.Show vbModal
            Set frmF = Nothing
    End Select
    PonerFoco txtAux(Index)
End Sub


Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
            
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            
            If Not Data1.Recordset.EOF Then
                PonerModo 2
                Data1.Recordset.MoveFirst
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else 'No hay Registros en la Tabla
                PonerModo 0
            End If
            LLamaLineas 10
            
        Case 4  'MODIFICAR
            TerminaBloquear
            DeseleccionaGrid Me.DataGrid1
'            CargaTxtAux False, False
            PonerModo 2
            LLamaLineas 10
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Not Data1.Recordset.EOF And (Modo <> 3 And Modo <> 4) Then
        'Poner descripcion del Cliente
        Text2(0).Text = DevuelveDesdeBDNew(1, "sprove", "nomprove", "codprove", CStr(Data1.Recordset!Codprove), "N")
        
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de laLa toolbar
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 42 'Generar
        .Buttons(12).Image = 16 'Imprimir
        .Buttons(13).Image = 15 'Salir
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    PonerModo 0
    
    NombreTabla = "sdtomp" 'Tabla Descuentos Familia/Marca Proveedor
    Ordenacion = " ORDER BY codprove, codfamia, codmarca "
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codprove = -1" 'No recupera datos
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    CargaGrid (Modo = 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim B As Boolean
Dim SQL As String
On Error GoTo ECarga

    B = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, False

    CargaGrid2

    DataGrid1.Enabled = B
    DataGrid1.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


Private Sub CargaGrid2()
Dim tots As String
On Error GoTo ECarga2

    '"SELECT codprove, " & NombreTabla & ".codfamia, nomfamia, " & NombreTabla & ".codmarca, nommarca, fechadto, dtoline1, dtoline2 "
    tots = "S|txtAux(0)|T|Prove.|850|;S|cmdAux(0)|B||0|;S|txtAux(1)|T|Fami.|650|;S|cmdAux(1)|B||0|;S|txtaux2(1)|T|Desc. Familia|2600|;"
    tots = tots & "S|txtAux(2)|T|Marca|680|;S|cmdAux(2)|B||0|;S|txtaux2(2)|T|Desc. Marca|2600|;"
    tots = tots & "S|txtAux(3)|T|Fecha dto.|1150|;S|cmdAux(3)|B||0|;S|txtAux(4)|T|Dto 1|800|;S|txtAux(5)|T|Dto 2|800|;"
    'Sept 2010
    tots = tots & "S|txtAux(6)|T|Rap 1|800|;S|txtAux(7)|T|Rap 2|800|;"
    'Ago 2016
    tots = tots & "S|txtAux(8)|T|S/Cargo|850|;"
    arregla tots, DataGrid1, Me

    'dtos alineados a la dcha
    DataGrid1.Columns(6).Alignment = dbgRight
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(8).Alignment = dbgRight
    DataGrid1.Columns(9).Alignment = dbgRight
    DataGrid1.Columns(10).Alignment = dbgRight
    
ECarga2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
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
Dim indice As Byte
    indice = 3
    txtAux(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Familias
    txtAux(1).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtaux2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmM_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento MARCAS
    txtAux(2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtaux2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento Proveedores
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
   BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
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
            
        Case 10
            CadenaDesdeOtroForm = ""
            frmListado3.Opcion = 33
            frmListado3.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                'CadenaDesdeOtroForm
                
                CadenaBusqueda = "WHERE " & CadenaDesdeOtroForm
                CadenaBusqueda = Replace(CadenaBusqueda, "codfamia =", "sdtomp.codfamia =", , , vbTextCompare)
                CadenaBusqueda = Replace(CadenaBusqueda, "codprove ", "sdtomp.codprove ", , , vbTextCompare)
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadenaDesdeOtroForm & Ordenacion
                
                PonerCadenaBusqueda
                
            End If
        
        Case 12 'Imprimir
            BotonImprimir
        Case 13  'Salir
            mnSalir_Click
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
      
    Select Case Kmodo
        Case 1 'Modo Buscar
            PonerFoco txtAux(0)
            txtAux(0).BackColor = vbYellow
        Case 2    'Preparamos para que pueda Modificar
            Me.cmdRegresar.visible = False
    End Select
           
     'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = b
'        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
'    Else
'        cmdRegresar.visible = False
'    End If
                 
    
    '-----------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    Me.DataGrid1.Enabled = (Modo <> 3 And Modo <> 4)
    BloquearClavesP (Modo = 4) ' si modificar


    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
    
    B = (Modo = 2)
    'Insertar
    Toolbar1.Buttons(5).Enabled = (B Or (Modo = 0))
    Me.mnNuevo.Enabled = (B Or (Modo = 0))
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnEliminar.Enabled = B
    
    
    B = (Modo >= 3) Or Modo = 1
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
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
    
    SQL = "SELECT " & NombreTabla & ".codprove, " & NombreTabla & ".codfamia, nomfamia, " & NombreTabla & ".codmarca, nommarca, "
    SQL = SQL & "fechadto, dtoline1, dtoline2, rap1, rap2,dtosincargo "
    SQL = SQL & " FROM (" & NombreTabla & " LEFT JOIN sfamia ON " & NombreTabla & ".codfamia =sfamia.codfamia) "
    SQL = SQL & " LEFT JOIN smarca ON " & NombreTabla & ".codmarca=smarca.codmarca "

    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
    Else
        SQL = SQL & " WHERE sdtomp.codprove = -1"
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
'        CargaTxtAux True, True
        anc = ObtenerAlto(Me.DataGrid1)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            txtAux(kCampo).Text = ""
            txtAux(kCampo).BackColor = vbYellow
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
'Ver todos
    EsBusqueda = False
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
'    CargaGrid False
    
'    If chkVistaPrevia.Value = 1 Then
'        MandaBusquedaPrevia ""
'    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
'    End If
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
       
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
    
'    CargaTxtAux True, True
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc
    
    PonerFoco txtAux(0)
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
    
'    CargaTxtAux True, False
    anc = ObtenerAlto(Me.DataGrid1)
    LLamaLineas anc
    
    '---- poner valores grabados
    'prove y familia
    For i = 0 To 1
        txtAux(i).Text = DBLet(DataGrid1.Columns(i).Value, "N")
        FormateaCampo txtAux(i)
    Next i
    txtaux2(1).Text = DBLet(Me.DataGrid1.Columns(2).Value, "T")
    
    'Marca
    txtAux(2).Text = DBLet(Data1.Recordset!codmarca, "T")
    FormateaCampo txtAux(2)
    txtaux2(2).Text = DBLet(Data1.Recordset!nommarca, "T")
    
    txtAux(3).Text = DBLet(Me.DataGrid1.Columns(5).Value, "F")
    For i = 4 To 8
        txtAux(i).Text = DBLet(DataGrid1.Columns(i + 2), "N")
        FormateaCampo txtAux(i)
    Next i
    '-----
    If BLOQUEADesdeFormulario(Me) Then
        PonerFoco txtAux(3)
    Else
        cmdCancelar_Click
    End If
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        'Ciertas comprobaciones
        If Data1.Recordset.EOF Then Exit Function
        
        SQL = "¿Seguro que desea eliminar el Descuento para?" & vbCrLf
        SQL = SQL & vbCrLf & "Proveedor: " & Format(Data1.Recordset.Fields(0).Value, "000000") & " - " & Text2(0).Text
        SQL = SQL & vbCrLf & "Familia: " & Format(Data1.Recordset.Fields(1).Value, "0000") & " - " & Data1.Recordset!nomfamia
        SQL = SQL & vbCrLf & "Marca : " & Format(Data1.Recordset.Fields(3).Value, "0000") & " - " & Data1.Recordset!nommarca
        
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            NumRegElim = Me.Data1.Recordset.AbsolutePosition
            SQL = "Delete from " & NombreTabla & " where codprove=" & Val(Data1.Recordset!Codprove)
            SQL = SQL & " and codfamia=" & Val(Data1.Recordset!Codfamia) & " and codmarca"
            If IsNull(Data1.Recordset!codmarca) Then
                SQL = SQL & " is null "
            Else
                SQL = SQL & " = " & Val(Data1.Recordset!codmarca)
            End If
            conn.Execute SQL
            CancelaADODC Me.Data1
            CargaGrid True
            CancelaADODC Me.Data1
            SituarDataPosicion Me.Data1, NumRegElim, SQL
        End If
        
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Descuento", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me, 3)
    If Not B Then Exit Function
    
    
    If txtAux(0).Text <> "" Then
        If Text2(0).Text = "" Then
            MsgBox "Error en proveedor", vbExclamation
            PonerFoco txtAux(0)
            Exit Function
        End If
    End If
    
    If txtAux(1).Text <> "" Then
        If txtaux2(1).Text = "" Then
            MsgBox "Error en familia", vbExclamation
            PonerFoco txtAux(1)
            Exit Function
        End If
    End If
    'NUEVO Sept 2010
    'Si esta OK
    'comprobaremos si  ya existe. Ya que hemmos tenido que crear la clave
    If Modo = 3 Then
        If txtAux(2).Text = "" Then
            CadenaConsulta = " is null"
        Else
            'Si pone la marca, DEBE existir
            If txtaux2(2).Text = "" Then
                MsgBox "No existe la marca: " & txtAux(2).Text, vbExclamation
                Exit Function
            End If
        
        
            CadenaConsulta = " = " & txtAux(2).Text
        End If
        CadenaConsulta = "codmarca " & CadenaConsulta
        CadenaConsulta = CadenaConsulta & " and codfamia = " & txtAux(1).Text & " and codprove "
        CadenaConsulta = DevuelveDesdeBD(conAri, "codprove", "sdtomp", CadenaConsulta, txtAux(0).Text, "N")
        If CadenaConsulta <> "" Then
            MsgBox "Ya ha introducido un descuento para proveedor/familia/marca con estos valores", vbExclamation
            Exit Function
        End If
        
    
    End If
    
    DatosOk = True
End Function



Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If cadB <> "" Then 'Se muestran en el mismo form
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
        PonerModo Modo
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
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

Private Sub BotonImprimir()
    frmListado2.Opcion = 33
    frmListado2.Show vbModal

End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub

Private Sub BloquearClavesP(bol As Boolean)
'Si BloquearClavesPrimarias=true deshablilita los textbox de codigos y lo pone amarillo
'y habilita el resto de campos para introducir nuevos valores
'Si BloquearClavesPrimarias=false habilita los textbox de codigos para introducir
Dim i As Byte

    For i = 0 To 2 'Codigos
        BloquearTxt txtAux(i), bol
        Me.cmdAux(i).Enabled = Not bol
    Next i
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim campo As String, tabla As String
  
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'cod. Proveedor
            If PonerFormatoEntero(txtAux(Index)) Then
                campo = "nomprove"
                tabla = "sprove"
                Text2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, tabla, campo)
            Else
                Text2(Index).Text = ""
            End If
        Case 1 'Cod. Familia
            If PonerFormatoEntero(txtAux(Index)) Then
                campo = "nomfamia"
                tabla = "sfamia"
                txtaux2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, tabla, campo)
            Else
                txtaux2(Index).Text = ""
            End If
            
        Case 2 'Cod. Marca
            If PonerFormatoEntero(txtAux(Index)) Then
                campo = "nommarca"
                tabla = "smarca"
                txtaux2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, tabla, campo)
            Else
                txtaux2(Index).Text = ""
            End If
            
        Case 3 'Fecha Descuento
            PonerFormatoFecha txtAux(Index)
        Case 4 To 8 'Descuentos
            'Formato tipo 4: Decimal(4,2)
            If Not PonerFormatoDecimal(txtAux(Index), 4) Then txtAux(Index).Text = ""
    End Select
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim B As Boolean

        DeseleccionaGrid Me.DataGrid1
        'PonerModo xModo + 1

        B = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

        For jj = 0 To txtAux.Count - 1
            txtAux(jj).Height = DataGrid1.RowHeight
            txtAux(jj).Top = alto
            txtAux(jj).visible = B
        Next jj
        
        For jj = 1 To txtaux2.Count
            txtaux2(jj).Height = Me.DataGrid1.RowHeight
            txtaux2(jj).Top = alto
            txtaux2(jj).visible = B
        Next jj
        
        For jj = 0 To Me.cmdAux.Count - 1
            Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
            Me.cmdAux(jj).Top = alto
            Me.cmdAux(jj).visible = B
        Next jj
End Sub



Private Function ModificaAqui() As Boolean
Dim SQL As String

On Error GoTo eModificaAqui
    ModificaAqui = False

    SQL = "update `sdtomp` set fechadto=" & DBSet(txtAux(3).Text, "F")
    SQL = SQL & ", dtoline1=" & DBSet(txtAux(4).Text, "N")
    SQL = SQL & ", dtoline2=" & DBSet(txtAux(5).Text, "N")
    SQL = SQL & ", rap1=" & DBSet(txtAux(6).Text, "N", "S")
    SQL = SQL & ", rap2=" & DBSet(txtAux(7).Text, "N", "S")
    SQL = SQL & ", dtosincargo=" & DBSet(txtAux(8).Text, "N", "S")
    SQL = SQL & " where `codprove`=" & txtAux(0).Text & " and codfamia= " & txtAux(1).Text & " AND codmarca "
    If txtAux(2).Text = "" Then
        SQL = SQL & " is null"
    Else
        SQL = SQL & " = " & txtAux(2).Text
    End If
    conn.Execute SQL
    ModificaAqui = True

    
eModificaAqui:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function
