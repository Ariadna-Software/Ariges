VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmMarcasGr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Marcas"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   6525
   Icon            =   "frmAlmMarcasGr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   7
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   8
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
      Height          =   330
      Index           =   0
      Left            =   120
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "C�digo Marca|N|N|0|9999|smarca|codmarca|0000|S|"
      Text            =   "Dat"
      Top             =   5040
      Width           =   800
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
      Height          =   330
      Index           =   1
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "Marca|T|N|||smarca|nommarca||N|"
      Text            =   "Dato2"
      Top             =   5040
      Width           =   1395
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
      Left            =   5310
      TabIndex        =   3
      Top             =   7020
      Width           =   1065
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
      Left            =   4065
      TabIndex        =   2
      Top             =   7005
      Width           =   1065
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
      Left            =   5310
      TabIndex        =   6
      Top             =   7020
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   135
      TabIndex        =   4
      Top             =   6870
      Width           =   1755
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
         Height          =   255
         Left            =   225
         TabIndex        =   5
         Top             =   180
         Width           =   1200
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   495
      Left            =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
      Bindings        =   "frmAlmMarcasGr.frx":000C
      Height          =   5985
      Left            =   135
      TabIndex        =   9
      Top             =   810
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   10557
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
         Caption         =   "&Ver todos"
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
Attribute VB_Name = "frmAlmMarcasGr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)

Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos

Dim FormatoCod As String 'formato del campo de codigo
Dim Modo As Byte
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas


Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
    
    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b

    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    End If

    'Si estamos insertando o busqueda
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu seg�n Modo
    'PonerModoUsuarioGnral Modo, "ariagro"
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
   
    b = b And Not DeConsulta
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    Toolbar1.Buttons(10).Enabled = b
End Sub

'Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
'Dim Rs As ADODB.Recordset
'Dim cad As String
'Dim b As Boolean
'Dim bAux As Boolean
'
'    On Error Resume Next
'
'    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
'    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.ID, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Not Rs.EOF Then
'        b = (Modo = 2 And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
'
'        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2) And Not DeConsulta
'        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And b
'        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And b
'
'        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 2)
'        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 2)
'
'        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And b
'
'    End If
'
'    Rs.Close
'    Set Rs = Nothing
'
'End Sub



Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub BotonAnyadir()
Dim anc As Single
    
    'Situamos el grid al final
    
    AnyadirLinea DataGrid1, adodc1
      
    anc = ObtenerAlto(DataGrid1, 10)
'    anc = DataGrid1.Top
'    If DataGrid1.Row < 0 Then
'        anc = anc + 240
'    Else
'        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
'    End If
    'Obtenemos la siguiente numero de c�digo de Marca
    txtAux(0).Text = SugerirCodigoSiguienteStr("smarca", "codmarca")
    FormateaCampo txtAux(0)
    txtAux(1).Text = ""
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid "codmarca= -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""

    LLamaLineas DataGrid1.Top + 240, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid ""
    If adodc1.Recordset.RecordCount <= 0 Then
         'MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
         MsgBox "No hay ning�n registro en la tabla smarca", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
Dim cad As String
Dim anc As Single
Dim i As Integer
On Error GoTo EModificar

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    'El registro de codigo 0 no se puede Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    anc = ObtenerAlto(DataGrid1, 10)
    
    cad = ""
    For i = 0 To 1
        cad = cad & DataGrid1.Columns(i).Text & "|"
    Next i
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    LLamaLineas anc, 4
   
    Screen.MousePointer = vbDefault
EModificar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar", Err.Description
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(0).Left = DataGrid1.Left + 340
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
End Sub


Private Sub BotonEliminar()
Dim SQL As String
On Error GoTo Error2

    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
    
    '### a mano
    SQL = "�Seguro que desea eliminar la Marca?" & vbCrLf
    SQL = SQL & vbCrLf & "C�digo: " & Format(adodc1.Recordset.Fields(0), FormatoCod)
    SQL = SQL & vbCrLf & "Denominaci�n: " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
        SQL = "Delete from smarca where codmarca=" & adodc1.Recordset!codmarca
        conn.Execute SQL
        CancelaADODC adodc1
        CargaGrid ""
        CancelaADODC Me.adodc1
        SituarDataPosicion Me.adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Marca", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String
On Error Resume Next

    Select Case Modo
        Case 3   'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                If ModificaDesdeFormulario(Me, 3) Then
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CancelaADODC Me.adodc1
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                End If
                DataGrid1.SetFocus
            End If
            
        Case 1 'BUSQUEDA
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                'Encuentra registros
                PonerModo 2
                CargaGrid cadB
                DataGrid1.SetFocus
            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
        Case 1 'Buscar
            CargaGrid
    End Select
    PonerModo 2
    DataGrid1.SetFocus
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = adodc1.Recordset.Fields(0) & "|"
    cad = cad & adodc1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    ' ICONITOS DE LA BARRA
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
    FormatoCod = FormatoCampo(txtAux(0))
    
    '## A mano
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    DataGrid1.ClearFields

    CadAncho = False
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 2
      
    'Cadena consulta
    CadenaConsulta = "Select * from smarca"
    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnNuevo_Click
        Case 2: mnModificar_Click
        Case 3: mnEliminar_Click
        
        Case 5: mnBuscar_Click
        Case 6: mnVerTodos_Click
        
        Case 8   'Imprimir Listado de Marcas
            Me.Hide
            AbrirListado (1) 'OpcionListado=1
            Me.Show vbModal
    End Select
End Sub



Private Sub CargaGrid(Optional SQL As String)
Dim i As Byte
Dim b As Boolean
    
    b = DataGrid1.Enabled

    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codmarca"

    CargaGridGnral DataGrid1, Me.adodc1, SQL, False
    DataGrid1.RowHeight = 350
    
    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "Marca"
        DataGrid1.Columns(i).Width = 1000
        DataGrid1.Columns(i).NumberFormat = FormatoCod
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Denominaci�n"
        DataGrid1.Columns(i).Width = 4630 '2930
            
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        CadAncho = True
    End If
   
   'No permitir cambiar tama�o de columnas
   For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
   Next i
   
    'Habilitamos botones Modificar y Eliminar
   If Toolbar1.Buttons(6).Enabled = True Then
        Toolbar1.Buttons(6).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        mnModificar.Enabled = Not adodc1.Recordset.EOF
        mnEliminar.Enabled = Not adodc1.Recordset.EOF
   End If
   
   DataGrid1.Enabled = b
   DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not adodc1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
   PonerOpcionesMenu
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    If Index = 0 Then PonerFormatoEntero txtAux(Index) 'codmarca
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    'Comprobar si ya existe el cod de marca en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(txtAux(0)) Then b = False
    End If
    
    DatosOk = b
End Function


Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub