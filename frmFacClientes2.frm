VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacClientes3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "frmFacClientes2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   4320
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Nombre Comercial|T|N|||sclien|nomcomer||N|"
      Text            =   "Dato2"
      Top             =   5040
      Width           =   3195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Código Cliente|N|N|0|999999|sclien|codclien|000000|S|"
      Text            =   "Dat"
      Top             =   5040
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Nombre Cliente|T|N|||sclien|nomclien||N|"
      Text            =   "Dato2"
      Top             =   5040
      Width           =   3195
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacClientes2.frx":000C
      Height          =   4710
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   540
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8308
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7620
      TabIndex        =   4
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6420
      TabIndex        =   3
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7620
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   3435
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
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3120
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
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
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Busqueda avanzada"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
      Begin VB.Menu mnBusqAvan 
         Caption         =   "Busqueda &Avanzada"
         HelpContextID   =   2
         Shortcut        =   ^A
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnOrdenado 
      Caption         =   "Ordenacion"
      Begin VB.Menu mnOrdenadoPor 
         Caption         =   "Código"
         Index           =   0
      End
      Begin VB.Menu mnOrdenadoPor 
         Caption         =   "Nombre"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmFacClientes3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)



Private WithEvents frmC As frmFacClientes
Attribute frmC.VB_VarHelpID = -1

Private CadenaConsulta As String

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
'-------------------------------------------------------


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(vModo As Byte)
Dim B As Boolean

    Modo = vModo
    B = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    Me.txtAux(0).visible = Not B
    txtAux(1).visible = Not B
    txtAux(2).visible = Not B
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    'Si estamos en insertar o modificar
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    
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
    'Buscar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ber Todos
    Toolbar1.Buttons(2).Enabled = B
    Me.mnVerTodos.Enabled = B
    
'     b = b And Not DeConsulta
'    'Añadir
'    Toolbar1.Buttons(5).Enabled = b
'    Me.mnNuevo.Enabled = b
'    'Modificar
'    Toolbar1.Buttons(6).Enabled = b
'    Me.mnModificar.Enabled = b
'    'Eliminar
'    Toolbar1.Buttons(7).Enabled = b
'    Me.mnEliminar.Enabled = b
'    'Imprimir
'    Toolbar1.Buttons(10).Enabled = b
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub BotonAnyadir()
'Dim anc As Single
'
'    'Situamos el grid al final
'    AnyadirLinea DataGrid1, adodc1
'
'    'Obtenemos la siguiente numero de factura
'    txtAux(0).Text = SugerirCodigoSiguienteStr("sactiv", "codactiv")
'    txtAux(0).Text = Format(txtAux(0).Text, FormatoCod)
'    txtAux(1).Text = ""
'
'    anc = ObtenerAlto(DataGrid1)
'    LLamaLineas anc, 3
'
'    'Ponemos el foco
'    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid2 "codclien= -1"  'para vaciar los datos del Grid
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    LLamaLineas 750, 1
    PonerFoco txtAux(1)
End Sub

Private Sub BotonVerTodos()
On Error Resume Next

    CargaGrid2 ""
    If Adodc1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla sactiv", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
'Dim anc As Single
'Dim i As Integer
'
'    If adodc1.Recordset.EOF Then Exit Sub
'    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'
'    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
'        i = DataGrid1.Bookmark - DataGrid1.FirstRow
'        DataGrid1.Scroll 0, i
'        DataGrid1.Refresh
'    End If
'
'    'Llamamos al form
'    txtAux(0).Text = DataGrid1.Columns(0).Text
'    txtAux(1).Text = DataGrid1.Columns(1).Text
'
'    anc = ObtenerAlto(DataGrid1)
'    LLamaLineas anc, 4
'
'    'Como es modificar
''    PonerFoco txtAux(1)
'    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
'    txtAux(0).Left = DataGrid1.Left + 340
'    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
'    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
End Sub


Private Sub BotonEliminar()
'Dim SQL As String
'
'    On Error GoTo Error2
'
'    'Ciertas comprobaciones
'    If adodc1.Recordset.EOF Then Exit Sub
'
'    '### a mano
'    SQL = "¿Seguro que desea eliminar la Actividad?" & vbCrLf
'    SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), FormatoCod)
'    SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
'    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
'        'Hay que eliminar
'        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
'        SQL = "Delete from sactiv where codactiv=" & adodc1.Recordset!codactiv
'        Conn.Execute SQL
'        CancelaADODC Me.adodc1
'        CargaGrid ""
'        CancelaADODC Me.adodc1
'        SituarDataPosicion Me.adodc1, NumRegElim, SQL
'    End If
'
'Error2:
'    Screen.MousePointer = vbDefault
'    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Actividad Cliente", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String

    On Error Resume Next

    Select Case Modo
        Case 1 'HacerBusqueda
        
            'Modificacion para que no tenga que poner *
            If txtAux(1).Text <> "" Then
                'No lo ha puesto el. Se lo pongo YO
                If InStr(1, txtAux(1).Text, "*") = 0 Then txtAux(1).Text = "*" & txtAux(1).Text & "*"
            End If
            If txtAux(2).Text <> "" Then
                'No lo ha puesto el. Se lo pongo YO
                If InStr(1, txtAux(2).Text, "*") = 0 Then txtAux(2).Text = "*" & txtAux(2).Text & "*"
            End If
            
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                PonerModo 2
                CargaGrid2 cadB
                DataGrid1.SetFocus
            End If
        
'        Case 3  'Hacemos insertar
'            If DatosOk Then
'                If InsertarDesdeForm(Me) Then
'                    CargaGrid
'                    BotonAnyadir
'                End If
'            End If
'
'        Case 4 'Modificar
'             If DatosOk And BLOQUEADesdeFormulario(Me) Then
'                 If ModificaDesdeFormulario(Me, 3) Then
'                      TerminaBloquear
'                      i = adodc1.Recordset.Fields(0)
'                      PonerModo 2
'                      CancelaADODC Me.adodc1
'                      CargaGrid
'                      adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
'                  End If
'                  DataGrid1.SetFocus
'            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
    On Error Resume Next

    Select Case Modo
        Case 1 'busqueda
            CargaGrid2

'        Case 3 'Insertar
'            DataGrid1.AllowAddNew = False
'            'CargaGrid
'            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
'
'        Case 4 'Modificar
'            'CargaGrid
'            TerminaBloquear
''            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
'            Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    End Select
    PonerModo 2
    DataGrid1.SetFocus
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    Cad = Adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & Adodc1.Recordset.Fields(1) & "|"
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
    If Not Adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Me.txtAux(1)
End Sub


Private Sub Form_Load()
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1    'Botón Busqueda
        .Buttons(2).Image = 2    'Botón Recuperar Todos
        .Buttons(5).Image = 19    'Botón busqu avanz
'        .Buttons(6).Image = 4    'Botón Modificar Registro
'        .Buttons(7).Image = 5    'Botón Borrar Registro
'        .Buttons(10).Image = 16  'Botón Imprimir
        .Buttons(11).Image = 15  'Botón Salir
    End With
    
    
    FormatoCod = CheckValueLeer(Me.Name)
    If FormatoCod <> "1" Then FormatoCod = "0"
    Me.mnOrdenadoPor(CInt(FormatoCod)).Checked = True
    
    
    FormatoCod = FormatoCampo(txtAux(0))
    
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
'    PonerModo 2
    
    'Cadena consulta
    CadenaConsulta = "Select codclien,nomclien,nomcomer from sclien"
    BotonBuscar
'    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Modo = 0
    If Me.mnOrdenadoPor(1).Checked Then Modo = 1
    CheckValueGuardar Me.Name, Modo
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    FormatoCod = CadenaSeleccion
End Sub

'Private Sub Form_Unload(Cancel As Integer)
''    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
'End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub







Private Sub mnBusqAvan_Click()
    FormatoCod = ""
    Set frmC = New frmFacClientes
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.Show vbModal
    Set frmC = Nothing
    If FormatoCod <> "" Then
        RaiseEvent DatoSeleccionado(FormatoCod)
        Unload Me
    Else
        FormatoCod = FormatoCampo(txtAux(0))
    End If
End Sub

Private Sub mnOrdenadoPor_Click(Index As Integer)
        mnOrdenadoPor(0).Checked = Index = 0
        mnOrdenadoPor(1).Checked = Index = 1
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
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5:
                mnBusqAvan_Click
'        Case 6: mnModificar_Click
'        Case 7: mnEliminar_Click
'        Case 10  'Informes
'                Me.Hide
'                AbrirListado (20)  'OpcionListado=20
'                Me.Show vbModal
        Case 11 'Salir
                mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid2(Optional SQL As String)
Dim B As Boolean
Dim tots As String
    
    B = DataGrid1.Enabled
    
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.CodigoAgente > 0 Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & " codagent = " & vUsu.CodigoAgente
        End If
    End If
    
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
    Else
        SQL = CadenaConsulta
    End If
    
    
    If mnOrdenadoPor(1).Checked Then
        SQL = SQL & " ORDER BY sclien.nomclien"
    Else
        SQL = SQL & " ORDER BY sclien.codclien"
    End If


    CargaGridGnral DataGrid1, Me.Adodc1, SQL, False
    
    '### a mano
    tots = "S|txtAux(0)|T|Cliente|900|;S|txtAux(1)|T|Nombre|3900|;S|txtAux(2)|T|Nombre Comercial|3000|;"
    arregla tots, DataGrid1, Me
    
    DataGrid1.Enabled = B
    DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    If Index = 0 Then PonerFormatoEntero txtAux(Index)
End Sub


Private Function DatosOk() As Boolean
'Dim b As Boolean
'
'    b = CompForm(Me, 3)
'    If Not b Then Exit Function
'
'    If Modo = 3 Then 'Insertar
'        If ExisteCP(txtAux(0)) Then b = False
'    End If
'    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

