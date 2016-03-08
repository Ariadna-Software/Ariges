VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacSituaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Situaciones Especiales"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmFacSituaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboBusqweb 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmFacSituaciones.frx":000C
      Left            =   5160
      List            =   "frmFacSituaciones.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Ocultar busqueda clientes WEB|N|N|||ssitua|ocultarbus||N|"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ComboBox cboclioferped 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmFacSituaciones.frx":0010
      Left            =   3840
      List            =   "frmFacSituaciones.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Bloquea ofertas pedidos cliente|N|N|||ssitua|clioferped||N|"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ComboBox CboTipoSitu 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmFacSituaciones.frx":0014
      Left            =   2400
      List            =   "frmFacSituaciones.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Bloquea|N|N|||ssitua|tipositu||N|"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "Código Situaciones Especiales|N|N|0|99|ssitua|codsitua|00|S|"
      Text            =   "Codigo"
      Top             =   4920
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   960
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Descripción Situaciones Especiales|T|N|||ssitua|nomsitua||N|"
      Text            =   "Descripcion"
      Top             =   4920
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacSituaciones.frx":0018
      Height          =   4845
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   540
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8546
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
      Left            =   6660
      TabIndex        =   6
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5460
      TabIndex        =   5
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6660
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   8
      Top             =   5355
      Width           =   2115
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
         Left            =   210
         TabIndex        =   9
         Top             =   210
         Width           =   1800
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
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
            ImageIndex      =   16
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
Attribute VB_Name = "frmFacSituaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)

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


Private Sub PonerModo(vModo As Byte)
Dim b As Boolean

    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    CboTipoSitu.visible = Not b
    cboclioferped.visible = Not b
    cboBusqweb.visible = Not b
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    End If
    
    'Si estamos insertando o busqueda
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Buscar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ber Todos
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    b = b And Not DeConsulta
    'Añadir
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


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub BotonAnyadir()
Dim anc As Single
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Adodc1
    
    'Obtenemos la siguiente numero de codigo de Situaciones
    txtAux(0).Text = SugerirCodigoSiguienteStr("ssitua", "codsitua")
    FormateaCampo txtAux(0)
    txtAux(1).Text = ""
    CboTipoSitu.ListIndex = 1
    cboclioferped.ListIndex = 1
    cboBusqweb.ListIndex = 0
    anc = ObtenerAlto(DataGrid1)
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid "codsitua= -1"
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    Me.CboTipoSitu.ListIndex = -1
    cboclioferped.ListIndex = -1
    cboBusqweb.ListIndex = -1
    LLamaLineas 750, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next

    CargaGrid ""
    If Adodc1.Recordset.RecordCount <= 0 Then
         MsgBox "No hay ningún registro en la tabla ssitua", vbInformation
         Screen.MousePointer = vbDefault
          Exit Sub
    Else
        PonerModo 2
         DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
Dim anc As Single
Dim i As Integer

    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    If EsCodigoCero(CStr(Adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    Select Case DataGrid1.Columns(2).Value
        Case "Si"
            CboTipoSitu.ListIndex = 1
        Case "No"
            CboTipoSitu.ListIndex = 0
    End Select
    If CStr(DataGrid1.Columns(3).Value) = "Si" Then
            cboclioferped.ListIndex = 1
    Else
            cboclioferped.ListIndex = 0
    End If
    If CStr(DataGrid1.Columns(4).Value) = "Si" Then
            cboBusqweb.ListIndex = 1
    Else
            cboBusqweb.ListIndex = 0
    End If
    anc = ObtenerAlto(DataGrid1)
    LLamaLineas anc, 4
   
    'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
'Pone posicion TOP y LEFT de los controles en el form
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    CboTipoSitu.Top = alto - 15
    cboclioferped.Top = alto - 15
    cboBusqweb.Top = alto - 15
    txtAux(0).Left = DataGrid1.Left + 340
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
    CboTipoSitu.Left = txtAux(1).Left + txtAux(1).Width + 65
    cboclioferped.Left = CboTipoSitu.Left + CboTipoSitu.Width + 9
    cboBusqweb.Left = cboclioferped.Left + cboclioferped.Width + 0
    
End Sub


Private Sub BotonEliminar()
Dim SQL As String
On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    If EsCodigoCero(CStr(Adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub

    '### a mano
    SQL = "¿Seguro que desea eliminar la Situación?" & vbCrLf
    SQL = SQL & vbCrLf & "Código: " & Format(Adodc1.Recordset.Fields(0), FormatoCod)
    SQL = SQL & vbCrLf & "Denominación: " & Adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Adodc1.Recordset.AbsolutePosition
        SQL = "Delete from ssitua where codsitua=" & Adodc1.Recordset!codsitua
        conn.Execute SQL
        CancelaADODC Me.Adodc1
        CargaGrid ""
        CancelaADODC Me.Adodc1
        SituarDataPosicion Me.Adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Situaciones Especiales", Err.Description
End Sub


Private Sub cboBusqweb_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub CboTipoSitu_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboclioferped_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String
On Error Resume Next

    Select Case Modo
        Case 1 'HacerBusqueda
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                PonerModo 2
                CargaGrid cadB
                DataGrid1.SetFocus
            End If
        Case 3 'Insertar
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If
        Case 4 'Modificar
            If DatosOk And BLOQUEADesdeFormulario(Me) Then
                If ModificaDesdeFormulario(Me, 3) Then
                   TerminaBloquear
                   i = Adodc1.Recordset.Fields(0)
                   PonerModo 2
                   CancelaADODC Me.Adodc1
                   CargaGrid
                   Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & i)
                End If
                DataGrid1.SetFocus
            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Select Case Modo
        Case 1 'Busqueda
            CargaGrid
        Case 3
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            TerminaBloquear
            Me.lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End Select
    PonerModo 2
    DataGrid1.SetFocus
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = Adodc1.Recordset.Fields(0) & "|"
    cad = cad & Adodc1.Recordset.Fields(1) & "|"
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
    If Not Adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Recuperar Todos
        .Buttons(5).Image = 3   'Botón Añadir Nuevo Registro
        .Buttons(6).Image = 4   'Botón Modificar Registro
        .Buttons(7).Image = 5   'Botón Borrar Registro
        .Buttons(10).Image = 16  'Botón Imprimir
        .Buttons(11).Image = 15  'Botón Salir
    End With
    
    FormatoCod = FormatoCampo(txtAux(0))
    
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 2
    
    'Cadena consulta
    CadenaConsulta = "Select codsitua,nomsitua,If(tipositu=1,""Si"",""No"") AS tipositu,"
    CadenaConsulta = CadenaConsulta & "If(clioferped=1,""Si"",""No"") AS bloqoferped ,If(ocultarbus=1,""Si"",""No"") AS bloqoferped "
    CadenaConsulta = CadenaConsulta & " from ssitua"
    
    CargaGrid
    CargaCombo
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
        Case 1: BotonBuscar
        Case 2: BotonVerTodos
        Case 5: BotonAnyadir
        Case 6: BotonModificar
        Case 7: BotonEliminar
        Case 10 'Imprimir listado de Situaciones Especiales
                Me.Hide
                AbrirListado (27) 'OpcionListado=27
                Me.Show vbModal
        Case 11 'Salir
                mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
Dim b As Boolean
Dim tots As String
    
    b = DataGrid1.Enabled
    
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codsitua"
    
    CargaGridGnral DataGrid1, Me.Adodc1, SQL, False
    
    '### a mano
    tots = "S|txtAux(0)|T|Situación|900|;S|txtAux(1)|T|Denominación|2700|;S|CboTipoSitu|C|Bloquea|800|;"
    tots = tots & "S|cboclioferped|C|Blq ofe/ped|1050|;S|cboBusqweb|C|Busq. web|1050|;"
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If

'    'Habilitamos botones Modificar y Eliminar
'   If Toolbar1.Buttons(6).Enabled Then
'        Toolbar1.Buttons(6).Enabled = Not adodc1.Recordset.EOF
'        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
'        mnModificar.Enabled = Not adodc1.Recordset.EOF
'        mnEliminar.Enabled = Not adodc1.Recordset.EOF
'   End If
'   DataGrid1.Enabled = b
'   DataGrid1.ScrollBars = dbgAutomatic
'
'   PonerOpcionesMenu
End Sub

Private Sub CargaCombo()
    'Carga la lista de impresión de etiquetas
    CboTipoSitu.Clear
    
    CboTipoSitu.AddItem "No"
    CboTipoSitu.ItemData(CboTipoSitu.NewIndex) = 0
    
    CboTipoSitu.AddItem "Si"
    CboTipoSitu.ItemData(CboTipoSitu.NewIndex) = 1
    
    
    
    'Carga la lista de impresión de etiquetas
    cboclioferped.Clear
    
    cboclioferped.AddItem "No"
    cboclioferped.ItemData(cboclioferped.NewIndex) = 0
    
    cboclioferped.AddItem "Si"
    cboclioferped.ItemData(cboclioferped.NewIndex) = 1
    
    
    'Carga bloquea en web
    cboBusqweb.Clear
    
    cboBusqweb.AddItem "No"
    cboBusqweb.ItemData(cboBusqweb.NewIndex) = 0
    
    cboBusqweb.AddItem "Si"
    cboBusqweb.ItemData(cboBusqweb.NewIndex) = 1
    
    
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
    If Index = 0 Then PonerFormatoEntero txtAux(Index)
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function


    'Si no bloquea no puede bloquear ofertas pedidos
    If Me.cboclioferped.ListIndex = 1 Then
        'Si bloquea oferta/pedido
        If Me.CboTipoSitu.ListIndex = 0 Then
            MsgBox "No puede bloquear la oferta/pedido sin bloquear la situacion del cliente", vbExclamation
            Exit Function
        End If
    End If


    'comprobar si ya existe el codigo de situacion en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(txtAux(0)) Then b = False
    End If
    
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

