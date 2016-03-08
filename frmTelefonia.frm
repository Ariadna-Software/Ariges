VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTelefonia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recarga teléfonos"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   Icon            =   "frmTelefonia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      Height          =   290
      Index           =   1
      Left            =   8400
      TabIndex        =   20
      Top             =   5520
      Width           =   135
   End
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      Height          =   290
      Index           =   0
      Left            =   3960
      TabIndex        =   19
      Top             =   4920
      Width           =   150
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "descForpa"
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   7680
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "Trabajador|T|S|||stelefonia|codtraba|||"
      Text            =   "CCOst"
      Top             =   5520
      Width           =   675
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   8880
      MaxLength       =   1
      TabIndex        =   5
      Tag             =   "Tipo|T|N|||stelefonia|tipo|||"
      Text            =   "Descripcion"
      Top             =   4920
      Width           =   435
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTelefonia.frx":000C
      Left            =   8160
      List            =   "frmTelefonia.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "Cob.|N|N|||stelefonia|cobrado|||"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   6720
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Ref.|T|N|||stelefonia|referencia|||"
      Text            =   "Descripcion"
      Top             =   4920
      Width           =   915
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "descForpa"
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3360
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Forma pago|N|N|||stelefonia|fpago|0||"
      Text            =   "Fp"
      Top             =   4920
      Width           =   915
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Importe|N|N|||stelefonia|importe|##0.00||"
      Text            =   "Descripcio"
      Top             =   4920
      Width           =   1395
   End
   Begin VB.ComboBox CboCtrLotes 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTelefonia.frx":0022
      Left            =   7560
      List            =   "frmTelefonia.frx":0024
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Fac|N|N|||stelefonia|facturado|||"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "Telefono|T|N|||stelefonia|telefono||N|"
      Text            =   "te"
      Top             =   4920
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   960
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||stelefonia|fecha|||"
      Text            =   "Descripcion"
      Top             =   4920
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTelefonia.frx":0026
      Height          =   4935
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8705
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
      Left            =   10920
      TabIndex        =   10
      Top             =   5600
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9720
      TabIndex        =   9
      Top             =   5600
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10920
      TabIndex        =   14
      Top             =   5600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   12
      Top             =   5400
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
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1680
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
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
      Height          =   330
      Left            =   1440
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   7
      Left            =   11280
      MaxLength       =   30
      TabIndex        =   18
      Tag             =   "id|N|N|||stelefonia|id||S|"
      Text            =   "id"
      Top             =   5040
      Width           =   75
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
Attribute VB_Name = "frmTelefonia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'RECARGA DE TELEFONIA
'Autor: David
'Fecha creación: 10/09/2007
'Fecha modificacion:
'+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+


'El txt que lleva la columna clave esta oculto(pero visible)

Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

'Public Event DatoSeleccionado(CadenaSeleccion As String)

'

Private CadenaConsulta As String

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores 'Mto de Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmF As frmFacFormasPago
Attribute frmF.VB_VarHelpID = -1

Private miCodTra As Integer


'Dim FormatoCod As String 'formato del campo de codigo
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
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
 Dim I  As Integer
 
    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    For I = 0 To 6                  'El 7 esta oculto
        txtAux(I).visible = Not b
    Next
    cmdAux(0).visible = Not b
    'cmdAux(1).visible = vEmpresa.TieneAnalitica And (Not b)
    cmdAux(1).visible = Not b
    Me.CboCtrLotes.visible = Not b
    
    Combo1.visible = Not b
    Text1(0).visible = Not b
    Text1(1).visible = Not b
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    
    CboCtrLotes.Enabled = (Modo = 1) 'solo en buqueda
    txtAux(6).Enabled = (Modo = 1)
    cmdAux(1).Enabled = (Modo = 1)
    'Si es regresar
    'If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    cmdRegresar.visible = False
    
        
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                            'de permisos del usuario
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
    If miCodTra < 0 Then
        MsgBox "No tiene asignado trabajador", vbExclamation
        Exit Sub
    End If
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Adodc1
    
    'Limpiamos los campos para insertar
    limpiar Me
    
    'por defecto control de lotes vale NO
    Me.CboCtrLotes.ListIndex = 0
    Me.Combo1.ListIndex = 0
    
    anc = ObtenerAlto(DataGrid1)
    LLamaLineas anc, 3
    
    'Campos por defecto
    txtAux(1).Text = Format(Now, "dd/mm/yyyy")
    txtAux(5).Text = "A"
    
    'Nombre trbajador
    txtAux(6).Text = miCodTra
    Text1(1).Text = vUsu.Nombre
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    limpiar Me
    CargaGrid "id= -1"
    Me.CboCtrLotes.ListIndex = -1
    Me.Combo1.ListIndex = -1
    LLamaLineas 690, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
    On Error Resume Next

    CargaGrid ""
    If Adodc1.Recordset.RecordCount <= 0 Then
         MsgBox "No hay ningún registro en la tabla de telefonia", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        PonerModo 2
        PonerFocoGrid DataGrid1
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
Dim anc As Single
Dim I As Integer

    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    
    If Not SePuedeCambiar Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    'Llamamos al form
    For I = 0 To 3
        txtAux(I).Text = DataGrid1.Columns(I).Text
    Next I
    Text1(0).Text = DataGrid1.Columns(4).Text
    txtAux(4).Text = DataGrid1.Columns(5).Text
    txtAux(5).Text = DataGrid1.Columns(6).Text
   ' If vEmpresa.TieneAnalitica Then
        I = 9   'Para los dos ultimos campos, los dos combos
        txtAux(6).Text = DataGrid1.Columns(7).Text
        Text1(1).Text = DataGrid1.Columns(8).Text
   ' Else
   '     'Si no tiene analitica entonces los campos 7 y 8 son el cobrado y el facturad
   '     I = 7
   ' End If
    
    PosicionarComboDes Me.CboCtrLotes, DataGrid1.Columns(I).Text
    PosicionarComboDes Me.Combo1, DataGrid1.Columns(I + 1).Text
    txtAux(7).Text = CStr(Adodc1.Recordset!Id)
    anc = ObtenerAlto(DataGrid1)
    LLamaLineas anc, 4
   
    'Como es modificar
    PonerFoco txtAux(0)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Integer
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    For I = 0 To 6
        txtAux(I).Top = alto
    Next I
    For I = 0 To 1
        Text1(I).Top = alto
        cmdAux(I).Top = alto
    Next I
    'If Not vEmpresa.TieneAnalitica Then
    '    'Los campos de analitica van a FALSE
    '    txtAux(6).visible = False
    '    Text1(1).visible = False
    'End If
    Me.CboCtrLotes.Top = alto - 15
    Me.Combo1.Top = alto - 15
    
    
    
'    txtAux(0).Left = DataGrid1.Left + 340
'    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
'    Me.CboCtrLotes.Left = txtAux(1).Left + txtAux(1).Width + 55
End Sub

Private Function SePuedeCambiar() As Boolean
Dim SePuede As Boolean

    SePuede = False
    
    If CStr(Adodc1.Recordset.Fields(9)) = "Si" Then
        If vUsu.Codigo = 0 Then
            MsgBox "NO DEBERIA modificar la recarga", vbExclamation
            SePuede = True
        Else
            MsgBox "La recarga ha sido facturada. No puede modificarse", vbExclamation
        End If
    Else
        If miCodTra < 0 Then
            SePuede = False
        Else
            If vUsu.Codigo = 0 Then
                SePuede = True
            Else
                If Adodc1.Recordset!CodTraba <> miCodTra Then
                    MsgBox "No es el usuario que genero la recarga", vbExclamation
                Else
                    SePuede = True
                End If
            End If
        End If
    End If
    SePuedeCambiar = SePuede
    
End Function

Private Sub BotonEliminar()
Dim SQL As String

    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub


    If Not SePuedeCambiar Then Exit Sub

    '### a mano
    SQL = "¿Seguro que desea eliminar la recarga?" & vbCrLf
    SQL = SQL & vbCrLf & "Teléfono: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Fecha: " & Adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Importe: " & Adodc1.Recordset.Fields(2)
    SQL = SQL & vbCrLf & "REF: " & Adodc1.Recordset.Fields(5)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Adodc1.Recordset.AbsolutePosition
        SQL = "Delete from stelefonia where id=" & Adodc1.Recordset!Id
        conn.Execute SQL
        CancelaADODC Me.Adodc1
        CargaGrid ""
'        CancelaADODC Me.adodc1
        SituarDataPosicion Me.Adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Categoría.", Err.Description
End Sub



Private Sub CboCtrLotes_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim I As String
Dim cadB As String

    On Error GoTo EAceptar

    Select Case Modo
        Case 3 'Insertar
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If
        
        Case 4  'Modificar
            If DatosOk Then
                If BLOQUEADesdeFormulario(Me) Then
                    If ModificaDesdeFormulario(Me, 3) Then
                        TerminaBloquear
                        I = Adodc1.Recordset!Id
                        PonerModo 2
                        CancelaADODC Me.Adodc1
                        CargaGrid
                        Adodc1.Recordset.Find (" id =" & I)
                    End If
                    PonerFocoGrid DataGrid1
                End If
            End If
            
        Case 1 'HacerBusqueda
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                PonerModo 2
                CargaGrid cadB
                PonerFocoGrid DataGrid1
            End If
    End Select
    
EAceptar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)


    If Index = 1 Then
        'TRABAJADORES
        
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
            
    Else
        'Case 0.  FORPA
'                      'Llamamos a al form
'                      CadenaDesdeOtroForm = ""
'
'                      'Estamos en Modo de Cabeceras
'                      'Registro de la tabla de cabeceras: slista
'                      Set frmB = New frmBuscaGrid
'                      frmB.vSQL = ""
'                      frmB.vDevuelve = "0|1|"
'                      frmB.vselElem = 1
'                      frmB.vCargaFrame = False
'
'                      NumRegElim = Index
'
'                '      If Index = 0 Then
'                          'Forpa
'                          'Cod Diag.|tabla|columna|tipo|formato|10·
'                          CadenaDesdeOtroForm = "Código|sforpa|codforpa|N||20·Descripcion|sforpa|nomforpa|T||80·"
'
'
'                          frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
'                          frmB.vTabla = "sforpa"
'                          frmB.vTitulo = "Formas de pago"
'                          frmB.vCampos = CadenaDesdeOtroForm
'
'
'                '       'YA NO ES CENRO DE COSTE. Es trabajador ahi arriba
'                '      Else
'                '          'CENTRO COSTE
'                '          'Cod Diag.|tabla|columna|tipo|formato|10·
'                '          CadenaDesdeOtroForm = "Código|cabccost|codccost|T||20·Descripcion|cabccost|nomccost|T||80·"
'                '
'                '
'                '          frmB.vConexionGrid = conConta
'                '          frmB.vTabla = "cabccost"
'                '          frmB.vTitulo = "Centros de coste"
'                '          frmB.vCampos = CadenaDesdeOtroForm
'                '      End If
'                      frmB.Show vbModal
'                      Set frmB = Nothing
'                      CadenaDesdeOtroForm = ""
'
'

            Set frmF = New frmFacFormasPago
            frmF.DatosADevolverBusqueda = "0|1|"
            frmF.Show vbModal
            Set frmF = Nothing
    End If
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next

    Select Case Modo
        Case 1 'Busqueda
            CargaGrid
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            TerminaBloquear
            Me.lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End Select
    
    PonerModo 2
    PonerFocoGrid DataGrid1
    
    If Err.Number <> 0 Then Err.Clear
End Sub


'Private Sub cmdRegresar_Click()
'Dim cad As String
'
'    If adodc1.Recordset.EOF Then
'        MsgBox "Ningún registro devuelto.", vbExclamation
'        Exit Sub
'    End If
'
'    cad = adodc1.Recordset.Fields(0) & "|"
'    cad = cad & adodc1.Recordset.Fields(1) & "|"
'    RaiseEvent DatoSeleccionado(cad)
'    Unload Me
'End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
   ' If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     If Not Adodc1.Recordset.EOF Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()

    FijarCodigoTrabajador   '
    
    
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
    
    
   ' cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 2
    
    'Cadena consulta
    CadenaConsulta = "select telefono,fecha,importe,fpago,nomforpa,referencia,tipo"
    
    CadenaConsulta = CadenaConsulta & ",stelefonia.codtraba,nomtraba"
    'Los dos COMBOS al final
    CadenaConsulta = CadenaConsulta & ",if(facturado=1,""Si"",""No""),if(cobrado=1,""Si"",""No"")"
    CadenaConsulta = CadenaConsulta & ",id from stelefonia,sforpa"
    CadenaConsulta = CadenaConsulta & ",straba"
    
    CadenaConsulta = CadenaConsulta & " WHERE stelefonia.fpago = sforpa.codforpa"
    CadenaConsulta = CadenaConsulta & " AND straba.codtraba = stelefonia.codtraba"
    CargaGrid
    CargaCombo
    

    'Posicionamos el form y los botones segun tenga, o no, analitica
    NumRegElim = 0
    'If Not vEmpresa.TieneAnalitica Then NumRegElim = 2210
    Me.Width = Me.Width - NumRegElim
    Me.cmdAceptar.Left = cmdAceptar.Left - NumRegElim
    Me.cmdCancelar.Left = Me.cmdCancelar.Left - NumRegElim
    Me.cmdRegresar.Left = Me.cmdCancelar.Left
    DataGrid1.Width = DataGrid1.Width - NumRegElim
    Me.Height = Me.Height + 120
    NumRegElim = 0
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    If NumRegElim = 0 Then
        'Forpa
        txtAux(3).Text = RecuperaValor(CadenaDevuelta, 1)
        Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
    Else
        'Centro de coste
        txtAux(6).Text = RecuperaValor(CadenaDevuelta, 1)
        Text1(1).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
End Sub

Private Sub frmF_DatoSeleccionado(CadenaSeleccion As String)
        txtAux(3).Text = RecuperaValor(CadenaSeleccion, 1)
        Text1(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
        txtAux(6).Text = RecuperaValor(CadenaSeleccion, 1)
        Text1(1).Text = RecuperaValor(CadenaSeleccion, 2)
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
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5: mnNuevo_Click
        Case 6: mnModificar_Click
        Case 7: mnEliminar_Click
        'Imprime el justificante del pago
        Case 10: ImprimeJusti
                                
        Case 11: mnSalir_Click  'Salir
    End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
Dim b As Boolean
Dim tots As String

    b = DataGrid1.Enabled

    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY fecha,id"
    
    CargaGridGnral DataGrid1, Me.Adodc1, SQL, False


    '### a mano
    '        telefono                           Fecha
    tots = "S|txtAux(0)|T|Teléfono|1100|;S|txtAux(1)|T|Fecha|1000|;S|txtAux(2)|T|Importe|900|;S|txtAux(3)|T|F.P.|800|;"
    'el botoncito
    tots = tots & "S|cmdAux(0)|B||150|;"
    tots = tots & "S|Text1(0)|T|Descripción FP|2300|;S|txtAux(4)|T|Referencia|1150|;S|txtAux(5)|T|Tipo|650|;"

    tots = tots & "S|txtAux(6)|T|C.Tra|700|;S|cmdAux(1)|B||150|;S|Text1(1)|T|Nombre|1500|;"

    'Los dos cmbos al final
    tots = tots & "S|CboCtrLotes|C|Fact.|600|;S|Combo1|C|Pag.|600|;"
    'El ID ocilto
    tots = tots & "N|txtAux(7)|T|||;"
    
    arregla tots, DataGrid1, Me


    'Reajuste tamaño
    txtAux(5).Left = txtAux(5).Left - 30
    txtAux(5).Width = 320
    txtAux(6).Left = txtAux(6).Left - 30
    Text1(1).Left = Text1(1).Left - 30
    Text1(0).Left = Text1(0).Left - 30
    Text1(0).Width = Text1(0).Width + 30
    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    Me.CboCtrLotes.Left = Me.CboCtrLotes.Left - 30
    Me.Combo1.Left = Me.Combo1.Left - 30
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
   
End Sub

Private Sub CargaCombo()
    'Carga la lista de impresión de etiquetas
    Me.CboCtrLotes.Clear
    Combo1.Clear
    
    CboCtrLotes.AddItem "No"
    CboCtrLotes.ItemData(CboCtrLotes.NewIndex) = 0
    CboCtrLotes.AddItem "Si"
    CboCtrLotes.ItemData(CboCtrLotes.NewIndex) = 1
    
    
    Combo1.AddItem "No"
    Combo1.ItemData(Combo1.NewIndex) = 0
    
    Combo1.AddItem "Si"
    Combo1.ItemData(Combo1.NewIndex) = 1
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
    
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then
            If Index = 3 Then Text1(0).Text = ""
            If Index = 6 Then Text1(1).Text = ""
            Exit Sub
    End If
    
    
    
    Select Case Index
    Case 1
        'La fecha
        PonerFormatoFecha txtAux(1)
        If txtAux(1).Text = "" Then PonerFoco txtAux(1)
    
    Case 2
        If Not PonerFormatoDecimal(txtAux(2), 6) Then PonerFoco txtAux(2)
        
    Case 3
        Text1(0).Text = PonerNombreDeCod(txtAux(Index), conAri, "sforpa", "nomforpa", "codforpa")
        
    Case 6
        
        Text1(1).Text = PonerNombreDeCod(txtAux(Index), conAri, "straba", "nomtraba", "codtraba")
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    
    DatosOk = False
    
    'La forpa y si tiene anal el ccost
    If Text1(0).Text = "" Then
        MsgBox "Ponga la forma de pago", vbExclamation
        Exit Function
    End If
    
    
    'La forma de pago tiene que ser efectivo
    If Not FormaPagoCorrecta Then Exit Function
    
    
    If Text1(1).Text = "" Then
        MsgBox "Trabajador incorrecto", vbExclamation
        Exit Function
    End If
        
        'Sugerimos codigo siguiente
    If Modo = 3 Then txtAux(7).Text = SugerirCodigoSiguienteStr("stelefonia", "id")
    b = CompForm(Me, 3)
    If Not b Then Exit Function

    


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






Private Sub ImprimeJusti()
Dim C As String

    C = ""
    If Not (Adodc1.Recordset Is Nothing) Then
        If Adodc1.Recordset.RecordCount > 0 Then C = "O"
    End If
    If C = "" Then Exit Sub
    
    
    'Para evitar declarar mas variables entodavia
    NumRegElim = Modo
    If PonerParamRPT(22, C, Modo, CadenaDesdeOtroForm, pImprimeDirecto, pPdfRpt) Then
        With frmImprimir
            .ConSubInforme = False
            .FormulaSeleccion = "{stelefonia.id} = " & Adodc1.Recordset!Id
            .NombreRPT = CadenaDesdeOtroForm
            .NombrePDF = pPdfRpt
            .OtrosParametros = C
            .NumeroParametros = Modo
            .Opcion = 2003 'Esta libre
            .Show vbModal
        End With
    End If
    Modo = CByte(NumRegElim)
    NumRegElim = 0




    
End Sub



Private Sub FijarCodigoTrabajador()
    miCodTra = -1
    CadenaConsulta = DevuelveDesdeBDNew(conAri, "straba", "codtraba", "login", vUsu.Login, "T")
    If CadenaConsulta = "" Then
            MsgBox "El usuario: " & vUsu.Nombre & " (" & vUsu.Login & ")    NO tiene trabajador asignado", vbExclamation
    Else
        miCodTra = Val(CadenaConsulta)
    End If
    CadenaConsulta = ""
        
End Sub

Private Function FormaPagoCorrecta() As Boolean
    FormaPagoCorrecta = False
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", txtAux(3).Text, "N")
    NumRegElim = -1
    If CadenaDesdeOtroForm <> "" Then NumRegElim = Val(CadenaDesdeOtroForm)
    If NumRegElim <> 0 Then
        MsgBox "Forma de pago INCORRECTA. Solo contado.", vbExclamation
    Else
        FormaPagoCorrecta = True
    End If
    CadenaDesdeOtroForm = ""
End Function
