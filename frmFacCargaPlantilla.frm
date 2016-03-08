VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacCargaPlantilla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Plantillas a Ofertas"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8715
   ClipControls    =   0   'False
   Icon            =   "frmFacCargaPlantilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   3
      Text            =   "existencia"
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   5355
      Width           =   3255
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
         TabIndex        =   8
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7395
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7395
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cargar Plantilla"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacCargaPlantilla.frx":000C
      Height          =   4730
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8334
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Left            =   3720
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
      TabIndex        =   6
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnCargar 
         Caption         =   "&Cargar Plantilla"
         HelpContextID   =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu mnBarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacCargaPlantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
'Solo utilizamos el Modo=4 -> Modificar

Dim kCampo As Integer

'Dim CadenaConsulta As String

'Dim CadenaBusqueda As String
'Cadena para la consulta de busqueda en Grid

Public Event CargarPlantillas()


Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
                           'Para el RowColChange, si el grid no esta totalmente cargado el CargaTxtAux da error.

Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
'Dim cad As String,
Dim Indicador As String
Dim NumReg As Long
'On Error GoTo Error1
'
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 4 'Modificar Cantidad (Introducir Valores a Cargar)
'        If DatosOk Then
            If ActualizarCantidad Then
'                TerminaBloquear
                CargaTxtAux False, False
                NumReg = Data1.Recordset.AbsolutePosition
                CargaGrid True
                If SituarDataPosicion(Data1, NumReg, Indicador) Then
                    If Not Data1.Recordset.EOF Then
                        If NumReg < Data1.Recordset.RecordCount Then 'No es Último Registro
                            'DataGrid1.Row = NumReg
                            Data1.Recordset.Move NumReg, 1
'                        Else 'Es el último Registro
'                            CargaTxtAux False, False
'                            PonerModo 2
                        End If
                        CargaTxtAux True, True
                    End If
                End If
            End If
'        End If
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    If Modo = 4 Then
          CargaTxtAux False, False
          PonerModo 0
'          Me.cmdAceptar.visible = False
    End If
    PonerModo 0
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF And gridCargado And Modo = 4 Then 'Modo4: Modificar
       CargaTxtAux True, True
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'La toolbar
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 4 'Modificar
        .Buttons(2).Image = 21 'Cargar Plantilla y Salir
        .Buttons(4).Image = 15 'Salir
    End With
    
    
'    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    PrimeraVez = True
    
    NombreTabla = "tmpscapla"
    Ordenacion = " ORDER BY codusu, codplant"
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    PonerModo 0
    
    'Carga la Tabla Temporal (tmpscapla) con las plantillas de la tabla scapla
    CargarDatos
    
'    Data1.ConnectionString = Conn
'    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codusu = " & vUsu.Codigo
'    Data1.RecordSource = CadenaConsulta
'    Data1.Refresh
    
    CargaGrid True
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
'Dim b As Boolean
Dim i As Byte
Dim SQL As String
On Error GoTo ECarga

'    b = DataGrid1.Enabled
    gridCargado = False
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
        
'    DataGrid1.Columns(0).visible = False
    'Cod. Grupo
    DataGrid1.Columns(0).Caption = "Grupo"
    DataGrid1.Columns(0).Width = 700
    DataGrid1.Columns(0).NumberFormat = "00"
    DataGrid1.Columns(0).Alignment = dbgCenter
    
    DataGrid1.Columns(1).Caption = "Nom. Grupo"
    DataGrid1.Columns(1).Width = 1700
       
    'Cod. Plantilla
    DataGrid1.Columns(2).Caption = "Plant."
    DataGrid1.Columns(2).Width = 700
    DataGrid1.Columns(2).NumberFormat = "000"
    DataGrid1.Columns(2).Alignment = dbgCenter
    
    'Nombre Plantilla
    DataGrid1.Columns(3).Caption = "Nom. Plant."
    DataGrid1.Columns(3).Width = 3200
    
    'Cantidad
    DataGrid1.Columns(4).Caption = "Cantidad"
    DataGrid1.Columns(4).Width = 1200
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(4).NumberFormat = FormatoImporte
    
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
        DataGrid1.Columns(i).Locked = True
    Next i
    
    Me.DataGrid1.ScrollBars = dbgAutomatic
    Me.DataGrid1.Enabled = (Modo = 0)
    gridCargado = True
    PrimeraVez = False
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
'Dim i As Integer

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
        'txtAux.visible = visible
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            txtAux.Text = DBLet(Data1.Recordset!Cantidad)
            txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If

        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux.Left = DataGrid1.Columns(4).Left + 130 'cantidad
        txtAux.Width = DataGrid1.Columns(4).Width - 10
    End If
    'Los ponemos Visibles o No
    '--------------------------
    txtAux.visible = visible
    
    PonerFoco txtAux
    
'    If visible Then
'        txtAux.TabIndex = 2
'    Else
'        txtAux.TabIndex = 5
'    End If
End Sub



Private Sub mnCargar_Click()
    'Cargar las lineas de las plantilla como lineas de la Oferta y Salir (Volver a Mto Ofertas)
    Unload Me
    RaiseEvent CargarPlantillas
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

'Private Sub imgBuscar_Click(Index As Integer)
'
'    If Modo = 2 Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'    imgBuscar(0).Tag = Index
'
'    Select Case Index
'        Case 0 'Codigo Almacen
'            Set frmA = New frmAlmAlPropios
'            frmA.DatosADevolverBusqueda = "0"
'            frmA.Show vbModal
'            Set frmA = Nothing
'        Case 1 'Codigo Familia / Cod. Proveedor
'            If vParamAplic.InventarioxProv Then
'                'Realizar inventario por Proveedor
'                Set frmP = New frmComProveedores
'                frmP.DatosADevolverBusqueda = "0"
'                frmP.Show vbModal
'                Set frmP = Nothing
'            Else 'Cod. Familia
'                Set frmFA = New frmAlmFamiliaArticulo
'                frmFA.DatosADevolverBusqueda = "0"
'                frmFA.Show vbModal
'                Set frmFA = Nothing
'            End If
'    End Select
'    PonerFoco Text1(0)
'    Screen.MousePointer = vbDefault
'End Sub



Private Sub txtAux_GotFocus()
    ConseguirFoco txtAux, 3
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Row > 0 Then
                    DataGrid1.Row = DataGrid1.Row - 1
                    CargaTxtAux True, True
                End If
        Case 40 'Desplazamiento Flecha Hacia Abajo
                If Data1.Recordset.AbsolutePosition <> Data1.Recordset.RecordCount Then
                    If DataGrid1.Row < Data1.Recordset.RecordCount - 1 Then
                        DataGrid1.Row = DataGrid1.Row + 1
                        CargaTxtAux True, True
                    End If
                End If
    End Select
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)
On Error GoTo EKeyPress
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then cmdCancelar_Click 'ESC
    End If
EKeyPress:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtAux_LostFocus()
    With txtAux
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        'Formato tipo 1: Decimal(12,2)
        PonerFormatoDecimal txtAux, 1
    End With
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Modificar
            mnModificar_Click
        Case 2 'Cargar Plantillas y Salir
            mnCargar_Click
        Case 4 'Salir
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
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
                
    Me.cmdAceptar.visible = (Modo = 4)
    Me.cmdCancelar.visible = (Modo = 4)
    Me.DataGrid1.Enabled = (Modo = 0)
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
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

    SQL = "SELECT scapla.codgrupl, grupo.nomgrupl, " & NombreTabla & ".codplant, scapla.nomplant, " & NombreTabla & ".cantidad "
    SQL = SQL & " FROM (" & NombreTabla & " LEFT JOIN scapla ON " & NombreTabla & ".codplant=scapla.codplant) "
    SQL = SQL & " LEFT JOIN sgrupl AS grupo ON scapla.codgrupl=grupo.codgrupl "
    If enlaza Then
        SQL = SQL & " WHERE " & NombreTabla & ".codusu = " & vUsu.Codigo
    Else
        SQL = SQL & " WHERE " & NombreTabla & ".codusu = -1"
    End If

    SQL = SQL & " ORDER BY " & NombreTabla & ".codusu, " & NombreTabla & ".codplant"
    MontaSQLCarga = SQL
End Function


'Private Sub BotonBuscar()
'    If Modo <> 1 Then
'        LimpiarCampos
'        PonerModo 1
'        'Ponemos el grid lineasfacturas enlazando a ningun sitio
'        CargaGrid False
'        CargaTxtAux False, False
'    Else
'        'Ya estamos en Modo de Busqueda
''        HacerBusqueda
'        If Data1.Recordset.EOF Then
''            Text1(kCampo).Text = ""
''            Text1(kCampo).BackColor = vbYellow
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'End Sub


Private Sub BotonModificar()
    PonerModo 4
    CargaTxtAux True, True
End Sub


'Private Function DatosOk() As Boolean
'Solo se actualiza el campo de Existencia Real
'    txtAux.Text = Trim(txtAux.Text)
'    If txtAux.Text <> "" And EsNumerico(txtAux.Text) Then
'        DatosOk = True
'    Else
'        DatosOk = False
'    End If
'End Function



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function ActualizarCantidad() As Boolean
'Actualizar en la tabla temporal tmpscapla la cantidad de cada plantilla que
'vamos a cargar en las lineas de la Oferta
Dim SQL As String
'Dim valor As Currency
On Error GoTo EActualizar
'
'        Conn.BeginTrans
        'Actualizar la Tabla: tmpscapla con la cantidad introducida
'        '-------------------------------------------------------
'        ADonde = "Modificando datos de Inventario (Tabla: sinven)."
'        valor = ImporteFormateado(txtAux.Text) 'cantidad
        SQL = "UPDATE " & NombreTabla & " SET cantidad = " & DBSet(txtAux.Text, "N")
        SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND codplant=" & Data1.Recordset!codPlant
    
        conn.Execute SQL
'
EActualizar:
    If Err.Number <> 0 Then
'         'Hay error , almacenamos y salimos
'          SQL = "Actualizando Diferencias de Inventario." & vbCrLf & "--------------------------------------------" & vbCrLf
'          SQL = SQL & ADonde
          SQL = "Actualizando Cantidad a Cargar"
          MuestraError Err.Number, SQL, Err.Description
'          Conn.RollbackTrans
          ActualizarCantidad = False
    Else
        ActualizarCantidad = True
'        Conn.CommitTrans
    End If
End Function



Private Function CargarDatos()
'Al salir de la aplicacion se borran los datos de la tabla temporal pero no se cargan
Dim SQL As String
On Error GoTo ECargaDatos

    '------------- AHORA
    SQL = "DELETE from " & NombreTabla & " where codusu= " & vUsu.Codigo
    conn.Execute SQL
            
    CargaDatosTMPplantilla
    'CargaGrid
    'CargaImportes
    Exit Function
ECargaDatos:
        MuestraError Err.Number, "Carga Plantillas a Ofertas", Err.Description
End Function



Public Function CargaDatosTMPplantilla() As Boolean
On Error GoTo ECargaDatosTMP
Dim SQL As String
Dim RS As ADODB.Recordset

    'Obtener las plantillas disponibles de la tabla scapla
    SQL = "select codplant from scapla "
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'Insertar en la temporal las plantillas disponibles
        SQL = "INSERT INTO " & NombreTabla & " (codusu, codplant, cantidad) VALUES ("
        SQL = SQL & vUsu.Codigo & ", " & RS!codPlant & ", " & ValorNulo & ")"
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    CargaDatosTMPplantilla = True
Exit Function
ECargaDatosTMP:
    'CargaDatosConExt = 2
    CargaDatosTMPplantilla = False
    MuestraError Err.Number, "Gargando datos temporales. Cta: ", Err.Description
End Function

