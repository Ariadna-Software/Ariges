VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelDtoCuotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descuentos cuotas"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11250
   ClipControls    =   0   'False
   Icon            =   "frmTelDtoCuotas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "TipoRef|N|S|0|9999|tel_desc_cuotas|Refacturar|0||"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   8760
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Valor|N|S|||tel_desc_cuotas|valor|##0.00||"
      Text            =   "Valor"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "% Coop|N|N|0|100.01|tel_desc_cuotas|Porcentaje|##0.00||"
      Text            =   "fecha"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   2640
      MaxLength       =   16
      TabIndex        =   2
      Tag             =   "%Ope|N|N|0|100|tel_desc_cuotas|PorcentajeOperador|##0.00||"
      Text            =   "zona"
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Descripcion|T|N|||tel_desc_cuotas|DescCuota|||"
      Text            =   "numlote"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   12
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
         TabIndex        =   13
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   480
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "codigo|T|N|||tel_desc_cuotas|CodCuota||S|"
      Text            =   "codartic codarti"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   5280
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   5280
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
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
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   4680
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9480
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
      Bindings        =   "frmTelDtoCuotas.frx":000C
      Height          =   4425
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   10935
      _ExtentX        =   19288
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
End
Attribute VB_Name = "frmTelDtoCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Modo As Byte
Dim kCampo As Integer

Dim EsBusqueda As Boolean

Dim CadenaConsulta As String
Dim CadenaBusqueda As String



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
                    EsBusqueda = True
                    CargaGrid True
                    mnNuevo_Click
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk Then
                 If ModificaDesdeFormulario(Me, 3) Then
                     
                     NumReg = Data1.Recordset.AbsolutePosition
                     
                     CancelaADODC Me.Data1
                     CargaGrid True
                     LLamaLineas 10, 2
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




Private Sub cmdCancelar_Click()
Dim Indicador As String

    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            
            LLamaLineas 10, 0
            EsBusqueda = False
           
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            DataGrid1.Enabled = True
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
      
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else
                PonerModo 0
            End If
            LLamaLineas 10, 2
            
        Case 4  'Modificar
            TerminaBloquear
            NumRegElim = Data1.Recordset.AbsolutePosition
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            
            LLamaLineas 10, 2
            DataGrid1.Enabled = True
            SituarDataPosicion Data1, NumRegElim, Indicador
            DeseleccionaGrid Me.DataGrid1
            lblIndicador.Caption = Indicador
            PonerFocoGrid DataGrid1
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub








Private Sub Combo1_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
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
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 15 'Salir
    End With

    CargaCombo


    LimpiarCampos   'Limpia los campos TextBox
   
    DataGrid1.ClearFields
    
    
    
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    

    
    CadenaConsulta = MontaSQLCarga(False)
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
   
        PonerModo 2
 
'    CargaGrid (Modo = 2 Or Modo = 0)
    CargaGrid True
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String
    
    On Error GoTo ECarga
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, False
    
    tots = "S|txtAux(0)|T|Tipo|1000|;S|txtAux(1)|T|Descripcion|4000|;"
    tots = tots & "S|txtAux(2)|T|% Operador|950|;S|txtAux(3)|T|% Coop|950|;"
    tots = tots & "S|Combo1|C|Refacturar|1350|;S|txtAux(4)|T|Importe o %|1150|;N|||||;"
    


    arregla tots, DataGrid1, Me



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



Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim Jj As Integer
Dim b As Boolean

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar
    
    For Jj = 0 To txtAux.Count - 1
     
        txtAux(Jj).Height = DataGrid1.RowHeight
        txtAux(Jj).Top = alto
        txtAux(Jj).visible = b
    Next Jj
    Combo1.visible = b
    Combo1.Top = alto
End Sub






Private Sub mnBuscar_Click()
    BotonBuscar
End Sub



Private Sub mnEliminar_Click()
     '### a mano
    CadenaBusqueda = "¿Seguro que desea eliminar la linea de descuento de cuotas? " & vbCrLf
    CadenaBusqueda = CadenaBusqueda & vbCrLf & "Código: " & Data1.Recordset.Fields(0)
    CadenaBusqueda = CadenaBusqueda & vbCrLf & "Denominación: " & Data1.Recordset.Fields(1)
    
    
    If MsgBox(CadenaBusqueda, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Data1.Recordset.AbsolutePosition
        CadenaBusqueda = "CodCuota=" & DBSet(Me.Data1.Recordset.Fields(0), "T")
        CadenaBusqueda = "Delete from tel_desc_cuotas where " & CadenaBusqueda
        ejecutar CadenaBusqueda, False
        CancelaADODC Me.Data1
        CargaGrid True
        CancelaADODC Me.Data1
        SituarDataPosicion Me.Data1, NumRegElim, CadenaBusqueda
    End If
    CadenaBusqueda = ""
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub



Private Sub mnNuevo_Click()
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
   
    NumRegElim = Val(ObtenerAlto(DataGrid1, 10))
    
    'Obtenemos la siguiente numero de factura
    LimpiarCampos
   
    
    LLamaLineas CSng(NumRegElim), 3
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.index
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
        Case 10 'Imprimir
'            BotonImprimir
            
        Case 11  'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    
    'Modo Buscar
    If Kmodo = 1 Then
        PonerFoco txtAux(0)
    End If
                                 
    BloquearTxt txtAux(0), (Modo = 4)

                   
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

    b = (Modo = 2)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    
    
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
    Combo1.ListIndex = -1
'    Combo2.ListIndex = -1
End Sub


Private Sub Desplazamiento(index As Integer)
'Botones de Desplazamiento de la Toolbar
    DesplazamientoData Data1, index
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
    
    SQL = "select CodCuota,DescCuota,PorcentajeOperador,Porcentaje,"
    SQL = SQL & "if(refacturar=0,'',if(refacturar=1,'Increm €','% Dto'))"
    SQL = SQL & "Texto,Valor,Refacturar  from tel_desc_cuotas "
    SQL = SQL & " where 1=1 "
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
    Else
        SQL = SQL & " AND  codcuota = '-1'"
    End If
    SQL = SQL & " ORDER BY codcuota"
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
        LLamaLineas anc, 1
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


Private Sub BotonVerTodos()
    On Error Resume Next

    EsBusqueda = False
    LimpiarCampos
    
    CadenaConsulta = MontaSQLCarga(True)
    PonerCadenaBusqueda
    PonerFocoGrid DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub BotonModificar()
Dim I As Integer


    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    I = CInt(ObtenerAlto(Me.DataGrid1, 10))
    LLamaLineas CSng(I), 4
    
 
    For I = 0 To 3
        txtAux(I).Text = DBLet(DataGrid1.Columns(I).Text, "T")
    Next I
    I = Val(Data1.Recordset!refacturar)
    Combo1.ListIndex = I
    txtAux(4).Text = DBLet(DataGrid1.Columns(5).Text, "T")
    
'
'    If UCase(DBLet(DataGrid1.Columns(9).Value, "T")) = "SI" Then
'        Combo1.ListIndex = 0
'    Else
'        Combo1.ListIndex = 1
'    End If
'
'    If UCase(DBLet(DataGrid1.Columns(10).Value, "T")) = "NO" Then
'        Combo2.ListIndex = 1
'    Else
'        Combo2.ListIndex = 0
'    End If
    
    DataGrid1.Enabled = False
   
End Sub





Private Function DatosOk() As Boolean
Dim b As Boolean
Dim J As Currency

    On Error GoTo ErrDatosOK

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    CadenaConsulta = ""
    If Combo1.ListIndex = 0 Then
        If Me.txtAux(4).Text <> "" Then CadenaConsulta = "No debe especificar nada en importe / porcentaje"
    Else
        If Me.txtAux(4).Text = "" Then CadenaConsulta = "Indique algun valor para importe / porcentaje"
    End If
    If CadenaConsulta <> "" Then
        MsgBox CadenaConsulta, vbExclamation
        CadenaConsulta = ""
        b = False
    End If
    
    If b And Combo1.ListIndex = 2 Then
        J = ImporteFormateado(txtAux(4).Text)
        If J < -100 Or J > 100 Then
            MsgBox "Descuento porcentual. Valores entre -100 y 100", vbExclamation
            b = False
        End If
    End If
    
    DatosOk = b
    Exit Function
    
ErrDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos OK.", Err.Description
End Function





Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
 '   If chkVistaPrevia = 1 Then
 '       MandaBusquedaPrevia cadB
 '   ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaBusqueda = " AND " & CadB
        CadenaConsulta = MontaSQLCarga(True)
        'CadenaBusqueda = " AND " & cadB
        PonerCadenaBusqueda
 '   End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        Exit Sub
    Else
        PonerModo 2
        PonerCampos
    End If
    LLamaLineas 10, 2
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


Private Sub txtAux_GotFocus(index As Integer)
    ConseguirFoco txtAux(index), Modo
End Sub


Private Sub TxtAux_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.

    

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If index > 0 Then PonerFoco txtAux(index - 1)
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
                If index = 3 Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    SendKeys "{tab}"
                End If
    End Select
End Sub

Private Sub txtAux_KeyPress(index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(index As Integer)

    On Error Resume Next
    
    If Not PerderFocoGnral(txtAux(index), Modo) Then Exit Sub
    
    Select Case index
    
       
        Case 2, 3, 4

            If txtAux(index).Text <> "" Then
                If Not PonerFormatoDecimal(txtAux(index), 3) Then txtAux(index).Text = ""
            End If
              
        
        
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub CargaCombo()
    Me.Combo1.Clear
    Combo1.AddItem "NADA"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Increm. €"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo1.AddItem "Porcentaje"
    Combo1.ItemData(Combo1.NewIndex) = 2
    
End Sub
