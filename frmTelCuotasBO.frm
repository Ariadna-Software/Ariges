VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTelBolbaite 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   Icon            =   "frmTelCuotasBO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboOperadora 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTelCuotasBO.frx":000C
      Left            =   1440
      List            =   "frmTelCuotasBO.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   4320
      TabIndex        =   4
      Text            =   "Descripcion"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ComboBox CboMensual 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTelCuotasBO.frx":0010
      Left            =   3360
      List            =   "frmTelCuotasBO.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox CboAcciones 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTelCuotasBO.frx":0014
      Left            =   2400
      List            =   "frmTelCuotasBO.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "Codigo"
      Top             =   4920
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Text            =   "Descripcion"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTelCuotasBO.frx":0018
      Height          =   4710
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   585
      Width           =   10095
      _ExtentX        =   17806
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
      Left            =   9180
      TabIndex        =   6
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7980
      TabIndex        =   5
      Top             =   5400
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   8
      Top             =   5280
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
         TabIndex        =   9
         Top             =   240
         Width           =   1680
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10365
      _ExtentX        =   18283
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
      Height          =   330
      Left            =   360
      Top             =   5400
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
Attribute VB_Name = "frmTelBolbaite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
'
'Autor: DAVID
'Fecha creación: 18/10/2013
'Fecha modificacion:
'+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

Option Explicit


'Es comun para las cuotas y para los conceptos ya que comparten muchas carctarisiticas
' operadora, solo se modifica o elimina
'y ahorrar un poco de fuente
Public QueOpcion As Byte
    '0  Cuotas
    '1  conceptos llamadas
    '2  Cuotas propias coooperativa
    '3  Cargos varios

Private CadenaConsulta As String

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
    
    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    
    
    cboOperadora.Enabled = Modo = 1 Or (Modo = 3)
    
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    txtAux(2).visible = Not b And QueOpcion <> 1 And QueOpcion <> 3
    Me.cboOperadora.visible = Not b
    Me.CboAcciones.visible = Not b And QueOpcion < 2  'visibles 0,1
    Me.CboMensual.visible = Not b And QueOpcion <> 2
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    

    
    'Si estamos insertando o busqueda
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    BloquearTxt txtAux(1), (Modo <> 3 And Modo <> 1)  'tampoco podemos cambiar la cuota
        
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
    
    b = b And QueOpcion = 2 'Solo el 2
    'Añadir
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    
    
     b = (Modo = 2)
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
    AnyadirLinea DataGrid1, adodc1
    'De momento solo para queopcion=2
    
    'Limpiamos los campos para insertar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    'por defecto control de lotes vale NO
    'Me.CboAcciones.ListIndex = 0
    'Me.CboMensual.ListIndex = 0

    anc = ObtenerAlto(DataGrid1)
    LLamaLineas anc, 3

    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub


Private Sub BotonBuscar()
    CargaGrid "1= -1"
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    Me.CboAcciones.ListIndex = -1
    Me.CboMensual.ListIndex = -1
    Me.cboOperadora.ListIndex = -1
    LLamaLineas 790, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
    On Error Resume Next

    CargaGrid ""
    If adodc1.Recordset.RecordCount <= 0 Then
         MsgBox "No hay ningún registro en la tabla ", vbInformation
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

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    SituarCombo Me.cboOperadora, CByte(adodc1.Recordset!Operadora)
'    If LCase(DataGrid1.Columns(0).Text) = "movistar" Then
'        Me.cboOperadora.ListIndex = 0
'    Else
'        Me.cboOperadora.ListIndex = 1
'    End If
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(1).Text
    txtAux(1).Text = DataGrid1.Columns(2).Text
    
    If QueOpcion = 0 Then
        txtAux(2).Text = DataGrid1.Columns(7).Text
         PosicionarCombo Me.CboAcciones, CInt(DataGrid1.Columns(3).Text)
        PosicionarCombo CboMensual, CInt(DataGrid1.Columns(5).Text)
    ElseIf QueOpcion = 1 Then
        PosicionarCombo CboMensual, CInt(DataGrid1.Columns(3).Text)
        PosicionarCombo CboAcciones, CInt(DataGrid1.Columns(5).Text)
    ElseIf QueOpcion = 2 Then
        txtAux(2).Text = DataGrid1.Columns(3).Text
    Else
        PosicionarCombo CboMensual, CInt(DataGrid1.Columns(4).Text)
    End If
    
    anc = ObtenerAlto(DataGrid1)
    LLamaLineas anc, 4
   
    'Como es modificar
'    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    Me.CboMensual.Top = alto - 15
    Me.CboAcciones.Top = alto - 15
    Me.cboOperadora.Top = alto - 15
'    txtAux(0).Left = DataGrid1.Left + 340
'    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
'    Me.CboCtrLotes.Left = txtAux(1).Left + txtAux(1).Width + 55
End Sub

Private Sub BotonEliminar()
Dim SQL As String

    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub

    '### a mano
    If QueOpcion <> 1 Then
        SQL = "cuota"
    Else
        SQL = "concepto"
    End If
    SQL = "¿Seguro que desea eliminar la " & SQL & "?" & vbCrLf
    SQL = SQL & vbCrLf & "Operadora: " & adodc1.Recordset!laoperadora
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Descripción: " & adodc1.Recordset.Fields(2)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
        If QueOpcion = 0 Then
            SQL = "Delete from tel_descuentoscuotas where  codigo_de_cuota=" & DBSet(adodc1.Recordset!codigo_de_cuota, "T")
        ElseIf QueOpcion = 1 Then
            SQL = "Delete from el_conceptosllamadas where Codigo_de_tipo_de_trafico=" & DBSet(adodc1.Recordset!Codigo_de_tipo_de_trafico, "N")
        ElseIf QueOpcion = 2 Then
            SQL = "Delete from stfnocuotaspropias where  codigoCuota=" & DBSet(adodc1.Recordset!codigoCuota, "T")
        Else
            SQL = "Delete from tel_cargo_varios where  CodigoVario=" & DBSet(adodc1.Recordset!CodigoVario, "T")
        End If
        SQL = SQL & " AND operadora= " & adodc1.Recordset!Operadora
        
        conn.Execute SQL
        CancelaADODC Me.adodc1
        CargaGrid ""
'        CancelaADODC Me.adodc1
        SituarDataPosicion Me.adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar cuota.", Err.Description
End Sub





Private Sub CboAcciones_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub CboMensual_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub


Private Sub cboOperadora_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim I As String
Dim cadB As String

    On Error GoTo EAceptar

    Select Case Modo
        Case 3 'Insertar
            If DatosOk Then
                If InsertModifica Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If
        
        Case 4  'Modificar
            If DatosOk Then
                
                    If InsertModifica Then
                        TerminaBloquear
                        I = adodc1.Recordset.Fields(0)
                        PonerModo 2
                        CancelaADODC Me.adodc1
                        CargaGrid
                        adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & DBSet(I, "T"))
                    End If
                    PonerFocoGrid DataGrid1
           
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


Private Sub cmdCancelar_Click()
    On Error Resume Next

    Select Case Modo
        Case 1 'Busqueda
            CargaGrid
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            TerminaBloquear
            Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    End Select
    
    PonerModo 2
    PonerFocoGrid DataGrid1
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub DataGrid1_DblClick()
    If Modo = 2 Then BotonModificar
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     If Not adodc1.Recordset.EOF Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub MontaSQL()

    If QueOpcion = 0 Then

        CadenaConsulta = "codigo_de_cuota,Descripcion_de_cuota,accion,"
        CadenaConsulta = CadenaConsulta & "if(accion=0,'NADA',if (accion=1,'Eliminar',if(accion=2,'Refacturar x importe','Refacturar x descuento'))),"
        CadenaConsulta = CadenaConsulta & "Mensual,if(Mensual=0,'Mes','Prorratea'),Valor,operadora from tel_descuentoscuotas"
        CadenaConsulta = CadenaConsulta & ",stfnooperador where codoperador = operadora"
        
        
    ElseIf QueOpcion = 1 Then
        CadenaConsulta = "Codigo_de_tipo_de_trafico,Tipo_de_trafico,refacturar,if(refacturar=0,'','Si'),TraficoSegundos,if(TraficoSegundos=0,'','Si'),operadora"
        CadenaConsulta = CadenaConsulta & " FROM tel_conceptosllamadas,stfnooperador where codoperador = operadora"
        
        
    ElseIf QueOpcion = 2 Then
        CadenaConsulta = "codigoCuota,stfnocuotaspropias.nombre,importe,operadora from stfnocuotaspropias,stfnooperador where codoperador = operadora"
    Else
        CadenaConsulta = "CodigoVario ,Descripcion ,if(CargarEnFactura=0,'','Si'),CargarEnFactura,operadora FROM tel_cargo_varios "
        CadenaConsulta = CadenaConsulta & ",stfnooperador where codoperador = operadora"
    End If
    
    CadenaConsulta = "select stfnooperador.nombre LaOperadora," & CadenaConsulta

  

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
    
    
    
    
    Modo = 0
    
    'Cadena consulta
    'tel_descuentoscuotas  tel_conceptosllamadas stfnocuotaspropias
    If QueOpcion = 0 Then
        Caption = "Cuotas telefonía recalculables"
        
        Me.DataGrid1.Width = 11095
        
        Me.txtAux(0).Tag = "codigo_de_cuota|T|N|||@@|codigo_de_cuota|||"
        Me.txtAux(1).Tag = "Descripcion_de_cuota|T|N|||@@|Descripcion_de_cuota|||"
        Me.txtAux(2).Tag = "Valor|N|N|0||@@|Valor|#0.0000||"
        Me.CboAcciones.Tag = "accion|N|N|0||@@|accion|0||"
        Me.CboMensual.Tag = "mensual|N|N|0||@@|mensual|0||"
        
    ElseIf QueOpcion = 1 Then
    
    
        Caption = "Conceptos telefonía refacturables"
        
        Me.txtAux(0).Tag = "Codigo_de_tipo_de_trafico|T|N|||@@|Codigo_de_tipo_de_trafico|||"
        Me.txtAux(1).Tag = "Descripcion_de_cuota|T|N|||@@|Descripcion_de_cuota|||"
        Me.txtAux(2).Tag = ""
        Me.CboAcciones.Tag = "accion|N|N|0||@@|accion|0||"
        Me.DataGrid1.Width = 8995
        Me.txtAux(2).Tag = "Valor|N|N|0|||@@|Valor|#0.0000||"
        Me.CboMensual.Tag = "refacturar|N|N|0||@@|refacturar|0||"
        Me.CboAcciones.Tag = "TraficoSegundos|N|N|0||@@|TraficoSegundos|0||"
    ElseIf QueOpcion = 2 Then
        Caption = "Cuotas propias cooperativa"
        
        Me.txtAux(0).Tag = "codigoCuota|T|N|||@@|codigoCuota|||"
        Me.txtAux(1).Tag = "nombre|T|N|||@@|nombre|||"
        Me.txtAux(2).Tag = "importe|N|N|0||@@|importe|#0.0000||"
        
        Me.DataGrid1.Width = 7995
    
    Else
        Caption = "Cargos VARIOS"
        'tel_cargo_varios CodigoVario Descripcion CargarEnFactura
        Me.txtAux(0).Tag = "codigoCuota|T|N|||@@|CodigoVario|||"
        Me.txtAux(1).Tag = "nombre|T|N|||@@|Descripcion|||"
        Me.txtAux(2).Tag = ""
        
        Me.DataGrid1.Width = 7995
    End If
    cboOperadora.Tag = "Operadora|N|N|0||@@|operadora|0||"
    CadenaConsulta = "tel_descuentoscuotas|tel_conceptosllamadas|stfnocuotaspropias|tel_cargo_varios|"
    For Modo = 0 To 2
        Me.txtAux(Modo).Tag = Replace(Me.txtAux(Modo).Tag, "@@", RecuperaValor(CadenaConsulta, QueOpcion + 1), 1)
    Next
    Me.CboAcciones.Tag = Replace(CboAcciones.Tag, "@@", RecuperaValor(CadenaConsulta, QueOpcion + 1), 1)
    Me.CboMensual.Tag = Replace(CboMensual.Tag, "@@", RecuperaValor(CadenaConsulta, QueOpcion + 1), 1)
    Me.cboOperadora.Tag = Replace(cboOperadora.Tag, "@@", RecuperaValor(CadenaConsulta, QueOpcion + 1), 1)
    'accion|N|N|||tel_descuentoscuotas|accion||
    CargarCombo_Tabla cboOperadora, "stfnoOperador", "codoperador", "nombre", "codoperador <=3"
    
    Modo = 0
    Me.cboOperadora.ListIndex = 0
    Me.Width = DataGrid1.Width + 360
    Me.cmdCancelar.Left = Me.Width - Me.cmdCancelar.Width - 240
    Me.cmdAceptar.Left = Me.cmdCancelar.Left - 240 - Me.cmdAceptar.Width
    CargaGrid
    CargaCombo
    
    PonerModo 2
    
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
'        Case 10 'Imprimir listado de Formas de Envío
'                Me.Hide
'                AbrirListado (23) 'OpcionListado=23
'                Me.Show vbModal
        Case 11: mnSalir_Click  'Salir
    End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
Dim b As Boolean
Dim tots As String

    b = DataGrid1.Enabled

    'Carga el sql
    MontaSQL
    

    'EL SQL LLEVA AND
    If SQL <> "" Then SQL = " AND " & SQL
    SQL = CadenaConsulta & SQL

    SQL = SQL & " ORDER BY operadora,"
    If QueOpcion = 0 Then
        SQL = SQL & "codigo_de_cuota"
    ElseIf QueOpcion = 1 Then
        SQL = SQL & "Codigo_de_tipo_de_trafico"
    ElseIf QueOpcion = 2 Then
        SQL = SQL & "codigoCuota"
    Else
        SQL = SQL & "CodigoVario"
    End If
    
    CargaGridGnral DataGrid1, Me.adodc1, SQL, False

    '### a mano
    If QueOpcion = 0 Then
        tots = "S|txtAux(0)|T|Cod.|800|;S|txtAux(1)|T|Descripción|4500|;N||||0|;"
        tots = tots & "S|CboAcciones|C|Accion|1900|;N||||0|;S|CboMensual|C|Aplicar|1000|;S|txtAux(2)|T|Valor|1200|;N||||0|;"
        
    ElseIf QueOpcion = 1 Then
        tots = "S|txtAux(0)|T|Cod.|800|;S|txtAux(1)|T|Descripción|4500|;N||||0|;S|CboMensual|C|Refacturar|1000|;"
        tots = tots & "N||||0|;S|CboAcciones|C|Seg.|800|;N||||0|;"
    
    ElseIf QueOpcion = 2 Then
        
        tots = "S|txtAux(0)|T|Cod.|800|;S|txtAux(1)|T|Descripción|4500|;S|txtAux(2)|T|Importe|1000|;N||||0|;"
        
    Else
        'tel_cargo_varios CodigoVario Descripcion CargarEnFactura
        tots = "S|txtAux(0)|T|Cod.|800|;S|txtAux(1)|T|Descripción|4500|;S|CboMensual|C|Facturar|1000|;N||||0|;N||||0|;"
    End If
        
         
    tots = "S|cboOperadora|C|Operadora|1100|;" & tots
        
    arregla tots, DataGrid1, Me

    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
End Sub

Private Sub CargaCombo()
    'Carga la lista de impresión de etiquetas
    
    Me.CboMensual.Clear
    
    If QueOpcion = 0 Then
    
        CboMensual.AddItem "Mes"
        CboMensual.ItemData(CboMensual.NewIndex) = 0
        
        CboMensual.AddItem "Prorratea"
        CboMensual.ItemData(CboMensual.NewIndex) = 1
    ElseIf QueOpcion = 1 Or QueOpcion = 3 Then
        CboMensual.AddItem "No"
        CboMensual.ItemData(CboMensual.NewIndex) = 0
        
        CboMensual.AddItem "Si"
        CboMensual.ItemData(CboMensual.NewIndex) = 1
    End If
    
    'if (accion=1,'Eliminar',if(accion=2,'Refacturar x importe','Refacturar x descuento')),"
    CboAcciones.Clear
    If QueOpcion = 0 Then
        CboAcciones.AddItem "NADA"
        CboAcciones.ItemData(CboAcciones.NewIndex) = 0
        CboAcciones.AddItem "Eliminar"
        CboAcciones.ItemData(CboAcciones.NewIndex) = 1
        CboAcciones.AddItem "Refacturar x importe"
        CboAcciones.ItemData(CboAcciones.NewIndex) = 2
        CboAcciones.AddItem "Refacturar x descuento"
        CboAcciones.ItemData(CboAcciones.NewIndex) = 3
    ElseIf QueOpcion = 1 Then
        'TraficoSegundos
        CboAcciones.AddItem "No"
        CboAcciones.ItemData(CboAcciones.NewIndex) = 0
        
        CboAcciones.AddItem "Si"
        CboAcciones.ItemData(CboAcciones.NewIndex) = 1
    End If
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
    
    If Index = 0 And QueOpcion = 2 Then
        'SOLO PARA CUOTAS PROPIAS
        If Not PonerFormatoEntero(Me.txtAux(0)) Then txtAux(0).Text = ""
    End If
        
    If Index = 2 Then
        If Not PonerFormatoDecimal(txtAux(Index), 2) Then txtAux(Index).Text = ""
    End If
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    If Me.CboAcciones.ListIndex = 3 Then

        'Es sobre dto VALOR no puede ser superirior a 100
        If ImporteFormateado(Me.txtAux(2).Text) >= 100 Then
            MsgBox "Importe no puede ser >= de 100", vbExclamation
            Exit Function
        End If
    End If

    If Modo = 3 Then
        If cboOperadora.ListIndex < 0 Then
            MsgBox "Seleccione operadora", vbExclamation
            b = False
        End If
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



Private Function InsertModifica() As Boolean
Dim C As String

    InsertModifica = False
    If Modo = 3 Then
        'Insertar
        
        
        
        If QueOpcion <> 2 Then Exit Function
        
        
        
        'stfnocuotaspropias (operadora,codigoCuota,nombre,importe)
        C = "INSERT INTO stfnocuotaspropias (operadora,codigoCuota,nombre,importe)"
        C = C & " VALUES (" & Me.cboOperadora.ItemData(cboOperadora.ListIndex) & ","
        C = C & DBSet(Me.txtAux(0).Text, "T") & "," & DBSet(Me.txtAux(1).Text, "T") & ","
        C = C & DBSet(Me.txtAux(2).Text, "N", "N") & ")"
    
    Else
        'MOdificar
        If QueOpcion = 0 Then
            C = "UPDATE tel_descuentoscuotas SET "
            C = C & "Descripcion_de_cuota=" & DBSet(Me.txtAux(1).Text, "T")
            C = C & ",Valor=" & DBSet(Me.txtAux(2).Text, "N")
            C = C & ",accion=" & Me.CboAcciones.ItemData(CboAcciones.ListIndex)
            C = C & ",Mensual=" & Me.CboMensual.ItemData(CboMensual.ListIndex)
            C = C & " WHERE operadora =" & Me.cboOperadora.ItemData(cboOperadora.ListIndex)
            C = C & " and codigo_de_cuota =" & DBSet(Me.txtAux(0), "T")
    
        ElseIf QueOpcion = 1 Then
            C = "UPDATE tel_conceptosllamadas SET "
            C = C & "refacturar=" & Me.CboMensual.ItemData(CboMensual.ListIndex)
            C = C & ",TraficoSegundos=" & Me.CboAcciones.ItemData(CboAcciones.ListIndex)
            C = C & " WHERE operadora =" & Me.cboOperadora.ItemData(cboOperadora.ListIndex) & " and Codigo_de_tipo_de_trafico =" & DBSet(Me.txtAux(0), "T")
            
        ElseIf QueOpcion = 2 Then
            C = "UPDATE stfnocuotaspropias SET "
            C = C & "importe=" & DBSet(Me.txtAux(2).Text, "N")
            C = C & " WHERE operadora =" & Me.cboOperadora.ItemData(cboOperadora.ListIndex) & " and codigoCuota =" & DBSet(Me.txtAux(0), "T")
            
        ElseIf QueOpcion = 3 Then
            
            'tel_cargo_varios CodigoVario Descripcion CargarEnFactura
            C = "UPDATE tel_cargo_varios SET "
            C = C & "CargarEnFactura=" & Me.CboMensual.ItemData(CboMensual.ListIndex)
            C = C & " WHERE operadora =" & Me.cboOperadora.ItemData(cboOperadora.ListIndex) & " and CodigoVario =" & DBSet(Me.txtAux(0), "T")
            
        End If
    
    End If
    If Ejecutar(C, False) Then InsertModifica = True
End Function
