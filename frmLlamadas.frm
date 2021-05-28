VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLlamadas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Llamadas"
   ClientHeight    =   10470
   ClientLeft      =   -105
   ClientTop       =   -15
   ClientWidth     =   16215
   Icon            =   "frmLlamadas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   16215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   4
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   5
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLlamadas.frx":000C
      Height          =   8865
      Left            =   135
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   765
      Width           =   15930
      _ExtentX        =   28099
      _ExtentY        =   15637
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   14985
      TabIndex        =   0
      Top             =   9855
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   9795
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
         TabIndex        =   2
         Top             =   195
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
Attribute VB_Name = "frmLlamadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CadenaConsulta As String



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
Dim ElSQL As String
Dim PrimeraVez As Boolean

Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
    
    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    

    


    'DataGrid1.Enabled = b

  

    
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
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
   

    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = b
    
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub BotonAnyadir()
    
    CadenaDesdeOtroForm = ""
    frmLLamadasDatos2.SoloVer = False
    frmLLamadasDatos2.vModo = 3
    frmLLamadasDatos2.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        CadenaConsulta = "Select * from sllama"
        CargaGrid ElSQL
    End If
End Sub


Private Sub BotonBuscar()
  CadenaDesdeOtroForm = ""
    frmLLamadasDatos2.SoloVer = False
    frmLLamadasDatos2.vModo = 1
    frmLLamadasDatos2.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        CargaGrid CadenaDesdeOtroForm
    End If
    CadenaDesdeOtroForm = ""
End Sub


Private Sub BotonVerTodos()
On Error Resume Next
    ElSQL = ""
    CargaGrid ElSQL
    If Adodc1.Recordset.RecordCount <= 0 Then
         'MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
         'MsgBox "No hay ningún registro en la tabla sllama", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
Dim N As Long
    CadenaDesdeOtroForm = "`feholla`=" & DBSet(Adodc1.Recordset.Fields(0), "FH") & " and `usuario`=" & DBSet(Adodc1.Recordset.Fields(1), "T")
    N = Adodc1.Recordset.AbsolutePosition
    frmLLamadasDatos2.SoloVer = False
    frmLLamadasDatos2.vModo = 4
    frmLLamadasDatos2.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        CadenaConsulta = "Select * from sllama"
        CargaGrid ElSQL
        If N > 1 Then SituarDataPosicion Adodc1, N, Me.lblIndicador
    End If
End Sub




Private Sub BotonEliminar()
Dim SQL As String
On Error GoTo Error2

    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub

    'Si no es de usuario solo dejamos
    If Adodc1.Recordset!Usuario <> vUsu.Login Then
        MsgBox "Llamada no insertada por usuario actual", vbExclamation
        If vUsu.Nivel > 1 Then Exit Sub
    End If
    
    '### a mano
    SQL = "Fecha: " & Adodc1.Recordset.Fields(0) & vbCrLf
    SQL = SQL & "Usuario:  " & Adodc1.Recordset.Fields(1) & vbCrLf
    SQL = SQL & "Cliente:  " & DBLet(Adodc1.Recordset!codClien, "T") & "  " & DBLet(Adodc1.Recordset!NomClien, "T") & vbCrLf & vbCrLf
    SQL = SQL & vbCrLf & "¿Seguro que desea eliminar la llamada?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Adodc1.Recordset.AbsolutePosition
        SQL = "Delete from sllama where usuario='" & Adodc1.Recordset!Usuario & "' AND feholla = " & DBSet(Adodc1.Recordset!feholla, "FH")
        conn.Execute SQL
        CancelaADODC Adodc1
        CargaGrid ElSQL
        CancelaADODC Me.Adodc1
        SituarDataPosicion Me.Adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Marca", Err.Description
End Sub










Private Sub cmdCancelar1_Click()
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If Not Adodc1.Recordset.EOF Then BotonModificar
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
         
'            'Cadena consulta
'            CadenaConsulta = "Select * from sllama"
'            ElSQL = "  date(feholla)='" & Format(Now, FormatoFecha) & "'"
'            CargaGrid ElSQL
'            PonerModo 2
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    PrimeraVez = True
    
    'ICONOS de La toolbar.
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

    'Cadena consulta
    CadenaConsulta = "Select * from sllama"
    ElSQL = "  date(feholla)='" & Format(Now, FormatoFecha) & "'"
    
    CargaGrid ElSQL
    'PonerModo 2
      

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
            frmListado2.Opcion = 22
            frmListado2.Show vbModal
    End Select
End Sub


Private Sub CargaGrid(Optional ByVal SQL As String)
Dim i As Byte

    
    lblIndicador.Caption = "Leyendo ..."
    lblIndicador.Refresh
    
    
    CadenaConsulta = "select feholla,usuario,codclien,nomclien,telefono,codtraba,nomtraba,nomllama1 from"
    CadenaConsulta = CadenaConsulta & " sllama,sllama1 where sllama.codllama1=sllama1.codllama1"
    If SQL <> "" Then
        CadenaConsulta = CadenaConsulta & " AND " & SQL
      
    End If
    
    SQL = CadenaConsulta
    SQL = SQL & " ORDER BY feholla desc"

    CargaGridGnral DataGrid1, Me.Adodc1, SQL, False

    
    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "Fecha/Hora"
        DataGrid1.Columns(i).Width = 2250
        DataGrid1.Columns(i).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Usuario"
        DataGrid1.Columns(i).Width = 1175
            
    i = 2
        DataGrid1.Columns(i).Caption = "Código"
        DataGrid1.Columns(i).Width = 970
                    
    i = 3
        DataGrid1.Columns(i).Caption = "Cliente"
        DataGrid1.Columns(i).Width = 3430
               
    i = 4
        DataGrid1.Columns(i).Caption = "Teléfono"
        DataGrid1.Columns(i).Width = 1475
                          
    i = 5
        DataGrid1.Columns(i).Caption = "Trab."
        DataGrid1.Columns(i).Width = 800
    'select feholla,usuario,codclien,nomclien,codtraba,nomtraba,nomllama1 from"
    
    i = 6
        DataGrid1.Columns(i).Caption = "Nombre"
        DataGrid1.Columns(i).Width = 2690
    
    i = 7
        DataGrid1.Columns(i).Caption = "Motivo"
        DataGrid1.Columns(i).Width = 2575
    
    

   'No permitir cambiar tamaño de columnas
   For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
   Next i
   
    'Habilitamos botones Modificar y Eliminar
   'If Toolbar1.Buttons(6).Enabled = True Then
        Toolbar1.Buttons(2).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(3).Enabled = Not Adodc1.Recordset.EOF
        mnModificar.Enabled = Not Adodc1.Recordset.EOF
        mnEliminar.Enabled = Not Adodc1.Recordset.EOF
   ' End If
   
   DataGrid1.ScrollBars = dbgAutomatic
   '++
   DataGrid1.RowHeight = 350
   
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
   PonerOpcionesMenu
End Sub





Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


