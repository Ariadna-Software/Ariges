VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLlamadas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Llamadas"
   ClientHeight    =   6780
   ClientLeft      =   -105
   ClientTop       =   285
   ClientWidth     =   14220
   Icon            =   "frmLlamadas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLlamadas.frx":000C
      Height          =   5670
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   10001
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
   Begin VB.CommandButton cmdCancelar1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   13080
      TabIndex        =   0
      Top             =   6360
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   2
      Top             =   6240
      Width           =   1755
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
         TabIndex        =   3
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
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
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
   

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
    If adodc1.Recordset.RecordCount <= 0 Then
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
    CadenaDesdeOtroForm = "`feholla`=" & DBSet(adodc1.Recordset.Fields(0), "FH") & " and `usuario`=" & DBSet(adodc1.Recordset.Fields(1), "T")
    N = adodc1.Recordset.AbsolutePosition
    frmLLamadasDatos2.SoloVer = False
    frmLLamadasDatos2.vModo = 4
    frmLLamadasDatos2.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        CadenaConsulta = "Select * from sllama"
        CargaGrid ElSQL
        If N > 1 Then SituarDataPosicion adodc1, N, Me.lblIndicador
    End If
End Sub




Private Sub BotonEliminar()
Dim SQL As String
On Error GoTo Error2

    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub

    'Si no es de usuario solo dejamos
    If adodc1.Recordset!Usuario <> vUsu.Login Then
        MsgBox "Llamada no insertada por usuario actual", vbExclamation
        If vUsu.Nivel > 1 Then Exit Sub
    End If
    
    '### a mano
    SQL = "Fecha: " & adodc1.Recordset.Fields(0) & vbCrLf
    SQL = SQL & "Usuario:  " & adodc1.Recordset.Fields(1) & vbCrLf
    SQL = SQL & "Cliente:  " & DBLet(adodc1.Recordset!codClien, "T") & "  " & DBLet(adodc1.Recordset!nomclien, "T") & vbCrLf & vbCrLf
    SQL = SQL & vbCrLf & "¿Seguro que desea eliminar la llamada?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.adodc1.Recordset.AbsolutePosition
        SQL = "Delete from sllama where usuario='" & adodc1.Recordset!Usuario & "' AND feholla = " & DBSet(adodc1.Recordset!feholla, "FH")
        conn.Execute SQL
        CancelaADODC adodc1
        CargaGrid ElSQL
        CancelaADODC Me.adodc1
        SituarDataPosicion Me.adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Marca", Err.Description
End Sub










Private Sub cmdCancelar1_Click()
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If Not adodc1.Recordset.EOF Then BotonModificar
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not adodc1.Recordset.EOF Then 'And Modo = 0 Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
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
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1 'Buscar
        .Buttons(2).Image = 2 'VerTodos
        .Buttons(5).Image = 3   'Botón Añadir Nuevo Registro
        .Buttons(6).Image = 4  'Modificar
        .Buttons(7).Image = 5  'Eliminar
        .Buttons(10).Image = 16  'Botón Imprimir
        .Buttons(11).Image = 15  'Botón Salir
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
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5: mnNuevo_Click
        Case 6: mnModificar_Click
        Case 7: mnEliminar_Click
        Case 10   'Imprimir Listado de Marcas

            frmListado2.Opcion = 22
            frmListado2.Show vbModal
        Case 11: mnSalir_Click
    End Select
End Sub



Private Sub CargaGrid(Optional ByVal SQL As String)
Dim I As Byte

    
    lblIndicador.Caption = "Leyendo ..."
    lblIndicador.Refresh
    
   

    
    
    CadenaConsulta = "select feholla,usuario,codclien,nomclien,telefono,codtraba,nomtraba,nomllama1 from"
    CadenaConsulta = CadenaConsulta & " sllama,sllama1 where sllama.codllama1=sllama1.codllama1"
    If SQL <> "" Then
        CadenaConsulta = CadenaConsulta & " AND " & SQL
      
    End If
    
    SQL = CadenaConsulta
    SQL = SQL & " ORDER BY feholla desc"

    CargaGridGnral DataGrid1, Me.adodc1, SQL, False

    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Fecha/Hora"
        DataGrid1.Columns(I).Width = 1700
        DataGrid1.Columns(I).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Usuario"
        DataGrid1.Columns(I).Width = 1175
            
    I = 2
        DataGrid1.Columns(I).Caption = "Cod.Cli."
        DataGrid1.Columns(I).Width = 970
                    
    I = 3
        DataGrid1.Columns(I).Caption = "Cliente"
        DataGrid1.Columns(I).Width = 2900
               
    I = 4
        DataGrid1.Columns(I).Caption = "Teléfono"
        DataGrid1.Columns(I).Width = 1475
                          
               
               
    I = 5
        DataGrid1.Columns(I).Caption = "Trab."
        DataGrid1.Columns(I).Width = 800
    'select feholla,usuario,codclien,nomclien,codtraba,nomtraba,nomllama1 from"
    
    I = 6
        DataGrid1.Columns(I).Caption = "Nombre"
        DataGrid1.Columns(I).Width = 2190
    
    I = 7
        DataGrid1.Columns(I).Caption = "Motivo"
        DataGrid1.Columns(I).Width = 2175
    
    

   
   'No permitir cambiar tamaño de columnas
   For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
   Next I
   
    'Habilitamos botones Modificar y Eliminar
   'If Toolbar1.Buttons(6).Enabled = True Then
        Toolbar1.Buttons(6).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        mnModificar.Enabled = Not adodc1.Recordset.EOF
        mnEliminar.Enabled = Not adodc1.Recordset.EOF
   ' End If
   
   DataGrid1.ScrollBars = dbgAutomatic
   
   'Actualizar indicador
   If Not adodc1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
   PonerOpcionesMenu
End Sub





Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


