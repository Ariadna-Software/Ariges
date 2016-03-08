VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmTipoUnidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Unidad"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmAlmTipoUnidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "EsTrabajo|N|N|||sunida|EsTrabajo||N|"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdAux2 
      Caption         =   "+"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   1200
      TabIndex        =   18
      Tag             =   "Nombre Tipo de unidad|T|N|||sunida|nomunida||N|"
      Text            =   "Descripcion"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Tag             =   "Código Tipo de unidad|N|N|0|99|sunida|codunida|00|S|"
      Text            =   "co"
      Top             =   5640
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.TextBox txtAux2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Tag             =   "Tasa reciclado|N|S|||sunida|tasareciclado|0.00000|N|"
      Text            =   "Tasa"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2160
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Adodc2"
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
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Tag             =   "Tasa reciclado|N|S|||sunida|tasareciclado|0.00000|N|"
      Text            =   "Tasa"
      Top             =   3480
      Width           =   1155
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "Abreviatura|T|N|||sunida|nomunbre||N|"
      Text            =   "Abreviatura"
      Top             =   3480
      Width           =   675
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4260
      TabIndex        =   9
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3060
      TabIndex        =   8
      Top             =   6960
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   120
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "Código Tipo de unidad|N|N|0|99|sunida|codunida|00|S|"
      Text            =   "Codigo"
      Top             =   3480
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   960
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Nombre Tipo de unidad|T|N|||sunida|nomunida||N|"
      Text            =   "Descripcion"
      Top             =   3480
      Width           =   1395
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   4260
      TabIndex        =   14
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   11
      Top             =   6860
      Width           =   2475
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2280
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
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
            Object.ToolTipText     =   "Lineas"
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
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
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
      Bindings        =   "frmAlmTipoUnidad.frx":000C
      Height          =   885
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   540
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1561
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmAlmTipoUnidad.frx":0021
      Height          =   2205
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3889
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
   Begin VB.Label Label1 
      Caption         =   "Costes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   2655
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
      Begin VB.Menu mnMtoLineas 
         Caption         =   "Mantenimiento lineas"
      End
      Begin VB.Menu mnbarra3 
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
      Begin VB.Menu mnbarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAlmTipoUnidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)



Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

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
'-------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas


Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
    
    
    
    ActualizarToolbarGnral Me.Toolbar1, Modo, vModo, 5
    
    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    b = Modo = 1 Or Modo = 3 Or Modo = 4
    txtAux(0).visible = b
    txtAux(1).visible = b
    txtAux(2).visible = b
    txtAux(3).visible = b
    Combo1.visible = b
    
    cmdAceptar.visible = b
    cmdCancelar.visible = b
    b = b Or Modo = 5
    DataGrid1.Enabled = Not b
   
    b = (Modo = 2)
    'Si es regresar    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If b Then
        cmdRegresar.Caption = "&Regresar"
        cmdRegresar.visible = DatosADevolverBusqueda <> ""
    End If
    
    'Si estamos insertando o busqueda
    BloquearTxt txtAux(0), (Modo <> 3 And Modo <> 1)
    
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
    Toolbar1.Buttons(1).Enabled = b 'Buscar
    Me.mnBuscar.Enabled = b
    Toolbar1.Buttons(2).Enabled = b 'Todos
    Me.mnVerTodos.Enabled = b
    Toolbar1.Buttons(9).Enabled = b
    Me.mnMtoLineas.Enabled = b
    If b Then
        b = b And Not DeConsulta
    Else
        b = Modo = 5 And ModificaLineas = 0
    End If
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
    
    
    If Modo = 5 Then

        If ModificaLineas = 2 Then Exit Sub
        AnyadirLinea DataGrid2, Adodc2
        ModificaLineas = 1
        PonerBotonCabecera False
        'Los txts
        txtAux2(0).Text = "": txtAux2(1).Text = "": txtAux2(2).Text = ""
        Campos_2_Visibles True
        anc = ObtenerAlto(DataGrid2, 10)
        LLamaLineas2 anc
        PonerFoco txtAux2(0)
        
    Else
    
        'Situamos el grid al final
        AnyadirLinea DataGrid1, adodc1
          
        anc = ObtenerAlto(DataGrid1, 10)
        
        'Obtenemos la siguiente numero de factura
        LimpiarCampos
        txtAux(0).Text = SugerirCodigoSiguienteStr("sunida", "codunida")
        txtAux(0).Text = Format(txtAux(0).Text, FormatoCod)
    
        LLamaLineas anc, 3
        Combo1.ListIndex = 0 'NO es un TRABAJO
        'Ponemos el foco
        PonerFoco txtAux(0)
    End If
End Sub


Private Sub BotonBuscar()
    CargaGrid "codunida= -1"
    LimpiarCampos
    LLamaLineas 770, 1
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid ""
    If adodc1.Recordset.RecordCount <= 0 Then
         'MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
         MsgBox "No hay ningún registro en la tabla sunida", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
        PonerModo 2
'        adodc1.Recordset.MoveFirst
'        PonerCampos
         DataGrid1.SetFocus
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub BotonModificar()
Dim cad As String
Dim anc As Single
Dim i As Integer

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If Modo = 5 Then
        If Adodc2.Recordset.EOF Then Exit Sub
        If Adodc2.Recordset.RecordCount < 1 Then Exit Sub
        If ModificaLineas = 1 Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    
    
    If Modo = 5 Then
        ModificaLineas = 2
        PonerBotonCabecera False
        'Los txts
        For i = 0 To 2
             txtAux2(i).Text = DataGrid2.Columns(i).Text
         Next i
        txtAux2(2).visible = True
        anc = ObtenerAlto(DataGrid2, 10)
        LLamaLineas2 anc
        PonerFoco txtAux2(2)
    Else
         If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
             i = DataGrid1.Bookmark - DataGrid1.FirstRow
             DataGrid1.Scroll 0, i
             DataGrid1.Refresh
         End If
         
         anc = ObtenerAlto(DataGrid1, 10)
         
         cad = ""
         For i = 0 To 2
             cad = cad & DataGrid1.Columns(i).Text & "|"
         Next i
         'Llamamos al form
         txtAux(0).Text = DataGrid1.Columns(0).Text
         txtAux(1).Text = DataGrid1.Columns(1).Text
         txtAux(2).Text = DataGrid1.Columns(2).Text
         txtAux(3).Text = DataGrid1.Columns(3).Text
  
         If DataGrid1.Columns(4).Text <> "" Then
            Combo1.ListIndex = 1
         Else
            Combo1.ListIndex = 0
         End If
         
         LLamaLineas anc, 4
         PonerFoco txtAux(1)
   End If
   Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(3).Top = alto
    txtAux(0).Left = DataGrid1.Left + 340
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
    txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 65
    txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 65
    Me.Combo1.Top = alto
    Combo1.Left = txtAux(3).Left + txtAux(3).Width
End Sub

Private Sub LLamaLineas2(alto As Single)
    
    txtAux2(0).Top = alto
    txtAux2(1).Top = alto
    txtAux2(2).Top = alto
    cmdAux2.Top = alto
    cmdAux2.visible = ModificaLineas = 1
    txtAux2(0).Locked = ModificaLineas = 2
    txtAux2(0).Left = DataGrid2.Left + 340
    cmdAux2.Left = txtAux2(0).Left + txtAux2(0).Width + 15
    txtAux2(1).Left = txtAux2(0).Left + txtAux2(0).Width + 65
    txtAux2(2).Left = txtAux2(1).Left + txtAux2(1).Width + 65

End Sub



Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2

    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    If Modo = 5 Then
        If Adodc2.Recordset.EOF Then Exit Sub
        SQL = "¿Seguro que desea eliminar la linea de coste? " & vbCrLf
        SQL = SQL & vbCrLf & "Código: " & Format(Adodc2.Recordset.Fields(0), FormatoCod)
        SQL = SQL & vbCrLf & "Descripcion: " & Adodc2.Recordset.Fields(1)
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
        SQL = "DELETE FROM sunilin"
        SQL = SQL & " WHERE codunida =" & adodc1.Recordset!codunida & " AND codigo =" & Adodc2.Recordset!codigo
        conn.Execute SQL
        CargaGrid2 True
    
    Else
        'Eliminar normal
        SQL = DevuelveDesdeBD(conAri, "codunida", "sartic", "codunida", CStr(adodc1.Recordset!codunida))
        If SQL <> "" Then
            MsgBox "Existen articulos con este tipo de unidad", vbExclamation
            Exit Sub
        End If
        
        '### a mano
        SQL = "¿Seguro que desea eliminar el Tipo de Unidad? " & vbCrLf
        SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), FormatoCod)
        SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
        
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            NumRegElim = Me.adodc1.Recordset.AbsolutePosition
            'Hay que eliminar
            SQL = "Delete from sunida where codunida=" & adodc1.Recordset!codunida
            conn.Execute SQL
            CancelaADODC Me.adodc1
            CargaGrid ""
            CancelaADODC Me.adodc1
            SituarDataPosicion Me.adodc1, NumRegElim, SQL
        End If

    End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Tipo Unidad", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String
On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
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
        Case 1  'HacerBusqueda
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                PonerModo 2
                CargaGrid cadB
                DataGrid1.SetFocus
            End If
            
        Case 5
            If InsertarModificar Then
                If ModificaLineas = 2 Then
                    'MODIFICARç
                    NumRegElim = Adodc2.Recordset!codigo
                    CargaGrid2 True
                    Adodc2.Recordset.Find (" codigo =" & NumRegElim)
    
                    PonerBotonCabecera True
                    PonerFocoBtn Me.cmdAceptar
                    ModificaLineas = 0
                Else
                    CargaGrid2 True
                    BotonAnyadir
                End If
            End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdAux2_Click()
Dim cad As String
        
        
    cad = "Código|sunidesc|codigo|N||20·Descripcion|sunidesc|descripcion|T||70·"
    
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = "sunidesc"
        frmB.vSQL = ""
    
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Tipos de coste"
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri
        frmB.vCargaFrame = False
        '#
        frmB.Show vbModal
        Set frmB = Nothing
End Sub

Private Sub cmdCancelar_Click()
On Error Resume Next

    Select Case Modo
    Case 3 'Insertar
        DataGrid1.AllowAddNew = False
        'CargaGrid
        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    Case 4 'Modificar
        TerminaBloquear
        Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    Case 1 'Busqueda
        CargaGrid
    Case 5
        DataGrid2.AllowAddNew = False
        Campos_2_Visibles False
        ModificaLineas = 0
        DataGrid2.Enabled = True
        CargaGrid2 True
        PonerBotonCabecera True
        cmdRegresar.visible = True
        Exit Sub
    End Select
    PonerModo 2
    DataGrid1.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Modo = 5 Then
        Me.lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        If DataGrid2.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid2
            DataGrid2.Bookmark = 1
        End If
        
        Campos_2_Visibles False
        PonerModo 2

    Else

        If adodc1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
    
        cad = adodc1.Recordset.Fields(0) & "|"
        cad = cad & adodc1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub




Private Sub Combo1_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

On Error GoTo Error1

    If Not adodc1.Recordset.EOF Then lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
        
        
    


    If Modo = 2 Or Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not adodc1.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Devuelve = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedcl", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
            'Poner descripcion de ampliacion lineas
            CargaGrid2 True
        Else
            
        End If
    End If
    
Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
        
        

End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    ' ICONITOS DE LA BARRA
    Me.Icon = frmPpal.Icon
    If vParamAplic.Descriptores Then Me.Caption = "Formatos"
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Recuperar Todos
        .Buttons(5).Image = 3   'Botón Añadir Nuevo Registro
        .Buttons(6).Image = 4   'Botón Modificar Registro
        .Buttons(7).Image = 5   'Botón Borrar Registro
        .Buttons(9).Image = 10  '
        .Buttons(10).Image = 16  'Botón Imprimir
        .Buttons(11).Image = 15  'Botón Salir
    End With
    
    '------------------------------------------------
    DataGrid2.visible = vParamAplic.Produccion
    Label1.visible = vParamAplic.Produccion
    Me.Toolbar1.Buttons(9).visible = vParamAplic.Produccion
    If vParamAplic.Produccion Then
        DataGrid1.Height = 3525
    Else
        DataGrid1.Height = 6285
    End If
    
    CargarCombo_SiNo Combo1
    
    FormatoCod = FormatoCampo(txtAux(0))
    
    '## A mano
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    CadAncho = False
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 2
    
    'Cadena consulta
    CadenaConsulta = "Select codunida,nomunida,nomunbre,tasareciclado,if(EsTrabajo=1,""Si"","""")   from sunida"
    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    txtAux2(0).Text = RecuperaValor(CadenaDevuelta, 1)
    txtAux2(1).Text = RecuperaValor(CadenaDevuelta, 2)
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

Private Sub mnMtoLineas_Click()
    MtoLineas
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
        Case 9: MtoLineas
        Case 10 'Imprimir listado Tipos de Unidades
                Me.Hide
                AbrirListado (3) 'OpcionListado=3
                Me.Show vbModal
        Case 11: mnSalir_Click
    End Select
End Sub

Private Sub MtoLineas()
    ModificaLineas = 0
    PonerModo 5
    PonerBotonCabecera True
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
    SQL = SQL & " ORDER BY codunida"
    
    CargaGridGnral DataGrid1, Me.adodc1, SQL, False
    
    i = 0 'Cod. Tipo Unidad
        DataGrid1.Columns(i).Caption = "Unidad"
        DataGrid1.Columns(i).Width = 700
        DataGrid1.Columns(i).NumberFormat = FormatoCod
    
    i = 1 'Desc. Tipo Unidad
        DataGrid1.Columns(i).Caption = "Denominación"
        DataGrid1.Columns(i).Width = 2100
        
    i = 2 'Abrev.
        DataGrid1.Columns(i).Caption = "Abrev."
        DataGrid1.Columns(i).Width = 700
            
            
    i = 3 'Tasa reciclado
        DataGrid1.Columns(i).Caption = "Tasa"
        DataGrid1.Columns(i).Width = 1200
        DataGrid1.Columns(i).Alignment = dbgRight
        DataGrid1.Columns(i).NumberFormat = "0.00000"
            
    i = 4 'Es trabajo
        DataGrid1.Columns(i).Caption = "Dosis"
        DataGrid1.Columns(i).Width = 600
        DataGrid1.Columns(i).Alignment = dbgCenter
        
            
            
            
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        txtAux(2).Width = DataGrid1.Columns(2).Width - 60
        txtAux(3).Width = DataGrid1.Columns(3).Width - 30
        Me.Combo1.Width = DataGrid1.Columns(4).Width
        CadAncho = True
    End If
   
   'No permitir cambiar tamaño de columnas
   For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
   Next i
   
    'Habilitamos botones Modificar y Eliminar
   If Toolbar1.Buttons(6).Enabled Then
        Toolbar1.Buttons(6).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        mnModificar.Enabled = Not adodc1.Recordset.EOF
        mnEliminar.Enabled = Not adodc1.Recordset.EOF
   End If
   DataGrid1.Enabled = b
   DataGrid1.ScrollBars = dbgAutomatic
   
   CargaGrid2 Not adodc1.Recordset.EOF
   
   
   PonerOpcionesMenu
   
   'Actualizar indicador
   If Not adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
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
    If Index = 0 Then PonerFormatoEntero txtAux(Index) 'Cod. Tipo Unidad
    If Index = 3 Then
        ' lo que ponga en su TAG  (8)
        If Not PonerFormatoDecimal_Single(txtAux(Index), 8) Then txtAux(Index).Text = ""  'Cod. Tipo Unidad
    End If
End Sub




Private Function DatosOk() As Boolean
Dim b As Boolean

    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    'Comprobar si ya existe el cod de tipo unidad en la tabla
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


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "&Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Lineas Detalle"
        PonerFocoBtn Me.cmdRegresar
    Else
        Campos_2_Visibles False
        Me.lblIndicador.Caption = ""
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Campos_2_Visibles(visibles As Boolean)
        txtAux2(0).visible = visibles: txtAux2(1).visible = visibles: txtAux2(2).visible = visibles
        Me.cmdAux2.visible = visibles
End Sub

Private Sub CargaGrid2(Enlaza As Boolean)
Dim i As Byte
Dim b As Boolean
Dim SQL As String
Dim PriVe As Boolean


    If Not vParamAplic.Produccion Then Exit Sub

    b = DataGrid2.Enabled
    DataGrid2.Enabled = False
    SQL = "select sunilin.codigo,descripcion,importe from sunilin,sunidesc where sunilin.codigo=sunidesc.codigo AND codunida = "
    If Enlaza Then
        SQL = SQL & adodc1.Recordset!codunida
    Else
        SQL = SQL & " -1"
    End If
    SQL = SQL & " ORDER BY codigo"
    
    PriVe = Adodc2.Recordset Is Nothing
    
    CargaGridGnral DataGrid2, Me.Adodc2, SQL, PriVe
    
    i = 0 'Cod. Tipo Unidad
        DataGrid2.Columns(i).Caption = "Coste"
        DataGrid2.Columns(i).Width = 700
        DataGrid2.Columns(i).NumberFormat = FormatoCod
    
    i = 1 'Desc. Tipo Unidad
        DataGrid2.Columns(i).Caption = "Denominación"
        DataGrid2.Columns(i).Width = 2500
        
    i = 2 'Tasa reciclado
        DataGrid2.Columns(i).Caption = "Importe"
        DataGrid2.Columns(i).Width = 1200
        DataGrid2.Columns(i).Alignment = dbgRight
        DataGrid2.Columns(i).NumberFormat = "0.00000"
    'Fiajamos el cadancho
    If PriVe Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux2(0).Width = DataGrid2.Columns(0).Width - 60
        txtAux2(1).Width = DataGrid2.Columns(1).Width - 60
        txtAux2(2).Width = DataGrid2.Columns(2).Width - 60
        'txtAux(3).Width = DataGrid1.Columns(3).Width - 30
        
    End If
   
   'No permitir cambiar tamaño de columnas
   For i = 0 To DataGrid2.Columns.Count - 1
        DataGrid2.Columns(i).AllowSizing = False
   Next i
   
   
   DataGrid2.Enabled = b
   DataGrid2.ScrollBars = dbgAutomatic
   
   
End Sub



Private Sub txtAux2_GotFocus(Index As Integer)
    ConseguirFoco txtAux2(Index), Modo
End Sub

Private Sub txtAux2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux2_LostFocus(Index As Integer)
Dim cad As String

    If Index = 0 Then
        cad = ""
        If PonerFormatoEntero(txtAux2(Index)) Then
            cad = DevuelveDesdeBD(conAri, "descripcion", "sunidesc", "codigo", txtAux2(Index))
            If cad = "" Then
                MsgBox "No existe el coste: " & txtAux2(Index).Text, vbExclamation
                txtAux2(0).Text = ""
                PonerFoco txtAux2(Index)
            End If
        Else
            txtAux2(0).Text = ""
        End If
        txtAux2(1).Text = cad
        If txtAux2(0).Text <> "" Then PonerFoco txtAux2(2)
    End If
    If Index = 2 Then
        ' lo que ponga en su TAG  (8)
        If Not PonerFormatoDecimal_Single(txtAux2(Index), 8) Then txtAux2(Index).Text = ""  'Cod. Tipo Unidad
    End If
End Sub




Private Function InsertarModificar() As Boolean
Dim C As String
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    For NumRegElim = 0 To 2
        txtAux2(NumRegElim).Text = Trim(txtAux2(NumRegElim).Text)
        If txtAux2(NumRegElim).Text = "" Then
            MsgBox "todos los campos son obligatorios", vbExclamation
            Exit Function
        End If
    Next
    
    
    C = TransformaComasPuntos(txtAux2(2).Text)
    If ModificaLineas = 1 Then
        
        '               codigo              importe
        C = "," & Val(txtAux2(0).Text) & "," & C & ")"
        C = "INSERT INTO sunilin (codunida ,codigo ,importe) VALUES (" & adodc1.Recordset!codunida & C
        
    
    Else
        C = "UPDATE sunilin set importe = " & C
        C = C & " WHERE codunida =" & adodc1.Recordset!codunida & " AND codigo =" & Adodc2.Recordset!codigo
    End If
    
    conn.Execute C
    InsertarModificar = True
    Exit Function
EInsertarModificar:
        MuestraError Err.Number, Err.Description

End Function




