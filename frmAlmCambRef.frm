VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmCambRef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio código articulo-referencia"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   17985
   Icon            =   "frmAlmCambRef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   17985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   90
      TabIndex        =   21
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   22
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
   Begin VB.Frame FrameBotonGnral2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3765
      TabIndex        =   19
      Top             =   90
      Width           =   840
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   20
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Actualizar"
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
      Index           =   7
      Left            =   13440
      MaxLength       =   16
      TabIndex        =   6
      Tag             =   "Compra|T|S|||sarticcambioref|referprov|||"
      Text            =   " "
      Top             =   4920
      Width           =   675
   End
   Begin VB.CommandButton cmdArticulo 
      Caption         =   "+"
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
      Index           =   4
      Left            =   9720
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdArticulo 
      Caption         =   "+"
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
      Index           =   3
      Left            =   7440
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
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
      Index           =   6
      Left            =   12720
      MaxLength       =   16
      TabIndex        =   5
      Tag             =   "Compra|N|S|||sarticcambioref|precom|0,0000||"
      Text            =   " "
      Top             =   4920
      Width           =   675
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
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
      Index           =   5
      Left            =   12120
      MaxLength       =   16
      TabIndex        =   4
      Tag             =   "Venta|N|S|||sarticcambioref|prevta|0,0000||"
      Text            =   "Dato2"
      Top             =   4920
      Width           =   675
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
      Index           =   4
      Left            =   9000
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Familia|N|S|||sarticcambioref|codfamia|||"
      Text            =   "Dato2"
      Top             =   4920
      Width           =   675
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Index           =   4
      Left            =   9960
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   18
      Text            =   "Dato2"
      Top             =   4920
      Width           =   2475
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
      Index           =   3
      Left            =   6600
      MaxLength       =   16
      TabIndex        =   2
      Tag             =   "Prov|N|S|||sarticcambioref|codprove|||"
      Text            =   "Dato2"
      Top             =   4920
      Width           =   675
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Index           =   3
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   16
      Text            =   "Dato2"
      Top             =   4920
      Width           =   795
   End
   Begin VB.CommandButton cmdArticulo 
      Caption         =   "+"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   14
      Text            =   "Dato2"
      Top             =   4920
      Width           =   2475
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
      Index           =   2
      Left            =   5160
      MaxLength       =   16
      TabIndex        =   1
      Tag             =   "Dest|T|N|||sarticcambioref|codarti1||N|"
      Text            =   "Dato2"
      Top             =   4920
      Width           =   1395
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
      Left            =   1260
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Origen|T|N|||sarticcambioref|codartic||S|"
      Text            =   "Dato2"
      Top             =   4920
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmCambRef.frx":000C
      Height          =   8500
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   855
      Width           =   17715
      _ExtentX        =   31247
      _ExtentY        =   15002
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
      Left            =   16800
      TabIndex        =   8
      Top             =   9585
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
      Left            =   15645
      TabIndex        =   7
      Top             =   9585
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   11
      Top             =   9480
      Width           =   2715
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
         Left            =   120
         TabIndex        =   12
         Top             =   195
         Width           =   2520
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
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   10
      Tag             =   "Linea|N|N|0||sarticcambioref|numlinea|000|N|"
      Text            =   "Dat"
      Top             =   4920
      Width           =   800
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
Attribute VB_Name = "frmAlmCambRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos

Private WithEvents frmAr As frmBasico2
Attribute frmAr.VB_VarHelpID = -1
Private WithEvents frmP As frmBasico2 '%=%=frmComProveedores
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmF As frmBasico2 'frmAlmFamiliaArticulo
Attribute frmF.VB_VarHelpID = -1

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
Dim PulsadoMas2 As Boolean
Dim J As Integer

Private Sub PonerModo(vModo As Byte)
Dim b As Boolean


    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
         
    For J = 0 To txtAux.Count - 1
        txtAux(J).visible = Not b
        If J = 1 Or J = 3 Or J = 4 Then
            cmdArticulo(J).visible = Not b
            txtAux2(J).visible = Not b
        End If
    Next
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b



    'Si estamos insertando o busqueda
    b = Modo <> 3 And Modo <> 1
    BloquearTxt txtAux(0), b
    BloquearTxt txtAux(1), b
    
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCamposGnral Me, Modo, 3
    
    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
                        
                        
    PulsadoMas2 = False
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



Private Sub BotonAnyadir()
Dim anc As Single
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Adodc1
   
    anc = ObtenerAlto(DataGrid1, 10)
    
    limpiar Me
    
    'Obtenemos la siguiente numero de Almacen
    txtAux(0).Text = SugerirCodigoSiguienteStr("sarticcambioref", "numlinea")
    FormateaCampo txtAux(0)
    
    
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonerFoco txtAux(1)
End Sub


Private Sub BotonBuscar()
Dim T As TextBox

    CargaGrid "false"  'para vaciar los datos del Grid
    
    'Buscar
    For Each T In txtAux
        T.Text = ""
    Next
    For Each T In txtAux2
        T.Text = ""
    Next
    
    LLamaLineas DataGrid1.Top + 240, 1
    PonerFoco txtAux(1)
End Sub





Private Sub BotonVerTodos()
    On Error Resume Next

    CargaGrid ""
    If Adodc1.Recordset.RecordCount <= 0 Then
         MsgBox "No hay ningún registro en la tabla sarticcambioref", vbInformation
         Screen.MousePointer = vbDefault
         Exit Sub
    Else
'        adodc1.Recordset.MoveFirst
        PonerFocoGrid Me.DataGrid1
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificar()
Dim cad As String
Dim anc As Single

    
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        J = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
    anc = ObtenerAlto(DataGrid1, 10)
    
    'Cad = ""
    'For J = 0 To 1
    '    Cad = Cad & DataGrid1.Columns(J).Text & "|"
    'Next J
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux2(1).Text = DataGrid1.Columns(2).Text
    txtAux(2).Text = DataGrid1.Columns(3).Text
    
    
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux2(3).Text = DataGrid1.Columns(5).Text
    txtAux(4).Text = DataGrid1.Columns(6).Text
    txtAux2(4).Text = DataGrid1.Columns(7).Text
    txtAux(5).Text = DataGrid1.Columns(8).Text
    txtAux(6).Text = DataGrid1.Columns(9).Text
    txtAux(7).Text = DataGrid1.Columns(10).Text
    
    
    
    LLamaLineas anc, 4
    PonerFoco txtAux(2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    For J = 0 To txtAux.Count - 1
        txtAux(J).Top = alto
        If J = 1 Or J = 3 Or J = 4 Then
            cmdArticulo(J).Top = alto
            txtAux2(J).Top = alto
        End If
    Next
    
    
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    
    '### a mano
    SQL = "¿Seguro que desea eliminar el artículo?" & vbCrLf
    SQL = SQL & vbCrLf & "Linea: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Artículo: " & Adodc1.Recordset.Fields(1) & " " & Adodc1.Recordset.Fields(2)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Adodc1.Recordset.AbsolutePosition
        SQL = "Delete from sarticcambioref where codartic=" & DBSet(Adodc1.Recordset!codArtic, "T")
        conn.Execute SQL
        CancelaADODC Me.Adodc1
        CargaGrid ""
        CancelaADODC Me.Adodc1
        SituarDataPosicion Me.Adodc1, NumRegElim, SQL
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar ", Err.Description
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String

    On Error Resume Next

    Select Case Modo
        Case 3  'Insertar
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
                        i = Adodc1.Recordset.Fields(0)
                        PonerModo 2
                        CancelaADODC Me.Adodc1
                        CargaGrid
                        Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & i)
                    End If
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
            
        Case 1  'HacerBusqueda
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                PonerModo 2
                CargaGrid cadB
                PonerFocoGrid Me.DataGrid1
            End If
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdArticulo_Click(Index As Integer)
    If Modo = 2 Then Exit Sub
    
    Select Case Index
    Case 1
        If Modo = 4 Then Exit Sub
        CadenaConsulta = ""
        Set frmAr = New frmBasico2
'        frmAr.DatosADevolverBusqueda = "0|1|"
'        frmAr.Show vbModal
        AyudaArticulos frmAr, txtAux(Index), , , , True
        Set frmAr = Nothing
        
        
    Case 3
        
'        Set frmP = New frmComProveedores
'        frmP.DatosADevolverBusqueda = "0|1|"
'        frmP.Show vbModal
        Set frmP = New frmBasico2
        AyudaProveedores frmP, txtAux(Index)
        Set frmP = Nothing
        
    Case 4
        
'        Set frmF = New frmAlmFamiliaArticulo
'        frmF.DatosADevolverBusqueda = "0|1|"
'        frmF.Show vbModal
        Set frmF = New frmBasico2
        AyudaFamilias frmF, txtAux(Index)
        Set frmF = Nothing
    
    
    End Select
    
    If CadenaConsulta <> "" Then
        Me.txtAux(Index).Text = RecuperaValor(CadenaConsulta, 1)
        Me.txtAux2(Index).Text = RecuperaValor(CadenaConsulta, 2)
        CadenaConsulta = ""
        
    End If

End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar

    Select Case Modo
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
        Case 4 'Modificar
            Me.lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
        Case 1 'Buscar
            CargaGrid
    End Select
    PonerModo 2
    PonerFocoGrid Me.DataGrid1
    
ECancelar:
    If Err.Number <> 0 Then Err.Clear
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

'    ' ICONITOS DE LA BARRA
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Busqueda
'        .Buttons(2).Image = 2   'Botón Recuperar Todos
'        .Buttons(5).Image = 3   'Botón Añadir Nuevo Registro
'        .Buttons(6).Image = 4   'Botón Modificar Registro
'        .Buttons(7).Image = 5   'Botón Borrar Registro
'        .Buttons(10).Image = 16  'Botón Imprimir
'        .Buttons(12).Image = 42  'Actualizar referencias
'
'        .Buttons(14).Image = 15  'Botón Salir
'    End With
    
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
    
    '++ añdimos los puntos de utilidades
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 42 ' actualizar referencias
    End With
    
    
    CadAncho = False
    
    PonerModo 2
    
    'Cadena consulta
    
    CargaGrid
End Sub


Private Sub frmAr_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub frmF_DatoSeleccionado(CadenaSeleccion As String)
     CadenaConsulta = CadenaSeleccion
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
 CadenaConsulta = CadenaSeleccion
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
        Case 8 ' imprimir
            If Modo <> 2 Then Exit Sub
            If Adodc1.Recordset.EOF Then Exit Sub
            
            Imprime
    End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
Dim i As Byte
Dim b As Boolean
    
    


    CadenaConsulta = "SELECT numlinea ,sarticcambioref.codartic, sartic.nomartic, sarticcambioref.codarti1,"
    CadenaConsulta = CadenaConsulta & " sarticcambioref.codprove,nomprove,sarticcambioref.codfamia,nomfamia"
    CadenaConsulta = CadenaConsulta & " ,prevta,precom, sarticcambioref.referprov"
    CadenaConsulta = CadenaConsulta & " ,trim(concat("
    CadenaConsulta = CadenaConsulta & " if(sartic_1.nomartic is null,'','A'),' ',"
    CadenaConsulta = CadenaConsulta & " if(sarticcambioref.codfamia>=0 and nomfamia is null,'F',''),' ',"
    CadenaConsulta = CadenaConsulta & " if(sarticcambioref.codprove>=0 and nomprove is null,'P',''))) verror"
    
    CadenaConsulta = CadenaConsulta & " FROM  sarticcambioref sarticcambioref  LEFT JOIN sartic sartic_1 ON sarticcambioref.codarti1=sartic_1.codartic"
    CadenaConsulta = CadenaConsulta & " LEFT JOIN sartic sartic ON sarticcambioref.codartic=sartic.codartic"
    CadenaConsulta = CadenaConsulta & " LEFT JOIN sprove ON sarticcambioref.codprove=sprove.codprove"
    CadenaConsulta = CadenaConsulta & " LEFT JOIN sfamia ON sarticcambioref.codfamia=sfamia.codfamia"


    
    b = DataGrid1.Enabled
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY numlinea"
    
    CargaGridGnral DataGrid1, Me.Adodc1, SQL, False

    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "Linea"
        DataGrid1.Columns(i).Width = 0 '600
   
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Origen"
        DataGrid1.Columns(i).Width = 2000 '1400
    i = 2
        DataGrid1.Columns(i).Caption = "Descripcion"
        DataGrid1.Columns(i).Width = 3500 '3000
    i = 3
        DataGrid1.Columns(i).Caption = "Destino"
        DataGrid1.Columns(i).Width = 2000 '1400
    i = 4
        DataGrid1.Columns(i).Caption = "Prov"
        DataGrid1.Columns(i).Width = 800
    
    i = 5
        DataGrid1.Columns(i).Caption = "Proveedor"
        DataGrid1.Columns(i).Width = 1800
    
    i = 6
        DataGrid1.Columns(i).Caption = "Fam"
        DataGrid1.Columns(i).Width = 700
    i = 7
        DataGrid1.Columns(i).Caption = "Familia"
        DataGrid1.Columns(i).Width = 1200
        
    i = 8
        DataGrid1.Columns(i).Caption = "€ Vta"
        DataGrid1.Columns(i).Width = 1150 '950
        DataGrid1.Columns(i).NumberFormat = FormatoPrecio
        DataGrid1.Columns(i).Alignment = dbgRight
    i = 9
        DataGrid1.Columns(i).Caption = "€ Compra"
        DataGrid1.Columns(i).Width = 1150 '950
        DataGrid1.Columns(i).NumberFormat = FormatoPrecio
        DataGrid1.Columns(i).Alignment = dbgRight
        
    i = 10
        DataGrid1.Columns(i).Caption = "Ref. prov"
        DataGrid1.Columns(i).Width = 2100 '1600
    
    i = 11
        DataGrid1.Columns(i).Caption = "Error"
        DataGrid1.Columns(i).Width = 750 '1300
        
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
'--
'        txtAux(0).Left = DataGrid1.Columns(0).Left + DataGrid1.Left
'        txtAux(0).Width = DataGrid1.Columns(0).Width - 30
        txtAux(1).Left = DataGrid1.Columns(1).Left + DataGrid1.Left
        txtAux(1).Width = DataGrid1.Columns(1).Width - 30
        txtAux2(1).Left = DataGrid1.Columns(2).Left + DataGrid1.Left
        txtAux2(1).Width = DataGrid1.Columns(2).Width - 60
        Me.cmdArticulo(1).Left = txtAux2(1).Left - 120
        txtAux(2).Left = DataGrid1.Columns(3).Left + DataGrid1.Left
        txtAux(2).Width = DataGrid1.Columns(3).Width - 60
        
        'Codprove y famia
        For i = 3 To 4
            J = 4
            If i = 4 Then J = 6
            txtAux(i).Left = DataGrid1.Columns(J).Left + DataGrid1.Left
            txtAux(i).Width = DataGrid1.Columns(J).Width - 30
            txtAux2(i).Left = DataGrid1.Columns(J + 1).Left + DataGrid1.Left
            txtAux2(i).Width = DataGrid1.Columns(J + 1).Width - 60
            Me.cmdArticulo(i).Left = txtAux2(i).Left - 120
        Next
        
        For i = 5 To 7
            txtAux(i).Left = DataGrid1.Columns(i + 3).Left + DataGrid1.Left
            txtAux(i).Width = DataGrid1.Columns(i + 3).Width - 30
        Next
        
        
        CadAncho = True
    End If
   
   'No permitir cambiar tamaño de columnas
   For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
   Next i
   
   'Habilitamos botones Modificar y Eliminar
   If Toolbar1.Buttons(2).Enabled Then
        Toolbar1.Buttons(2).Enabled = Not Adodc1.Recordset.EOF
        Toolbar1.Buttons(3).Enabled = Not Adodc1.Recordset.EOF
        mnModificar.Enabled = Not Adodc1.Recordset.EOF
        mnEliminar.Enabled = Not Adodc1.Recordset.EOF
   End If
   DataGrid1.Enabled = b
   DataGrid1.ScrollBars = dbgAutomatic
   
   DataGrid1.RowHeight = 350
   
   
   'Actualizar indicador
   If Not Adodc1.Recordset.EOF And (Modo = 2) Then
        lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
   PonerOpcionesMenu
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' actualizar
            Screen.MousePointer = vbHourglass
            ActualizarReferencias
            Screen.MousePointer = vbDefault
            
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Index = 1 Or Index = 3 Or Index = 4) And KeyCode = vbKeyAdd Then
        If Modo <> 2 Then
            PulsadoMas2 = True
            cmdArticulo_Click Index
            PulsadoMas2 = False
        End If
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If PulsadoMas2 Then Exit Sub

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    Select Case Index
    Case 0
        PonerFormatoEntero txtAux(Index) 'codalmpr
    Case 1
        CadenaConsulta = ""
        If Me.txtAux(Index).Text <> "" Then
            CadenaConsulta = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtAux(Index).Text, "T")
            If CadenaConsulta = "" Then
                MsgBox "No existe el articulo", vbExclamation
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
        End If
        Me.txtAux2(1).Text = CadenaConsulta
    Case 2
        CadenaConsulta = "OK"
        If Me.txtAux(Index).Text <> "" Then
            CadenaConsulta = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtAux(Index).Text, "T")
            If CadenaConsulta <> "" Then
                CadenaConsulta = "ERROR"
                MsgBox "YA existe el articulo", vbExclamation
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
        End If
    Case 3
        'Prooveedor
        CadenaConsulta = ""
        If Me.txtAux(Index).Text <> "" Then
            If PonerFormatoEntero(txtAux(Index)) Then
                CadenaConsulta = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(Index).Text)
                If CadenaConsulta = "" Then
                    MsgBox "No existe el proveedor", vbExclamation
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            End If
        End If
        Me.txtAux2(Index).Text = CadenaConsulta
    Case 4
        'Familia
        CadenaConsulta = ""
        If Me.txtAux(Index).Text <> "" Then
            If PonerFormatoEntero(txtAux(Index)) Then
                CadenaConsulta = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtAux(Index).Text)
                If CadenaConsulta = "" Then
                    MsgBox "No existe la familia", vbExclamation
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            End If
        End If
        Me.txtAux2(Index).Text = CadenaConsulta
        
    Case 5, 6 'precios
        PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
    
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    

    'El articulo NO puede ser de varios
    b = False
    CadenaConsulta = DevuelveDesdeBD(conAri, "artvario", "sartic", "codartic", Me.txtAux(1).Text, "T")
    If CadenaConsulta = "" Then
        MsgBox "No existe el articulo", vbExclamation
    ElseIf CadenaConsulta = "1" Then
        MsgBox "El articulo NO puede ser de varios", vbExclamation
    Else
        b = True
    End If
    
    
    
    
    
    
    If b Then
        'El articulo destino NO puede estar ya en la tabla
        If Modo = 3 Then
            CadenaConsulta = ""
        Else
            CadenaConsulta = "codartic <> " & DBSet(txtAux(1).Text, "T") & " AND "
        End If
        CadenaConsulta = CadenaConsulta & "codarti1"
        CadenaConsulta = DevuelveDesdeBD(conAri, "numlinea", "sarticcambioref", CadenaConsulta, Me.txtAux(2).Text, "T")
        If CadenaConsulta <> "" Then
            MsgBox "Ya existe la nueva referencia. Linea: " & CadenaConsulta, vbExclamation
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


Private Sub Imprime()
    With frmImprimir
            .ConSubInforme = False
            .FormulaSeleccion = ""
            .NombreRPT = "rAlmCambioReferen.rpt"
            .NombrePDF = .NombreRPT
            .OtrosParametros = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            .NumeroParametros = 1
            .Opcion = 2003 'Esta libre
            .Titulo = "Cambio referencia"
            .Show vbModal
        End With
End Sub


Private Sub ActualizarReferencias()
Dim H As Integer
Dim tabla As String
Dim Cole As Collection
Dim k As Integer
Dim J As Integer
Dim Aux As String
Dim CambiosOk As Boolean
    
    'Primera combrobacion. No existe ninguna referencia
    Me.lblIndicador.Caption = "Comprobaciones"
    Me.lblIndicador.Refresh
    
    Set miRsAux = New ADODB.Recordset
    CadenaConsulta = "select * from sarticcambioref where codarti1 in (select codartic from sartic)"
    miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaConsulta = ""
    While Not miRsAux.EOF
        CadenaConsulta = CadenaConsulta & vbCrLf & "  -" & miRsAux!codarti1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If CadenaConsulta <> "" Then CadenaConsulta = "A) Ya existen los artículos" & vbCrLf & CadenaConsulta
    
    tabla = "select * from sarticcambioref where not codartic in (select codartic from sartic)"
    miRsAux.Open tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    tabla = ""
    While Not miRsAux.EOF
        tabla = tabla & vbCrLf & "  -" & miRsAux!codArtic
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If tabla <> "" Then CadenaConsulta = CadenaConsulta & vbCrLf & "B) No  existen los artículos" & vbCrLf & tabla

    
    'PROVEEDORES Y FAMILIAS
    tabla = "select * from sarticcambioref where not codprove in (select codprove from sprove)"
    miRsAux.Open tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    tabla = ""
    While Not miRsAux.EOF
        tabla = tabla & vbCrLf & "  -" & miRsAux!codArtic
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If tabla <> "" Then CadenaConsulta = CadenaConsulta & vbCrLf & "C) No  existe el proceedor para los artículos" & vbCrLf & tabla

    tabla = "select * from sarticcambioref where not codfamia in (select codfamia from sfamia)"
    miRsAux.Open tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    tabla = ""
    While Not miRsAux.EOF
        tabla = tabla & vbCrLf & "  -" & miRsAux!codArtic
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If tabla <> "" Then CadenaConsulta = CadenaConsulta & vbCrLf & "D) No  existe la familia para los artículos" & vbCrLf & tabla


    'Vemos si hay articulos con cambio de precios
    Aux = "Select codartic from sarticcambioref where prevta >=0"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not miRsAux.EOF
        Aux = Aux & ", " & DBSet(miRsAux!codArtic, "T")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Aux <> "" Then
        'Hay articulos con cambio de precio de venta. Veremos si existe en slista y si existe, si
        'no tiene fechanue. Ya que el cambio de articulos provocara que pongamos en fechanue=now and prenue=prevta
        Aux = Mid(Aux, 2)
        
        Aux = "select * from slista WHERE fechanue>='1900-01-01' AND codartic IN (" & Aux & ") "
        miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Aux = ""
        While Not miRsAux.EOF
            Aux = Aux & "   -" & miRsAux!codArtic & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If Aux <> "" Then CadenaConsulta = CadenaConsulta & vbCrLf & "E) Lista de precios venta sin actualizar" & vbCrLf & Aux
    
    End If

    'Vemos si hay articulos con cambio de preciocompra
    Aux = "Select codartic from sarticcambioref where precom >=0"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not miRsAux.EOF
        Aux = Aux & ", " & DBSet(miRsAux!codArtic, "T")
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Aux <> "" Then
        'Hay articulos con cambio de precio de venta. Veremos si existe en slista y si existe, si
        'no tiene fechanue. Ya que el cambio de articulos provocara que pongamos en fechanue=now and prenue=prevta
        Aux = Mid(Aux, 2)
        
        Aux = "select * from slispr WHERE fechanue>='1900-01-01' AND codartic IN (" & Aux & ") "
        miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Aux = ""
        While Not miRsAux.EOF
            Aux = Aux & "   -" & miRsAux!codArtic & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If Aux <> "" Then CadenaConsulta = CadenaConsulta & vbCrLf & "F) Lista de precios compra sin actualizar" & vbCrLf & Aux
    
    End If

    If CadenaConsulta <> "" Then
        CadenaConsulta = "ERRORES.  " & vbCrLf & CadenaConsulta
        MsgBox CadenaConsulta, vbCritical
        Exit Sub
    End If
    
    
    Set Cole = New Collection
    CadenaConsulta = "select * from sarticcambioref ORDER BY numlinea,codartic"
    miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cole.Add CStr(miRsAux!codArtic)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    CadenaConsulta = ""
    
    'Vemos que no exista en otros parametros
    'advparametros artagrupa1   artagrupa2  artagrupa3   codarticGRUPO
    CadenaConsulta = CadenaConsulta & TablasParametros("artagrupa1,artagrupa2,artagrupa3,codarticGRUPO", "advparametros", Cole)
    CadenaConsulta = CadenaConsulta & TablasParametros("codartid,codartictel,ArtReciclado,ArticuloPortes,artRecargoFina,artSeparador,artTfoniaIvaExento", "spara1", Cole)
    If vParamAplic.TieneTelefonia2 > 0 Then CadenaConsulta = CadenaConsulta & TablasParametros("artiTelefNorORAN,artiTelefNorVOD", "spara2", Cole)
      
    Set Cole = Nothing
    If CadenaConsulta <> "" Then
        MsgBox CadenaConsulta, vbExclamation
        Exit Sub
    End If
    
    
    Me.lblIndicador.Caption = ""
    CadenaConsulta = "Si hay equipos trabajando el proceso podria llevar mucho tiempo." & vbCrLf & "¿Continuar?"
    If MsgBox(CadenaConsulta, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    CadenaConsulta = InputBox("Escriba password de seguridad para continuar con el proceso")
    
    If UCase(CadenaConsulta) <> "ARIADNA" Then Exit Sub
    
    
    
    
    Screen.MousePointer = vbHourglass
    
    'Preparando datos
    Set Cole = New Collection
    
    Cole.Add "advpartes_lineas#codartic|"
    Cole.Add "advtrata_lineas#codartic|"
    Cole.Add "salmac#codartic|"
    
    'Nov 2020
    'Cole.Add "salmagrupo#codartic|"
    
    Cole.Add "sarpmp#codartic|"
    Cole.Add "sarti1#codartic|codarti1|"
    Cole.Add "sarti2#codartic|"
    Cole.Add "sarti3#codartic|"
    Cole.Add "sarti5#codartic|"
    Cole.Add "sarti6#codartic|codarti1|"
    Cole.Add "sarti7#codartic|"
    Cole.Add "sbonif#codartic|codarti1|"
    Cole.Add "scarep#codartic|"
    Cole.Add "schrep#codartic|"
    Cole.Add "sconsulta#codartic|"
    Cole.Add "shinve#Codartic|"
    Cole.Add "sinven#codartic|"
    Cole.Add "slhalb#codartic|"
    Cole.Add "slhalp#codartic|"
    Cole.Add "slhmov#codartic|"
    Cole.Add "slhped#codartic|"
    Cole.Add "slhppr#codartic|"
    Cole.Add "slhpre#codartic|"
    Cole.Add "slhtra#codartic|"
    Cole.Add "slialb#codartic|"
    Cole.Add "slialp#codartic|"
    Cole.Add "slienvpr#codartic|"
    Cole.Add "slienvpr2#codartic|codarti2|"
    Cole.Add "slifac#codartic|"
    Cole.Add "slifpc#codartic|"
    Cole.Add "slimov#codartic|"
    Cole.Add "sliordpr#codartic|"
    Cole.Add "sliordpr2#codartic|codarti2|"
    Cole.Add "sliped#codartic|"
    Cole.Add "slipedb#codartic|"
    Cole.Add "slipedrma#codartic|"
    Cole.Add "slipla#codartic|"
    Cole.Add "slippr#codartic|"
    Cole.Add "slipre#codartic|"
    Cole.Add "slirep#codartic|"
    Cole.Add "slisp1#codartic|"
    Cole.Add "slispr#codartic|"
    Cole.Add "slist1#codartic|"
    Cole.Add "slista#codartic|"
    Cole.Add "slitra#codartic|"
    Cole.Add "sliven#codartic|"
    Cole.Add "slotes#codartic|"
    Cole.Add "smoval#codartic|"
    Cole.Add "spedidos#codartic|"
    Cole.Add "spree1#codartic|"
    Cole.Add "sprees#codartic|"
    Cole.Add "spromo#codartic|"
    Cole.Add "sserie#codartic|"
    Cole.Add "sserlin#codartic|"
    Cole.Add "stelem#codartic|"
    Cole.Add "stipco#codartic|"
    Cole.Add "straspaso#codartic|"
    
    
    
    'Para cada articulo
    CadenaConsulta = "select * from sarticcambioref ORDER BY numlinea,codartic"
    miRsAux.Open CadenaConsulta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Set LOG = New cLOG
    
    While Not miRsAux.EOF
    
        
        conn.Execute "SET FOREIGN_KEY_CHECKS=0"
        
        'S lleva cambio de precios comprobamos
        CambiosOk = True
        If Not IsNull(miRsAux!prevta) Then
            If Not ActualizarPrecios(True) Then CambiosOk = False
        End If
        If CambiosOk Then
            If Not IsNull(miRsAux!precom) Then
                If Not ActualizarPrecios(False) Then CambiosOk = False
            End If
        End If
    
        If CambiosOk Then
            Espera 1  'Noviembre 2020
            CadenaConsulta = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", CStr(miRsAux!codArtic), "T")
            CadenaConsulta = "[ACTUALIZAR] " & miRsAux!codArtic & "   Desc: " & CadenaConsulta & " -> " & miRsAux!codarti1
            LOG.Insertar 7, vUsu, CadenaConsulta
 
        
        
            CadenaConsulta = "UPDATE sartic SET codartic = " & DBSet(miRsAux!codarti1, "T")
            If Not IsNull(miRsAux!Codprove) Then CadenaConsulta = CadenaConsulta & ", codprove =" & miRsAux!Codprove
            If Not IsNull(miRsAux!Codfamia) Then CadenaConsulta = CadenaConsulta & ", codfamia =" & miRsAux!Codfamia
            If Not IsNull(miRsAux!referprov) Then CadenaConsulta = CadenaConsulta & ", referprov =" & DBSet(miRsAux!referprov, "T")
            
            CadenaConsulta = CadenaConsulta & " WHERE codartic = " & DBSet(miRsAux!codArtic, "T")
            conn.Execute CadenaConsulta
            
            'Cambamos en las tablas
            For k = 1 To Cole.Count
            
                J = InStr(1, Cole.Item(k), "#")
                
                Me.lblIndicador.Caption = miRsAux!codArtic & " " & k & "/" & Cole.Count
                Me.lblIndicador.Refresh
                tabla = Mid(Cole.Item(k), 1, J - 1)
                
                DoEvents
                Screen.MousePointer = vbHourglass
                
                Aux = Mid(Cole.Item(k), J + 1)
                
                While Aux <> ""
                    J = InStr(1, Aux, "|")
                    If J = 0 Then
                        Aux = ""
                    Else
                        CadenaConsulta = Mid(Aux, 1, J - 1)
                        Aux = Mid(Aux, J + 1)
                        
                        CadenaConsulta = "UPDATE " & tabla & " SET " & CadenaConsulta & " = " & DBSet(miRsAux!codarti1, "T") & " WHERE " & CadenaConsulta & " = " & DBSet(miRsAux!codArtic, "T")
                        'Debug.Print CadenaConsulta
                        conn.Execute CadenaConsulta
                    End If
                        
                Wend
            Next k
        
            'Reestablecemos
            DoEvents
            Me.lblIndicador.Caption = "Ajuste BD "
            Me.lblIndicador.Refresh
            
            
           
            
            conn.Execute "DELETE FROM  sarticcambioref WHERE codartic = " & DBSet(miRsAux!codArtic, "T")
        End If
        
        
        conn.Execute "SET FOREIGN_KEY_CHECKS=1"
            
            
        conn.Execute "commit"
        Espera 0.2
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set LOG = Nothing
    Set miRsAux = Nothing
    
    
    Me.lblIndicador.Caption = "Ajuste BD "
    Me.lblIndicador.Refresh
    Espera 0.5
    
    
    CargaGrid ""
    If Not Me.Adodc1.Recordset.EOF Then MsgBox "Llame soporte técnico", vbCritical
    
    Screen.MousePointer = vbDefault
End Sub


Private Function ActualizaArticuloENBD() As Boolean
Dim Donde As String
    On Error GoTo eActualizaArticuloENBD

    Donde = "Inicio"

    conn.Execute "SET FOREIGN_KEY_CHECKS=0;"
    
    Donde = "Actualizar sartic"
    
    
    
    
    

eActualizaArticuloENBD:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    conn.Execute "SET FOREIGN_KEY_CHECKS=1;"
    
End Function



Private Function TablasParametros(ByVal Campos As String, tabla As String, ByRef ColArticulos As Collection) As String
Dim Aux As String
Dim i As Byte
Dim k As Integer
    
    On Error GoTo eTablasParametros
    
    TablasParametros = ""
    Aux = "Select " & Campos & " FROM " & tabla
    miRsAux.Open Aux, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        For k = 1 To ColArticulos.Count
            Aux = ""
            For i = 0 To miRsAux.Fields.Count - 1
                Campos = DBLetMemo(miRsAux.Fields(Val(i)))
                If Campos = CStr(ColArticulos.Item(k)) Then Aux = Aux & "  " & miRsAux.Fields(i).Name
            Next
            If Aux <> "" Then TablasParametros = TablasParametros & vbCrLf & ColArticulos.Item(k) & Aux
        Next k
    End If
    miRsAux.Close
    
    
    If TablasParametros <> "" Then TablasParametros = vbCrLf & tabla & " " & TablasParametros
    
eTablasParametros:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar parametros", Err.Description
        Set miRsAux = Nothing
        Set miRsAux = New ADODB.Recordset
    End If
End Function
        



Private Function ActualizarPrecios(Venta As Boolean) As Boolean

    On Error GoTo eActualizarPrecios
    ActualizarPrecios = False
    
    CadenaConsulta = ""
    NumRegElim = -1
    If Venta Then
        CadenaConsulta = "UPDATE slista SET fechanue=" & DBSet(Now, "F") & ", precionu =" & DBSet(miRsAux!prevta, "N")
    Else
        CadenaConsulta = DevuelveDesdeBD(conAri, "codprove", "sartic", "codartic", miRsAux!codArtic, "T")
        If CadenaConsulta <> "" Then NumRegElim = Val(CadenaConsulta)
        
        
        CadenaConsulta = "UPDATE slispr SET fechanue=" & DBSet(Now, "F") & ", precionu =" & DBSet(miRsAux!precom, "N")
        If Not IsNull(miRsAux!Codprove) Then CadenaConsulta = CadenaConsulta & ", codprove = " & miRsAux!Codprove

    End If
    
    CadenaConsulta = CadenaConsulta & " WHERE codartic = " & DBSet(miRsAux!codArtic, "T")
    If Not Venta Then
        If NumRegElim >= 0 Then CadenaConsulta = CadenaConsulta & " AND codprove = " & NumRegElim
    End If
    
    conn.Execute CadenaConsulta
    
    
    
    If Not Venta Then
        If Not IsNull(miRsAux!Codprove) Then
            CadenaConsulta = "UPDATE slisp1 set codprove =" & miRsAux!Codprove & " WHERE codartic = " & DBSet(miRsAux!codArtic, "T")
            CadenaConsulta = CadenaConsulta & " AND codprove = " & NumRegElim
            conn.Execute CadenaConsulta
        End If
    End If
    
    
    
    ActualizarPrecios = True
eActualizarPrecios:
    If Err.Number <> 0 Then MuestraError Err.Number, "Avise soporte técnico"
    
End Function
