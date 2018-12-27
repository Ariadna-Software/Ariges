VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPortesMatriculasChofer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Matriculas"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13395
   Icon            =   "frmPortesMatriculas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   10080
      MaxLength       =   15
      TabIndex        =   15
      Text            =   "Ca"
      Top             =   9360
      Width           =   1545
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
      Left            =   4080
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "t|T|S|||smatriculas|titulo|||"
      Text            =   "Ca"
      Top             =   4920
      Width           =   1545
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPortesMatriculas.frx":000C
      Left            =   5040
      List            =   "frmPortesMatriculas.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Def|N|N|0||smatriculas|defecto|||"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   13
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
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   1170
      TabIndex        =   11
      Top             =   4950
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
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
      Index           =   0
      Left            =   900
      MaskColor       =   &H00000000&
      TabIndex        =   10
      ToolTipText     =   "Buscar Variedad"
      Top             =   4950
      Visible         =   0   'False
      Width           =   195
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
      Left            =   3060
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "Matricula|T|S|||smatriculas|matricula||S|"
      Text            =   "Ca"
      Top             =   4950
      Width           =   1545
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
      Left            =   10560
      TabIndex        =   4
      Top             =   9240
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   11880
      TabIndex        =   5
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
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
      Index           =   0
      Left            =   90
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "codigo transportista|N|N|0|999999|smatriculas|codenvio|0000|S|"
      Text            =   "Var"
      Top             =   4950
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPortesMatriculas.frx":0022
      Height          =   8145
      Left            =   135
      TabIndex        =   8
      Top             =   870
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   14367
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   11880
      TabIndex        =   9
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   9120
      Width           =   2385
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
         Left            =   45
         TabIndex        =   7
         Top             =   210
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2790
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13740
      TabIndex        =   14
      Top             =   180
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label LblMostr 
      BackStyle       =   0  'Transparent
      Caption         =   "MOSTRADOR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   3000
      TabIndex        =   16
      Top             =   9120
      Width           =   6735
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmPortesMatriculasChofer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es faça referència a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Private Const IdPrograma = 2023


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Transportista As String

Public VerMatriculas As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmEnv As frmFacFormasEnvio
Attribute frmEnv.VB_VarHelpID = -1


Public DeConsulta As Boolean

' utilizado para buscar por checks
Private BuscaChekc As String

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    BuscaChekc = ""
    
    B = (Modo = 2)
    If B Then
        PonLblIndicador lblIndicador, adodc1
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 2
        txtAux(i).visible = Not B
        txtAux(i).BackColor = vbWhite
    Next i
    
    txtAux2(2).visible = Not B
    btnBuscar(0).visible = Not B
    Me.Combo1.visible = Not B
    txtAux(3).Left = 15000
    txtAux(3).visible = Not B And Not VerMatriculas




    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    B = False
    If Modo = 4 Then
        B = True       'modificando
    Else
        If Transportista >= 0 Then B = True
    End If
    BloquearTxt txtAux(0), B
    btnBuscar(0).Enabled = B
    BloquearTxt txtAux(1), (Modo = 4)
    
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    B = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'Imprimir
    Toolbar1.Buttons(8).Enabled = B
    Me.mnImprimir.Enabled = B
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    

    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    For i = 1 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    
    txtAux2(2).Text = ""
    Combo1.ListIndex = 1
    
    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
    
    If Not VerMatriculas Then
        NumF = DevuelveDesdeBD(conAri, "max(chofer)", "sconductor", "1", "1")
        NumF = Val(NumF) + 1
        txtAux(3).Text = NumF
    End If
    
    If Transportista >= 0 Then
        txtAux(0).Text = Transportista
        txtAux2(2).Text = Me.Tag
        PonerFoco txtAux(1)

    Else
        'Ponemos el foco
        PonerFoco txtAux(0)
    End If
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "false"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i

    txtAux2(2).Text = ""
    Combo1.ListIndex = -1
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    If Transportista >= 0 Then
        txtAux(0).Text = Transportista
        txtAux2(2).Text = Me.Tag
        PonerFoco txtAux(1)
    Else
        PonerFoco txtAux(0)
    End If
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top '545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux2(2).Text = DataGrid1.Columns(1).Text
    txtAux(1).Text = DataGrid1.Columns(2).Text
    txtAux(2).Text = DataGrid1.Columns(3).Text
    Combo1.ListIndex = IIf(Trim(DataGrid1.Columns(4)) <> "", 0, 1)
    
    If Not VerMatriculas Then txtAux(3).Text = DataGrid1.Columns(5).Text
        
    
    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(2)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 2
        txtAux(i).Top = alto
    Next i
    

   
    txtAux2(2).Top = alto
    btnBuscar(0).Top = alto
    Me.Combo1.Top = alto

End Sub


Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SePuedeBorrar Then Exit Sub



    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el registro?"
    SQL = SQL & vbCrLf & "Transportista: " & adodc1.Recordset.Fields(0) & " " & adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & IIf(VerMatriculas, "Matricula", "Chofer") & ": " & adodc1.Recordset.Fields(3)
    If Not VerMatriculas Then SQL = SQL & vbCrLf & "DNI: " & adodc1.Recordset!dni
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        'matricula , codenvio ,titulo ,defecto
        If VerMatriculas Then
            SQL = "Delete from matriculas where matricula=" & DBSet(adodc1.Recordset!Matricula, "T")
        Else
            SQL = "Delete from sconductor  where chofer=" & DBSet(adodc1.Recordset!Chofer, "T") & "  " & DBSet(adodc1.Recordset.Fields(3), "T")
        End If
        SQL = SQL & " and codenvio = " & adodc1.Recordset!CodEnvio
        conn.Execute SQL
        
        SQL = DBLet(adodc1.Recordset!defec, "T")
        If Trim(SQL) <> "" Then
            SQL = "update  sconductor , (select max(chofer) chofer from sconductor where codenvio=" & adodc1.Recordset!CodEnvio & " ) aa"
            SQL = SQL & " Set sconductor.defecto = 1 where sconductor.chofer=aa.chofer"
            ejecutar SQL, False
        End If
        CargaGrid CadB
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
 TerminaBloquear
    
    Select Case Index
        Case 0 'variedades de comercial
            

    
        Case 1 '
            
  
            PonerFoco txtAux(Indice)
    
    End Select
    
    
End Sub


Private Sub cmdAceptar_Click()
    Dim i As Long
    
    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me, False)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    
                    PorDefecto
                    
                    CargaGrid CadB
                    If (DatosADevolverBusqueda <> "") Then 'And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            CadB = "dni"
                            If VerMatriculas Then CadB = "matricula"
                            CadB = CadB & " = " & DBSet(txtAux(1), "T")
                            SituarData adodc1, CadB, lblIndicador.Caption
                            CadB = ""
                            'SituarDataMULTI adodc1, "codenvio = " & txtAux(0) & " and matricula= " & DBSet(txtAux(1), "T"), lblIndicador.Caption
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If

                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 3) Then
                 TerminaBloquear
                    PorDefecto
                
                    
                    i = adodc1.Recordset.AbsolutePosition
                    PonerModo 2
                    CargaGrid CadB
                    If i > 0 Then adodc1.Recordset.Move i - 1
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
''    Else
''        lblIndicador.Caption = ""
'    End If
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    
    If VerMatriculas Then
        cad = adodc1.Recordset!Matricula & "|" & adodc1.Recordset!Titulo & "|"
    Else
        cad = adodc1.Recordset!Chofer & "|" & adodc1.Recordset!Nombre & "(" & adodc1.Recordset!dni & ")|"
    End If
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'If Modo = 2 Then PonerContRegIndicador lblIndicador, adodc1, CadB
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        
        If Transportista >= 0 Then
            BuscaChekc = DevuelveDesdeBD(conAri, "nomenvio", "senvio", "codenvio", CStr(Transportista))
            Me.Caption = Me.Caption & "    -- " & UCase(BuscaChekc)
            Me.Tag = BuscaChekc
        End If
        'If (DatosADevolverBusqueda <> "") Then   ' And NuevoCodigo <> "" Then
        '    BotonAnyadir
        'Else
            PonerModo 2
           
        'End If
    End If
End Sub

Private Sub Form_Load()

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        'el 1 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(8).Image = 10  'imprimir
    End With

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With


    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    If VerMatriculas Then
        Me.Caption = "Matriculas"
        '****************** canviar la consulta *********************************+
        CadenaConsulta = "SELECT  smatriculas.codenvio,nomenvio,matricula,titulo,if(defecto=1,'Si','') defec"
        CadenaConsulta = CadenaConsulta & " FROM smatriculas ,senvio WHERE smatriculas.codenvio=senvio.codenvio"
        '************************************************************************
        txtAux(0).Tag = "codigo transportista|N|N|0|999999|smatriculas|codenvio|0000|S|"
        txtAux(1).Tag = "Matricula|T|S|||smatriculas|matricula||S|"
        txtAux(2).Tag = "t|T|S|||smatriculas|titulo|||"
        txtAux(3).Tag = ""
       Combo1.Tag = "Def|N|N|0||smatriculas|defecto|||"
    Else
        Me.Caption = "Conductor / Chofer"
        '****************** canviar la consulta *********************************+
        CadenaConsulta = "SELECT  sconductor.codenvio,nomenvio,dni,nombre,if(defecto=1,'Si','') defec,chofer "
        CadenaConsulta = CadenaConsulta & " FROM sconductor ,senvio WHERE sconductor.codenvio=senvio.codenvio"
        '************************************************************************
        
        txtAux(0).Tag = "Codigo transportista|N|N|0|999999|sconductor|codenvio|0000|S|"
        txtAux(1).Tag = "Matricula|T|S|||sconductor|dni||S|"
        txtAux(2).Tag = "t|T|S|||sconductor|nombre|||"
        txtAux(3).Tag = "t|N|N|0||sconductor|chofer|||"
        Combo1.Tag = "Def|N|N|0||sconductor|defecto|||"
       
    End If
    LblMostr.Caption = Me.Caption
    CadB = ""
    CargaGrid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Calidad
    txtAux(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 2), "00") 'codcalid
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 3) 'nombre calidad
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedad comercial
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codvarie
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    
    'Preparamos para modificar
    '-------------------------
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 8
                'MsgBox "Imprimir...under construction"
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim Tabla As String
    Dim SQL As String
    Dim tots As String
    
    Tabla = "sconductor"
     If VerMatriculas Then Tabla = "smatriculas"
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " AND " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    
    If Transportista >= 0 Then
        SQL = SQL & " and " & Tabla & ".codenvio = " & Transportista
    End If
    
    
    '********************* canviar el ORDER BY *********************++
    If VerMatriculas Then
        SQL = SQL & " ORDER BY smatriculas.codenvio,matricula "
    Else
        SQL = SQL & " ORDER BY sconductor.codenvio,dni"
    End If
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Codigo|1000|;S|btnBuscar(0)|B|||;S|txtAux2(2)|T|Transportista|4900|;"
    
    If VerMatriculas Then
        tots = tots & "S|txtAux(1)|T|Matricula|2100|"
    Else
        tots = tots & "S|txtAux(1)|T|dni|2100|"
    End If
    
    tots = tots & ";S|txtAux(2)|T|" & IIf(VerMatriculas, "Titulo", "Nombre") & "|2900|;S|Combo1|C|Defecto|950|;"
    
    If Not VerMatriculas Then tots = tots & "N||||0|;"
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgRight
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            'LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'codigo de variedad
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(2).Text = PonerNombreDeCod(txtAux(Index), conAri, "senvio", "nomenvio", "codenvio", "N")
        
        Case 1 'codigo de calibre
        
            
        Case 2 'numlinea
           
        
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim SQL As String
Dim Mens As String

    B = CompForm(Me, 0)
    If Not B Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        'SQL = DevuelveDesdeBDNew(cAgro, "rcalidad_calibrador", "codcalid", "codvarie", txtAux(0).Text, "N", , "codcalid", txtAux(1).Text, "N", "numlinea", txtAux(2).Text, "N")
        If SQL <> "" Then
            MsgBox "Linea de calibrador existente para esta calidad. Reintroduzca.", vbExclamation
            PonerFoco txtAux(0)
            B = False
        End If
    End If
    
    If B And (Modo = 3 Or Modo = 4) Then

    End If
    
    If B And (Modo = 3) Then
        SQL = "select count(*) from matriculas where codenvio = " & DBSet(txtAux(0).Text, "N")
        SQL = SQL & " and matricula = " & DBSet(txtAux(1).Text, "T")
        
    
        If TotalRegistros(SQL) <> 0 Then
            MsgBox "Ya esta la matricula para el transportista. Revise.", vbExclamation
            PonerFoco txtAux(1)
            B = False
        End If
    End If
    
    If B And (Modo = 3 Or Modo = 4) Then
        
    End If
    DatosOk = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    'PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub printNou()
    
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 43 Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 0: KEYBusqueda KeyAscii, 0 'variedad
                Case 1: KEYBusqueda KeyAscii, 1 'calidad
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub

Private Sub PorDefecto()
Dim C As String
    
    If Combo1.ListIndex = 0 Then
         
        C = IIf(VerMatriculas, "smatriculas", "sconductor")
        C = "UPDATE " & C & " SET defecto=0 WHERE codenvio="
        If Modo = 3 Then
            C = C & txtAux(0).Text
        Else
            C = C & adodc1.Recordset!CodEnvio
        End If
        If VerMatriculas Then
            C = C & " AND matricula <> " & DBSet(txtAux(1).Text, "T")
        Else
            C = C & " AND chofer <> " & DBSet(txtAux(3).Text, "T")
        End If
        ejecutar C, True
    End If
End Sub


Private Function SePuedeBorrar() As Boolean
    SePuedeBorrar = False
    Screen.MousePointer = vbHourglass
    BuscaChekc = ""
    If VerMatriculas Then
        Stop
        'FALTA###
    Else
        For i = 1 To 3
            BuscaChekc = DevuelveDesdeBD(conAri, "count(*)", RecuperaValor("scaalb|schalb|scafac1|", i), "chofer", CStr(adodc1.Recordset!Chofer), "N")
            If Val(BuscaChekc) > 0 Then Exit For
        Next i
        If i > 3 Then BuscaChekc = ""
    End If
    
    If BuscaChekc <> "" Then
        MsgBox "Datos relacionados con este chófer", vbExclamation
    Else
        SePuedeBorrar = True
    End If
    Screen.MousePointer = vbDefault
End Function
