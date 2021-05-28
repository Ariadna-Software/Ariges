VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBasico2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario basico"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   7830
   Icon            =   "frmBasico2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRealizadasPorMi 
      Caption         =   "Realizadas por mi"
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
      Height          =   420
      Left            =   5280
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      Index           =   0
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4905
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CheckBox ChkCaduca 
      Caption         =   "Excluir Bloqueados-Caducados"
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
      Height          =   420
      Left            =   5520
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Frame FrameFiltro 
      Enabled         =   0   'False
      Height          =   705
      Left            =   4320
      TabIndex        =   20
      Top             =   45
      Visible         =   0   'False
      Width           =   2895
      Begin VB.ComboBox cboFiltro 
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
         ItemData        =   "frmBasico2.frx":000C
         Left            =   120
         List            =   "frmBasico2.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Enabled         =   0   'False
      Height          =   705
      Left            =   1755
      TabIndex        =   18
      Top             =   90
      Visible         =   0   'False
      Width           =   885
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   19
         Top             =   180
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Busqueda avanzada"
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
      Index           =   8
      Left            =   7110
      MaxLength       =   30
      TabIndex        =   10
      Top             =   4905
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
      Index           =   7
      Left            =   6390
      MaxLength       =   30
      TabIndex        =   9
      Top             =   4905
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
      Index           =   6
      Left            =   5625
      MaxLength       =   30
      TabIndex        =   8
      Top             =   4905
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
      Index           =   5
      Left            =   4860
      MaxLength       =   30
      TabIndex        =   7
      Top             =   4905
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
      Left            =   4095
      MaxLength       =   30
      TabIndex        =   6
      Top             =   4905
      Width           =   720
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
      Left            =   3195
      MaxLength       =   30
      TabIndex        =   5
      Top             =   4905
      Width           =   855
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
      Left            =   2370
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Descripción|T|N|||inciden|nomincid|||"
      Top             =   4920
      Width           =   765
   End
   Begin VB.Frame FrameBotonGnral 
      Enabled         =   0   'False
      Height          =   705
      Left            =   120
      TabIndex        =   16
      Top             =   90
      Visible         =   0   'False
      Width           =   1545
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   17
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
               Object.Tag             =   "0"
            EndProperty
         EndProperty
      End
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
      Left            =   5445
      TabIndex        =   2
      Tag             =   "   "
      Top             =   8745
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   6615
      TabIndex        =   3
      Top             =   8745
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Descripción|T|N|||inciden|nomincid|||"
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
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Código|N|N|0|999|inciden|codincid|000|S|"
      Top             =   4920
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBasico2.frx":0050
      Height          =   8145
      Left            =   135
      TabIndex        =   14
      Top             =   405
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   14367
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   23
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
      Left            =   6630
      TabIndex        =   15
      Top             =   8745
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   8595
      Width           =   2985
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   13
         Top             =   180
         Width           =   2895
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2205
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
         Shortcut        =   ^I
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
Attribute VB_Name = "frmBasico2"
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

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Public CadenaConsulta As String

Public pConn As Byte


Public Formulario As String
Public LenCta As Integer


Public Tag1 As String
Public Tag2 As String
Public Tag3 As String
Public Tag4 As String
Public Tag5 As String
Public Tag6 As String
Public Tag7 As String
Public Tag8 As String
Public Tag9 As String
Public Tag10 As String

Public Maxlen1 As Integer
Public Maxlen2 As Integer
Public Maxlen3 As Integer
Public Maxlen4 As Integer
Public Maxlen5 As Integer
Public Maxlen6 As Integer
Public Maxlen7 As Integer
Public Maxlen8 As Integer
Public Maxlen9 As Integer
Public Maxlen10 As Integer

Public CadenaTots As String
Public tabla As String
Public CampoCP As String
Public TipoCP As String
Public Report As String

Public Titulo As String

Private cadB As String

Private WithEvents frmProv As frmComProveedoresGr
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientesGr
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmBan As frmFacBancosPropios
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmSer As frmRepNumSerie2GR
Attribute frmSer.VB_VarHelpID = -1

Dim CampoOrden As String
Dim TipoOrden As String
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
Dim I As Integer

Dim vTag1 As cTag
Dim vTag2 As cTag
Dim vTag3 As cTag
Dim vTag4 As cTag
Dim vTag5 As cTag
Dim vTag6 As cTag
Dim vTag7 As cTag
Dim vTag8 As cTag
Dim vTag9 As cTag
Dim vTag10 As cTag

Dim cadFiltro As String
Dim cadFiltro2 As String


Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).BackColor = vbWhite
    Next I
    
    For I = 0 To txtAux.Count - 1
        If txtAux(I).Tag <> "" Then
            txtAux(I).visible = Not b
        Else
            txtAux(I).visible = False
        End If
    Next I
    If Combo1(0).Tag <> "" Then
        Combo1(0).visible = Not b
    Else
        Combo1(0).visible = False
    End If
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(1).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid 'primer de tot carregue tot el grid
    cadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    End If
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    For I = 1 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    cadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid CampoCP & " is null "
    
    'Prueba
    DataGrid1.ClearSelCols
    
    
    
    
    '*******************************************************************************
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    If tabla = "sdirpr" Then
        Combo1(0).ListIndex = -1
        Combo1(0).visible = True
        Combo1(0).Enabled = True
    End If
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    If Formulario <> "" And Formulario <> "Cuentas" Then
        PonerFoco txtAux(1)
    Else
        PonerFoco txtAux(0)
    End If
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    
    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For I = 0 To txtAux.Count - 1
        txtAux(I).Top = alto
    Next I
    
    If tabla = "sdirpr" Then
        Combo1(0).Top = alto
    End If
    ' ### [Monica] 12/09/2006
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el registro de " & Me.Caption & "?"
    SQL = SQL & vbCrLf & "Código: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Descripción: " & Adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Adodc1.Recordset.AbsolutePosition
        SQL = "Delete from " & tabla & " where " & CampoCP & "=" & Adodc1.Recordset.Fields(0).Value
        conn.Execute SQL
        CargaGrid cadB
        PonerModoOpcionesMenu
        Adodc1.Recordset.Cancel
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

Private Sub cboFiltro_Click()
    CargaGrid
End Sub


Private Sub ChkCaduca_Click()
    cadFiltro2 = ""
    If ChkCaduca.Value = 1 Then
        cadFiltro2 = "sartic.codstatu < 2"
    End If
    BotonVerTodos
End Sub

Private Sub cmdAceptar_Click()
    Dim I As Variant ' Integer

    Select Case Modo
        Case 1 'BUSQUEDA
            cadB = ObtenerBusqueda(Me, False)
            If cadB <> "" Then
                CargaGrid cadB
                PonerModo 2
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not Adodc1.Recordset.EOF Then
                            Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & DBSet(NuevoCodigo, vTag1.TipoDato))
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    cadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 0) Then
                    TerminaBloquear
                    I = Adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid cadB

                    Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & DBSet(I, RecuperaValor(Tag1, 2)))
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid cadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            cad = cad & Adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim cad As String
Dim CampoOrdenAnt As String

    If Adodc1.Recordset Is Nothing Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    
    CampoOrdenAnt = CampoOrden
    
    If ColIndex <= 8 Then
        Me.Refresh
        DoEvents
        Screen.MousePointer = vbHourglass
        Select Case ColIndex
            Case 0
                CampoOrden = RecuperaValor(txtAux(0).Tag, 7)
            Case 1
                CampoOrden = RecuperaValor(txtAux(1).Tag, 7)
            Case 2
                CampoOrden = RecuperaValor(txtAux(2).Tag, 7)
            Case 3
                CampoOrden = RecuperaValor(txtAux(3).Tag, 7)
            Case 4
                CampoOrden = RecuperaValor(txtAux(4).Tag, 7)
            Case 5
                CampoOrden = RecuperaValor(txtAux(5).Tag, 7)
            Case 6
                CampoOrden = RecuperaValor(txtAux(6).Tag, 7)
            Case 7
                CampoOrden = RecuperaValor(txtAux(7).Tag, 7)
            Case 8
                CampoOrden = RecuperaValor(txtAux(8).Tag, 7)
        End Select
        If CampoOrden <> "" Then
            Select Case TipoOrden
                Case "ASC"
                    TipoOrden = "DESC"
                Case "DESC"
                    TipoOrden = "ASC"
            End Select
        Else
            'dejamos el grid ordenado por el que estaba anteriormente
            CampoOrden = CampoOrdenAnt
        End If
        CargaGrid cadB
        Screen.MousePointer = vbDefault
        Else
        MsgBox "Error cargando tabla. Imposible ordenacion", vbCritical
    End If

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    PonerContRegIndicador
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
'        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'            BotonAnyadir
        If Formulario = "Articulos" Then
            ChkCaduca.Left = Me.Width - ChkCaduca.Width - 300 'ChkCaduca.Left + 4700
        End If
        
        
        If DeConsulta Then
            PonerModo 2
            If Me.CodigoActual <> "" Then
                SituarDataNew Me.Adodc1, CampoCP & "=" & DBSet(Me.CodigoActual, Me.TipoCP), ""
            End If
        Else
            If Formulario <> "" And Formulario <> "Cuentas" Then 'And Formulario <> "Bancos" Then
                BotonBuscar
            Else
                PonerModo 2
                If Me.CodigoActual <> "" Then
                    SituarDataNew Me.Adodc1, CampoCP & "=" & DBSet(Me.CodigoActual, Me.TipoCP), ""
                End If
            End If
        End If
    End If
End Sub



Private Sub Form_Load()
    PrimeraVez = True
    Me.Caption = Titulo
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2   'Todos
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 24  'busqueda experta
    End With

    
    txtAux(0).Tag = Tag1
    txtAux(1).Tag = Tag2
    txtAux(2).Tag = Tag3
    txtAux(3).Tag = Tag4
    txtAux(4).Tag = Tag5
    txtAux(5).Tag = Tag6
    txtAux(6).Tag = Tag7
    txtAux(7).Tag = Tag8
    txtAux(8).Tag = Tag9
    Combo1(0).Tag = Tag10
    
    If Tag1 <> "" Then txtAux(0).TabIndex = 0
    If Tag2 <> "" Then txtAux(1).TabIndex = 1
    If Tag3 <> "" Then txtAux(2).TabIndex = 2
    If Tag4 <> "" Then txtAux(3).TabIndex = 3
    If Tag5 <> "" Then txtAux(4).TabIndex = 4
    If Tag6 <> "" Then txtAux(5).TabIndex = 5
    If Tag7 <> "" Then txtAux(6).TabIndex = 6
    If Tag8 <> "" Then txtAux(7).TabIndex = 7
    If Tag9 <> "" Then txtAux(8).TabIndex = 8
    If Tag10 <> "" Then Combo1(0).TabIndex = 9
    
    If tabla = "sdirpr" Then
        Combo1(0).TabIndex = 1
    End If
    
    txtAux(0).MaxLength = Maxlen1
    txtAux(1).MaxLength = Maxlen2
    txtAux(2).MaxLength = Maxlen3
    txtAux(3).MaxLength = Maxlen4
    txtAux(4).MaxLength = Maxlen5
    txtAux(5).MaxLength = Maxlen6
    txtAux(6).MaxLength = Maxlen7
    txtAux(7).MaxLength = Maxlen8
    txtAux(8).MaxLength = Maxlen9
    
    CampoOrden = CampoCP
    TipoOrden = "ASC"
    
    '[Monica]27/06/2019: cargamos el filtro
    If Formulario = "Socios" Then
        FrameFiltro.visible = True
        FrameFiltro.Enabled = True
        FrameFiltro.Left = FrameFiltro.Left + 5700
    
        CargaFiltros
        PosicionarCombo cboFiltro, 1
    End If
    
    If Formulario = "Articulos" Then
        ChkCaduca.visible = Not DeConsulta
        ChkCaduca.Enabled = Not DeConsulta
    End If
    
    cadB = ""
    If DeConsulta Then
        CargaGrid
    Else
        If Formulario <> "" And Formulario <> "Cuentas" Then 'And Formulario <> "Bancos" Then
            CargaGrid "false"
        Else
            CargaGrid
        End If
    End If
    Set vTag1 = New cTag
    vTag1.Cargar txtAux(0)
    Set vTag2 = New cTag
    vTag2.Cargar txtAux(1)
    Set vTag3 = New cTag
    vTag3.Cargar txtAux(2)
    Set vTag4 = New cTag
    vTag4.Cargar txtAux(3)
    Set vTag5 = New cTag
    vTag5.Cargar txtAux(4)
    Set vTag6 = New cTag
    vTag6.Cargar txtAux(5)
    Set vTag7 = New cTag
    vTag7.Cargar txtAux(6)
    Set vTag8 = New cTag
    vTag8.Cargar txtAux(7)
    Set vTag9 = New cTag
    vTag9.Cargar txtAux(8)
    Set vTag10 = New cTag
    vTag10.Cargar Combo1(0)
            
           
    If pConn = 2 And Formulario = "Cuentas" Then
        FrameBotonGnral.visible = True
        FrameBotonGnral.Enabled = True
        
        Me.DataGrid1.Top = Me.DataGrid1.Top + 500
        Me.Frame1(1).Top = Me.Frame1(1).Top + 500
        Me.cmdAceptar.Top = Me.cmdAceptar.Top + 500
        Me.cmdCancelar.Top = Me.cmdCancelar.Top + 500
        Me.cmdRegresar.Top = Me.cmdRegresar.Top + 500
        
        Me.Height = Me.Height + 500
    End If
                       
    '[Monica]26/06/2019: boton de poder buscar en el propio mantenimiento
    If DeConsulta Then
        FrameBotonGnral2.visible = False
        FrameBotonGnral2.Enabled = False
    Else
        If Formulario <> "" And Formulario <> "Cuentas" Then
            FrameBotonGnral2.visible = True
            FrameBotonGnral2.Enabled = True
        End If
    End If
    
    If tabla = "sdirpr" Then CargaCombo
    
End Sub

Private Sub CargaCombo()
    Combo1(0).Clear
    Combo1(0).AddItem "Albaran"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Combo1(0).AddItem "Factura"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
    Set vTag1 = Nothing
    Set vTag3 = Nothing
End Sub



Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If Adodc1.Recordset.EOF Then Exit Sub
    
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
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
                mnBuscar_Click
        Case 2
                mnVerTodos_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim SQL As String
    Dim tots As String
    
    
    If pConn = 1 Then
        Adodc1.ConnectionString = conn
    Else
        Adodc1.ConnectionString = ConnConta
    End If
    
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " AND " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    
    '[Monica]27/06/2019: solo para la ayuda de socios para no tener en cuenta los que tienen fecha de baja
    If cboFiltro.ListIndex <> -1 Then
        CargarSqlFiltro
        
        SQL = SQL & " and " & cadFiltro
    End If
    
    If cadFiltro2 <> "" Then
        SQL = SQL & " and " & cadFiltro2
    End If
    
    
    '********************* canviar el ORDER BY *********************++
    SQL = SQL & " ORDER BY " & CampoOrden & " " & TipoOrden
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.Adodc1, SQL, PrimeraVez
    
    
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = CadenaTots
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.RowHeight = 350

End Sub

Private Sub CargaGridGnral(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, SQL As String, PrimeraVez As Boolean)
On Error GoTo ECargaGrid

    vDataGrid.Enabled = True
    '    vdata.Recordset.Cancel
'    vData.ConnectionString = conn
    vData.RecordSource = SQL
    vData.CursorType = adOpenDynamic
    vData.LockType = adLockPessimistic
    vDataGrid.ScrollBars = dbgNone
    vData.Refresh
    
    Set vDataGrid.DataSource = vData
    vDataGrid.AllowRowSizing = False
    vDataGrid.AllowUpdate = False
    vDataGrid.RowHeight = 290

    If PrimeraVez Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "CargaGrid", Err.Description
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Formulario <> "" And Formulario <> "Cuentas" Then
        AbrirFormulario Formulario
    End If
End Sub

Private Sub AbrirFormulario(vFormulario As String)
    Select Case vFormulario
        Case "Proveedores"
            VariePublic = ""
            
            Set frmProv = New frmComProveedoresGr
            
            frmProv.DatosADevolverBusqueda = "0|"
            frmProv.Show vbModal

            Set frmProv = Nothing

            cadB = VariePublic
            If cadB <> "" Then
                cadB = "sprove.codprove = " & VariePublic
                CargaGrid cadB
                PonerModo 2
                cmdRegresar_Click
            End If
            
        Case "Articulos"
            VariePublic = ""
            
            frmAlmArticulosGr.DatosADevolverBusqueda = "@1@"
            frmAlmArticulosGr.Show vbModal

            cadB = VariePublic
            If cadB <> "" Then
                cadB = "sartic.codartic = " & DBSet(VariePublic, "T")
                CargaGrid cadB
                PonerModo 2
                cmdRegresar_Click
            End If
        
        Case "Clientes"
            VariePublic = ""
            
            Set frmCli = New frmFacClientesGr
            
            frmCli.DatosADevolverBusqueda = "0|"
            frmCli.Show vbModal

            Set frmCli = Nothing

            cadB = VariePublic
            If cadB <> "" Then
                cadB = "sclien.codclien = " & VariePublic
                CargaGrid cadB
                PonerModo 2
                cmdRegresar_Click
            End If
        
        Case "Series"
            VariePublic = ""
            
            Set frmSer = New frmRepNumSerie2GR
            
            frmSer.DatosADevolverBusqueda = "0|"
            frmSer.Show vbModal

            Set frmSer = Nothing

            cadB = VariePublic
            If cadB <> "" Then
                cadB = "sserie.numserie = " & VariePublic
                CargaGrid cadB
                PonerModo 2
                cmdRegresar_Click
            End If
        
        
        
'        Case "Bancos"
'            VariePublic = ""
'
'            Set frmBan = New frmFacBancosPropios
'
'            frmBan.DatosADevolverBusqueda = "0|"
'            frmBan.Show vbModal
'
'            Set frmBan = Nothing
'
'            cadB = VariePublic
'            If cadB <> "" Then
'                cadB = "sbanpr.codbanpr = " & VariePublic
'                CargaGrid cadB
'                PonerModo 2
'                cmdRegresar_Click
'            End If
        
            
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    If txtAux(Index) = "" Then Exit Sub
    
    Select Case Index
        Case 0
            If vTag1.Formato <> "" Then
                Select Case vTag1.TipoDato
                    Case "N"
                        txtAux(Index).Text = Format(txtAux(Index).Text, vTag1.Formato)
                    Case "F"
                        PonerFormatoFecha txtAux(Index)
                End Select
            End If
            
            If Formulario = "Cuentas" And LenCta = 0 Then
                txtAux(1).Text = PonerNombreCuenta(txtAux(Index), 3, "")
            End If
            
        Case 1
            'txtAux(Index).Text = UCase(txtAux(Index).Text)
            If vTag2.Formato <> "" Then
                Select Case vTag2.TipoDato
                    Case "N"
                        txtAux(Index).Text = Format(txtAux(Index).Text, vTag2.Formato)
                    Case "F"
                        PonerFormatoFecha txtAux(Index)
                End Select
            End If
            
        Case 2
            If vTag3.Formato <> "" Then
                Select Case vTag3.TipoDato
                    Case "N"
                        txtAux(Index).Text = Format(txtAux(Index).Text, vTag3.Formato)
                    Case "F"
                        PonerFormatoFecha txtAux(Index)
                End Select
            End If
            
        Case 3
            If vTag4.Formato <> "" Then
                Select Case vTag4.TipoDato
                    Case "N"
                        txtAux(Index).Text = Format(txtAux(Index).Text, vTag4.Formato)
                    Case "F"
                        PonerFormatoFecha txtAux(Index)
                End Select
            End If
            
        Case 4
            If vTag5.Formato <> "" Then
                Select Case vTag5.TipoDato
                    Case "N"
                        txtAux(Index).Text = Format(txtAux(Index).Text, vTag5.Formato)
                    Case "F"
                        PonerFormatoFecha txtAux(Index)
                End Select
            End If
        
        Case 5
            If vTag6.Formato <> "" Then
                Select Case vTag6.TipoDato
                    Case "N"
                        txtAux(Index).Text = Format(txtAux(Index).Text, vTag6.Formato)
                    Case "F"
                        PonerFormatoFecha txtAux(Index)
                End Select
            End If
        
        Case 6
            If vTag7.Formato <> "" Then
                Select Case vTag7.TipoDato
                    Case "N"
                        txtAux(Index).Text = Format(txtAux(Index).Text, vTag7.Formato)
                    Case "F"
                        PonerFormatoFecha txtAux(Index)
                End Select
            End If
            
        Case 7
            If vTag8.Formato <> "" Then
                Select Case vTag8.TipoDato
                    Case "N"
                        txtAux(Index).Text = Format(txtAux(Index).Text, vTag8.Formato)
                    Case "F"
                        PonerFormatoFecha txtAux(Index)
                End Select
            End If
            
        Case 8
            If vTag9.Formato <> "" Then
                Select Case vTag9.TipoDato
                    Case "N"
                        txtAux(Index).Text = Format(txtAux(Index).Text, vTag9.Formato)
                    Case "F"
                        PonerFormatoFecha txtAux(Index)
                End Select
            End If
            
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim SQL As String
Dim Mens As String


    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtAux(0)) Then b = False
    End If
    
    DatosOk = b
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Adodc1)
        If cadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub CargaFiltros()
Dim Aux As String
    
    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Sólo Altas "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
    cboFiltro.AddItem "Sólo Bajas"
    cboFiltro.ItemData(cboFiltro.NewIndex) = 2

End Sub
    
    
Private Sub cboFiltro_KeyPress(KeyAscii As Integer)
    CargaGrid
End Sub
    
Private Sub cboFiltro_Change()
    CargaGrid
End Sub
    
Private Sub CargarSqlFiltro()

    Screen.MousePointer = vbHourglass
    cadFiltro = ""
    
    Select Case Me.cboFiltro.ListIndex
        Case -1, 0 ' sin filtro
            cadFiltro = "(1=1)"
        
        Case 1 ' solo altas
            cadFiltro = "(rsocios.fechabaja is null or rsocios.fechabaja = '') "
        
        Case 2 ' solo bajas
            cadFiltro = " not (rsocios.fechabaja is null or rsocios.fechabaja = '') "
    
    End Select
    Screen.MousePointer = vbDefault

End Sub

Private Function SituarDataNew(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String, Optional Refresca As Boolean) As Boolean
'Situa un DataControl en el registo que cumple vwhere
'para cuando la clave primaria esta formada por 1 campo
On Error GoTo ESituarData
        'Actualizamos el recordset
        If Refresca Then vData.Refresh
        vData.Recordset.MoveFirst
        'El sql para que se situe en el registro en especial es el siguiente
        vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataNew = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        If vData.Recordset.RecordCount > 0 Then vData.Recordset.MoveFirst
        SituarDataNew = False
End Function

