VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacLotesGeneralitat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control lotes subvencionados"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14565
   Icon            =   "frmFacLotesGeneralitat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   6
      Left            =   7920
      MaxLength       =   18
      TabIndex        =   6
      Tag             =   "NºSerie|T|N|||slotesgeneralitat|numserie|||"
      Text            =   "ser"
      Top             =   2640
      Width           =   675
   End
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   36
      Top             =   2640
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   35
      Top             =   2640
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox txtAux2 
      Height          =   315
      Index           =   7
      Left            =   12000
      TabIndex        =   15
      Text            =   "co"
      Top             =   5040
      Width           =   1395
   End
   Begin VB.TextBox txtAux2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   5760
      TabIndex        =   11
      Text            =   "cantidad"
      Top             =   5640
      Width           =   1155
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||slotesgeneralitat|fecha|dd/mm/yyyy||"
      Text            =   "Fecha alb"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   4200
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Arti|T|N|||slotesgeneralitat|codartic|||"
      Text            =   "Descripcion"
      Top             =   2640
      Width           =   1395
   End
   Begin VB.TextBox txtAux2 
      Height          =   1755
      Index           =   8
      Left            =   9240
      TabIndex        =   16
      Text            =   "24348588Y"
      Top             =   5760
      Width           =   4755
   End
   Begin VB.TextBox txtAux2 
      Height          =   315
      Index           =   6
      Left            =   10440
      TabIndex        =   14
      Text            =   "co"
      Top             =   5040
      Width           =   1395
   End
   Begin VB.TextBox txtAux2 
      Height          =   315
      Index           =   5
      Left            =   9240
      TabIndex        =   13
      Text            =   "co"
      Top             =   5040
      Width           =   1035
   End
   Begin VB.TextBox txtAux2 
      Height          =   315
      Index           =   4
      Left            =   9240
      TabIndex        =   12
      Text            =   "24348588Y"
      Top             =   4320
      Width           =   4755
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   6720
      MaxLength       =   16
      TabIndex        =   5
      Text            =   "Nomprov"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   5880
      MaxLength       =   18
      TabIndex        =   4
      Tag             =   "Prov|N|N|||slotesgeneralitat|codprove|||"
      Text            =   "Prov"
      Top             =   2640
      Width           =   675
   End
   Begin VB.CommandButton cmdAux2 
      Caption         =   "+"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
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
      Index           =   2
      Left            =   4080
      TabIndex        =   28
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
      TabIndex        =   9
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
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Text            =   "Tasa"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4560
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
      Index           =   8
      Left            =   9480
      TabIndex        =   8
      Tag             =   "Cantidad|N|N|||slotesgeneralitat|cantidad|0.00|N|"
      Text            =   "cantidad"
      Top             =   2640
      Width           =   1155
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   7
      Left            =   8760
      MaxLength       =   18
      TabIndex        =   7
      Tag             =   "Lote|T|N|||slotesgeneralitat|numlote|||"
      Text            =   "lote"
      Top             =   2640
      Width           =   675
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12780
      TabIndex        =   18
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11580
      TabIndex        =   17
      Top             =   7800
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Cod|N|N|0||slotesgeneralitat|id|0000|S|"
      Text            =   "id"
      Top             =   2640
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   2520
      MaxLength       =   16
      TabIndex        =   2
      Tag             =   "Arti|T|N|||slotesgeneralitat|codartic|||"
      Text            =   "codartic"
      Top             =   2640
      Width           =   1395
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12780
      TabIndex        =   24
      Top             =   7800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   21
      Top             =   7695
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
         TabIndex        =   22
         Top             =   240
         Width           =   2280
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   14565
      _ExtentX        =   25691
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
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3840
         TabIndex        =   23
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
      Bindings        =   "frmFacLotesGeneralitat.frx":000C
      Height          =   2805
      Left            =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   540
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   4948
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
      Bindings        =   "frmFacLotesGeneralitat.frx":0021
      Height          =   3405
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6006
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
   Begin VB.Label Label3 
      Caption         =   "Movimientos lote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2400
      TabIndex        =   34
      Top             =   3600
      Width           =   11415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   4
      Left            =   9240
      TabIndex        =   33
      Top             =   5520
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo carnet "
      Height          =   195
      Index           =   3
      Left            =   12000
      TabIndex        =   32
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fec. vigencia"
      Height          =   195
      Index           =   2
      Left            =   10440
      TabIndex        =   31
      Top             =   4800
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nº carnet "
      Height          =   195
      Index           =   1
      Left            =   9240
      TabIndex        =   30
      Top             =   4800
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre carnet manipulador"
      Height          =   195
      Index           =   0
      Left            =   9240
      TabIndex        =   29
      Top             =   4080
      Width           =   1950
   End
   Begin VB.Label Label1 
      Caption         =   "Movimientos lote"
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
      TabIndex        =   26
      Top             =   3600
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
Attribute VB_Name = "frmFacLotesGeneralitat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public DeConsulta As Boolean 'Muestra Form para consulta, solo buscar y ver todos activos

Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmC As frmBasico2
Attribute frmC.VB_VarHelpID = -1
Private WithEvents FrmArt As frmBasico2   'Form Articulos
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmProv As frmBasico2 '%=%=frmComProveedores
Attribute frmProv.VB_VarHelpID = -1

Dim PrimeraVez As Boolean


Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos


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

Dim i As Integer

Private Sub PonerModo(vModo As Byte)
Dim b As Boolean
    
    
    
    ActualizarToolbarGnral Me.Toolbar1, Modo, vModo, 5
    
    Modo = vModo
    b = (Modo = 2)
    PonerIndicador Me.lblIndicador, Modo
    b = Modo = 1 Or Modo = 3 Or Modo = 4
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = b
    Next i
    Me.cmdAux(0).visible = False
    Me.cmdAux(1).visible = False
    
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
    BloquearTxt txtAux(0), (Modo <> 1), True
    BloquearTxt txtAux(3), True
    BloquearTxt txtAux(5), True
    
    b = False
    If Modo = 5 Then
        If ModificaLineas > 0 Then MsgBox "Hb st"
    End If
    Campos_2_Visibles b
    
    
    
    Label3.visible = Modo = 5
    
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
        For i = 0 To txtAux2.Count - 1
            txtAux2(i).Text = ""
        Next
        
        txtAux2(0).Text = Format(Now, "dd/mm/yyyy")
        Campos_2_Visibles True
        anc = ObtenerAlto(DataGrid2, 10)
        LLamaLineas2 anc
        PonerFoco txtAux2(0)
        
    Else
        PonerModo 3
        'Situamos el grid al final
        
        AnyadirLinea DataGrid1, adodc1
        CargaGrid2 False
        anc = ObtenerAlto(DataGrid1, 10)
        
        'Obtenemos la siguiente numero de factura
        LimpiarCampos
    
        
        LLamaLineas anc, 1
        BloquearTxt txtAux(2), False
        txtAux(0).Text = "0"
        'Ponemos el foco
        PonerFoco txtAux(0)
    End If
End Sub


Private Sub BotonBuscar()
    PonerModo 1
    CargaGrid "id= -1"
    LimpiarCampos
    
    LLamaLineas DataGrid1.Top + 220, 0
    PonerFoco txtAux(0)
End Sub


Private Sub BotonVerTodos()
On Error Resume Next
    CargaGrid ""
    If adodc1.Recordset.RecordCount <= 0 Then
         'MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        ' MsgBox "No hay ningún registro en la tabla sunida", vbInformation
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

        DeseleccionaGrid DataGrid2
            
        ModificaLineas = 2
        PonerBotonCabecera False
        'Los txts
        For i = 0 To txtAux2.Count - 1
             txtAux2(i).Text = DataGrid2.Columns(i).Text
         Next i
        
        
        
        
        Campos_2_Visibles True
        anc = ObtenerAlto(DataGrid2, 10)
        LLamaLineas2 anc
        
        
        PonerFoco txtAux2(3)
        txtAux2_GotFocus 3
    Else
         If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
             i = DataGrid1.Bookmark - DataGrid1.FirstRow
             DataGrid1.Scroll 0, i
             DataGrid1.Refresh
         End If
         
         anc = ObtenerAlto(DataGrid1, 10)
         PonerModo 4
         cad = ""
         For i = 0 To 2
             cad = cad & DataGrid1.Columns(i).Text & "|"
         Next i
         'Llamamos al form
         For i = 0 To txtAux.Count - 1
            txtAux(i).Text = DataGrid1.Columns(i).Text
         Next
         LLamaLineas anc, 2
         
        'SI ya tiene lineas, NO dejo cambiar el articulo
        i = IIf(Adodc2.Recordset.EOF, 0, 1)
        BloquearTxt txtAux(2), i = 1
        Me.cmdAux(0).visible = i = 0
        Me.cmdAux(1).visible = i = 0
         
         
         PonerFoco txtAux(1)
   End If
   Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    
  
    
    
    
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next
    cmdAux(0).Top = alto
    cmdAux(1).Top = alto
    
    txtAux(0).Left = DataGrid1.Left + 340
    txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
    txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 65
    txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 65
    txtAux(4).Left = txtAux(3).Left + txtAux(3).Width + 65
    txtAux(5).Left = txtAux(4).Left + txtAux(4).Width + 65
    txtAux(6).Left = txtAux(5).Left + txtAux(5).Width + 70
    txtAux(7).Left = txtAux(6).Left + txtAux(6).Width + 65
    txtAux(8).Left = txtAux(7).Left + txtAux(7).Width + 65
    If Modo = 3 Then
        cmdAux(0).visible = True
        cmdAux(0).Left = txtAux(3).Left - 90
        cmdAux(1).visible = True
        cmdAux(1).Left = txtAux(5).Left - 90
    Else
        cmdAux(0).visible = False
        cmdAux(1).visible = False
    End If
End Sub

Private Sub LLamaLineas2(alto As Single)
    
    For i = 0 To 3
        txtAux2(i).Top = alto
    Next
    cmdAux2.Top = alto
    txtAux2(0).Left = DataGrid2.Left + 340
    cmdAux2.Left = txtAux2(0).Left + txtAux2(0).Width + 15
    
    txtAux2(1).Left = txtAux2(0).Left + txtAux2(0).Width + 65
    txtAux2(2).Left = txtAux2(1).Left + txtAux2(1).Width + 65
    txtAux2(3).Left = txtAux2(2).Left + txtAux2(2).Width + 65

End Sub



Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2

    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    If Modo = 5 Then
        If Adodc2.Recordset.EOF Then Exit Sub
        SQL = "¿Seguro que desea eliminar la linea de entrega de articulo/lote? " & vbCrLf
        SQL = SQL & vbCrLf & "Fecha: " & Adodc2.Recordset.Fields(0)
        SQL = SQL & vbCrLf & "Cliente: " & Adodc2.Recordset.Fields(2)
        SQL = SQL & vbCrLf & "Cantidad: " & Adodc2.Recordset.Fields(3)
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
        SQL = "DELETE FROM slotesgeneralitatmov"
        SQL = SQL & " WHERE idlote =" & adodc1.Recordset!ID & " AND idmov=" & Adodc2.Recordset!idmov
        conn.Execute SQL
        CargaGrid2 True
        PonerFora True
    Else
        'Eliminar normal
        SQL = DevuelveDesdeBD(conAri, "idlote", "slotesgeneralitatmov", "idlote", CStr(adodc1.Recordset!ID))
        If SQL <> "" Then
            MsgBox "Existen movimientos de ese LOTE", vbExclamation
            Exit Sub
        End If
        
        '### a mano
        SQL = "¿Seguro que desea eliminar el lote? " & vbCrLf

        
        
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            NumRegElim = Me.adodc1.Recordset.AbsolutePosition
            'Hay que eliminar
            SQL = "Delete from slotesgeneralitat where id=" & adodc1.Recordset!ID
            conn.Execute SQL
            limpiar Me
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


Private Function InserarModificar() As Boolean
Dim vtag As cTag
Dim SQL As String

    On Error GoTo eInserarModificar
    InserarModificar = False
    CadenaConsulta = "(ID"
    SQL = ""
    If Modo = 3 Then SQL = Val(DevuelveDesdeBD(conAri, "max(id)", "slotesgeneralitat", "1", "1")) + 1
    Set vtag = New cTag
    'vtAG(0) NO entra
    For i = 1 To Me.txtAux.Count - 1
        If Not (i = 3 Or i = 5) Then '3l 3 es nomartic
            vtag.Cargar txtAux(i)
                            
            If Modo = 3 Then
                'INSERTAR
                SQL = SQL & ", " & DBSet(txtAux(i).Text, vtag.TipoDato, vtag.Vacio)
                CadenaConsulta = CadenaConsulta & "," & vtag.columna
            
            Else
                
                SQL = SQL & "," & vtag.columna & " = " & DBSet(txtAux(i).Text, vtag.TipoDato, vtag.Vacio)
                
            
            
            End If
            
            
        End If
    Next i
    
    If Modo = 3 Then
        SQL = "INSERT INTO slotesgeneralitat" & CadenaConsulta & ") VALUES (" & SQL & ")"
    
    Else
        SQL = Mid(SQL, 2)
        SQL = "UPDATE slotesgeneralitat SET " & SQL & " WHERE id =" & adodc1.Recordset!ID
    End If
    
    conn.Execute SQL
    InserarModificar = True
    
eInserarModificar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        
    End If
    Set vtag = Nothing
    CadenaConsultaSelect
End Function


Private Sub Adodc2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    PonerFora False
End Sub

Private Sub PonerFora(limpiar As Boolean)


    On Error Resume Next
    
    If limpiar Then
        For i = 4 To txtAux2.Count - 1
            txtAux2(i).Text = ""
        Next i
    Else
        If Modo = 2 Or Modo = 5 Then
            If Not Adodc2.Recordset.EOF Then
                'Ponemos campos foragrid
                For i = 4 To txtAux2.Count - 1
                    txtAux2(i).Text = DataGrid2.Columns(i).Text
                Next i
            End If
        End If
    End If
    Err.Clear
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cadB As String
On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            txtAux(0).Text = "0"
            If DatosOk Then
                If InserarModificar Then
                    CargaGrid
                    BotonAnyadir
                End If
            End If

        Case 4  'MODIFICAR
            If DatosOk Then
                If InserarModificar Then
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
                    NumRegElim = Adodc2.Recordset!idmov
                    CargaGrid2 True
                    Adodc2.Recordset.Find (" idmov =" & NumRegElim)
    
                    PonerBotonCabecera True
                    PonerFocoBtn Me.cmdAceptar
                    ModificaLineas = 0
                    txtAux2(8).Enabled = False
                    PonerModo 5
                Else
                    CargaGrid2 True
                    BotonAnyadir
                End If
             End If
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdAux_Click(Index As Integer)
    
    
    CadenaConsulta = ""
    If Index = 0 Then
        Set FrmArt = New frmBasico2
        'FrmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en Modo busqueda
'        FrmArt.DesdeTPV = False
'        FrmArt.Show vbModal
        AyudaArticulos FrmArt, txtAux(2).Text
        Set FrmArt = Nothing
        If CadenaConsulta <> "" Then
            Me.txtAux(2).Text = RecuperaValor(CadenaConsulta, 1)
            Me.txtAux(3).Text = RecuperaValor(CadenaConsulta, 2)
        End If
    
    
    Else
        'Proveedor
        
'        Set frmProv = New frmComProveedores
'        frmProv.DatosADevolverBusqueda = "1"
'        frmProv.Show vbModal
        Set frmProv = New frmBasico2
        AyudaProveedores frmProv, txtAux(4)
        Set frmProv = Nothing
        
        If CadenaConsulta <> "" Then
            Me.txtAux(4).Text = RecuperaValor(CadenaConsulta, 1)
            Me.txtAux(5).Text = RecuperaValor(CadenaConsulta, 2)
        End If
    End If
    
    CadenaConsultaSelect
End Sub

Private Sub cmdAux2_Click()
Dim cad As String
Dim MostrarAutorizados As Boolean

        cad = ""
        MostrarAutorizados = False
        'Llamamos al manipulador de carnet fitosnaitarios
        'NO tiene puesto cliente
        If txtAux2(1).Text = "" Then
            'LLAMAMOS A CLIENTE
              'Cliente
            CadenaConsulta = ""
            Set frmC = New frmBasico2
            AyudaClientes frmC, txtAux2(1)
            Set frmC = Nothing
        
            If CadenaConsulta = "" Then
                CadenaConsultaSelect
                Exit Sub
            End If
            txtAux2(1).Text = RecuperaValor(CadenaConsulta, 1)
            txtAux2(2).Text = RecuperaValor(CadenaConsulta, 2)
            cad = RecuperaValor(CadenaConsulta, 1)
        Else
            cad = txtAux2(1).Text
            CadenaConsulta = cad & "|" & txtAux2(2).Text & "|"
        End If
        
        'Veremos si tiene autrizados
        cad = DevuelveDesdeBD(conAri, "count(*)", "sclienmani", "codclien", cad)
        
        'Si tiene autirzados muestro el frm de seleeccionar
        If Val(cad) > 0 Then
            MostrarAutorizados = True
            CadenaDesdeOtroForm = ""
        Else
            'Voy a poner los datos del cliente ya que NO tiene autirzados
            'ListView1.ListItems(NumRegElim).SubItems(4) = IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
            cad = "concat(coalesce(ManipuladorNumCarnet,''),'|',coalesce(DATE_FORMAT(ManipuladorFecCaducidad,'%d/%m/%Y'),''),'|',coalesce(nomclien,''),'|'"
            cad = cad & ",coalesce(If(ManipuladortipoCarnet = 2, ""Cualificado"", ""Básico""),'|'),'|')"
            cad = DevuelveDesdeBD(conAri, cad, "sclien", "codclien", RecuperaValor(CadenaConsulta, 1))
            
            If Len(cad) <= 4 Then
                cad = ""
            Else
                If Trim(RecuperaValor(cad, 1)) = "" Then cad = ""
            End If
            
            If cad = "" Then
                MsgBox "Sin carnet de manipulador ni autorizados" & vbCrLf & RecuperaValor(CadenaConsulta, 2), vbExclamation
                txtAux2(2).Text = ""
                txtAux2(1).Text = ""
                CadenaConsultaSelect
                Exit Sub
            End If
            
            
            If txtAux2(1).Text = "" Then
                txtAux2(1).Text = RecuperaValor(CadenaConsulta, 1)
                txtAux2(2).Text = RecuperaValor(CadenaConsulta, 2)
            End If
            CadenaDesdeOtroForm = cad
        End If
        CadenaConsultaSelect
            
            
                    
        
        If MostrarAutorizados Then
            frmFitoCarnet.Cliente = -1
            If Val(txtAux2(1).Text) > 0 Then frmFitoCarnet.Cliente = Val(txtAux2(1).Text)
            frmFitoCarnet.Show vbModal
        End If
        If CadenaDesdeOtroForm <> "" Then
        
            If txtAux2(0).Text <> "" Then
                If CDate(RecuperaValor(CadenaDesdeOtroForm, 2)) < CDate(txtAux2(0).Text) Then
                    If MsgBox("Carnet caducado.  ¿Desea continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                End If
            End If
            
            'numcarnet|feccad|Nombre|Básico|

            
            Me.txtAux2(4).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
            Me.txtAux2(5).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Me.txtAux2(6).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            
            'IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
            Me.txtAux2(7).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
            
            'Ponemos foco en cantidad
            PonerFoco txtAux2(3)
        End If
        
   
End Sub

Private Sub cmdCancelar_Click()
On Error Resume Next

    Select Case Modo
    Case 3 'Insertar
        DataGrid1.AllowAddNew = False
        'CargaGrid
        If Not adodc1.Recordset.EOF Then
            adodc1.Recordset.MoveFirst
            CargaGrid2 True
            Modo = 2
            PonerFora False
        Else
            limpiar Me
        End If
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
        txtAux2(8).Enabled = False
        CargaGrid2 True

        PonerFora False
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
        cmdCancelar.Cancel = True
        
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



Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

On Error GoTo Error1

    If Not adodc1.Recordset.EOF Then lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
        
        
    
    PonerFora True
    
    If Modo = 2 Or Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not adodc1.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Devuelve = DevuelveDesdeBDNew(conAri, NomTablaLineas, "ampliaci", "numpedcl", Text1(0).Text, "N", , "numlinea", Data2.Recordset!numlinea, "N")
            'Poner descripcion de ampliacion lineas
            CargaGrid2 True
            PonerFora False
       
        End If
    End If
    
Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
        
        

End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    
'    If Not Adodc2.Recordset.EOF Then
'        If Not DGrid_CambiarFila(DataGrid2) Then Exit Sub
'    End If
    
    If Not Adodc2.Recordset.EOF And ModificaLineas <> 1 Then PonerFora False
    
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        
        'Cadena consulta
        CadenaConsultaSelect  'Pone el select
        CargaGrid
        PonerModo 2
        PonerFora False
        PrimeraVez = False
    End If
End Sub


Private Sub Form_Load()

    PrimeraVez = True
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
    
   
    DataGrid2.visible = True
    Label1.visible = True
    Me.Toolbar1.Buttons(9).visible = True
    

    '## A mano
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    CadAncho = False
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    Modo = 0
    ModificaLineas = 0
    limpiar Me
    
    
End Sub


Private Sub CadenaConsultaSelect()
    CadenaConsulta = "select id,fecha,slotesgeneralitat.codartic,nomartic,"
    'slotesgeneralitat.numserie,slotesgeneralitat.fecvigen,"
    CadenaConsulta = CadenaConsulta & " slotesgeneralitat.codprove,nomprove, "
    CadenaConsulta = CadenaConsulta & " slotesgeneralitat.numserie,numlote ,cantidad from slotesgeneralitat,sartic,sprove"
    CadenaConsulta = CadenaConsulta & " where slotesgeneralitat.codartic=sartic.codartic AND slotesgeneralitat.codprove=sprove.codprove"
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    txtAux2(0).Text = RecuperaValor(CadenaDevuelta, 1)
    txtAux2(1).Text = RecuperaValor(CadenaDevuelta, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
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
                If Modo <> 2 Then Exit Sub
                frmListado3.Opcion = 65
                frmListado3.Show vbModal
                
               
        Case 11: mnSalir_Click
    End Select
End Sub

Private Sub MtoLineas()
    
    If Modo <> 2 Then Exit Sub
    If adodc1.Recordset.EOF Then Exit Sub
    ModificaLineas = 0
    
    Label3.Caption = adodc1.Recordset!NomArtic & "   LOTE: " & adodc1.Recordset!numLote
    
    
    PonerModo 5
    PonerBotonCabecera True
End Sub
Private Sub CargaGrid(Optional SQL As String)
Dim i As Byte
Dim b As Boolean
    
    b = DataGrid1.Enabled
    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY id"
    
    CargaGridGnral DataGrid1, Me.adodc1, SQL, False
    
    'select id,fecha,codartic,nomartic,,numlote,cantidad,horamov,numserie,fecvigen from slotesgeneralitat
    ' id, fecha, codartic, nomartic, serie, fecvigen,lote ,cantidad
    i = 0 'Cod.
        DataGrid1.Columns(i).Caption = "ID"
        DataGrid1.Columns(i).Width = 700
    
    i = 1 'Fecha
        DataGrid1.Columns(i).Caption = "Fecha"
        DataGrid1.Columns(i).Width = 1100
    
    
    i = 2 'Artic
        DataGrid1.Columns(i).Caption = "Codartic"
        DataGrid1.Columns(i).Width = 1400
        
    i = 3 '
        DataGrid1.Columns(i).Caption = "Nombre"
        DataGrid1.Columns(i).Width = 3000
    
    i = 4 '
        DataGrid1.Columns(i).Caption = "Prov"
        DataGrid1.Columns(i).Width = 850
    
    i = 5 '
        DataGrid1.Columns(i).Caption = "Nombre proveedor"
        DataGrid1.Columns(i).Width = 2400
    
    i = 6 '
        DataGrid1.Columns(i).Caption = "NºSerie"
        DataGrid1.Columns(i).Width = 1200
    
    
    
    i = 7 '
        DataGrid1.Columns(i).Caption = "Lote"
        DataGrid1.Columns(i).Width = 1200
                        
            
    i = 8 'Lote
        DataGrid1.Columns(i).Caption = "Cantidad"
        DataGrid1.Columns(i).Width = 1300
        DataGrid1.Columns(i).Alignment = dbgRight
        DataGrid1.Columns(i).NumberFormat = FormatoCantidad
            
            
            
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        For i = 0 To DataGrid1.Columns.Count - 1
            txtAux(i).Width = DataGrid1.Columns(i).Width - 60
            txtAux(i).Height = Me.DataGrid1.RowHeight - 10
        Next i
        
        'txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        'txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        'txtAux(2).Width = DataGrid1.Columns(2).Width - 60
        'txtAux(3).Width = DataGrid1.Columns(3).Width - 30
        
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
Dim devuelve As String
    If Modo = 1 Then Exit Sub
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    Select Case Index
    Case 0
        PonerFormatoEntero txtAux(Index) 'Cod. Tipo Unidad
    Case 1
        PonerFormatoFecha txtAux(Index)
         
    Case 2
        
        devuelve = ""
        If txtAux(Index).Text <> "" Then
            devuelve = DevuelveDesdeBD(conAri, "nomartic", "sartic", "artvario=0 AND codartic", txtAux(Index).Text, "T")
            If devuelve = "" Then MsgBox "no existe artículo", vbExclamation
        End If
        
        If devuelve = "" Then
            If txtAux(Index).Text <> "" Then
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
        End If
        txtAux(3).Text = devuelve
    Case 4
        devuelve = ""
        If txtAux(Index).Text <> "" Then
            devuelve = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtAux(Index).Text, "N")
            If devuelve = "" Then MsgBox "no existe proveedor", vbExclamation
        End If
        
        If devuelve = "" Then
            If txtAux(Index).Text <> "" Then
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
        End If
        txtAux(5).Text = devuelve
    Case 8
         If Not PonerFormatoDecimal(txtAux(Index), 3) Then txtAux(Index).Text = ""   'Cod. Tipo Unidad
    End Select
  
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
 '   If cerrar Then Unload Me  de momneto lo comento
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
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
        cmdRegresar.Cancel = True
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.cmdCancelar.Cancel = True
        Campos_2_Visibles False
        Me.lblIndicador.Caption = ""
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Campos_2_Visibles(visibles As Boolean)

Dim J As Integer
    For J = 0 To 3
        txtAux2(J).visible = visibles
    Next
    
    
    For J = 4 To txtAux2.Count - 1
        txtAux2(J).Enabled = False
    Next
    txtAux2(8).Enabled = visible
    Me.cmdAux2.visible = visibles

End Sub

Private Sub CargaGrid2(enlaza As Boolean)
Dim i As Byte
Dim b As Boolean
Dim SQL As String


    If Not Label1.visible Then Exit Sub

    b = DataGrid2.Enabled
    DataGrid2.Enabled = False
    SQL = "select slotesgeneralitatmov.fechaMov ,slotesgeneralitatmov.codclien,nomclien,cantidad"
    SQL = SQL & " ,slotesgeneralitatmov.ManipuladorNombre,slotesgeneralitatmov.ManipuladorNumCarnet,"
    SQL = SQL & " slotesgeneralitatmov.ManipuladorFecCaducidad ,If(TipoCarnet = 2, ""Cualificado"", ""Básico"") , observa, idMov"
    SQL = SQL & " from slotesgeneralitatmov,sclien where slotesgeneralitatmov.codclien=sclien.codclien and idLote = "
    If enlaza Then
        SQL = SQL & adodc1.Recordset!ID
    Else
        SQL = SQL & " -1"
    End If
    SQL = SQL & " ORDER BY idMov"
    

    
    CargaGridGnral DataGrid2, Me.Adodc2, SQL, PrimeraVez
    
    'select id,codartic,fecha,numlote,cantidad,horamov,numserie,fecvigen from slotesgeneralitat
  

    i = 0 '
        DataGrid2.Columns(i).Caption = "Fecha"
        DataGrid2.Columns(i).Width = 1200
    
    i = 1 '
        DataGrid2.Columns(i).Caption = "Socio"
        DataGrid2.Columns(i).Width = 1300
        DataGrid2.Columns(i).NumberFormat = "0000"
        
    i = 2 '
        DataGrid2.Columns(i).Caption = "Nombre"
        DataGrid2.Columns(i).Width = 3400
        
        
    i = 3 'Cantidad
        DataGrid2.Columns(i).Caption = "Cantidad"
        DataGrid2.Columns(i).Width = 1000
        DataGrid2.Columns(i).Alignment = dbgRight
        DataGrid2.Columns(i).NumberFormat = FormatoCantidad
    
    For i = 4 To DataGrid2.Columns.Count - 1
        DataGrid2.Columns(i).visible = False
    Next
    
    'Fiajamos el cadancho
    
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux2(0).Width = DataGrid2.Columns(0).Width - 60
        txtAux2(1).Width = DataGrid2.Columns(1).Width - 60
        txtAux2(2).Width = DataGrid2.Columns(2).Width - 60
        txtAux2(3).Width = DataGrid2.Columns(3).Width - 30
        
    
   
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

    Select Case Index
    Case 0
        PonerFormatoFecha txtAux2(Index)
            
    
    Case 1
        cad = ""
        
        If PonerFormatoEntero(txtAux2(Index)) Then
            
            cad = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtAux2(Index))
            If cad = "" Then
                MsgBox "No existe el cliente: " & txtAux2(Index).Text, vbExclamation
                txtAux2(Index) = ""
                PonerFoco txtAux2(Index)
                
            Else
                txtAux2(Index + 1) = cad
                cmdAux2_Click
            End If
            
        End If
        
        If cad = "" Then
            For i = 1 To 7
                If i <> 3 Then txtAux2(i).Text = ""
            Next
        End If

        
    
    Case 3
        ' lo que ponga en su TAG  (8)
        If Not PonerFormatoDecimal(txtAux2(Index), 3) Then txtAux2(Index).Text = ""
    End Select
End Sub




Private Function InsertarModificar() As Boolean
Dim C As String
Dim Suma As Currency
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    For NumRegElim = 0 To 7
        txtAux2(NumRegElim).Text = Trim(txtAux2(NumRegElim).Text)
        If txtAux2(NumRegElim).Text = "" Then
            MsgBox "Campos son obligatorios (Excepto observaciones)", vbExclamation
            Exit Function
        End If
    Next
    
    'Llega AQUI. Vemos si la suma de lo que hay es esto
    C = ""
    If ModificaLineas = 2 Then C = " idmov<>" & Adodc2.Recordset!idmov & " AND "
    C = C & " idLote "
    C = DevuelveDesdeBD(conAri, "sum(cantidad)", "slotesgeneralitatmov", C, CStr(adodc1.Recordset!ID))
    If C = "" Then C = "0"
    Suma = ImporteFormateado(C)
    
    'Veamos si la suma mas lo que hay aqui da bastante
    If Suma + ImporteFormateado(txtAux2(3).Text) > adodc1.Recordset!cantidad Then
        If MsgBox("Total vendido excede cantidad lote. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    
    'idMov|idLote|
    CadenaDesdeOtroForm = "fechaMov|codclien|#|cantidad|ManipuladorNombre|ManipuladorNumCarnet|ManipuladorFecCaducidad|TipoCarnet|observa|"
     CadenaConsulta = ""
    C = ""
    For i = 0 To txtAux2.Count - 1
        If ModificaLineas = 1 Then
            'INSERTAR
            If i = 0 Or i = 6 Then
                C = C & ", " & DBSet(txtAux2(i).Text, "F")
            ElseIf i = 3 Then
                C = C & ", " & DBSet(txtAux2(i).Text, "N")
            ElseIf i = 2 Then
                'NADA
            Else
                C = C & ", " & DBSet(txtAux2(i).Text, "T")
            End If
        Else
            'MODIFICAR
            If i <> 2 Then
                C = C & ", " & RecuperaValor(CadenaDesdeOtroForm, i + 1) & " = "
                If i = 0 Or i = 6 Then
                    C = C & DBSet(txtAux2(i).Text, "F")
                ElseIf i = 3 Then
                    C = C & DBSet(txtAux2(i).Text, "N")
                Else
                    C = C & DBSet(txtAux2(i).Text, "T")
                End If
                
                
            End If
        End If
     
    
        
    Next i
    
    If ModificaLineas = 1 Then
        
        CadenaDesdeOtroForm = Replace(CadenaDesdeOtroForm, "|#|", "|")  'el campo|#| para la modificacion
        CadenaConsulta = Replace(CadenaDesdeOtroForm, "|", ",")
        CadenaConsulta = Mid(CadenaConsulta, 1, Len(CadenaConsulta) - 1) 'quitamos el ultimo pipe
        CadenaConsulta = "INSERT INTO slotesgeneralitatmov(idMov,idLote," & CadenaConsulta
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "max(idMov)", "slotesgeneralitatmov", "idlote", CStr(adodc1.Recordset!ID))
        CadenaConsulta = CadenaConsulta & ") VALUES (" & Val(CadenaDesdeOtroForm) + 1 & "," & CStr(adodc1.Recordset!ID)
        C = CadenaConsulta & C & ")"
    
    Else
        C = Mid(C, 2)
        C = "UPDATE slotesgeneralitatmov SET " & C
        C = C & " WHERE idmov=" & Adodc2.Recordset!idmov & " AND idlote =" & adodc1.Recordset!ID
    End If
    conn.Execute C
    InsertarModificar = True

EInsertarModificar:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    CadenaConsultaSelect
End Function




