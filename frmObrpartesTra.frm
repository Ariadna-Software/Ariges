VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmObrpartesTra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Partes  diarios de trabajo"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   15060
   Icon            =   "frmObrpartesTra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOT 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   8280
      TabIndex        =   30
      ToolTipText     =   "Buscar artículo"
      Top             =   5640
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmObrpartesTra.frx":000C
      Left            =   11640
      List            =   "frmObrpartesTra.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   8
      Left            =   10200
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "nombre direc"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   7
      Left            =   9360
      MaxLength       =   16
      TabIndex        =   7
      Text            =   "direc"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmObrpartesTra.frx":0022
      Left            =   12360
      List            =   "frmObrpartesTra.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   5
      Left            =   12840
      MaxLength       =   16
      TabIndex        =   10
      Text            =   "cantidad"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   2520
      MaxLength       =   16
      TabIndex        =   4
      Text            =   "direc"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   6
      Left            =   3840
      Locked          =   -1  'True
      MaxLength       =   45
      TabIndex        =   29
      Text            =   "nombre artic"
      Top             =   5520
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox THoras 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FEF7E4&
      Height          =   315
      Left            =   13920
      MaxLength       =   7
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   550
      Width           =   855
   End
   Begin VB.CommandButton cmdAux1 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   2160
      TabIndex        =   14
      ToolTipText     =   "Buscar artículo"
      Top             =   5040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   5400
      MaxLength       =   50
      TabIndex        =   5
      Text            =   "actua"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   6
      Text            =   "desc"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   2
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   19
      Text            =   "nombre direc"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   3
      Text            =   "codclien"
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   12600
      TabIndex        =   11
      Top             =   5595
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13800
      TabIndex        =   12
      Top             =   5595
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   13800
      TabIndex        =   26
      Top             =   5595
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Top             =   5430
      Width           =   3000
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
         TabIndex        =   25
         Top             =   180
         Width           =   2595
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   550
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   2
      Tag             =   "Trabajador|N|N|0||scaparte|codtraba|000|N|"
      Text            =   "Text1"
      Top             =   550
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||scaparte|fecha|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   550
      Width           =   975
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
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
            Object.ToolTipText     =   "Lineas"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Actualizar"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   23
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8280
      Top             =   480
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FEF7E4&
      Height          =   315
      Index           =   0
      Left            =   960
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "Nº parte|N|N|0||scaparte|numparte|0000000|S|"
      Text            =   "Text1"
      Top             =   550
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   7560
      Top             =   480
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmObrpartesTra.frx":0038
      Height          =   4320
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   7620
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Alignment       =   1  'Right Justify
      Caption         =   "Total Horas"
      Height          =   255
      Left            =   12960
      TabIndex        =   28
      Top             =   600
      Width           =   855
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   5160
      Picture         =   "frmObrpartesTra.frx":004D
      ToolTipText     =   "Buscar almacen"
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   2520
      Picture         =   "frmObrpartesTra.frx":014F
      ToolTipText     =   "Buscar fecha"
      Top             =   585
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "Cód. Trabajador"
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   555
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Parte "
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   555
      Width           =   855
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
      TabIndex        =   17
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
Attribute VB_Name = "frmObrpartesTra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'--------------------------------------------------------------------------

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores 'Mto de Trabajadores
Attribute frmT.VB_VarHelpID = -1
'Private WithEvents frmCl As frmFacClientes
Private WithEvents frmAc As frmObraActua
Attribute frmAc.VB_VarHelpID = -1
Private WithEvents frmOT As frmObraOT
Attribute frmOT.VB_VarHelpID = -1


Dim NombreTabla As String
Dim NomTablaLineas As String
Dim Ordenacion As String

Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CadenaConsulta As String
Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe

Private HaDevueltoDatos As Boolean



Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()

On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        Text1(kCampo).BackColor = vbWhite
        cadSeleccion = ""
        HacerBusqueda
        
    Case 3 'INSERTAR
        If DatosOk() Then
            If InsertarDesdeForm(Me) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonLineas
                BotonAnyadirLineas
            End If
        End If
    Case 4 'MODIFICAR
        If DatosOk() Then
             If ModificaDesdeFormulario(Me, 1) Then
                 TerminaBloquear
                 PosicionarData
             End If
         End If
            
    Case 5 'LINEAS Traspaso Almacenes
        If InsertarModificarLinea Then
        
            
        
            'Reestablecemos los campos
            'y ponemos el grid
            DataGrid1.AllowAddNew = False
            If ModificaLineas = 2 Then
                TerminaBloquear
                NumRegElim = Data2.Recordset!linea
            End If
            CargaGrid True
            
            PonerSumaHoras
            
            If ModificaLineas = 1 Then 'Insertar
                ModificaLineas = 0
                BotonAnyadirLineas
            ElseIf ModificaLineas = 2 Then 'Modificar
                Data2.Recordset.Find (" linea =" & NumRegElim)
                ModificaLineas = 0
                PonerBotonCabecera True
                CargaTxtAux False, False
                Me.lblIndicador.Caption = ""
                DataGrid1.Enabled = True
'                DataGrid1.SetFocus
                PonerFocoGrid Me.DataGrid1
            End If
        End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub cmdAux1_Click()

    cadSeleccion = ""
    Set frmAc = New frmObraActua
    'Nos retornara SIEMPRE codclien,coddirec,actua
    'Pero le enivaremos datos para montar la busqueda
    frmAc.DatosADevolverBusqueda = txtAux(0).Text & "|" & txtAux(1).Text & "|" & txtAux(3).Text & "|"
    frmAc.Show vbModal
    Set frmAc = Nothing
    If cadSeleccion <> "" Then
        'cadSeleccion
        For NumRegElim = 0 To 2
            If NumRegElim = 2 Then
                txtAux(NumRegElim + 1).Text = RecuperaValor(cadSeleccion, NumRegElim + 1)
            Else
                txtAux(NumRegElim).Text = RecuperaValor(cadSeleccion, NumRegElim + 1)
            End If
            txtAux_LostFocus CInt(NumRegElim)
            
        Next
        PonerFoco txtAux(4)
    End If
    cadSeleccion = ""
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Mantenimiento Lineas traspasos
            CargaTxtAux False, False
            DataGrid1.Enabled = True
            DataGrid1.AllowAddNew = False
            If Not ModificaLineas = 2 Then 'Modificar
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
             PonerBotonCabecera True
            DataGrid1.Refresh
            PonerFocoBtn Me.cmdRegresar
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdOT_Click()
    cadSeleccion = ""
    Set frmOT = New frmObraOT
    frmOT.DatosADevolverBusqueda = "0|1|"
    frmOT.Show vbModal
    Set frmOT = Nothing
    If cadSeleccion <> "" Then
        txtAux(7).Text = RecuperaValor(cadSeleccion, 1)
        txtAux(8).Text = RecuperaValor(cadSeleccion, 2)
        PonerFoco txtAux(7)
    End If
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then 'modo 5: Mantenimiento Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        Me.cmdRegresar.visible = False
    End If
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
    btnAnyadir = 5 'Posicion del boton Añadir en la toolbar1
    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(9).Image = 10 'Mantenimiento Líneas
        '.Buttons(10).Image = 39 'Actualizar
        .Buttons(12).Image = 16 'Imprimir
        .Buttons(13).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
       
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'campo situacio solo en tabla scatra

    
    cadSeleccion = ""
    
  
        NombreTabla = "scaparte"
        NomTablaLineas = "sliparte"
        

    
    Ordenacion = " ORDER BY numparte"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE numparte = -1"

    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    
    CargaGrid False
    
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim B As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo ECarga

    B = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, False
      
    DataGrid1.Columns(0).visible = False 'linea

    
    I = 1
    'Cod. Artículo
    DataGrid1.Columns(I).Caption = "Cod clien"
    DataGrid1.Columns(I).Width = 900
    DataGrid1.Columns(I).NumberFormat = "00000"
    
    'Nombre Artículo
    I = I + 1
    DataGrid1.Columns(I).Caption = "Cod.Obra"
    DataGrid1.Columns(I).Width = 900
    DataGrid1.Columns(I).NumberFormat = "0000"
    I = I + 1
    DataGrid1.Columns(I).Caption = "Obra"
    DataGrid1.Columns(I).Width = 2500
    
    
    I = I + 1
    DataGrid1.Columns(I).Caption = "Actuacion"
    DataGrid1.Columns(I).Width = 1200
    
    
    I = I + 1
    DataGrid1.Columns(I).Caption = "Descripción"
    DataGrid1.Columns(I).Width = 3200
    
    
    I = I + 1
    DataGrid1.Columns(I).Caption = "OT"
    DataGrid1.Columns(I).Width = 700
    I = I + 1
    DataGrid1.Columns(I).Caption = "Trabajo"
    DataGrid1.Columns(I).Width = 2000
    
    
    I = I + 1
    DataGrid1.Columns(I).Caption = "P.T."
    DataGrid1.Columns(I).Width = 600
    
    I = I + 1
    DataGrid1.Columns(I).Caption = "P.P."
    DataGrid1.Columns(I).Width = 600
    
    'Cantidad
    I = I + 1
    DataGrid1.Columns(I).Caption = "Cantidad"
    DataGrid1.Columns(I).Width = 1400
    DataGrid1.Columns(I).Alignment = dbgRight
    DataGrid1.Columns(I).NumberFormat = FormatoImporte & " "
    I = I + 1
    DataGrid1.Columns(I).visible = False 'linea
    DataGrid1.Columns(I + 1).visible = False 'linea
    DataGrid1.Columns(I + 2).visible = False 'linea
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
    Next I
       
    DataGrid1.Enabled = B
    DataGrid1.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim I As Byte
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        alto = 290
        For I = 0 To txtAux.Count - 1
            If I <> 6 Then txtAux(I).Top = alto
        Next I
        cmdAux1.Top = alto
        cmdOT.Top = alto
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
            Next I
            Combo1.ListIndex = -1
            Combo2.ListIndex = -1
        End If
        
        If ModificaLineas = 1 Then 'Insertar
            For I = 0 To txtAux.Count - 1
                
                BloquearTxt txtAux(I), I = 2 Or I = 6 Or I = 8 ' he puesto un 7 en vez de un 6 masl
                
            Next I
           
        ElseIf ModificaLineas = 2 Then
            'Poner valor a los txtAux
            
            For I = 0 To 4
                txtAux(I).Text = DataGrid1.Columns(I + 1).Text
            Next I
            txtAux(I).Text = DataGrid1.Columns(I + 5).Text  'i es 5
            Combo1.ListIndex = DBLet(Data2.Recordset!partet, "N")
            Combo2.ListIndex = DBLet(Data2.Recordset!partep, "N")
            For I = 7 To 8
                txtAux(I).Text = DataGrid1.Columns(I - 1).Text ' he cambiado 2 por 3 masl
            Next I
            
        End If
        
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        
        
        'Fijamos altura y posición Top
        cmdAux1.Top = alto
        cmdAux1.Height = DataGrid1.RowHeight
        cmdOT.Top = alto
        cmdOT.Height = DataGrid1.RowHeight
        
        For I = 0 To txtAux.Count - 1  'EL 6 no lo toco
            If I <> 6 Then
                txtAux(I).Top = alto
                txtAux(I).Height = DataGrid1.RowHeight
            End If
        Next I
        
        
        'Fijamos anchura y posicion Left
        txtAux(0).Left = DataGrid1.Left + 340 'codartic
        txtAux(0).Width = DataGrid1.Columns(1).Width - 35
        txtAux(1).Left = DataGrid1.Columns(2).Left + DataGrid1.Left + 15
        txtAux(1).Width = DataGrid1.Columns(2).Width - 35
        
        

        For I = 2 To 8  ' el ultimo NO esta aqui
            If I <> 6 Then
                If I = 5 Then
                    txtAux(I).Left = DataGrid1.Columns(I + 5).Left + DataGrid1.Left + 15
                    txtAux(I).Width = DataGrid1.Columns(I + 5).Width - 35
                Else
                
                    If I > 6 Then
                
                        txtAux(I).Left = DataGrid1.Columns(I - 1).Left + DataGrid1.Left + 15
                        txtAux(I).Width = DataGrid1.Columns(I - 1).Width - 35
                        If I = 8 Then
                            cmdOT.Left = txtAux(8).Left - 150
                        End If
                    Else
                        txtAux(I).Left = DataGrid1.Columns(I + 1).Left + DataGrid1.Left + 15
                        txtAux(I).Width = DataGrid1.Columns(I + 1).Width - 35
                    End If
                    If I = 3 Then
                        txtAux(I).Width = txtAux(I).Width - 170
                        cmdAux1.Left = txtAux(3).Left + txtAux(3).Width
                    End If
                End If
            End If
        Next I
        Me.Combo1.Top = alto
        Combo1.Left = DataGrid1.Columns(8).Left + DataGrid1.Left + 15
        Me.Combo2.Top = alto
        Combo2.Left = DataGrid1.Columns(9).Left + DataGrid1.Left + 15
    End If

    'Los ponemos Visibles o No
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = visible
    Next I
    cmdAux1.visible = visible
    cmdOT.visible = visible
    Me.Combo1.visible = visible
    Me.Combo2.visible = visible
    
End Sub



Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    cadSeleccion = CadenaSeleccion
End Sub

'Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
''Almacenes Propios
'Dim Indice As Byte
'    Indice = CByte(Me.imgBuscar(0).Tag)
'    Text1(Indice + 2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
'    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)
'End Sub
'
'Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
''Mantenimiento de Articulos
'    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
'    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
'End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Modo <> 5 Then 'Estamos en Cabecera
            'Recupera todo el registro de Traspaso Almacenes
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else 'Estamos en Lineas
            'Llamamos desde el boton auxiliar de Artículos
            txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
            txtAux(1).Text = RecuperaValor(CadenaDevuelta, 2)
            PonerFoco txtAux(2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmOT_DatoSeleccionado(CadenaSeleccion As String)
    cadSeleccion = CadenaSeleccion
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Trabajadores
    cadSeleccion = CadenaSeleccion
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    cadSeleccion = ""
    Select Case Index

        Case 2  'Cod. Trabajador
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
    End Select
    If cadSeleccion <> "" Then
        Text1(Index).Text = Format(RecuperaValor(cadSeleccion, 1), "0000")
        Text2(Index).Text = RecuperaValor(cadSeleccion, 2)
        PonerFocoBtn cmdAceptar
    End If
    cadSeleccion = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte

   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = Index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   indice = 1
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(1)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then   'Eliminar lineas Traspaso Almacenes
        BotonEliminarLinea
    Else 'Eliminar Cabecera Traspaso Almacenes
        BotonEliminar
    End If
End Sub

Private Sub mnModificar_Click()
    If Modo = 5 Then  'Modificar lineas Traspaso Almacenes
        If BLOQUEADesdeFormulario(Me) Then BotonModificarLinea
    Else 'Modificar Cabecera Traspaso Almacenes
        If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then  'Añadir lineas Traspaso Almacenes
        BotonAnyadirLineas
    Else 'Añadir Cabecera Traspaso Almacenes
        BotonAnyadir
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index <> 5 Then ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And Index = 5 And Modo = 1 Then
        PonerFocoBtn cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub Text1_LostFocus(Index As Integer)
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Bloquear el contador si no estamos en busquedas
    If (Modo <> 1) And (Index = 0) Then BloquearTxt Text1(0), True, True
    cadSeleccion = ""
    Select Case Index
        Case 0 'Codigo Traspaso Almacen
            
            If Not PonerFormatoEntero(Text1(Index)) Then Text1(Index).Text = ""
                
        Case 1 'Fecha
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 2
        
            If PonerFormatoEntero(Text1(Index)) Then
                cadSeleccion = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
            End If
            If cadSeleccion = "" Then
                If Text1(Index).Text <> "" Then PonerFoco Text1(Index)
                Text1(Index).Text = ""
                
            End If
            Text2(Index).Text = cadSeleccion
            
        Case 5 'Observaciones
            If Text1(Index).Text <> "" Then Text1(Index).Text = QuitarCaracterEnter(Text1(Index).Text)
    End Select
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub



Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
  
        KEYpress KeyAscii

End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String
    
    
    
    'Quitar espacios en blanco por los lados
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    devuelve = ""
    Select Case Index
        Case 0 'Cliente
            If txtAux(0).Text <> "" Then
                If PonerFormatoEntero(txtAux(0)) Then
                    devuelve = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtAux(0).Text)
                    If devuelve = "" Then MsgBox "NO existe el cliente: " & txtAux(0).Text, vbExclamation
                End If
                If devuelve = "" Then
                    If txtAux(0).Text <> "" Then txtAux(0).Text = ""
                    PonerFoco txtAux(0)
                End If
            End If
            txtAux(6).Text = devuelve
        Case 1
              'Cod direc. Si no esta el cliente no hacemos nada
              If txtAux(1).Text <> "" Then
                 'Tiene datos
                 If txtAux(0).Text = "" Then
                    MsgBox "Ponga primero el cliente", vbExclamation
                    PonerFoco txtAux(0)
                 Else
                    If PonerFormatoEntero(txtAux(1)) Then
                        devuelve = "codclien = " & txtAux(0).Text & " AND coddirec "
                        devuelve = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", devuelve, txtAux(1).Text)
                        If devuelve = "" Then MsgBox "NO existe la obra: " & txtAux(1).Text, vbExclamation
                    End If
                    If devuelve = "" Then
                        If txtAux(1).Text <> "" Then txtAux(1).Text = ""
                        PonerFoco txtAux(1)
                    End If
                  End If
              End If
              txtAux(2).Text = devuelve
        Case 3
            'ACTUACION.  Tiene que existir.
            If txtAux(3).Text <> "" Then
                 If txtAux(1).Text = "" Or txtAux(0).Text = "" Then
                    MsgBox "Ponga primero el cliente/obra", vbExclamation
                    If txtAux(0).Text = "" Then
                        PonerFoco txtAux(0)
                    Else
                        PonerFoco txtAux(1)
                    End If
                 Else
                    devuelve = "codclien = " & txtAux(0).Text & " AND coddirec = " & txtAux(1).Text & " AND actuacion"
                    devuelve = DevuelveDesdeBD(conAri, "fechaini", "sactuaobra", devuelve, txtAux(3).Text, "T")
                    If devuelve = "" Then
                        devuelve = "No existe la actuacion: " & txtAux(3).Text & vbCrLf & vbCrLf
                        devuelve = devuelve & " Cliente: " & txtAux(0).Text & " " & txtAux(6).Text & vbCrLf
                        devuelve = devuelve & " Obra: " & txtAux(1).Text & " " & txtAux(2).Text & vbCrLf
                        MsgBox devuelve, vbExclamation
                        txtAux(3).Text = ""
                        PonerFoco txtAux(3)
                    Else
                        PonerFoco txtAux(4)
                    End If
                    devuelve = ""
                End If
            End If
        Case 5 'Cantidad (Comprobamos formato como si fuera un Importe)
            'Formato tipo 1: Decimal(12,2)
            If txtAux(Index) <> "" Then PonerFormatoDecimal txtAux(Index), 1
            
        Case 7
            devuelve = ""
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            If txtAux(Index).Text <> "" Then
                devuelve = DevuelveDesdeBD(conAri, "nomtipor", "stipor", "codtipor", txtAux(Index).Text, "T")
                If devuelve = "" Then MsgBox "No existe la orden de trabajo", vbExclamation
            End If
            If devuelve = "" Then
                If txtAux(Index).Text <> "" Then PonerFoco txtAux(Index)
                txtAux(Index).Text = ""
            End If
            txtAux(8).Text = devuelve
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
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
        Case 9 'Mantenimiento Lineas
            BotonLineas
            
        Case 12 'Imprimir
            frmObraListado.opcion = 0
            frmObraListado.Show vbModal
            
        Case 13  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim B As Boolean
Dim NumReg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo

    'Modo 2. Hay datos y estamos visualizandolos
    '-------------------------------------------
    B = (Kmodo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
              
    'Como el campo 0 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), (Modo <> 1), True
    
    'Modo 1:Busqueda / Modo 3: Insertar / Modo 4: Modificar
    '-------------------------------------------------------
    B = (Modo = 3 Or Modo = 4 Or Modo = 1)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = B
    Next I
    
    
    Me.imgBuscar(2).Enabled = B
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
Dim I As Byte
    


    
 
         B = (Modo = 2) Or (Modo = 5)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (B Or Modo = 0)
        Me.mnNuevo.Enabled = (B Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = B
        Me.mnModificar.Enabled = B
        'eliminar
        Toolbar1.Buttons(7).Enabled = B
        Me.mnEliminar.Enabled = B
        
        '--------------------------------
        B = (Modo = 2)
        'Lineas Traspaso Almacenes
        Toolbar1.Buttons(9).Enabled = B
        'Actualizar
        Toolbar1.Buttons(10).Enabled = B
        'Imprimir
        Toolbar1.Buttons(12).Enabled = B Or Modo = 0
            
        '-------------------------------
        B = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not B
        Me.mnBuscar.Enabled = Not B
        'VerTodos
        Toolbar1.Buttons(2).Enabled = Not B
        Me.mnVerTodos.Enabled = Not B
  
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox

    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones(Flechas) de Desplazamiento de Registros de la Toolbar

    Select Case Modo
        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
            If Data2.Recordset.EOF Then Exit Sub
            DesplazamientoData Data2, Index
        Case Else 'Datos de Cabecera
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, Index
            PonerCampos
    End Select
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
On Error GoTo EMontaSQL
 


    SQL = "SELECT  sliparte.linea, sliparte.codclien,  sliparte.coddirec,nomdirec,actuacion,"
    SQL = SQL & " sliparte.descr, sliparte.codtipor,nomtipor,"
    SQL = SQL & " if(partet=0,""No"",""Si""),if(partep=0,""No"",""Si""), sliparte.horas,nomclien,partet,partep FROM"
    SQL = SQL & " sliparte  left join stipor on sliparte.codtipor=stipor.codtipor "
    SQL = SQL & " ,sclien,sdirec  WHERE sliparte.codclien = sclien.codclien AND"
    SQL = SQL & " sliparte.codClien = sdirec.codClien And sliparte.CodDirec = sdirec.CodDirec AND "
    
    
    If enlaza Then
        SQL = SQL & ObtenerWhereCP(False)  '" WHERE codtrasp = " & Data1.Recordset!codtrasp
    Else
        SQL = SQL & " numparte = -1"
    End If
    SQL = SQL & " ORDER BY " & NomTablaLineas & ".linea"
    MontaSQLCarga = SQL
    
EMontaSQL:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
        PonerFoco Text1(0)
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonLineas()
On Error GoTo ErrorLineas

    Screen.MousePointer = vbHourglass
    PonerModo 5
    ModificaLineas = 0
    PonerBotonCabecera True
    'CargaGrid True
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lineas"
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonAnyadir()
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
        
    'Ponemos el grid lineas Traspaso enlazando a ningun sitio
    CargaGrid False
    
    'Sugerir codigo siguiente
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, "numparte")
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    'Poner Trabajador por defecto el trabajador conectado
    Text1(2).Text = PonerTrabajadorConectado(NomTraba)
    Text2(2).Text = NomTraba
    
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLineas()
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    
    ModificaLineas = 1
    
    vWhere = ObtenerWhereCP(False)
    
    
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data2

    CargaTxtAux True, True
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    'Como el campo 0 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), True, True
    PonerFoco Text1(1)
End Sub

Private Sub BotonModificarLinea()
Dim I As Integer

    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.RecordCount < 1 Then Exit Sub

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub
    
    ModificaLineas = 2 'Modificar

    Screen.MousePointer = vbHourglass
    
    PonerBotonCabecera False
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    


    CargaTxtAux True, False
    PonerFoco txtAux(4) 'Poner el foco
    Screen.MousePointer = vbDefault
    Me.DataGrid1.Enabled = False
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    SQL = "Parte trabajo." & vbCrLf
    SQL = SQL & "------------------------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a eliminar el parte:" & vbCrLf
    SQL = SQL & vbCrLf & "Numero   : " & Text1(0).Text
    SQL = SQL & vbCrLf & "Fecha   : " & CStr(Data1.Recordset.Fields(2))
    SQL = SQL & vbCrLf & "Trabajador: " & Text1(2).Text & "  " & Text2(2).Text
    SQL = SQL & vbCrLf & vbCrLf & " ¿Desea continuar ? "
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub

        NumRegElim = Data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
     
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
On Error GoTo FinEliminar
    
    conn.BeginTrans
    SQL = ObtenerWhereCP(True)  '" WHERE  codtrasp=" & Data1.Recordset!codtrasp
    
    'Lineas
    conn.Execute "Delete  from " & NomTablaLineas & SQL
    
    'Cabeceras
    conn.Execute "Delete  from " & NombreTabla & SQL
                      
                      

                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub BotonEliminarLinea()
Dim SQL As String
On Error GoTo Error2
    'Ciertas comprobaciones
    If Data2.Recordset.EOF Then Exit Sub

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    '### a mano
    SQL = "Seguro que desea eliminar la línea del parte:"
    SQL = SQL & vbCrLf & "Actuacion: " & Data2.Recordset!codclien & " " & Data2.Recordset!Nomclien & " / " & Data2.Recordset!nomdirec
    SQL = SQL & vbCrLf & "Descripción: " & Data2.Recordset.Fields(5)
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from sliparte where numparte=" & Data1.Recordset!numparte
        SQL = SQL & " and linea=" & Data2.Recordset!linea
        conn.Execute SQL
        CancelaADODC Me.Data2
        CargaGrid True
        PonerSumaHoras
        CancelaADODC Me.Data2
    End If
    ModificaLineas = 0
Error2:
        Screen.MousePointer = vbDefault
        ModificaLineas = 0
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Línea del partes de trabajo", Err.Description
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me, 1)

    If B Then
        'No se repite dia 7trabajador
        cadSeleccion = "fecha=" & DBSet(Text1(1).Text, "F") & " AND codtraba"
        cadSeleccion = DevuelveDesdeBD(conAri, "count(*)", "scaparte", cadSeleccion, Text1(2).Text, "N")
        If cadSeleccion <> "" Then
            If Val(cadSeleccion) > 0 Then
                MsgBox "Ya tiene datos el trabajador " & Me.Text2(2).Text & " para el dia " & Text1(1).Text, vbExclamation
                B = False
            End If
            cadSeleccion = ""
        End If
     End If


    DatosOk = B
End Function






Private Function DatosOkLinea() As Boolean
Dim B As Boolean

    DatosOkLinea = False
    B = True
        
        
    For NumRegElim = 0 To txtAux.Count - 1
        Select Case NumRegElim
        Case 6, 7, 8, 4
        
        Case Else
            If txtAux(NumRegElim).Text = "" Then
                MsgBox "Campos obligatorios", vbExclamation
                PonerFoco txtAux(NumRegElim)
                B = False
                Exit Function
            End If
        End Select
    Next
        
    'Comprobamos el campo Cantidad
    If Not IsNumeric(txtAux(5).Text) Then
        MsgBox "El campo horas debe ser numérico", vbExclamation
        B = False
    End If
    If Not B Then
        PonerFoco txtAux(5)
        Exit Function
    End If
    
    If Combo1.ListIndex < 0 Then
        MsgBox "Selecciona Parte Trabajo", vbExclamation
        PonerFocoCbo Combo1
        B = False
    End If
    If Combo2.ListIndex < 0 Then
        MsgBox "Selecciona Parte Privado", vbExclamation
        PonerFocoCbo Combo2
        B = False
    End If
        
    If txtAux(7).Text = "" Xor txtAux(8).Text = "" Then
        MsgBox "Error en orden de trabajo", vbExclamation
        If B Then PonerFoco txtAux(7)
        B = False
    End If

    DatosOkLinea = B
End Function


Private Sub PonerBotonCabecera(B As Boolean)
On Error Resume Next
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "LINEAS DETALLE"
    Else
        Me.lblIndicador.Caption = ""
    End If
     'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function InsertarModificarLinea() As Boolean
Dim SQL As String
On Error GoTo EInsertarModificarLinea
    
    SQL = ""
    InsertarModificarLinea = False
    Select Case ModificaLineas
    Case 1 'Insertar
        If DatosOkLinea() Then 'INSERTAR
            SQL = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
            SQL = SugerirCodigoSiguienteStr(NomTablaLineas, "linea", SQL)
            SQL = "INSERT INTO sliparte (numparte,linea,codclien ,coddirec ,actuacion ,descr ,partet ,partep,horas ,codtipor) VALUES (" & Val(Text1(0).Text) & ", " & SQL & ","
            SQL = SQL & DBSet(txtAux(0).Text, "N") & ", "
            SQL = SQL & DBSet(txtAux(1).Text, "N") & ","
            SQL = SQL & DBSet(txtAux(3).Text, "T") & ", "
            SQL = SQL & DBSet(txtAux(4).Text, "T") & ", "
            SQL = SQL & Combo1.ItemData(Combo1.ListIndex) & ","
            SQL = SQL & Combo2.ItemData(Combo2.ListIndex) & ","
            SQL = SQL & DBSet(txtAux(5).Text, "N") & ", "
            SQL = SQL & DBSet(txtAux(7).Text, "T", "S") & ") "
        Else
'            PonerFoco txtAux(3)
        End If
    Case 2 'Modificar
        If DatosOkLinea() Then
            SQL = "UPDATE sliparte Set horas = " & DBSet(txtAux(5).Text, "N")
            SQL = SQL & ", partet = " & Combo1.ItemData(Combo1.ListIndex)
            SQL = SQL & ", partep = " & Combo2.ItemData(Combo2.ListIndex)
            SQL = SQL & ", descr = " & DBSet(txtAux(4).Text, "T")
            SQL = SQL & ", actuacion = " & DBSet(txtAux(3).Text, "T")
            SQL = SQL & ", coddirec = " & DBSet(txtAux(1).Text, "N")
            SQL = SQL & ", codclien = " & DBSet(txtAux(0).Text, "N")
            SQL = SQL & ", codtipor = " & DBSet(txtAux(7).Text, "T", "S")
            SQL = SQL & ObtenerWhereCP(True) & " AND " '" WHERE codtrasp =" & Val(Text1(0).Text) & " AND "
            SQL = SQL & " linea =" & Data2.Recordset!linea
        End If
    End Select
            
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLinea = True
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Lineas Traspaso Almacenes" & vbCrLf & Err.Description
End Function


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    If Modo <> 5 Then 'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: scatra
        Cad = Cad & ParaGrid(Text1(0), 15, "Numero")
        Cad = Cad & ParaGrid(Text1(1), 18, "Fecha")
        Cad = Cad & ParaGrid(Text1(2), 13, "Cod.trab")
        Cad = Cad & "Trabajador|straba|nomtraba|T||30·"
        
        
        tabla = "(" & NombreTabla & " LEFT JOIN straba ON " & NombreTabla & ".codtraba=straba.codtraba" & ")"
        
        
        ' tabla = "scatra"
        Titulo = Me.Caption
    Else 'Estamos en modo Lineas
        Cad = Cad & "Código|sclien|nomclien|N||30·Nombre|sclien|nomartic|T||70·"
        tabla = "sartic"
        Titulo = "Articulos"
    End If
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda(Me, False)
    'cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
            PonerCadenaBusqueda
        Else
'            CadenaConsulta = "select * from " & NombreTabla & Ordenacion
'            PonerCadenaBusqueda
            MsgBox "Introducir criterios de búsqueda", vbExclamation
            PonerFoco Text1(0)
        End If
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
        Me.DataGrid1.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    If Err.Number <> 0 Then MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "straba", "nomtraba")
    CargaGrid True
    
    PonerSumaHoras
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Function ActualizarStocks() As Boolean
Dim SQL As String
Dim Cantidad As Single
Dim devuelve As String
Dim RS As ADODB.Recordset

    On Error GoTo EActualizarStock

    ActualizarStocks = False
    
    '---- Laura: 27/09/2006
    'sustituir el data2 por el RecordSEt
    Set RS = New ADODB.Recordset
    RS.Open Data2.RecordSource, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
'    While Not Data2.Recordset.EOF

        'Actualizar el stock si el articulo tiene control de stock
        devuelve = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", RS!codArtic, "T")
        If Val(devuelve) = 1 Then

            Cantidad = CSng(RS!Cantidad) 'Cant a traspasar
            
            '==== Almacen Origen
            SQL = "UPDATE salmac Set canstock = canstock - " & DBSet(Cantidad, "N")
            SQL = SQL & " WHERE codartic =" & DBSet(RS!codArtic, "T") & " AND "
            SQL = SQL & " codalmac =" & Data1.Recordset!almaorig
            conn.Execute SQL
        
            '==== Almacen Destino
            'Comprobar que existe el articulo en Almacen Destino
            devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codalmac", "codartic", RS!codArtic, "T", , "codalmac", Text1(3).Text, "N")
            If devuelve = "" Then 'No hay de ese artículo en Destino
                SQL = "INSERT INTO salmac (codartic,codalmac,ubialmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
                SQL = SQL & " VALUES (" & DBSet(RS!codArtic, "T") & "," & Val(Text1(3).Text) & ",''," & DBSet(Cantidad, "N") & ",0,0,0,0,NULL,NULL,0)"
            Else 'Existe el artic en almac. Dest -> Aumentar stock
                SQL = "UPDATE salmac Set canstock = canstock + " & DBSet(Cantidad, "N")
                SQL = SQL & " WHERE codartic =" & DBSet(RS!codArtic, "T") & " AND "
                SQL = SQL & " codalmac =" & Data1.Recordset!almadest
            End If
            
            conn.Execute SQL
        End If
        'Data2.Recordset.MoveNext
        RS.MoveNext
    Wend
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        ActualizarStocks = False
    Else
        ActualizarStocks = True
    End If
    
EActualizarStock:
End Function










Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function ObtenerWhereCP(conWhere As Boolean) As String
On Error Resume Next
    ObtenerWhereCP = ""
    If conWhere Then ObtenerWhereCP = " WHERE "
    ObtenerWhereCP = ObtenerWhereCP & " numparte= " & Val(Text1(0).Text)
    
End Function


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             PonerCampos
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    End If
End Sub



Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerSumaHoras()
Dim Horas As Currency
    If Data2.Recordset.EOF Then
        Horas = 0
    Else
        cadSeleccion = DevuelveDesdeBD(conAri, "sum(horas)", "sliparte", "numparte", Text1(0).Text)
        If cadSeleccion = "" Then cadSeleccion = "0"
        Horas = CCur(cadSeleccion)
    End If
    Me.THoras.Text = Format(Horas, FormatoCantidad)
End Sub
