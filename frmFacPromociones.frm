VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacPromociones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Promociones Tarifas"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7800
   ClipControls    =   0   'False
   Icon            =   "frmFacPromociones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameNuevos 
      Caption         =   "Valores Nuevos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   2175
      Left            =   4050
      TabIndex        =   27
      Top             =   1800
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   1480
         MaxLength       =   12
         TabIndex        =   8
         Tag             =   "Precio Nuevo|N|S|0|999999.0000|spromo|precionu|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   1160
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   1480
         MaxLength       =   12
         TabIndex        =   9
         Tag             =   "Precio Caja Nuevo|N|S|0|999999.0000|spromo|precion1|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1480
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Fecha Inicio Nueva|F|S|||spromo|fechain1|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   1480
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Fecha Fin Nueva|F|S|||spromo|fechafi1|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   760
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Precio"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1160
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Precio Caja"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   760
         Width           =   735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1190
         Picture         =   "frmFacPromociones.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1200
         Picture         =   "frmFacPromociones.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame FrameActuales 
      Caption         =   "Valores Actuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   2175
      Left            =   480
      TabIndex        =   23
      Top             =   1800
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1480
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Fin Actual|F|N|||spromo|fechafin|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   760
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1480
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Inicio Actual|F|N|||spromo|fechaini|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1480
         MaxLength       =   12
         TabIndex        =   5
         Tag             =   "Precio Caja Actual|N|S|0|999999.0000|spromo|precioa1|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   1480
         MaxLength       =   12
         TabIndex        =   4
         Tag             =   "Precio Actual|N|N|0|999999.0000|spromo|precioac|###,##0.0000|N|"
         Text            =   "Text1"
         Top             =   1160
         Width           =   1335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1190
         Picture         =   "frmFacPromociones.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   760
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1190
         Picture         =   "frmFacPromociones.frx":01AD
         ToolTipText     =   "Buscar fecha"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Fin"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   760
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Caja"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Precio"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1160
         Width           =   495
      End
   End
   Begin VB.CheckBox chkPermiteDto 
      Caption         =   "Permite Descuento"
      Enabled         =   0   'False
      Height          =   195
      Left            =   5520
      TabIndex        =   22
      Tag             =   "Permite Descuento|N|N|||spromo|dtopermi||N|"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6075
      TabIndex        =   11
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6075
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   480
      TabIndex        =   20
      Top             =   4230
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   2350
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   1110
      Width           =   2445
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   735
      Width           =   3885
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Cod. Artículo|T|N|||spromo|codartic||S|"
      Text            =   "Text1"
      Top             =   735
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "Cod. Tarifa|N|N|0|999|spromo|codlista|000|S|"
      Text            =   "Text1"
      Top             =   1110
      Width           =   510
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
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
            Object.ToolTipText     =   "Preparar precios promocion"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar precios promocion"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mailing"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6120
         TabIndex        =   19
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3360
      Top             =   4440
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
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1500
      ToolTipText     =   "Buscar artículo"
      Top             =   765
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1500
      ToolTipText     =   "Buscar tarifa"
      Top             =   1125
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "Cod. Artículo"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   735
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Cod. Tarifa"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1110
      Width           =   975
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
      TabIndex        =   14
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
Attribute VB_Name = "frmFacPromociones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmT As frmFacTarifas 'Form Mantenimiento Tarifas
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmA As frmAlmArticu2  'Form Mantenimiento Articulos
Attribute frmA.VB_VarHelpID = -1

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Dim CadenaConsulta As String

Private HaDevueltoDatos As Boolean


Private Sub chkPermiteDto_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkPermiteDto_LostFocus()
    PonerFoco Text1(2)
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim SQL As String
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1
        HacerBusqueda
    Case 3 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then PosicionarData
        End If
        
    Case 4 'MODIFICAR
           If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
                'PonerOpcionesMenu
            Else
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

'Private Sub cmdRegresar_Click()
''Este es el boton Cabecera
'Dim cad As String
'Dim Indicador As String
'
'    'Quitar lineas y volver a la cabecera
'    If Modo = 5 Then 'modo 5: Lineas Articulos x Almacen
'        DataGrid1.ClearFields
'        cad = "(codmovim=" & Val(Text1(0).Text) & ")"
'        If SituarData(Data1, cad, Indicador) Then
'            PonerModo 2
'            lblIndicador.Caption = Indicador
'            Me.Toolbar1.Buttons(9).Enabled = True
'            Me.Toolbar1.Buttons(10).Enabled = True
'        End If
'    End If
'End Sub



Private Sub Form_activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Me.imgBuscar(1).Picture = frmPpal.imgListComun.ListImages(19).Picture
    

    'ICONOS de La toolbar
    btnPrimero = 16 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(9).Image = 27 'Para cambiar precios
        .Buttons(10).Image = 21 'Para cambiar precios
        .Buttons(11).Image = 20 'Para precios revita-email
        .Buttons(12).Image = 16 'Imprimir
        .Buttons(14).Image = 15 'Salir
        .Buttons(16).Image = 6 'Primero
        .Buttons(17).Image = 7 'Anterior
        .Buttons(18).Image = 8 'Siguiente
        .Buttons(19).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox

    'Pone el Tag del primer botón de busqueda a -1
    'Si tag =-1 abre busqueda en la tabla del mantenimiento, BD: Ariges
    'Si tag>0 abre busqueda en la tabla asociada al indice.
    '### se puede borrar??
'    imgBuscar(0).Tag = "-1"
    '###
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "spromo" 'Tabla Promociones Tarifas
    Ordenacion = " ORDER BY codartic, codlista"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codartic = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Articulos
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
Dim Indice As Integer
      
    If CadenaDevuelta <> "" Then
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            
            'Estamos en Cabecera
            'Recupera todo el registro de Tarifas de Precios
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            CadB = CadB & " and " & Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
Dim Indice As Byte
    Indice = Val(Me.imgFecha(0).Tag)
    Select Case Indice
        Case 0, 1: Indice = Indice + 2
        Case 2, 3: Indice = Indice + 4
    End Select
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    'Formulario Mantenimiento Tarifas
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    
    Select Case index
        Case 0 'Codigo Articulo
            Set frmA = New frmAlmArticu2
            'frmA.DatosADevolverBusqueda3 = "@1@" 'Abre en modo busqueda
            frmA.DesdeTPV = False
            frmA.Show vbModal
            Set frmA = Nothing
        Case 1  'Cod. Tarifa
            Set frmT = New frmFacTarifas
            frmT.DatosADevolverBusqueda = "0"
            frmT.Show vbModal
            Set frmT = Nothing
    End Select
    PonerFoco Text1(index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(index As Integer)
Dim Indice As Integer

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   imgFecha(0).Tag = index
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case index
    Case 0, 1: Indice = index + 2
    Case 2, 3: Indice = index + 4
   End Select

   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.Fecha = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
     If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
     Unload Me
End Sub

Private Sub Text1_GotFocus(index As Integer)
    kCampo = index
    ConseguirFoco Text1(index), Modo
End Sub


Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(index As Integer)
Dim campo As String
Dim tabla As String

    If Not PerderFocoGnral(Text1(index), Modo) Then Exit Sub

    Select Case index
        Case 0 'Codigo Articulo
            campo = "nomartic"
            tabla = "sartic"
            Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, tabla, campo)
        Case 1 'Codigo Tarifa
            If PonerFormatoEntero(Text1(index)) Then
                Text2(index).Text = PonerNombreDeCod(Text1(index), conAri, "starif", "nomlista")
            Else
                Text2(index).Text = ""
            End If
        
        Case 2, 3, 6, 7 'Fechas Inicio/Fin Actuales y Nuevas
            If Text1(index).Text <> "" Then PonerFormatoFecha Text1(index)
        Case 4, 5, 8, 9 'Precios Actuales y Nuevos
            'Formato tipo 2: Decimal(10,4)
            If Text1(index).Text <> "" And Modo <> 1 Then PonerFormatoDecimal Text1(index), 2
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            BotonVerTodos
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
            
        Case 9, 10
            frmListado2.Opcion = 29 + Button.index  '38,39
            frmListado2.Show vbModal
            BotonBuscar
        Case 11
            If Modo = 2 Or Modo = 0 Then
                frmMailPromo.Show vbModal
                If Modo = 2 Then PosicionarData
            End If
            
        Case 12 'Imprimir
            AbrirListado (29) '29: Informe Promociones de Articulos
        Case 14  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim Cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, Cerrar
    If Cerrar Then Unload Me
    If KeyAscii = 27 And Modo = 1 Then cmdCancelar_Click 'busqueda
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
   
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
        
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
    Else
        cmdRegresar.visible = False
    End If
        
    '-------------------------------------------
    'Bloquear Registros
    BloquearText1 Me, Modo
            
    'Modo INSERTAR
    If Modo = 3 Then Me.chkPermiteDto.Value = 1
    
    '------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b And Modo <> 4 'Si modificar no activado pq son claves ajenas
    Next I
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b 'Si es insertar o modificar
    Next I
    
    Me.chkPermiteDto.Enabled = (Modo = 3) Or (Modo = 4)
    
    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    '-------------------------------------
    b = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub

Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkPermiteDto.Value = 0
End Sub



Private Sub Desplazamiento(index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, index
    PonerCampos
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
    
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
'Ver todos
    LimpiarCampos
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
'    cmdAceptar.Caption = "Aceptar"
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    
    'Para que si no se ha cargado el Data1 inicialmente, tenga valor cuando situamos el Data
'    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
'    Data1.RecordSource = CadenaConsulta
    
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    PonerFoco Text1(2)
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    SQL = "Promociones." & vbCrLf
    SQL = SQL & "-------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a Eliminar la Promoción para el Articulo:"
    SQL = SQL & vbCrLf & "Cod. Artic. : " & Text1(0).Text
    SQL = SQL & vbCrLf & "Cod. Tarifa : " & Text1(1).Text
    
    
    SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            'MsgBox Err.Number & ": " & Err.Description, vbExclamation
            MuestraError Err.Number, "Eliminar Promoción", Err.Description
            Data1.Recordset.CancelUpdate
        End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String

    On Error GoTo FinEliminar
        
'        Conn.BeginTrans
    SQL = " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T")
    SQL = SQL & " AND codlista = " & Val(Data1.Recordset!codlista)
       
    'Cabeceras
    conn.Execute "Delete  from " & NombreTabla & SQL
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
 '       Conn.RollbackTrans
        Eliminar = False
    Else
  '      Conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cadMen As String

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    'Comprobar que fecha fin promocion actual es igual o posterior a fecha inicio
    If Not EsFechaIgualPosterior(Text1(2).Text, Text1(3).Text, True) Then Exit Function
    
    'Comprobar que fecha fin promocion nueva es igual o posterior a fecha inicio
    'si existe valor para estos campos
    If Text1(6).Text <> "" And Text1(7).Text <> "" Then
        If Not EsFechaIgualPosterior(Text1(6).Text, Text1(7).Text, True) Then Exit Function
    End If
    
    'Para que no se solapen promociones
    'Comprobar que la Fecha Inicio de la Promocion Nueva es posterior a la
    'Fecha Fin de la promocion Actual
    cadMen = "No se pueden solapar las fechas de las promociones." & vbCrLf
    cadMen = cadMen & "La Fecha Inicio de la nueva promoción debe ser posterior a la Fecha Fin de la actual promoción."
    If Not EsFechaPosterior(Text1(3).Text, Text1(6).Text, True, cadMen) Then Exit Function
        
    DatosOk = True
End Function



Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    cad = cad & ParaGrid(Text1(0), 14, "Articulo")
    cad = cad & "Desc. Artic|sartic|nomartic|T||50·"
    cad = cad & ParaGrid(Text1(1), 8, "Tarifa")
    cad = cad & "Desc. Tarifa|starif|nomlista|T||27·"
    
    tabla = "(" & NombreTabla & " LEFT JOIN sartic ON " & NombreTabla & ".codartic=sartic.codartic" & ")"
    tabla = tabla & " LEFT JOIN starif ON " & NombreTabla & ".codlista=starif.codlista"
    
    Titulo = "Precios Especiales"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|2|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then
            MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla & ".", vbInformation
        End If
        Screen.MousePointer = vbDefault
        PonerModo Modo
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
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
    
    'Poner el nombre del cod. Articulo
    Text2(0).Text = PonerNombreDeCod(Text1(0), 1, "sartic", "nomartic")
    'Poner el nombre del cod. Tarifa
    Text2(1).Text = PonerNombreDeCod(Text1(1), 1, "starif", "nomlista")
    
    'Si los campos de precios nuevos son cero mostrar cadena vacia
    If Text1(8).Text <> "" Then
        If Text1(8).Text = 0 Then Text1(8).Text = ""
    End If
    If Text1(9).Text <> "" Then
        If Text1(9).Text = 0 Then Text1(9).Text = ""
    End If
    
'    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = "(codartic=" & DBSet(Text1(0).Text, "T") & " AND codlista=" & Text1(1).Text & ")"
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
'    If SituarData2(Data1, Text1(0).Text, Text1(1).Text, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'        LimpiarCampos
        PonerModo 0
    End If
End Sub
