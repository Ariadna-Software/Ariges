VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTaxcoTarifa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarifas taxímetro"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   16665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTaxcoTarif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   16665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3825
      TabIndex        =   71
      Top             =   45
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   72
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "ï¿½ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   69
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   180
         TabIndex        =   70
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
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      Height          =   195
      Left            =   14895
      TabIndex        =   68
      Top             =   360
      Width           =   1530
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   46
      Left            =   14880
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   5415
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   45
      Left            =   14880
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   4725
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   44
      Left            =   14880
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   4035
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   43
      Left            =   14880
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3345
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   42
      Left            =   14880
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2655
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   41
      Left            =   13200
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   5415
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   40
      Left            =   13200
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   4725
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   39
      Left            =   13200
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   4035
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   38
      Left            =   13200
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3345
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   37
      Left            =   13200
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2655
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   36
      Left            =   11520
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   5415
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   35
      Left            =   11520
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   4725
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   34
      Left            =   11520
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   4035
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   33
      Left            =   11520
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3345
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   32
      Left            =   11520
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2655
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   31
      Left            =   9840
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   5415
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   30
      Left            =   9840
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   4725
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   29
      Left            =   9840
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   4035
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   28
      Left            =   9840
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3345
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   27
      Left            =   9840
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2655
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   26
      Left            =   8160
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   5415
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   25
      Left            =   8160
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   4725
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   24
      Left            =   8160
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   4035
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   23
      Left            =   8160
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3345
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   22
      Left            =   8160
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2655
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   21
      Left            =   6480
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   5415
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   20
      Left            =   6480
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   4725
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   19
      Left            =   6480
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   4035
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   18
      Left            =   6480
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3345
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   17
      Left            =   6480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2655
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   16
      Left            =   4800
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   5415
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   15
      Left            =   4800
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   4725
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   14
      Left            =   4800
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   4035
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   13
      Left            =   4800
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3345
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   12
      Left            =   4800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2655
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   11
      Left            =   3120
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   5415
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   10
      Left            =   3120
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   4725
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   9
      Left            =   3120
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4035
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   8
      Left            =   3120
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3345
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   7
      Left            =   3120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2655
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   6
      Left            =   1440
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   5415
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   5
      Left            =   1440
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   4725
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   4
      Left            =   1440
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   4035
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   3
      Left            =   1440
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3345
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2655
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   0
      Left            =   150
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "Código|N|N|0|999|slista_taxi|codtarifa|000|S|"
      Text            =   "Text1"
      Top             =   1365
      Width           =   690
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   15420
      TabIndex        =   49
      Top             =   7035
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   1
      Left            =   1455
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||slista_taxi|descripcion|||"
      Text            =   "Text1"
      Top             =   1365
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   90
      TabIndex        =   50
      Top             =   6750
      Width           =   3135
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   210
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   15420
      TabIndex        =   48
      Top             =   7020
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   14160
      TabIndex        =   47
      Top             =   7035
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   90
      Top             =   6030
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
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      X1              =   90
      X2              =   16530
      Y1              =   6075
      Y2              =   6075
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      X1              =   120
      X2              =   16560
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Label Label2 
      Caption         =   "Percep. min."
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   67
      Top             =   5415
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Precio Km"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   66
      Top             =   4725
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Cada salto"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   65
      Top             =   4035
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Tar. horaria"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   64
      Top             =   3345
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Baj. bandera"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   63
      Top             =   2655
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   8
      Left            =   14880
      TabIndex        =   62
      Top             =   2085
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   240
      Index           =   7
      Left            =   13200
      TabIndex        =   61
      Top             =   2085
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   6
      Left            =   11520
      TabIndex        =   60
      Top             =   2085
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   240
      Index           =   5
      Left            =   9840
      TabIndex        =   59
      Top             =   2085
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   4
      Left            =   8160
      TabIndex        =   58
      Top             =   2085
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   240
      Index           =   3
      Left            =   6480
      TabIndex        =   57
      Top             =   2085
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   56
      Top             =   2085
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   240
      Index           =   1
      Left            =   3120
      TabIndex        =   55
      Top             =   2085
      Width           =   630
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   54
      Top             =   2085
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Denominación"
      Height          =   255
      Index           =   1
      Left            =   1455
      TabIndex        =   53
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   52
      Top             =   1080
      Width           =   720
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
Attribute VB_Name = "frmTaxcoTarifa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBasico2
Attribute frmB.VB_VarHelpID = -1

'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSCAR
            HacerBusqueda
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                   
                    PonerModo 0
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                   
                    TerminaBloquear
                    cad = "(codtarifa=" & Text1(0).Text & ")"
                    If SituarData(Data1, cad, Indicador) Then
                        PonerModo 2
                        lblIndicador.Caption = Indicador
                        PonerFoco Text1(0)
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3
            LimpiarCampos
            PonerModo 0
        Case 4
            'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr("slista_taxi", "codtarifa")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos
        PonerModo 1
        '### A mano
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


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index, True
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    '### A mano

    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    If Not PuedeEliminar Then Exit Sub
    
    '### a mano
    cad = "¿Seguro que desea eliminar la tarifa?" & vbCrLf
    cad = cad & vbCrLf & "Codigo: " & Format(Data1.Recordset.Fields(0), "000")
    cad = cad & vbCrLf & "Descripcion: " & Data1.Recordset.Fields(1)
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        Screen.MousePointer = vbHourglass

        
        Data1.Recordset.Delete
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            PonerModo 0
        End If
        
        
    End If
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Forma de Pago" & vbCrLf & cad, Err.Description
    End If
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim J As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
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
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    LimpiarCampos

           
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    'Los label de la tarifa
    For kCampo = 0 To 8
        Label2(kCampo).Caption = "Tarifa " & kCampo + 1
    Next
    
    NumRegElim = 2
    For kCampo = 1 To 9
        For J = 1 To 5
        
            Ordenacion = "tarifa" & J & kCampo
            Text1(NumRegElim).Tag = Ordenacion & "|T|S|||slista_taxi|" & Ordenacion & "|||"
            Text1(NumRegElim).Text = "" ' Text1(NumRegElim).Tag
            NumRegElim = NumRegElim + 1
        Next
    Next
    'ASignamos un SQL al DATA1
    NombreTabla = "slista_taxi"
    Ordenacion = " ORDER BY codtarifa"

    Data1.ConnectionString = conn
    '## A mano
    Data1.RecordSource = "Select * from " & NombreTabla & " where false"
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano

End Sub





Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
Dim Aux As String
Dim Indice As Integer

    If CadenaSeleccion <> "" Then
        
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
            cadB = Aux
            '   Como la clave principal es unica, con poner el sql apuntando
            '   al valor devuelto sobre la clave ppal es suficiente
            'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            'If CadB <> "" Then CadB = CadB & " AND "
            'CadB = CadB & Aux
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault

    End If

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
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
   
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 '
           If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
           End If
        
      
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
Dim cad As String

'        'Llamamos a al form
'        '##A mano
'        cad = ""
'        cad = cad & ParaGrid(Text1(0), 30, "Código")
'        cad = cad & ParaGrid(Text1(1), 70, "Denominacion")
'        If cad <> "" Then
'            Screen.MousePointer = vbHourglass
'            Set frmB = New frmBuscaGrid
'            frmB.vCampos = cad
'            frmB.vTabla = NombreTabla
'            frmB.vSQL = cadB
'
'            HaDevueltoDatos = False
'            frmB.vDevuelve = "0|1|" 'Campos de la tabla que devuelve
'            frmB.vTitulo = "Tarifas taxímetro"
'            frmB.vselElem = 1
'            frmB.vConexionGrid = 1 'Conexión a BD: Ariges
''            If imgFPago(0).Tag = -1 Then
''                frmB.vBuscaPrevia = chkVistaPrevia
''            Else
''                frmB.vBuscaPrevia = True
''            End If
'            frmB.vCargaFrame = False
'            frmB.Show vbModal
'            Set frmB = Nothing
'            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'            'tendremos que cerrar el form lanzando el evento
'''            If HaDevueltoDatos Then
''                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                    cmdRegresar_Click
''            Else   'de ha devuelto datos, es decir NO ha devuelto datos
''                PonerFoco Text1(kCampo)
''                PonerModo Modo
''            End If
'        End If

    Set frmB = New frmBasico2
    AyudaTarifasTaxi frmB, Text1(0), cadB
    Set frmB = Nothing


End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    
    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
'         MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
         Screen.MousePointer = vbDefault
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
    PonerCamposForma Me, Me.Data1
   
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    
    '----------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = b Or Modo = 1
    cmdCancelar.visible = b Or Modo = 1
    If b Or Modo = 1 Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    
    BloquearText1 Me, Modo
    

  
    

     chkVistaPrevia.Enabled = (Modo <= 2)

    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    mnEliminar.Enabled = b
    
    b = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = False
    
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    
    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
     
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then b = False
    End If
     
    If Not b Then Exit Function
        

    
    DatosOk = b
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5 'Buscar
            mnBuscar_Click
        Case 6  'Todos
            mnVerTodos_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Function PuedeEliminar() As Boolean
Dim cad As String
    PuedeEliminar = False
    
          
    cad = DevuelveDesdeBD(conAri, "count(*)", "sclien_taxi", "codtarifa", Text1(0).Text)
    If Val(cad) > 0 Then
        MsgBox "Existen clientes asociados a la tarifa", vbExclamation
    Else
        PuedeEliminar = True
    End If
    
End Function


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
