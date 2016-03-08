VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacClientesV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes Varios"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmFacClientesV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame FrameManipuladorFito 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   240
      TabIndex        =   26
      Top             =   3840
      Width           =   6375
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "F. caducidad|F|S|||sclvar|fcaducidad|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   1365
      End
      Begin VB.ComboBox cboFitos 
         Height          =   315
         ItemData        =   "frmFacClientesV.frx":000C
         Left            =   1200
         List            =   "frmFacClientesV.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "Tipo|N|S|||sclvar|ManipuladortipoCarnet|||"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "Num. carnet|T|S|||sclvar|ManipuladorNumCarnet|||"
         Text            =   "Text1"
         Top             =   120
         Width           =   2325
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Left            =   960
         Picture         =   "frmFacClientesV.frx":003E
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Caducidad"
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   28
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Carnet"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   27
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   7
      Left            =   1440
      MaxLength       =   200
      TabIndex        =   7
      Tag             =   "Teléfono|T|S|||sclvar|observa||N|"
      Text            =   "Text1"
      Top             =   3120
      Width           =   4965
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   6
      Tag             =   "Teléfono|T|S|||sclvar|telclien||N|"
      Text            =   "Text1"
      Top             =   2640
      Width           =   1600
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Provincia|T|S|||sclvar|proclien||N|"
      Text            =   "Text1"
      Top             =   2640
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "N.I.F.|T|N|||sclvar|nifclien||S|"
      Text            =   "Text1"
      Top             =   720
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   3360
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Población|T|S|||sclvar|pobclien||N|"
      Text            =   "Text1"
      Top             =   2160
      Width           =   2925
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "C. Postal|T|S|||sclvar|codpobla||N|"
      Text            =   "Text1"
      Top             =   2160
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Domicilio|T|N|||sclvar|domclien||N|"
      Text            =   "Text1"
      Top             =   1680
      Width           =   4845
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre Cliente Varios|T|N|||sclvar|nomclien||N|"
      Text            =   "Text1"
      Top             =   1200
      Width           =   4845
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   210
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   5160
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   5160
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   5520
      Top             =   720
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
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
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5160
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Observa."
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   25
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Left            =   1000
      Picture         =   "frmFacClientesV.frx":00C9
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   2190
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Teléfono"
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   24
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   23
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
      Height          =   255
      Index           =   4
      Left            =   2535
      TabIndex        =   22
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "C.Postal"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   21
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "N.I.F."
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   495
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
Attribute VB_Name = "frmFacClientesV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public vNif As String   'Por si he pinchado ese nif

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin ningun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private btnPrimero As Byte
'Numero del boton donde empiezan las flecha de desplazamiento de registros

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


Dim PrimeraVez As Boolean





Private Sub cboFitos_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Cad As String
Dim Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSCAR
            HacerBusqueda
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    Espera 0.25
                    Data1.RecordSource = "Select * from " & NombreTabla
                    Cad = "(nifclien=" & DBSet(Text1(0).Text, "T") & ")"
                    If SituarData(Data1, Cad, Indicador) Then
                        PonerModo 2
                        lblIndicador.Caption = Indicador
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    
                    Cad = "(nifclien=" & DBSet(Text1(0).Text, "T") & ")"
                    If SituarData(Data1, Cad, Indicador) Then
                        PonerModo 2
                        lblIndicador.Caption = Indicador
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
        Case 1, 3 '1:Buscar / 3:Insertar
            LimpiarCampos
            PonerModo 0
        Case 4  'Modificar
            lblIndicador.Caption = ""
            TerminaBloquear
            PonerModo 2
            PonerCampos
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    If vParamAplic.ManipuladorFitosanitarios2 Then cboFitos.ListIndex = 0
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then 'Modo 1: Buscar
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


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Cad = "¿Seguro que desea eliminar el Cliente de Varios?"
    Cad = Cad & vbCrLf & "Código: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Descripción: " & Data1.Recordset.Fields(1)
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
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
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Clientes Varios", Err.Description
    End If
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = Data1.Recordset.Fields(0) & "|"
    Cad = Cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Not Data1.Recordset.EOF Then
            If Modo = 2 Then PonerCampos
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim AbreModo1 As Boolean
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    PrimeraVez = True

    ' ICONITOS DE LA BARRA
    btnPrimero = 13 'Boton donde empiezan las Flechas de desplazamiento de Registros
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(9).Image = 16  'Imp
        .Buttons(10).Image = 15  'Salir
        .Buttons(13).Image = 6  'Primero
        .Buttons(14).Image = 7  'Anterior
        .Buttons(15).Image = 8  'Siguiente
        .Buttons(16).Image = 9  'Último
    End With
    
    LimpiarCampos
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "sclvar"
    Ordenacion = " ORDER BY nifclien"
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    FrameManipuladorFito.visible = vParamAplic.ManipuladorFitosanitarios2
    If vParamAplic.ManipuladorFitosanitarios2 Then
        Me.Height = 6420
        Frame1.Top = 5040
        Me.cmdAceptar.Top = 5160
    Else
        Me.Height = 5205
        Frame1.Top = 3960
        Me.cmdAceptar.Top = 3960
    End If
    Me.cmdCancelar.Top = cmdAceptar.Top
    Me.cmdRegresar.Top = cmdAceptar.Top
    
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    If vNif = "" Then
        Data1.RecordSource = "Select * from " & NombreTabla & " where nifclien=-1"
    Else
        Data1.RecordSource = "Select * from " & NombreTabla & " where nifclien= " & DBSet(vNif, "T")
    End If
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        If vNif <> "" And Data1.Recordset.RecordCount > 0 Then
            
            PonerModo 2
            
        Else
            '        BotonBuscar
            PonerModo 1
            Text1(0).BackColor = vbYellow
            PonerFoco Text1(0)
        End If
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    If vParamAplic.ManipuladorFitosanitarios2 Then Me.cboFitos.ListIndex = -1
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    vNif = ""
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim devuelve As String

    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(4).Text = ObtenerPoblacion(Text1(3).Text, devuelve)  'Poblacion
    'provincia
    Text1(5).Text = devuelve
End Sub


Private Sub frmF_Selec(vFecha As Date)
    Text1(9).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgBuscar_Click()
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    'Codigo Postal
    Set frmCP = New frmCPostal
    frmCP.DatosADevolverBusqueda = "0"
    frmCP.Show vbModal
    Set frmCP = Nothing

    PonerFoco Text1(3)
    VieneDeBuscar = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click()
    If Me.Text1(9).Locked Then Exit Sub
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Text1(9).Text <> "" Then frmF.Fecha = CDate(Text1(9).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    
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

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


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
Dim devuelve As String
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'NIF
            If Modo = 3 Then 'Insertar (Solo en modo=3.Es Clave primaria y no se Modifica)
                Text1(Index).Text = UCase(Text1(Index).Text)
                ValidarNIF Text1(Index).Text
                If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
            End If
             
        Case 3 'CPostal
            If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
                Text1(Index + 2).Text = ""
            ElseIf Not VieneDeBuscar Then
                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                 Text1(Index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
        Case 9
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
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
Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 30, "N.I.F.")
        Cad = Cad & ParaGrid(Text1(1), 70, "Nombre")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = cadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|" 'Campos de la tabla que devuelve
            frmB.vTitulo = "Clientes Varios "
            frmB.vselElem = 1
            frmB.vConexionGrid = 1 'Conexión a BD: Ariges
'            frmB.vBuscaPrevia = chkVistaPrevia
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(kCampo)
            End If
        End If
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
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '--------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    
    '-----------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) Or Modo = 1 '-->Luego not b sera kmodo<3
    cmdAceptar.visible = b
    cmdCancelar.visible = b
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    If vParamAplic.ManipuladorFitosanitarios2 Then BloquearCmb cboFitos, Not b
        
    Me.chkVistaPrevia.Enabled = (Modo <= 2)

    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    mnEliminar.Enabled = b
    'imprimir
    Toolbar1.Buttons(9).Enabled = b Or Modo = 0
    
    '---------------------------------------------
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


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar datos OK
    If Not b Then Exit Function
        
    'comprobamos si ya existe el cliente de varios
    If Modo = 3 Then 'Insertar
            If ExisteCP(Text1(0)) Then b = False
    End If

    If vParamAplic.ManipuladorFitosanitarios2 Then
        If cboFitos.ListIndex > 0 Then
            If Trim(Text1(8).Text) = "" Or Trim(Text1(9).Text) = "" Then
                MsgBox "Ha indicado que tiene carnet de fitosanitarios." & vbCrLf & "Debe indicar numero y fecha de caducidad", vbExclamation
                b = False
            End If
        End If
    End If

    DatosOk = b
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
        Case 9
            frmListado3.Opcion = 15
            frmListado3.Show vbModal
            
        Case 10  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
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
