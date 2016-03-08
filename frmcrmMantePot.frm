VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmcrmMantePot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Mantenimiento acciones comerciales sobre clientes potenciales"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13080
   Icon            =   "frmcrmMantePot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   9480
      MaxLength       =   30
      TabIndex        =   29
      Tag             =   "Tipo|N|N|||scrmaccionclipot|tipo|||"
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   9480
      MaxLength       =   30
      TabIndex        =   28
      Tag             =   "Tipo|N|N|||scrmaccionclipot|tipo|||"
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "no visible"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   11280
      TabIndex        =   25
      Top             =   1680
      Width           =   1815
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   26
         Tag             =   "medio|T|N|||scrmaccionclipot|medio|||"
         Text            =   "Text"
         Top             =   360
         Width           =   2805
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      ItemData        =   "frmcrmMantePot.frx":000C
      Left            =   5400
      List            =   "frmcrmMantePot.frx":000E
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   840
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      ItemData        =   "frmcrmMantePot.frx":0010
      Left            =   3720
      List            =   "frmcrmMantePot.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Estado|N|N|||scrmaccionclipot|estado|||"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6720
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   1605
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   1605
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   2475
      Index           =   6
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "O|T|S|||scrmaccionclipot|observaciones||N|"
      Text            =   "frmcrmMantePot.frx":0014
      Top             =   2040
      Width           =   9165
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Tag             =   "usuario|T|N|||scrmaccionclipot|usuario||S|"
      Text            =   "Text"
      Top             =   840
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   120
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Cliente|N|N|||scrmaccionclipot|codclipot||S|"
      Text            =   "Text1"
      Top             =   1605
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   5400
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Tipo|N|N|||scrmaccionclipot|tipo|||"
      Text            =   "Text1"
      Top             =   1605
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1380
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Fecha/Horal|H|N|||scrmaccionclipot|fechora|dd/mm/yyyy hh:mm:ss|S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   2535
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
         TabIndex        =   11
         Top             =   210
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   4800
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   4800
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   120
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13080
      _ExtentX        =   23072
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
            Style           =   3
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
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblQue 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   2880
      TabIndex        =   27
      Top             =   4680
      Width           =   5175
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   6240
      Picture         =   "frmcrmMantePot.frx":001A
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   1327
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   1080
      Picture         =   "frmcrmMantePot.frx":011C
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   1327
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Medio"
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   21
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Estado"
      Height          =   255
      Index           =   8
      Left            =   3720
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion"
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   19
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   18
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Cod.Cliente"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1350
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo accion"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha/hora"
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   13
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   975
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
Attribute VB_Name = "frmcrmMantePot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public TipoPredefinido As Byte   'Del 1 al 20 nos los hemos reservado

Public DesdeElCliente As Long  '0  -Ningun cliente

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmAcc As frmCRMtipos
Attribute frmAcc.VB_VarHelpID = -1
Private WithEvents frmCli As frmFacClientes3
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmAg As frmFacAgentesCom
Attribute frmAg.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1


'  Variables comunes a todos los formularios
Private Modo As Byte
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

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

'Private VieneDeBuscar As Boolean
''Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
''de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private PrimeraVez As Boolean

Private ModificandoEnCliente As Boolean

Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    If DesdeElCliente > 0 Then
                        Unload Me
                        Exit Sub
                    End If
                    PosicionarData
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
        Case 1  'BUSCAR
            HacerBusqueda
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1, 3 'Insertar
        LimpiarCampos
        PonerModo 0
        PonerOpcionesMenu
    Case 4  'Modificar
        lblIndicador.Caption = ""
        TerminaBloquear
        PonerModo 2
        PonerCampos
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3
    Text1(0).Text = vUsu.Login
    Text1(1).Text = Format(Now, "dd/mm/yyyy hh:mm:ss ")
    Me.Combo1(0).ListIndex = 0
    'Nuevo para un cliente
    If DesdeElCliente > 0 Then
        Text1(3).Text = DesdeElCliente
        Text1_LostFocus 3
        BloquearTxt Text1(3), True
    End If
    If TipoPredefinido > 0 Then
        Text1(2).Text = TipoPredefinido
        Text1_LostFocus 2
        BloquearTxt Text1(2), True
        
        
        'Si el tipo preddefinido es el UNO, pondresmo como medio Telefono
        If TipoPredefinido = 1 Then Me.Combo1(1).ListIndex = 0
        If TipoPredefinido = 2 Then Me.Combo1(1).ListIndex = 4
        
        lblQue.Caption = Text2(2).Text
    End If
    
    
   ' Text1(5).Text = PonerTrabajadorConectado(CadenaDesdeOtroForm)
   ' Text2(5).Text = CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    
    PonerFoco Text1(6)
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
Dim C As String
    C = ""
    If vUsu.Nivel > 0 Then C = " tipo >20"
    'Ver todos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia C
    Else
        LimpiarCampos
        If C <> "" Then C = " WHERE " & C
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
    If Me.TipoPredefinido > 0 Then BloquearTxt Text1(2), True
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    

    
    Cad = "¿Seguro que desea eliminar la accion comercial? " & vbCrLf
    Cad = Cad & vbCrLf & "Usuario: " & Format(Data1.Recordset.Fields(0), "0000")
    Cad = Cad & vbCrLf & "Cliente: " & Text1(2).Text & "     " & Text2(2).Text
    Cad = Cad & vbCrLf & "Fecha/Hora: " & Text1(1).Text
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub


        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Cad = DevuelveWHERE
        Cad = "Delete from scrmaccionclipot where " & Cad
        conn.Execute Cad
        

        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If

    Screen.MousePointer = vbDefault
    
Error2:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Agente Comercial", Err.Description
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

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    
    If PrimeraVez Then
        PrimeraVez = False
        If DesdeElCliente > 0 Then
            If DatosADevolverBusqueda = "" Then
                'Es para añadir nueva
                BotonAnyadir
            
            Else
                If Data1.Recordset.EOF Then
                    MsgBox "No se ha encontrado el registro", vbExclamation
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                    lblQue.Caption = Text2(2).Text
                    ModificandoEnCliente = True
                End If
            End If
        End If
        DatosADevolverBusqueda = ""
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim Vac As Boolean
    'Icono del formulario
    PrimeraVez = True
    ModificandoEnCliente = False
    Me.Icon = frmPpal.Icon
    Me.Width = 10800
    lblQue.Caption = ""
    ' ICONITOS DE LA BARRA
    btnPrimero = 13
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 15  'Salir
        .Buttons(13).Image = 6  'Primero
        .Buttons(14).Image = 7  'Anterior
        .Buttons(15).Image = 8  'Siguiente
        .Buttons(16).Image = 9  'Último
    End With
    
    CargaComboMediosCRM Me.Combo1(1)
    CargaComboEstadoCRM Me.Combo1(0)
    
    LimpiarCampos
    'VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "scrmaccionclipot"
    Ordenacion = " ORDER BY fechora ,codclien"
        
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario

    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Vac = True 'Que lo carge vacio
    If DesdeElCliente > 0 Then
        If DatosADevolverBusqueda <> "" Then Vac = False 'cargara datos en el activate
    End If
    
    
    If Vac Then
        'Modo normal o para modo insertar
        Data1.RecordSource = "Select * from " & NombreTabla & " where codclipot=-1"
    Else
        Data1.RecordSource = "Select * from " & NombreTabla & " where " & DatosADevolverBusqueda
        
    End If
    Data1.Refresh
    
    'MODO
    If DesdeElCliente > 0 Then
        'No pongo modo y en el activate hare cosas
    
    Else
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
            Text1(0).BackColor = vbYellow
        End If
    End If

End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmAcc_DatoSeleccionado(CadenaSeleccion As String)
    HaDevueltoDatos = True
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmAg_DatoSeleccionado(CadenaSeleccion As String)
    HaDevueltoDatos = True
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2)
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
        Aux = Aux & " AND " & ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        Aux = Aux & " AND " & ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 3)
        cadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub
    




Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    HaDevueltoDatos = True
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    HaDevueltoDatos = True
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    
    'Clientes
    If Index = 3 Then
        If Modo = 4 Then Exit Sub
        If Modo = 3 And Me.DesdeElCliente > 0 Then Exit Sub
    End If
    
    If Index = 2 And TipoPredefinido > 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    HaDevueltoDatos = False
    
    Select Case Index
    Case 2
        Set frmAcc = New frmCRMtipos
        frmAcc.DatosADevolverBusqueda = "0|1|"
        frmAcc.Show vbModal
        Set frmAcc = Nothing
        
    Case 3
        Set frmCli = New frmFacClientes3
        frmCli.DatosADevolverBusqueda = "0|1|"
        frmCli.Show vbModal
        Set frmCli = Nothing
        
    Case 4
        Set frmAg = New frmFacAgentesCom
        frmAg.DatosADevolverBusqueda = "0|1|"
        frmAg.Show vbModal
        Set frmAg = Nothing


        
    Case 5
        Set frmT = New frmAdmTrabajadores
        frmT.DatosADevolverBusqueda = "0|1|"
        frmT.Show vbModal
        Set frmT = Nothing
    
    End Select

    If HaDevueltoDatos Then PonerFoco Text1(Index)
    HaDevueltoDatos = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    If ModificandoEnCliente Then Exit Sub
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    If ModificandoEnCliente Then Exit Sub
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    If ModificandoEnCliente Then Exit Sub
    BotonVerTodos
End Sub


'### A mano
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
Dim devuelve As String
Dim Agente As String
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    Select Case Index
        Case 2, 3, 4, 5
            devuelve = ""
                    
            If PonerFormatoEntero(Text1(Index)) Then
                If Index = 3 Then
                    Agente = ""
                    devuelve = DevuelveDesdeBD(conAri, "nomclien", "sclipot", "codclien", Text1(Index).Text, "N", Agente)
                ElseIf Index = 4 Then
                  '      devuelve = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", Text1(Index).Text, "N")
                ElseIf Index = 5 Then
                  '      devuelve = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", Text1(Index).Text, "N")
                Else
                    'index =2
                    devuelve = DevuelveDesdeBD(conAri, "denominacion", "scrmtipo", "codigo", Text1(Index).Text, "N")
                End If
                If Index = 2 And devuelve <> "" Then
                    'SI NO tiene accion x defecto NO puede eligir menores <= 20
                    If TipoPredefinido > 0 Then
                        If TipoPredefinido <> Text1(Index).Text Then
                            MsgBox "Accion comercial debe ser " & TipoPredefinido, vbExclamation
                            Text1(Index).Text = TipoPredefinido
                            devuelve = ""
                        End If
                    Else
                        If Val(Text1(Index).Text) <= 20 Then
                            If vUsu.Nivel > 0 Then
                                MsgBox "Accion comercial debe ser mayor de 20", vbExclamation
                                Text1(Index).Text = ""
                                devuelve = ""
                            End If
                        End If
                    End If
                End If
            End If
            Text2(Index).Text = devuelve
            If Text1(Index).Text <> "" And devuelve = "" Then
                MsgBox "No existe el codigo", vbExclamation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            'If Index = 3 And Agente <> "" Then Text1(4).Text = Agente
        Case 1
            'FECHA HORA
            If Not IsDate(Text1(Index).Text) Then
                MsgBox "Fecha incorrecta", vbExclamation
                Text1(Index).Text = ""
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    If Me.Combo1(1).Text <> "" Then Text1(7).Text = Combo1(1).Text
    cadB = ObtenerBusqueda(Me, False)

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then     'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 10, "Usuario")
        Cad = Cad & ParaGrid(Text1(1), 25, "Fecha")
        Cad = Cad & ParaGrid(Text1(3), 10, "Codigo")
        Cad = Cad & "Nombre|sclien|nomclien|T||50·"

            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla & ",sclien"
            
            Cad = "sclien.codclien=" & NombreTabla & ".codclien"
            If cadB <> "" Then Cad = Cad & " AND " & cadB
            frmB.vSQL = Cad
            HaDevueltoDatos = False
            '###A mano

            frmB.vDevuelve = "0|1|2|" 'Campos de la tabla que devuelve
            frmB.vTitulo = "Acciones realizadas"
            frmB.vselElem = 1
            frmB.vConexionGrid = conAri 'Conexión a BD: Ariges
'            frmB.vBuscaPrevia = chkVistaPrevia
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                Modo = 3
                PonerCampos2
                Modo = 2
                PonerFocoBtn Me.cmdRegresar
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerModo Modo
                PonerFoco Text1(kCampo)
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
         'PonerModo 0
         Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
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

Private Sub PonerCampos2()
    Text1_LostFocus 3
  '  Text1_LostFocus 4
   ' Text1_LostFocus 5
    Text1_LostFocus 2
End Sub
Private Sub PonerCampos()
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    Combo1(1).Text = Text1(7).Text
    Modo = 3
    PonerCampos2
    Modo = 2
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim NumReg As Byte

    Modo = Kmodo

    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    
    BloquearText1 Me, Modo
    
    
    B = Modo = 2 Or Modo = 0
    BloquearCmb Me.Combo1(0), B, False
    BloquearCmb Me.Combo1(1), B, False
    
    '---------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = B Or Modo = 1
    cmdCancelar.visible = B Or Modo = 1
    
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub



Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
    
    B = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
    '----------------------------------------
    B = (Modo >= 3) 'Insertar/Modificar
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
'Dim cad As String

    

    Text1(7).Text = Combo1(1).Text
    DatosOk = False
    B = CompForm(Me, 1) 'Comprobar datos OK
    If Not B Then Exit Function
        
        
    If DevuelveDesdeBD(conConta, "codusu", "usuarios.usuarios", "login", Text1(0).Text, "T") = "" Then
        MsgBox "No existe el usuario: " & Text1(0).Text, vbExclamation
        Text1(0).Text = vUsu.Login
        Exit Function
    End If
        
        
    DatosOk = B
End Function

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


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


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String
Dim Situar As Boolean

    If Modo = 3 Then
        'Seleccion el registro
        Data1.RecordSource = "Select * from " & NombreTabla & " WHERE " & DevuelveWHERE
        Data1.Refresh
        Situar = True
    Else
        Situar = Not Data1.Recordset.EOF
    End If
    If Situar Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & DevuelveWHERE & ")"
         
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
'         If SituarDataGral(Data1, Text1(30).Text, "T", Text1(0).Text, "N", Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             PonerModo 0
         End If
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub



Private Function DevuelveWHERE() As String
    If Modo = 3 Then
        DevuelveWHERE = "codclipot = " & Text1(2).Text & " AND usuario=" & DBSet(Text1(0).Text, "T")
        DevuelveWHERE = DevuelveWHERE & " AND fechora = " & DBSet(Text1(1).Text, "FH")
    Else
        DevuelveWHERE = "codclipot = " & Data1.Recordset!codclipot & " AND usuario=" & DBSet(Data1.Recordset!Usuario, "T")
        DevuelveWHERE = DevuelveWHERE & " AND fechora = " & DBSet(Data1.Recordset!fechora, "FH")
    End If
End Function


