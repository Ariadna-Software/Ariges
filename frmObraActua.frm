VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmObraActua 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actuaciones en obra"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7800
   ClipControls    =   0   'False
   Icon            =   "frmObraActua.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   1200
      Width           =   4605
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   720
      Width           =   4605
   End
   Begin VB.TextBox Text1 
      Height          =   1875
      Index           =   5
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "Ob|T|S|||sactuaobra|observa|||"
      Text            =   "frmObraActua.frx":000C
      Top             =   2640
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   4200
      TabIndex        =   4
      Tag             =   "F.Fin|F|S|||sactuaobra|fechafin|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Tag             =   "F.ini|F|S|||sactuaobra|fechaini|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   2
      Tag             =   "Actuacion|T|N|||sactuaobra|actuacion||S|"
      Text            =   "Text1"
      Top             =   1689
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Tag             =   "Cod. Obra|N|N|||sactuaobra|coddirec|0000|S|"
      Text            =   "Text1"
      Top             =   1212
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   4800
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6435
      TabIndex        =   8
      Top             =   4800
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6435
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   4710
      Width           =   2535
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Tag             =   "Cliente|N|N|0||sactuaobra|codclien|0000|S|"
      Text            =   "Text1"
      Top             =   735
      Width           =   975
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
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
         NumButtons      =   17
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   5520
         TabIndex        =   13
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2640
      Top             =   4920
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
      Index           =   2
      Left            =   1080
      Picture         =   "frmObraActua.frx":0012
      Tag             =   "-1"
      ToolTipText     =   "Buscar actuación"
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion envio"
      Height          =   255
      Left            =   6240
      TabIndex        =   23
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   4
      Left            =   3840
      Picture         =   "frmObraActua.frx":0114
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   3
      Left            =   1200
      Picture         =   "frmObraActua.frx":019F
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1080
      Picture         =   "frmObraActua.frx":022A
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1080
      Picture         =   "frmObraActua.frx":032C
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label7 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha fin"
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   2190
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha inicio"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Actuacion"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1695
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Obra"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   735
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
      TabIndex        =   11
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
Attribute VB_Name = "frmObraActua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatoForzarBusqueda As String   'Para cuando busque algo
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String

Private HaDevueltoDatos2 As String
Private PimeraVez As Boolean

'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


'Private Sub cboTipoDirec_KeyPress(KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo EAceptar
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'BUSCAR
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CadenaConsulta = "Select * from " & NombreTabla & " WHERE codclien=" & Text1(0).Text & " AND coddirec = " & Text1(1).Text
                    Me.Data1.RecordSource = CadenaConsulta
                    PosicionarData
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                 If ModificaDesdeFormulario(Me, 1) Then
                     TerminaBloquear
                     PosicionarData
                 End If
            End If
    End Select
EAceptar:
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


Private Sub cmdRegresar_Click()
Dim Cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    If DatosADevolverBusqueda <> "||" Then
    
        Cad = Val(RecuperaValor(DatosADevolverBusqueda, 1))
        If Val(Cad) <> Val(Data1.Recordset.Fields(0)) Then
            MsgBox "No coincide el cliente (" & Cad & ")", vbExclamation
            Exit Sub
        End If
        Cad = RecuperaValor(DatosADevolverBusqueda, 2)
        If Cad <> "" Then
            If Val(Cad) <> Val(Data1.Recordset.Fields(1)) Then
                MsgBox "No coincide la obra (" & Cad & ")", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    Cad = Data1.Recordset.Fields(0) & "|"
    Cad = Cad & Data1.Recordset.Fields(1) & "|"
     Cad = Cad & DBLet(Data1.Recordset.Fields(2), "F") & "|"
    'Pongo la fecha ini y el txt
    Cad = Cad & DBLet(Data1.Recordset!FechaIni, "T") & "|"
    Cad = Cad & DBLet(Data1.Recordset!observa, "T") & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub Form_Activate()
    If PimeraVez Then
        PimeraVez = False
        If Text1(0).Text <> "" Then cmdAceptar_Click
    End If
    Screen.MousePointer = vbDefault
    'If Modo = 1 Then PonerFoco Text1(0)
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    PimeraVez = True
    'ICONOS de La toolbar
    btnPrimero = 14 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 16 'Imprimir
        .Buttons(11).Image = 15 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    

    
    NombreTabla = "sactuaobra" 'Tabla Promociones Tarifas
    Ordenacion = " ORDER BY coddirec"
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE coddirec = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'Modo Busqueda

            For Modo = 0 To 2
                Text1(Modo).Text = RecuperaValor(DatosADevolverBusqueda, CInt(Modo) + 1)
            Next

            

        
        Modo = 1
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    'Formulario para Busqueda
    HaDevueltoDatos2 = CadenaDevuelta

End Sub






Private Sub frmC_Selec(vFecha As Date)
    HaDevueltoDatos2 = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Cad As String
Dim cad1 As String
Dim nada As String

Dim devuelve As String
    
    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    Select Case Index
    Case 0
        Cad = "Codigo|sclien|codclien|N|000|15·"
        Cad = Cad & "Nombre|sclien|nomclien|T||55·"
        frmB.vTabla = "sclien"
        frmB.vSQL = ""
        frmB.vTitulo = "Clientes"
    Case 1
        If Text1(0).Text = "" Then
           MsgBox "Primero debe indicar el cliente", vbExclamation
           Exit Sub
        End If
        Cad = "Obra|sdirec|coddirec|N|000|15·"
        Cad = Cad & "Desc.|sdirec|nomdirec|T||55·"
        
        frmB.vTabla = "sdirec"
        frmB.vSQL = "codclien = " & Text1(0).Text
        frmB.vTitulo = "Obras cliente: " & Text2(0).Text
        
    Case 2
        If Text1(0).Text = "" Or Text1(1).Text = "" Then
           MsgBox "Primero debe indicar un cliente y obra", vbExclamation
           Exit Sub
        End If
        Cad = "Actuacion|sactuaobra|actuacion|T||25·"
        Cad = Cad & "Observaciones|sactuaobra|observa|T||55·"
        
        frmB.vTabla = "sactuaobra"
        frmB.vSQL = "codclien = " & Text1(0).Text & " and coddirec = " & Text1(1).Text
        
    
         cad1 = "Actuaciones en obra : "
        
         devuelve = "codclien = " & Text1(0).Text & " AND coddirec "
         devuelve = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", devuelve, Val(Text1(1)))
         
        
        cad1 = cad1 & devuelve
                
        frmB.vTitulo = cad1
        
                
    End Select
    
    HaDevueltoDatos2 = ""
    frmB.vCampos = Cad
    frmB.vDevuelve = "0|1"
    frmB.vselElem = 0
    frmB.vCargaFrame = False
    frmB.vConexionGrid = 1
    frmB.Show vbModal
    Set frmB = Nothing
    
    If HaDevueltoDatos2 <> "" And Index <> 2 Then
        'Stop
        Text1(Index).Text = RecuperaValor(HaDevueltoDatos2, 1)
        Text2(Index).Text = RecuperaValor(HaDevueltoDatos2, 2)
        HaDevueltoDatos2 = ""
    End If
    
   If HaDevueltoDatos2 <> "" And Index = 2 Then
        'Stop
     
        Text1(Index).Text = RecuperaValor(HaDevueltoDatos2, 1)
        HaDevueltoDatos2 = ""
    End If
    
  
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
     If Modo = 2 Or Modo = 0 Then Exit Sub
     
     Set frmC = New frmCal
     HaDevueltoDatos2 = ""
     frmC.Fecha = Now
     If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
     frmC.Show vbModal
     If HaDevueltoDatos2 <> "" Then Text1(Index).Text = HaDevueltoDatos2
     
     
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


Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
On Error Resume Next

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    If Modo = 1 Then Exit Sub
    
    
    
    With Text1(Index)
        Select Case Index
            Case 0, 1 'cliente, obra
                devuelve = ""
                Text1(Index).Text = Trim(Text1(Index).Text)
                If Index = 1 Then
                    If Text1(0).Text = "" Then
                        MsgBox "Primero debe indicar el cliente", vbExclamation
                        Text1(1).Text = ""
                        PonerFoco Text1(0)
                    End If
                End If
                If Text1(Index).Text <> "" Then
                    If PonerFormatoEntero(Text1(Index)) Then
                       If Index = 0 Then
                            devuelve = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text)
                       Else
                            devuelve = "codclien = " & Text1(0).Text & " AND coddirec "
                            devuelve = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", devuelve, Text1(Index).Text)
                        End If
                        If devuelve = "" Then
                            MsgBox "No existe el codigo: " & Text1(Index).Text, vbExclamation
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    Else
                        Text1(Index).Text = ""
                    End If
                End If
                Text2(Index).Text = devuelve
            Case 2
                Text1(2).Text = UCase(Text1(2).Text)
            Case 3, 4
                PonerFormatoFecha Text1(Index)
                    
 
        End Select
    End With
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
        Case 10 'Imprimir
            
            'El formulario es compartido con el proyecto EULER
            'Cuando generemos aquel exe este trozo estara comentado
            'frmObraListado.opcion = 1
            'frmObraListado.Show vbModal
        Case 11  'Salir
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
Dim b As Boolean
Dim NumReg As Byte 'Solo para saber que hay + de 1 Registro

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    '--------------------------------------------
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
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea clave primaria
    BloquearText1 Me, Modo
    
    
    'Bloquear Registro sino es Insert o Update
    b = (Modo = 0) Or (Modo = 2)

    
           
    '------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    Me.Check1.Enabled = b
    
    Me.imgBuscar(0).Enabled = b
    Me.imgBuscar(1).Enabled = b
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activa las Opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b

    '-------------------------------------
    b = (Modo >= 3)
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
   Check1.Value = 0
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
    
        'Si pasamos el control
        If Me.DatosADevolverBusqueda <> "" Then
            Text1(0).Text = RecuperaValor(DatosADevolverBusqueda, 1)
            Text1(1).Text = RecuperaValor(DatosADevolverBusqueda, 2)
            PonerFoco Text1(2)
            Text1(2).BackColor = vbYellow
        Else
            PonerFoco Text1(0)
            Text1(0).BackColor = vbYellow
        End If
        
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
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    Text1(3).Text = Format(Now, "dd/mm/yyyy")
    
    If Me.DatosADevolverBusqueda <> "" Then
        Text1(0).Text = RecuperaValor(Me.DatosADevolverBusqueda, 1)
        Text1(1).Text = RecuperaValor(Me.DatosADevolverBusqueda, 2)
        Text1_LostFocus 0
        Text1_LostFocus 1
        PonerFoco Text1(2)
    Else
        PonerFoco Text1(0)
    End If
    
    'sugerir siguiente codigo. Si fuera secuencial el text1(3)......
    'Text1(3).Text = SugerirCodigoSiguienteStr(NombreTabla, "camopi")

    
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    SQL = SQL & "¿Seguro que desea eliminar la Dirección de Compras?"
    SQL = SQL & vbCrLf & "Cliente.  : " & Format(Text1(0).Text, "000") & " " & Text2(0).Text
    SQL = SQL & vbCrLf & "Obra      : " & Text1(1).Text & " " & Text2(1).Text
    SQL = SQL & vbCrLf & "Actuación : " & Text1(2).Text
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    If PuedeEliminar Then
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
        MuestraError Err.Number, "Eliminar Dirección", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
On Error GoTo FinEliminar
        
        SQL = " WHERE coddirec=" & Data1.Recordset!CodDirec
        SQL = SQL & " AND codclien=" & Data1.Recordset!codclien
        SQL = SQL & " AND actuacion=" & DBSet(Data1.Recordset!actuacion, "T")
        'Cabeceras
        conn.Execute "Delete  from " & NombreTabla & SQL
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Eliminar = False
    Else
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
        
    DatosOk = True
End Function


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String

    'Llamamos a al form
    Cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    Cad = Cad & ParaGrid(Text1(0), 8, "Cod. Cli.")
    
    Cad = Cad & "Nombre|sclien|nomclien|T||35·"
    Cad = Cad & ParaGrid(Text1(1), 8, "Obra")
    Cad = Cad & "Desc. obra|sdirec|nomdirec|T||30·"
    Cad = Cad & ParaGrid(Text1(2), 18, "Actuacion")
    
               
 
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = NombreTabla & ",sclien,sdirec"
        
        Cad = "sactuaobra.codclien=sclien.codclien and sdirec.codclien=sactuaobra.codclien and sdirec.coddirec=sactuaobra.coddirec"

        If cadB <> "" Then Cad = Cad & " AND " & cadB
        
        frmB.vSQL = Cad
        HaDevueltoDatos2 = ""
        '###A mano
        frmB.vDevuelve = "0|2|4|"
        frmB.vTitulo = "Actuaciones en obras"
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos2 <> "" Then
            'Creamos una cadena consulta y ponemos los datos
            Screen.MousePointer = vbHourglass
            cadB = ""
            Cad = ValorDevueltoFormGrid(Text1(0), HaDevueltoDatos2, 1)
            cadB = Cad
            Cad = ValorDevueltoFormGrid(Text1(1), HaDevueltoDatos2, 2)
            cadB = cadB & " and " & Cad
            Cad = ValorDevueltoFormGrid(Text1(2), HaDevueltoDatos2, 3)
            cadB = cadB & " and " & Cad
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            
            
        End If

        

    Screen.MousePointer = vbDefault
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
    
    Modo = 3  'truco parqa que haga el lostfocus
    Text1_LostFocus 0
    Text1_LostFocus 1
    Modo = 2  'A ponercampos siempre entra con modo=2
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub





Private Sub PosicionarData()
Dim vWhere As String, Indicador As String

    vWhere = "codclien=" & Text1(0).Text & " AND coddirec = " & Text1(1).Text & " AND actuacion = " & DBSet(Text1(2).Text, "T")
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'        LimpiarCampos
        PonerModo 0
    End If
End Sub

Private Function PuedeEliminar() As Boolean
Dim C As String
Dim Aux As String
    Aux = ""
    C = " coddirec = " & Text1(1).Text & " AND actuacion = " & DBSet(Text1(2).Text, "T") & " AND codclien "
    C = DevuelveDesdeBD(conAri, "count(*)", "scapre", C, Text1(0).Text, "N")
    If C = "" Then C = "0"
    If Val(C) <> 0 Then Aux = Aux & "   -Ofertas" & vbCrLf
        
        
    C = " coddirec = " & Text1(1).Text & " AND actuacion = " & DBSet(Text1(2).Text, "T") & " AND codclien "
    C = DevuelveDesdeBD(conAri, "count(*)", "scaped", C, Text1(0).Text, "N")
    If C = "" Then C = "0"
    If Val(C) <> 0 Then Aux = Aux & "   -Pedidos" & vbCrLf
        
        
    C = " coddirec = " & Text1(1).Text & " AND actuacion = " & DBSet(Text1(2).Text, "T") & " AND codclien "
    C = DevuelveDesdeBD(conAri, "count(*)", "scaalb", C, Text1(0).Text, "N")
    If C = "" Then C = "0"
    If Val(C) <> 0 Then Aux = Aux & "   -Albaranes" & vbCrLf
    
    'Partes de trabajo
    C = " coddirec = " & Text1(1).Text & " AND actuacion = " & DBSet(Text1(2).Text, "T") & " AND codclien "
    C = DevuelveDesdeBD(conAri, "count(*)", "sliparte", C, Text1(0).Text, "N")
    If C = "" Then C = "0"
    If Val(C) <> 0 Then Aux = Aux & "   -Partes de trabajo" & vbCrLf
    
    
    If Aux <> "" Then
        PuedeEliminar = False
        Aux = "Existen datos relacionados con la actuacion en : " & vbCrLf & Aux
        MsgBox Aux, vbQuestion
    Else
        PuedeEliminar = True
    End If
    
End Function
