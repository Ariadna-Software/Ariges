VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelematMto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fichero TELEMATEL"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "frmTelematMto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   6840
      TabIndex        =   8
      Tag             =   "Fecha|F|N|||stelem|fechacambio|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   2490
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   3600
      TabIndex        =   7
      Tag             =   "Precio|N|N|||stelem|precio|#,##0.0000||"
      Text            =   "Text1"
      Top             =   2490
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Tag             =   "Uds precio|N|N|1||stelem|uniprec|||"
      Text            =   "Text1"
      Top             =   2490
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   5760
      TabIndex        =   5
      Tag             =   "EAN|T|S|||stelem|codean||N|"
      Text            =   "Text1"
      Top             =   2010
      Width           =   2445
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Tag             =   "Referencia provedor|T|N|||stelem|referprov|||"
      Text            =   "Text1"
      Top             =   2010
      Width           =   2325
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   3
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   1560
      Width           =   3585
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Tag             =   "Codigo proveedor|N|S|||stelem|codprove|||"
      Text            =   "Text1"
      Top             =   1560
      Width           =   645
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   1080
      Width           =   5025
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||stelem|nombre|||"
      Top             =   600
      Width           =   5055
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Tag             =   "Codigo en ariges|T|S|||stelem|codartic|||"
      Text            =   "Text1"
      Top             =   1080
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Tag             =   "Código telematel|N|N|0||stelem|codtelem|00000000|S|"
      Text            =   "Text"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   3000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   3000
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   270
      Top             =   3075
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
      TabIndex        =   16
      Top             =   0
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Capturar y actualizar codigos"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar precios"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Crear articulo"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar de un proveedor"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Importar fichero"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir descuadre referencias"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6960
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   8
      Left            =   6480
      Picture         =   "frmTelematMto.frx":000C
      ToolTipText     =   "Buscar centro coste"
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha cambio"
      Height          =   195
      Index           =   7
      Left            =   5880
      TabIndex        =   25
      Top             =   2550
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Precio"
      Height          =   195
      Index           =   6
      Left            =   2880
      TabIndex        =   24
      Top             =   2550
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Ud. precio"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "EAN"
      Height          =   195
      Index           =   4
      Left            =   5280
      TabIndex        =   22
      Top             =   2070
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Ref. proveedor"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   960
      Picture         =   "frmTelematMto.frx":0596
      ToolTipText     =   "Buscar centro coste"
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Cod.artículo"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Cod."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   735
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
         HelpContextID   =   1
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   1
         Shortcut        =   ^E
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnActCodigos 
         Caption         =   "Capturar y actualizar codigos"
      End
      Begin VB.Menu mnActImportes 
         Caption         =   "Actualiza importes"
      End
      Begin VB.Menu mnCrearArticulo 
         Caption         =   "Crear artículo"
         HelpContextID   =   1
      End
      Begin VB.Menu mnTelematel 
         Caption         =   "Importar fichero telematel"
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
Attribute VB_Name = "frmTelematMto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmP As frmComProveedores
Attribute frmP.VB_VarHelpID = -1

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
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
Private ModoAnterior As Byte

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1







Private Sub cmdAceptar_Click()
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
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
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1 'Busqueda
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
    End Select
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadir()
'    LimpiarCampos
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    ModoAnterior = Modo 'Para el botón cancelar en Modo Insertar
'    PonerModo 3
'    Text1(0).Text = SugerirCodigoSiguienteStr("stelem", "codtelem")
'    FormateaCampo Text1(0)
'    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else 'Modo=1 Busqueda
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
        MandaBusquedaPrevia "", True
    Else
        CadenaConsulta = "Select * from " & NombreTabla
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
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    
    '### a mano
    Cad = "¿Seguro que desea eliminar la Familia de Artículo?:" & vbCrLf
    Cad = Cad & vbCrLf & "Cod. : " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Desc.: " & Data1.Recordset.Fields(1)

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
        MuestraError Err.Number, "Eliminar Familia de Articulo", Err.Description
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
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    If vParamAplic.Descriptores Then Me.Caption = "Categorias Art."
    ' ICONITOS DE LA BARRA
    btnPrimero = 23 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        
        '.Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(9).Image = 37   '
        .Buttons(10).Image = 31  '
        .Buttons(11).Image = 39  '
        .Buttons(12).Image = 14  'borrar proveedor
        .Buttons(13).Image = 13  'importar fich
        
        
        .Buttons(19).Image = 16  ' Imprimir
        .Buttons(20).Image = 40  ' Imprimir
        .Buttons(21).Image = 15  'Salir
        
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    
    'Pone el Tag del primer botón de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: stelem, BD: Ariges
    'Si tag>0 abre busqueda en la tabla: Cuentas, BD: Conta

        
  
    '## A mano
    NombreTabla = "stelem"
    Ordenacion = " ORDER BY codtelem"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
       
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE codtelem=-1"
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
Dim Indice As Byte
    
    If CadenaDevuelta <> "" Then
            
            If Val(Me.Tag) = "1" Then
   
                    'Recupera todo el registro de Banco Propio
        
                    Screen.MousePointer = vbHourglass
                    'Sabemos que campos son los que nos devuelve
                    'Creamos una cadena consulta y ponemos los datos
                    CadB = ""
                    Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
                    CadB = Aux
                    '   Como la clave principal es unica, con poner el sql apuntando
                    '   al valor devuelto sobre la clave ppal es suficiente
                    'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
                    'If CadB <> "" Then CadB = CadB & " AND "
                    'CadB = CadB & Aux
                    'Se muestran en el mismo form
                    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                    PonerCadenaBusqueda
                    Screen.MousePointer = vbDefault
            Else
            
                'prove
                Text1(3).Text = RecuperaValor(CadenaDevuelta, 1)
                Text2(3).Text = RecuperaValor(CadenaDevuelta, 2)
            End If
    End If
End Sub


Private Sub imgBuscar_Click(Index As Integer)
    ' ---- [19/10/2009] [LAURA]: añadir campo centro de coste familia
    
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    
    Select Case Index
        Case 3 '
            MandaBusquedaPrevia "", False
            
        Case 8
            Set frmC = New frmCal
            frmC.Fecha = Now
            frmC.Show vbModal
            Set frmC = Nothing
    End Select
End Sub




Private Sub mnActCodigos_Click()
    BotonActualizarCodigos
End Sub

Private Sub mnActImportes_Click()
    BotonActualizarImportes
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
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
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
Dim C As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo

             If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar si ya existe el cod
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If


        Case 3
            'Codprove
            C = ""
            If Text1(3).Text <> "" Then
                If Not PonerFormatoEntero(Text1(3)) Then
                    Text1(3).Text = ""
                Else
                    C = PonerNombreDeCod(Text1(3), conAri, "sprove", "nomprove", "codprove")
                    
                End If
            End If
            Me.Text2(3).Text = C
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
    
    CadB = ObtenerBusqueda(Me, False)

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB, True
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String, EsBusqueda As Boolean)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String

            Screen.MousePointer = vbHourglass
            'Llamamos a al form
            '##A mano
            Cad = ""
      
            'Busqueda de una Família de Artículo
             Set frmB = New frmBuscaGrid
            If EsBusqueda Then
                '
                Cad = Cad & ParaGrid(Text1(0), 14, "Código")
                Cad = Cad & ParaGrid(Text1(1), 61, "Nombre")
                Cad = Cad & ParaGrid(Text1(4), 25, "Ref prove")
                frmB.vTabla = "stelem"
                frmB.vTitulo = "Fichero telematel"
            Else
                'PREOVEEDORES
                Cad = Cad & "Código|sprove|codprove|N|000000|18·"
                Cad = Cad & "Nombre|sprove|nomprove|T||40·"
                Cad = Cad & "Nom.Comer.|sprove|nomcomer|T||40·"
            
                frmB.vTabla = "sprove"
                frmB.vTitulo = "Proveedores"
            End If
            frmB.vCampos = Cad
            
            Me.Tag = Abs(EsBusqueda)   '0- Proveedores 1-Buscaprevia
            frmB.vSQL = CadB
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vselElem = 1
            frmB.vConexionGrid = conAri
            frmB.vCargaFrame = False
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            Me.Tag = ""
 
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
        Data1.Recordset.MoveFirst
        PonerModo 2
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
Dim I As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "sartic", "nomartic", "codartic")
    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "sprove", "nomprove", "codprove")
    
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
    PonerIndicador lblIndicador, Modo
    
    '-------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    
    'Poner Boton de Cabecera o Aceptar/Cancelar
    PonerBotonCabecera B Or (Modo = 0)
        
    'Bloquear Registros si modo distinto de Insertar o Modificar
    BloquearText1 Me, Modo
    
    BloquearChecks Me, Modo
        
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según MODO
    PonerOpcionesMenu   'Activar opciones de menu según NIVEL
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
On Error Resume Next

    'b = (Modo = 2) Or (Modo = 0) Or (Modo = 1)
    B = Modo <= 2
    
    Toolbar1.Buttons(9).Enabled = B
    mnActCodigos.Enabled = B
    
    Toolbar1.Buttons(10).Enabled = B
    mnActImportes.Enabled = B
    
    Toolbar1.Buttons(13).Enabled = B
    mnTelematel.Enabled = B
    
    'Añadir
   ' Toolbar1.Buttons(5).Enabled = b
   ' Me.mnNuevo.Enabled = b
    
    B = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    mnEliminar.Enabled = B
    
    Toolbar1.Buttons(11).Enabled = B
    mnCrearArticulo.Enabled = B
    

    
    
    
    
     '---------------------------------
    B = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean

    DatosOk = False
    B = CompForm(Me, 1)
    If Not B Then Exit Function
    
    'Comprobar si ya existe el cod de familia en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then B = False
    End If
    
    DatosOk = B
End Function





Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Precio As Currency


    'If Button.index >= 12 And Button.index <= 13 Then
    If Button.Index = 11 Then
        If Modo <> 2 Then Exit Sub
    End If
    If Button.Index = 12 Then
        If Modo <> 2 And Modo <> 0 Then Exit Sub
    End If
    
    Select Case Button.Index
        Case 1: mnBuscar_Click
        Case 2: mnVerTodos_Click
        Case 5  'Nuevo
                mnNuevo_Click
        Case 6  'Modificar
                mnModificar_Click
        Case 7  'Borrar
                mnEliminar_Click
                
               
        Case 9
                BotonActualizarCodigos
        Case 10
                BotonActualizarImportes
           
        Case 11
            'CrearArticulo
                'codprove|nomprove|refprove|precio|nomartic|ean|codtelem|
                If Data1.Recordset.EOF Then Exit Sub
                If DBLet(Data1.Recordset!codArtic, "T") <> "" Then
                    MsgBox "Ya tiene asignado articulo", vbExclamation
                    Exit Sub
                End If
                Precio = ImporteFormateado(Text1(7).Text)
                NumRegElim = Val(Text1(6).Text)
                Precio = Precio / NumRegElim
                CadenaDesdeOtroForm = ""
                With frmAlmArticulos
                    'codprove|nomprove|refprove|precio|nomartic|ean|codtelem|
                   .DatosADevolverBusqueda = "··" & Text1(3).Text & "|" & Text2(3).Text & "|" & Text1(4).Text & "|" & CStr(Precio) & "|" & Text1(1).Text & _
                        "|" & Text1(5) & "|" & Text1(0).Text & "|"
                   .Show vbModal
                End With
                
                If CadenaDesdeOtroForm <> "" Then
                    Screen.MousePointer = vbHourglass
                    'SE HA INSERTADO
                    conn.Execute "commit"
                    Espera 0.5
                    CadenaDesdeOtroForm = ""
                    PosicionarData
                    Screen.MousePointer = vbHourglass
                    If Not Data1.Recordset.EOF Then PonerCampos
                    
                End If
               Screen.MousePointer = vbDefault
        
        Case 12
            CadenaDesdeOtroForm = ""
            frmListado3.opcion = 56
            frmListado3.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                Screen.MousePointer = vbHourglass
                If CadenaDesdeOtroForm = "null" Then
                    CadenaDesdeOtroForm = " is " & CadenaDesdeOtroForm
                Else
                    CadenaDesdeOtroForm = " = " & CadenaDesdeOtroForm
                End If
                ejecutar " delete from stelem where codprove " & CadenaDesdeOtroForm, False
                If Modo = 2 Then BotonVerTodos
                Screen.MousePointer = vbDefault
            End If
        Case 13
            BotonTelematel
            
        Case 19 'Imprimir listado
            BotonImprimir False
            
        Case 20
            BotonImprimir True
            
        Case 21: mnSalir_Click
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


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    If B Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codtelem=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub


Private Sub BotonImprimir(Descuadre As Boolean)
    If Descuadre Then
        frmTelematVarios.opcion = 3
    Else
        frmTelematVarios.opcion = 2
    End If
    frmTelematVarios.Show vbModal
End Sub




Private Sub BotonActualizarCodigos()
    frmTelematVarios.opcion = 1
    frmTelematVarios.Show vbModal
End Sub


Private Sub BotonActualizarImportes()
    frmTelematVarios.opcion = 0
    frmTelematVarios.Show vbModal
End Sub


Private Sub BotonTelematel()
    frmTelematImportar.Show vbModal
    PonerModo 0
End Sub
