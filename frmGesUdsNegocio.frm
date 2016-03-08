VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGesUdsNegocio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unidades de negocio"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmGesUdsNegocio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   13
      Left            =   5400
      MaxLength       =   30
      TabIndex        =   37
      Tag             =   "Sitaucion de baja en Ariges|N|S|||unidadesnegocio|codsituabaja|||"
      Text            =   "Text1"
      Top             =   3240
      Width           =   405
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   12
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   36
      Text            =   "Text2"
      Top             =   5160
      Width           =   3045
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   12
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   34
      Tag             =   "Forma de pago|N|S|||unidadesnegocio|forpa||N|"
      Text            =   "Text1"
      Top             =   5160
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Seccion euroagro|T|N|||unidadesnegocio|CodSeccEuroagro||N|"
      Text            =   "Text1"
      Top             =   3240
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Código de Banco Propio|N|N|0|9999|unidadesnegocio|IdUnidad|0000|S|"
      Text            =   "Text1"
      Top             =   510
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Denominación|T|N|||unidadesnegocio|Nombre||N|"
      Text            =   "Text1"
      Top             =   990
      Width           =   4245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "Domicilio del Banco Propio|T|S|||unidadesnegocio|Direccion||N|"
      Text            =   "Text1"
      Top             =   1470
      Width           =   4245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "CP|N|S|||unidadesnegocio|CodPostal||N|"
      Text            =   "Text1"
      Top             =   1965
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   3360
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "Población|T|S|||unidadesnegocio|Poblacion||N|"
      Text            =   "Text1"
      Top             =   1965
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   5
      Tag             =   "Teléfono|T|S|||unidadesnegocio|Telefono||N|"
      Text            =   "Text1"
      Top             =   2520
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   3360
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "Empresa conta|N|S|||unidadesnegocio|empresa_conta||N|"
      Text            =   "Text1"
      Top             =   3240
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   6
      Tag             =   "Identif. Cedente|T|S|||unidadesnegocio|Fax||N|"
      Text            =   "Text1"
      Top             =   2520
      Width           =   1365
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   9
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   3840
      Width           =   3045
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   10
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   4200
      Width           =   3045
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   11
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   4560
      Width           =   3045
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   11
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   11
      Tag             =   "s|T|S|||unidadesnegocio|raiz_proveedor||N|"
      Text            =   "Text1"
      Top             =   4560
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "Cta.|T|S|||unidadesnegocio|raiz_cliente_asociado||N|"
      Text            =   "Text1"
      Top             =   4200
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "Cta.|T|S|||unidadesnegocio|raiz_cliente_socio||N|"
      Text            =   "Text1"
      Top             =   3840
      Width           =   645
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   4770
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   315
      TabIndex        =   18
      Top             =   5835
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4770
      TabIndex        =   14
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   6000
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   270
      Top             =   6075
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
      TabIndex        =   22
      Top             =   0
      Width           =   6090
      _ExtentX        =   10742
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
         Left            =   4680
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Situa. baja(ariges)"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   38
      Top             =   3270
      Width           =   1455
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   3
      Left            =   1800
      ToolTipText     =   "Buscar cuenta contable"
      Top             =   5160
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Forma de pago(ges)"
      Height          =   255
      Index           =   2
      Left            =   195
      TabIndex        =   35
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5880
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Seccion euroagro"
      Height          =   255
      Index           =   18
      Left            =   165
      TabIndex        =   33
      Top             =   3270
      Width           =   1455
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Left            =   960
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   1995
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa conta"
      Height          =   255
      Index           =   16
      Left            =   2040
      TabIndex        =   32
      Top             =   3270
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fax"
      Height          =   195
      Index           =   15
      Left            =   2880
      TabIndex        =   31
      Top             =   2580
      Width           =   255
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   1
      Left            =   1725
      ToolTipText     =   "Buscar cuenta contable"
      Top             =   4245
      Width           =   240
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   2
      Left            =   1725
      ToolTipText     =   "Buscar cuenta contable"
      Top             =   4590
      Width           =   240
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   0
      Left            =   1725
      Tag             =   "-1"
      ToolTipText     =   "Buscar cuenta contable"
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Raiz cliente socio"
      Height          =   255
      Index           =   5
      Left            =   195
      TabIndex        =   30
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Raiz cliente asociado"
      Height          =   255
      Index           =   6
      Left            =   195
      TabIndex        =   29
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Raiz proveedor"
      Height          =   255
      Index           =   7
      Left            =   195
      TabIndex        =   28
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Teléfono"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   27
      Top             =   2550
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
      Height          =   255
      Index           =   11
      Left            =   2280
      TabIndex        =   26
      Top             =   1995
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "C.Postal"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   25
      Top             =   1995
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   24
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Denominación"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   510
      Width           =   615
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
Attribute VB_Name = "frmGesUdsNegocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1

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

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                
                    If Data1.Recordset.EOF Then 'No estaba cargado Inicialmente
                        Data1.RecordSource = "Select * from " & NombreTabla & ObtenerWhereCP
                        Data1.Refresh
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
    End Select
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            PonerModo 0
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
    End Select
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    Text1(0).Text = Format(SugerirCodigoSiguienteStr("unidadesnegocio", "codbanpr"), "0000")
    PonerFoco Text1(0)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
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
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
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
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()


    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    



Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Banco Propio", Err.Description
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

    'Icono de busqueda
    Me.imgBuscar.Picture = frmPpal.imgListComun.ListImages(19).Picture
    For kCampo = 0 To Me.imgCuentas.Count - 1
        Me.imgCuentas(kCampo).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Next kCampo


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
    
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "unidadesnegocio"
    Ordenacion = " ORDER BY IdUnidad"
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where IdUnidad=-1"
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
Dim indice As Byte
    
    If CadenaDevuelta <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde un botón de busqueda de Cuentas
            'Recuperar solo el campo código y Descripción
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            indice = Val(Me.imgCuentas(0).Tag)
            Me.Text1(indice + 9).Text = RecuperaValor(CadenaDevuelta, 1)
            Me.Text2(indice + 9).Text = RecuperaValor(CadenaDevuelta, 2)

        Else
            'Recupera todo el registro de Banco Propio
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(indice + 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Poblacion
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


Private Sub imgCuentas_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    
    If Modo = 3 Then
        MsgBox "No se puede asignar cuentas en el alta", vbExclamation
        Exit Sub
    End If
    
    If Index = 3 Then
        'FORMA DE PAGO
        NumRegElim = DBLet(Data1.Recordset!IdUnidad, "T")
        If NumRegElim = 1 Or NumRegElim = 3 Then
            MsgBox "Valido solo para ariges", vbExclamation
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbHourglass

    imgCuentas(0).Tag = Index
    If Index = 3 Then
        MandaBusquedaPrevia ""
    Else
        MandaBusquedaPrevia "length(codmacta)=5"
    End If
    PonerFoco Text1(Index)
    imgCuentas(0).Tag = -1
    Screen.MousePointer = vbDefault
End Sub

Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    If Index = 0 Then dirMail = Text1(7).Text
    If LanzaMailGnral(dirMail) Then Espera 2
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2
    If LanzaHomeGnral(Text1(8).Text) Then Espera 2
    Screen.MousePointer = vbDefault
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
    
    MsgBox "Opcion NO disponible", vbExclamation
    Exit Sub
    BotonAnyadir
End Sub


Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
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
      
    'en el campo ID de norma 34 no se hace Trim ni nada. Lo q pongan
    If Index = 18 Then Exit Sub
      
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
         Case 0 'Cod. Banco Propio
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    'Detectamos aki si ya existe y no esperamos hasta boton Aceptar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
            
         Case 3 'CPostal
            If Text1(Index).Text = "" Then
                Text1(Index + 1).Text = ""
            ElseIf Not VieneDeBuscar Then
                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
            End If
            VieneDeBuscar = False
            
         Case 10, 11 'codbanco, codsucursal
            PonerFormatoEntero Text1(Index)
            
         Case 12, 13 'DC, numero cta
            FormateaCampo Text1(Index)
            
         Case 14, 15, 16, 17 'Cuentas
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)
            If Text1(Index).Text <> "" And Text2(Index).Text = "" Then PonerFoco Text1(Index)
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Conexion As Byte
Dim CargaF As Boolean 'Para saber si se carga el frame o no en el BuscaGrid
        
        'Llamamos a al form
        '##A mano
        Cad = ""
        
        If Val(Me.imgCuentas(0).Tag) = 3 Then
            'FORMA DE PAGO
            NumRegElim = DBLet(Data1.Recordset!empresa_conta, "T")
            Cad = Cad & "Código|sforpa|codforpa|N|000|20·Denominacion|sforpa|nomforpa|T||70·"
            tabla = "ariges" & NumRegElim & ".sforpa"
            Titulo = "Formas de pago"
            Conexion = conAri
            CargaF = False
            
        ElseIf Val(Me.imgCuentas(0).Tag) >= 0 Then
        'Se llama a Busqueda desde un campo de Cuenta
            '#A MANO: Porque busca en la tabla Cuentas
            'de la base de datos de Contabilidad
            
            Cad = Cad & "Código|cuentas|codmacta|T||30·Denominacion|cuentas|nommacta|T||70·"
            tabla = "cuentas"
            Titulo = "Cuentas"
            Conexion = conConta    'Conexión a BD: Conta
            CargaF = True
        Else
            'Busqueda de un Banco Propio
            Cad = Cad & ParaGrid(Text1(0), 30, "Código")
            Cad = Cad & ParaGrid(Text1(1), 70, "Denominacion")
            tabla = "unidadesnegocio"
            Titulo = "Bancos Propios"
            Conexion = conAri    'Conexión a BD: Ariges
            CargaF = False
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
            frmB.vselElem = 1
            frmB.vConexionGrid = Conexion
'            frmB.vBuscaPrevia = chkVistaPrevia
            frmB.vCargaFrame = CargaF
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
'                If kCampo < 17 Then Text1(kCampo + 1).SetFocus
'                If kCampo = 17 Then cmdAceptar.SetFocus
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(kCampo)
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
    For I = 9 To 12
        If Text1(I).Text = "" Then
            Text2(I).Text = ""
        Else
            
            
            
            If I = 12 Then
                CadenaConsulta = DBLet(Data1.Recordset!empresa_conta, "T")
                CadenaConsulta = "ariges" & CadenaConsulta & ".sforpa"
                Text2(I).Text = PonerNombreDeCod(Text1(I), conConta, CadenaConsulta, "nomforpa", "codforpa", , "N")
            Else
                'Cuentas
                CadenaConsulta = DBLet(Data1.Recordset!empresa_conta, "T")
                If CadenaConsulta <> "" Then CadenaConsulta = "conta" & CadenaConsulta & "."
                CadenaConsulta = CadenaConsulta & "cuentas"
                Text2(I).Text = PonerNombreDeCod(Text1(I), conConta, CadenaConsulta, "nommacta", "codmacta", , "T")
            End If
        End If
    Next I
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte
   
    Modo = Kmodo
        
    '----------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    PonerIndicador lblIndicador, Modo
    
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
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, CLng(NumReg)
    
    
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = b Or Modo = 1
    cmdCancelar.visible = b Or Modo = 1
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa botones de la Toolbar segun el Modo
Dim b As Boolean
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    mnEliminar.Enabled = b
    
    '-----------------------------------------
    b = (Modo >= 3) 'Insertar/Modificar
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
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then b = False
    End If

    DatosOk = b
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click  'Buscar
        Case 2: mnVerTodos_Click  'Todos
            
        Case 5: mnNuevo_Click  'Nuevo
        Case 6: mnModificar_Click  'Modificar
        Case 7: mnEliminar_Click  'Borrar
            
        Case 10
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
Dim Cad As String
Dim Indicador As String

    Cad = "(codbanpr=" & Text1(0).Text & ")"
    If SituarData(Data1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'        LimpiarCampos
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
On Error Resume Next
    ObtenerWhereCP = " WHERE codbanpr= " & Text1(0).Text
    If Err.Number <> 0 Then Err.Clear
End Function


