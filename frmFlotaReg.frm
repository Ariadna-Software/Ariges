VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFlotaReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro entrada flotas"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   5985
   ClipControls    =   0   'False
   Icon            =   "frmFlotaReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   3480
      TabIndex        =   9
      Tag             =   "ProxKM|F|S|||sflotasregistro|proximo|||"
      Text            =   "Text1"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   3120
      TabIndex        =   2
      Tag             =   "Base imp.|N|N|||sflotasregistro|baseimp|##,##0.00||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame FrameConKim 
      Height          =   1095
      Left            =   240
      TabIndex        =   33
      Top             =   4830
      Width           =   5655
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   480
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   4560
         TabIndex        =   13
         Tag             =   "Base imp.|N|S|0|9999|sflotasregistro|consumo|##,##0.00||"
         Text            =   "Text1"
         Top             =   450
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   3360
         TabIndex        =   12
         Tag             =   "Litros ult. tiquet|N|S|0||sflotasregistro|litrosulttik|##,##0.00||"
         Text            =   "Text1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   1200
         TabIndex        =   11
         Tag             =   "Litros ante.|N|S|0||sflotasregistro|litros|##,##0.00||"
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Tag             =   "Km Iniciales|N|S|0||sflotasregistro|kminciales|||"
         Text            =   "Text1"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Anterior"
         Height          =   195
         Index           =   11
         Left            =   2520
         TabIndex        =   42
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Consumo %"
         Height          =   195
         Index           =   10
         Left            =   4560
         TabIndex        =   40
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Litros ult.tiket"
         Height          =   195
         Index           =   9
         Left            =   3360
         TabIndex        =   39
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Litros (Sin ultim)"
         Height          =   195
         Index           =   8
         Left            =   1200
         TabIndex        =   38
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Km iniciales"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   35
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   4560
      TabIndex        =   3
      Tag             =   "Horas KM|N|S|||sflotasregistro|horaskm|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   1440
      TabIndex        =   8
      Tag             =   "Caduca|F|S|||sflotasregistro|caduca|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "Concepto|N|N|||sflotasregistro|codconcef|||"
      Top             =   2160
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   1680
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||sflotasregistro|fecha|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   6
      Left            =   240
      MaxLength       =   35
      TabIndex        =   0
      Tag             =   "Registro|T|S|||sflotasregistro|registro||S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   1560
      Width           =   4125
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   3600
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      Height          =   1155
      Index           =   5
      Left            =   240
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Tag             =   "Ob|T|S|||sflotasregistro|Observaciones|||"
      Text            =   "frmFlotaReg.frx":000C
      Top             =   6360
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   240
      MaxLength       =   35
      TabIndex        =   6
      Tag             =   "A|T|S|||sflotasregistro|ampliacion|||"
      Text            =   "Text1"
      Top             =   2880
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Tag             =   "Vehiculo|T|N|||sflotasregistro|codflota|||"
      Text            =   "Text1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4875
      TabIndex        =   16
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   4875
      TabIndex        =   17
      Top             =   7800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   7680
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Tag             =   "Proveedor|N|S|||sflotasregistro|codprove|||"
      Text            =   "Text1"
      Top             =   3600
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   5985
      _ExtentX        =   10557
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
         Left            =   4800
         TabIndex        =   19
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4440
      Top             =   6720
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
   Begin VB.Label Label1 
      Caption         =   "PROXIMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   37
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Kms"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   36
      Top             =   4230
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Base Imp"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   34
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   32
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Horas-Km"
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   31
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   29
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   1920
      Width           =   735
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   3
      Left            =   2280
      Picture         =   "frmFlotaReg.frx":0012
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Nº Regsitro"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   960
      Picture         =   "frmFlotaReg.frx":009D
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1080
      Picture         =   "frmFlotaReg.frx":019F
      Tag             =   "-1"
      ToolTipText     =   "Buscar población"
      Top             =   3360
      Width           =   240
   End
   Begin VB.Label Label7 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Ampliacion"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Vehiculo"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   30
      Top             =   3360
      Width           =   1095
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
   Begin VB.Menu mnOrdenacion 
      Caption         =   "Orden&ación"
      Begin VB.Menu mnOrden1 
         Caption         =   "Registro entrada"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnOrden1 
         Caption         =   "Fecha "
         Index           =   1
      End
      Begin VB.Menu mnOrden1 
         Caption         =   "Vehiculo"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmFlotaReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IncidenciaLlenarDeposito = 7

Public DatoForzarBusqueda As String   'Para cuando busque algo
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1


Dim NombreTabla As String

Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String

Private HaDevueltoDatos2 As String
Private PimeraVez As Boolean

Private DatosVehiculo As String
Private ConceptosRequierenKm As String  'Llevara empipados los conceptos que requieran KM

Dim cad As String

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
                    CadenaConsulta = "Select * from " & NombreTabla & " WHERE registro=" & Text1(6).Text
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
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = Val(RecuperaValor(DatosADevolverBusqueda, 1))
    If Val(cad) <> Val(Data1.Recordset.Fields(0)) Then
        MsgBox "No coincide el cliente (" & cad & ")", vbExclamation
        Exit Sub
    End If
    cad = Val(RecuperaValor(DatosADevolverBusqueda, 2))
    If Val(cad) <> Val(Data1.Recordset.Fields(1)) Then
        MsgBox "No coincide la obra (" & cad & ")", vbExclamation
        Exit Sub
    End If
        
    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(1) & "|"
     cad = cad & DBLet(Data1.Recordset.Fields(2), "F") & "|"
    'Pongo la fecha ini y el txt
    cad = cad & DBLet(Data1.Recordset!FechaIni, "T") & "|"
    cad = cad & DBLet(Data1.Recordset!observa, "T") & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Combo1_Click()
    
    If Modo = 2 Then Exit Sub
    
    
    HaDevueltoDatos2 = "|" & Combo1.ListIndex & "|"
    Me.FrameConKim.visible = InStr(1, ConceptosRequierenKm, HaDevueltoDatos2) > 0
    HaDevueltoDatos2 = ""
    
     
    If Modo >= 3 Then PonerProximo
    
        
        
    
End Sub

Private Sub FijarProximaRevisionKmFecha()
    If Combo1.ListIndex < 0 Then Exit Sub
    If Me.Text1(1).Text = "" Then Exit Sub
    
    'Modificar e insertar
    'Proxima revision en fecha y/o KM
        
    'flotascon_x_tipo
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PimeraVez Then
        PimeraVez = False
       
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
    
    CargaComboConceptos

    
    NombreTabla = "sflotasregistro" 'Tabla Promociones Tarifas
   
    CadenaConsulta = "Select * from " & NombreTabla & " WHERE registro = -1" 'No recupera datos
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'Modo Busqueda
        
        Modo = 1
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Function Ordenacion() As String
    Ordenacion = " ORDER BY "
    If mnOrden1(0).Checked = True Then
        Ordenacion = Ordenacion & "registro "
    ElseIf mnOrden1(0).Checked = True Then
        Ordenacion = Ordenacion & "fecha,codflota "
    Else
        Ordenacion = Ordenacion & "codflota,fecha "
    End If
End Function
Private Sub frmB_Selecionado(CadenaDevuelta As String)
    'Formulario para Busqueda
    HaDevueltoDatos2 = CadenaDevuelta

End Sub






Private Sub frmC_Selec(vFecha As Date)
    HaDevueltoDatos2 = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim cad As String
Dim cad1 As String
Dim nada As String

Dim devuelve As String
    
    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    Select Case Index
    Case 0
        cad = "Codigo|sprove|codprove|N|000|15·"
        cad = cad & "Nombre|sprove|nomprove|T||55·"
        frmB.vTabla = "sprove"
        frmB.vSQL = ""
        frmB.vTitulo = "Proveedores"
    Case 1
        
        cad = "Codigo|sflotas|codflota|T||15·"
        cad = cad & "Desc.|sflotas|nomflota|T||45·"
        cad = cad & "Desc.|sflotatipo|desctipflota|T||30·"
        
        frmB.vTabla = "sflotas,sflotatipo"
        frmB.vSQL = "sflotas.tipo=tipflota "
        frmB.vTitulo = "Vehiculos"
        
    Case 2
'        If Text1(0).Text = "" Or Text1(1).Text = "" Then
'           MsgBox "Primero debe indicar un cliente y obra", vbExclamation
'           Exit Sub
'        End If
'        Cad = "Actuacion|sflotasregistro|actuacion|T||25·"
'        Cad = Cad & "Observaciones|sflotasregistro|observa|T||55·"
'
'        frmB.vTabla = "sflotasregistro"
'        frmB.vSQL = "codclien = " & Text1(0).Text & " and coddirec = " & Text1(1).Text
'
'
'         cad1 = "Actuaciones en obra : "
'
'         devuelve = "codclien = " & Text1(0).Text & " AND coddirec "
'         devuelve = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", devuelve, Val(Text1(1)))
'
'
'        cad1 = cad1 & devuelve
'
'        frmB.vTitulo = cad1
        
                
    End Select
    
    HaDevueltoDatos2 = ""
    frmB.vCampos = cad
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 0
    frmB.vCargaFrame = False
    frmB.vConexionGrid = 1
    frmB.Show vbModal
    Set frmB = Nothing
    
    If HaDevueltoDatos2 <> "" Then
        'St op
        Text1(Index).Text = RecuperaValor(HaDevueltoDatos2, 1)
        
        If Index = 1 Then
            PonerDatosVehiculo
        Else
            Text2(Index).Text = RecuperaValor(HaDevueltoDatos2, 2)
        End If
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

Private Sub mnOrden1_Click(Index As Integer)
    mnOrden1(0).Checked = False
    mnOrden1(1).Checked = False
    mnOrden1(2).Checked = False
    mnOrden1(Index).Checked = True
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
    
   
    

    Select Case Index
        Case 0 'PROVEED
            devuelve = ""
            Text1(Index).Text = Trim(Text1(Index).Text)
            
            
            If Text1(Index).Text <> "" Then
                If PonerFormatoEntero(Text1(Index)) Then
                   
                    devuelve = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", Text1(Index).Text)
                    If devuelve = "" Then
                        MsgBox "No existe el vehiculo: " & Text1(Index).Text, vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                Else
                    Text1(Index).Text = ""
                End If
            End If
            Text2(Index).Text = devuelve
        Case 1
            'VEHICULO
            devuelve = ""
            Text1(Index).Text = Trim(Text1(Index).Text)
            
            PonerDatosVehiculo
            
        
        Case 2
            Text1(2).Text = UCase(Text1(2).Text)
        Case 3
            PonerFormatoFecha Text1(Index)
            PonerProximo
        Case 8
            PonerFormatoDecimal Text1(Index), 3

    End Select
    
    If Modo >= 3 Then
        If Me.FrameConKim.visible Then
            If Index = 4 Or Index = 9 Or Index = 11 Then FijarConsumo
        End If
    End If
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
            frmListado3.Opcion = 30
            frmListado3.Show vbModal
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
    BloquearTxt Text1(6), Modo <> 1, True
    BloquearTxt Text1(7), Modo <> 1
    BloquearTxt Text1(9), Modo <> 1
    BloquearTxt Text1(10), Modo <> 1
    
    'Bloquear Registro sino es Insert o Update
    b = (Modo = 0) Or (Modo = 2)

    
           
    '------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    Me.Combo1.Enabled = b
    
    
    
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
   Combo1.ListIndex = -1
   DatosVehiculo = ""
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
    FrameConKim.visible = False
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
    Text1(3).Text = Format(Now, "dd/mm/yyyy")
    
    If Me.DatosADevolverBusqueda <> "" Then
'        Text1(0).Text = RecuperaValor(Me.DatosADevolverBusqueda, 1)
'        Text1(1).Text = RecuperaValor(Me.DatosADevolverBusqueda, 2)
'        Text1_LostFocus 0
'        Text1_LostFocus 1
 
    End If
    PonerFoco Text1(3)
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
    
    SQL = SQL & "¿Seguro que desea eliminar el registro de entrada?" & vbCrLf & vbCrLf
    SQL = SQL & vbCrLf & "Vechiculo : " & Text1(1).Text & " " & Text2(1).Text
    SQL = SQL & vbCrLf & "Concepto : " & Combo1.Text
    SQL = SQL & vbCrLf & "Base imponible : " & Text1(8).Text
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
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
   
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Dirección", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
'Dim SQL As String
On Error GoTo FinEliminar
        
        
        'SQL = SQL & " AND codclien=" & Data1.Recordset!codclien
        'SQL = SQL & " AND actuacion=" & DBSet(Data1.Recordset!actuacion, "T")
        'Cabeceras
        conn.Execute "Delete  from " & NombreTabla & " WHERE registro= " & Data1.Recordset!registro
                      
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
    
    
    'Los Kms actuales NO pueden ser menor que los que ya habian
    If Modo = 3 Then
        If Text1(9).Text <> "" Then
            If Val(Text1(4).Text) < Val(Text1(9).Text) Then
                MsgBox "Km actuales menor que los anteriores", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    
    
    If Modo = 3 Then Text1(6).Text = SugerirCodigoSiguienteStr("sflotasregistro", "registro")
     
    DatosOk = True
End Function


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String

    'Llamamos a al form
    cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: slista
    cad = cad & ParaGrid(Text1(6), 10, "Registro")
    
        cad = cad & "Codigo|sflotas|codflota|T||15·"
        cad = cad & "Desc.|sflotas|nomflota|T||35·"
        cad = cad & "Desc.|sflotatipo|desctipflota|T||30·"
        
        
               
 
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vTabla = "sflotasregistro,sflotas,sflotatipo"
        If cadB <> "" Then cadB = " AND " & cadB
        cadB = "sflotasregistro.codflota=sflotas.codflota AND sflotas.tipo=tipflota " & cadB
        frmB.vSQL = cadB
        frmB.vTitulo = "Registro"
        frmB.vCampos = cad
        HaDevueltoDatos2 = ""
        '###A mano
        frmB.vDevuelve = "0|"
        
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri 'Conexion a BD Ariges

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos2 <> "" Then
            'Creamos una cadena consulta y ponemos los datos
            Screen.MousePointer = vbHourglass
            cadB = ""
            cad = RecuperaValor(HaDevueltoDatos2, 1)
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE registro = " & cad & Ordenacion
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
    
    PonerDatosVehiculo
    Text2(0).Text = PonerNombreDeCod(Text1(0), conAri, "sprove", "nomprove", "codprove", "N")
    
    
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

    vWhere = "registro=" & Text1(6).Text
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
'        LimpiarCampos
        PonerModo 0
    End If
End Sub

'Private Function PuedeEliminar() As Boolean
'Dim C As String
'Dim Aux As String
''    Aux = ""
''    C = " coddirec = " & Text1(1).Text & " AND actuacion = " & DBSet(Text1(2).Text, "T") & " AND codclien "
''    C = DevuelveDesdeBD(conAri, "count(*)", "scapre", C, Text1(0).Text, "N")
''    If C = "" Then C = "0"
''    If Val(C) <> 0 Then Aux = Aux & "   -Ofertas" & vbCrLf
''
''
''    C = " coddirec = " & Text1(1).Text & " AND actuacion = " & DBSet(Text1(2).Text, "T") & " AND codclien "
''    C = DevuelveDesdeBD(conAri, "count(*)", "scaped", C, Text1(0).Text, "N")
''    If C = "" Then C = "0"
''    If Val(C) <> 0 Then Aux = Aux & "   -Pedidos" & vbCrLf
''
''
''    C = " coddirec = " & Text1(1).Text & " AND actuacion = " & DBSet(Text1(2).Text, "T") & " AND codclien "
''    C = DevuelveDesdeBD(conAri, "count(*)", "scaalb", C, Text1(0).Text, "N")
''    If C = "" Then C = "0"
''    If Val(C) <> 0 Then Aux = Aux & "   -Albaranes" & vbCrLf
''
''    'Partes de trabajo
''    C = " coddirec = " & Text1(1).Text & " AND actuacion = " & DBSet(Text1(2).Text, "T") & " AND codclien "
''    C = DevuelveDesdeBD(conAri, "count(*)", "sliparte", C, Text1(0).Text, "N")
''    If C = "" Then C = "0"
''    If Val(C) <> 0 Then Aux = Aux & "   -Partes de trabajo" & vbCrLf
''
'
'    If Aux <> "" Then
'        PuedeEliminar = False
'        Aux = "Existen datos relacionados con la actuacion en : " & vbCrLf & Aux
'        MsgBox Aux, vbQuestion
'    Else
'        PuedeEliminar = True
'    End If
'
'End Function
'


Private Sub CargaComboConceptos()
    ConceptosRequierenKm = "|"
    NombreTabla = "Select  * from sflotasconce order by nomconcef"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open NombreTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Combo1.Clear
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!nomconcef
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!codconcef
        If DBLet(miRsAux!solicitakm, "N") = 1 Then ConceptosRequierenKm = ConceptosRequierenKm & miRsAux!codconcef & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
End Sub


Private Sub PonerDatosVehiculo()


    'A partir de la matricula pondra nombre-tipo....
    ' Kms iniciales .. caducidad
    Text1(9).Text = ""  'Km anteriores
    Text2(2).Text = ""
    DatosVehiculo = ""
    If Text1(1).Text = "" Then
        Text2(1).Text = ""
    Else
        
        cad = "Select  desctipflota, nomflota,tipflota from sflotas,sflotatipo WHERE tipo=tipflota AND codflota = " & DBSet(Text1(1).Text, "T")
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            Text2(1).Text = ""
            
            MsgBox "No existe el vehiculo" & Text1(1).Text, vbExclamation
            Text1(1).Text = ""
        Else
            'OK
            Text2(1).Text = miRsAux.Fields(1) & "  (" & miRsAux.Fields(0) & ")"
            DatosVehiculo = miRsAux!tipflota & "|" 'Tipo de flota el 1
            If Modo = 3 Then
                'Fijaremos los KM anteriores
                miRsAux.Close
                cad = "Select horaskm,litrosulttik FROM sflotasregistro WHERE codflota = " & DBSet(Text1(1).Text, "T")
                cad = cad & " AND codconcef = " & IncidenciaLlenarDeposito   'Llenar deposito
                cad = cad & " ORDER BY fecha desc" 'Cogemos el ultimo tiket
                miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then
                    DatosVehiculo = DatosVehiculo & DBLet(miRsAux.Fields(0), "N") & "|"
                    DatosVehiculo = DatosVehiculo & DBLet(miRsAux.Fields(1), "N") & "|"
                    'Esto lo hace poner proximo
              '      Text1(9).Text = DBLet(miRsAux.Fields(0), "N")
              '      Text2(2).Text = DBLet(miRsAux.Fields(1), "N")
                End If
                
            End If
        End If
        miRsAux.Close
        Set miRsAux = Nothing

    End If
    PonerProximo
End Sub


Private Sub PonerProximo()
    If Modo = 2 Then Exit Sub
    
    Text1(7).Text = "": Text1(10).Text = ""

    If Text1(3).Text = "" Then Exit Sub
    If Combo1.ListIndex < 0 Then Exit Sub
    If DatosVehiculo = "" Then Exit Sub
    'sflotascon_x_tipo tipflota freqKm freqMes
    cad = "Select freqKm ,freqMes FROM sflotascon_x_tipo WHERE tipflota = " & RecuperaValor(DatosVehiculo, 1)
    cad = cad & " AND codconcef = " & Combo1.ItemData(Combo1.ListIndex)
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux!freqMes, "N") > 0 Then Text1(7).Text = DateAdd("m", miRsAux!freqMes, CDate(Text1(3).Text))
        Text1(10).Text = miRsAux!freqKm
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If Combo1.ItemData(Combo1.ListIndex) = IncidenciaLlenarDeposito Then
       
        Text1(9).Text = RecuperaValor(DatosVehiculo, 2)
        Text2(2).Text = RecuperaValor(DatosVehiculo, 3)
    End If
End Sub


Private Sub FijarConsumo()
Dim C As Currency
    On Error Resume Next
    If Text1(4).Text = "" Or Text1(9).Text = "" Or Text1(11).Text = "" Then Exit Sub
    If Val(Text1(11).Text) = 0 Then Exit Sub
    
    C = (Val(Text1(4).Text) - Val(Text1(9).Text)) / Text1(11).Text
    Text1(13).Text = Format(C, FormatoCantidad)
    
End Sub
