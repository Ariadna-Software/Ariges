VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmagrupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento comunicaci�n datos ALMAGRUPO "
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmAlmagrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTipoRegistro 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   375
      Left            =   2760
      TabIndex        =   27
      Top             =   3240
      Width           =   3135
      Begin VB.OptionButton Option1 
         Caption         =   "Stocks"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   29
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Consumos"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame FrameConsumo2 
      Caption         =   "Consumo. Es invisible"
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   8175
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   120
         MaxLength       =   16
         TabIndex        =   5
         Tag             =   "Prov.compra|T|S|||salmagrupo|cifproveedor|||"
         Text            =   "Text1"
         Top             =   405
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Uds|N|S|||salmagrupo|udscompra|#,##0.00||"
         Text            =   "Text1"
         Top             =   405
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Importe|N|S|||salmagrupo|importe|#,##0.00||"
         Text            =   "Text1"
         Top             =   405
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   1680
         MaxLength       =   55
         TabIndex        =   6
         Tag             =   "Proveedor|T|S|||salmagrupo|nomproveedor|||"
         Text            =   "Text1"
         Top             =   405
         Width           =   3525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1680
         Picture         =   "frmAlmagrupo.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor compra"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Uds"
         Height          =   255
         Index           =   6
         Left            =   5400
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         Height          =   255
         Index           =   7
         Left            =   6720
         TabIndex        =   24
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   1800
      MaxLength       =   55
      TabIndex        =   22
      Tag             =   "Proveedor|T|S|||salmagrupo|nomprovhabitual|||"
      Text            =   "Text1"
      Top             =   1725
      Width           =   4005
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   840
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1020
      TabIndex        =   1
      Tag             =   "mes|N|N|1|12|salmagrupo|mes|00|S|"
      Text            =   "Text"
      Top             =   840
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Stock|N|N|||salmagrupo|stock|#,##0.00||"
      Text            =   "Text1"
      Top             =   1725
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   180
      MaxLength       =   16
      TabIndex        =   3
      Tag             =   "Prov. habitual|T|N|||salmagrupo|cifprovhabitual|||"
      Text            =   "Text1"
      Top             =   1725
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1800
      MaxLength       =   16
      TabIndex        =   2
      Tag             =   "Articulo|T|N|||salmagrupo|codartic||S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Tag             =   "A�o|N|N|1900|2200|salmagrupo|anyo|00|S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   12
      Top             =   3195
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
         Left            =   120
         TabIndex        =   13
         Top             =   210
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   3360
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   3360
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6720
      Top             =   480
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
      TabIndex        =   16
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Comunicar datos"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar datos para el proceso"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6840
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   2400
      Picture         =   "frmAlmagrupo.frx":010E
      Tag             =   "-1"
      ToolTipText     =   "Buscar poblaci�n"
      Top             =   600
      Width           =   240
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8160
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1800
      Picture         =   "frmAlmagrupo.frx":0210
      Tag             =   "-1"
      ToolTipText     =   "Buscar poblaci�n"
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Stock"
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   20
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor habitual"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   19
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Art�culo"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   18
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "A�o"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Mes"
      Height          =   255
      Index           =   0
      Left            =   1020
      TabIndex        =   14
      Top             =   600
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
Attribute VB_Name = "frmAlmagrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents FrmArt As frmBasico2
Attribute FrmArt.VB_VarHelpID = -1
Private WithEvents frmP As frmBasico2 '%=%=frmComProveedores
Attribute frmP.VB_VarHelpID = -1

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

Private CadenaConsulta2 As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1





Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Select Case Modo
        Case 3  'INSERTAR
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
    If MsgBox("No deberia insertar los consumos por aqu�. �Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    LimpiarCampos
    PonerModo 3
   
    PonerFoco Text1(1)

End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(1)
        Text1(1).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Function DevuelveCadenaBusquedaTipo() As String
    
    If Me.Option1(1).Value Then
        DevuelveCadenaBusquedaTipo = "="
    Else
        DevuelveCadenaBusquedaTipo = "<>"
    End If
    DevuelveCadenaBusquedaTipo = "cifproveedor " & DevuelveCadenaBusquedaTipo & "'S'"


End Function
Private Sub BotonVerTodos()

    'Ver todos
    'las ventas tendran como CIF proveedor el de la fra venta
    'si en el campo CIF viene un 'S' significa que es stock
    
    
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia DevuelveCadenaBusquedaTipo
    Else
        LimpiarCampos
        CadenaConsulta2 = "Select * from " & NombreTabla & " WHERE " & DevuelveCadenaBusquedaTipo & Ordenacion
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
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()


    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    

    
  
  
        CadenaConsulta2 = "�Seguro que desea eliminar el articulo-mes-anyo? " & vbCrLf
        CadenaConsulta2 = CadenaConsulta2 & vbCrLf & "Articulo: " & Data1.Recordset.Fields(0)
        CadenaConsulta2 = CadenaConsulta2 & vbCrLf & "Descripci�n: " & Text3(2).Text
        CadenaConsulta2 = CadenaConsulta2 & vbCrLf & "mes-anyo: " & Format(Data1.Recordset!mes, "00/") & Data1.Recordset!Anyo
        If MsgBox(CadenaConsulta2, vbQuestion + vbYesNo) = vbNo Then Exit Sub

        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        CadenaConsulta2 = "Delete from " & NombreTabla & " where mes=" & Data1.Recordset!mes
        CadenaConsulta2 = CadenaConsulta2 & " AND anyo=" & Data1.Recordset!Anyo
        CadenaConsulta2 = CadenaConsulta2 & " AND codartic=" & DBSet(Data1.Recordset!codArtic, "T")
        conn.Execute CadenaConsulta2
        
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If


    
Error2:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Agente Comercial", Err.Description
     CadenaConsulta2 = ""
End Sub

'
'Private Sub cmdRegresar_Click()
'Dim Cad As String
'
'    If Data1.Recordset.EOF Then
'        MsgBox "Ning�n registro devuelto.", vbExclamation
'        Exit Sub
'    End If
'
'    Cad = Data1.Recordset.Fields(0) & "|"
'    Cad = Cad & Data1.Recordset.Fields(1) & "|"
'    RaiseEvent DatoSeleccionado(Cad)
'    Unload Me
'End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 20
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Buscar
        .Buttons(2).Image = 2   'Bot�n Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        
        .Buttons(9).Image = 40   'immprimir
        .Buttons(10).Image = 22   'exsportar
        .Buttons(12).Image = 42   'generar
        
        .Buttons(17).Image = 15  'Salir
        
        
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
    End With
    
    LimpiarCampos
    
    FrameConsumo2.BorderStyle = vbBSNone
    FrameConsumo2.Top = 2280
    
    '## A mano
    NombreTabla = "salmagrupo"
    Ordenacion = " ORDER BY anyo,mes,codartic"
        
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario

    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where mes=-1"
    Data1.Refresh
   ' If DatosADevolverBusqueda = "" Then
        PonerModo 0
   ' Else
   '     PonerModo 1
   '     Text1(0).BackColor = vbYellow
   ' End If
    
    

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

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta2 = CadenaSeleccion
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
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
        cadB = cadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
        cadB = cadB & " AND " & Aux
        
        Aux = DevuelveCadenaBusquedaTipo
        cadB = cadB & " AND " & Aux
        
        CadenaConsulta2 = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub
    



Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select nifprove,nomprove from sprove where codprove=" & RecuperaValor(CadenaSeleccion, 1), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then CadenaConsulta2 = miRsAux!nifProve & "|" & miRsAux!nomprove & "|"
    miRsAux.Close
    Set miRsAux = Nothing
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    CadenaConsulta2 = ""
    If Index = 2 Then
        'Articulos
        Set FrmArt = New frmBasico2
        'FrmArt.DatosADevolverBusqueda3 = "@1@" 'Poner en Modo busqueda
'        FrmArt.DesdeTPV = False
'        FrmArt.Show vbModal
        AyudaArticulos FrmArt, Text1(2)
        Set FrmArt = Nothing
        If CadenaConsulta2 <> "" Then
            Me.Text1(2).Text = RecuperaValor(CadenaConsulta2, 1)
            Me.Text3(2).Text = RecuperaValor(CadenaConsulta2, 2)
            FijarDatosProvHabitual
        End If
        
    
    Else
        'Proveedor
'        Set frmP = New frmComProveedores
'        frmP.DatosADevolverBusqueda = "0|1|"
'        frmP.Show vbModal
        Set frmP = New frmBasico2
        AyudaProveedores frmP, IIf(Index = 0, Text1(3), Text1(5))
        Set frmP = Nothing
        If CadenaConsulta2 <> "" Then
            If Index = 0 Then
                Me.Text1(3).Text = RecuperaValor(CadenaConsulta2, 1)
                Me.Text1(8).Text = RecuperaValor(CadenaConsulta2, 2)
            Else
                Me.Text1(5).Text = RecuperaValor(CadenaConsulta2, 1)
                Me.Text1(8).Text = RecuperaValor(CadenaConsulta2, 2)
            End If
        End If
            
    End If
    
    CadenaConsulta2 = ""
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
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Option1_Click(Index As Integer)
    If Modo > 2 Then Exit Sub
    If Me.Option1(0).Value Then
        FrameConsumo2.Top = 2280
    Else
        FrameConsumo2.Top = 12280
    End If
    Me.FrameConsumo2.Enabled = Me.Option1(0).Value
    If Modo = 2 Then BotonBuscar

    
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

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0, 1
            If Not PonerFormatoEntero(Text1(Index)) Then Text1(Index).Text = ""
               
            

        Case 2
            'codartic
            devuelve = ""
            If Me.Text1(Index).Text <> "" Then
                devuelve = PonerNombreDeCod(Text1(Index), conAri, "sartic", "nomartic", "codartic", "Articulos", "T")
                If devuelve = "" Then PonerFoco Text1(Index)
    
            End If
            Text3(Index).Text = devuelve
            'Datos del proveedor habitual
            If devuelve <> "" Then FijarDatosProvHabitual
                
                
        Case 3, 5 'NIF prove
            devuelve = ""
            If Me.Text1(Index).Text <> "" Then
                devuelve = PonerNombreDeCod(Text1(Index), conAri, "sprove", "nomprove", "nifprove", "Proveedores", "T")
                If devuelve = "" Then PonerFoco Text1(Index)
            End If
            If Index = 3 Then
                Text1(8).Text = devuelve
            Else
                Text1(9).Text = devuelve
            End If
            
        Case 4, 6, 7
            'Tipo 4: Decimal(4,2)
            If Not PonerFormatoDecimal(Text1(Index), 3) Then
               Text1(Index).Text = ""
            End If
            
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If cadB <> "" Then cadB = cadB & " AND " & DevuelveCadenaBusquedaTipo
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then     'Se muestran en el mismo form
        CadenaConsulta2 = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
Dim cad As String
        'Llamamos a al form
        '##A mano
        cad = ""
        cad = cad & ParaGrid(Text1(1), 7, "A�o")
        cad = cad & ParaGrid(Text1(0), 10, "Mes")
        cad = cad & ParaGrid(Text1(2), 25, "Articulo")
        cad = cad & "Descripcion|sartic|nomartic|T||50�"
        '"Cod Diag.|tabla|columna|tipo|formato|10
        
        If cadB <> "" Then cadB = " AND " & cadB
        cadB = NombreTabla & ".codartic = sartic.codartic" & cadB
        If cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = cad
            frmB.vTabla = NombreTabla & ",sartic"
            frmB.vSQL = cadB
            HaDevueltoDatos = False
            '###A mano

            frmB.vDevuelve = "0|1|2|" 'Campos de la tabla que devuelve
            frmB.vTitulo = "Datos ALMAGRUPO"
            frmB.vselElem = 1
            frmB.vConexionGrid = conAri 'Conexi�n a BD: Ariges
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                'PonerFocoBtn Me.cmdRegresar
               
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerModo Modo
                PonerFoco Text1(kCampo)
            End If
        End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta2
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de B�squeda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
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


Private Sub PonerCampos()
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    Text3(2).Text = PonerNombreDeCod(Text1(2), conAri, "sartic", "nomartic", "codartic", "Articulos", "T")
    
    'No hace falta. El nomprove va guardado
'    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "sprove", "nomprove", "nifprove", "Proveedores", "T")
'    Text2(5).Text = PonerNombreDeCod(Text1(5), conAri, "sprove", "nomprove", "nifprove", "Proveedores", "T")
'
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

    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
   ' If DatosADevolverBusqueda <> "" Then
   '     cmdRegresar.visible = b
   ' Else
        cmdRegresar.visible = False
   ' End If
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    b = Modo = 1 Or Modo >= 3 'busqueda o inser/mod
    
    'Nombre proveedor
    b = Modo <> 1
    BloquearTxt Me.Text1(8), b
    BloquearTxt Me.Text1(9), b

    '---------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = b Or Modo = 1
    cmdCancelar.visible = b Or Modo = 1
    
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    FrameTipoRegistro.Enabled = Not b
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub



Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    
    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    mnEliminar.Enabled = b
    
    '----------------------------------------
    b = (Modo >= 3) 'Insertar/Modificar
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cad As String

    DatosOk = False
    b = CompForm(Me, 1) 'Comprobar datos OK
    If Not b Then Exit Function
        
        
    'Que exista NOMARTIC
    CadenaConsulta2 = ""
    If Text3(2).Text = "" Then CadenaConsulta2 = "-Articulo incorrecto"
    
    If Text1(8).Text = "" Then CadenaConsulta2 = CadenaConsulta2 & "-Proveedor habitual incorrecto" & vbCrLf
    If Me.Option1(0).Value Then
        If Text1(5).Text = "" Xor Text1(9).Text = "" Then CadenaConsulta2 = CadenaConsulta2 & "-Proveedor compra incorrecto" & vbCrLf
    Else
        Text1(5).Text = ""
    End If
    
    
    
    If Text1(5).Text <> "" Then
        'Es compra.
        'Campos obligados
        If Text1(6).Text = "" Then CadenaConsulta2 = CadenaConsulta2 & "-Compra: Unidades" & vbCrLf
        If Text1(7).Text = "" Then CadenaConsulta2 = CadenaConsulta2 & "-Compra: importe" & vbCrLf
    Else
        'NO es compra. No debe poner valor para uds
        
        If Text1(6).Text <> "" Then CadenaConsulta2 = CadenaConsulta2 & "-Unidades no debe tener valor" & vbCrLf
        If Text1(7).Text <> "" Then CadenaConsulta2 = CadenaConsulta2 & "-Importe no debe tener valor" & vbCrLf
        If CadenaConsulta2 = "" Then
            Text1(5).Text = "S"  'stock
            Text1(9).Text = ""   'nom
        End If
    End If
    
    
    
    
    If CadenaConsulta2 <> "" Then
        MsgBox CadenaConsulta2, vbExclamation
        b = False
        CadenaConsulta2 = ""
    End If
    
    
    'Comprobaciones
    cad = "codtelem"
    CadenaConsulta2 = "sartic.codfamia=sfamia.codfamia AND codartic"
    CadenaConsulta2 = DevuelveDesdeBD(conAri, "comunica", "sartic,sfamia", CadenaConsulta2, Text1(2).Text, "T", cad)
    If CadenaConsulta2 = "" Then CadenaConsulta2 = "0"
    If Val(CadenaConsulta2) <> 1 Then
        CadenaConsulta2 = "Famila NO se comunica"
    Else
        CadenaConsulta2 = ""
    End If
    If cad = "" Then CadenaConsulta2 = CadenaConsulta2 & vbCrLf & "No tiene codigo telematel"
    If CadenaConsulta2 <> "" Then
        MsgBox CadenaConsulta2, vbExclamation
        b = False
        CadenaConsulta2 = ""
    End If
    
    DatosOk = b
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
            
            
        Case 10
                If Modo <> 2 And Modo <> 0 Then Exit Sub
                frmListado3.Opcion = 29
                frmListado3.Show vbModal
        Case 12
            'Generar
            GenerarDatosMes
        
        Case 17  'Salir
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
Dim vWhere As String
Dim Indicador As String

    
    vWhere = DBSet(Text1(5).Text, "T")
    vWhere = "mes = " & Text1(0).Text & " and anyo = " & Val(Text1(1).Text) & " and  codartic = " & DBSet(Text1(2).Text, "T") & " AND cifproveedor = " & vWhere
    If Modo = 3 Then Data1.RecordSource = "select * from " & NombreTabla & " WHERE " & vWhere & " " & Ordenacion
    
    
    If SituarDataMULTI(Data1, vWhere, Indicador) Then
    
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        LimpiarCampos
        PonerModo 0
    End If
End Sub


'Private Function ObtenerWhereCP() As String
'Dim SQL As String
'On Error Resume Next
'    SQL = " WHERE codagent= " & Text1(0).Text
'    ObtenerWhereCP = SQL
'End Function






Private Sub GenerarDatosMes()
Dim FI As Date
    
    If Modo <> 2 And Modo <> 0 Then Exit Sub
    Set miRsAux = New ADODB.Recordset
    CadenaConsulta2 = "select * from salmagrupoparam "
    miRsAux.Open CadenaConsulta2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        CadenaConsulta2 = "01/" & Format(miRsAux!ultmes, "00") & "/" & miRsAux!ultanyo
        FI = CDate(CadenaConsulta2)
        FI = DateAdd("m", 1, FI) 'el mes siguiente
    Else
        MsgBox "Error ultimo periodo generado", vbExclamation
        CadenaConsulta2 = ""
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If CadenaConsulta2 = "" Then Exit Sub
    
    frmListado3.Opcion = 28
    frmListado3.OtrosDatos = FI
    frmListado3.Show vbModal
    
    'Cargaremos datos
    If CadenaDesdeOtroForm <> "" Then
        'Ha generado
        CadenaConsulta2 = "mes = " & Month(FI) & " AND anyo =" & Year(FI)
        CadenaConsulta2 = "select * from " & NombreTabla & " WHERE " & CadenaConsulta2 & " " & Ordenacion
        PonerCadenaBusqueda
                
    End If
    
    

End Sub



Private Sub FijarDatosProvHabitual()
Dim D As String
Dim D2 As String
    D2 = "nifprove"
    D = "sartic.codprove=sprove.codprove AND codartic"
    D = DevuelveDesdeBD(conAri, "nomprove", "sartic,sprove", D, Text1(2).Text, "T", D2)
    Text1(3).Text = D2
    Text1(8).Text = D
End Sub
