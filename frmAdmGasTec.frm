VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAdmGasTec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gastos técnicos"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frmAdmGasTec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   9
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   3
      Tag             =   "Nº viajes|N|S|0|99|sgaste|numviaje|0|N|"
      Text            =   "viaje"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   8760
      MaxLength       =   16
      TabIndex        =   9
      Tag             =   "Importe varios|N|S|0||sgaste|impvario|#,###,###,##0.00|N|"
      Text            =   "Imp varios"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   7560
      MaxLength       =   16
      TabIndex        =   8
      Tag             =   "Importe taxi|N|S|0||sgaste|impotaxi|#,###,###,##0.00|N|"
      Text            =   "Imp taxi"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   6480
      MaxLength       =   16
      TabIndex        =   7
      Tag             =   "Importe parking|N|S|0||sgaste|impparki|#,###,###,##0.00|N|"
      Text            =   "Imp parking"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   5280
      MaxLength       =   16
      TabIndex        =   6
      Tag             =   "Importe autopista|N|S|0||sgaste|impautop|#,###,###,##0.00|N|"
      Text            =   "Imp autopista"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   4200
      MaxLength       =   16
      TabIndex        =   5
      Tag             =   "Importe comida|N|S|0||sgaste|impcomid|#,###,###,##0.00|N|"
      Text            =   "Imp comida"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   2
      Left            =   3000
      TabIndex        =   24
      ToolTipText     =   "Buscar artículo"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   6570
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   5040
      Width           =   4365
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Kilometros|N|S|0|99999|sgaste|kilometr|0|N|"
      Text            =   "kms"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   5400
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
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   18
      ToolTipText     =   "Buscar artículo"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   5040
      Width           =   4245
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   240
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "Fecha|F|N|||sgaste|fecgasto|dd/mm/yyyy|S|"
      Text            =   "fecha"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Cod. Cliente|N|N|0|999999|sgaste|codclien|000000|S|"
      Text            =   "codclien"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   5565
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9675
      TabIndex        =   11
      Top             =   5565
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9675
      TabIndex        =   12
      Top             =   5565
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "Técnico|N|N|0|9999|sgaste|codtecni|0000|S|"
      Text            =   "tecn"
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   9240
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3000
      Top             =   5640
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAdmGasTec.frx":000C
      Height          =   4425
      Left            =   240
      TabIndex        =   13
      Top             =   495
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7805
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   23
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Técnico"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   5040
      Width           =   615
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
      TabIndex        =   15
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
Attribute VB_Name = "frmAdmGasTec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
'Private WithEvents frmF As frmCal 'Calendario de Fechas
Private WithEvents frmT As frmAdmTrabajadores  'Form Mantenimiento Trabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes   'Form Mantenimiento Clientes
Attribute frmC.VB_VarHelpID = -1


Private NombreTabla As String
Private Ordenacion As String
Private Modo As Byte

Dim kCampo As Integer

Dim EsBusqueda As Boolean
'Para cargar el DataGrid con la consulta de busqueda y no con todos los registros

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As Boolean


Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim NumReg As Long

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    
    Select Case Modo
        Case 1 'BUSQUEDA
            HacerBusqueda
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid True
                    BotonAnyadir
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk Then
                 If ModificaDesdeFormulario(Me, 3) Then
                     TerminaBloquear
                     NumReg = Data1.Recordset.AbsolutePosition
                     PonerModo 2
                     CancelaADODC Me.Data1
                     CargaGrid True
                     LLamaLineas 10
                     SituarDataPosicion Data1, NumReg, Indicador
                 End If
                 lblIndicador.Caption = Indicador
                 PonerFocoGrid DataGrid1
            End If
    End Select
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 1 'cod. tecnico
            Set frmT = New frmAdmTrabajadores
            frmT.DatosADevolverBusqueda = "0|1|" 'Poner Modo Busqueda
            frmT.Show vbModal
            Set frmT = Nothing
            
        Case 2 'cod. cliente
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0|1|" 'Poner Modo Busqueda
            frmC.Show vbModal
            Set frmC = Nothing
    End Select
    PonerFoco txtAux(Index)
End Sub


Private Sub cmdCancelar_Click()
Dim Indicador As String

    On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
            LLamaLineas 10
            EsBusqueda = False
           
        Case 3 'Insertar
            DataGrid1.AllowAddNew = False
            DataGrid1.Enabled = True
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                PonerModo 2
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            Else
                PonerModo 0
            End If
            LLamaLineas 10
            
        Case 4  'Modificar
            TerminaBloquear
            NumRegElim = Data1.Recordset.AbsolutePosition
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            PonerModo 2
            LLamaLineas 10
            DataGrid1.Enabled = True
            SituarDataPosicion Data1, NumRegElim, Indicador
            DeseleccionaGrid Me.DataGrid1
            lblIndicador.Caption = Indicador
    End Select
    Exit Sub
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1

    If Not Data1.Recordset.EOF Then
        txtAux(1).Text = DBLet(Data1.Recordset!codtecni, "T")
        txtAux2(1).Text = PonerNombreDeCod(txtAux(1), conAri, "straba", "nomtraba", "codtraba", "N")

        txtAux(2).Text = DBLet(Data1.Recordset!codclien, "T")
        Me.txtAux2(2).Text = PonerNombreDeCod(txtAux(2), conAri, "sclien", "nomclien", "codclien", "N")
        
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
    Exit Sub
    
Error1:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'ICONOS de La toolbar
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
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "sgaste" 'Tabla Gastos Tecnicos
    Ordenacion = " ORDER BY fecgasto desc,codtecni,codclien "
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        'se le llama desde otro form
        BotonBuscar
    End If
    CargaGrid False
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim tots As String
    
    On Error GoTo ECarga
    
    tots = MontaSQLCarga(enlaza)

    DataGrid1.Columns(0).Width = 700
    DataGrid1.Columns(1).Width = 700
    
    CargaGridGnral DataGrid1, Me.Data1, tots, False
    
    tots = "S|txtAux(0)|T|Fecha|1050|;S|txtAux(1)|T|Técnico|850|;S|cmdAux(1)|B||0|;S|txtAux(2)|T|Cliente|850|;S|cmdAux(2)|B||0|;"
    tots = tots & "S|txtAux(9)|T|Viajes|650|;S|txtAux(3)|T|Kms|700|;S|txtAux(4)|T|Imp. Comida|1200|;S|txtAux(5)|T|Imp. Autopista|1350|;S|txtAux(6)|T|Imp. Parking|1200|;S|txtAux(7)|T|Imp. Taxi|1100|;S|txtAux(8)|T|Imp. Varios|1150|;"
    
    arregla tots, DataGrid1, Me

    'dtos alineados a la dcha
    DataGrid1.Columns(3).Alignment = dbgRight
    DataGrid1.Columns(4).Alignment = dbgRight
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(6).Alignment = dbgRight
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(8).Alignment = dbgRight
    DataGrid1.Columns(9).Alignment = dbgRight

    DataGrid1.ScrollBars = dbgAutomatic
    
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0 Or Modo = 2) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
   Exit Sub
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Byte
Dim b As Boolean

    On Error Resume Next
    
    DeseleccionaGrid Me.DataGrid1
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To txtAux.Count - 1
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = b
    Next jj
    
    For jj = 1 To 2
        Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(jj).Top = alto
        Me.cmdAux(jj).visible = b
        Me.cmdAux(jj).Enabled = b And (Modo <> 4)
    Next jj
    
    If Err.Number Then Err.Clear
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Cliente
    txtAux(2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod clien
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom clien
End Sub


'Private Sub frmF_Selec(vFecha As Date)
''Calendario de Fecha
'    txtAux(3).Text = Format(vFecha, "dd/mm/yyyy")
'End Sub


Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Trabajadores
    txtAux(1).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod traba
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom traba
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
Dim SQL As String

    If Me.Data1.Recordset.EOF Then Exit Sub
    
    SQL = " fecgasto=" & DBSet(Me.Data1.Recordset!fecgasto, "F") & " AND codtecni=" & Data1.Recordset!codtecni
    SQL = SQL & " AND codclien=" & Data1.Recordset!codclien
    If BloqueaRegistro(NombreTabla, SQL) Then BotonModificar
    
'     If BLOQUEADesdeFormulario(Me) Then BotonModificar
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
            AbrirListadoOfer (92)
            
        Case 11  'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim i As Byte
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
     'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
                      
    'Modo Buscar
    If Kmodo = 1 Then PonerFoco txtAux(0)
                      
    'Bloquear los campos de clave primaria al modificar
    For i = 0 To 2
        BloquearTxt txtAux(i), (Modo = 4)
    Next i
                      
    '-----------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    PonerModoOpcionesMenu  'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
      PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
    
    On Error Resume Next

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
    
    b = ((Modo >= 3))
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo opciones del menú.", Err.Description

End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
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
    
    On Error GoTo ErrSQL
    
    SQL = "SELECT fecgasto, codtecni, " & NombreTabla & ".codclien,numviaje,kilometr,impcomid,impautop,impparki,impotaxi,impvario "
    SQL = SQL & " FROM " & NombreTabla
'    SQL = SQL & " INNER JOIN sclien ON " & NombreTabla & ".codclien = sclien.codclien "
        
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then
            SQL = SQL & CadenaBusqueda
        Else
            If Modo = 3 Then
                If Data1.Recordset.RecordCount < 1 Then SQL = SQL & " WHERE fecgasto=" & DBSet(txtAux(0).Text, "F")
            End If
        End If
    Else
        SQL = SQL & " WHERE codtecni = -1"
    End If
    
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
    Exit Function
    
ErrSQL:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cadena SQL", Err.Description
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    LimpiarCampos
    
    If Modo <> 1 Then
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        anc = ObtenerAlto(Me.DataGrid1, 10)
        LLamaLineas anc
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            txtAux(kCampo).Text = ""
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    On Error Resume Next

    EsBusqueda = False
    LimpiarCampos
    
    CadenaConsulta = MontaSQLCarga(True)
    PonerCadenaBusqueda
    PonerFocoGrid Me.DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
   
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    txtAux(0).Text = Format(Now, "dd/mm/yyyy")
    txtAux(9).Text = "1"
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
Dim i As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    On Error Resume Next
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc

    'poner valores grabados
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "F")
    txtAux(1).Text = DBLet(DataGrid1.Columns(1).Value, "N")
    txtAux(2).Text = DBLet(DataGrid1.Columns(2).Value, "N")
    
    For i = 3 To 8
        txtAux(i).Text = DBLet(Data1.Recordset.Fields(i + 1).Value, "T")
    Next i
    txtAux(9).Text = DBLet(Data1.Recordset!numviaje, "T")
    
    FormateaCampo txtAux(1)
    FormateaCampo txtAux(2)
    For i = 4 To 8
        FormateaCampo txtAux(i)
    Next i

    DataGrid1.Enabled = False
    PonerFoco txtAux(3)
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Botón modificar.", Err.Description
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String
        
    On Error GoTo FinEliminar
        
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Function
        
    SQL = "¿Seguro que desea eliminar los Gastos técnicos?" & vbCrLf
    SQL = SQL & vbCrLf & "Fecha: " & Data1.Recordset.Fields(0).Value
    SQL = SQL & vbCrLf & "Técnico: " & Format(Data1.Recordset.Fields(1).Value, "0000") & " - " & txtAux2(1).Text
    SQL = SQL & vbCrLf & "Cliente: " & Format(Data1.Recordset.Fields(2).Value, "000000") & " - " & txtAux2(2).Text
            
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Me.Data1.Recordset.AbsolutePosition
        SQL = "Delete from " & NombreTabla & " where fecgasto=" & DBSet(Data1.Recordset!fecgasto, "F")
        SQL = SQL & " AND codtecni=" & Data1.Recordset!codtecni & " AND codclien=" & Data1.Recordset!codclien
        
        conn.Execute SQL
        CancelaADODC Me.Data1
        CargaGrid True
        CancelaADODC Me.Data1
        SituarDataPosicion Me.Data1, NumRegElim, SQL
    End If
    Exit Function
        
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Gastos Técnicos.", Err.Description
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String

    DatosOk = False
    b = CompForm(Me, 3)
    If Not b Then Exit Function
    
    'Comprobar que existe un Albaran de venta para ese técnico(realizado por del alb)
    'para ese cliente y en esa fecha. Si no avisar
    SQL = "SELECT count(*) FROM scaalb WHERE "
    SQL = SQL & " fechaalb=" & DBSet(txtAux(0).Text, "F") & " AND codtraba=" & txtAux(1).Text
    SQL = SQL & " AND codclien=" & txtAux(2).Text & " AND codtipom='ALV'"
    
    If Not (RegistrosAListar(SQL) > 0) Then
        SQL = "No existe un Albaran de fecha: " & txtAux(0).Text
        SQL = SQL & " para ese técnico y cliente." & vbCrLf
        SQL = SQL & "¿Desea continuar?"
        If MsgBox(SQL, vbYesNo) = vbNo Then b = False
    End If
    
    DatosOk = b
End Function


Private Sub HacerBusqueda()
Dim cadB As String

    On Error Resume Next

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
'        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then 'Se muestran en el mismo form
'        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
        CadenaBusqueda = " WHERE " & cadB
        CadenaConsulta = MontaSQLCarga(True)
        PonerCadenaBusqueda
        PonerFocoGrid Me.DataGrid1
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerCadenaBusqueda()
Dim cad As String

    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
      

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        cad = "No hay ningún registro en la tabla " & NombreTabla
         If EsBusqueda Then cad = cad & " para ese criterio de Búsqueda."
        MsgBox cad, vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        Exit Sub
    Else
        PonerModo 2
        CargaGrid True
'        DataGrid1.Refresh
'        PonerCampos
    End If
    LLamaLineas 10
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            If Index > 0 Then PonerFoco txtAux(Index - 1)
                
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    On Error Resume Next
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'Fecha
            PonerFormatoFecha txtAux(Index)
        
        Case 1 'cod tecnico
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "straba", "nomtraba", "codtraba")
                If txtAux2(Index).Text = "" Then PonerFoco txtAux(Index)
            Else
                txtAux2(Index).Text = ""
            End If
            
        Case 2 'Cod. clien
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), conAri, "sclien", "nomclien", "codclien")
                If txtAux2(Index).Text = "" Then
                    PonerFoco txtAux(Index)
                ElseIf Modo = 3 Then 'Insertar
                    txtAux(3).Text = DevuelveDesdeBDNew(conAri, "sclien", "kilometr", "codclien", txtAux(Index).Text, "N")
                End If
            Else
                txtAux2(Index).Text = ""
            End If

        Case 3 'kms
            PonerFormatoEntero txtAux(Index)
            
        Case 4 To 8 'Importes
             PonerFormatoDecimal txtAux(Index), 1
    End Select
    
    If Err.Number <> 0 Then MuestraError Err.Number, "", Err.Description
End Sub

