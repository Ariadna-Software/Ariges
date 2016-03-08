VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmFacTPVParamT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros terminales TPV"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10050
   Icon            =   "frmFacTPVParamT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Desc. terminal|T|S|||spatpvt|destermi||N|"
   Begin VB.TextBox txtAux 
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   6600
      MaxLength       =   3
      TabIndex        =   4
      Tag             =   "COM|N|N|1|100|spatpvt|CualCOM|0|N|"
      Text            =   "COM"
      Top             =   3480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      BackColor       =   &H8000000B&
      Height          =   405
      Index           =   4
      Left            =   480
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   21
      Tag             =   "Impresora|T|S|||spatpvt|nomimpre||N|"
      Text            =   "impreso"
      Top             =   4440
      Width           =   4635
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   4680
      MaxLength       =   7
      TabIndex        =   3
      Tag             =   "Contador|N|N|0||spatpvt|contador|0|N|"
      Text            =   "contado"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3120
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Nombre PC|T|N|||spatpvt|nombrepc||N|"
      Text            =   "nom pc"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   1500
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "Desc. terminal|T|S|||spatpvt|destermi||N|"
      Text            =   "desc. termi"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   360
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Nº terminal|N|N|0|999|spatpvt|numtermi|000|S|"
      Text            =   "ter"
      Top             =   3480
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.Frame FrameVisor 
      Caption         =   " Visor "
      Height          =   1695
      Left            =   5400
      TabIndex        =   16
      Top             =   4080
      Width           =   4215
      Begin VB.CheckBox chkAbreCajon 
         Caption         =   "Abre cajón"
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Tag             =   "Cajon|N|S|||spatpvt|abrecajon|||"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboPuertos 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Nº Puerto|N|S|||spatpvt|numpuerto|||"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkHayVisor 
         Caption         =   "Utiliza visor "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "Hay visor|N|S|||spatpvt|hayvisor|||"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboVelocidad 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   17
         Tag             =   "Velocidad Puerto|N|S|||spatpvt|velocpue|||"
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Puerto visor"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Velocidad (Baudios)"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame FrameImpresoras 
      Caption         =   " Impresora "
      Height          =   1695
      Left            =   240
      TabIndex        =   15
      Top             =   4080
      Width           =   5055
      Begin VB.CommandButton cmdConfigImpre 
         Caption         =   "Sel. &impresora..."
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2400
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   360
      TabIndex        =   13
      Top             =   5880
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   210
         Width           =   2280
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8595
      TabIndex        =   9
      Top             =   6045
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   6045
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8595
      TabIndex        =   11
      Top             =   6045
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3600
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Bindings        =   "frmFacTPVParamT.frx":000C
      Height          =   3165
      Left            =   240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   600
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5583
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   1
         Shortcut        =   ^N
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
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacTPVParamT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmMtoCli As frmFacClientes
Attribute frmMtoCli.VB_VarHelpID = -1
Private WithEvents frmMtoBanPr As frmFacBancosPropios
Attribute frmMtoBanPr.VB_VarHelpID = -1

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
'4: Modificar

Dim PrimeraVez As Boolean





Private Sub cboPuertos_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboVelocidad_Click()
    txtAux(5).Text = Me.cboVelocidad.List(Me.cboVelocidad.ListIndex)
End Sub

Private Sub chkBasesImp_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkBasesImp_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkBasesImp_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cboVelocidad_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkHayVisor_Click()
    If (Modo <> 3 And Modo <> 4) Then Exit Sub
    If Me.chkHayVisor.Value = 1 Then
        Me.cboPuertos.Enabled = True
        Me.cboVelocidad.Enabled = True
    Else
        Me.cboPuertos.Enabled = False
        Me.cboVelocidad.Enabled = False
        Me.cboPuertos.ListIndex = -1
        Me.cboVelocidad.ListIndex = -1
    End If
End Sub

Private Sub chkImprtick_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkImprtick_KeyDown(KeyCode As Integer, Shift As Integer)
     KEYdown KeyCode
End Sub

Private Sub chkImprtick_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub



Private Sub chkHayVisor_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Indicador As String

    On Error GoTo Error1
    
    Select Case Modo
        Case 3 'Insertar
             If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid ""
                    BotonAnyadir
                End If
             End If
             
        Case 4 'Modificar
            If DatosOk Then
                txtAux(2).Text = UCase(txtAux(2).Text)
                If ModificaDesdeFormulario(Me, 3) Then
                    TerminaBloquear
                    NumRegElim = Data1.Recordset.AbsolutePosition
                    PonerModo 2
                    CancelaADODC Me.Data1
                    CargaGrid ""
                    LLamaLineas 5
                    SituarDataPosicion Data1, NumRegElim, Indicador
                    lblIndicador.Caption = Indicador
                    PonerFocoGrid DataGrid1
'                    DataGrid1.SetFocus
'                    'PonerModo 2
'                    PonerFocoBtn Me.cmdSalir
                End If
            End If
    End Select
    DataGrid1.Enabled = True
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
Dim Indicador As String

    On Error GoTo ECancelar
    
    Select Case Modo
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
            LLamaLineas 5
            

        Case 4 'Modificar
            TerminaBloquear
            NumRegElim = Data1.Recordset.AbsolutePosition
            If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
            PonerModo 2
            LLamaLineas 5
            DataGrid1.Enabled = True
            SituarDataPosicion Data1, NumRegElim, Indicador
            DeseleccionaGrid Me.DataGrid1
            lblIndicador.Caption = Indicador
    End Select


'    TerminaBloquear
'    If Data1.Recordset.EOF Then
'        PonerModo 0
'        LimpiarCampos
'    Else
'        PonerCampos
'        PonerModo 2
'    End If
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdConfigImpre_Click()
    Screen.MousePointer = vbHourglass
    'Me.CommonDialog1.Flags = cdlPDPageNums
    CommonDialog1.ShowPrinter
    txtAux(4).Text = PonerNombreImpresora
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub









Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error1

    If Not Data1.Recordset.EOF Then
        txtAux(4).Text = DBLet(Data1.Recordset!NomImpre, "T")
        Me.chkHayVisor = Data1.Recordset!HayVisor
        Me.chkAbreCajon = Data1.Recordset!AbreCajon
        Me.cboPuertos.ListIndex = DBLet(Data1.Recordset!NumPuerto, "N") - 1

        txtAux(5).Text = DBLet(Data1.Recordset!velocpue, "N")
        PosicionarComboDes Me.cboVelocidad, txtAux(5).Text
        
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
    
Error1:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
       
        If Not Data1.Recordset.EOF Then DataGrid1_RowColChange 0, 0
    End If
    PrimeraVez = False
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Eliminar
        .Buttons(6).Image = 15  'Salir
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
'    Me.SSTab1.Tab = 0


    NombreTabla = "spatpvt"
    Ordenacion = " ORDER BY numtermi"
    'ASignamos un SQL al DATA1
'    Data1.ConnectionString = Conn
    CadenaConsulta = "Select * from " & NombreTabla
'    Data1.RecordSource = CadenaConsulta
'    Data1.Refresh
    
    
'    Label1(3).Caption = PonerNombreImpresora
    CargaComboPuertos
    CargaComboVelocidad
    
    PonerModo 0
    
    CargaGrid ""
    Screen.MousePointer = vbDefault

End Sub




Private Sub CargaGrid(SQL As String)
'Dim SQL As String
Dim tots As String
    
    On Error GoTo ECarga
    
'    SQL = MontaSQLCarga(enlaza)
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & Ordenacion

'    DataGrid1.Columns(0).Width = 700
'    DataGrid1.Columns(1).Width = 700
    DataGrid1.ScrollBars = dbgNone
    
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez
    
    tots = "S|txtAux(0)|T|NºTerm.|800|;S|txtAux(1)|T|Desc. terminal|1700|;S|txtAux(2)|T|Nombre PC|1600|;"
    tots = tots & "S|txtAux(3)|T|contador|1200|;N||||0|;N||||0|;N||||0|;N||||0|;"
    tots = tots & "N||||0|;N||||0|;S|txtAux(6)|T|Puerto COM|1200|;"  'abrecajon y secuencia cajon y puerto com
    arregla tots, DataGrid1, Me

    'dtos alineados a la dcha
    DataGrid1.Columns(3).Alignment = dbgRight


    DataGrid1.ScrollBars = dbgAutomatic
    
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0 Or Modo = 2) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub







Private Sub CargaComboPuertos()
    'Carga la lista del Combo de Puertos
    Me.cboPuertos.Clear
  
    
    cboPuertos.AddItem "COM1"
    cboPuertos.ItemData(cboPuertos.NewIndex) = 1
    
    cboPuertos.AddItem "COM2"
    cboPuertos.ItemData(cboPuertos.NewIndex) = 2
    
    cboPuertos.AddItem "COM3"
    cboPuertos.ItemData(cboPuertos.NewIndex) = 3
    
    cboPuertos.AddItem "COM4"
    cboPuertos.ItemData(cboPuertos.NewIndex) = 4
    
    cboPuertos.AddItem "COM5"
    cboPuertos.ItemData(cboPuertos.NewIndex) = 5
    
'    cboPuertos.ListIndex = 1
End Sub




Private Sub CargaComboVelocidad()
    'Carga la lista del Combo de Puertos
    Me.cboVelocidad.Clear
    
    cboVelocidad.AddItem "9600"
    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 0
    
    cboVelocidad.AddItem "14400"
    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 1
    
    cboVelocidad.AddItem "19200"
    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 2
    
    cboVelocidad.AddItem "28800"
    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 3
    
    cboVelocidad.AddItem "38400"
    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 4
    
    cboVelocidad.AddItem "56000"
    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 5
    
    cboVelocidad.AddItem "128000"
    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 6
    
    cboVelocidad.AddItem "256000"
    cboVelocidad.ItemData(cboVelocidad.NewIndex) = 7
'    cboPuertos.ListIndex = 1
End Sub





Private Sub PonerCadenaBusqueda()
    On Error GoTo EEPonerBusq
    
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Me.Toolbar1.Buttons(1).Enabled = False 'Modificar
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Byte
Dim b As Boolean

        DeseleccionaGrid Me.DataGrid1
        b = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

        For jj = 0 To 6
            If jj < 4 Or jj > 5 Then
                txtAux(jj).Height = DataGrid1.RowHeight
                txtAux(jj).Top = alto
                txtAux(jj).visible = b
            End If
        Next jj
        
End Sub




Private Sub BotonAnyadir()
Dim anc As Single

    LimpiarCampos 'Vacía los TextBox
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data1
   
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    anc = ObtenerAlto(Me.DataGrid1, 5)
    LLamaLineas anc
'    txtAux(0).Text = Format(Now, "dd/mm/yyyy")
'    txtAux(9).Text = "1"
    txtAux(0).Text = SugerirCodigoSiguienteStr("spatpvt", "numtermi")
    txtAux(0).Text = Format(txtAux(0).Text, "000")
    PonerFoco txtAux(0)
End Sub


Private Function BotonEliminar() As Boolean
Dim SQL As String
        
        On Error GoTo FinEliminar
        
        'Ciertas comprobaciones
        If Data1.Recordset.EOF Then Exit Function
        
        SQL = "¿Seguro que desea eliminar el terminal?" & vbCrLf
        SQL = SQL & vbCrLf & "Term.: " & Format(Data1.Recordset.Fields(0).Value, "000") & " - " & Data1.Recordset.Fields(1).Value
        SQL = SQL & vbCrLf & "Equipo: " & Data1.Recordset.Fields(2).Value
                
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            'Hay que eliminar
            NumRegElim = Me.Data1.Recordset.AbsolutePosition
            SQL = "Delete from " & NombreTabla & " where numtermi=" & DBSet(Data1.Recordset!NumTermi, "N")
            'SQL = SQL & " AND codtecni=" & Data1.Recordset!codtecni & " AND codclien=" & Data1.Recordset!CodClien
            
            conn.Execute SQL
            CancelaADODC Me.Data1
            CargaGrid ""
            CancelaADODC Me.Data1
            SituarDataTrasEliminar Data1, NumRegElim, True
            
            'SituarDataPosicion Me.Data1, NumRegElim, SQL
        End If
        
FinEliminar:
     Screen.MousePointer = vbDefault
     If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar terminal del TPV", Err.Description
End Function


Private Sub mnEliminar_Click()
'Eliminar
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If Data1.Recordset.EOF Then Exit Sub
    
    
    If BloqueaRegistro("spatpvt", "numtermi=" & Data1.Recordset!NumTermi) Then BotonModificar
'    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Anyadir
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3 'Eliminar
            mnEliminar_Click
        Case 6 'Salir
            mnSalir_Click
    End Select
End Sub


'Private Sub BotonAnyadir()
'    LimpiarCampos
'    PonerModo 3
'
'    Text1(0).Text = SugerirCodigoSiguienteStr("scryst", "codcryst")
'    PonerFoco Text1(0)
'End Sub


Private Sub BotonModificar()
Dim I As Integer
Dim anc As Single

    PonerModo 4
'    PonerFoco Text1(8)

    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
'    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
'        i = DataGrid1.Bookmark - DataGrid1.FirstRow
'        DataGrid1.Scroll 0, i
'        DataGrid1.Refresh
'    End If
    
    anc = ObtenerAlto(Me.DataGrid1, 5)
    LLamaLineas anc

    'poner valores grabados
    txtAux(0).Text = DBLet(DataGrid1.Columns(0).Value, "N")
'    txtAux(0).Text = Format(txtAux(0).Text, "000")
    txtAux(1).Text = DBLet(DataGrid1.Columns(1).Value, "T")
    txtAux(2).Text = DBLet(DataGrid1.Columns(2).Value, "T")
    txtAux(3).Text = DBLet(DataGrid1.Columns(3).Value, "N")
    txtAux(4).Text = DBLet(Data1.Recordset!NomImpre, "T")
    txtAux(6).Text = DBLet(Data1.Recordset!CualCom, "N")
    
'    For i = 3 To 8
'        txtAux(i).Text = DBLet(Data1.Recordset.Fields(i + 1).Value, "T")
'    Next i
'    txtAux(9).Text = DBLet(Data1.Recordset!numviaje, "T")
    
    FormateaCampo txtAux(0)
'    FormateaCampo txtAux(2)
'    For i = 4 To 8
'        FormateaCampo txtAux(i)
'    Next i

    DataGrid1.Enabled = False
    PonerFoco txtAux(1)

End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me, 3)
    
    
    If b Then
        If Me.chkHayVisor.Value = 1 Then
            If Me.cboPuertos.ListIndex = -1 Or Me.cboVelocidad.ListIndex = -1 Then
                MsgBox "Debe indicar el valor del puerto/velocidad", vbExclamation
                b = False
            End If
        End If
    End If
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
'On Error GoTo EPonerCampos
'
'    If Data1.Recordset.EOF Then Exit Sub
'    PonerCamposForma Me, Data1
'
'    'El combo del numpuerto lo ponermos en el numpuerto -1
'    'Me.cboPuertos.ListIndex = Me.cboPuertos.ListIndex - 1
'
'
'    BloquearChecks Me, Modo
'
'    PosicionarCombo Me.cboVelocidad, Text1(9).Text
'
'
'EPonerCampos:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
   
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not b
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
'    BloquearText1 Me, Modo
    
    'Bloquear los checkbox
    BloquearChecks Me, Modo

    b = (Modo = 0) Or (Modo = 2)
    Me.FrameImpresoras.Enabled = (Not b)
    Me.FrameVisor.Enabled = (Not b)
    Me.cboPuertos.Enabled = (Not b)
    Me.cboVelocidad.Enabled = (Not b)
    
    
     'Bloquear los campos de clave primaria al modificar
    BloquearTxt txtAux(0), (Modo = 4)
    
    
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
      PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not b 'Nuevo
    Me.mnNuevo.Enabled = Not b
    
    Me.Toolbar1.Buttons(2).Enabled = Not b 'Modificar
    Me.mnModificar.Enabled = Not b
End Sub




Private Sub txtAux_GotFocus(Index As Integer)
    If Index = 4 Then Exit Sub
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If Modo < 3 Then Exit Sub
    
    
    If Index = 6 Then
        If Not PonerFormatoEntero(txtAux(Index)) Then txtAux(Index).Text = ""
    End If
End Sub
