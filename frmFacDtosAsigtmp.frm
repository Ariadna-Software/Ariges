VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacDtosAsigtmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion descuentos"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   12210
   ClipControls    =   0   'False
   Icon            =   "frmFacDtosAsigtmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   11040
      TabIndex        =   16
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   8880
      MaxLength       =   16
      TabIndex        =   2
      Text            =   "codartic codarti"
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   4560
      MaxLength       =   16
      TabIndex        =   1
      Tag             =   "Famia|N|S|0|9999|tmpGenDtos|codfamia|0000|S|"
      Text            =   "codartic codarti"
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   4440
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   4800
      TabIndex        =   14
      ToolTipText     =   "Buscar artículo"
      Top             =   4080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmFacDtosAsigtmp.frx":000C
      Left            =   8280
      List            =   "frmFacDtosAsigtmp.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Clasifica|N|S|0|9999|tmpGenDtos|clasifica|0000|S|"
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   5160
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
         TabIndex        =   13
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   11
      ToolTipText     =   "Buscar artículo"
      Top             =   4080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   4440
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   960
      MaxLength       =   16
      TabIndex        =   0
      Tag             =   "Marca|N|S|0|9999|tmpGenDtos|codmarca|0000|S|"
      Text            =   "codartic codarti"
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   5280
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11040
      TabIndex        =   5
      Top             =   5280
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12210
      _ExtentX        =   21537
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
         Left            =   9120
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9000
      Top             =   5280
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
      Bindings        =   "frmFacDtosAsigtmp.frx":0022
      Height          =   4425
      Left            =   240
      TabIndex        =   6
      Top             =   705
      Width           =   11775
      _ExtentX        =   20770
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
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   5280
      Width           =   4695
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
      TabIndex        =   8
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
Attribute VB_Name = "frmFacDtosAsigtmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vFec As Date
Public vSoloNuevos As Boolean
Public vClien As Long
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1





Dim Ordenacion As String
Private Modo As Byte
Dim kCampo As Integer

Dim EsBusqueda As Boolean

Dim CadenaConsulta As String
Dim CadenaBusqueda As String
'Cadena para la consulta de de busqueda en Grid

Private HaDevueltoDatos As String
Dim Cad As String


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
                If InsertarMod Then
                    CargaGrid True
                    BotonAnyadir
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk Then
                         
                 If InsertarMod Then
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
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub cmdActualizar_Click()
    If Data1.Recordset.EOF Then
        MsgBox "No hay datos", vbExclamation
        Exit Sub
    End If
    
    Set miRsAux = New ADODB.Recordset
    Cad = "SELECT codfamia,COUNT(*) FROM tmpGenDtos WHERE codusu =" & vUsu.codigo & " GROUP BY codfamia HAVING COUNT(*)>1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & Format(miRsAux!Codfamia, "0000") & ":  " & miRsAux.Fields(1) & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If Cad <> "" Then
        Set miRsAux = Nothing
        Cad = "Existe valores duplicados para algunas familias" & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
        
    If MsgBox("¿Continuar con el proceso?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    
    
    Updatear
    
    Set miRsAux = Nothing
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
    CadenaDesdeOtroForm = "OK"   'para que refresque
    Unload Me
End Sub

Private Sub Updatear()
Dim Existe As Boolean
    On Error GoTo EU
    
    Data1.Recordset.MoveFirst
    
    While Not Data1.Recordset.EOF
    
        Label1.Caption = Data1.Recordset!codmarca & "/" & Data1.Recordset!Codfamia
        Label1.Refresh
        
        
        Cad = "Select * from sdtofm where codclien =" & Me.vClien & " AND  codfamia = " & Data1.Recordset!Codfamia
        
        'NO PONGO LA MARCA
        Cad = Cad & " AND codmarca is null"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Existe = False
        If Not miRsAux.EOF Then Existe = True
        miRsAux.Close
        
        Cad = ""
        If Not Existe Then
            'NUEVO
            Cad = "INSERT INTO sdtofm(codclien,codfamia,codmarca,fechadto,dtoline1,dtoline2) VALUES (" & vClien
            Cad = Cad & "," & Data1.Recordset!Codfamia & ",NULL,'"
            Cad = Cad & Format(Me.vFec, FormatoFecha) & "'," & TransformaComasPuntos(CStr(Data1.Recordset!dtoline1)) & ",0)"
        Else
            If Not Me.vSoloNuevos Then
                Cad = "UPDATE sdtofm SET dtoline1=" & TransformaComasPuntos(CStr(Data1.Recordset!dtoline1))
                Cad = Cad & " ,fechadto = " & DBSet(vFec, "F")
                Cad = Cad & " WHERE codclien = " & vClien & " AND codfamia = " & Data1.Recordset!Codfamia
                Cad = Cad & " AND codmarca is NULL "
                
            End If
        End If
        If Cad <> "" Then Ejecutar Cad, False
        Data1.Recordset.MoveNext
    Wend
    Exit Sub
EU:
    MuestraError Err.Number
End Sub


Private Sub cmdAux_Click(Index As Integer)
    HaDevueltoDatos = ""

            
    MandaBusquedaPrevia2 Index
            


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
            PonerFocoGrid DataGrid1
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub






Private Sub Combo1_Change()
    If Modo < 3 Then Exit Sub
    
End Sub

Private Sub Combo1_Click()
Dim Cad As String
    
    If Modo < 3 Then Exit Sub
    Cad = ""
    If Me.txtAux(1).Text = "" Then
        Combo1.ListIndex = -1
    
    Else
        If Combo1.ListIndex >= 0 Then
            Cad = "clasifica = " & Combo1.ItemData(Combo1.ListIndex) & " and codfamia "
            Cad = DevuelveDesdeBD(conAri, "dtoline1", "sfamiadtos", Cad, txtAux(1).Text)
            If Cad = "" Then Cad = "0"
        End If
    End If
    txtAux(2).Text = Cad
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
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
    EsBusqueda = False
    CadenaBusqueda = ""
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    Modo = 0
    CargaCombo
    Ordenacion = " ORDER BY codmarca,codfamia "
    CadenaConsulta = MontaSQLCarga(True)
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
   
    
   
        
 
'    CargaGrid (Modo = 2 Or Modo = 0)
    CargaGrid True
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim SQL As String
Dim tots As String
    
    On Error GoTo ECarga
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, SQL, False
    

    tots = "S|txtAux(0)|T|Marca|900|;S|cmdAux(0)|B||0|;S|txtAux2(0)|T|Marca|2200|;"
    tots = tots & "S|txtAux(1)|T|Fam.|900|;S|cmdAux(1)|B||0|;S|txtAux2(1)|T|Familia|2900|;"
    tots = tots & "N||||0|;S|Combo1|C|Desc.|2000|;S|txtAux(2)|T|Dto|900|;"
    arregla tots, DataGrid1, Me


'    'dtos alineados a la dcha
'    DataGrid1.Columns(6).Alignment = dbgCenter

    DataGrid1.ScrollBars = dbgAutomatic
    
   'Actualizar indicador
   If Not Data1.Recordset.EOF And (Modo = 0) Then
        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
   Else
        Me.lblIndicador.Caption = ""
   End If
   
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub



Private Sub LLamaLineas(alto As Single)
Dim jj As Integer
Dim B As Boolean

    DeseleccionaGrid Me.DataGrid1
    B = (Modo = 3 Or Modo = 4 Or Modo = 1) 'Insertar o Modificar

    For jj = 0 To txtAux.Count - 1
        If jj < 2 Then
            txtAux2(jj).Height = Me.DataGrid1.RowHeight
            txtAux2(jj).Top = alto
            txtAux2(jj).visible = B
        End If
        txtAux(jj).Height = DataGrid1.RowHeight
        txtAux(jj).Top = alto
        txtAux(jj).visible = B
    Next jj

    Me.Combo1.visible = B
    Me.Combo1.Top = alto
    
    For jj = 0 To Me.cmdAux.Count - 1
        Me.cmdAux(jj).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(jj).Top = alto
        Me.cmdAux(jj).visible = B
    Next jj
End Sub





Private Sub frmB_Selecionado(CadenaDevuelta As String)
    HaDevueltoDatos = CadenaDevuelta
End Sub



Private Sub mnBuscar_Click()
    BotonBuscar
End Sub



Private Sub mnEliminar_Click()
Dim SQL As String
    
    If Me.Data1.Recordset.EOF Then Exit Sub
    If MsgBox("Eliminar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    SQL = "DELETE from tmpGenDtos WHERE codusu= " & vUsu.codigo & " AND codmarca = " & Data1.Recordset!codmarca & " AND codfamia = " & Data1.Recordset!Codfamia
    SQL = SQL & " AND clasifica= " & Data1.Recordset!clasifica
    conn.Execute SQL
    BotonVerTodos
End Sub

Private Sub mnModificar_Click()
    BotonModificar
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
            frmListado2.Opcion = 35
            frmListado2.Show vbModal
            
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
Dim B As Boolean
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)

                      
    If Kmodo = 1 Then 'Modo Buscar
        PonerFoco txtAux(0)
    End If
                           
                           
    BloquearTxt txtAux(2), True
                           
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)

    Me.cmdAux(0).Enabled = (Modo <> 4)
     Me.cmdAux(1).Enabled = (Modo <> 4)

     
    '-----------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    cmdActualizar.visible = Modo = 2
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
Dim B As Boolean
    
    
    
    'Insertar
    B = (Modo = 2) Or Modo = 0
    Toolbar1.Buttons(5).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2)
    
    'eliminar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnEliminar.Enabled = B
    
    
    'Modificar
    Toolbar1.Buttons(6).Enabled = B
    Me.mnModificar.Enabled = B
    
    
    B = ((Modo >= 3))
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not B
    Me.mnBuscar.Enabled = Not B
    'VerTodos
    Toolbar1.Buttons(2).Enabled = Not B
    Me.mnVerTodos.Enabled = Not B
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    DesplazamientoData Data1, Index
    PonerCampos
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
    
    SQL = "SELECT tmpGenDtos.codmarca,nommarca,tmpGenDtos.codfamia,nomfamia,tmpGenDtos.clasifica,nombre,sfamiadtos.dtoline1 "
    SQL = SQL & " FROM tmpGenDtos,smarca,sfamia,sfamiadtos,sfamiatipodto WHERE tmpgendtos.codusu = " & vUsu.codigo
    SQL = SQL & " AND tmpGenDtos.codmarca=smarca.codmarca  AND tmpGenDtos.codfamia=sfamia.codfamia "
    SQL = SQL & " AND tmpGenDtos.codfamia=sfamia.codfamia AND sfamiadtos.codfamia=tmpGenDtos.codfamia AND sfamiadtos.clasifica=tmpGenDtos.clasifica"
    SQL = SQL & " AND sfamiatipodto.clasifica=tmpGenDtos.clasifica"
    
    If enlaza Then
        If EsBusqueda And CadenaBusqueda <> "" Then SQL = SQL & CadenaBusqueda
    Else
        SQL = SQL & " AND  tmpGenDtos.codmarca = -1"
    End If
    SQL = SQL & Ordenacion
    MontaSQLCarga = SQL
End Function


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    EsBusqueda = True
    If Modo <> 1 Then
        LimpiarCampos
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
    PonerFocoGrid DataGrid1

    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub BotonModificar()
Dim I As Integer
Dim anc As Single

    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo1, campo2 y campo3 es clave primaria, NO se puede modificar
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    anc = ObtenerAlto(Me.DataGrid1, 10)
    LLamaLineas anc
    
 
    For I = 0 To 1
        txtAux(I).Text = DBLet(DataGrid1.Columns(I * 2).Value, "T")
        txtAux2(I).Text = DBLet(DataGrid1.Columns((I * 2) + 1).Value, "T")
    Next I

 
    
    SituarCombo Combo1, DBLet(Data1.Recordset!clasifica, "N")
    
 
 
    
    DataGrid1.Enabled = False
    PonerFoco txtAux(2)
End Sub





Private Function DatosOk() As Boolean
Dim B As Boolean


    On Error GoTo ErrDatosOK
    DatosOk = False
    B = True
    If txtAux(0).Text = "" Or txtAux(1).Text = "" Then B = False
    If Combo1.ListIndex < 0 Then B = False
    If Not B Then
        MsgBox "Campos obligatorios", vbExclamation
    Else
        If Me.txtAux(2).Text = "" Then
            MsgBox "No hay descuento asignado", vbExclamation
            B = False
        End If
    End If
    If Not B Then Exit Function
    

    
    DatosOk = B
    Exit Function
    
ErrDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos OK.", Err.Description
End Function

Private Sub BotonAnyadir()
Dim anc As Single
    
    'Situamos el grid al final
    limpiar Me
    Combo1.ListIndex = -1
    AnyadirLinea DataGrid1, Data1
      
    anc = ObtenerAlto(DataGrid1, 10)
    PonerModo 3
    
    LLamaLineas anc
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub MandaBusquedaPrevia2(Index As Integer)
''Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
'Dim Tabla As String
'Dim Titulo As String
'
'    'Llamamos a al form
'    cad = ""
'    'Estamos en Modo de Cabeceras
'    'Registro de la tabla de cabeceras: slista
        'Cod Diag.|tabla|columna|tipo|formato|10·
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        
        Select Case Index
        Case 0
            Cad = "Codigo|smarca|codmarca|N||20·"
            Cad = Cad & "Decripcion|smarca|nommarca|T||60·"
            frmB.vTabla = "smarca"
            frmB.vTitulo = "Marcas"
        Case 1
            Cad = "Codigo|sfamia|codfamia|N||20·"
            Cad = Cad & "Decripcion|sfamia|nomfamia|T||60·"
            frmB.vTabla = "sfamia"
            frmB.vTitulo = "Familias"
        End Select
        frmB.vCampos = Cad
        
        'frmB.vTabla = tabla
        frmB.vSQL = ""
        
        
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri       'Conexión a BD: Ariges
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos <> "" Then
            txtAux(Index).Text = RecuperaValor(HaDevueltoDatos, 1)
            txtAux2(Index).Text = RecuperaValor(HaDevueltoDatos, 2)
        End If
    

End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    
        
    
    
    
 '   If chkVistaPrevia = 1 Then
 '       MandaBusquedaPrevia cadB
 '   ElseIf cadB <> "" Then 'Se muestran en el mismo form
        CadenaBusqueda = " AND " & cadB
        CadenaConsulta = MontaSQLCarga(True)
        'CadenaBusqueda = " AND " & cadB
        PonerCadenaBusqueda
 '   End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CargaGrid False
        MsgBox "No hay ningún registro en la tabla para ese criterio de Búsqueda.", vbInformation
        Screen.MousePointer = vbDefault
        PonerModo Modo
        Exit Sub
    Else
        PonerModo 2
        PonerCampos
    End If
    LLamaLineas 10
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
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
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
                If Index = 3 Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    SendKeys "{tab}"
                End If
    End Select
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Cad As String

    On Error Resume Next
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    If Modo = 1 Then Exit Sub
    
    Select Case Index
        Case 0 To 1
            Cad = ""
            If txtAux(Index).Text <> "" Then
                If Not IsNumeric(txtAux(Index).Text) Then
                    MsgBox "Campo numerico", vbExclamation
                    txtAux(Index).Text = ""
                Else
                    If Index = 0 Then
                        Cad = DevuelveDesdeBD(conAri, "nommarca", "smarca", "codmarca", txtAux(Index).Text)
                    Else
                     Cad = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtAux(Index).Text)
                    End If
                    If Cad = "" Then MsgBox "No existe el valor en la BD: " & txtAux(Index).Text, vbExclamation
                End If
                If Cad = "" And txtAux(Index).Text <> "" Then
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
                      
            End If
            txtAux2(Index).Text = Cad
            If Modo = 3 Then Combo1.ListIndex = -1
                
            
        
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function InsertarMod() As Boolean
Dim C As String
    
    If Modo = 3 Then
        C = "INSERT INTO tmpGenDtos(codusu,codmarca,codfamia,clasifica) VALUES ("
        C = C & vUsu.codigo & "," & txtAux(0).Text & "," & txtAux(1).Text & "," & Combo1.ItemData(Combo1.ListIndex) & ")"
    Else
        C = "UPDATE tmpGenDtos set clasifica = " & Combo1.ItemData(Combo1.ListIndex)
        C = C & " WHERE CODUSU = " & vUsu.codigo & " AND codmarca = " & txtAux(0).Text & " AND codfamia = " & txtAux(1).Text
        C = C & " AND clasifica = " & Data1.Recordset!clasifica
    End If
    
    If Ejecutar(C, False) Then
        InsertarMod = True
    Else
        InsertarMod = False
    End If
        
End Function





Private Sub CargaCombo()

    Combo1.Clear

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select clasifica elcodigo,nombre elNombre from sfamiatipodto ORDER BY clasifica", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!ElCodigo & " - " & miRsAux!ElNombre
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!ElCodigo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub
