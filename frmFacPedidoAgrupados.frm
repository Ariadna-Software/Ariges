VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacPedidoAgrupados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agrupacion pedidos cliente"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17205
   ClipControls    =   0   'False
   Icon            =   "frmFacPedidoAgrupados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   17205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   16575
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   8280
         MaxLength       =   16
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmFacPedidoAgrupados.frx":000C
         Left            =   11400
         List            =   "frmFacPedidoAgrupados.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   1
         Tag             =   "Forpa|N|N|0|9999||codforpa|0000|N|"
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "Cliente|N|N|0|||codclien|0000|N|"
         Text            =   "Text1"
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   200
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   10320
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   7200
         TabIndex        =   15
         Top             =   240
         Width           =   675
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   8040
         Picture         =   "frmFacPedidoAgrupados.frx":0020
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Forma pago"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   195
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         ToolTipText     =   "Buscar almacen"
         Top             =   195
         Width           =   240
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   9120
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   180
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14280
      TabIndex        =   2
      Top             =   9120
      Width           =   1135
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15720
      TabIndex        =   3
      Top             =   9120
      Width           =   1135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacPedidoAgrupados.frx":00AB
      Height          =   7995
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   14102
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8880
      Top             =   5160
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
   Begin VB.Label lblInfInv 
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   9240
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmFacPedidoAgrupados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cliente As Long
Private WithEvents frmC3 As frmBasico2
Attribute frmC3.VB_VarHelpID = -1
Private WithEvents frmFP As frmBasico2 'frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1



Dim kCampo As Integer

Dim CadenaConsulta As String

Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
                           'Para el RowColChange, si el grid no esta totalmente cargado el CargaTxtAux da error.

Dim PrimeraVez As Boolean
Dim Orden As Integer
Dim Desce As Boolean

Private HaDevueltoDatos As Boolean

Private ColAlbaran As Collection



Private Sub cmdAceptar_Click()
Dim cad As String
Dim i As Integer
Dim TodoOk As Boolean

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
   
    Set ColAlbaran = New Collection
    If PrimerasComprobaciones Then
        TodoOk = True
        For i = 1 To ColAlbaran.Count
            conn.BeginTrans
            lblIndicador.Caption = "(" & i & "/" & ColAlbaran.Count & ") ..." & Mid(ColAlbaran.Item(i), 1, 0)
            lblIndicador.Refresh
            If RealizarAlbaran(ColAlbaran.Item(i)) Then
                conn.CommitTrans
                
            Else
                conn.RollbackTrans
                TodoOk = False
                Exit For
            End If
            lblIndicador.Caption = "actualizando"
            lblIndicador.Refresh
            Espera 1
        Next
        
        CargaGrid
       
    End If
    CadenaConsulta = ""
    
    txtAux.visible = False
Error1:
    Screen.MousePointer = vbDefault
    Set ColAlbaran = Nothing
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
    lblIndicador.Caption = ""
     If TodoOk Then
            MsgBox "Proceso finalizado", vbInformation
            Unload Me
        End If
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar
    
     If txtAux.visible Then
        txtAux.visible = False
        Exit Sub
    End If

    lblInfInv.Caption = ""
    Unload Me
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    
    If ColIndex > 2 Then Exit Sub
    If Orden = ColIndex Then
        Desce = Not Desce
    Else
        Orden = ColIndex
        Desce = False
    End If
     txtAux.visible = False
    CargaGrid
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF And gridCargado Then
       CargaTxtAux True, False
    End If
End Sub

Private Sub DataGrid1_Scroll(Cancel As Integer)
    If txtAux.visible Then Cancel = 1
End Sub

Private Sub Form_Activate()

    If PrimeraVez Then
       
        If Cliente >= 0 Then
            
            Text1(0).Text = Format(Cliente, "0000")
            CadenaConsulta = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(0).Text)
            Text2(0).Text = CadenaConsulta
            CargaGrid
                
        End If
    End If

    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
   'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListComun.ListImages(1).Picture
    Next kCampo

    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True
    
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
    PonerModo 1
    Combo1.ListIndex = 0
    Orden = 0
    If Cliente < 0 Then CargaGrid
        
        
    
            
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub CargaGrid()
Dim i As Byte
Dim SQL As String
On Error GoTo ECarga

    gridCargado = False
    
    lblIndicador.Caption = "leyendo BD"
    lblIndicador.Refresh
    
    
    
    SQL = MontaSQLCarga()
    CargaGridGnral DataGrid1, Me.Data1, SQL, PrimeraVez, 360
    
    PrimeraVez = False
        
    'Cod. Articulo
    DataGrid1.Columns(0).Caption = "NºPed."
    DataGrid1.Columns(0).Width = 1050
    DataGrid1.Columns(0).NumberFormat = "00000"
    DataGrid1.Columns(1).Caption = "Fecha"
    DataGrid1.Columns(1).Width = 1250
    DataGrid1.Columns(1).NumberFormat = "dd/mm/yyyy"
    
    DataGrid1.Columns(2).visible = False
    
    DataGrid1.Columns(3).Caption = "Referencia"
    DataGrid1.Columns(3).Width = 2250
   
    
    i = 3
    DataGrid1.Columns(i + 1).Caption = "Articulo"
    DataGrid1.Columns(i + 1).Width = 1450

    DataGrid1.Columns(i + 2).Caption = "Descripcion"
    DataGrid1.Columns(i + 2).Width = 4300
       
    
    DataGrid1.Columns(i + 3).Caption = "Solicitadas"
    DataGrid1.Columns(i + 3).Width = 1250
    
    DataGrid1.Columns(i + 4).Caption = "Pdtes"
    DataGrid1.Columns(i + 4).Width = 1260
    
    
    DataGrid1.Columns(i + 5).Caption = "Servir"
    DataGrid1.Columns(i + 5).Width = 1300
    
    
    DataGrid1.Columns(i + 6).Caption = "Resto"
    DataGrid1.Columns(i + 6).Width = 1300
    
    
    DataGrid1.Columns(i + 7).Caption = " P"
    DataGrid1.Columns(i + 7).Width = 600
    
    
    
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
        DataGrid1.Columns(i).Locked = True
        If i > 5 Then
            DataGrid1.Columns(i).Alignment = dbgRight
            DataGrid1.Columns(i).NumberFormat = FormatoCantidad
        End If
    Next i
    DataGrid1.ScrollBars = dbgAutomatic
    gridCargado = True
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
    lblIndicador.Caption = ""
    lblIndicador.Refresh
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
        txtAux.visible = visible
    Else
        DeseleccionaGrid Me.DataGrid1
        If Not limpiar Then 'Vaciar los textBox (Vamos a Insertar)
                txtAux.Text = Data1.Recordset!servidas
                txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        'Fijamos altura y posición Top
        '-------------------------------
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        txtAux.Left = DataGrid1.Columns(8).Left + 130 'codalmac
        txtAux.Width = DataGrid1.Columns(8).Width - 10
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux.visible = visible
    End If
    PonerFoco txtAux
    
    If visible Then
        txtAux.TabIndex = 2
        txtAux.SelStart = 0
        txtAux.SelLength = Len(txtAux.Text)
    Else
        txtAux.TabIndex = 5
    End If
End Sub




Private Sub frmC3_DatoSeleccionado(CadenaSeleccion As String)
     CadenaConsulta = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(2).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
    CadenaConsulta = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)

 
    Screen.MousePointer = vbHourglass
    imgBuscar(0).Tag = Index
    CadenaConsulta = ""
    Select Case Index
        Case 0 'Codigo Almacen
            Set frmC3 = New frmBasico2
            AyudaClientes frmC3, Text1(Index)
            Set frmC3 = Nothing
        Case 1 'Codigo Familia / Cod. Proveedor
'            Set frmFP = New frmFacFormasPago
'            frmFP.DatosADevolverBusqueda = "0|1|"
'            frmFP.Show vbModal
            Set frmFP = New frmBasico2
            AyudaFormasPago frmFP, Text1(Index)
            Set frmFP = Nothing
    End Select
    If CadenaConsulta <> "" Then
        Text1(Index).Text = RecuperaValor(CadenaConsulta, 1)
        Text2(Index).Text = RecuperaValor(CadenaConsulta, 2)
        CadenaConsulta = ""
        If Index = 0 Then
            Screen.MousePointer = vbHourglass
            txtAux.Text = ""
            Text1_LostFocus 0
            
        
        End If
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now

    PonerFormatoFecha Text1(2)
   If Text1(2).Text <> "" Then frmF.Fecha = CDate(Text1(2).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(2)
   
End Sub

Private Sub Label1_Click(Index As Integer)
    If Index = 2 Then
        If vUsu.Nivel <= 1 Then
            If Combo1.ListCount = 1 Then
                Combo1.AddItem "Presupuesto"
                HaMostradoCanal2_El_B = True
            End If
        End If
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim campo As String
Dim tabla As String

    If Not PerderFocoGnral(Text1(Index), 3) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    
    If Index < 2 Then
        If Text1(Index).Text = "" Then
            Text2(Index).Text = ""
        Else
            If Index = 1 Then
                campo = "nomforpa"
                tabla = "sforpa"
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, tabla, campo)
                
            
            
            Else
                'Cliente
                campo = "nomclien"
                tabla = "sclien"
                CadenaConsulta = "codforpa"
                Text2(Index).Text = DevuelveDesdeBD(conAri, campo, tabla, "codclien", Text1(Index).Text, "N", CadenaConsulta)
                If Text2(Index).Text <> "" Then
                    Text1(1).Text = CadenaConsulta
                    CadenaConsulta = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", CadenaConsulta)
                    Text2(1).Text = CadenaConsulta
                End If
            End If
            
            If Text1(Index).Text <> "" And Text2(Index).Text = "" Then
                Text1(Index).Text = ""
                If Index = 0 Then PonerFoco Text1(Index)
            Else
                If Index = 0 Then CargaGrid
            End If
            
        End If
    Else
        PonerFormatoFecha Text1(Index)
    End If

End Sub


Private Sub txtAux_GotFocus()
    ConseguirFocoLin txtAux
End Sub

Private Sub TxtAux_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If KeyCode = 38 Or KeyCode = 40 Then
        ModificarExistencia
    End If

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                End If
        
'                If DataGrid1.Row > 0 Then
'                    DataGrid1.Row = DataGrid1.Row - 1
'                    CargaTxtAux True, True
''                elseif
'                End If
        Case 40 'Desplazamiento Flecha Hacia Abajo
                PasarSigReg
                
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)

'    KEYpress KeyAscii
    
   If KeyAscii = 13 Then 'ENTER
        If ModificarExistencia Then PasarSigReg
   ElseIf KeyAscii = 27 Then
        CargaTxtAux True, False
   End If
End Sub


Private Sub txtAux_LostFocus()
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtAux
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        'Formato tipo 1: Decimal(12,2)
        PonerFormatoDecimal txtAux, 1
    End With
End Sub





Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 3, cerrar
   ' If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
       
    
    'PonerIndicador lblIndicador, 3
    
    b = False
    BloquearTxt Text1(0), b
    BloquearTxt Text1(1), b
    
    b = True
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i



    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Function MontaSQLCarga() As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim miOrden As String
    '                                                                                           round((importel/cantidad),4) precio ,
    SQL = "select sliped.numpedcl,fecpedcl,numlinea,referenc,codartic,nomartic,solicitadas,cantidad,servidas,"
    SQL = SQL & " cantidad - servidas Pendiente, if(servidas>0 ,'S','') ped"
    SQL = SQL & " from sliped,scaped where sliped.numpedcl=scaped.numpedcl and cerrado=0 and cantidad >0 AND "
    
    
    
    If Text1(0).Text <> "" Then
        SQL = SQL & " codclien= " & Val(Text1(0).Text)
    Else
        SQL = SQL & " false"
    End If

    SQL = SQL & " order by "
    
    If Orden = 0 Then
        miOrden = " numpedcl  " & IIf(Desce, "DESC", "") & " ,fecpedcl"
    Else
        miOrden = " fecpedcl " & IIf(Desce, "DESC", "") & ",numpedcl"
    End If
    
    miOrden = miOrden & " ,numlinea"
    MontaSQLCarga = SQL & miOrden
End Function






Private Function DatosOk() As Boolean
'Solo se actualiza el campo de Existencia Real
Dim canti As Currency
    txtAux.Text = Trim(txtAux.Text)
    If txtAux.Text = "" Then txtAux.Text = "0"
    If txtAux.Text <> "" And EsNumerico(txtAux.Text) Then
        If PonerFormatoDecimal(txtAux, 1) Then
            DatosOk = True
        Else
            DatosOk = False
        End If
        'DatosOk = True
    Else
        DatosOk = False
    End If
    
    
    If DatosOk Then
        
        canti = ImporteFormateado(txtAux.Text)
        If canti < 0 Then
            MsgBox "Cantidad servir negativa", vbExclamation
            DatosOk = False
            
        Else
            If canti > Data1.Recordset!cantidad Then
                MsgBox "Cantidad superior a la pendiente", vbExclamation
                 '  DatosOk = False
            End If
        End If
        If Not DatosOk Then txtAux.Text = ""
    End If
    
End Function





Private Sub PonerOpcionesMenu()
   ' no tiene menu nui toolbar PonerOpcionesMenuGeneral Me
End Sub


Private Function ActualizarExistencia(canti As String) As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim SQL As String
Dim ADonde As String

    On Error GoTo EActualizar

    ADonde = "Modificando datos de Inventario (Tabla: sinven)."
    SQL = "UPDATE sliped Set servidas = " & DBSet(canti, "N")
    SQL = SQL & " WHERE numpedcl =" & Data1.Recordset!NumPedcl
    SQL = SQL & " AND numlinea =" & Data1.Recordset!numlinea
    conn.Execute SQL
    
    
    ActualizarExistencia = True
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         MuestraError Err.Number, SQL, Err.Description
         
         ActualizarExistencia = False
    Else
        ActualizarExistencia = True
        
    End If
End Function


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid1.Bookmark < Data1.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
    ElseIf DataGrid1.Bookmark = Data1.Recordset.RecordCount Then
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long
Dim Indicador As String

    If DatosOk Then
        If ActualizarExistencia(txtAux.Text) Then
            TerminaBloquear
            NumReg = Data1.Recordset.AbsolutePosition
            CargaGrid
            If SituarDataPosicion(Data1, NumReg, Indicador) Then

            End If
            ModificarExistencia = True
        Else
            ModificarExistencia = False
        End If
    End If
End Function





Private Function PrimerasComprobaciones() As Boolean
Dim Aux As String
Dim A_Servir As Currency
Dim NuevoAlbaran As String
Dim NumpedPed As String
Dim RN As ADODB.Recordset
Dim C1 As Currency
Dim PedidosArticulosdistintos As String
Dim vCli As CCliente

On Error GoTo ePrimerasComprobaciones
    PrimerasComprobaciones = False
    
    
    If Text1(0).Text = "" Then Exit Function
    
    
    CadenaConsulta = ""
    Set vCli = New CCliente
    If vCli.LeerDatos(Text1(0).Text) Then
        If vCli.ClienteBloqueado(2, False) Then CadenaConsulta = "N"
    Else
        MsgBox "Error leyendo cliente", vbExclamation
        CadenaConsulta = "N"
    End If
    Set vCli = Nothing
    
    If CadenaConsulta <> "" Then Exit Function
    
    CadenaConsulta = "scaped.numpedcl=Sliped.numpedcl AND cantidad>0 AND codclien = " & Text1(0).Text & " AND 1"
    CadenaConsulta = DevuelveDesdeBD(conAri, "sum(servidas)", "scaped,sliped", CadenaConsulta, "1")
    If CadenaConsulta = "" Then CadenaConsulta = "0"
    
    If CCur(CadenaConsulta) = 0 Then
        MsgBox "Ningun datos seleccionado para servir", vbExclamation
        Exit Function
    End If
    
    'Fecha
    Aux = ""
    If Not EsFechaOK(Text1(2).Text) Then
        Aux = "Fecha incorrecta"
        
    Else
        If CDate(Text1(2).Text) < vEmpresa.FechaIni Then
            Aux = "Fecha anterior inicio ejercicios"
        Else
            If CDate(Text1(2).Text) > DateAdd("yyyy", 1, vEmpresa.FechaFin) Then Aux = "Fecha posterior fin ejercicios"
        End If
    End If
    If Aux <> "" Then
        MsgBox Aux, vbExclamation
        Exit Function
    End If
    
    If PonerTrabajadorConectado(Aux) = "" Then
        Aux = "Error trabajador conectado " & vbCrLf & vbCrLf
    Else
        Aux = ""
    End If
    If Text1(1).Text = "" Then Aux = Aux & vbCrLf & "Forma de pago"
    
    
    If Aux <> "" Then
        MsgBox Aux, vbExclamation
        Exit Function
    End If
    
    Set miRsAux = New ADODB.Recordset
    
    
    PedidosArticulosdistintos = ""
    Aux = "scaped.numpedcl=Sliped.numpedcl AND servidas>0 AND codclien = " & Text1(0).Text & " AND 1"
    Aux = "Select scaped.numpedcl from scaped,sliped WHERE " & Aux & "  GROUP BY scaped.numpedcl"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Aux = "0"
    While Not miRsAux.EOF
        Aux = Val(Aux) + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Val(Aux) > 0 Then PedidosArticulosdistintos = vbCrLf & "Pedidos: " & Aux
             
    
    Aux = "scaped.numpedcl=Sliped.numpedcl AND servidas>0 AND codclien = " & Text1(0).Text & " AND 1"
    Aux = "Select codartic from scaped,sliped WHERE " & Aux & "  GROUP BY codartic"
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Aux = "0"
    While Not miRsAux.EOF
        Aux = Val(Aux) + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Val(Aux) > 0 Then PedidosArticulosdistintos = PedidosArticulosdistintos & vbCrLf & "Articulos: " & Aux
    
    
    
    
    Aux = "select coddirec,coalesce(referenc,'') referenc,codagent from scaped,sliped WHERE scaped.numpedcl=Sliped.numpedcl AND codclien = " & Text1(0).Text & " AND cantidad>0 AND servidas>0"
    Aux = Aux & " group by 1,2,3"
    Set RN = New ADODB.Recordset
    NuevoAlbaran = "-1"
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        'Comprobaremos stocks,cuantos albaranes.......
        Aux = "select * from scaped,sliped WHERE scaped.numpedcl=Sliped.numpedcl AND codclien = " & Text1(0).Text & " AND servidas>0"
        Aux = Aux & " AND coddirec  "
        If IsNull(RN!CodDirec) Then
            Aux = Aux & " is null"
        Else
            Aux = Aux & "= " & RN!CodDirec
        End If
        Aux = Aux & " AND coalesce(referenc ,'')"
        If IsNull(RN!referenc) Then
            Aux = Aux & " is null"
        Else
            Aux = Aux & " = " & DBSet(RN!referenc, "T", "N")
        End If
        
        Aux = Aux & " AND codagent =" & RN!CodAgent
        
        Aux = Aux & " ORDER BY scaped.numpedcl,numlinea"
        
        
        miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumpedPed = ""
        While Not miRsAux.EOF
                
            
      
            If InStr(1, NumpedPed, Format(miRsAux!NumPedcl, "000000")) = 0 Then NumpedPed = NumpedPed & ", " & Format(miRsAux!NumPedcl, "000000")
           
            A_Servir = A_Servir + miRsAux!servidas
            
        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If NumpedPed <> "" Then ColAlbaran.Add NumpedPed
        RN.MoveNext
    Wend
    RN.Close
    
    Aux = "select codalmac,sliped.codartic,sliped.nomartic,sum(servidas) Cuantas from scaped,sliped,sartic "
    Aux = Aux & " WHERE scaped.numpedcl=Sliped.numpedcl AND sliped.codartic=sartic.codartic AND artvario=0 AND codclien = " & Text1(0).Text & " AND cantidad>0 AND servidas>0 GROUP BY 1,2"
    RN.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    While Not RN.EOF
        NuevoAlbaran = "codalmac = " & RN!codAlmac & " AND codartic "
        NuevoAlbaran = DevuelveDesdeBD(conAri, "canstock", "salmac", NuevoAlbaran, RN!codArtic, "T")
        C1 = NuevoAlbaran
        If DBLet(RN!cuantas, "N") > 0 Then
            If RN!cuantas > C1 Then Aux = Aux & "  -" & RN!NomArtic & "  St: " & C1 & " Pd:" & RN!cuantas & vbCrLf
        End If
        
        RN.MoveNext
    Wend
    RN.Close
    
    If Aux <> "" Then Aux = "STOCKS insuficiente" & vbCrLf & Aux & vbCrLf
    Aux = Aux & "Albaranes a generar: " & Format(ColAlbaran.Count, "000") & vbCrLf & vbCrLf & vbCrLf
    Aux = Aux & PedidosArticulosdistintos
    Aux = Aux & vbCrLf & "Unidades totales a servir: " & Format(A_Servir, FormatoCantidad) & vbCrLf
    
    
    If Combo1.ListIndex = 1 Then Aux = Aux & vbCrLf & vbCrLf & "***** CANAL ******"
    If MsgBox(Aux, vbQuestion + vbYesNoCancel) = vbYes Then PrimerasComprobaciones = True
    



ePrimerasComprobaciones:
    If Err.Number <> 0 Then MuestraError Err.Number, , CadenaConsulta
    Set miRsAux = Nothing
    Set RN = Nothing
End Function



Private Function RealizarAlbaran(ListaPedidos As String) As Boolean
Dim Aux As String
Dim TrabajadorC As String
Dim bol As Boolean, Existe As Boolean
Dim devuelve As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim codtipom As String
Dim NumAlb As Long
Dim linea As Integer
Dim vCStock As CStock
Dim vSQL As String
Dim PesoAlbaran As Currency
Dim Incrementa As Boolean

    On Error GoTo eRealizarAlbaran
    RealizarAlbaran = False

    
    TrabajadorC = PonerTrabajadorConectado(Aux)
    
    Aux = "select scaped.*,sliped.*"
    Aux = Aux & ",codenvio,codzonas"
    Aux = Aux & "  from scaped,sliped,sclien WHERE scaped.codclien=sclien.codclien AND "
    Aux = Aux & " scaped.numpedcl=Sliped.numpedcl AND scaped.codclien = " & Text1(0).Text & " AND servidas>0"
    Aux = Aux & " AND scaped.numpedcl IN (" & Mid(ListaPedidos, 2) & ")"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then Err.Raise 513, "", "No existe ningun registro"
    

    
    bol = False
    
    
    'Siempre van a ALZ
    codtipom = IIf(Combo1.ListIndex = 1, "ALZ", "ALV")
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(codtipom) Then
        'Comprobar si mientras tanto se incremento el contador de Pedidos
        'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
        Do
        
            NumAlb = -1
            Incrementa = True
            If vParamAplic.NumeroInstalacion = vbFenollar Then
                devuelve = BuscaHueco(codtipom)
                If devuelve <> "" Then
                    NumAlb = Val(devuelve)
                    Incrementa = False
                End If
            End If
            If NumAlb <= 0 Then devuelve = vTipoMov.ConseguirContador(codtipom)
            devuelve = DevuelveDesdeBDNew(conAri, "scaalb", "numalbar", "codtipom", codtipom, "T", , "numalbar", CStr(NumAlb), "N")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                If Incrementa Then
                    vTipoMov.IncrementarContador (codtipom)
                    NumAlb = vTipoMov.ConseguirContador(codtipom)
                End If
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    
    
    If NumAlb < 0 Then Err.Raise 513, , "Error consiguiendo contador"
    
    'Acabar la sql con el contador seleccionado
    vSQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien"
    vSQL = vSQL & ",nifclien,telclien,coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,"
    vSQL = vSQL & "dtoppago,dtognral,tipofact,observa01,observa02,observa03,observa04,observa05,numofert,fecofert,"
    vSQL = vSQL & "numpedcl,fecpedcl,fecentre,sementre,coddiren,tipAlbaran,codzonas,fecenvio,codinter,codnatura,chofer) VALUES ("
    vSQL = vSQL & "'" & codtipom & "'," & NumAlb & ", " & DBSet(Text1(2).Text, "F") & ",0," & miRsAux!codClien
    vSQL = vSQL & "," & DBSet(miRsAux!NomClien, "T") & "," & DBSet(miRsAux!domclien, "T", "N") & "," & DBSet(miRsAux!codpobla, "N")
    vSQL = vSQL & "," & DBSet(miRsAux!pobclien, "T", "N") & "," & DBSet(miRsAux!proclien, "T", "N") & "," & DBSet(miRsAux!nifClien, "T")
    vSQL = vSQL & "," & DBSet(miRsAux!telclien, "T") & "," & DBSet(miRsAux!CodDirec, "N", "S") & "," & DBSet(miRsAux!nomdirec, "T", "S")
    vSQL = vSQL & "," & DBSet(miRsAux!referenc, "T", "N") & "," & TrabajadorC & "," & TrabajadorC & "," & DBSet(miRsAux!CodTraba, "T")
    
    vSQL = vSQL & "," & DBSet(miRsAux!CodAgent, "T") & ","
    If Combo1.ListIndex = 0 Then
        vSQL = vSQL & DBSet(miRsAux!codforpa, "N")
    Else
        vSQL = vSQL & "2"  'EFECTIVO CONTADO
    End If
    vSQL = vSQL & "," & DBSet(miRsAux!CodEnvio, "N") & ",0,0," & DBSet(miRsAux!TipoFact, "N")
    vSQL = vSQL & "," & DBSet(miRsAux!observa01, "T") & "," & DBSet(miRsAux!observa02, "T") & "," & DBSet(miRsAux!observa03, "T")
    vSQL = vSQL & "," & DBSet(miRsAux!observa04, "T") & "," & DBSet(miRsAux!observa05, "T") & ",null,null,0," & DBSet(Now, "F") 'pedido:0 fecha :actual
    vSQL = vSQL & ",null,null," & DBSet(miRsAux!coddiren, "N", "S") & ",0," & DBSet(miRsAux!codzonas, "T")
    
    'fecenvio
    vSQL = vSQL & "," & DBSet(Text1(2).Text, "F")
    
    'codinter,codnatura,chofer
    Aux = " codenvio = " & miRsAux!CodEnvio & " AND defecto"
    Aux = DevuelveDesdeBD(conAri, "chofer", "sconductor", Aux, "1")
    vSQL = vSQL & ",1,4," & DBSet(Aux, "T", "S") & ")"
    
    'Insertar Cabecera
    conn.Execute vSQL, , adCmdText
    
    
    
    
    'Portes
    
    
    
    
    
    
    

    Set vCStock = New CStock
    linea = 0
    
    While Not miRsAux.EOF
        linea = linea + 1


        'STOCK
        vCStock.Documento = Format(NumAlb, "0000000")
        vCStock.tipoMov = "S"
        vCStock.DetaMov = codtipom
        vCStock.Trabajador = CLng(TrabajadorC) 'En codigope ponemos el Cliente
        vCStock.codArtic = miRsAux!codArtic
        vCStock.codAlmac = CInt(miRsAux!codAlmac)
        vCStock.FechaMov = CDate(Text1(2).Text)
        vCStock.HoraMov = CDate(vCStock.FechaMov & " " & Format(Now, "hh:nn:ss"))

        vCStock.cantidad = CSng(miRsAux!servidas)
        vCStock.Importe = CCur(CalcularImporte(miRsAux!servidas, miRsAux!precioar, miRsAux!dtoline1, miRsAux!dtoline2, vParamAplic.TipoDtos))
    
        vCStock.LineaDocu = linea
        If Not vCStock.ActualizarStock(False, True) Then Err.Raise 513, , "Actualizando stock"

    
        vSQL = "INSERT INTO slialb (codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        vSQL = vSQL & "cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre,codprovex,numlote,codccost,idL) "
        vSQL = vSQL & " VALUES('" & codtipom & "', " & NumAlb & ", " & linea & " , "
        vSQL = vSQL & vCStock.codAlmac & ", " & DBSet(miRsAux!codArtic, "T") & ", " & DBSet(miRsAux!NomArtic, "T") & ", " & DBSet(miRsAux!Ampliaci, "T") & ", "
        vSQL = vSQL & DBSet(miRsAux!servidas, "N") & ", " & DBSet(miRsAux!servidas, "N") & ", "
        vSQL = vSQL & DBSet(miRsAux!precioar, "N") & ", " & DBSet(miRsAux!dtoline1, "N") & ", " & DBSet(miRsAux!dtoline2, "N") & ", "
        vSQL = vSQL & DBSet(vCStock.Importe, "N") & ", " & DBSet(miRsAux!origpre, "T") & ",0,null,"  '0:codprove
        vSQL = vSQL & DBSet(miRsAux!CodCCost, "T", "S") & "," & DBSet(miRsAux!idL, "N") & ")"
        conn.Execute vSQL


        
            
            
        'Actualizamos la linea de pedido
        vSQL = "update sliped set cantidad=cantidad - " & DBSet(vCStock.cantidad, "N")
        vSQL = vSQL & ",servidas=0"
        vSQL = vSQL & " WHERE numpedcl=" & miRsAux!NumPedcl & " AND numlinea = " & miRsAux!numlinea
        conn.Execute vSQL
           
            
       If miRsAux!servidas > miRsAux!cantidad Then
            vSQL = "update sliped set cantidad=0"
            vSQL = vSQL & ",servidas=0"
            vSQL = vSQL & " WHERE numpedcl=" & miRsAux!NumPedcl & " AND numlinea = " & miRsAux!numlinea
            conn.Execute vSQL
        End If
        'Sguiente linea
       miRsAux.MoveNext

    Wend
    miRsAux.Close

    
    Espera 0.25
    Aux = "select scaped.numpedcl,sum(cantidad)"
    Aux = Aux & "  from scaped,sliped WHERE scaped.numpedcl=Sliped.numpedcl AND scaped.codclien = " & Text1(0).Text
    Aux = Aux & " AND scaped.numpedcl IN (" & Mid(ListaPedidos, 2) & ") GROUP BY 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If miRsAux.Fields(1) = 0 Then
            Aux = "UPDATE scaped set sementre=52,cerrado=1 WHERE numpedcl=" & miRsAux!NumPedcl
            conn.Execute Aux
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    

    vTipoMov.IncrementarContador (codtipom)
    
    
    Espera 0.25
    Aux = "slialb.codartic=sartic.codartic and codtipom='" & codtipom & "' and numalbar"
    Aux = DevuelveDesdeBD(conAri, " sum(cantidad*(coalesce(pesoarti,0)))", "slialb,sartic", Aux, CStr(NumAlb))
    If Aux <> "" Then
        Aux = "UPDATE scaalb set pesoalba =" & DBSet(Aux, "N") & " WHERE "
        Aux = Aux & " codtipom='" & codtipom & "' and numalbar = " & NumAlb
        ejecutar Aux, False
    End If
    
    Aux = "slialb.codartic=sartic.codartic and codtipom='" & codtipom & "' and numalbar"
    
    Aux = DevuelveDesdeBD(conAri, " sum(if(slialb.codartic in ('11000','11001','11003','11007','11010','11012'),cantidad,0))", "slialb,sartic", Aux, CStr(NumAlb))
    If Aux = "0" Then Aux = ""
    If Aux <> "" Then
        Aux = TransformaPuntosComas(Aux)
        Aux = "UPDATE scaalb set numbultos =" & DBSet(Aux, "N") & " WHERE "
        Aux = Aux & " codtipom='" & codtipom & "' and numalbar = " & NumAlb
        ejecutar Aux, False
    End If
    
    
    
    
    
    RealizarAlbaran = True
    
    
eRealizarAlbaran:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, "Generando albaran"
    Set miRsAux = Nothing
    Set vCStock = Nothing
    Set vTipoMov = Nothing
End Function





Private Function BuscaHueco(hcoCodTipoM As String) As String
Dim RN As ADODB.Recordset
Dim C As String
Dim Co As Long
    BuscaHueco = ""
    C = "Select numalbar from scaalb where codtipom='" & hcoCodTipoM & "' AND year(fechaalb)=" & Year(CDate(Text1(2).Text))
    C = C & " UNION Select numalbar from scafac1 where codtipoa='" & hcoCodTipoM & "' AND year(fechaalb)=" & Year(CDate(Text1(2).Text))
    C = C & " ORDER BY numalbar DESC"
    Set RN = New ADODB.Recordset
    RN.Open C, conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not RN.EOF Then
        Co = RN.Fields(0)
        
        While Not RN.EOF
            If RN!Numalbar <> Co Then
                
                BuscaHueco = Co
                RN.MoveLast
            Else
                Co = Co - 1
            End If
            RN.MoveNext
        Wend
    End If
    RN.Close
    Set RN = Nothing
End Function
