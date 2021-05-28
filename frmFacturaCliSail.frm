VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturaCliSail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación por cliente SAIL"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   Icon            =   "frmFacturaCliSail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "&Imprimir"
      Height          =   375
      Index           =   1
      Left            =   9360
      TabIndex        =   22
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "Facturar"
      Height          =   375
      Index           =   0
      Left            =   9360
      TabIndex        =   16
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      Top             =   7920
      Width           =   1335
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6135
      Left            =   6000
      TabIndex        =   4
      Top             =   1320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10821
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.ComboBox cboTipo2 
         Height          =   315
         ItemData        =   "frmFacturaCliSail.frx":000C
         Left            =   9720
         List            =   "frmFacturaCliSail.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtFec 
         Height          =   285
         Left            =   8640
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtCopia 
         Height          =   285
         Left            =   8640
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtSitua 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text5"
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtclien 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   9
         Left            =   9720
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   0
         Left            =   9360
         Picture         =   "frmFacturaCliSail.frx":0024
         ToolTipText     =   "Buscar actividad"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Copias"
         Height          =   255
         Index           =   8
         Left            =   8640
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Situacion"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   100
         Left            =   840
         Picture         =   "frmFacturaCliSail.frx":0126
         ToolTipText     =   "Buscar actividad"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4048
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha Vto"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Factura"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "F. Factura"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Pendiente"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   23
      Top             =   7800
      Width           =   3495
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   8640
      Picture         =   "frmFacturaCliSail.frx":0228
      ToolTipText     =   "seleccionar todos"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   8280
      Picture         =   "frmFacturaCliSail.frx":0372
      ToolTipText     =   "Quitar seleccion"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
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
      Index           =   1
      Left            =   1080
      TabIndex        =   17
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   19
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Pendiente"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   3960
      TabIndex        =   15
      Top             =   4560
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Albaranes pendientes facturar"
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
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Albaranes para facturar"
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
      Index           =   4
      Left            =   6000
      TabIndex        =   13
      Top             =   1080
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Cobros pendientes"
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
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Label lblInd 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   8040
      Width           =   4095
   End
End
Attribute VB_Name = "frmFacturaCliSail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public ImprimirCertificacion As Boolean
    'Esto solo es para SAIL    'Habra dos puntos de menu:
    '               1- Facturar
    '               2- Imprimir una certifcacion
    '                   A partir de unos albaraes seleccionados mostrara una especie de "factura"
    '                   pero sin haber pasado los datos a scafac,scafac1 y slifac
    '
    
Private WithEvents frmCli As frmBasico2 'frmFacClientesGr
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Dim SQL As String
Dim Im As Currency
Dim PriVez As Boolean
    
    
    
    


Private Sub cboTipo2_Click()
   ' If Me.txtclien.Text = "" Then Exit Sub
   ' CargarDatos
End Sub

Private Sub cmdFacturar_Click(Index As Integer)
Dim I As Integer

    
    

    If Me.txtclien.Text = "" Then Exit Sub
    If Me.txtNombre.Text = "" Then Exit Sub
    
    If TreeView1.Nodes.Count = 0 Then Exit Sub
    
    If ImprimirCertificacion Then
        If txtFec.Text = "" Then
            MsgBox "Ponga una fecha", vbExclamation
            PonerFoco txtFec
            Exit Sub
        End If
    Else
        If txtCopia.Text = "" Then txtCopia.Text = "1"
        SQL = ""
        If Val(txtCopia.Text) > 10 Then
            SQL = "Numero copias excesivo"
        Else
            If Val(txtCopia.Text) <= 0 Then SQL = "Numero copias incorrecto"
        End If
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            PonerFoco txtCopia
        End If
    End If
    
    
    'Vere si hay alguno marcado para facturar
    SQL = ""
    For I = 1 To TreeView1.Nodes.Count
        If Not TreeView1.Nodes(I).Parent Is Nothing Then
            If TreeView1.Nodes(I).Checked Then
                If Not TreeView1.Nodes(I).Parent.Checked Then
                    MsgBox "Deberia estar marcado: " & TreeView1.Nodes(I).Parent.Text, vbExclamation
                    TreeView1.Nodes(I).Parent.Checked = True
                    Exit Sub
                End If
                
                SQL = "OK"
                Exit For
            End If
        End If
    Next
    If SQL = "" Then
        If ImprimirCertificacion Then
            SQL = "imprimir certificación"
        Else
            SQL = "facturar"
        End If
        MsgBox "Ninguna albarán marcado para " & SQL, vbExclamation
        Exit Sub
    End If


    'Si hay alguno para facturar compruebo que todas las facturas tengan al menos un albaran seleccionado
    For I = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(I).Parent Is Nothing Then
            If TreeView1.Nodes(I).Checked Then
                'Comprobar si algun nodo esta seleccionado
                If Not ComprobarSubNodoSeleccionado(I) Then Exit Sub
            End If
        End If
    Next


    'SI NO ES ALBARAN VENTA, preguntaremos
    If InstalacionEsEulerTaxco Then
    
        If Not HacerComprobarProyectosEuler Then Exit Sub
    
    
    
        If Me.cboTipo2.ItemData(Me.cboTipo2.ListIndex) > 0 Then
            CadenaDesdeOtroForm = String(43, "**") & vbCrLf
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & "Va a generar factura    " & UCase(Me.cboTipo2.Text) & vbCrLf & vbCrLf & CadenaDesdeOtroForm
            MsgBox CadenaDesdeOtroForm, vbInformation
        End If
        
        
    End If
    
    'AQUI segun sea Imprimir o facturar hara unas cosas U otras
    CadenaDesdeOtroForm = ""
    
    If ImprimirCertificacion Then
        'IMPRIMIR CERTIF
        ImprimirCerti
    Else
        'FACTURAR
        frmListado2.Opcion = 25
        frmListado2.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            'OK Vamos a facturar
            Set miRsAux = Nothing
            Screen.MousePointer = vbHourglass
            HacerFacturacion
            CargarDatos
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PriVez Then
        PriVez = False
        
        If Not ImprimirCertificacion Then CargaComboTipos
        If InstalacionEsEulerTaxco Then
           If Me.cboTipo2.ListCount > 0 Then Me.cboTipo2.ListIndex = 0
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    lblInd.Caption = ""
    PriVez = True
    txtclien.Text = ""
    If Me.ImprimirCertificacion Then
        Label2.Caption = "Certificación"
        Label2.ForeColor = &H80&
        SQL = "Impresión certificación"
        'Fecha "certificacion"
        Label1(8).Caption = "F. Certif."
        Me.txtFec = Format(Now, "dd/mm/yyyy")
        
    Else
        Label2.Caption = "Facturar"
        Label2.ForeColor = vbRed
        SQL = "Facturar Albaranes x Cliente"
        Label1(8).Caption = "Copias"
    End If
    Me.Caption = SQL
    Me.cmdFacturar(1).visible = ImprimirCertificacion
    Me.cmdFacturar(0).visible = Not ImprimirCertificacion
    
    'Tipo Albaran para la certificacion
    If ImprimirCertificacion Then CargaComboTipos
    Me.Label1(9).visible = True ' ImprimirCertificacion
    cboTipo2.visible = True 'ImprimirCertificacion2
   
    
    
    'nºCopias
    txtCopia.visible = Not ImprimirCertificacion
    txtFec.visible = ImprimirCertificacion
    imgBuscarG(0).visible = ImprimirCertificacion
    
    limpiar Me
    Set TreeView1.ImageList = frmPpal.imgListComun
    Set TreeView2.ImageList = frmPpal.imgListComun
    Set ListView1.SmallIcons = frmPpal.imgListComun
    Me.txtCopia.Text = vParamAplic.NumCopiasFacturacion
End Sub



Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtclien.Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Me.txtFec.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgBuscarG_Click(Index As Integer)

    If Index = 0 Then
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Me.txtFec.Text <> "" Then frmF.Fecha = CDate(txtFec.Text)
        frmF.Show vbModal
        Set frmF = Nothing
        
    Else
        SQL = txtclien.Text
'        Set frmCli = New frmFacClientesGr
'        frmCli.DatosADevolverBusqueda = "0|1|"
'        frmCli.Show vbModal
        Set frmCli = New frmBasico2
        AyudaClientes frmCli, txtclien.Text
        Set frmCli = Nothing
        If txtclien.Text <> SQL Then
            PonerFoco txtclien
            txtclien_LostFocus
        End If
    End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
    For NumRegElim = 1 To TreeView1.Nodes.Count
        TreeView1.Nodes(NumRegElim).Checked = Index = 1
    Next
End Sub

Private Sub TreeView1_DblClick()

    If TreeView1.Nodes.Count = 0 Then Exit Sub
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If TreeView1.SelectedItem.Parent Is Nothing Then Exit Sub
    
        NumRegElim = InStr(1, TreeView1.SelectedItem.Text, " ") - 4   'Numero de albaran
        NumRegElim = Val(Mid(TreeView1.SelectedItem.Text, 4, NumRegElim))
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmFacEntAlbaranes2.hcoCodMovim = NumRegElim
            frmFacEntAlbaranes2.hcoCodTipoM = Mid(TreeView1.SelectedItem.Text, 1, 3)
            frmFacEntAlbaranes2.Show vbModal
            Set frmFacEntAlbaranes2 = Nothing
        Else
            frmFacEntAlbSAIL.hcoCodMovim = NumRegElim
            frmFacEntAlbSAIL.hcoCodTipoM = Mid(TreeView1.SelectedItem.Text, 1, 3)
            frmFacEntAlbSAIL.Show vbModal
            Set frmFacEntAlbSAIL = Nothing
        End If
        'Vuelvo a cargar los datos
        
        CargarDatos
  
  

  
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim N As Node
Dim Im As Currency

    If Node.Parent Is Nothing Then
        'Ha checkeado(quitado) uno padre. Todos los hijos haran los mismo
        Set N = Node.Child
        'If Node.Checked = False Then N.Tag = 0
            
        Im = 0
        While Not N Is Nothing
            
            N.Checked = N.Parent.Checked
            If N.Checked Then Im = Im + N.Tag
            Set N = N.Next
            
            
        Wend
        Node.Tag = Im
        PonerCadenaImporte Node, True
    Else
        If Node.Checked Then
            Node.Parent.Tag = Node.Parent.Tag + Node.Tag
        Else
            Node.Parent.Tag = Node.Parent.Tag - Node.Tag
        End If
        PonerCadenaImporte Node.Parent, True
        
        
        
        'Comprobare que si hay marcado alguno el ppal este maracdo y al reves
        Im = 0
        Set N = Node.FirstSibling
        
            
        Im = 0  'Ninguno chekeado
        While Not N Is Nothing
            
            If N.Checked Then Im = Im + 1
            Set N = N.Next
            
            
        Wend
        Node.Parent.Checked = Im > 0
    End If
    
    
                        
End Sub

Private Sub PonerCadenaImporte(ByRef N As Node, Padre As Boolean)
Dim I As Integer
Dim J As Integer
    'Silo puede llevar los dops puntos UNA vez

    If Padre Then
        J = 10
    Else
        J = 45
    End If
    I = InStr(1, N.Text, ":")
    If I > 0 Then
        N.Text = Mid(N.Text, 1, I)
        N.Text = N.Text & Right(Space(J) & Format(N.Tag, FormatoImporte), J)
    End If
End Sub

Private Sub TreeView2_DblClick()
    If TreeView2.Nodes.Count = 0 Then Exit Sub
    If TreeView2.SelectedItem Is Nothing Then Exit Sub

    
        NumRegElim = InStr(1, TreeView2.SelectedItem.Text, " ") - 4   'Numero de albaran
        NumRegElim = Val(Mid(TreeView2.SelectedItem.Text, 4, NumRegElim))
        If vParamAplic.TipoFormularioClientes = 0 Then
                frmFacEntAlbaranes2.hcoCodMovim = NumRegElim
                frmFacEntAlbaranes2.hcoCodTipoM = Mid(TreeView2.SelectedItem.Text, 1, 3)
                frmFacEntAlbaranes2.Show vbModal
                Set frmFacEntAlbaranes2 = Nothing
            
        Else
        
            frmFacEntAlbSAIL.hcoCodMovim = NumRegElim
            frmFacEntAlbSAIL.hcoCodTipoM = Mid(TreeView2.SelectedItem.Text, 1, 3)
            frmFacEntAlbSAIL.Show vbModal
            Set frmFacEntAlbSAIL = Nothing
        End If
        'Vuelvo a cargar los datos
        
        CargarDatos
End Sub

Private Sub txtclien_GotFocus()
   ConseguirFoco txtclien, 3
End Sub

Private Sub txtclien_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtclien_LostFocus()


    SQL = ""
    txtclien.Text = Trim(txtclien.Text)
    txtSitua.Text = ""
    txtNombre.Text = ""
    If txtclien.Text <> "" Then
    
        If PonerFormatoEntero(txtclien) Then
            
            Set miRsAux = New ADODB.Recordset
            SQL = PonerCliente
            If SQL = "" Then
                MsgBox "No existe el cliente: " & txtclien.Text, vbExclamation
                txtclien.Text = ""
                PonerFoco txtclien
            Else
                'Cargar DATOS
                CargaDatosPpal
                'If vParamAplic.NumeroInstalacion <> 4 Then SQL = ""
                
            End If

        End If
    End If
    If SQL = "" Then
        Me.ListView1.ListItems.Clear
        Me.TreeView1.Nodes.Clear
        Me.TreeView2.Nodes.Clear
        lblTot(0).Caption = ""
        lblTot(1).Caption = ""
    End If
    
    
End Sub

Private Function PonerCliente() As String
    Set miRsAux = New ADODB.Recordset
    SQL = "Select nomclien,nomsitua,codmacta from sclien,ssitua WHERE sclien.codsitua=ssitua.codsitua AND codclien=" & txtclien.Text
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not miRsAux.EOF Then
        SQL = miRsAux!NomClien
        Me.txtNombre.Text = SQL
        PonerCliente = miRsAux!NomClien
        txtSitua = miRsAux!nomsitua
        txtSitua.Tag = miRsAux!Codmacta
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Function


Private Sub CargaDatosPpal()
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset

    'Cargamos cobros pendientes
    lblInd.Caption = "Vencimientos"
    lblInd.Refresh
    CargarVtos
    
    
    'CargaComboTipos
    CargarDatos
    
    Screen.MousePointer = vbDefault
    Set miRsAux = Nothing
    


End Sub

Private Sub CargarDatos()
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset

    
    
    'Cargamos albarananes pendientes de facturar
    lblInd.Caption = "Albaranes pendientes facturar"
    lblInd.Refresh
    CargaAlbaranes
    
    
    lblInd.Caption = "Albaranes sin marca facturar"
    lblInd.Refresh
    CargaAlbaranesSin
    
    
    lblInd.Caption = ""
    Screen.MousePointer = vbDefault
    Set miRsAux = Nothing
    
End Sub


Private Sub CargarVtos()
Dim IT As ListItem
Dim Im2 As Currency
Dim Pend As Currency

    ListView1.ListItems.Clear
     
    If vParamAplic.ContabilidadNueva Then
        SQL = "SELECT ImpVenci,gastos,impcobro,FecVenci,numSerie,numfactu codfaccl,fecfactu fecfaccl FROM cobros scobro INNER JOIN formapago ON scobro.codforpa=formapago.codforpa  "
    Else
        SQL = "SELECT scobro.* FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
    End If
    
    
    SQL = SQL & " WHERE scobro.codmacta = '" & txtSitua.Tag & "'"
    'SQL = SQL & " AND fecvenci <= ' " & Format(Now, FormatoFecha) & "' "
    ' SQL = SQL & " AND (sforpa.tipforpa between 0 and 3)
    SQL = SQL & " ORDER BY fecvenci"
    miRsAux.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Im = 0
    Pend = 0
    While Not miRsAux.EOF
        Im2 = miRsAux!ImpVenci + DBLet(miRsAux!gastos, "N") - DBLet(miRsAux!impcobro, "N")
        If Im2 <> 0 Then
    
            Set IT = ListView1.ListItems.Add()
            IT.Text = miRsAux!FecVenci
            IT.SmallIcon = 23
            'If miRsAux!FecVenci > Now Then
            IT.SubItems(1) = miRsAux!numSerie & Format(miRsAux!Codfaccl, "00000")
            IT.SubItems(2) = Format(miRsAux!fecfaccl, "dd/mm/yyyy")
            
            
            IT.SubItems(3) = Format(Im2, FormatoImporte)
            Im = Im + Im2
            If miRsAux!FecVenci < Now Then Pend = Pend + Im2
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If ListView1.ListItems.Count = 0 Then
        lblTot(1).Caption = ""
        lblTot(0).Caption = ""
    Else
        lblTot(0).Caption = Format(Im, FormatoImporte)
        lblTot(1).Caption = Format(Pend, FormatoImporte)
    End If
End Sub



Private Sub CargaAlbaranes()
Dim Anterior As String
Dim Col As Collection
    TreeView1.Nodes.Clear
    'SAIL.  LLeva la actuacion TB
    'Todo estara en una cadena    direc|actuacion|forpa|dtopp|dtogn|   Si cambia algo sera salto factura
    'antClien = 0 'cliente SIEMPRE ES EL MISMO
    'antDirec = 0 'direccion/departamento
    
    ''SAIL
    'Llevara actuacion y salto por actuacion
    
    
    'antForpa = 0 'forma de pago
    'antDtoPP = 0 'dto pronto pago
    'antDtoGn = 0 'dto general
    
    
    
    SQL = "Select *  FROM  scaalb  WHERE "
    '(scaalb.fechaalb <= '2010-04-06') AND
    SQL = SQL & " (scaalb.codclien = " & txtclien.Text
    SQL = SQL & ") AND ( scaalb.codtipom IN ("
    SQL = SQL & DevuelveTipoDocumento2
    SQL = SQL & ")) AND ( scaalb.factursn=1)  and ((scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb))"
    
 
    SQL = SQL & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec"
    SQL = SQL & ",actuacion"
    SQL = SQL & ", codforpa, dtoppago, dtognral "
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Anterior = ""
    NumRegElim = 1
    Set Col = New Collection
    
    While Not miRsAux.EOF
    
        If miRsAux!TipoFact = 1 Then
            'Factura x albaran
            
            
            'Hay que meter una factura anterior
            If Anterior <> "" Then InsertarLineaFactura Col, Anterior
                
            'Meto esta
            SQL = CadenaIndentificacionAlbaran
            CadenaAlbaran Col
            InsertarLineaFactura Col, SQL
            
            Anterior = ""
        Else
            SQL = CadenaIndentificacionAlbaran
            If SQL <> Anterior Then
                'Ha cambiado algun valor
                If Anterior <> "" Then InsertarLineaFactura Col, Anterior
                
                
                Anterior = SQL
            End If
            CadenaAlbaran Col 'Meto el albaran en el collection
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Col.Count > 0 Then InsertarLineaFactura Col, Anterior
End Sub

Private Function CadenaIndentificacionAlbaran() As String
  '  direc|forpa|dtopp|dtogn|
    CadenaIndentificacionAlbaran = Format(DBLet(miRsAux!CodDirec, "N"), "0000") & "|" & UCase(DBLet(miRsAux!actuacion, "T")) & "|"
    CadenaIndentificacionAlbaran = CadenaIndentificacionAlbaran & Format(DBLet(miRsAux!codforpa, "N"), "000") & "|"
    CadenaIndentificacionAlbaran = CadenaIndentificacionAlbaran & Format(miRsAux!DtoPPago * 100, "0000") & "|" & Format(miRsAux!DtoGnral * 100, "0000") & "|"
End Function

Private Sub CadenaAlbaran(ByRef Cole As Collection)
Dim C As String

    C = " codtipom = '" & miRsAux!codtipom & "' AND numalbar"
    C = DevuelveDesdeBD(conAri, "sum(importel)", "slialb", C, miRsAux!Numalbar)
    
    'Ira codtipomNumalbar sapacioblanco fecha  espacios importe
    Cole.Add miRsAux!codtipom & Format(miRsAux!Numalbar, "000000") & "  " & Format(miRsAux!FechaAlb, "dd/mm/yyyy") & "|" & C & "|"
    
End Sub

Private Function DevuelveNumeroAlbaran(linea As String) As String
Dim J As Integer
    
    
    

        DevuelveNumeroAlbaran = "('PPP',0)"
        
        J = InStr(1, linea, " ")
        If J > 0 Then
            DevuelveNumeroAlbaran = Mid(linea, 1, J - 1)
            DevuelveNumeroAlbaran = "('" & Mid(DevuelveNumeroAlbaran, 1, 3) & "'," & Mid(DevuelveNumeroAlbaran, 4) & ")" 'los tres primeros son el codtipom
        End If


End Function


Private Sub InsertarLineaFactura(ByRef Cole As Collection, CadenaFactura As String)
Dim I As Integer
Dim N As Node
Dim TotalFra As Currency
Dim Aux As String


    If Cole.Count = 0 Then
        'Msgbox
        'No tiene albaranes a facturar? algo raro ha pasado
        
    End If
       

    'Meto el raiz
    
    
    ' "  " & Format(SQL, "000") & "-" & SQL
    Set N = TreeView1.Nodes.Add(, , "FRA" & Format(NumRegElim, "000"), "F" & Format(NumRegElim, "00"))
    
    'Añado al tooltip el departamento
    Aux = RecuperaValor(CadenaFactura, 1)
    Aux = "  Ob" & Format(Aux, "000") & "- "
    N.Text = N.Text & Aux
    Aux = RecuperaValor(CadenaFactura, 2)
    Aux = Mid(Aux & Space(20), 1, 20)
    N.Text = N.Text & Aux
    
    
    N.Image = 43
    N.Checked = True
    TotalFra = 0
    'Los albaranes que iran
    For I = 1 To Cole.Count
        'El importe
        Aux = RecuperaValor(Cole.Item(I), 2)
        Im = CCur(Aux)
        TotalFra = TotalFra + Im
        
        'El importe
        Aux = Right(Space(10) & Format(Im, FormatoImporte), 10)
        Aux = RecuperaValor(Cole.Item(I), 1) & Aux
        Set N = TreeView1.Nodes.Add("FRA" & Format(NumRegElim, "000"), tvwChild)
        N.Text = Aux
        N.Image = 44
        N.Checked = True
        N.Tag = Im
        
        
        
    Next
    N.Parent.Text = N.Parent.Text & " Imp: "
    N.Parent.Tag = TotalFra
    PonerCadenaImporte N.Parent, True
    
    N.Parent.Expanded = True
    NumRegElim = NumRegElim + 1
    Set Cole = Nothing
    Set Cole = New Collection
End Sub



Private Sub CargaAlbaranesSin()
Dim Col As Collection
Dim N As Node

    TreeView2.Nodes.Clear
    
    SQL = "Select *  FROM  scaalb  WHERE "
    '(scaalb.fechaalb <= '2010-04-06') AND
    SQL = SQL & " (scaalb.codclien = " & txtclien.Text
    SQL = SQL & ") AND ( scaalb.codtipom IN ("
    SQL = SQL & DevuelveTipoDocumento2
    
    SQL = SQL & ")) AND ( scaalb.factursn=0)  and ((scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb))"
    SQL = SQL & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec, codforpa, dtoppago, dtognral "
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set Col = New Collection
        CadenaAlbaran Col
        
        SQL = RecuperaValor(Col.Item(1), 2)
        
        'El importe
        SQL = Right(Space(10) & Format(SQL, FormatoImporte), 10)
        SQL = RecuperaValor(Col.Item(1), 1) & SQL
        Set N = TreeView2.Nodes.Add()
        N.Text = SQL
        N.Image = 44
            
            
        miRsAux.MoveNext
        Set Col = Nothing
    Wend
    miRsAux.Close
    
End Sub


Private Sub HacerFacturacion()
Dim NO As Node
Dim Aux As String 'por si acaso pierde datos en cadenadesdeotroform
Dim OK As Byte
Dim Mal As Byte
    OK = 0: Mal = 0
    Set NO = TreeView1.Nodes(1)
    Aux = CadenaDesdeOtroForm
    Do
        CadenaDesdeOtroForm = Aux
        If NO.Checked Then
            If HacerFacturacionClienteSAIL(NO.Index) Then
                OK = OK + 1
            Else
                Mal = Mal + 1
                lblInd.Caption = "Error factura: " & NO.Index
                lblInd.Refresh
                DoEvents
                Espera 0.5
            End If
        End If
        Set NO = NO.Next
    Loop Until NO Is Nothing
    
    lblInd.Caption = "Proceso finalizado"
    lblInd.Refresh
    
    If Mal > 0 Then
        Aux = "Proceso facturacion " & vbCrLf & String(20, "=") & vbCrLf & vbCrLf & "Correctas= " & OK & vbCrLf & "Errores= " & Mal
        MsgBox Aux, vbExclamation
    Else
        MsgBox "Proceso finalizado", vbInformation
    End If
    
    Espera 0.2
End Sub


Private Function HacerFacturacionClienteSAIL(ind As Integer) As Boolean
Dim CadenaSQL As String
Dim N As Node
Dim TipoFactura As String
    
    lblInd.Caption = TreeView1.Nodes(ind).Text
    lblInd.Refresh
    HacerFacturacionClienteSAIL = True
    SQL = ""
    Set N = TreeView1.Nodes(ind).Child
    Do
        If N.Checked Then SQL = SQL & ", " & DevuelveNumeroAlbaran(N.Text)
        Set N = N.Next
        
    Loop Until N Is Nothing
    If SQL <> "" Then
        SQL = Mid(SQL, 3)
        
        
        
        If InstalacionEsEulerTaxco Then
            TipoFactura = Replace(DevuelveTipoDocumento2, ",", "|") & "|" 'Coge todo
            TipoFactura = Replace(TipoFactura, "'", "")   'quita comillas
            TipoFactura = RecuperaValor(TipoFactura, Me.cboTipo2.ItemData(cboTipo2.ListIndex))
        Else
            'SAIL
            TipoFactura = DevuelveTipoDocumento2
            TipoFactura = Replace(TipoFactura, "'", "")
        End If
        
        CadenaSQL = " (scaalb.codtipom,scaalb.numalbar) IN (" & SQL & ") AND scaalb.codclien=" & Me.txtclien.Text
        SQL = "SELECT scaalb.*,sclien.nomclien FROM scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien  WHERE " & CadenaSQL
        
        NumRegElim = Val(RecuperaValor(CadenaDesdeOtroForm, 3))
        If Not TraspasoAlbaranesFacturasCliente(SQL, CadenaSQL, RecuperaValor(CadenaDesdeOtroForm, 1), RecuperaValor(CadenaDesdeOtroForm, 2), Nothing, Me.lblInd, NumRegElim = 1, TipoFactura, "", CByte(txtCopia.Text), False) Then
            'Ha ido bien
            HacerFacturacionClienteSAIL = False 'Ha ido mal
        End If
    End If
End Function



Private Function HacerComprobarProyectosEuler() As Boolean
Dim N As Node

    HacerComprobarProyectosEuler = True
    Set N = TreeView1.Nodes(1)
    Do
        lblInd.Caption = "Comprobar: " & N.Text
        lblInd.Refresh
        If N.Checked Then
            If Not HacerComprobarProyectosEuler_nodo(N.Index) Then HacerComprobarProyectosEuler = False
        End If
        Set N = N.Next
        
    Loop Until N Is Nothing
    
    lblInd.Caption = ""
    
End Function

Private Function HacerComprobarProyectosEuler_nodo(ind As Integer) As Boolean
 Dim N As Node
 
    lblInd.Caption = TreeView1.Nodes(ind).Text
    lblInd.Refresh
    HacerComprobarProyectosEuler_nodo = True
    SQL = ""
    Set N = TreeView1.Nodes(ind).Child
    Do
        If N.Checked Then SQL = SQL & ", " & DevuelveNumeroAlbaran(N.Text)
        Set N = N.Next
        
    Loop Until N Is Nothing
    If SQL <> "" Then
        SQL = Mid(SQL, 3)
        
        
        
        SQL = " (codtipoa,numalbar) IN (" & SQL & ")  AND 1"
        SQL = DevuelveDesdeBD(conAri, "numproyec", "sproyectolin", SQL, "1")
        If Val(SQL) > 0 Then
            SQL = "Albaranes para la factura: " & TreeView1.Nodes(ind).Text & vbCrLf & " vinculados en el proyecto: " & SQL
            MsgBox SQL, vbExclamation
            HacerComprobarProyectosEuler_nodo = False
        End If

    End If
End Function






Private Sub txtCopia_GotFocus()
    ConseguirFoco txtCopia, 3
End Sub

Private Sub txtCopia_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCopia_LostFocus()

    If Not PonerFormatoEntero(txtCopia) Then txtCopia.Text = ""
    
End Sub

Private Function ComprobarSubNodoSeleccionado(Indi As Integer) As Boolean
Dim NO As Node
Dim AlgunoSeleccionado As Boolean

       Set NO = TreeView1.Nodes(Indi).Child
       
       Do
            If NO.Checked Then
                AlgunoSeleccionado = True
                Set NO = Nothing
            Else
                Set NO = NO.Next
            End If
        Loop Until NO Is Nothing
        ComprobarSubNodoSeleccionado = AlgunoSeleccionado
        If Not AlgunoSeleccionado Then MsgBox "Existe una factura sin seleccionar albaranes: " & vbCrLf & TreeView1.Nodes(Indi).Text, vbExclamation
            
End Function


Private Sub ImprimirCerti()
Dim I As Byte
Dim NO As Node

    SQL = "DELETE from tmpnlotes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    
    Set NO = TreeView1.Nodes(1)
    Do
        If NO.Checked Then
            'Insertamos albaranes
            If Not InsertarAlbaranesCertificacion(NO.Index, Val(Mid(NO.Key, 4))) Then
                Set NO = Nothing
                Exit Sub
            End If
        End If
        Set NO = NO.Next
    Loop Until NO Is Nothing
    
    
    'Llegados aquin vemos si hay alguno reg insertado
    SQL = DevuelveDesdeBD(conAri, "count(*)", "tmpnlotes", "codusu", CStr(vUsu.Codigo))
    If SQL = "" Then SQL = "0"
    If Val(SQL) = 0 Then
        MsgBox "Ningun registro generado", vbExclamation
        Exit Sub
    End If
    
    If PonerParamRPT2(47, SQL, I, CadenaDesdeOtroForm, False, "", pRptvMultiInforme) Then
        SQL = SQL & "|pFecha=""" & txtFec.Text & """|"
        I = I + 1
        With frmImprimir
            .ConSubInforme = False
            .FormulaSeleccion = "{tmpnlotes.codusu} = " & vUsu.Codigo
            .NombreRPT = CadenaDesdeOtroForm
            .NombrePDF = pPdfRpt
            .Titulo = "Certificación"
            .OtrosParametros = SQL
            .NumeroParametros = I
            .SeleccionaRPTCodigo = pRptvMultiInforme
            .Opcion = 2003 'Esta libre
            .Show vbModal
        End With
    End If
End Sub


Private Function InsertarAlbaranesCertificacion(Indi As Integer, NumeroFac As Integer) As Boolean
Dim NO As Node
Dim Aux As String
Dim I As Integer
       Set NO = TreeView1.Nodes(Indi).Child
       '                                    nsecuen             nºALBA      codtipom
       'insert into `tmpnlotes` (`codusu`,`numalbar`,`codprove`,`numlotes`) values
       InsertarAlbaranesCertificacion = False
       SQL = ""
       
       
       Do
            If NO.Checked Then
                'Insertamos albaranes
                Aux = DevuelveNumeroAlbaran(NO.Text)
                Aux = Mid(Aux, InStr(1, Aux, ",") + 1)
                Aux = Mid(Aux, 1, Len(Aux) - 1)
                
                SQL = SQL & ", (" & vUsu.Codigo & "," & NumeroFac & "," & Aux
                Aux = Mid(NO.Text, 1, 3)
                SQL = SQL & ",'" & Aux & "')"
                
            End If
            Set NO = NO.Next
        Loop Until NO Is Nothing
        
        If SQL <> "" Then
            SQL = Mid(SQL, 2) 'quito la coma
            SQL = "insert into `tmpnlotes` (`codusu`,`numalbar`,`codprove`,`numlotes`) values " & SQL
            If Not ejecutar(SQL, False) Then Exit Function
        End If
        InsertarAlbaranesCertificacion = True
            
End Function



Private Sub txtFec_GotFocus()
    ConseguirFoco txtCopia, 3
End Sub

Private Sub txtFec_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtFec_LostFocus()
    txtFec.Text = Trim(txtFec.Text)
    If txtFec.Text <> "" Then
        PonerFormatoFecha txtFec
    End If
End Sub
 

Private Function DevuelveTipoDocumento2() As String

    If InstalacionEsEulerTaxco Then
        DevuelveTipoDocumento2 = "'ALV','ALR','ALO','ALE'"
    Else
        If cboTipo2.ListIndex = -1 Then
            DevuelveTipoDocumento2 = "'PPP'"
        Else
            If ImprimirCertificacion Then
                DevuelveTipoDocumento2 = "'" & RecuperaValor("ALV|ALZ|", cboTipo2.ListIndex + 1) & "'"
            Else
                DevuelveTipoDocumento2 = "'" & RecuperaValor("ALV|ALR|ALO|ALE|", cboTipo2.ItemData(cboTipo2.ListIndex)) & "'"
            End If
            
        End If
    End If
End Function


Private Sub CargaComboTipos()
Dim I As Byte


    
    
    If ImprimirCertificacion Then
    
        If cboTipo2.ListCount = 0 Then
            cboTipo2.AddItem "Venta"
            cboTipo2.ItemData(cboTipo2.NewIndex) = 1
            cboTipo2.AddItem "Presu"
            cboTipo2.ItemData(cboTipo2.NewIndex) = 2
            cboTipo2.ListIndex = 0
        Else
            If vParamAplic.NumeroInstalacion <> 4 Then cboTipo2.ListIndex = 0 'SAIL
        End If
    Else
        Me.cboTipo2.Clear
        

        If vParamAplic.NumeroInstalacion <> 4 Then
            
            cboTipo2.AddItem "VENTA"
            cboTipo2.ItemData(cboTipo2.NewIndex) = 1
            cboTipo2.ListIndex = 0
        Else
        
            SQL = "Venta|Reparación|Orden trabajo|Trabajo externo|"
            For I = 1 To 4
                    cboTipo2.AddItem RecuperaValor(SQL, CInt(I))
                    cboTipo2.ItemData(cboTipo2.NewIndex) = I
            Next
            
'            SQL = "Select codtipom from scaalb where codclien=" & txtclien.Text & " AND codtipom IN ('ALV','ALR','ALO','ALE') GROUP BY 1"
'            miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            SQL = "Venta|Reparación|Orden trabajo|Trabajo externo|"
'
'            While Not miRsAux.EOF
'                Select Case CStr(miRsAux!codtipom)
'                Case "ALV"
'                    i = 1
'                Case "ALR"
'                    i = 2
'                Case "ALO"
'                    i = 3
'                Case "ALE"
'                    i = 4
'                End Select
'
'                If i > 0 Then
'                    cboTipo2.AddItem RecuperaValor(SQL, CInt(i))
'                    cboTipo2.ItemData(cboTipo2.NewIndex) = i
'                End If
'                miRsAux.MoveNext
'            Wend
'            miRsAux.Close
        End If
 
    End If
End Sub
