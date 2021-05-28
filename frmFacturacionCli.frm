VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturacionCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación por cliente"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
   Icon            =   "frmFacturacionCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "Facturar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9465
      TabIndex        =   16
      Top             =   9255
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10905
      TabIndex        =   5
      Top             =   9255
      Width           =   1335
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7305
      Left            =   6390
      TabIndex        =   4
      Top             =   1725
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   12885
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   12095
      Begin VB.TextBox txtCopia 
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
         Left            =   10980
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   600
         Width           =   660
      End
      Begin VB.TextBox txtSitua 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text5"
         Top             =   600
         Width           =   4680
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   600
         Width           =   4980
      End
      Begin VB.TextBox txtclien 
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
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   900
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   100
         Left            =   945
         Picture         =   "frmFacturacionCli.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar cliente"
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Copias"
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
         Index           =   8
         Left            =   10980
         TabIndex        =   20
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Situacion"
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
         Index           =   2
         Left            =   6240
         TabIndex        =   11
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
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
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   315
         Width           =   975
      End
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   2520
      Left            =   120
      TabIndex        =   7
      Top             =   6465
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   4445
      _Version        =   393217
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3225
      Left            =   120
      TabIndex        =   8
      Top             =   1725
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   5689
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha Vto"
         Object.Width           =   2824
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Factura"
         Object.Width           =   2383
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "F. Factura"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Pendiente"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   11565
      Picture         =   "frmFacturacionCli.frx":0A0E
      ToolTipText     =   "Quitar seleccion"
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   11925
      Picture         =   "frmFacturacionCli.frx":0B58
      ToolTipText     =   "seleccionar todos"
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1215
      TabIndex        =   17
      Top             =   5280
      Width           =   1830
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
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
      Index           =   7
      Left            =   3585
      TabIndex        =   19
      Top             =   5310
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Pendiente"
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
      Index           =   6
      Left            =   165
      TabIndex        =   18
      Top             =   5310
      Width           =   975
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4185
      TabIndex        =   15
      Top             =   5280
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Albaranes pendientes facturar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   6180
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Albaranes para facturar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   6390
      TabIndex        =   13
      Top             =   1440
      Width           =   2385
   End
   Begin VB.Label Label1 
      Caption         =   "Cobros pendientes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1320
   End
   Begin VB.Label lblInd 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   9225
      Width           =   4335
   End
End
Attribute VB_Name = "frmFacturacionCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmCli As frmBasico2
Attribute frmCli.VB_VarHelpID = -1

Dim SQL As String
Dim Im As Currency

Private Sub cmdFacturar_Click()
Dim I As Integer

    
    

    If Me.txtclien.Text = "" Then Exit Sub
    If Me.txtNombre.Text = "" Then Exit Sub
    
    If TreeView1.Nodes.Count = 0 Then Exit Sub
    
    If txtCopia.Text = "" Then txtCopia.Text = "1"
    SQL = ""
    If Val(txtCopia.Text) > 10 Then
        SQL = "Numero copias excesivo"
    Else
        If Val(txtCopia.Text) <= 0 Then SQL = "Numero copias incorrecto"
    End If
    
    If SQL = "" Then
        SQL = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Me.txtclien.Text)
        If SQL = "1" Then
            SQL = "Cliente de varios. No se permite su facturacion por este proceso"
        Else
            SQL = ""
        End If
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        PonerFoco txtCopia
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
         MsgBox "Ninguna albarán marcado para facturar", vbExclamation
         Exit Sub
     End If



    CadenaDesdeOtroForm = ""
    frmListado2.Opcion = 25
    frmListado2.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        'OK Vamos a facturar
        Set miRsAux = Nothing
        Screen.MousePointer = vbHourglass
        HacerFacturacionCliente
        CargarDatos2
        Screen.MousePointer = vbDefault
        
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    lblInd.Caption = ""
    limpiar Me
    Set TreeView1.ImageList = frmPpal.imgListComun
    Set TreeView2.ImageList = frmPpal.imgListComun
    Set ListView1.SmallIcons = frmPpal.imgListComun
    Me.txtCopia.Text = vParamAplic.NumCopiasFacturacion
End Sub



Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtclien.Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub imgBuscarG_Click(Index As Integer)
    SQL = txtclien.Text
    Set frmCli = New frmBasico2
    AyudaClientes frmCli, SQL
    Set frmCli = Nothing
    If txtclien.Text <> SQL Then PonerFoco txtclien
        
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
    
    
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmFacEntAlbaranes2.hcoCodMovim = DevuelveNumeroAlbaran(TreeView1.SelectedItem.Text)
            frmFacEntAlbaranes2.hcoCodTipoM = Mid(TreeView1.SelectedItem.Text, 1, 3)
            frmFacEntAlbaranes2.Show vbModal
            Set frmFacEntAlbaranes2 = Nothing
        Else
            frmFacEntAlbSAIL.hcoCodMovim = DevuelveNumeroAlbaran(TreeView1.SelectedItem.Text)
            frmFacEntAlbSAIL.hcoCodTipoM = Mid(TreeView1.SelectedItem.Text, 1, 3)
            frmFacEntAlbSAIL.Show vbModal
            Set frmFacEntAlbSAIL = Nothing
        End If
        'Vuelvo a cargar los datos
        
        CargarDatos2
  
  

  
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
    If Padre Then
        J = 24
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

    
    
        If vParamAplic.TipoFormularioClientes = 0 Then
                frmFacEntAlbaranes2.hcoCodMovim = DevuelveNumeroAlbaran(TreeView2.SelectedItem.Text)
                frmFacEntAlbaranes2.hcoCodTipoM = Mid(TreeView2.SelectedItem.Text, 1, 3)
                frmFacEntAlbaranes2.Show vbModal
                Set frmFacEntAlbaranes2 = Nothing
            
        Else
        
            frmFacEntAlbSAIL.hcoCodMovim = DevuelveNumeroAlbaran(TreeView2.SelectedItem.Text)
            frmFacEntAlbSAIL.hcoCodTipoM = Mid(TreeView2.SelectedItem.Text, 1, 3)
            frmFacEntAlbSAIL.Show vbModal
            Set frmFacEntAlbSAIL = Nothing
        End If
        'Vuelvo a cargar los datos
        
        CargarDatos2
End Sub

Private Sub txtclien_GotFocus()
   ConseguirFoco txtclien, 3
End Sub

Private Sub txtclien_KeyPress(KeyAscii As Integer)
'    KEYpressGnral KeyAscii, 3, False
    If KeyAscii = teclaBuscar Then
        KEYBusquedaCli KeyAscii, 100 'cliente desde
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub KEYBusquedaCli(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscarG_Click (Indice)
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
                CargarDatos2
               
            
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

Private Sub CargarDatos2()
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset

    'Cargamos cobros pendientes
    lblInd.Caption = "Vencimientos"
    lblInd.Refresh
    CargarVtos
    
    
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
    'Todo estara en una cadena    direc|forpa|dtopp|dtogn|   Si cambia algo sera salto factura
    'antClien = 0 'cliente SIEMPRE ES EL MISMO
    'antDirec = 0 'direccion/departamento
    'antForpa = 0 'forma de pago
    'antDtoPP = 0 'dto pronto pago
    'antDtoGn = 0 'dto general
    SQL = "Select *  FROM  scaalb  WHERE "
    '(scaalb.fechaalb <= '2010-04-06') AND
    SQL = SQL & " (scaalb.codclien = " & txtclien.Text
    SQL = SQL & ") AND ( scaalb.codtipom='ALV' ) AND ( scaalb.factursn=1)  and ((scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb))"
    If vParamAplic.HayDeparNuevo > 0 Then
        SQL = SQL & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec, codforpa, dtoppago, dtognral "
    Else
        SQL = SQL & " ORDER BY scaalb.tipofact, scaalb.codclien, codforpa, dtoppago, dtognral, scaalb.coddirec "
    End If
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Anterior = ""
    NumRegElim = 1
    Set Col = New Collection
    
    While Not miRsAux.EOF
        If miRsAux!TipoFact = 1 Then
            'Factura x albaran
            
            
            'Hay que meter una factura anterior
            If Anterior <> "" Then InsertarLineaFactura Col
                
            'Meto esta
            CadenaAlbaran Col
            InsertarLineaFactura Col
            
            Anterior = ""
        Else
            SQL = CadenaIndentificacionAlbaran
            If SQL <> Anterior Then
            
                'Ha cambiado algun valor
                CadenaDesdeOtroForm = SQL
                If Anterior <> "" Then InsertarLineaFactura Col
                
                
                Anterior = CadenaDesdeOtroForm
            End If
            CadenaAlbaran Col 'Meto el albaran en el collection
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    CadenaDesdeOtroForm = ""
    If Col.Count > 0 Then InsertarLineaFactura Col
End Sub

Private Function CadenaIndentificacionAlbaran() As String
  '  direc|forpa|dtopp|dtogn|
    If vParamAplic.HayDeparNuevo > 0 Then
        CadenaIndentificacionAlbaran = Format(DBLet(miRsAux!CodDirec, "N"), "000")
    Else
        CadenaIndentificacionAlbaran = ""
    End If
    CadenaIndentificacionAlbaran = CadenaIndentificacionAlbaran & "|" & Format(DBLet(miRsAux!codforpa, "N"), "000") & "|"
    CadenaIndentificacionAlbaran = CadenaIndentificacionAlbaran & Format(miRsAux!DtoPPago * 100, "0000") & "|" & Format(miRsAux!DtoGnral * 100, "0000") & "|"
End Function

Private Sub CadenaAlbaran(ByRef Cole As Collection)
Dim C As String

    C = " codtipom = '" & miRsAux!codtipom & "' AND numalbar"
    C = DevuelveDesdeBD(conAri, "sum(importel)", "slialb", C, miRsAux!Numalbar)
    
    'Ira codtipomNumalbar sapacioblanco fecha  espacios importe
    
'-- sustituido por lo de abajo
'    Cole.Add miRsAux!codtipom & Format(miRsAux!Numalbar, "000000") & "  " & Format(miRsAux!FechaAlb, "dd/mm/yyyy") & "|" & C & "|"
    Cole.Add miRsAux!codtipom & Format(miRsAux!Numalbar, "000000") & "    " & Format(miRsAux!FechaAlb, "dd/mm/yyyy") & "|" & C & "|"
End Sub

Private Function DevuelveNumeroAlbaran(linea As String) As String
Dim J As Integer
    
    DevuelveNumeroAlbaran = "0"
    
    J = InStr(1, linea, " ")
    If J > 0 Then
        DevuelveNumeroAlbaran = Mid(linea, 1, J - 1)
        DevuelveNumeroAlbaran = Mid(DevuelveNumeroAlbaran, 4) 'los tres primeros son el codtipom
    End If
End Function


Private Sub InsertarLineaFactura(ByRef Cole As Collection)
Dim I As Integer
Dim N As Node
Dim TotalFra As Currency



    If Cole.Count = 0 Then
        'Msgbox
        'No tiene albaranes a facturar? algo raro ha pasado
        
    End If
       

    'Meto el raiz
    Set N = TreeView1.Nodes.Add(, , "FRA" & Format(NumRegElim, "000"), "Factura " & NumRegElim)
    N.Image = 43
    N.Checked = True
    TotalFra = 0
    'Los albaranes que iran
    For I = 1 To Cole.Count
        'El importe
        SQL = RecuperaValor(Cole.Item(I), 2)
        Im = CCur(SQL)
        TotalFra = TotalFra + Im
        
        'El importe
'--
'        SQL = Right(Space(10) & Format(Im, FormatoImporte), 10)
        SQL = Right(Space(18) & Format(Im, FormatoImporte), 18)
        SQL = RecuperaValor(Cole.Item(I), 1) & SQL
        Set N = TreeView1.Nodes.Add("FRA" & Format(NumRegElim, "000"), tvwChild)
        N.Text = SQL
        N.Image = 44
        N.Checked = True
        N.Tag = Im
        
        
        
    Next
'--
'    N.Parent.Text = N.Parent.Text & "   Imp: "
    N.Parent.Text = N.Parent.Text & "   Base Imponible: "
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
    SQL = SQL & ") AND ( scaalb.codtipom='ALV' ) AND ( scaalb.factursn=0)  and ((scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb))"
    SQL = SQL & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec, codforpa, dtoppago, dtognral "
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set Col = New Collection
        CadenaAlbaran Col
        
        SQL = RecuperaValor(Col.Item(1), 2)
        
        'El importe
'--sustituido por
'SQL = Right(Space(10) & Format(SQL, FormatoImporte), 10)
        SQL = Right(Space(33) & Format(SQL, FormatoImporte), 34)
        SQL = RecuperaValor(Col.Item(1), 1) & SQL
        Set N = TreeView2.Nodes.Add()
        N.Text = SQL
        N.Image = 44
            
            
        miRsAux.MoveNext
        Set Col = Nothing
    Wend
    miRsAux.Close
    
End Sub



Private Sub HacerFacturacionCliente()
Dim CadenaSQL As String
Dim I As Integer
    
    SQL = ""
    For I = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(I).Parent Is Nothing Then
            'NADA
            
        Else
            If TreeView1.Nodes(I).Checked Then SQL = SQL & ", " & DevuelveNumeroAlbaran(TreeView1.Nodes(I).Text)
   
        End If
    Next I
    
    SQL = Mid(SQL, 3)
    
    CadenaSQL = "scaalb.codtipom = 'ALV' AND scaalb.codclien=" & Me.txtclien.Text & " AND  scaalb.numalbar IN (" & SQL & ")"
    SQL = "SELECT scaalb.*,sclien.nomclien FROM scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien  WHERE " & CadenaSQL
    
    I = Val(RecuperaValor(CadenaDesdeOtroForm, 3))
    
     Dim AuxCadena As String
        
     If vParamAplic.ManipuladorFitosanitarios2 Then
        Screen.MousePointer = vbHourglass
        
        
        AuxCadena = ""
        If Not ComprobarFitosAlbaranesFacturasCliente(AuxCadena, CadenaSQL) Then AuxCadena = "NO"
        Screen.MousePointer = vbDefault
        
        If AuxCadena <> "" Then
            AuxCadena = App.Path & "\errfacFito.txt"
            
            AuxCadena = "Hay incidencias en fitosanitarios. Vea el fichero " & AuxCadena
            AuxCadena = AuxCadena & vbCrLf & vbCrLf & "¿Continuar de igualmente? "
            If MsgBox(AuxCadena, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If

    If vParamAplic.NumeroInstalacion = 2 Then
        Screen.MousePointer = vbHourglass
        AuxCadena = ""
        If Not ComprobarPrecioMinimoFacturacion(AuxCadena, CadenaSQL) Then AuxCadena = "NO"
        Screen.MousePointer = vbDefault
        If AuxCadena <> "" Then
            AuxCadena = App.Path & "\errfacFito.txt"
            
            AuxCadena = "Hay precios inferiores al precio míminmo. Ver fichero:  " & AuxCadena
            
            If vUsu.Nivel = 0 Then AuxCadena = AuxCadena & vbCrLf & vbCrLf & "¿Continuar de igual modo? "
            If MsgBox(AuxCadena, IIf(vUsu.Nivel = 0, vbQuestion + vbYesNoCancel, vbExclamation)) <> vbYes Then Exit Sub
        End If
    End If
        
    TraspasoAlbaranesFacturas SQL, CadenaSQL, RecuperaValor(CadenaDesdeOtroForm, 1), RecuperaValor(CadenaDesdeOtroForm, 2), Nothing, Me.lblInd, I = 1, "ALV", "", CByte(txtCopia.Text), True, False, False
End Sub


Private Sub txtCopia_GotFocus()
    ConseguirFoco txtCopia, 3
End Sub

Private Sub txtCopia_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCopia_LostFocus()

    If Not PonerFormatoEntero(txtCopia) Then txtCopia.Text = ""
    
End Sub
