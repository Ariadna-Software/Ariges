VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmfacFacturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturacion"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   17520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "frmfacFacturacion.frx":0000
      Left            =   6600
      List            =   "frmfacFacturacion.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdFra 
      Height          =   495
      Left            =   16680
      Picture         =   "frmfacFacturacion.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "FACTURAR"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   15000
      TabIndex        =   4
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
      ForeColor       =   &H00808080&
      Height          =   360
      ItemData        =   "frmfacFacturacion.frx":6866
      Left            =   9480
      List            =   "frmfacFacturacion.frx":686D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   13361
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Albaran"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   9825
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Obra"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Forma pago"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "B. imponible"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Origen"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   240
      Width           =   630
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   14640
      Picture         =   "frmfacFacturacion.frx":687C
      ToolTipText     =   "Buscar fecha"
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Pendiente facturar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   720
      Picture         =   "frmfacFacturacion.frx":6907
      ToolTipText     =   "Quitar seleccion"
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   240
      Picture         =   "frmfacFacturacion.frx":6A51
      ToolTipText     =   "Seleccionar todo"
      Top             =   600
      Width           =   240
   End
End
Attribute VB_Name = "frmfacFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private frmAlbG As frmFacEntAlbaranesGR
Dim cad As String
Dim primeravez As Boolean


Dim Orden As Integer
Dim Asc As Boolean

Dim Marcado As Long
Dim ImporteTot As Currency


Dim TipoAlbaSeleccionado As String


Private Sub CargaDatos(Numalbar As Long)
Dim IT As ListItem
    On Error GoTo ECargaDatos
    
    Screen.MousePointer = vbHourglass
     
    lw1.ListItems.Clear
    Marcado = 0
    ImporteTot = 0
    cad = "select c.numalbar,c.fechaalb,codclien,nomclien,substring(stippa.destippa,1,5),nomforpa ,coddirec,nomdirec,referenc,factursn,sum(importel) base"
    cad = cad & " from scaalb c inner join sforpa on c.codforpa=sforpa.codforpa"
    cad = cad & " left join stippa   on stippa.tipforpa =sforpa.tipforpa"
    cad = cad & " left join slialb l on c.codtipom=l.codtipom and c.numalbar=l.numalbar where c.codtipom='" & TipoAlbaSeleccionado & "' group by 1"
    
    cad = cad & " ORDER BY "
    Select Case Orden

    Case 2
         cad = cad & " fechaalb " & IIf(Not Asc, "DESC", "") & ", numalbar "
    Case 3
         cad = cad & " codclien " & IIf(Not Asc, "DESC", "") & ", numalbar  "
    
    Case 4
        cad = cad & " nomclien " & IIf(Not Asc, "DESC", "") & ", codclien, numalbar  "
    Case 5
        cad = cad & " nomdirec " & IIf(Not Asc, "DESC", "") & ", referenc " & IIf(Not Asc, "DESC", "") & ", codclien, numalbar  "
    Case 6
        cad = cad & " nomforpa " & IIf(Not Asc, "DESC", "") & ", fechaalb, numalbar  "
    
    Case 7
        cad = cad & " base " & IIf(Not Asc, "DESC", "") & ", fechaalb  "
    Case Else
        cad = cad & " numalbar " & IIf(Not Asc, "DESC", "") & ", fechaalb "
    End Select
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = Format(miRsAux!Numalbar, "000000")
        IT.SubItems(1) = Format(miRsAux!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(2) = Format(miRsAux!codClien, "00000")
        IT.SubItems(3) = miRsAux!NomClien
        IT.SubItems(4) = DBLet(miRsAux!nomdirec, "T")
        If IT.SubItems(4) = "" Then
            IT.SubItems(4) = DBLet(miRsAux!referenc, "T")
        Else
            IT.ListSubItems(4).ForeColor = vbBlue
            IT.ListSubItems(4).ToolTipText = "obra: " & DBLet(miRsAux!CodDirec, "T")
        End If
        IT.SubItems(5) = miRsAux!nomforpa
        IT.SubItems(6) = Format(DBLet(miRsAux!Base, "N"), FormatoImporte)
        
        
        
        
        
        IT.Checked = False
        If Val(miRsAux!factursn) = 1 Then
            Marcado = Marcado + 1
            ImporteTot = ImporteTot + DBLet(miRsAux!Base, "N")
            IT.Checked = True
        End If
        
        
        
        If miRsAux!Numalbar = Numalbar Then
            IT.Selected = True
            Set lw1.SelectedItem = IT
            PonerFocoOBj lw1
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
ECargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    Lbl
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFra_Click()
Dim Numalbar As String
Dim vCli As CCliente

    If Combo1.ListIndex = 0 Then
        MsgBox "Seleccione banco", vbExclamation
        Exit Sub
    End If
    
    Numalbar = ""
    For NumRegElim = 1 To lw1.ListItems.Count
        If lw1.ListItems(NumRegElim).Checked Then Numalbar = Numalbar & ", " & lw1.ListItems(NumRegElim).Text
    Next
     
    If Numalbar = "" Then
        MsgBox "Seleccione alguna albaran para facturar", vbExclamation
        Exit Sub
    End If
        
    Numalbar = Mid(Numalbar, 2) 'quitamos la primera coma
        
    If Text1.Text = Text1.Tag Then
        MsgBox "Seleccione fecha facturación", vbExclamation
        Exit Sub
    End If
    
     'FechaOK
    ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1.Text), True)
    If ResultadoFechaContaOK <> 0 Then
        If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
        Exit Sub
    End If
    
    
    Set miRsAux = New ADODB.Recordset
    cad = "Select codclien , nomclien from scaalb WHERE scaalb.factursn=1  and scaalb.codtipom='" & TipoAlbaSeleccionado & "' AND scaalb.numalbar in (" & Numalbar & ") GROUP BY codclien"
    Set vCli = New CCliente
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "OK"
    While Not miRsAux.EOF
        If Not vCli.LeerDatos(CStr(miRsAux!codClien)) Then
            cad = ""
        Else
            If vCli.ClienteBloqueado Then
                MsgBox "No se le puede realizar facturas : " & vbCrLf & "   -" & miRsAux!NomClien, vbExclamation
                cad = "N"
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Set vCli = Nothing
    If cad <> "OK" Then Exit Sub
    
    cad = vbCrLf & Replace(Me.label2.Caption, "Selec", "Selecionados")
    cad = "Fecha factura: " & Text1.Text & vbCrLf & "Banco: " & Combo1.Text & cad & vbCrLf & vbCrLf & "¿Desea continuar con la facturación?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    'Clientes BLOQUEADOS
    cad = "scaalb.factursn=1  and scaalb.codtipom='" & TipoAlbaSeleccionado & "' AND scaalb.numalbar in (" & Numalbar & ")"
    

    
    
    
    
    
    '------------------------------------------------------------------------------
    '  LOG de acciones.
    Set LOG = New cLOG
    cad = "Albaranes: " & Numalbar
    LOG.Insertar 2, vUsu, cad
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    
    
    Screen.MousePointer = vbHourglass
    cad = "scaalb.factursn=1  and scaalb.codtipom='" & TipoAlbaSeleccionado & "' AND scaalb.numalbar in (" & Numalbar & ")"
    TraspasoAlbaranesFacturas "Select *  FROM  scaalb WHERE " & cad, cad, Text1.Text, Combo1.ItemData(Combo1.ListIndex), Nothing, label2, True, CStr(TipoAlbaSeleccionado), "|||", CByte(vParamAplic.NumCopiasFacturacion), True, False, False
    
    DoEvents
    label2.Caption = "Leyendo BD"
    label2.Refresh
    Espera 0.75
    CargaDatos -1
    Screen.MousePointer = vbDefault
    
    
    
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 0 Then
        Combo1.ForeColor = &H808080
    Else
        Combo1.ForeColor = vbBlack
    End If
End Sub

Private Sub Combo2_Click()
    If primeravez Then Exit Sub
    If Combo2.ListIndex > 0 Then
        TipoAlbaSeleccionado = "ALZ"
        Me.label1(0).ForeColor = vbRed
    Else
        TipoAlbaSeleccionado = "ALV"
        Me.label1(0).ForeColor = 8388608
    End If
    CargaDatos -1
End Sub

Private Sub Form_activate()
    If primeravez Then
        primeravez = False
        
        CargaDatos -1
                
    End If
    
End Sub

Private Sub Form_Load()
    primeravez = True
    Me.Icon = frmPpal.Icon
   ' Me.cmdFra.Picture = frmPpal.imgListComun.ListImages(5)
    Orden = 1
    Asc = True
    TipoAlbaSeleccionado = "ALV"
    CargarCombo_Tabla Me.Combo1, "sbanpr", "codbanpr", "nombanpr", , True
    Combo1.List(0) = "Seleccione banco...."
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    
    
    Text1.Tag = "Fecha factura"
    Text1.Text = Text1.Tag
    
End Sub

Private Sub frmC_Selec(vFecha As Date)
    cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(index As Integer)

    
    cad = IIf(index = 0, "Marcar", "Desmarcar")
    cad = cad & " los albaranes para facturar?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    cad = "UPDATE scaalb set factursn = " & IIf(index = 0, "1", "0") & " WHERE codtipom = '" & TipoAlbaSeleccionado & "'"
    ejecutar cad, False
    CargaDatos -1
End Sub

Private Sub imgFecha_Click(index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    cad = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If cad <> "" Then
        Text1.Text = cad
        Text1_LostFocus
    End If
End Sub

Private Sub Label1_Click(index As Integer)
    If index = 2 Then
        If vUsu.Nivel <= 1 Then
            If Combo2.ListCount = 1 Then
                Combo2.AddItem "Presupuesto"
                HaMostradoCanal2_El_B = True
            End If
        End If
    End If
End Sub

Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.index = Orden Then
        Asc = Not Asc
    Else
        Asc = True
        Orden = ColumnHeader.index
    End If
    CargaDatos -1
End Sub

Private Sub lw1_DblClick()
     Dim N As Long
     
     If lw1.ListItems.Count = 0 Then Exit Sub
     If lw1.SelectedItem Is Nothing Then Exit Sub
     N = Val(lw1.SelectedItem)
     
        Set frmAlbG = New frmFacEntAlbaranesGR
                frmAlbG.hcoCodTipoM = IIf(Combo1.ListIndex > 0, "ALZ", "ALV")
                frmAlbG.hcoCodMovim = lw1.SelectedItem.Text
                frmAlbG.Show vbModal
                Set frmAlbG = Nothing
                    
                    
    label2.Caption = "Leyendo BD"
    label2.Refresh
    Espera 0.75
    CargaDatos N
    Screen.MousePointer = vbDefault
                
End Sub

Private Sub lw1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    cad = "UPDATE scaalb set factursn = " & IIf(Item.Checked, "1", "0") & " WHERE codtipom = '" & TipoAlbaSeleccionado & "' AND numalbar = " & Item.Text
    
    If Item.Checked Then
        Marcado = Marcado + 1
        ImporteTot = ImporteTot + ImporteFormateado(Item.SubItems(6))
    Else
        Marcado = Marcado - 1
        ImporteTot = ImporteTot - ImporteFormateado(Item.SubItems(6))
    End If
    ejecutar cad, False
    Lbl
End Sub


Private Sub Lbl()
    If Marcado > 0 Then
        label2.Caption = "Selec: " & Marcado & "     €:" & Format(ImporteTot, FormatoImporte)
    Else
        label2.Caption = ""
    End If
End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, 3
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus()
    If Text1.Text = Text1.Tag Then Exit Sub
    
    PonerFormatoFecha Text1
    If Text1.Text = "" Then
        Text1.ForeColor = &H808080
        Text1.Text = Text1.Tag
    Else
        Text1.ForeColor = vbBlack
    End If
End Sub
