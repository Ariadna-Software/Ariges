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
   Begin VB.CommandButton cmdFra 
      Height          =   495
      Left            =   16800
      Picture         =   "frmfacFacturacion.frx":0000
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
      ItemData        =   "frmfacFacturacion.frx":6852
      Left            =   9480
      List            =   "frmfacFacturacion.frx":6859
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   13785
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
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   14640
      Picture         =   "frmfacFacturacion.frx":6868
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
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Albaranes para facturar"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   5040
      Picture         =   "frmfacFacturacion.frx":68F3
      ToolTipText     =   "Quitar seleccion"
      Top             =   240
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   4680
      Picture         =   "frmfacFacturacion.frx":6A3D
      ToolTipText     =   "Seleccionar todo"
      Top             =   240
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


Dim cad As String
Dim PrimeraVez As Boolean


Dim Orden As Integer
Dim Asc As Boolean

Dim Marcado As Long
Dim ImporteTot As Currency


Private Sub CargaDatos()
Dim IT As ListItem
    On Error GoTo ECargaDatos
    
    Screen.MousePointer = vbHourglass
    
    lw1.ListItems.Clear
    Marcado = 0
    ImporteTot = 0
    cad = "select c.numalbar,c.fechaalb,codclien,nomclien,substring(stippa.destippa,1,5),nomforpa ,nomdirec,factursn,sum(importel) base"
    cad = cad & " from scaalb c inner join sforpa on c.codforpa=sforpa.codforpa"
    cad = cad & " left join stippa   on stippa.tipforpa =sforpa.tipforpa"
    cad = cad & " left join slialb l on c.codtipom=l.codtipom and c.numalbar=l.numalbar where c.codtipom='ALV' group by 1"
    
    cad = cad & " ORDER BY "
    Select Case Orden

    Case 2
         cad = cad & " fechaalb " & IIf(Not Asc, "DESC", "") & ", numalbar "
    Case 3
         cad = cad & " codclien " & IIf(Not Asc, "DESC", "") & ", numalbar  "
    
    Case 4
        cad = cad & " nomclien " & IIf(Not Asc, "DESC", "") & ", codclien, numalbar  "
    Case 5
        cad = cad & " nomdirec " & IIf(Not Asc, "DESC", "") & ", codclien, numalbar  "
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
        IT.Text = Format(miRsAux!NUmAlbar, "000000")
        IT.SubItems(1) = Format(miRsAux!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(2) = Format(miRsAux!codClien, "00000")
        IT.SubItems(3) = miRsAux!NomClien
        IT.SubItems(4) = DBLet(miRsAux!nomdirec, "T")
        IT.SubItems(5) = miRsAux!nomforpa
        IT.SubItems(6) = Format(DBLet(miRsAux!Base, "N"), FormatoImporte)
        
        
        
        
        
        IT.Checked = False
        If Val(miRsAux!factursn) = 1 Then
            Marcado = Marcado + 1
            ImporteTot = ImporteTot + DBLet(miRsAux!Base, "N")
            IT.Checked = True
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
Dim NUmAlbar As String
    If Combo1.ListIndex = 0 Then
        MsgBox "Seleccione banco", vbExclamation
        Exit Sub
    End If
    
    NUmAlbar = ""
    For NumRegElim = 1 To lw1.ListItems.Count
        If lw1.ListItems(NumRegElim).Checked Then NUmAlbar = NUmAlbar & ", " & lw1.ListItems(NumRegElim).Text
    Next
     
    If NUmAlbar = "" Then
        MsgBox "Seleccione alguna albaran para facturar", vbExclamation
        Exit Sub
    End If
        
    NUmAlbar = Mid(NUmAlbar, 2) 'quitamos la primera coma
        
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
    
    
    
    
    
    cad = vbCrLf & Replace(Me.Label2.Caption, "Selec", "Selecionados")
    cad = "Fecha factura: " & Text1.Text & vbCrLf & "Banco: " & Combo1.Text & cad & vbCrLf & vbCrLf & "¿Desea continuar con la facturación?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    '------------------------------------------------------------------------------
    '  LOG de acciones.
    Set LOG = New cLOG
    cad = "Albaranes: " & NUmAlbar
    LOG.Insertar 2, vUsu, cad
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    
    
    Screen.MousePointer = vbHourglass
    cad = "scaalb.factursn=1  and scaalb.codtipom='ALV' AND scaalb.numalbar in (" & NUmAlbar & ")"
    TraspasoAlbaranesFacturas "Select *  FROM  scaalb WHERE " & cad, cad, Text1.Text, Combo1.ItemData(Combo1.ListIndex), Nothing, Label2, True, "ALV", "|||", CByte(vParamAplic.NumCopiasFacturacion), True, False
    
    DoEvents
    Label2.Caption = "Leyendo BD"
    Label2.Refresh
    Espera 0.75
    CargaDatos
    Screen.MousePointer = vbDefault
    
    
    
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 0 Then
        Combo1.ForeColor = &H808080
    Else
        Combo1.ForeColor = vbBlack
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        CargaDatos
                
    End If
    
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
   ' Me.cmdFra.Picture = frmPpal.imgListComun.ListImages(5)
    Orden = 1
    Asc = True
    
    CargarCombo_Tabla Me.Combo1, "sbanpr", "codbanpr", "nombanpr", , True
    Combo1.List(0) = "Seleccione banco...."
    Combo1.ListIndex = 0
    
    
    Text1.Tag = "Fecha factura"
    Text1.Text = Text1.Tag
    
End Sub

Private Sub frmC_Selec(vFecha As Date)
    cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)

    
    cad = IIf(Index = 0, "Marcar", "Desmarcar")
    cad = cad & " los albaranes para facturar?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    cad = "UPDATE scaalb set factursn = " & IIf(Index = 0, "1", "0") & " WHERE codtipom = 'ALV'"
    ejecutar cad, False
    CargaDatos
End Sub

Private Sub imgFecha_Click(Index As Integer)
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

Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index = Orden Then
        Asc = Not Asc
    Else
        Asc = True
        Orden = ColumnHeader.Index
    End If
    CargaDatos
End Sub

Private Sub lw1_ItemCheck(ByVal item As MSComctlLib.ListItem)
    cad = "UPDATE scaalb set factursn = " & IIf(item.Checked, "1", "0") & " WHERE codtipom = 'ALV' AND numalbar = " & item.Text
    
    If item.Checked Then
        Marcado = Marcado + 1
        ImporteTot = ImporteTot + ImporteFormateado(item.SubItems(6))
    Else
        Marcado = Marcado - 1
        ImporteTot = ImporteTot - ImporteFormateado(item.SubItems(6))
    End If
    ejecutar cad, False
    Lbl
End Sub


Private Sub Lbl()
    If Marcado > 0 Then
        Label2.Caption = "Selec: " & Marcado & "     €:" & Format(ImporteTot, FormatoImporte)
    Else
        Label2.Caption = ""
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
