VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmfacAlbaraYFac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revision datos albaranes /facturas / pedidos"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   20235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
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
      Index           =   3
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   240
      Width           =   5415
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
      Index           =   2
      Left            =   12720
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
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
      Index           =   1
      Left            =   8880
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
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
      Index           =   0
      Left            =   6120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "frmfacAlbaraYFact.frx":0000
      Left            =   1080
      List            =   "frmfacAlbaraYFact.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   19815
      _ExtentX        =   34951
      _ExtentY        =   12303
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   9
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
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Factura"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "fecfac"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lwL 
      Height          =   2535
      Left            =   2160
      TabIndex        =   9
      Top             =   8040
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Text            =   "Codigo"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Articulo"
         Object.Width           =   10354
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Precio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Dto1"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Dto2"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Importe"
         Object.Width           =   3422
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Left            =   11760
      TabIndex        =   12
      Top             =   240
      Width           =   675
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   2
      Left            =   12480
      Picture         =   "frmfacAlbaraYFact.frx":001F
      ToolTipText     =   "Buscar fecha"
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LIneas"
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
      Index           =   4
      Left            =   1080
      TabIndex        =   8
      Top             =   8040
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   16200
      TabIndex        =   7
      Top             =   8160
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
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
      Index           =   3
      Left            =   7920
      TabIndex        =   6
      Top             =   240
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
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
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   600
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   1
      Left            =   8520
      Picture         =   "frmfacAlbaraYFact.frx":0A21
      ToolTipText     =   "Buscar fecha"
      Top             =   240
      Width           =   240
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
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   630
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   5760
      Picture         =   "frmfacAlbaraYFact.frx":0AAC
      ToolTipText     =   "Buscar fecha"
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frmfacAlbaraYFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2
Attribute frmCli.VB_VarHelpID = -1
Private frmAlbG As frmFacEntAlbaranesGR


Dim cad As String
Dim PrimeraVez As Boolean


Dim Orden As Integer
Dim Asc As Boolean

Dim Marcado As Long


Dim TipoAlbaSeleccionado2 As String


Private Sub CargaDatos3(Identeficador As Long)
     If Combo3.ListIndex = 0 Then
        CargaDatosPedido Identeficador
    Else
        CargaDatos2 Identeficador
    End If
End Sub


Private Sub CargaDatos2(Numalbar As Long)
Dim IT As ListItem
Dim Indice As Integer
    On Error GoTo ECargaDatos
    
    Screen.MousePointer = vbHourglass
      Label2.Caption = "Leyendo BD"
        Label2.Refresh
    lw1.ListItems.Clear
    Marcado = 0
    cad = "select c.numalbar numalbar,c.fechaalb fechaalb,codclien,nomclien,substring(stippa.destippa,1,5),nomforpa ,coddirec,nomdirec,referenc, sum(importel) base, null codtipom,null numfactu ,null fecfactu"
    cad = cad & " ,c.codtipom codtipoa from scaalb c inner join sforpa on c.codforpa=sforpa.codforpa"
    cad = cad & " left join stippa   on stippa.tipforpa =sforpa.tipforpa"
    cad = cad & " left join slialb l on c.codtipom=l.codtipom and c.numalbar=l.numalbar where c.codtipom IN " & TipoAlbaSeleccionado2
    
     cad = cad & " AND c.fechaalb between " & DBSet(Me.Text1(0).Text, "F") & " AND  " & DBSet(Me.Text1(1).Text, "F")
    If Text1(2).Text <> "" Then cad = cad & " AND c.codclien =" & Text1(2).Text
    
    
    
    cad = cad & "  group by 1"

    
    
    
    cad = cad & " UNION "
    
    cad = cad & "select c.numalbar numalbar,c.fechaalb fechaalb,codclien,nomclien,substring(stippa.destippa,1,5),nomforpa "
     cad = cad & " ,coddirec,nomdirec,referenc,SUM(IMPORTEL) base, c.codtipom codtipom, c.numfactu numfactu ,c.fecfactu"
    cad = cad & " ,c.codtipoa from scafac inner join "
    cad = cad & " scafac1 c on scafac.codtipom=c.codtipom and scafac.fecfactu=c.fecfactu and scafac.numfactu=C.numfactu "
    cad = cad & " inner join  SLIFAC L on scafac.codtipom=L.codtipom and scafac.fecfactu=L.fecfactu and scafac.numfactu=L.numfactu"
    cad = cad & " and c.CODTIPOA=L.codtipoa and c.numalbar=L.numalbar"
    cad = cad & " inner join sforpa on scafac.codforpa=sforpa.codforpa"
    cad = cad & " left join stippa   on stippa.tipforpa =sforpa.tipforpa"
    cad = cad & "  where c.codtipoa IN " & TipoAlbaSeleccionado2
     cad = cad & " AND c.fechaalb between " & DBSet(Me.Text1(0).Text, "F") & " AND  " & DBSet(Me.Text1(1).Text, "F")
    If Text1(2).Text <> "" Then cad = cad & " AND scafac.codclien =" & Text1(2).Text
    
    cad = cad & " GROUP BY numalbar,fechaalb,NUMFACTU,CODTIPOM"
    
    
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
    Case 8
        cad = cad & " numfactu " & IIf(Not Asc, "DESC", "") & ", fechaalb "
    Case Else
        cad = cad & " numalbar " & IIf(Not Asc, "DESC", "") & ", fechaalb "
    End Select
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Indice = 0
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
        
        
        
        If Not IsNull(miRsAux!codtipom) Then
            cad = miRsAux!codtipom & Format(miRsAux!Numfactu, "000000")
            IT.SubItems(8) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        Else
            cad = " "
            IT.SubItems(8) = miRsAux!Codtipoa
        End If
        IT.SubItems(7) = cad
        
        
        If miRsAux!Numalbar = Numalbar Then
            IT.Selected = True
            Set lw1.SelectedItem = IT
            IT.EnsureVisible
            PonerFocoOBj lw1
            Indice = IT.Index
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Indice = 0 Then
         If lw1.ListItems.Count > 0 Then Indice = 1
    End If
    If Indice > 0 Then CargaLineas Indice
ECargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
     Label2.Caption = ""
        Label2.Refresh
    Screen.MousePointer = vbDefault
End Sub



Private Sub combo3_Click()
    If PrimeraVez Then Exit Sub
    
    Me.lw1.ColumnHeaders.Item(1).Text = "Albarán"
    Me.lw1.ColumnHeaders.Item(8).Text = "Factura"
    
    If Combo3.ListIndex = 2 Then
        TipoAlbaSeleccionado2 = "'ALV','ALZ'"
      
    ElseIf Combo3.ListIndex = 3 Then
        TipoAlbaSeleccionado2 = "'ALZ'"
        
    ElseIf Combo3.ListIndex = 1 Then
        TipoAlbaSeleccionado2 = "'ALV'"
    ElseIf Combo3.ListIndex = 0 Then
        TipoAlbaSeleccionado2 = "PED" 'pedidos
        Me.lw1.ColumnHeaders.Item(1).Text = "Pedido"
        Me.lw1.ColumnHeaders.Item(8).Text = "Cerrado"
    Else
        TipoAlbaSeleccionado2 = "'ALV'"
    End If
    
    TipoAlbaSeleccionado2 = "(" & TipoAlbaSeleccionado2 & ")"
    Me.lwL.ListItems.Clear
   CargaDatos3 -1
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        CargaDatos3 -1
                
    End If
    
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
   ' Me.cmdFra.Picture = frmPpal.imgListComun.ListImages(5)
    Orden = 1
    Asc = True
   TipoAlbaSeleccionado2 = "('PED')"
    Text1(2).Text = ""
    Text1(3).Text = ""
    Me.Combo3.ListIndex = 0
    
    Text1(0).Tag = DateAdd("d", -4, Now)
    Text1(0).Text = Format(Text1(0).Tag, "dd/mm/yyyy")
    
    Text1(1).Tag = Now
    Text1(1).Text = Format(Text1(1).Tag, "dd/mm/yyyy")
    
    
End Sub

Private Sub frmC_Selec(vFecha As Date)
    cad = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    cad = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub imgFecha_Click(Index As Integer)

    If Index = 2 Then

            cad = ""

            
            Set frmCli = New frmBasico2
            AyudaClientes frmCli, Text1(2).Text
            Set frmCli = Nothing

            
            
            
            If cad <> "" Then
                Text1(2).Text = cad
                Text1_LostFocus 2
            End If



    Else
        Set frmC = New frmCal
        frmC.Fecha = Now
        cad = ""
        frmC.Show vbModal
        Set frmC = Nothing
        If cad <> "" Then
            Text1(Index).Text = cad
            Text1_LostFocus Index
        End If
    End If
End Sub

Private Sub Label1_Click(Index As Integer)
    If Index = 2 Then
        If vUsu.Nivel <= 1 Then
            If Combo3.ListCount = 2 Then
                
                Combo3.AddItem "Alb y fact. *"
                
                Combo3.AddItem "Presupuesto"
                
                
                
                HaMostradoCanal2_El_B = True
            End If
        End If
    End If
End Sub

Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index = Orden Then
        Asc = Not Asc
    Else
        Asc = True
        Orden = ColumnHeader.Index
    End If
    CargaDatos3 -1
End Sub

Private Sub lw1_DblClick()
Dim Carga As Boolean
     Dim N As Long
     
     If lw1.ListItems.Count = 0 Then Exit Sub
     If lw1.SelectedItem Is Nothing Then Exit Sub
     N = Val(lw1.SelectedItem)
     Carga = False
     If Combo3.ListIndex = 0 Then
         frmFacEntPedidos.DatosADevolverBusqueda2 = lw1.SelectedItem.Text
        frmFacEntPedidos.EsHistorico = False
        frmFacEntPedidos.Show vbModal
        Carga = True
     Else
        
         If Trim(lw1.SelectedItem.SubItems(7)) = "" Then
                Set frmAlbG = New frmFacEntAlbaranesGR
                frmAlbG.hcoCodTipoM = lw1.SelectedItem.SubItems(8)
                frmAlbG.hcoCodMovim = lw1.SelectedItem.Text
                frmAlbG.Show vbModal
                Set frmAlbG = Nothing
        
        Else
    
             With frmFacHcoFacturas2
                    .DesdeFichaCliente = True
                    .hcoCodMovim = Mid(lw1.SelectedItem.SubItems(7), 4)
                    .hcoCodTipoM = Mid(lw1.SelectedItem.SubItems(7), 1, 3)
                    .hcoFechaMov = lw1.SelectedItem.SubItems(8)
                    
                        .Show vbModal
                End With
    
    
        End If
    End If
    If Carga Then
        Label2.Refresh
        Espera 0.75
        lwL.ListItems.Clear
        CargaDatos3 N
       
        
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub lw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CargaLineas Item.Index
End Sub

Private Sub CargaLineas(Indice As Integer)
 Dim cad As String
    Dim IT
    
    On Error GoTo eCargaLineas
    
    Me.lwL.ListItems.Clear
    If Indice = 0 Then Exit Sub
    
    If Me.Combo3.ListIndex = 0 Then
        cad = " , numlinea as ordenlin FROM sliped where  numpedcl =" & lw1.ListItems(Indice).Text
        
    Else
        If Trim(lw1.ListItems(Indice).SubItems(7)) = "" Then
            cad = " FROM slialb where codtipom = '" & lw1.ListItems(Indice).SubItems(8) & "' AND numalbar =" & lw1.ListItems(Indice).Text
        Else
            cad = " FROM slifac where codtipom = '" & Mid(lw1.ListItems(Indice).SubItems(7), 1, 3) & "' AND numfactu =" & Mid(lw1.ListItems(Indice).SubItems(7), 4)
            cad = cad & " AND fecfactu =" & DBSet(lw1.ListItems(Indice).SubItems(8), "F")
            
            'Y EL ALBARAN
            cad = cad & " AND numalbar =" & lw1.ListItems(Indice).Text
            cad = cad & " AND codtipoa = "
            If Mid(lw1.ListItems(Indice).SubItems(7), 1, 3) = "FAZ" Then
                cad = cad & "'ALZ'"
            Else
                cad = cad & "'ALV'"
            End If
            
        End If
    End If
    
    cad = "Select codartic,nomartic,cantidad,precioar,dtoline1,dtoline2,importel  " & cad
    
    cad = cad & " ORDER BY ordenlin, numlinea"
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
         Set IT = lwL.ListItems.Add()
        IT.Text = miRsAux!codArtic
        IT.SubItems(1) = miRsAux!NomArtic
        IT.SubItems(2) = Format(miRsAux!cantidad, FormatoCantidad)
        IT.SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
        IT.SubItems(4) = Format(miRsAux!dtoline1, FormatoCantidad)
        IT.SubItems(5) = Format(miRsAux!dtoline2, FormatoCantidad)
        IT.SubItems(6) = Format(miRsAux!ImporteL, FormatoPrecio)
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Exit Sub
eCargaLineas:
    MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)


    If Index = 2 Then

        
        cad = ""
        If Text1(Index).Text = "" Then
            
        Else
            If Not IsNumeric(Text1(Index).Text) Then
                Text1(Index).Text = ""
            Else
                cad = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", CStr(CLng(Text1(Index).Text)))
                If cad = "" Then cad = "No existe el cliente"
            End If
        End If
        Text1(3).Text = cad

    Else
        If Index = 3 Then
            'Nomclien, nbo hgao nada
            
        Else
            If Text1(Index).Text = Text1(Index).Tag Then Exit Sub
            
            PonerFormatoFecha Text1(Index)
            If Text1(Index).Text = "" Then
                Text1(Index).ForeColor = &H808080
                Text1(Index).Text = Text1(Index).Tag
            Else
                Text1(Index).ForeColor = vbBlack
            End If
        End If
    End If
    
    CargaDatos3 -1
End Sub






Private Sub CargaDatosPedido(NumPedido As Long)
Dim IT As ListItem
Dim Indice As Integer
    On Error GoTo ECargaDatos
    
    Screen.MousePointer = vbHourglass
      Label2.Caption = "Leyendo BD"
        Label2.Refresh
    lw1.ListItems.Clear
    Marcado = 0
    cad = "select c.numpedcl,c.fecpedcl,codclien,nomclien,substring(stippa.destippa,1,5),nomforpa ,coddirec,nomdirec,referenc, sum(importel) base, null codtipom,null numfactu ,null fecfactu"
    cad = cad & " ,cerrado from scaped c inner join sforpa on c.codforpa=sforpa.codforpa"
    cad = cad & " left join stippa   on stippa.tipforpa =sforpa.tipforpa"
    cad = cad & " left join sliped l on c.numpedcl=l.numpedcl where 1=1"
    
     cad = cad & " AND c.fecpedcl between " & DBSet(Me.Text1(0).Text, "F") & " AND  " & DBSet(Me.Text1(1).Text, "F")
    
    If Text1(2).Text <> "" Then cad = cad & " AND c.codclien =" & Text1(2).Text
    
    cad = cad & "  group by 1"

    
    
    
    cad = cad & " ORDER BY "
    Select Case Orden

    Case 2
         cad = cad & " fecpedcl " & IIf(Not Asc, "DESC", "") & ", numalbar "
    Case 3
         cad = cad & " codclien " & IIf(Not Asc, "DESC", "") & ", numalbar  "
    
    Case 4
        cad = cad & " nomclien " & IIf(Not Asc, "DESC", "") & ", codclien, numalbar  "
    Case 5
        cad = cad & " nomdirec " & IIf(Not Asc, "DESC", "") & ", referenc " & IIf(Not Asc, "DESC", "") & ", codclien, numalbar  "
    Case 6
        cad = cad & " nomforpa " & IIf(Not Asc, "DESC", "") & ", fecpedcl, numpedcl  "
    
    Case 7
        cad = cad & " base " & IIf(Not Asc, "DESC", "") & ", fecpedcl  "
        
    Case 8
        
        cad = cad & " cerrado " & IIf(Not Asc, "DESC", "") & ", fecpedcl ,numpedcl"
    Case Else
        cad = cad & " numpedcl " & IIf(Not Asc, "DESC", "") & ", fecpedcl "
    End Select
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Indice = 0
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = Format(miRsAux!NumPedcl, "000000")
        IT.SubItems(1) = Format(miRsAux!fecpedcl, "dd/mm/yyyy")
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
        
        
        
        If Not IsNull(miRsAux!codtipom) Then
            cad = miRsAux!codtipom & Format(miRsAux!Numfactu, "000000")
            IT.SubItems(8) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        Else
            cad = IIf(miRsAux!cerrado = 1, "CERR", "")
            IT.SubItems(8) = " "
        End If
        IT.SubItems(7) = cad
        
        
        If miRsAux!NumPedcl = NumPedido Then
            IT.Selected = True
            Set lw1.SelectedItem = IT
            IT.EnsureVisible
            PonerFocoOBj lw1
            Indice = IT.Index
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Indice = 0 Then
         If lw1.ListItems.Count > 0 Then Indice = 1
    End If
    If Indice > 0 Then CargaLineas Indice
ECargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
     Label2.Caption = ""
        Label2.Refresh
    Screen.MousePointer = vbDefault
End Sub


