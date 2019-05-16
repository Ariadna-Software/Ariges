VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPedidoVincularTodasLineas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   13425
   StartUpPosition =   2  'CenterScreen
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
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "0"
      Top             =   7920
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
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Tag             =   "0"
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Aceptar"
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
      Index           =   1
      Left            =   10320
      TabIndex        =   7
      Top             =   7920
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Index           =   0
      Left            =   11760
      TabIndex        =   4
      Top             =   7920
      Width           =   1155
   End
   Begin MSComctlLib.ListView lwPed 
      Height          =   2775
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Text            =   "Lin"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Referencia"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Articulo"
         Object.Width           =   9172
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pendiente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Solicitadas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Servidas"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lwAlb 
      Height          =   3615
      Left            =   360
      TabIndex        =   3
      Top             =   4200
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   6376
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Albaran"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha alb"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Factura"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Referencia"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccionadas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   3360
      TabIndex        =   11
      Top             =   8002
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Servidas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   8002
      Width           =   1125
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   9000
      Picture         =   "frmPedidoVincularTodasLineas.frx":0000
      ToolTipText     =   "Quitar seleccion"
      Top             =   8040
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   9480
      Picture         =   "frmPedidoVincularTodasLineas.frx":014A
      ToolTipText     =   "Selec. todos"
      Top             =   8040
      Width           =   240
   End
   Begin VB.Label lblArt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   3840
      Width           =   8055
   End
   Begin VB.Label lblPed 
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   12855
   End
   Begin VB.Label Label1 
      Caption         =   "Albaranes     Articulo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmPedidoVincularTodasLineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public NumPedcl As Long



Dim IT As ListItem
Dim C As String
Dim RS As ADODB.Recordset

Private Sub cmdCancelar_Click(index As Integer)
Dim SinLineas As Boolean

    If index = 1 Then
    
        If lwPed.SelectedItem Is Nothing Then Exit Sub
    
    
        If lwAlb.ListItems.Count > 0 Then
            C = ""
            For NumRegElim = 1 To lwAlb.ListItems.Count
                If lwAlb.ListItems(NumRegElim).Checked Then C = C & "X"
            Next
            If C = "" Then
                If MsgBox("Ninguna linea seleccionada. ¿Continuar igualmente?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
              
            Else
                C = "Va a vincular la linea del pedido con " & Len(C) & " lineas de albaran/factura" & vbCrLf & "¿Continuar?"
                If MsgBox(C, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        ActualizaID
        Espera 0.5
        CargaLwPed
        Screen.MousePointer = vbDefault
        If lwPed.ListItems.Count > 0 Then Exit Sub
        
    End If
    Unload Me
End Sub

Private Sub ActualizaID()
Dim I As Integer
        'Ha dicho que si.
        ' Obtener un nuevo idL y asignamos lin pedido y las lineas alb/fra
        C = DevuelveDesdeBD(conAri, "contador", "stipom", "codtipom", "LPD", "T")
        NumRegElim = Val(C) + 1


        C = "UPDATE sliped set idl=" & NumRegElim & " WHERE numpedcl=" & NumPedcl & "  AND numlinea =" & lwPed.SelectedItem.Text
        If ejecutar(C, False) Then
            For I = 1 To lwAlb.ListItems.Count
                If lwAlb.ListItems(I).Checked Then
                    C = lwAlb.ListItems(I).Tag
                    C = Replace(C, "####", NumRegElim)
                    ejecutar C, False
                End If
            Next
        End If
        
        
        C = "UPDATE stipom SET contador= " & NumRegElim & " WHERE codtipom='LPD'"
        conn.Execute C
End Sub


Private Sub Form_activate()
    Screen.MousePointer = vbHourglass
    If Me.Tag = 0 Then
        Me.Tag = 1
        CargaLwPed
        
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
            
      Me.Tag = 0
      Me.Icon = frmPpal.Icon
        lwPed.ListItems.Clear
    Me.lwAlb.ListItems.Clear
                
End Sub


Private Sub CargaLwPed()
    'Cargamos del pedido los articulos que esten por vincular

    Set RS = New ADODB.Recordset
    lblPed.Caption = ""
    lwPed.ListItems.Clear
    C = "Select codclien,referenc,numlinea,codartic,nomartic,solicitadas,cantidad,importel,fecpedcl,NomClien FROM scaped inner join sliped on scaped.numpedcl=sliped.numpedcl "
    C = C & " WHERE idl=0 and scaped.numpedcl =" & NumPedcl & " ORDER BY numlinea"
    RS.Open C, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    lwPed.Tag = "0|0|"
    If Not RS.EOF Then
        C = "Ped " & Format(NumPedcl, "000000") & " - " & Format(RS!fecpedcl, "dd/mm/yy") & "   " & RS!NomClien
        If DBLet(RS!referenc, "T") <> "" Then C = C & "      (" & RS!referenc & ")"
        lblPed.Caption = C
        lwPed.Tag = RS!codClien & "|" & DBLet(RS!referenc, "T") & "|"
        While Not RS.EOF
            Set IT = lwPed.ListItems.Add(, "C" & Format(RS!numlinea, "000"))
            IT.Text = RS!numlinea
            IT.SubItems(1) = RS!codArtic
            IT.SubItems(2) = RS!NomArtic
            IT.SubItems(3) = RS!cantidad
            IT.SubItems(4) = RS!solicitadas
            IT.SubItems(5) = RS!ImporteL
            
            IT.SubItems(6) = RS!solicitadas - RS!cantidad
            
            RS.MoveNext
        Wend
    End If
    RS.Close
    
    lwAlb.ListItems.Clear
    If lwPed.ListItems.Count > 0 Then
        Set lwPed.SelectedItem = lwPed.ListItems(1)
        CargaLineas lwPed.ListItems(1).SubItems(1)
        Text1(0).Tag = lwPed.ListItems(1).SubItems(6)
        Text1(0).Text = Format(Text1(0).Tag, FormatoCantidad)
    End If
    
    
    Set RS = Nothing
    
End Sub

Private Sub imgCheck_Click(index As Integer)
Dim MultSelect As Boolean
    C = ""
    MultSelect = False
    For NumRegElim = 1 To lwAlb.ListItems.Count
        If lwAlb.ListItems(NumRegElim).Selected Then
            C = C & "S"
            If Len(C) > 1 Then
                MultSelect = True
                Exit For
            End If
        End If
    Next
    Text1(1).Tag = 0
    Text1(1).Text = ""
    For NumRegElim = 1 To lwAlb.ListItems.Count
        If MultSelect Then
            If lwAlb.ListItems(NumRegElim).Selected Then lwAlb.ListItems(NumRegElim).Checked = index = 1
        Else
            lwAlb.ListItems(NumRegElim).Checked = index = 1
        End If
        If lwAlb.ListItems(NumRegElim).Checked Then Text1(1).Tag = Text1(1).Tag + CCur(lwAlb.ListItems(NumRegElim).SubItems(6))
    Next
    If Text1(1).Tag <> 0 Then Text1(1).Text = Format(Text1(1).Tag, FormatoCantidad)
End Sub

Private Sub lwAlb_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim Impor As Currency
    Impor = ImporteFormateado(Item.SubItems(6))
    If Not Item.Checked Then Impor = -Impor
    
    Text1(1).Tag = Text1(1).Tag + Impor
    If Text1(1).Tag = 0 Then
        Text1(1).Text = ""
    Else
        Text1(1).Text = Format(Text1(1).Tag, FormatoCantidad)
    End If
    
End Sub

Private Sub lwPed_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    CargaLineas Item.SubItems(1)
    Text1(0).Tag = Item.SubItems(6)
    Text1(0).Text = Format(Text1(0).Tag, FormatoCantidad)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaLineas(Articulo As String)
Dim Aux As String
    On Error GoTo eCargaLineas
    
    
    Set RS = New ADODB.Recordset
    
    lwAlb.ListItems.Clear
    Me.lblArt.Caption = ""
    Text1(1).Tag = 0
    Text1(1).Text = ""
    C = "select scaalb.numalbar,if(scaalb.Codtipom='ALZ','*',' ') canal2,referenc,fechaalb,numlinea,cantidad,importel ,scaalb.codtipom,nomartic,"
    Aux = RecuperaValor(lwPed.Tag, 2)
    If Aux <> "" Then
      C = C & " if(ucase(referenc)= " & DBSet(UCase(Aux), "T") & ",0,1) "
    Else
      C = C & " 1 "
    End If
    C = C & " OrdRef FROM scaalb left join  slialb  on scaalb.numalbar =slialb.numalbar and scaalb.codtipom =slialb.codtipom "
    C = C & " WHERE codclien=" & RecuperaValor(lwPed.Tag, 1) & " AND codartic=" & DBSet(Articulo, "T") & " AND idl=0"
    C = C & " ORDER By OrdRef ,fechaalb desc"
    Set RS = New ADODB.Recordset
    RS.Open C, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Set IT = lwAlb.ListItems.Add()
        'ORden
        Aux = RS!OrdRef & "A" & Format(RS!FechaAlb, "yyyymmdd") & Format(RS!Numalbar, "00000000") & Format(RS!numlinea, "000")
        If Me.lblArt.Caption = "" Then Me.lblArt.Caption = RS!NomArtic
        'scaped.numpedcl,referenc,fecpedcl,numlinea,solicitadas,servidas
        IT.Text = " "
        IT.SubItems(1) = Aux
        IT.ToolTipText = Aux
        IT.SubItems(2) = RS!Numalbar
        
        IT.SubItems(3) = Format(RS!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(4) = " "
        
        IT.SubItems(5) = " "
        If DBLet(RS!referenc, "T") <> "" Then IT.SubItems(5) = RS!referenc
        If DBLet(RS!OrdRef, "N") = 0 Then
            IT.ListSubItems(5).ForeColor = vbRed
            IT.ListSubItems(5).ToolTipText = "Misma referencia/obra"
        End If
        
        
         IT.SubItems(6) = Format(RS!cantidad, FormatoCantidad)
         IT.SubItems(7) = Format(RS!ImporteL, FormatoCantidad)
        IT.Selected = False
        Aux = "UPDATE slialb set idl=#### WHERE slialb.codtipom=" & DBSet(RS!codtipom, "T")
        Aux = Aux & " AND slialb.numalbar=" & RS!Numalbar & " AND slialb.numlinea=" & RS!numlinea
        IT.Tag = Aux
        
        
        
        RS.MoveNext
    Wend
    RS.Close
    

    
    C = "select scafac1.numfactu,scafac1.codtipom,scafac1.numalbar,if(scafac1.Codtipom='ALZ','*',' ') canal2,scafac1.codtipoa,scafac1.fecfactu"
    C = C & ",referenc,fechaalb,numlinea,cantidad,importel,nomartic ,"
    Aux = RecuperaValor(lwPed.Tag, 2)
    If Aux <> "" Then
      C = C & " if(ucase(referenc)= " & DBSet(UCase(Aux), "T") & ",0,1) "
    Else
      C = C & " 1 "
    End If
    C = C & " OrdRef FROM  scafac left join scafac1 on scafac.codtipom=scafac1.codtipom "
    C = C & " AND scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu"
        
    C = C & " left join slifac on slifac.codtipom=scafac1.codtipom "
    C = C & " AND slifac.numfactu=scafac1.numfactu AND slifac.fecfactu=scafac1.fecfactu"
    C = C & " AND slifac.numalbar=scafac1.numalbar AND slifac.codtipoa=scafac1.codtipoa"
        

    C = C & " WHERE scafac.fecfactu >20180101 and codclien=" & RecuperaValor(lwPed.Tag, 1) & " AND codartic=" & DBSet(Articulo, "T") & " AND idl=0 "
    C = C & " ORDER By OrdRef ,fechaalb desc"
    
    RS.Open C, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Set IT = Me.lwAlb.ListItems.Add()
        
        If Me.lblArt.Caption = "" Then Me.lblArt.Caption = RS!NomArtic
        
        'orden
        Aux = RS!OrdRef & "F" & Format(RS!FechaAlb, "yyyymmdd") & Format(RS!Numalbar, "00000000") & Format(RS!numlinea, "000")
        IT.Text = IIf(RS!codtipom = "ALZ", "*", " ")
        IT.SubItems(1) = Aux
        IT.ToolTipText = Aux
        IT.SubItems(2) = RS!Numalbar
        
        IT.SubItems(3) = Format(RS!FechaAlb, "dd/mm/yyyy")
        IT.SubItems(4) = RS!FechaAlb
        
        IT.SubItems(5) = " "
        If DBLet(RS!referenc, "T") <> "" Then IT.SubItems(5) = RS!referenc
        If DBLet(RS!OrdRef, "N") = 0 Then
            IT.ListSubItems(5).ForeColor = vbRed
            IT.ListSubItems(5).ToolTipText = "Misma referencia/obra"
        End If
        
        
         IT.SubItems(6) = Format(RS!cantidad, FormatoCantidad)
         IT.SubItems(7) = Format(RS!ImporteL, FormatoCantidad)
        
        
        Aux = "UPDATE slifac set idl=#### WHERE slifac.codtipom=" & DBSet(RS!codtipom, "T")
        Aux = Aux & " AND slifac.numfactu=" & RS!Numfactu & " AND slifac.fecfactu=" & DBSet(RS!FecFactu, "F")
        Aux = Aux & " AND slifac.numalbar=" & RS!Numalbar & " AND slifac.codtipoa=" & DBSet(RS!codtipoa, "T")
        Aux = Aux & " AND slifac.numlinea=" & RS!numlinea
        IT.Tag = Aux
        IT.Selected = False
        RS.MoveNext
    Wend
    RS.Close
    
    
    If lwAlb.ListItems.Count > 0 Then
        lwAlb.ListItems(1).EnsureVisible
        Set lwAlb.SelectedItem = lwAlb.ListItems(1)
    End If
    Set RS = Nothing
    Exit Sub
eCargaLineas:
    MuestraError Err.Number, , Err.Description
    Set RS = New ADODB.Recordset
End Sub

