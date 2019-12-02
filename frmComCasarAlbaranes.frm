VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComCasarAlbaranes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   375
      Left            =   13080
      TabIndex        =   10
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Desvincular"
      Height          =   375
      Left            =   12960
      TabIndex        =   9
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vincular"
      Height          =   375
      Left            =   12960
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   240
      Width           =   5535
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3975
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7011
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Albaran"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Factura"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fec. fact."
         Object.Width           =   2011
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Albaran"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "codtipoa"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Factura"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fecha fact"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "codtipom"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Referenc"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Codigo"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Nombre cliente"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Base"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   3975
      Left            =   6240
      TabIndex        =   6
      Top             =   960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7011
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Albaran"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Factura"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fec. fact."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "codtipoa"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "numalbar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "F.Alb."
         Object.Width           =   1940
      EndProperty
      Picture         =   "frmComCasarAlbaranes.frx":0000
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Leyendo BD....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   555
      Index           =   4
      Left            =   8760
      TabIndex        =   11
      Top             =   240
      Width           =   5475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      X1              =   6000
      X2              =   6000
      Y1              =   720
      Y2              =   4920
   End
   Begin VB.Label Label2 
      Caption         =   "Asignados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   6240
      TabIndex        =   7
      Top             =   720
      Width           =   2835
   End
   Begin VB.Label Label2 
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Albaranes / facturas clientes pendientes de asignar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   6315
   End
   Begin VB.Label Label2 
      Caption         =   "Albaranes proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2835
   End
End
Attribute VB_Name = "frmComCasarAlbaranes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Codprove As Long

Dim miSQL As String
Dim PVez As Boolean
Dim IT As ListItem

Private Sub VincularAlbaran()

   ' If ListView1.SelectedItem Is Nothing Then Exit Sub
   ' If ListView2.SelectedItem Is Nothing Then Exit Sub
   ' miSQL = "¿Seguro que desea vincular el albaran/factura ?"
   ' If MsgBox(miSQL, vbQuestion + vbYesNo) = vbYes Then
        
        
    'codprove,numfacpr,fecfacpr,numalbPr,fecalbpr,numalbar,codtipom,fecfaccl)
    CadenaDesdeOtroForm = ""
    For NumRegElim = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Selected Then
        
            
            With ListView1.ListItems(NumRegElim)
                'pROVEE
                miSQL = ", (" & Codprove & "," & DBSet(Trim(ListView2.SelectedItem.SubItems(2)), "T") & ","
                miSQL = miSQL & DBSet(Trim(ListView2.SelectedItem.SubItems(3)), "F") & ","
                miSQL = miSQL & DBSet(Trim(ListView2.SelectedItem.Text), "T") & "," & DBSet(Trim(ListView2.SelectedItem.SubItems(1)), "F") & ","
                'vENTAS
                miSQL = miSQL & .Text & " ," & DBSet(Trim(.SubItems(2)), "T") & "," & DBSet(Trim(.SubItems(1)), "F") & ")"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & miSQL
            End With
        End If
    Next
    
    
    If CadenaDesdeOtroForm <> "" Then
        'scafpavinc(codprove,numfacpr,fecfacpr,numalbPr,fecalbpr,codtipom,numalbar,fecfaccl)
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2) 'quitamos la primera coma
        miSQL = "INSERT INTO scafpavinc(codprove,numfacpr,fecfacpr,numalbPr,fecalbpr,numalbar,codtipom,fecfaccl) VALUES  " & CadenaDesdeOtroForm
        ejecutar miSQL, False
    End If
    PVez = True
    Form_activate
End Sub

Private Sub Command1_Click()

    If Me.ListView2.SelectedItem Is Nothing Then Exit Sub
    
    miSQL = ""
    For NumRegElim = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Selected Then
            miSQL = miSQL & vbCrLf & "- " & ListView1.ListItems(NumRegElim).Text & " " & ListView1.ListItems(NumRegElim).SubItems(1)
        End If
    Next
    
    If miSQL = "" Then
        MsgBox "Ningun albaran seleccionado para vincular"
        Exit Sub
    End If
    
     
    
    miSQL = "Albaran proveedor: " & ListView2.SelectedItem.Text & "  " & ListView2.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & "Albaranes clientes:" & miSQL
    miSQL = miSQL & vbCrLf & "¿Continuar?"
    If MsgBox(miSQL, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    
    
        
    VincularAlbaran

End Sub

Private Sub Command2_Click()
    Desvincular
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_activate()
    Screen.MousePointer = vbHourglass
    If PVez Then
        PVez = False
        Set miRsAux = New ADODB.Recordset
        Label2(4).visible = True
        Label2(4).Refresh
        CargaProveedor
        DoEvents
        
        CargaClientes
        Set miRsAux = Nothing
        Label2(4).visible = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.Icon = frmPpal.Icon

    Text2.Text = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", CStr(Codprove))
        
    PVez = True
    
    Caption = "Vincular albaranes-facturas  proveedor  / cliente"
End Sub



Private Sub CargaProveedor()
    ListView2.ListItems.Clear
    'Si queremos meter los albaranes....
    miSQL = "select numalbar,fechaalb,null numfactu,null fecfactu ,albcli,codtipom,fecalbcli from scaalp where codprove =" & Codprove
    miSQL = miSQL & " and fechaalb >= '2019-09-01'"
    miSQL = miSQL & " and codtipom is null and albcli is null"
    miSQL = miSQL & " Union"
    
    
    miSQL = "" 'De momento NO entran los albaranes
    
    
    miSQL = miSQL & " select numalbar,fechaalb, numfactu, fecfactu  from scafpa where codprove =" & Codprove
    miSQL = miSQL & " and fecfactu >= '2019-09-01'"
    miSQL = miSQL & " and not (numfactu, fecfactu) IN (select numfacpr,fecfacpr from scafpaVinc where codprove = " & Codprove
    miSQL = miSQL & ") ORDER BY 1,2"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        Set IT = ListView2.ListItems.Add()
        IT.Text = miRsAux!Numalbar
        IT.SubItems(1) = miRsAux!FechaAlb
        IT.SubItems(2) = DBLet(miRsAux!Numfactu, "T") & " "
        IT.SubItems(3) = DBLet(miRsAux!FecFactu, "T") & " "
           
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Los que estaan
    ListView3.ListItems.Clear
'    miSQL = "select numalbar,fechaalb,null numfactu,null fecfactu ,albcli,codtipom,fecalbcli from scaalp where codprove =" & Codprove
'    miSQL = miSQL & " and  codtipom <>'' and albcli >=0"
'    miSQL = miSQL & " Union"
'    miSQL = miSQL & " select numalbar,fechaalb, numfactu, fecfactu ,albcli,codtipom,fecalbcli from scafpa where codprove =" & Codprove
'    miSQL = miSQL & " and codtipom <>'' and albcli >=0"
'    miSQL = miSQL & " ORDER BY 2,3"
'
    
    'MARZO 2015
        
    miSQL = " select  numalbPr,fecalbpr,numfacpr,fecfacpr,codtipom,numalbar,fecfaccl from scafpaVinc where codprove = " & Codprove
    miSQL = miSQL & " ORDER BY 1,2"
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        Set IT = ListView3.ListItems.Add()
        IT.Text = miRsAux!numalbPr
        IT.SubItems(1) = miRsAux!fecalbpr
        IT.SubItems(2) = DBLet(miRsAux!numfacpr, "T") & " "
        IT.SubItems(3) = DBLet(miRsAux!fecfacpr, "T") & " "
        IT.SubItems(4) = DBLet(miRsAux!codtipom, "T") & " "
        IT.SubItems(5) = DBLet(miRsAux!Numalbar, "T") & " "
        IT.SubItems(6) = DBLet(miRsAux!fecfaccl, "T") & " "
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Para cada Item busco el nombre factura
    
    
End Sub


Private Sub CargaClientes()
Dim RN As ADODB.Recordset
Dim F As Date

    F = DateAdd("m", -3, CDate("01/09/2019"))

    'Cargo un RS con todos los partes vinculados al proveedor
    miSQL = "select numparte from advpartes where codflota in (select codflota from sflotas where codprove=" & Codprove & ")"
    miRsAux.Open miSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText

    Label2(4).Caption = "Albaranes"
    Label2(4).Refresh
    ListView1.ListItems.Clear
    Screen.MousePointer = vbHourglass
    
    miSQL = "select numalbar,fechaalb,codtipom codtipoa,null numfactu,null fecfactu,null codtipom,referenc,codclien,nomclien,0.00 BrutoFac  from scaalb where  referenc like 'part%' "
    miSQL = miSQL & " and codtipom in ('ALS','ALI') AND not (numalbar,codtipom)"
    'miSQL = miSQL & " in (select codtipom,numalbar from scafpaVinc WHERE codprove=" & Codprove & ")  "
    miSQL = miSQL & " in (select codtipom,numalbar from scafpaVinc WHERE fecfaccl>=" & DBSet(F, "F") & ")"
    CargaLWClientes
    
    Screen.MousePointer = vbHourglass
    Label2(4).Caption = "Facturas"
    Label2(4).Refresh
    miSQL = " select numalbar,fechaalb,codtipoa, scafac1.numfactu, scafac1.fecfactu,scafac1.codtipom,"
    miSQL = miSQL & "  referenc,codclien,nomclien,brutofac from scafac,scafac1 where "
    miSQL = miSQL & "  scafac.numfactu = scafac1.numfactu and scafac.codtipom = scafac1.codtipom and scafac.fecfactu = scafac1.fecfactu "
    miSQL = miSQL & "  and codtipoa in ('ALS','ALI') AND referenc like 'part%' and scafac.fecfactu>="
    miSQL = miSQL & DBSet(F, "F")
    miSQL = miSQL & "  AND not (numalbar,codtipoa,scafac1.fecfactu)"
    'miSQL = miSQL & " in (select codtipom,numalbar,fecfaccl from scafpaVinc WHERE codprove=" & Codprove & ")  "
    miSQL = miSQL & " in (select codtipom,numalbar,fecfaccl from scafpaVinc WHERE fecfaccl>=" & DBSet(F, "F") & ")"

'
    CargaLWClientes
    
    
    
    
End Sub


Private Sub CargaLWClientes()
Dim RS As ADODB.Recordset
    
    
    Set RS = New ADODB.Recordset
    'MISQL llevará el sql vinculado a uno u otro
    RS.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
            'Parte: 14119
            Label2(4).Caption = "Refer: " & RS!referenc
            Label2(4).Refresh
            miSQL = Trim(Mid(RS!referenc, 7))
            
            If miSQL <> "" Then
                If Not IsNumeric(miSQL) Then
                    miSQL = ""
                Else
                    miSQL = "numparte = " & miSQL
                    miRsAux.Find miSQL, , adSearchForward, 1
                    If miRsAux.EOF Then miSQL = ""
                End If
                
                miSQL = "OK"
            End If
            If miSQL <> "" Then
                'Partte vinculado al proveedor
        
        
                Set IT = ListView1.ListItems.Add()
                IT.Text = RS!Numalbar
                IT.SubItems(1) = RS!FechaAlb
                IT.SubItems(2) = DBLet(RS!codtipoa, "T") & " "
                IT.SubItems(3) = DBLet(RS!Numfactu, "T") & " "
                IT.SubItems(4) = DBLet(RS!FecFactu, "T") & " "
                IT.SubItems(5) = DBLet(RS!codtipom, "T") & " "
                IT.SubItems(6) = RS!referenc
                IT.SubItems(7) = Format(RS!codClien, "0000")
                IT.SubItems(8) = RS!NomClien
                If RS!BrutoFac <> 0 Then IT.SubItems(9) = RS!BrutoFac
            End If
            
            RS.MoveNext
    Wend
    RS.Close
    
    Set RS = Nothing
End Sub


Private Sub Desvincular()
    If ListView3.SelectedItem Is Nothing Then Exit Sub
    miSQL = "¿Seguro que desea quitar la vinculación del albarán/factura ?"
    If MsgBox(miSQL, vbQuestion + vbYesNo) = vbYes Then
        
            'scafpavinc(codprove,numfacpr,fecfacpr,numalbPr,fecalbpr,codtipom,,fecfaccl)
        
            miSQL = "DELETE FROM scafpavinc"
            miSQL = miSQL & " WHERE codprove =" & Codprove & " AND numfacpr =" & DBSet(Trim(ListView3.SelectedItem.SubItems(2)), "T")
            miSQL = miSQL & " AND fecfacpr = " & DBSet(Trim(ListView3.SelectedItem.SubItems(3)), "F") & " AND numalbPr =" & DBSet(Trim(ListView3.SelectedItem.Text), "T")
            miSQL = miSQL & " AND fecalbpr = " & DBSet(Trim(ListView3.SelectedItem.SubItems(1)), "F") & " AND codtipom =" & DBSet(Trim(ListView3.SelectedItem.SubItems(4)), "T")
             miSQL = miSQL & " AND numalbar = " & DBSet(Trim(ListView3.SelectedItem.SubItems(5)), "N") & " AND fecfaccl =" & DBSet(Trim(ListView3.SelectedItem.SubItems(6)), "F")
            
            ejecutar miSQL, False
            PVez = True
            Form_activate
        
    End If
End Sub

