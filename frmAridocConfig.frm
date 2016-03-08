VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAridocConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Aridoc"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   240
      TabIndex        =   24
      Top             =   1320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5106
      View            =   2
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
         Text            =   "Titulo"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "campotitulo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "campofecha"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "seltitulo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "selfecha"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Carpeta"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "descricpion"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   8640
      TabIndex        =   23
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   7440
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   3960
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   225
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   225
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   3555
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3240
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Carpeta"
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   20
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   7
      Left            =   4080
      TabIndex        =   18
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   6
      Left            =   6960
      TabIndex        =   16
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   14
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   4
      Left            =   6960
      TabIndex        =   12
      Top             =   1455
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   10
      Top             =   1455
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Extension"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "PATH Aridoc"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmAridocConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rs As ADODB.Recordset

Dim Cad As String
Dim i As Integer

Private Sub Command1_Click()


    On Error GoTo EC
    'GUARDAR
    'Comprobar
    For i = 0 To 2
        Text1(i).Text = Trim(Text1(i).Text)
        If Text1(i).Text = "" Then
            MsgBox "Campos configuracion obligados", vbExclamation
            Exit Sub
        End If
    Next i
    If Me.Tag = "" Then
        'NUEVO
        'Cad = "insert into `paramaridoc` (`usuario`,`path`,`extension`,`activada`"
        Cad = "insert into `paramaridoc` (`usuario`,`path`,`extension`"
        Cad = Cad & ") values ("
        Cad = Cad & DBSet(Text1(0).Text, "T") & ","
        Cad = Cad & DBSet(CambiarBarrasPATH2(True, Text1(0).Text), "T") & ","
        Cad = Cad & DBSet(Text1(0).Text, "T") & ")"
        'Cad = Cad & Abs(Check1.Value) & ")"
    Else
        'UPDATE
        Cad = "update `paramaridoc` set "
        Cad = Cad & "usuario= " & DBSet(Text1(0).Text, "T")
        Cad = Cad & ",path='" & CambiarBarrasPATH2(True, Text1(1).Text)
        Cad = Cad & "',extension=" & Text1(2).Text
        'Cad = Cad & ",activada=" & Abs(Check1.Value)

    End If
    Conn.Execute Cad
    Unload Me
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    Me.Tag = ""
    limpiar Me
    If PonerCampos Then Me.Tag = "YA"
    Command1.visible = vUsu.Nivel <= 1
    
    PonerLabels
    
End Sub


Private Function PonerCampos() As Boolean
Dim IT
    On Error GoTo EP
    PonerCampos = False
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from paramaridoc", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Text1(0).Text = ""
    If Not Rs.EOF Then
        Text1(0).Text = Rs!Usuario
        Text1(1).Text = CambiarBarrasPATH2(False, CStr(Rs!Path))
        Text1(2).Text = Rs!extension
        'Check1.Value = Rs!activada
        Text1_LostFocus 2
        PonerCampos = True
    End If
    Rs.Close
    
    If Text1(0).Text <> "" Then
        'HAY DATOS
        Rs.Open "Select * FROM paramaridoc_lin ORDER BY codigo", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Set IT = ListView1.ListItems.Add
            IT.Key = "C" & Rs!Codigo
            IT.Text = Rs!Titulo
            
            IT.Tag = DBLet(Rs!camposel) 'Para marcar si esta seleccionado
            IT.SubItems(1) = DBLet(Rs!campoTitulo)
            IT.SubItems(2) = DBLet(Rs!camposel)
            IT.SubItems(3) = DBLet(Rs!fechatitulo)
            IT.SubItems(4) = DBLet(Rs!fechasel)
            IT.SubItems(5) = DBLet(Rs!Carpeta)
            IT.SubItems(6) = CambiarBarrasPATH2(False, DBLet(Rs!Descripcion))
            Rs.MoveNext
        Wend
        Rs.Close
        If ListView1.ListItems.Count > 0 Then
            ListView1.ListItems(1).Selected = True
            CargaDatosLineas ListView1.ListItems(1).Index
        End If
            
    End If
EP:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Function

Private Sub PonerDescripcionTextos()
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            For NumRegElim = 1 To 4
                Cad = RecuperaValor(Rs!camposel, CInt(NumRegElim))
                i = Val(Cad)
                If i = 0 Then
                    Cad = ""
                Else
                    Cad = RecuperaValor(Rs!campoTitulo, i)
                End If
            Next
        
    Rs.Close
    Set Rs = Nothing
End Sub

Private Sub CargaDatosLineas(Clave As Integer)
    If ListView1.ListItems(Clave).Tag = "" Then
        For i = 3 To 8
            Text1(i).Text = ""
        Next i
        Text2(1).Text = ""
    Else
        Cad = ListView1.ListItems(Clave).SubItems(2)
        For i = 1 To 4
            Cad = RecuperaValor(ListView1.ListItems(Clave).SubItems(2), i)
            If Cad <> "" Then
                NumRegElim = Val(Cad)
                Cad = RecuperaValor(ListView1.ListItems(Clave).SubItems(1), CInt(NumRegElim + 1))
            
            End If
            Text1(i + 2).Text = Cad
        Next i
        'Las fechas
        Cad = RecuperaValor(ListView1.ListItems(Clave).SubItems(4), 1)
        If Cad <> "" Then
            NumRegElim = Val(Cad)
            Cad = RecuperaValor(ListView1.ListItems(Clave).SubItems(3), CInt(NumRegElim + 1))
        
        End If
        Text1(7).Text = Cad
        
        Text1(8).Text = ListView1.ListItems(Clave).SubItems(5)
        Text2(1).Text = ListView1.ListItems(Clave).SubItems(6)
    End If
End Sub



Private Sub ListView1_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    CargaDatosLineas ListView1.SelectedItem.Index
End Sub

Private Sub ListView1_DblClick()
    If vUsu.Nivel > 1 Then Exit Sub
    
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    CadenaDesdeOtroForm = ""
    Cad = ""
    For i = 0 To 4
        
        
        frmAridocConfigCampos.Label1(i).Caption = Me.Label1(i + 3).Caption
        
        
       
    Next i
    frmAridocConfigCampos.Sel1 = ListView1.SelectedItem.SubItems(2)
    frmAridocConfigCampos.Sel2 = ListView1.SelectedItem.SubItems(4)
    frmAridocConfigCampos.ParaC1 = ListView1.SelectedItem.SubItems(1)
    frmAridocConfigCampos.ParaC2 = ListView1.SelectedItem.SubItems(3)
    frmAridocConfigCampos.Carpeta2 = ListView1.SelectedItem.SubItems(5) & "|" & ListView1.SelectedItem.SubItems(6) & "|"
    
    frmAridocConfigCampos.CargarCombos
    frmAridocConfigCampos.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        i = InStr(1, CadenaDesdeOtroForm, "·")
        ListView1.SelectedItem.SubItems(2) = Mid(CadenaDesdeOtroForm, 1, i - 1)
        Cad = "UPDATE paramaridoc_lin SET camposel = '" & ListView1.SelectedItem.SubItems(2)
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, i + 1)
        
        i = InStr(1, CadenaDesdeOtroForm, "·")
        ListView1.SelectedItem.SubItems(4) = Mid(CadenaDesdeOtroForm, 1, i - 1)
        Cad = Cad & "',fechasel = '" & ListView1.SelectedItem.SubItems(4) & "'"
        CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, i + 1)
        ListView1.SelectedItem.SubItems(5) = RecuperaValor(CadenaDesdeOtroForm, 1)
        ListView1.SelectedItem.SubItems(6) = RecuperaValor(CadenaDesdeOtroForm, 2)
        
        If ListView1.SelectedItem.SubItems(5) = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = ListView1.SelectedItem.SubItems(5)
        End If
        Cad = Cad & " , carpeta = " & CadenaDesdeOtroForm
        
        If ListView1.SelectedItem.SubItems(6) = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = "'" & DevNombreSQL(CambiarBarrasPATH2(True, ListView1.SelectedItem.SubItems(6))) & "'"
        End If
        Cad = Cad & " , descripcion = " & CadenaDesdeOtroForm
        
        Cad = Cad & " WHERE codigo = " & Mid(ListView1.SelectedItem.Key, 2)
        Conn.Execute Cad
        Espera 0.1
        CargaDatosLineas ListView1.SelectedItem.Index
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
        Text1(Index).Text = Trim(Text1(Index).Text)
        
        'Extension
        Cad = ""
        Select Case Index
        Case 2
            
            If Text1(2).Text <> "" Then
                If IsNumeric(Text1(2).Text) Then
                    Cad = DevuelveDesdeBDNew(conConta, "extension", "descripcion", "codext", Text1(2).Text, "N")
                    If Cad = "" Then MsgBox "NO se encuentra la extension: " & Text1(2).Text, vbExclamation
                Else
                    MsgBox "Campo debe ser numerico", vbExclamation
                End If
            End If
            Text2(0).Text = Cad
            If Cad = "" Then
                If Text1(2).Text <> "" Then PonerFoco Text1(2)
                Text1(2).Text = ""
            End If
'        Case 8
'            'CARPETA
'            'Cad = ""
'            If Text1(8).Text <> "" Then
'                If IsNumeric(Text1(8).Text) Then
'                    Cad = DevuelveDesdeBDNew(conConta, "carpetas", "nombre", "codcarpeta", Text1(8).Text, "N")
'                    If Cad = "" Then MsgBox "NO se encuentra la carpeta: " & Text1(8).Text, vbExclamation
'                Else
'                    MsgBox "Campo debe ser numerico", vbExclamation
'                End If
'            End If
'            Text2(1).Text = Cad
'            If Cad = "" Then
'                If Text1(8).Text <> "" Then PonerFoco Text1(8)
'                Text1(8).Text = ""
'            End If
'
'
        End Select
End Sub


Private Sub PonerLabels()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from configuracion", ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        For NumRegElim = 1 To 5
            Cad = DBLet(Rs.Fields(NumRegElim), "T")
            If Cad = "" Then Cad = Rs.Fields(NumRegElim).Name
            Label1(2 + NumRegElim).Caption = Cad
            
        Next NumRegElim
    End If
    Rs.Close
    Set Rs = Nothing
End Sub



