VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWH_SelExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WHOSE"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   1
      Left            =   5760
      Picture         =   "frmWH_SelExp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "buscar siguiente"
      Top             =   8880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   0
      Left            =   5280
      Picture         =   "frmWH_SelExp.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "bucar anterior"
      Top             =   8880
      Width           =   375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWH_SelExp.frx":0B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWH_SelExp.frx":7376
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdNuevo 
      Height          =   495
      Left            =   120
      Picture         =   "frmWH_SelExp.frx":7D88
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Nuevo expediente"
      Top             =   8760
      Width           =   495
   End
   Begin VB.CommandButton cmdVerClientePotencial 
      Height          =   495
      Left            =   720
      Picture         =   "frmWH_SelExp.frx":878A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Modificar datos cliente"
      Top             =   8760
      Width           =   495
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cerrar"
      Height          =   495
      Index           =   0
      Left            =   8280
      TabIndex        =   6
      Top             =   8760
      Width           =   1215
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   7815
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   13785
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
         Text            =   "Cliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Expediente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Año"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Buscar "
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   8940
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Clientes / Expedientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmWH_SelExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private WithEvents frmCliPot As frmFacClienPot
Attribute frmCliPot.VB_VarHelpID = -1
Dim Cad As String
Dim IT As ListItem
Dim OrdenIncial As Byte  'para no tocar el fichero a todas horas
Dim PrimVez As Boolean





Private Sub CargaClientesPotenciales2()
  
    Cad = "select whoExpedientePot.codclien,nomclien,fechaalt,telclie1 from whoExpedientePot,sclipot where sclipot.codclien=whoExpedientePot.codclien"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        'Metemos  el nodo
        NumRegElim = miRsAux!codclien
        Set IT = Me.lw1.ListItems.Add(, "P" & Format(NumRegElim, "000000"))
        IT.SmallIcon = 1
        IT.Text = Format(NumRegElim, "000000")
        IT.SubItems(1) = miRsAux!NomClien
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Sub CargaExpedientesClientes2()

    Cad = "select whoexpedientecli.codclien,nomclien,expediente,anoexp,nombre from whoexpedientecli"
    Cad = Cad & " inner join sclien on whoexpedientecli.codclien=sclien.codclien"
    Cad = Cad & " left join whoobrascli on whoexpedientecli.codclien=whoobrascli.codclien     "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        'Metemos  el nodo
        NumRegElim = miRsAux!codclien
        
        Set IT = Me.lw1.ListItems.Add()
        IT.SmallIcon = 2
        IT.Text = Format(NumRegElim, "000000")
        IT.SubItems(1) = miRsAux!NomClien
        
        If Not IsNull(miRsAux!expediente) Then
            IT.SubItems(2) = Format(DBLet(miRsAux!expediente, "N"), "000000")
            IT.SubItems(3) = DBLet(miRsAux!anoexp, "N")
        Else
            IT.SubItems(2) = " "
            IT.SubItems(3) = " "
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub




Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdEliminarExp_Click()
    MsgBox "No tiene permisos", vbExclamation
    
End Sub

Private Sub cmdNuevo_Click()
    NuevoExpediente
End Sub


Private Sub NuevoExpediente()
    'Seleccionamos Cliente potencial
    Cad = ""
    CadenaDesdeOtroForm = ""
    Set frmCliPot = New frmFacClienPot
    frmCliPot.DatosADevolverBusqueda = "0|"
    frmCliPot.Show vbModal
    Set frmCliPot = Nothing
    
    If Cad <> "" Then
        'OK. Ha seleccioado un cliente
        NumRegElim = RecuperaValor(Cad, 1)
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "codclien", "whoexpedientepot", "codclien", CStr(NumRegElim))
        If CadenaDesdeOtroForm <> "" Then
            MsgBox "Ya existe expediente", vbExclamation
        Else
            'CREAMOS EXPEDIENTE
            If GenerarEstructuraPotencial(NumRegElim) Then
                'Insertamos enla tabla, e insertamos el NODO
                CadenaDesdeOtroForm = "INSERT INTO whoexpedientepot(codclien) VALUES (" & NumRegElim & ")"
                If Ejecutar(CadenaDesdeOtroForm, False) Then
                
                    'Metemos  el nodo
                    Set IT = Me.lw1.ListItems.Add(, "P" & Format(NumRegElim, "000000"))
                    IT.SmallIcon = 1
                    IT.Text = Format(NumRegElim, "000000")
                    IT.SubItems(1) = RecuperaValor(Cad, 2)
                    IT.EnsureVisible
                    IT.Selected = True
                    Set lw1.SelectedItem = IT
                    
                    
                    'Vemos el expediente
                    lw1_DblClick
                    
                Else
                    'DEBERIAMOS BORRAR ESTRUCUTRA
                    MsgBox "ERROR CRITICO. Borrar estructura. Avise soporte técnico", vbCritical
                End If
                
            End If
        End If
        
        
    End If
    
    
End Sub

Private Sub cmdVerClientePotencial_Click()
    If Me.lw1.ListItems.Count = 0 Then Exit Sub
    If Me.lw1.SelectedItem Is Nothing Then Exit Sub
    
    If lw1.SelectedItem.SmallIcon = 1 Then
    
        'Abrir cliente potencial
        Set frmCliPot = New frmFacClienPot
        frmCliPot.DatosADevolverBusqueda = lw1.SelectedItem.Text
        frmCliPot.Show vbModal
        Set frmCliPot = Nothing
    
        Cad = DevuelveDesdeBD(conAri, "nomclien", "sclipot", "codclien", Mid(lw1.SelectedItem.Key, 2))
        
    Else
        frmFacClientes.VerCliente = Val(lw1.SelectedItem.Text)
        frmFacClientes.DatosADevolverBusqueda = ""
        frmFacClientes.Show vbModal
            
            
        Cad = ""
    End If
    If Cad <> "" Then lw1.SelectedItem.SubItems(1) = Cad
End Sub

Private Sub Command1_Click(Index As Integer)
Dim J As Integer
Dim I As Integer
Dim fin As Boolean
Dim Inicio As Integer

    If Trim(Me.Text1.Text) = "" Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    
    If lw1.SelectedItem Is Nothing Then Set lw1.SelectedItem = lw1.ListItems(1)
    
    'Desde el select item que se encuentre buscara encontrar la cadena
    'Desde el liw
        
    

    If Index = 1 Then
        'Siguiente
        For I = lw1.SelectedItem.Index + 1 To lw1.ListItems.Count
            If lw1.SortKey = 0 Then
                'TEXT
                J = InStr(1, lw1.ListItems(I).Text, Text1.Text, vbTextCompare)
            Else
                J = InStr(1, lw1.ListItems(I).SubItems(lw1.SortKey), Text1.Text, vbTextCompare)
            End If
            If J > 0 Then Exit For
        Next
        
    Else
        For I = lw1.SelectedItem.Index - 1 To 1 Step -1
            If lw1.SortKey = 0 Then
                'TEXT
                J = InStr(1, lw1.ListItems(I).Text, Text1.Text, vbTextCompare)
            Else
                J = InStr(1, lw1.ListItems(I).SubItems(lw1.SortKey), Text1.Text, vbTextCompare)
            End If
            If J > 0 Then Exit For
            
        Next
        
    End If
    If J > 0 Then
        lw1.ListItems(I).Selected = True
        Set lw1.SelectedItem = lw1.ListItems(I)
    Else
        MsgBox "No se ha encotrado ninguna coincidencia", vbExclamation
    End If
    lw1.SelectedItem.EnsureVisible
    lw1.SetFocus
        
    
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        
        CargaDatos_
        
        If Me.lw1.ListItems.Count > 0 Then Set lw1.SelectedItem = lw1.ListItems(1)
    End If
 
End Sub

Private Sub CargaDatos_()
    Screen.MousePointer = vbHourglass
    lw1.ListItems.Clear
    CargaClientesPotenciales2
    CargaExpedientesClientes2
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimVez = True
    Me.Icon = frmPPalWhose.Icon
    Set lw1.SmallIcons = Me.ImageList1
    
    
    lw1.SortOrder = lvwAscending
    lw1.Sorted = True
    Leerorden True
    
    cmdCancelar(0).Cancel = True
End Sub



Private Sub PonerFrameVisible(ByRef F As Frame)
    F.Top = 0
    F.Left = 120
    F.visible = True
    
    Height = F.Height + 420
    Width = F.Width + 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Leerorden False
End Sub

Private Sub frmCliPot_DatoSeleccionado(CadenaSeleccion As String)
    Cad = CadenaSeleccion
End Sub








Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lw1.SortKey = ColumnHeader.Index - 1 Then
        If lw1.SortOrder = lvwAscending Then
            lw1.SortOrder = lvwDescending
        Else
            lw1.SortOrder = lvwAscending
        End If
    Else
        lw1.SortKey = ColumnHeader.Index - 1
        lw1.SortOrder = lvwAscending
        Label2.Caption = "Buscar " & ColumnHeader.Text
    End If
End Sub

Private Sub lw1_DblClick()
Dim Clave As String


    If lw1.SelectedItem Is Nothing Then Exit Sub
    
    Clave = lw1.SelectedItem.Key
    
    If lw1.SelectedItem.SmallIcon = 1 Then
        'POTENCIAL
        frmWH_ExpedientesPot.ClientePot = Val(lw1.SelectedItem.Text)
        frmWH_ExpedientesPot.Show vbModal
        
        
        
    Else
        'NORMAL
        Volver_A_Cargar_Datos = False
        frmWH_Expedientes.Cliente = CLng(lw1.SelectedItem.Text)
        If Trim(lw1.SelectedItem.SubItems(2)) <> "" Then
            frmWH_Expedientes.Ano = CInt(lw1.SelectedItem.SubItems(3))
            frmWH_Expedientes.expediente = CLng(lw1.SelectedItem.SubItems(2))
        Else
            frmWH_Expedientes.expediente = 0
        End If
        frmWH_Expedientes.Show vbModal
        
        
        
    End If
    
    
    
    'Hay que refrescar
    If Volver_A_Cargar_Datos Then
        CadenaDesdeOtroForm = Me.lw1.SelectedItem.Text & "|" & Me.lw1.SelectedItem.SmallIcon & "|" & Trim(Me.lw1.SelectedItem.SubItems(2)) & "|" & Trim(Me.lw1.SelectedItem.SubItems(3)) & "|"
                
        CargaDatos_
            
            
        Cad = RecuperaValor(CadenaDesdeOtroForm, 1)
        For NumRegElim = 1 To lw1.ListItems.Count
            
            If Cad = lw1.ListItems(NumRegElim).Text Then
                If RecuperaValor(CadenaDesdeOtroForm, 2) = lw1.ListItems(NumRegElim).SmallIcon Then
                    If RecuperaValor(CadenaDesdeOtroForm, 3) = Trim(lw1.ListItems(NumRegElim).SubItems(2)) Then
                        If RecuperaValor(CadenaDesdeOtroForm, 4) = Trim(lw1.ListItems(NumRegElim).SubItems(3)) Then
                            'ESTE ES
                            Set lw1.SelectedItem = lw1.ListItems(NumRegElim)
                            lw1.ListItems(NumRegElim).EnsureVisible
                            Exit For
                        End If
                    End If
                End If
            End If
        
        Next NumRegElim

    End If
    
    
End Sub

Private Sub Leerorden(Leer As Boolean)
Dim NF As Integer

    On Error GoTo eLeerorden

    NF = FreeFile
    Cad = App.Path & "\whoselecexp.dat"
    If Leer Then
        
        If Dir(Cad, vbArchive) <> "" Then
            
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            If Cad <> "" Then
                If Not IsNumeric(Cad) Then
                    Cad = ""
                Else
                    If Val(Cad) > 3 Then Cad = ""
                End If
            End If
        Else
            Cad = ""
        End If
        If Cad = "" Then Cad = "0"
        Me.lw1.SortKey = CInt(Cad)
        OrdenIncial = lw1.SortKey
        Label2.Caption = "Buscar " & lw1.ColumnHeaders(OrdenIncial + 1).Text
        
    Else
        If OrdenIncial <> lw1.SortKey Then
            Open Cad For Output As #NF
            Print #NF, lw1.SortKey
            Close #NF
        End If
    End If
    Exit Sub
eLeerorden:
    Err.Clear
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub
