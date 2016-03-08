VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmADVvarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   Icon            =   "frmADVvarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDH 
      Height          =   4575
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdBusqueda 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   13
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdBusqueda 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3120
         TabIndex        =   12
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   3480
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3480
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   3000
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3000
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   2160
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2160
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1680
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1680
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   960
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   960
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   480
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Cod. "
         Text            =   "Text1"
         Top             =   480
         Width           =   765
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   5
         Left            =   960
         Picture         =   "frmADVvarios.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmADVvarios.frx":010E
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   3
         Left            =   960
         Picture         =   "frmADVvarios.frx":0210
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   2
         Left            =   960
         Picture         =   "frmADVvarios.frx":0312
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmADVvarios.frx":0414
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmADVvarios.frx":0516
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   28
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   27
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   26
         Top             =   3030
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   25
         Top             =   1710
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Variedad"
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
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
         Index           =   14
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame FrameSelecCampo 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton cmdBusq 
         Height          =   375
         Left            =   960
         Picture         =   "frmADVvarios.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   6240
         Width           =   375
      End
      Begin VB.CommandButton cmdSelCampo 
         Caption         =   "Regresar"
         Height          =   495
         Left            =   7440
         TabIndex        =   3
         Top             =   6240
         Width           =   975
      End
      Begin MSComctlLib.ListView lw11 
         Height          =   5775
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Partida"
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Variedad"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Sup(ha)"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Socio"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nombre"
            Object.Width           =   3880
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   495
         Index           =   0
         Left            =   8640
         TabIndex        =   1
         Top             =   6240
         Width           =   975
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "frmADVvarios.frx":101A
         ToolTipText     =   "Puntear al haber"
         Top             =   6240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmADVvarios.frx":1164
         ToolTipText     =   "Quitar al haber"
         Top             =   6240
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmADVvarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Public Opcion As Byte
    '0.- Mostrar campos para seleccionar en partes de trabajo


Public vCampos As String

Dim PrimVez As Boolean
Dim SQL As String
Dim It As ListItem

''''''
''''''Private Sub chkTodos_Click()
''''''    Screen.MousePointer = vbHourglass
''''''    CargaCampos
''''''    Screen.MousePointer = vbDefault
''''''End Sub

Private Sub cmdBusq_Click()
    
    Me.FrameDH.visible = True
    Me.FrameSelecCampo.Enabled = False
    PonerFoco Text1(0)
End Sub

Private Sub cmdBusqueda_Click(Index As Integer)
     If Index = 0 Then
        SQL = ""
            'rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
            'variedades on rcampos.codvarie = variedades.codvarie)"
            'SQL = SQL & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
        If Text1(0).Text <> "" Then SQL = SQL & " AND rcampos.codsocio >= " & Text1(0).Text
        If Text1(1).Text <> "" Then SQL = SQL & " AND rcampos.codsocio <= " & Text1(1).Text
        If Text1(2).Text <> "" Then SQL = SQL & " AND rcampos.codclien >= " & Text1(2).Text
        If Text1(3).Text <> "" Then SQL = SQL & " AND rcampos.codclien <= " & Text1(3).Text
        If Text1(4).Text <> "" Then SQL = SQL & " AND rcampos.codvarie >= " & Text1(4).Text
        If Text1(5).Text <> "" Then SQL = SQL & " AND rcampos.codvarie <= " & Text1(5).Text
        If SQL <> "" Then SQL = Mid(SQL, 5)
    Else
        SQL = "rcampos.codclien  = " & vCampos
    End If
    CargaCampos2 SQL
    
    'Uno u otro
    Me.FrameDH.visible = False
    Me.FrameSelecCampo.Enabled = True
    
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)

    CadenaDesdeOtroForm = ""  'por si las moscas
    Unload Me
End Sub

Private Sub cmdSelCampo_Click()
    If lw11.ListItems.Count = 0 Then Exit Sub
    
    SQL = ""
    For NumRegElim = 1 To lw11.ListItems.Count
        If lw11.ListItems(NumRegElim).Checked Then SQL = SQL & "1"
    Next
    If SQL = "" Then
        MsgBox "Seleccione algun campo", vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = ""
    NumRegElim = Len(SQL)
    If NumRegElim > 1 Then
        SQL = "Va a insertar " & NumRegElim & " campos. ¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        CadenaDesdeOtroForm = "@" 'comienza por arroba
    End If

    
    For NumRegElim = 1 To lw11.ListItems.Count
        If lw11.ListItems(NumRegElim).Checked Then
            SQL = lw11.ListItems(NumRegElim).Text & "|" & lw11.ListItems(NumRegElim).SubItems(1) & "|" & lw11.ListItems(NumRegElim).SubItems(2) & "|" & "·#"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
        End If
    Next
        
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        
        If Opcion = 0 Then CargaCampos2 "rcampos.codclien  = " & vCampos   'Martin. Enlaza con codclien
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    '
    Screen.MousePointer = vbHourglass
    PrimVez = True
    Me.Icon = frmPpal.Icon
    Me.FrameSelecCampo.visible = False
    limpiar Me
    Select Case Opcion
    Case 0
        Caption = "Campos"
        PonerFrameVisible Me.FrameSelecCampo
    End Select
    
    Me.cmdCancelar(Opcion).Cancel = True

End Sub


Private Sub PonerFrameVisible(Fr As Frame)
    Fr.visible = True
    Fr.Top = 0
    Fr.Left = 120
    Me.Height = Fr.Height + 480
    Me.Width = Fr.Width + 320
End Sub



'----------------------------------
Private Sub CargaCampos2(ByVal SQ As String)

    On Error GoTo ecargaCampos
    Set miRsAux = New ADODB.Recordset
    
    Me.lw11.ListItems.Clear
    'Para no meter MUCHOS ariagro.tabla
    'Pongo @# y luego lo reemplazo por vparamaplic.Ariagro.
    SQL = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,rcampos.supsigpa"
    SQL = SQL & " from ((@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
    SQL = SQL & " inner join @#variedades on rcampos.codvarie = variedades.codvarie)"
    SQL = SQL & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
    'where socio
    If SQ <> "" Then SQL = SQL & " WHERE " & SQ
    
    SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = lw11.ListItems.Add()
        It.Text = Format(miRsAux!codCampo, "0000000")
        It.SubItems(1) = miRsAux!nomparti
        It.SubItems(2) = miRsAux!nomvarie
        'Superficie
        
        It.SubItems(3) = Format(DBLet(miRsAux!supsigpa, "N"), FormatoPrecio)
        
        If IsNull(miRsAux!codclien) Then
            It.SubItems(4) = " "
        Else
            It.SubItems(4) = Format(miRsAux!codclien, "00000")
        End If
        It.SubItems(5) = Format(miRsAux!codsocio, "00000")
        It.SubItems(6) = miRsAux!nomsocio
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
ecargaCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    SQL = CadenaDevuelta
End Sub

Private Sub imgBusc_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    
    If Index < 2 Then
        frmB.vCampos = "Codigo|" & vParamAplic.Ariagro & ".rsocios|codsocio|N|0000000|20·Nombre|" & vParamAplic.Ariagro & ".rsocios|nomsocio|T||70·"
        frmB.vTabla = vParamAplic.Ariagro & ".rsocios"
        frmB.vTitulo = "Socios"
    ElseIf Index < 4 Then
        frmB.vCampos = "Codigo|sclien|codclien|N|0000000|20·Nombre|sclien|nomclien|T||70·"
        frmB.vTabla = "sclien"
        frmB.vTitulo = "Clientes"
    Else
        
        frmB.vCampos = "Codigo|" & vParamAplic.Ariagro & ".variedades|codvarie|N|0000000|20·Nombre|" & vParamAplic.Ariagro & ".variedades|nomvarie|T||70·"
        frmB.vTabla = vParamAplic.Ariagro & ".variedades"
        frmB.vTitulo = "Variedades"
    End If
    frmB.vSQL = ""
    
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
    SQL = ""
    frmB.Show vbModal
    
    If SQL <> "" Then
        Text1(Index).Text = RecuperaValor(SQL, 1)
        Text2(Index).Text = RecuperaValor(SQL, 2)
        SQL = ""
        If Index = 5 Then
            PonerFocoBtn Me.cmdBusqueda(0)
        Else
            PonerFoco Text1(Index + 1)
        End If
    End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
        
    For NumRegElim = 1 To lw11.ListItems.Count
        lw11.ListItems(NumRegElim).Checked = Index = 1
    Next
End Sub

Private Sub lw11_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index - 1 <> Me.lw11.SortKey Then
        lw11.SortKey = ColumnHeader.Index - 1
        lw11.SortOrder = lvwAscending
    Else
        If lw11.SortOrder = lvwAscending Then
            lw11.SortOrder = lvwDescending
        Else
            lw11.SortOrder = lvwAscending
        End If
    End If
        
End Sub

Private Sub lw11_DblClick()
    cmdSelCampo_Click
End Sub

Private Sub lw11_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Set lw11.SelectedItem = Item
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    SQL = ""
    If Text1(Index).Text <> "" Then
        If Not PonerFormatoEntero(Text1(Index)) Then
            Text1(Index).Text = ""
        
        Else
            If Index < 2 Then
                'Socio
                SQL = DevuelveDesdeBD(conAri, "nomsocio", vParamAplic.Ariagro & ".rsocios", "codsocio", Text1(Index).Text)
                If SQL = "" Then SQL = "NO existe el socio"
            ElseIf Index < 4 Then
                SQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text)
                If SQL = "" Then SQL = "NO existe el cliente"
                    
            Else
                'Variedad
                SQL = DevuelveDesdeBD(conAri, "nomvarie", vParamAplic.Ariagro & ".variedades", "codvarie", Text1(Index).Text)
                If SQL = "" Then SQL = "NO existe la variedad"
            End If
            
        End If
    End If
    Text2(Index).Text = SQL
End Sub
