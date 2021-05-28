VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión empresa"
   ClientHeight    =   6630
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6600
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3917.224
   ScaleMode       =   0  'User
   ScaleWidth      =   6197.042
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtlargo 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   3945
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   1440
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   4710
      Left            =   120
      TabIndex        =   0
      Top             =   1215
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   8308
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3006
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   3381
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   3900
      TabIndex        =   1
      Top             =   6120
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   5220
      TabIndex        =   2
      Top             =   6120
      Width           =   1140
   End
   Begin VB.Label lblAriges 
      Caption         =   "AriGes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Usuario:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione una de las empresas disponibles para el usuario:"
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
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   5595
   End
   Begin VB.Label lblLabels 
      Caption         =   "Empresas:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   885
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   120
      Top             =   6000
      Width           =   2280
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":0894
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cad As String
Dim ItmX As ListItem
Dim RS As Recordset

    
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
Dim OK As Boolean
  
    If lw1.ListItems.Count = 0 Then
        MsgBox "Ninguna empresa para seleccionar", vbExclamation
        Exit Sub
    End If
    If lw1.SelectedItem Is Nothing Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Sub
    End If

    
    Screen.MousePointer = vbHourglass
    

    CadenaDesdeOtroForm = lw1.SelectedItem.Tag
    'ASignamos la cadena de conexion
    vUsu.CadenaConexion = RecuperaValor(lw1.SelectedItem.Tag, 1)
            
    'Comprobamos ,k la empresa no este bloqueada
    conn.Execute "SET AUTOCOMMIT=0"
    If ComprobarEmpresaBloqueada(vUsu.Codigo, vUsu.CadenaConexion) Then
        Cad = "BLOQ"
    Else
        Cad = ""
    End If
    conn.Execute "SET AUTOCOMMIT=1"
        
    If Cad <> "" Then GoTo Salida   'Empresa bloqueada
            
    'Cerramos la ventana
    Unload Me
    
Salida:
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    CargaImagen
    lw1.SmallIcons = Me.ImageList1
    Me.txtUser.Text = vUsu.Login
    Me.txtlargo.Text = vUsu.Nombre
'    lw1.ColumnHeaders(1).Width = lw1.Width - 1500
'    lw1.ColumnHeaders(2).Width = 1100

    'Cargamos las empresas disponibles
    BuscaEmpresas
    NumeroEmpresaMemorizar True
End Sub


Private Sub CargaImagen()
On Error Resume Next
    Image2.Picture = LoadPicture(App.Path & "\minilogo.bmp")
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NumeroEmpresaMemorizar False
End Sub


Private Sub lblAriges_DblClick()
  '  If vParamAplic.NumeroInstalacion = vbFenollar Then
        HaMostradoCanal2_El_B = True
        lblAriges.Font.Bold = False
  '  End If
End Sub

Private Sub lw1_DblClick()
   cmdOk_Click
End Sub


'===== Laura : 16/02/05
Private Function DevuelveProhibidas() As String
'Private Function DevuelvePermitidas() As String
Dim i As Integer
Dim SQL As String
On Error GoTo EDevuelveProhibidas
    
    DevuelveProhibidas = ""
    Set RS = New ADODB.Recordset
    i = vUsu.Codigo Mod 1000
'    i = vUsu.Codigo
    'Si es Root todas permitidas
'    If vUsu.Nivel = 0 Then
'        SQL = "Select * from usuarios.empresasariges "
'    Else
        SQL = "Select * from usuarios.usuarioempresasariges WHERE codusu =" & i
'    End If
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        Cad = Cad & RS!codempre & "|"
        RS.MoveNext
    Wend
    If Cad <> "" Then Cad = "|" & Cad
    RS.Close
    DevuelveProhibidas = Cad
    
EDevuelveProhibidas:
    Err.Clear
    Set RS = Nothing
End Function


Private Sub BuscaEmpresas()
Dim Prohibidas As String
Dim SQL As String


'Cargamos las Empresas Permitidas para el usuario
Prohibidas = DevuelveProhibidas

'Cargamos las empresas
Set RS = New ADODB.Recordset
'SQL = "Select * from  usuarios.usuarioempresasariges INNER JOIN usuarios.empresasariges "
'SQL = SQL & " ON usuarios.usuarioempresasariges.codempre=usuarios.empresasariges.codempre "
'SQL = SQL & "WHERE codusu=" & vUsu.Codigo & " ORDER BY usuarios.usuarioempresasariges.codempre"
SQL = "Select * from usuarios.empresasariges ORDER BY Codempre"
RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

While Not RS.EOF
    Cad = "|" & RS!codempre & "|"
    If InStr(1, Prohibidas, Cad) = 0 Then
        Cad = RS!nomempre
        Set ItmX = lw1.ListItems.Add()
        
        ItmX.Text = Cad
        ItmX.SubItems(1) = RS!nomresum
        Cad = RS!AriGes & "|" & RS!nomresum & "|" & RS!Usuario & "|" & RS!Pass & "|"
        ItmX.Tag = Cad
        ItmX.ToolTipText = RS!AriGes & " - " & RS!codempre
        ItmX.SmallIcon = 1
    End If
    RS.MoveNext
Wend

'While Not RS.EOF
'    cad = "|" & RS!codempre & "|"
'    If InStr(1, Prohibidas, cad) = 0 Then
'        cad = RS!nomempre
'        Set ItmX = lw1.ListItems.Add()
'
'        ItmX.Text = cad
'        ItmX.SubItems(1) = RS!nomresum
'        cad = RS!CONTA & "|" & RS!nomresum & "|" & RS!Usuario & "|" & RS!Pass & "|"
'        ItmX.Tag = cad
'        ItmX.ToolTipText = RS!CONTA
'        ItmX.SmallIcon = 1
'    End If
'    RS.MoveNext
'Wend
'RS.Close


RS.Close
End Sub




'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim C1 As String
On Error GoTo ENumeroEmpresaMemorizar


    If Leer Then
        If CadenaDesdeOtroForm <> "" Then
            'Ya estabamos trabajando con la aplicacion
            
'            If Not (vEmpresa Is Nothing) Then
'                 For NF = 1 To Me.lw1.ListItems.Count
'                    If lw1.ListItems(NF).Text = vEmpresa.nomempre Then
'                        Set lw1.SelectedItem = lw1.ListItems(NF)
'                        Exit For
'                    End If
'                Next NF
'            End If
'
'                'El tercer pipe, si tiene es el ancho col1
'                Cad = AnchoLogin
'                C1 = RecuperaValor(Cad, 3)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 4360
'                End If
'                lw1.ColumnHeaders(1).Width = NF
'                'El cuarto pipe si tiene es el ancho de col2
'                C1 = RecuperaValor(Cad, 4)
'                If Val(C1) > 0 Then
'                    NF = Val(C1)
'                Else
'                    NF = 1400
'                End If
'                lw1.ColumnHeaders(2).Width = NF
'
'
'            CadenaDesdeOtroForm = ""
'            Exit Sub
        End If
    End If
    Cad = App.Path & "\ultempre.dat"
    If Leer Then
        If Dir(Cad) <> "" Then
            NF = FreeFile
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            If Cad <> "" Then
            
                
            
            
                'El primer pipe es el usuario. Como ya no lo necesito, no toco nada
                
                C1 = RecuperaValor(Cad, 2)
                'el segundo es el
                If C1 <> "" Then
                    For NF = 1 To Me.lw1.ListItems.Count
                        If lw1.ListItems(NF).Text = C1 Then
                            Set lw1.SelectedItem = lw1.ListItems(NF)
                            lw1.SelectedItem.EnsureVisible
                            Exit For
                        End If
                    Next NF
                End If
                
                'El tercer pipe, si tiene es el ancho col1
                C1 = RecuperaValor(Cad, 3)
                If Val(C1) > 0 Then
                    NF = Val(C1)
                Else
                    NF = 4360
                End If
                lw1.ColumnHeaders(1).Width = NF
                'El cuarto pipe si tiene es el ancho de col2
                C1 = RecuperaValor(Cad, 4)
                If Val(C1) > 0 Then
                    NF = Val(C1)
                Else
                    NF = 1400
                End If
                lw1.ColumnHeaders(2).Width = NF
            End If
        End If
    Else 'Escribir
        NF = FreeFile
        Open Cad For Output As #NF
        Cad = "NO ncesito|"
        If lw1.ListItems.Count > 0 Then
            Cad = Cad & lw1.SelectedItem.Text
        Else
            Cad = Cad & " "
        End If
        Cad = Cad & "|" & Int(Round(lw1.ColumnHeaders(1).Width, 2)) & "|" & Int(Round(lw1.ColumnHeaders(2).Width, 2)) & "|"
        AnchoLogin = Cad
        Print #NF, Cad
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub

