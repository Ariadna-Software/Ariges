VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSiiAvisos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avisos  SII"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContabilizar 
      Caption         =   "Contabilizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   7320
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11033
      SortKey         =   4
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   3952
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Factura"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre"
         Object.Width           =   8185
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Importe"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "FechaOculta"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ImporteOculto"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CheckBox chkTiket 
         Caption         =   "Excluir tickets"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   270
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSiiAvisos.frx":0000
         Left            =   360
         List            =   "frmSiiAvisos.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "0"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Facturas pendientes contabilizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   5280
         TabIndex        =   5
         Top             =   120
         Width           =   5805
      End
   End
End
Attribute VB_Name = "frmSiiAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Primeravez As Boolean
Dim SQL As String
Dim AgrupaTickets As Boolean

Private Sub chkTiket_Click()
    CargaFacturas
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdContabilizar_Click()
     AbrirListado 223
     Screen.MousePointer = vbHourglass
    CargaFacturas
    Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_Click()
    If Primeravez Then Exit Sub
    If Combo1.Tag = Combo1.ListIndex Then Exit Sub
    
    CargaFacturas
    Combo1.Tag = Combo1.ListIndex
    
End Sub

Private Sub Form_Activate()
    If Primeravez Then
        Primeravez = False
        CargaFacturas
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    Primeravez = True
    Set ListView1.SmallIcons = frmPpal.ImgListPpal
        
        
    AgrupaTickets = False
    SQL = DevuelveDesdeBD(conAri, "contador", "stipom", "codtipom", "FTG", "T")
    If SQL <> "" Then
        If Val(SQL) > 0 Then AgrupaTickets = True
    End If
    If AgrupaTickets Then
        chkTiket.visible = False
    Else
        SQL = DevuelveDesdeBD(conAri, "contador", "stipom", "codtipom", "'FTI'")
        If SQL = "" Then SQL = "0"
        chkTiket.visible = Val(SQL) > 0
    End If
    Combo1.ListIndex = 0
     
    If vParamAplic.SII_Tiene Then
        Caption = "Avisos SII"
    Else
        Caption = "Pendiente contabilidad"
        If vParamAplic.NumeroInstalacion = vbFenollar Then cmdContabilizar.visible = True
    End If
End Sub

'0 Todas
'1 Cli
'2 Pro
Private Sub CargaFacturas()
Dim Todas As Byte
    Screen.MousePointer = vbHourglass
    Me.ListView1.ListItems.Clear
    NumRegElim = 0
    
    Todas = Combo1.ListIndex
    If Todas <> 2 Then
        'Clientes
        SQL = DameSQL(True)
        CargaListView (True)
    End If
    If Todas <> 1 Then
        'Proeedores
        SQL = DameSQL(False)
        CargaListView (False)
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CargaListView(Cli As Boolean)
Dim It As ListItem
Dim FechaLimite As Date
Dim I As Byte
Dim Color As Long

    On Error GoTo eCargaListView
    Set miRsAux = New ADODB.Recordset
    
    'FechaLimite = DateAdd("d", -1 * vParamAplic.Sii_Dias, Now)
    FechaLimite = UltimaFechaCorrectaSII(vParamAplic.Sii_Dias, Now)
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Set It = ListView1.ListItems.Add(, "C" & Format(Now, "yymmdd" & Format(NumRegElim, "0000")))
        
        
        It.Text = Format(miRsAux!Fecha, "dd/mm/yyyy")
 
        
        It.SubItems(1) = miRsAux!Factura
        It.SubItems(2) = miRsAux!Nombre
        
        
        It.SubItems(3) = Format(miRsAux!Importe, FormatoImporte)
        
        If Not Cli Then It.SmallIcon = 11
                    
        If vParamAplic.SII_Tiene Then
            Color = -1
            If miRsAux!Fecha < FechaLimite Then
                Color = vbRed
            Else
                If miRsAux!Fecha = FechaLimite Then Color = vbBlue
            End If
                
            If Color <> -1 Then
                It.ForeColor = Color
                For I = 1 To It.ListSubItems.Count
                    It.ListSubItems(I).ForeColor = Color
                Next
            End If
        End If
        It.SubItems(4) = Format(miRsAux!Fecha, "yyyymmdd") & Val(Not Cli)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Me.Refresh
eCargaListView:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub



Private Function DameSQL(Clientes As Boolean) As String
Dim cad As String

    

    If Clientes Then
    
        cad = "Select fecfactu fecha, concat(codtipom,numfactu) factura, codclien,nomclien nombre,totalfac importe FROM scafac"
        cad = cad & " WHERE fecfactu>=" & DBSet(vParamAplic.Sii_Finicio, "F") & " AND  "
        cad = cad & "fecfactu<=DATE_ADD(now(), INTERVAL -1 DAY) AND  intconta =0 "
        
        If AgrupaTickets Then
            cad = cad & " AND codtipom <>'FTI'    "
        Else
            If Me.chkTiket.Value Then cad = cad & " AND codtipom <>'FTI'    "
        End If
    Else
        cad = "Select fecrecep fecha, numfactu factura, codprove,nomprove nombre,totalfac importe  FROM scafpc WHERE "
        cad = cad & "fecrecep>=" & DBSet(vParamAplic.Sii_Finicio, "F") & " AND  "
        cad = cad & "fecrecep<=DATE_ADD(now(), INTERVAL -1 DAY) AND  intconta =0 "

    
    End If
    
    DameSQL = cad

End Function
