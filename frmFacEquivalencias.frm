VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacEquivalencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equivalencias entre articulos"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Portapapeles"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3625
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   6421
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Stock"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmFacEquivalencias.frx":0000
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Equivalencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmFacEquivalencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vCodartic As String

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo EC
    If Me.ListView1.ListItems.Count = 0 Then Exit Sub
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub



    Clipboard.Clear
    Clipboard.SetText Me.ListView1.SelectedItem.Text
    Unload Me
    
EC:
    If Err.Number <> 0 Then MuestraError Err.Number, "Portapapeles"
End Sub

Private Sub Form_Load()
Dim RN As ADODB.Recordset
Dim IT

    Me.Icon = frmPpal.Icon
    Set RN = New ADODB.Recordset
    RN.Source = "Select codarti1,nomartic,canstock from sarti6,sartic,salmac where "
    
    RN.Source = RN.Source & " codarti1=sartic.codartic and sartic.codartic=salmac.codartic and codalmac=" & RecuperaValor(vCodartic, 2)
    vCodartic = RecuperaValor(vCodartic, 1)
    RN.Source = RN.Source & " and sarti6.codartic=" & DBSet(vCodartic, "T") & " ORDER BY nomartic"
    vCodartic = RN.Source
    RN.Open vCodartic, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        Set IT = Me.ListView1.ListItems.Add()
        IT.Text = RN!codarti1
        IT.SubItems(1) = RN!NomArtic
        IT.SubItems(2) = Format(RN!CanStock, FormatoCantidad)
        RN.MoveNext
    Wend
    RN.Close
    Set RN = Nothing
End Sub
