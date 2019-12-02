VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAvisosAlb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avisos  albaranes"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   9480
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9340
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
      Width           =   10695
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
         ItemData        =   "frmAvisosAlb.frx":0000
         Left            =   360
         List            =   "frmAvisosAlb.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "0"
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Albaranes entregados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   435
         Left            =   4440
         TabIndex        =   4
         Top             =   120
         Width           =   5805
      End
   End
End
Attribute VB_Name = "frmAvisosAlb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim primeravez As Boolean
Dim SQL As String


Private Sub chkTiket_Click()
    CargaAlbanres
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Sub Combo1_Click()
    If primeravez Then Exit Sub
    If Combo1.Tag = Combo1.ListIndex Then Exit Sub
    
    CargaAlbanres
    Combo1.Tag = Combo1.ListIndex
    
End Sub

Private Sub Form_activate()
    If primeravez Then
        primeravez = False
        CargaAlbanres
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    primeravez = True
    Set ListView1.SmallIcons = frmPpal.ImgListPpal

    Combo1.ListIndex = 0
     
    
        Caption = "Avisos albaranes "
        
End Sub


Private Sub CargaAlbanres()
Dim Todas As Byte
    Screen.MousePointer = vbHourglass
    Me.ListView1.ListItems.Clear
    NumRegElim = 0
    
    Todas = Combo1.ListIndex
    
        SQL = DameSQL()
        CargaListView
    
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CargaListView()
Dim IT As ListItem

Dim I As Byte
Dim Color As Long

    On Error GoTo eCargaListView
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        Set IT = ListView1.ListItems.Add(, "C" & Format(Now, "yymmdd" & Format(NumRegElim, "0000")))
        
        
        IT.Text = Format(miRsAux!Fecha, "dd/mm/yyyy")
 
        
        IT.SubItems(1) = miRsAux!Factura
        IT.SubItems(2) = miRsAux!Nombre
        
        
        IT.SubItems(3) = Format(miRsAux!Importe, FormatoImporte)
        
        IT.SmallIcon = 11

        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Me.Refresh
eCargaListView:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub



Private Function DameSQL() As String
Dim cad As String

    

   
        cad = "Select scaalb.fechaalb fecha, concat(scaalb.codtipom,scaalb.numalbar) factura, codclien,nomclien nombre,sum(importel) importe "
        cad = cad & "  FROM scaalb,slialb WHERE scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar"
        cad = cad & " AND not fechaent is null"
        cad = cad & " GROUP BY  scaalb.codtipom,scaalb.fechaalb "
        cad = cad & " ORDER BY  scaalb.codtipom,scaalb.fechaalb "
    
   
    
    DameSQL = cad

End Function

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
        With frmFacEntAlbSAIL
            .hcoCodTipoM = Mid(ListView1.SelectedItem.SubItems(1), 1, 3)
            .hcoCodMovim = Mid(ListView1.SelectedItem.SubItems(1), 4)
            .Show vbModal
        End With
    
End Sub
