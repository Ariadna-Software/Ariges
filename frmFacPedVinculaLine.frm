VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacPedVinculaLine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Viincular lineas pedido-albaran"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11445
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&No vincular"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   9960
      TabIndex        =   4
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   8160
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NºPed"
         Object.Width           =   2461
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2432
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Referencia"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "LIn."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Solicitadas"
         Object.Width           =   2170
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Pdte."
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   661
      EndProperty
   End
   Begin VB.Label lblArticulo 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   8775
   End
   Begin VB.Label lblCliente 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   10815
   End
End
Attribute VB_Name = "frmFacPedVinculaLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public C As String
Public A As String


Private Sub Command1_Click(index As Integer)
    CadenaDesdeOtroForm = ""
    If index = 0 Then
        davidCodtipom = ""
        For NumRegElim = 1 To ListView1.ListItems.Count
            If Me.ListView1.ListItems(NumRegElim).Checked Then
                davidCodtipom = davidCodtipom & "X"
                davidNumalbar = NumRegElim
            End If
        Next
        
        
        For NumRegElim = 1 To ListView1.ColumnHeaders.Count
            Debug.Print ListView1.ColumnHeaders.Item(NumRegElim).Text & " " & ListView1.ColumnHeaders.Item(NumRegElim).Width
        Next
        
        If Len(davidCodtipom) <> 1 Then
            MsgBox "Selecciona una (y solo una) linea para vincular", vbExclamation
            Exit Sub
        End If
        
        CadenaDesdeOtroForm = " numpedcl=" & ListView1.ListItems(davidNumalbar).Text & " AND numlinea = " & ListView1.ListItems(davidNumalbar).SubItems(3)
        
        
    End If
    davidCodtipom = ""
    davidNumalbar = 0
    Unload Me
End Sub

Private Sub Form_Activate()
Dim It As ListItem

    
    
    
    If Val(Me.Tag) = 0 Then
        Me.Tag = 1
        ListView1.ListItems.Clear
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Me.lblCliente.Tag, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
            Set It = ListView1.ListItems.Add()
            'scaped.numpedcl,referenc,fecpedcl,numlinea,solicitadas,servidas
            It.Text = Format(miRsAux!NumPedcl, "00000")
            It.SubItems(1) = miRsAux!fecpedcl
            It.SubItems(2) = miRsAux!referenc
            It.SubItems(3) = miRsAux!numlinea
            It.SubItems(4) = miRsAux!solicitadas
            It.SubItems(5) = miRsAux!cantidad
            It.SubItems(6) = " "
            If miRsAux!cerrado = 1 Then
                It.SubItems(6) = "*"
                It.ListSubItems(6).ToolTipText = "Cerrado"
            End If
            If DBLet(miRsAux!OrdRef, "N") = 0 Then
                It.ListSubItems(2).ForeColor = vbRed
                It.ListSubItems(2).ToolTipText = "Misma referencia/obra"
            End If
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    Me.Tag = 0
    frmFacPedVinculaLine.lblCliente.Caption = C
    frmFacPedVinculaLine.lblArticulo.Caption = A
    
End Sub

