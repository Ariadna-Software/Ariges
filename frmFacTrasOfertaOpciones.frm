VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacTrasOfertaOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   8640
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "codigo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Desc"
         Object.Width           =   8220
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   1720
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Precio"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Dtos"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Importe"
         Object.Width           =   2011
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   6120
      Width           =   7455
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "frmFacTrasOfertaOpciones.frx":0000
      ToolTipText     =   "Puntear al haber"
      Top             =   6120
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "frmFacTrasOfertaOpciones.frx":014A
      ToolTipText     =   "Quitar al haber"
      Top             =   6120
      Width           =   240
   End
End
Attribute VB_Name = "frmFacTrasOfertaOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cad As String
Dim I As Integer


Private Sub Command1_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    
    If Index = 1 Then
        'Cancelar
        CadenaDesdeOtroForm = "NO"
    
    Else
        Cad = ""
        NumRegElim = 0
        For I = 1 To ListView3.ListItems.Count
            If ListView3.ListItems(I).Text <> "" Then
                NumRegElim = NumRegElim + 1
                If ListView3.ListItems(I).Checked Then Cad = Cad & "O"
            End If
        Next
        
        If Cad = "" Then
            MsgBox "Seleccione alguna linea para pasar al pedido.", vbExclamation
            Exit Sub
        End If
        
        If Len(Cad) <> NumRegElim Then
            NumRegElim = NumRegElim - Len(Cad)
            Cad = "Hay " & NumRegElim & " linea(s)  que no pasan al pedido.  ¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        
        
        
            'Montamos la cadena que llevara los numlinea
            For I = 1 To ListView3.ListItems.Count
                If ListView3.ListItems(I).Text <> "" Then
                    If ListView3.ListItems(I).Checked Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & ", " & Me.ListView3.ListItems(I).Tag
                End If
            Next
        
        Else
            'Van todas las lineas. No pongo nada
            CadenaDesdeOtroForm = ""
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    If ListView3.ListItems.Count > 0 Then Exit Sub
    
    CargaList
    
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon
    
    Me.ListView3.ListItems.Clear
End Sub


Private Sub CargaList()
Dim IT As ListItem
Dim I As Byte
    Set miRsAux = New ADODB.Recordset
    Cad = "Select * from slipre where numofert=" & Me.Caption & " ORDER BY numlinea"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView3.ListItems.Add
        IT.Checked = DBLet(miRsAux!esopcion, "N") <> 1
        IT.SubItems(1) = DBLet(miRsAux!NomArtic, "T")
        If miRsAux!codArtic = vParamAplic.ArtSeparador Then
            
            IT.Checked = False
            IT.ListSubItems(1).Bold = True
            IT.ListSubItems(1).ForeColor = vbBlue
            Cad = " "
            IT.SubItems(2) = Cad
            IT.SubItems(3) = Cad
            IT.SubItems(4) = Cad
            IT.SubItems(5) = Cad
            Cad = ""
            
        Else
            Cad = miRsAux!codArtic
            
        End If
        IT.Text = Cad
        
        
        
        If Cad <> "" Then
            'El articulo no es separador
            IT.SubItems(2) = Format(miRsAux!Cantidad, FormatoCantidad)
            IT.SubItems(3) = Format(miRsAux!precioar, FormatoPrecio)
            Cad = ""
            If DBLet(miRsAux!Dtoline1, "N") <> 0 Then Cad = miRsAux!Dtoline1
            If DBLet(miRsAux!Dtoline2, "N") <> 0 Then Cad = Cad & "  " & miRsAux!Dtoline2
            IT.SubItems(4) = Cad & " "
            
   
            IT.SubItems(5) = Format(miRsAux!ImporteL, FormatoImporte)
        End If
    
        
            
        
    
        IT.Tag = miRsAux!numlinea
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Caption = "Oferta: " & Caption
    Set miRsAux = Nothing

End Sub

Private Sub imgCheck_Click(Index As Integer)

Dim B As Boolean
    For I = 1 To ListView3.ListItems.Count
        If ListView3.ListItems(I).Text = "" Then
            B = False
        Else
            B = Index = 1
        End If
        ListView3.ListItems(I).Checked = B
    Next
End Sub
