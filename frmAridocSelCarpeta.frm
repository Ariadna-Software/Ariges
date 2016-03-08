VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAridocSelCarpeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargando datos"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   2
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   1
      Top             =   6720
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   6600
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
            Picture         =   "frmAridocSelCarpeta.frx":0000
            Key             =   "abierto"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAridocSelCarpeta.frx":6862
            Key             =   "cerrado"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11245
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmAridocSelCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Rs As ADODB.Recordset

Public vCodCarpeta As Long   'Para marcar por defecto una carpeta

Dim PrimeraVez As Boolean


Dim CadenaCarpetas As String


Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        If Not TreeView1.SelectedItem Is Nothing Then
            CadenaDesdeOtroForm = Mid(TreeView1.SelectedItem.Key, 2) & "|" & TreeView1.SelectedItem.FullPath & "|"
            
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    
    If PrimeraVez Then
        PrimeraVez = False
        CargaArbol
        
    End If
    Screen.MousePointer = vbDefault
    Caption = "Seleccionar carpeta Aridoc"
End Sub

Private Sub Form_Load()
    PrimeraVez = True

    
    
    
End Sub







Private Sub CargaArbol()
Dim Cad As String
Dim Nod As Node
Dim C As Integer
Dim i As Integer
Dim Contador2 As Integer



    TreeView1.Nodes.Clear
    
    TreeView1.ImageList = Me.ImageList1

    CadenaCarpetas = "|"
    
    If Rs!padre <> 0 Then
        MsgBox "Error en primer NODO. Padre != 0", vbExclamation
        Exit Sub
    End If
    C = 0
    i = 0
    While i = 0
        INSERTAR_NODO Rs, 1
        Rs.MoveNext
        If Rs.EOF Then
            i = 1
        Else
            If Rs!padre <> 0 Then i = 1
        End If
        C = C + 1
    Wend
    
    'Cargo el segundo nivel
    Contador2 = TreeView1.Nodes.Count
    C = 0
    For i = 1 To Contador2
        Cad = Mid(TreeView1.Nodes(i).Key, 2)
        Rs.MoveFirst
        Rs.Find " padre = " & Cad, , adSearchForward, 1
        While Not Rs.EOF
            C = C + 1
            If Rs!padre = Cad Then
                INSERTAR_NODO Rs, 2
            Else
                Rs.MoveLast
                
            End If
            Rs.MoveNext
        Wend
    Next i
       
    If C > 0 Then
                'Cargo el tercer nivel
                C = Contador2 + 1
                Contador2 = TreeView1.Nodes.Count
                For i = C To Contador2
                    Cad = Mid(TreeView1.Nodes(i).Key, 2)
                    Rs.MoveFirst
                    Rs.Find " padre = " & Cad, , adSearchForward, 1
                    While Not Rs.EOF
                        C = C + 1
                        If Rs!padre = Cad Then
                            INSERTAR_NODO Rs, 2
                        Else
                            Rs.MoveLast
                        End If
                        Rs.MoveNext
                    Wend
                Next i
                
                            
                            
                       
                            
                            
                            C = Contador2 + 1
                            Contador2 = TreeView1.Nodes.Count
                            If Contador2 >= C Then
                                For i = C To Contador2
                                    
                                    CargaArbolRecursivo Mid(TreeView1.Nodes(i).Key, 2), Rs, 5
                                  
                                Next i
                            End If
 
                End If '3 nivel
    
    
    
        
    Rs.Close
    If vCodCarpeta = 0 Then
        If TreeView1.Nodes.Count > 2 Then TreeView1.Nodes(3).EnsureVisible
    End If
    TreeView1.SetFocus
End Sub




Private Function INSERTAR_NODO(ByRef RSS As Recordset, SubNivel As Integer) As Integer
Dim XNodo As Node
Dim Cortar11 As String


On Error GoTo EIns_Nodo

    
    

    INSERTAR_NODO = -1
    If RSS!padre = 0 Then
        'NODO RAIZ
        Set XNodo = TreeView1.Nodes.Add(, tvwChild, "C" & RSS!codcarpeta)
    Else
    
        'NODO HIJO
        Set XNodo = TreeView1.Nodes.Add("C" & RSS!padre, tvwChild, "C" & RSS!codcarpeta)
    End If
    
    XNodo.Text = RSS!Nombre
    'En el tag metemos la seguriad
    XNodo.Tag = RSS!escriturau & "|" & RSS!escriturag & "|"
    
    

    'XNODO.Expanded = True
    CadenaCarpetas = CadenaCarpetas & Mid(XNodo.Key, 2) & "|"
    
    
    XNodo.Image = "cerrado"
    XNodo.ExpandedImage = "abierto"

    If vCodCarpeta = RSS!codcarpeta Then
        XNodo.EnsureVisible
        XNodo.Selected = True
    End If

    If RSS!hijos > 0 Then INSERTAR_NODO = XNodo.Index
'    End If
Exit Function
EIns_Nodo:
    Cortar11 = "ERROR GRAVE." & vbCrLf & vbCrLf
    Cortar11 = Cortar11 & Err.Description & vbCrLf & vbCrLf
    Cortar11 = Cortar11 & RSS!codcarpeta & " " & DBLet(RSS!Nombre, "T")
   ' MsgBox Cortar11, vbCritical
    Cortar11 = Cortar11 & vbCrLf & vbCrLf
    Cortar11 = Cortar11 & "Verifique ARIDOC. Si persiste avise a soporte técnico"
    Cortar11 = Cortar11 & vbCrLf & vbCrLf & vbCrLf & "¿FINALIZAR?"
    If MsgBox(Cortar11, vbCritical + vbYesNo) = vbYes Then
        Conn.Close
        End
    End If
End Function




Private Sub CargaArbolRecursivo(CarpePadre As String, ByRef RS1 As ADODB.Recordset, ByVal Nivel As Integer)
Dim C As Integer
Dim i As Integer
Dim CADENA As String
Dim Fin As Boolean
 
    'Este esta puesto para cuando es el arranque, que si le cuesta leer que no
    'bloquee el equipo
    If (TreeView1.Nodes.Count Mod 30) = 0 Then DoEvents


    CADENA = ""
    C = 0
    RS1.MoveFirst
    RS1.Find " padre = " & CarpePadre, , adSearchForward, 1
    Fin = RS1.EOF
    While Not Fin
        If RS1!padre = CarpePadre Then
        
            i = INSERTAR_NODO(RS1, Nivel)
            If i > 0 Then
                CADENA = CADENA & RS1!codcarpeta & "|"
                C = C + 1
            End If
            RS1.MoveNext
            If RS1.EOF Then Fin = True
        Else
            Fin = True
        End If
        

    Wend

    If C > 0 Then
        For i = 1 To C
            CargaArbolRecursivo (RecuperaValor(CADENA, i)), RS1, Nivel + 1
        Next i
    End If

End Sub




