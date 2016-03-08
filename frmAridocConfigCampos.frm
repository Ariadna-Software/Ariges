VERSION 5.00
Begin VB.Form frmAridocConfigCampos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion campos Aridoc"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   2160
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Index           =   0
      Left            =   9240
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   4
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   480
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Image imgBuscarC 
      Height          =   240
      Index           =   0
      Left            =   960
      Picture         =   "frmAridocConfigCampos.frx":0000
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Carpeta"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmAridocConfigCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ParaC1 As String
Public ParaC2 As String
Public Sel1 As String
Public Sel2 As String
Public Carpeta2 As String
Dim i As Integer


Public Sub CargarCombos()

Dim J As Integer
Dim Cad As String

             For J = 0 To 4
                Combo1(J).Clear
            Next J
            


    Do
        i = InStr(1, ParaC1, "|")
        If i = 0 Then
            ParaC1 = ""
        Else
            Cad = Mid(ParaC1, 1, i - 1)
            ParaC1 = Mid(ParaC1, i + 1)
            'Añadimos a los combs
            For J = 0 To 3
                Combo1(J).AddItem Cad
            Next J
            
        End If
    Loop Until ParaC1 = ""
        
    Do
        i = InStr(1, ParaC2, "|")
        If i = 0 Then
            ParaC2 = ""
        Else
            Cad = Mid(ParaC2, 1, i - 1)
            ParaC2 = Mid(ParaC2, i + 1)
            'Añadimos a los combs
            
                Combo1(4).AddItem Cad
            
            
        End If
    Loop Until ParaC2 = ""
        
    
    
    'Establecemos el selecccionado
    If Sel1 <> "" Then
        For i = 0 To 3
            Cad = RecuperaValor(Sel1, i + 1)
            If Cad <> "" Then
                J = Val(Cad)
                If J >= 0 Then Combo1(i).ListIndex = J

            End If
        Next i
            
    End If
    
    
    'Establecemos el selecccionado
    If Sel2 <> "" Then
        
            Cad = RecuperaValor(Sel2, 1)
            If Cad <> "" Then
                J = Val(Cad)
                If J >= 0 Then Combo1(4).ListIndex = J

            End If
        
            
    End If
    
       'La carpeta
       Text1(8).Text = RecuperaValor(Carpeta2, 1)
       Text2(1).Text = RecuperaValor(Carpeta2, 2)
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click(Index As Integer)
Dim J As Integer
    CadenaDesdeOtroForm = ""
    If Index = 1 Then
             For i = 0 To 4
                If Combo1(i).ListIndex < 0 Then
                    CadenaDesdeOtroForm = ""
                    MsgBox "Seleccione todas las opciones", vbExclamation
                    Exit Sub
                End If
            
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Combo1(i).ListIndex & "|"
                'Para separar de las fechas
                If i = 3 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "·"
            Next i
            
            'La carpeta
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "·" & Text1(8).Text & "|"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text2(1).Text & "|"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Command1(1).visible = vUsu.Nivel < 2
End Sub

Private Sub imgBuscarC_Click(Index As Integer)
Dim RS1 As ADODB.Recordset
Dim Cad As String


    CadenaDesdeOtroForm = ""
    
    'Monto el recordset. DEberia montarlo en el objeto ARIDOC
     Cad = " from carpetas"
    
    'Es el usuario propietario
    'If vUsu.CodUsu > 0 Then
    '    cad = cad & " WHERE "
    '    cad = cad & "userprop = " & vUsu.CodUsu
   '
   '     'O el grupo tiene permiso
   '     cad = cad & " OR (lecturag & " & vUsu.Grupo & ")"
   '
   ' End If

    
    
    'Ordenado por padre
    Cad = Cad & " ORDER BY Padre,nombre"
    
    
    Set RS1 = New ADODB.Recordset
    RS1.Open "select * " & Cad, ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
    If RS1.EOF Then
        MsgBox "Error en ARIDOC (1)", vbExclamation
        RS1.Close
    Else
        Set frmAridocSelCarpeta.Rs = RS1
        If Text1(8).Text <> "" Then
            Cad = Text1(8).Text
        Else
            Cad = "0"
        End If
        frmAridocSelCarpeta.vCodCarpeta = Val(Cad)
        frmAridocSelCarpeta.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            Text1(8).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Text2(1).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        End If
        
    End If
    Set RS1 = Nothing
End Sub

'Private Sub Text1_LostFocus(Index As Integer)
'Dim Cad As String
'If Index = 8 Then
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
'End If
'End Sub
