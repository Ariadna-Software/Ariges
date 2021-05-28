VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmBackUP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmBackUP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl2.Animation Animation1 
      Height          =   915
      Left            =   600
      TabIndex        =   5
      Top             =   1620
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1614
      _Version        =   327681
      FullWidth       =   301
      FullHeight      =   61
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   1980
      TabIndex        =   3
      Top             =   3360
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   3600
      TabIndex        =   2
      Top             =   3360
      Width           =   1515
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   2595
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "sobre ficheros locales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   4860
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   4515
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Copia de seguridad :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   3480
   End
End
Attribute VB_Name = "frmBackUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

Private Tablas() As String
Private NumTablas As Integer

Dim RS As Recordset
Dim NF As Integer
Dim Archivo As String
Dim Izquierda As String
Dim Derecha As String


'En todas las futuros backups, se trata de cargar el array tablas con las tablas(-1) a copiar


Private Sub cmdAceptar_Click()
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de copia", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Select Case Combo1.ListIndex
    Case 0
        CopiaTodo
'    Case 1
'        RenumeracionAsientos
'    Case 2
'        Cierre
'    Case 3
'        Amorizacion
    End Select
    
    'Ahora hacemos las copias
    HacerBackUp
    
    MsgBox "Copia finalizada en: " & Archivo, vbInformation
    cmdAceptar.Enabled = False
    Label1.Caption = ""
    PonerVideo False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    
    Label1.Caption = ""
    Label2.Caption = "Empresa: " & vEmpresa.nomempre
    Caption = "Backup para " & UCase(vEmpresa.nomresum)
    
    CargaCombo
    Combo1.ListIndex = 0
End Sub


Private Sub CargaCombo()
    Combo1.Clear
    Combo1.AddItem "Copia todo "
'    Combo1.AddItem "Renumeración asientos"
'    Combo1.AddItem "Cierre ejercicio"
'    Combo1.AddItem "Amortización"
End Sub

Private Sub CopiaTodo()


    Set RS = New ADODB.Recordset
    RS.Open "SHOW TABLES", conn, adOpenKeyset, adLockOptimistic, adCmdText
    NumTablas = 0
    While Not RS.EOF
        If LCase(Mid(RS.Fields(0), 1, 3)) = "tmp" Then
            'Las temporales no hacemos nada
        Else
            NumTablas = NumTablas + 1
        End If
        RS.MoveNext
    Wend
    
    RS.MoveFirst
    
    ReDim Tablas(NumTablas - 1)
    NumTablas = 0
    While Not RS.EOF
        If LCase(Mid(RS.Fields(0), 1, 3)) = "tmp" Then
            'Las temporales no hacemos nada
        Else
            Tablas(NumTablas) = RS.Fields(0)
            NumTablas = NumTablas + 1
        End If
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing

End Sub




Private Sub PonerVideo(Encender As Boolean)
    If Encender Then
        Me.Animation1.Open App.Path & "\actua.avi"
        Me.Animation1.Play
        Me.Animation1.visible = True
    Else
        Me.Animation1.Stop
        Me.Animation1.visible = False
    End If
End Sub



Private Sub HacerBackUp()
Dim i As Integer

    On Error GoTo EBackUp

    If NumTablas > 3 Then PonerVideo True


    Archivo = FijarCarpeta
    If Archivo = "" Then
        MsgBox "no se ha creado correctamente la carpeta de copia.", vbExclamation
        Exit Sub
    End If
        
    
    For i = 0 To NumTablas - 1
        Label1.Caption = Tablas(i) & "     (" & i + 1 & " de " & NumTablas & ")"
        Label1.Refresh
        BKTablas (Tablas(i))
    Next i
    
    Exit Sub
    
EBackUp:
    If Err.Number <> 0 Then MuestraError Err.Number, "BackUp", Err.Description
End Sub



Private Function FijarCarpeta() As String
Dim FE As String
Dim i As Integer

    On Error GoTo EFijarCarpeta

    FijarCarpeta = ""
    If Dir(App.Path & "\BACKUP", vbDirectory) = "" Then MkDir App.Path & "\BACKUP"
    
    Derecha = App.Path & "\BACKUP\"
    Izquierda = Format(Now, "yymmdd")
    i = -1
    Do
        i = i + 1
        FE = Format(i, "00")
        FE = Derecha & Izquierda & FE
        If Dir(FE, vbDirectory) = "" Then
            'OK
            MkDir FE
            FijarCarpeta = FE
            i = 100
        End If
    Loop Until i > 99
    Exit Function
    
EFijarCarpeta:
    MuestraError Err.Number
End Function



Private Sub BKTablas(tabla As String)
Dim cad As String

    Set RS = New ADODB.Recordset
    RS.Open tabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    If RS.EOF Then
        'No hace falta hacer back up
    
    Else
        NF = FreeFile
        Open Archivo & "\" & tabla & ".sql" For Output As #NF
        BACKUP_TablaIzquierda RS, Izquierda
        While Not RS.EOF
            BACKUP_Tabla RS, Derecha
            cad = "INSERT INTO " & tabla & " " & Izquierda & " VALUES " & Derecha & ";"
            Print #NF, cad
            RS.MoveNext
        Wend
        Close #NF
    End If
    RS.Close
    Set RS = Nothing
End Sub



