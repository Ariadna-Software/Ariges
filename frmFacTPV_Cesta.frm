VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacTPV_Cesta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cesta"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCesta 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   9
      Top             =   5880
      Width           =   1185
   End
   Begin VB.CommandButton cmdCesta 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9000
      TabIndex        =   8
      Top             =   5880
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   9596
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   600
      Width           =   4605
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label label1 
      Caption         =   "Cesta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   7
      Top             =   240
      Width           =   900
   End
   Begin VB.Label label1 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   900
   End
   Begin VB.Label label1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "frmFacTPV_Cesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Sql As String


Private Sub cmdCesta_Click(Index As Integer)
    If Index = 0 Then
        Sql = "Insertar " & Me.ListView1.ListItems.Count & " articulos. " & vbCrLf & "¿Continuar?"
    Else
        Sql = "Cancelar el proceso"
    End If
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    If Index = 0 Then CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub Form_Activate()
    If Sql = "" Then
        Sql = "select cestas.*,cestas_lineas.*, nomartic from cestas inner join cestas_lineas on cestas.cestaId= cestas_lineas.cestaId "
        Sql = Sql & " inner join sartic on cestas_lineas.codartic=sartic.codartic where cestas.codclien=" & Text1(0).Text & " order by numlinea"
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        Sql = ""
        While Not miRsAux.EOF
            
            NumRegElim = NumRegElim + 1
            If Text1(1).Text = "" Then
                
                
                Sql = DevuelveDesdeBD(conAri, "login", "straba", "codtraba", miRsAux!CodUsu, "N")
                If Sql = "" Then Sql = "ERROR trb:" & miRsAux!CodUsu
                Text1(1).Text = Sql
                Sql = miRsAux!cestaId
                Text1(2).Text = Format(miRsAux!cestaId, "0000") & " " & Format(miRsAux!Fecha, "dd/mm/yy")
            Else
                If Val(Sql) <> miRsAux!cestaId Then
                    MsgBox "Dos cestas abiertas para el mismo cliente", vbExclamation
                      Sql = miRsAux!cestaId
                End If
            End If
            
            ListView1.ListItems.Add , "L" & Format(miRsAux!cestaLineaId, "000000")
            ListView1.ListItems(NumRegElim).Text = miRsAux!codArtic
            ListView1.ListItems(NumRegElim).SubItems(1) = miRsAux!NomArtic
            ListView1.ListItems(NumRegElim).SubItems(2) = Format(miRsAux!cantidad, FormatoCantidad)
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        

        
        
    End If
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    Me.Icon = frmPpal.Icon
    Sql = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    
    
    
End Sub
