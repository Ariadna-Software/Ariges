VERSION 5.00
Begin VB.Form frmImprimir2 
   Caption         =   "Imprimir"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3270
   Icon            =   "frmImprimir2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1980
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   " Intervalo de registros "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton Option1 
         Caption         =   "Registros buscados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Registro actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmImprimir2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: CÈSAR +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

'GENERALS PER A PASAR-LI A CRYSTAL REPORTS
Public cadRegSelec As String
Public cadRegActua As String
Public cadTodosReg As String
Public OtrosParametros2 As String
Public NumeroParametros2 As Integer
Public cadTabla2 As String 'Cadena en les taules
Public Informe2 As String 'Nom de l'informe
Public MostrarTree2 As Boolean
Public InfConta2 As Boolean
Public ConSubInforme2 As Boolean

Private MIPATH As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

'Private Sub InicializarVbles()
'    cadFormula2 = ""
'    cadSelect2 = ""
'    cadParam2 = ""
'    numParam2 = 0
'End Sub

Private Sub cmdImprimir_Click()
    With frmVisReport
               
        If Option1(0).Value = True Then 'registro actual
            .FormulaSeleccion = cadRegActua
        ElseIf Option1(1).Value = True Then 'todos
            .FormulaSeleccion = cadTodosReg
        ElseIf Option1(2).Value = True Then 'registros seleccionados
            .FormulaSeleccion = cadRegSelec
        End If
        
        .OtrosParametros = OtrosParametros2
        .NumeroParametros = NumeroParametros2
        .Informe = MIPATH & Informe2
        .SoloImprimir = False
        .MostrarTree = MostrarTree2
        '.InfConta = InfConta2
        .ConSubInforme = ConSubInforme2
        .ExportarPDF = False
        .Opcion = 2000
        .Show vbModal
    
    End With
        
    Unload Me
        
End Sub

Private Sub Form_Load()
    MIPATH = App.Path & "\Informes\"
End Sub
