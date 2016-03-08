VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVisReportExportar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar documento"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   495
      Index           =   1
      Left            =   6480
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar"
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmVisReportExportar.frx":0000
      Left            =   240
      List            =   "frmVisReportExportar.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   7455
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Left            =   840
      Picture         =   "frmVisReportExportar.frx":0043
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   525
   End
End
Attribute VB_Name = "frmVisReportExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Sub Combo1_Click()
    Text1.Text = ""
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        If Text1.Text = "" Then Exit Sub
        
        If Dir(Text1.Text, vbArchive) <> "" Then
            If MsgBox("El fichero ya existe. ¿Sobreescribir? ", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
        
        RaiseEvent DatoSeleccionado(Combo1.ListIndex & " " & Text1.Text)

    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    Combo1.ListIndex = 0
End Sub

Private Sub imgBuscarOfer_Click()


    cd1.CancelError = False
    If Combo1.ListIndex = 0 Then
        cd1.DefaultExt = ".pdf" 'extension por defecto
        cd1.Filter = "Acrobat dcoument |*.pdf|" 'extensiones a mostrar
    ElseIf Combo1.ListIndex = 1 Then
        cd1.DefaultExt = ".xls" 'extension por defecto
        cd1.Filter = "Hoja cálculo |*.xls|" 'extensiones a mostrar
    Else
        cd1.DefaultExt = ".doc" 'extension por defecto
        cd1.Filter = "Microsoft Word |*.doc|" 'extensiones a mostrar
    End If
    
    
    
    cd1.FilterIndex = 1
    cd1.FileName = ""
    
    Me.cd1.ShowSave
    
eDialog:
    If Err.Number <> 0 Then Err.Clear
    If cd1.FileName <> "" Then Text1.Text = cd1.FileName
End Sub
