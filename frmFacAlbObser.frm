VERSION 5.00
Begin VB.Form frmFacAlbObser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmFacAlbObser.frx":0000
      Top             =   120
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "frmFacAlbObser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Modificar As Boolean

Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = "0|"
    Else
        'Desde el 3 en adelante
        CadenaDesdeOtroForm = "1|" & Text1.Text
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.Tag = 1 Then
        Me.Tag = 0
        If Modificar Then
            If Text1.Text <> "" Then Text1.Text = Text1.Text & " "
            Text1.SelStart = Len(Text1.Text)
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Tag = 1
    Me.Caption = "Observaciones linea albarán"
    Me.Icon = frmPpal.Icon
    Text1.Locked = Not Modificar
    Me.Command1(0).Enabled = Modificar
    Screen.MousePointer = vbDefault
    
        
End Sub
