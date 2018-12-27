VERSION 5.00
Begin VB.Form frmFacClienteObser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmFacClienteObser.frx":0000
      Top             =   120
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "frmFacClienteObser"
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
    If Text1.Tag = 1 Then
        Text1.Tag = 0
        Me.Text1.SelStart = Len(Text1.Text)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Caption = "Observaciones"
    Me.Icon = frmPpal.Icon
    Text1.Locked = Not Modificar
    Me.Command1(0).Enabled = Modificar
    Text1.Tag = 1
    
End Sub
