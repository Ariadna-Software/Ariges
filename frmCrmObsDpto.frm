VERSION 5.00
Begin VB.Form frmCrmObsDpto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Observaciones CRM por Departamento"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmCrmObsDpto.frx":0000
      Top             =   1080
      Width           =   7935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmCrmObsDpto.frx":0006
      Left            =   360
      List            =   "frmCrmObsDpto.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Departamento"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCrmObsDpto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Nuevo As Boolean

Public Dpto As Byte




Private Sub cmdAceptar_Click()
Dim C As String

    If Nuevo Then
        C = "insert into `scrmobsclien` (`codclien`,`dpto`,`fecha`,`observa`) values ("
        C = C & Me.Tag & "," & Combo1.ListIndex + 1 & "," & DBSet(Now, "F") & "," & DBSet(Text1.Text, "T") & ")"
        
    Else
        C = "UPDATE scrmobsclien set fecha = " & DBSet(Now, "F")
        C = C & " , observa=" & DBSet(Text1.Text, "T")
        C = C & " WHERE codclien = " & Me.Tag & " AND dpto=" & Dpto
    End If
    If Ejecutar(C, False) Then Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    If Nuevo Then
        Text1.Text = ""
        Me.Combo1.ListIndex = -1
    Else
        Me.Combo1.ListIndex = Dpto - 1
        Text1.Text = CadenaDesdeOtroForm
        CadenaDesdeOtroForm = ""
    End If
    Combo1.Enabled = Nuevo
End Sub
