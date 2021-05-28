VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.Controls.v17.2.0.ocx"
Begin VB.Form frmMenBox 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TaskDialog TaskDialog1 
      Left            =   0
      Top             =   0
      _Version        =   1114114
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
      WindowTitle     =   "TaskDialog1"
   End
End
Attribute VB_Name = "frmMenBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub Form_Activate()
Dim Cad As String
    Cad = TaskDialog1.MainInstructionText
   
    frmMenBox.TaskDialog1.ShowDialog
    Unload Me
End Sub

Private Sub TaskDialog1_ButtonClicked(ByVal Id As Long, CloseDialog As Variant)

    If Id = 2 Then
        RespuestaMsgBox = vbYes
    ElseIf Id = 4 Then
        RespuestaMsgBox = vbNo
    ElseIf Id = 8 Then
        RespuestaMsgBox = vbCancel
    Else
        RespuestaMsgBox = Id
    End If
End Sub
