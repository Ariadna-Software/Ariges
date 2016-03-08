VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmEulerPDF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver PDF"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   9600
      Width           =   1215
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _cx             =   18018
      _cy             =   16325
   End
End
Attribute VB_Name = "frmEulerPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.AcroPDF1.LoadFile Me.Tag
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon

End Sub
