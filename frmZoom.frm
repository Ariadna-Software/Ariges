VERSION 5.00
Begin VB.Form frmZoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zoom"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "frmZoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   5450
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   5450
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   4845
      Index           =   0
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   440
      Width           =   9615
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   195
      Width           =   7455
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Actualizar(vCampo As String)


'variables que se pasan con valor al llamar al formulario de zoom desde otro formulario
Public pValor As String
Public pTitulo As String
Public pModo As Byte 'muestra el campo bloqueado y no se puede modificar segun Modo




Private Sub cmdActualizar_Click()
    'devolver el valor del campo al formulario que lo llamo
    RaiseEvent Actualizar(Text1(0).Text)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    'salimos y no hacemos nada
    Unload Me
End Sub

Private Sub Form_Load()

    'obtener el campo correspondiente y mostrarlo en el text
    Text1(0).Text = pValor
    Label1.Caption = pTitulo
    BloquearTxt Text1(0), (pModo <> 3 And pModo <> 4)
    Me.cmdActualizar.visible = (pModo = 3 Or pModo = 4)
    
    SendKeys "^{END}"
End Sub
