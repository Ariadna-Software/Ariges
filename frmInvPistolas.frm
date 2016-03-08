VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInvPistolas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura del fichero de inventario."
   ClientHeight    =   1800
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdBuscaFichero 
      Caption         =   "...."
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtFich 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin MSComDlg.CommonDialog cmdgFichero 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Fichero inventario"
      Filter          =   "*.txt"
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero de inventario:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmInvPistolas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- frmInvPistolas:
'   La intención de este formulario es simplemente obtener el fichero con el path completo donde reside la información de
'   captura del inventario de las pistolas.
'   El formulario llamante (frmAlmInventario) será el que procesando esta información actualice las tablas sinven y salmac con los
'   datos adecuados
Option Explicit
Public Event Seleccionado(Cadena As String)
Dim Fichero As String



Private Sub cmdAceptar_Click()
    If Not ComprobarFichero() Then
        MsgBox "El fichero es incorrecto o no existe", vbExclamation
        Exit Sub
    End If
    RaiseEvent Seleccionado(Fichero)
    Unload Me
End Sub

Private Sub cmdBuscaFichero_Click()
    cmdgFichero.Filter = "Texto *.txt"
    cmdgFichero.ShowOpen
    txtFich = cmdgFichero.FileName
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function ComprobarFichero() As Boolean
    '-- Comprueba que le fichero seleccionado realmente existe
    Fichero = txtFich
    If Dir(Fichero) <> "" Or Fichero <> "" Then ComprobarFichero = True
End Function

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
End Sub
