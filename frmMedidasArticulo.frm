VERSION 5.00
Begin VB.Form frmMedidasArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste lineal"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
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
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
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
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "PRECIO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Precio x 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Medida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmMedidasArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Valores As String


Public Event DatoSeleccionado(CadenaSeleccion As String)

Dim C As String
Dim CPrecioFact As CPreciosFact
Dim OrigP As String

Private Sub cmdAceptar_Click()
    C = ""
    If Text1(2).Text = "" Then C = "N"
    If Text1(1).Text = "" Then C = "N"
    If Text1(3).Text = "" Then C = "N"
    If C <> "" Then
        MsgBox "Campos obligatorios", vbExclamation
        PonerFoco Text1(2)
        Exit Sub
    End If
    If OrigP = "" Then
        MsgBox "Error obteniendo precio articulo", vbExclamation
    End If
    C = Text1(2).Text & "|" & Text1(3).Text & "|" & OrigP & "|" & CPrecioFact.Descuento1 & "|" & CPrecioFact.Descuento2 & "|"
    RaiseEvent DatoSeleccionado(C)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "N"
        PonerDatos

       
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Tag = ""
    Me.Icon = frmPpal.Icon
    
    Screen.MousePointer = vbHourglass
    Text1(0).Tag = RecuperaValor(Valores, 1)
    Text1(3).Text = RecuperaValor(Valores, 2)
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(0).Text = ""
    Caption = "n"
    BloquearTxt Text1(1), True
    BloquearTxt Text1(2), True

    
End Sub


Private Sub PonerDatos()


    Set CPrecioFact = New CPreciosFact
    

    CPrecioFact.CodigoArtic = Text1(0).Tag
    CPrecioFact.CodigoClien = RecuperaValor(Valores, 3)
    CPrecioFact.FijarTarifaActividad
    
    Text1(1).Tag = 0
    C = CPrecioFact.ObtenerPrecio(False, RecuperaValor(Valores, 4), OrigP, "")
    If C <> "" Then Text1(1).Tag = C
    Text1(1).Text = Format(Text1(1).Tag, FormatoPrecio)
    

    C = ""
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set CPrecioFact = Nothing
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim C As Currency
    If Me.Tag = "" Then Exit Sub
    If Index = 0 Then
        If Not PonerFormatoDecimal(Text1(0), 2) Then
            Text1(0).Text = ""
            Text1(2).Text = ""
        Else
            
             C = ImporteFormateado(Text1(0).Text)
             C = C * Text1(1).Tag
             Text1(2).Text = Format(C, FormatoPrecio)
            
        End If
    End If
End Sub
