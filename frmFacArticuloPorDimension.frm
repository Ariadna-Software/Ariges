VERSION 5.00
Begin VB.Form frmFacArticuloPorDimension 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
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
   ScaleHeight     =   1080
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   8640
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   1545
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   240
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "C.Postal|T|N|||sclien|codpobla||N|"
      Text            =   "Text1"
      Top             =   360
      Width           =   5505
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7320
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmFacArticuloPorDimension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim Cerrar As Boolean
    KEYpressGnral KeyAscii, 4, Cerrar
    If Cerrar Then
        CadenaDesdeOtroForm = ""
        Unload Me
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    
    If Not PonerFormatoDecimal(Text1(Index), 2) Then
        Text1(Index).Text = ""
    Else
        'OK
        If Text1(1).Text = "" Then
        
            PonerFoco Text1(1)
        Else
            If Text1(2).Text = "" Then
                PonerFoco Text1(2)
            Else
                If Text1(3).Text = "" Then
                    PonerFoco Text1(3)
                Else
                    'OK. Devolvemos este dato
                    CadenaDesdeOtroForm = Text1(1) & "|" & Text1(2) & "|" & Text1(3) & "|"
                    Unload Me
                End If
            End If
        End If
    End If
End Sub
