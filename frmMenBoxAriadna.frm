VERSION 5.00
Begin VB.Form frmMenBoxAriadna 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ariconta6"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenBoxAriadna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOtro 
      Caption         =   "otro dao"
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   2310
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdOtro 
      Caption         =   "otro dao"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   2310
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2310
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Si"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2310
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   4
      Left            =   120
      Picture         =   "frmMenBoxAriadna.frx":000C
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      Top             =   2160
      Width           =   7695
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblPie 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label lblTexto 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   6435
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   3
      Left            =   120
      Picture         =   "frmMenBoxAriadna.frx":0F97
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   2
      Left            =   120
      Picture         =   "frmMenBoxAriadna.frx":1973
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   120
      Picture         =   "frmMenBoxAriadna.frx":21AE
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "frmMenBoxAriadna.frx":2AA2
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmMenBoxAriadna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cmdCancelvisible As Boolean
Public cmdNovisible As Boolean
Public cmdOtro1visible As Boolean
Public cmdOtro2visible As Boolean
Public cmdYesvisible As Boolean
Public cmdOkvisible As Boolean

Dim Cursor As Integer


Public Sub Limpiar()

    cmdCancel.visible = False
    cmdNo.visible = False
    cmdOtro(1).visible = False
    cmdOtro(2).visible = False
    cmdYes.visible = False
    cmdOk.visible = False
    cmdCancelvisible = False
    cmdNovisible = False
    cmdOtro1visible = False
    cmdOtro2visible = False
    cmdYesvisible = False
    cmdOkvisible = False
    
    
    
    
    Image1(0).visible = False
    Image1(1).visible = False
    Image1(2).visible = False
    Image1(3).visible = False
    Image1(4).visible = False
    
    lblTexto.Caption = ""
    Me.lblPie.Caption = ""
    Me.lblTitulo.Caption = ""
    Me.Caption = "Ariconta6"
End Sub

Public Sub AjustaTamañosYPosicion()
Dim L As Integer
Dim W As Integer
Dim CambiarAncho As Boolean
Dim H As Integer
Dim SaltosDeLinea As Integer
    'Textos
    
        SaltosDeLinea = 0
        L = 1
        Do
            L = InStr(L, lblTexto.Caption, vbCrLf)
            If L > 0 Then
                SaltosDeLinea = SaltosDeLinea + 1
                L = L + 1
            End If
        Loop Until L = 0
        L = SaltosDeLinea * 20
        L = Len(lblTexto.Caption) + L
        W = 7680
        If L < 25 Then
            W = 4680
            H = 1900
        ElseIf L < 50 Then
            W = 5980
            H = 2000
        ElseIf L < 70 Then
            H = 2100
        ElseIf L < 100 Then
            H = 2300
        ElseIf L < 200 Then
            H = 2500
        ElseIf L < 300 Then
            H = 2900
        ElseIf L < 400 Then
            H = 3200
        ElseIf L < 500 Then
            H = 4200
        Else
            H = 4500
        End If
        If lblPie.Caption <> "" Then H = H + 300
        Me.Height = H
        Me.Width = W
        H = H - Me.Shape1.Height - 360
        Me.Shape1.top = H
        
        
        If lblPie.Caption <> "" Then
            H = Me.Shape1.top - 90
            lblPie.top = H - lblPie.Height
        End If
        
        If lblTitulo.Caption <> "" Then
            lblTexto.top = 480
        Else
            lblTexto.top = 210
        End If
        Me.lblTexto.Height = H - 30
    
    
    
    
    
    
    
    
    
    
    
    
    'Botonera
''''cmdCancelvisible
''''cmdOkvisible
''''cmdNovisible
''''cmdYesvisible
''''cmdOtro1visible
''''cmdOtro2visible


    W = Me.Width - 240
    
    'DE derecha a izquierda
    If cmdCancelvisible Then
        W = W - cmdCancel.Width - 30
        Me.cmdCancel.Left = W
        Me.cmdCancel.top = Shape1.top + 120
        cmdCancel.visible = True
        cmdCancel.Cancel = True
    End If
    
    If cmdOkvisible Then
        W = W - cmdOk.Width - 60
        Me.cmdOk.Left = W
        Me.cmdOk.top = Shape1.top + 120
        cmdOk.visible = True
        cmdOk.Cancel = True
    End If
    
    If cmdNovisible Then
        W = W - cmdNo.Width - 60
        Me.cmdNo.Left = W
        Me.cmdNo.top = Shape1.top + 120
        cmdNo.visible = True
    End If
    
    If cmdYesvisible Then
        W = W - cmdYes.Width - 60
        Me.cmdYes.Left = W
        Me.cmdYes.top = Shape1.top + 120
        cmdYes.visible = True
    End If
    ''''cmdOtro1visible
''''cmdOtro2visible
    If cmdOtro1visible Then
        W = W - cmdOtro(1).Width - 60
        Me.cmdOtro(1).Left = W
        Me.cmdOtro(1).top = Shape1.top + 120
        cmdOtro(1).visible = True
        If Not cmdCancelvisible Then cmdOtro(1).Cancel = True
    End If
    
    If cmdOtro2visible Then
        W = W - cmdOtro(2).Width - 60
        Me.cmdOtro(2).Left = W
        Me.cmdOtro(2).top = Shape1.top + 120
        cmdOtro(2).visible = True
    End If
    
    
End Sub

Private Sub cmdCancel_Click()
    RespuestaMsgBox = vbCancel
    Unload Me
End Sub

Private Sub cmdNo_Click()
    RespuestaMsgBox = vbNo
    Unload Me
End Sub

Private Sub cmdOk_Click()
    RespuestaMsgBox = vbOK
    Unload Me
End Sub

Private Sub cmdOtro_Click(Index As Integer)
    RespuestaMsgBox = cmdOtro(Index).Tag
    Unload Me
End Sub

Private Sub cmdYes_Click()
    RespuestaMsgBox = vbYes
    Unload Me
End Sub

Private Sub Form_Activate()

    
    If Me.cmdYes.Default Then PonerFocoBtn cmdYes
    If Me.cmdNo.Default Then PonerFocoBtn cmdNo
    If Me.cmdCancel.Default Then PonerFocoBtn cmdCancel
    If Me.cmdOk.Default Then PonerFocoBtn cmdOk
    If Me.cmdOtro(1).Default Then PonerFocoBtn cmdOtro(1)
    If Me.cmdOtro(2).Default Then PonerFocoBtn cmdOtro(2)

    

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Cursor = Screen.MousePointer
    Me.Icon = frmppal.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = Cursor
End Sub
