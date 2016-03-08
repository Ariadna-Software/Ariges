VERSION 5.00
Object = "{608009F3-E1FB-11D2-9BA1-0040D0002C80}#1.0#0"; "nslock15vb6.ocx"
Begin VB.Form frmLLave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de registro"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmLLave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin nslock15vb6.ActiveLock ActiveLock1 
      Left            =   5040
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   820
      Password        =   "D@BYZ"
      SoftwareName    =   "AriGes 4"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
   End
   Begin VB.CommandButton cmdCont 
      Caption         =   "Continuar sin registro."
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "Registrar el producto."
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox txt3 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   12
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox txt3 
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txt3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblInf 
      Height          =   1095
      Left            =   2520
      TabIndex        =   15
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lbl3 
      Caption         =   "Clave de activación:"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   11
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label lbl3 
      Caption         =   "Número de producto:"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   9
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lbl3 
      Caption         =   "Nombre de producto:"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "ariadnasoftware@ariadnasoftware.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   6
      Top             =   5760
      Width           =   3075
   End
   Begin VB.Label Label2 
      Caption         =   "Tel: 96 3580547"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   5
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "C/ Franco Tormo N.3 Bajo Izq. 46007 Valencia"
      Height          =   495
      Index           =   3
      Left            =   2640
      TabIndex        =   4
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Ariadna Software S.L."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmLLave.frx":0CCA
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmLLave.frx":0E0D
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmLLave.frx":0F36
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   720
      Picture         =   "frmLLave.frx":1378
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmLLave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PulsadoSalir As Boolean

Private Sub cmdCont_Click()
Dim m
    Screen.MousePointer = vbDefault
    PulsadoSalir = True
    If (Not ActiveLock1.RegisteredUser) And ActiveLock1.UsedDays > 30 Then
        m = " A T E N C I O N " & vbCrLf
        m = m & "___________________________________________" & vbCrLf & vbCrLf
        m = m & "Han expirado el tiempo de uso sin registrarse." & vbCrLf
        m = m & "Póngase en contacto con nosotros." & vbCrLf & vbCrLf
        m = m & "       La aplicación se detendrá." & vbCrLf & vbCrLf
        m = m & "Ariadna Software." & vbCrLf
        m = m & "C/ Franco Tormo N.3 Bajo Izq. " & vbCrLf
        m = m & "46007 Valencia" & vbCrLf & vbCrLf
        m = m & "Tel:    96 358 05 47 " & vbCrLf
        m = m & "e-mail: ariadnasoftware@ariadnasoftware.com" & vbCrLf
        m = m & "___________________________________________" & vbCrLf & vbCrLf
        MsgBox m, vbCritical
        End
        Else
            Unload Me
    End If
    '-- Aqui continuaria en eavluación
End Sub

Private Sub cmdReg_Click()
    If Trim(txt3(2)) = "" Then Exit Sub
    ActiveLock1.LiberationKey = txt3(2)
    If ActiveLock1.RegisteredUser Then
        MsgBox "Se ha registrado con éxito." & vbCrLf & "Deberá reiniciar la aplicación.", vbExclamation
        End
    Else
        MsgBox "Clave de activación incorrecta.", vbExclamation
    End If
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    PulsadoSalir = False
    txt3(0) = ActiveLock1.SoftwareName
    txt3(1) = ActiveLock1.SoftwareCode
    '-- Cuidadin con la password que hay que ponerla de manera oculta

    If ActiveLock1.UsedDays <= 30 Then
       lblInf.Caption = "Usted no es un usuario registrado, lleva " & _
       CStr(ActiveLock1.UsedDays) & " dias usando el producto. Recuerde que le quedan " & _
       CStr(30 - ActiveLock1.UsedDays) & " dias de uso sin registro."
    Else
       lblInf.Caption = "El plazo de uso de este producto ha expirado" & _
           ". Es necesario registrarse para poder utilizar el programa"
    End If
    If ActiveLock1.LastRunDate >= Now Then
        MsgBox "La fecha del sistema ha sido modificada." & vbCrLf & _
            "Por seguridad el programa finalizará.", vbCritical
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not PulsadoSalir Then
   If Not Me.ActiveLock1.RegisteredUser Then End
End If
Screen.MousePointer = vbHourglass
End Sub


Private Sub txt3_DblClick(Index As Integer)
txt3(1).Enabled = True
End Sub
