VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPpalGessocial 
   Caption         =   "Gestion avanzada socios. Gesocial"
   ClientHeight    =   4455
   ClientLeft      =   225
   ClientTop       =   615
   ClientWidth     =   12510
   Icon            =   "frmGessocial.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGessocial.frx":33E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGessocial.frx":36FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGessocial.frx":9F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGessocial.frx":D350
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Asociados"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Unidades de negocio"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar empresa"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   120
      Top             =   120
      Width           =   7815
   End
   Begin VB.Menu mnGeneral 
      Caption         =   "General"
      Begin VB.Menu mnGeneral1 
         Caption         =   "Asociados"
         Index           =   0
      End
      Begin VB.Menu mnGeneral1 
         Caption         =   "Unidades de negocio"
         Index           =   1
      End
      Begin VB.Menu mnGeneral1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnGeneral1 
         Caption         =   "Cambiar empresa"
         Index           =   3
      End
      Begin VB.Menu mnGeneral1 
         Caption         =   "Salir"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmPpalGessocial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PulsaSalir As Boolean
Dim QuiereCambiarEmpresa As Boolean




Private Sub AccionesCerrar()
    On Error Resume Next
    
   ' Unload frmPpal
    
    'cerrar las conexiones
    conn.Close
    CerrarConexionConta

End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    QuiereCambiarEmpresa = False
    PulsaSalir = False
      With Me.Toolbar1
        .ImageList = Me.ImageList1
        .Buttons(1).Image = 2
        .Buttons(2).Image = 3
        
        .Buttons(10).Image = 4
        .Buttons(11).Image = 1
    End With
    CargaImagen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not PulsaSalir Then Cancel = 1
    If QuiereCambiarEmpresa Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = "S"
        AccionesCerrar
              
        'Finalizamos
        End
    End If
End Sub

Private Sub mnGeneral1_Click(Index As Integer)



    Select Case Index
    Case 0
            frmGesSocAsociadosGR.Show vbModal
    Case 1
            frmGesUdsNegocio.Show vbModal
    Case 3
            PulsaSalir = True
            QuiereCambiarEmpresa = True
            Unload Me
    Case 4
            'CERRAR ARIGES
            PulsaSalir = True
            AccionesCerrar
                  
            'Finalizamos
            End
    
    End Select
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Top = 240
    Me.Left = 240
    Image1.Top = 0
    Image1.Left = 0
    Image1.Width = Me.Width
    Image1.Height = Me.Height
    Image1.Stretch = True
    
    Image1.Picture = LoadPicture(App.Path & "\arifon2.dll")
    If Err.Number <> 0 Then
        Me.Picture = LoadPicture()
        Err.Clear
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1, 2
        mnGeneral1_Click Button.Index - 1
        
    Case 10, 11
        mnGeneral1_Click Button.Index - 7
    
    
    
    End Select
    
End Sub
