VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTiposMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de mensaje"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   Icon            =   "frmTiposMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   3855
      Begin VB.CommandButton cmdIco 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   12
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdIco 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   11
         Top             =   3960
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   600
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   3625
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   3120
         Picture         =   "frmTiposMensajes.frx":000C
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Icono"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   495
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2520
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Color"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   8
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   5640
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   8916
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTiposMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cambios As Boolean

Dim It As ListItem

Private Sub cmdIco_Click(Index As Integer)
Dim miRsAux As ADODB.Recordset
Dim SQL As String

    If Index = 0 Then
        If Text1(0).Text = "" Then
            MsgBox "Descripcion no puede estar en blanco", vbExclamation
            Exit Sub
        End If
    
        If Frame1.Tag = 1 Then
            SQL = "update mailtipo set descripcion= '" & DevNombreSQL(Text1(0).Text)
            SQL = SQL & "', color = '" & Me.Shape1.FillColor & "',numico = " & ListView2.SelectedItem.Index - 1
            SQL = SQL & " WHERE tipo = " & Text1(1).Text
            
        Else
            On Error GoTo Error1
        
            SQL = "Select max(tipo) from mailtipo"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then Text1(1).Text = DBLet(miRsAux.Fields(0), "N") + 1
            miRsAux.Close
            Set miRsAux = Nothing
            
            SQL = "INSERT INTO mailtipo (tipo, Descripcion, color, numico) VALUES (" & Text1(1).Text
            SQL = SQL & ",'" & DevNombreSQL(Text1(0).Text) & "','" & Me.Shape1.FillColor & "',"
            SQL = SQL & ListView2.SelectedItem.Index - 1 & ")"
            
        End If
        Conn.Execute SQL
        Cambios = True
        CargaTipos
    End If
    Frame1.visible = False
    PonerModo
    
Error1:
    If Err.Number <> 0 Then MuestraError Err.Description, "Insertar tipo de mensaje.", Err.Description
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub PonerModo()

    If Frame1.visible Then
        cmdIco(1).Cancel = True
    Else
        Command1.Cancel = True
    End If
    
    Toolbar1.Enabled = Not Frame1.visible
    Command1.Enabled = Not Frame1.visible
End Sub

Private Sub Form_Load()
    'Icono del formulario principal
    Me.Icon = frmPpal.imgListComun.ListImages(20).Picture
    
    Cambios = False
          ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        
        .Buttons(6).Enabled = vUsu.Nivel < 4
        .Buttons(7).Enabled = vUsu.Nivel < 4
        .Buttons(8).Enabled = vUsu.Nivel < 4
        
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        
        .Buttons(14).visible = False
        .Buttons(15).visible = False
        .Buttons(16).visible = False
        .Buttons(17).visible = False

    End With
        
    Set ListView1.SmallIcons = frmPpal.ImageListMAIL
    Set ListView2.SmallIcons = frmPpal.ImageListMAIL
    'Set ListView2.Icons = Admin.ImageListMAIL


Dim I As Integer
    
    'Metemos el icono VACIO
    Set It = ListView2.ListItems.Add
    It.Text = "SIN ICO"
    For I = 1 To frmPpal.ImageListMAIL.ListImages.Count
        Set It = ListView2.ListItems.Add
        It.Text = "   nº: " & I
        It.SmallIcon = I
        'It.Icon = I
    Next I
    
    CargaTipos
    Frame1.visible = False
    Set ListView1.SelectedItem = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cambios Then PonerArrayTiposMensaje
End Sub

Private Sub Image1_Click()
On Error GoTo E1
    cd1.CancelError = True
    cd1.ShowColor
    Me.Shape1.FillColor = cd1.Color
E1:
    Err.Clear
End Sub



Private Sub ListView1_DblClick()
    HacerTool 7
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Caption = Button.Index
    HacerTool Button.Index
    
End Sub

Private Sub HacerTool(Indice As Integer)
    If Indice = 7 Or Indice = 8 Then
        If ListView1.SelectedItem Is Nothing Then Exit Sub
    End If
    
    Select Case Indice
        
      
    Case 6, 7
        
        limpiar Me
        If Indice = 7 Then
            Frame1.Tag = "1"
            Label2.Caption = "MODIFICAR"
            Text1(1).Text = Mid(ListView1.SelectedItem.Key, 2)
            Text1(0).Text = ListView1.SelectedItem.Text
            Me.Shape1.FillColor = ListView1.SelectedItem.ForeColor
            If ListView1.SelectedItem.SmallIcon = "" Then
                ListView2.SelectedItem = ListView2.ListItems(1)
            Else
                ListView2.SelectedItem = ListView2.ListItems(Val(ListView1.SelectedItem.SmallIcon))
                ListView2.SelectedItem.EnsureVisible
            End If
        Else
            Frame1.Tag = 0
            Label2.Caption = "NUEVO"
            ListView2.SelectedItem = ListView2.ListItems(1)
        End If
        Frame1.visible = True
        
        PonerModo
        Text1(0).SetFocus
        
    Case 8
        'Eliminar
        If vUsu.Codigo <> 0 Then
            MsgBox "Solo el administrador puede eliminar", vbExclamation
            Exit Sub
        End If
        
    Case 12
        Unload Me
    End Select
End Sub


Private Sub CargaTipos()
Dim SQL As String

    ListView1.ListItems.Clear
    SQL = "Select * from mailtipo order by tipo"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set It = ListView1.ListItems.Add(, "C" & miRsAux!Tipo)
        It.Text = miRsAux!Descripcion
        If miRsAux!Color <> 0 Then It.ForeColor = miRsAux!Color
        If miRsAux!numico <> 0 Then It.SmallIcon = Val(miRsAux!numico)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub
