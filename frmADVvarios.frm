VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmADVvarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13275
   Icon            =   "frmADVvarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDH 
      Height          =   4575
      Left            =   3600
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdBusqueda 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   13
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdBusqueda 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3120
         TabIndex        =   12
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   3480
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3480
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   3000
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3000
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   2160
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2160
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1680
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1680
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   960
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   960
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   480
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Cod. "
         Text            =   "Text1"
         Top             =   480
         Width           =   765
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   5
         Left            =   960
         Picture         =   "frmADVvarios.frx":000C
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   4
         Left            =   960
         Picture         =   "frmADVvarios.frx":010E
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   3
         Left            =   960
         Picture         =   "frmADVvarios.frx":0210
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   2
         Left            =   960
         Picture         =   "frmADVvarios.frx":0312
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmADVvarios.frx":0414
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   960
         Width           =   240
      End
      Begin VB.Image imgBusc 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmADVvarios.frx":0516
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   28
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   27
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   26
         Top             =   3030
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   25
         Top             =   1710
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Variedad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame FrameAgrupadoCampos 
      Height          =   7695
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   13095
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   495
         Index           =   1
         Left            =   12000
         TabIndex        =   33
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdRegresar2 
         Caption         =   "Regresar"
         Height          =   495
         Left            =   10920
         TabIndex        =   31
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdBUs2 
         Height          =   375
         Left            =   1080
         Picture         =   "frmADVvarios.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Buscar"
         Top             =   7200
         Width           =   375
      End
      Begin MSComctlLib.ListView lwCamposAgrupados 
         Height          =   6255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NroCampo"
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Partida"
            Object.Width           =   4075
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Variedad"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Sup(ha)"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Socio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Agrupado por NºCampo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   4215
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   240
         Picture         =   "frmADVvarios.frx":101A
         ToolTipText     =   "Quitar al haber"
         Top             =   7200
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   600
         Picture         =   "frmADVvarios.frx":1164
         ToolTipText     =   "Puntear al haber"
         Top             =   7200
         Width           =   240
      End
   End
   Begin VB.Frame FrameSelecCampo 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   13095
      Begin VB.CommandButton cmdBusq 
         Height          =   375
         Left            =   1080
         Picture         =   "frmADVvarios.frx":12AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Buscar"
         Top             =   7200
         Width           =   375
      End
      Begin VB.CommandButton cmdSelCampo 
         Caption         =   "Regresar"
         Height          =   495
         Left            =   10920
         TabIndex        =   3
         Top             =   7080
         Width           =   975
      End
      Begin MSComctlLib.ListView lw11 
         Height          =   6615
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   11668
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Partida"
            Object.Width           =   4075
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Variedad"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Sup(ha)"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Socio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   495
         Index           =   0
         Left            =   12000
         TabIndex        =   1
         Top             =   7080
         Width           =   975
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmADVvarios.frx":1CB0
         ToolTipText     =   "Puntear al haber"
         Top             =   7200
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmADVvarios.frx":1DFA
         ToolTipText     =   "Quitar al haber"
         Top             =   7200
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmADVvarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'[ARIGES-ARITPV]
'   Revisar en programa ARITPV
'------------------------------------------------


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Public Opcion As Byte
    '0.- Mostrar campos para seleccionar en partes de trabajo
    '1.-   " todo igual pero con campos agrupados

Public vCampos As String   'Si es -1 es que quiere lanzar el modo bsqueda

Dim PrimVez As Boolean
Dim SQL As String
Dim IT As ListItem

Private Sub cmdBUs2_Click()
  Me.FrameDH.visible = True
    Me.FrameAgrupadoCampos.Enabled = False
    PonerFoco Text1(0)
End Sub

''''''
''''''Private Sub chkTodos_Click()
''''''    Screen.MousePointer = vbHourglass
''''''    CargaCampos
''''''    Screen.MousePointer = vbDefault
''''''End Sub

Private Sub cmdBusq_Click()
    
    Me.FrameDH.visible = True
    Me.FrameSelecCampo.Enabled = False
    PonerFoco Text1(0)
End Sub

Private Sub cmdBusqueda_Click(index As Integer)
     If index = 0 Then
        SQL = "      rcampos.codclien > 0"
            'rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
            'variedades on rcampos.codvarie = variedades.codvarie)"
            'SQL = SQL & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
        If Text1(0).Text <> "" Then SQL = SQL & " AND rcampos.codsocio >= " & Text1(0).Text
        If Text1(1).Text <> "" Then SQL = SQL & " AND rcampos.codsocio <= " & Text1(1).Text
        If Text1(2).Text <> "" Then SQL = SQL & " AND rcampos.codclien >= " & Text1(2).Text
        If Text1(3).Text <> "" Then SQL = SQL & " AND rcampos.codclien <= " & Text1(3).Text
        If Text1(4).Text <> "" Then SQL = SQL & " AND rcampos.codvarie >= " & Text1(4).Text
        If Text1(5).Text <> "" Then SQL = SQL & " AND rcampos.codvarie <= " & Text1(5).Text
        If SQL <> "" Then SQL = Mid(SQL, 5)
    Else
        SQL = "rcampos.codclien  = " & vCampos
    End If
    
    Screen.MousePointer = vbHourglass
    
    
    
    'Uno u otro
    If Opcion = 0 Then
        CargaCampos2 SQL
        Me.FrameSelecCampo.Enabled = True
    Else
        CargaCamposAgr SQL
        Me.FrameAgrupadoCampos.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    Me.FrameDH.visible = False
   
    
End Sub

Private Sub cmdCancelar_Click(index As Integer)

    CadenaDesdeOtroForm = ""  'por si las moscas
    Unload Me
End Sub

Private Sub cmdRegresar2_Click()
Dim T1 As String
     
    If lwCamposAgrupados.ListItems.Count = 0 Then Exit Sub
    
    SQL = ""
    For NumRegElim = 1 To lwCamposAgrupados.ListItems.Count
        If lwCamposAgrupados.ListItems(NumRegElim).Checked Then SQL = SQL & "1"
    Next
    If SQL = "" Then
        MsgBox "Seleccione algun campo", vbExclamation
        Exit Sub
    End If
    
    
    If vCampos = "-1" Then
        'Multi parte. NO puedo coger campos que no tengan cliente asociado"
        T1 = ""
        For NumRegElim = 1 To lwCamposAgrupados.ListItems.Count
            If lwCamposAgrupados.ListItems(NumRegElim).Checked Then
               If Trim(lwCamposAgrupados.ListItems(NumRegElim).SubItems(4)) = "" Then T1 = T1 & "X"
            End If
        Next
        If T1 <> "" Then
            MsgBox "Existen " & Len(T1) & " campo" & IIf(Len(T1) > 1, "s", "") & " sin cliente asociado", vbExclamation
            Exit Sub
        End If
    End If
    
    CadenaDesdeOtroForm = ""
    NumRegElim = Len(SQL)
    If NumRegElim > 1 Then
        SQL = "Ha seleccionado " & NumRegElim & " AGRUPACIONES de campos. ¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        CadenaDesdeOtroForm = "@" 'comienza por arroba
    End If

    
    For NumRegElim = 1 To lwCamposAgrupados.ListItems.Count
        If lwCamposAgrupados.ListItems(NumRegElim).Checked Then
            SQL = lwCamposAgrupados.ListItems(NumRegElim).Text & "|" & lwCamposAgrupados.ListItems(NumRegElim).SubItems(1) & "|" & lwCamposAgrupados.ListItems(NumRegElim).SubItems(2) & "|" & lwCamposAgrupados.ListItems(NumRegElim).Tag & "|" & "·#"   'tag=codvarie
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
        End If
    Next
        
    
    Unload Me
End Sub

Private Sub cmdSelCampo_Click()
Dim T1 As String
    If lw11.ListItems.Count = 0 Then Exit Sub
    
    SQL = ""
    For NumRegElim = 1 To lw11.ListItems.Count
        If lw11.ListItems(NumRegElim).Checked Then SQL = SQL & "1"
    Next
    If SQL = "" Then
        MsgBox "Seleccione algun campo", vbExclamation
        Exit Sub
    End If
    
    
    If vCampos = "-1" Then
        'Multi parte. NO puedo coger campos que no tengan cliente asociado"
        T1 = ""
        For NumRegElim = 1 To lw11.ListItems.Count
            If lw11.ListItems(NumRegElim).Checked Then
               If Trim(lw11.ListItems(NumRegElim).SubItems(4)) = "" Then T1 = T1 & "X"
            End If
        Next
        If T1 <> "" Then
            MsgBox "Existen " & Len(T1) & " campo" & IIf(Len(T1) > 1, "s", "") & " sin cliente asociado", vbExclamation
            Exit Sub
        End If
    End If
    
    CadenaDesdeOtroForm = ""
    NumRegElim = Len(SQL)
    If NumRegElim > 1 Then
        SQL = "Ha seleccionado " & NumRegElim & " campos. ¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        CadenaDesdeOtroForm = "@" 'comienza por arroba
    End If

    
    For NumRegElim = 1 To lw11.ListItems.Count
        If lw11.ListItems(NumRegElim).Checked Then
            SQL = lw11.ListItems(NumRegElim).Text & "|" & lw11.ListItems(NumRegElim).SubItems(1) & "|" & lw11.ListItems(NumRegElim).SubItems(2) & "|" & "·#"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
        End If
    Next
        
    
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        
        If Opcion = 0 Then
            If vCampos >= 0 Then
                CargaCampos2 "rcampos.codclien  = " & vCampos   'Martin. Enlaza con codclien
            Else
                cmdBusq_Click
            End If
        
        Else
            cmdBUs2_Click
        
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    '
    Screen.MousePointer = vbHourglass
    PrimVez = True
    Me.Icon = frmPpal.Icon
    Me.FrameSelecCampo.visible = False
    limpiar Me
    Select Case Opcion
    Case 0
        Caption = "Campos"
        
        PonerFrameVisible Me.FrameSelecCampo
            
    Case 1
        Caption = "Campos agrupados"
        PonerFrameVisible Me.FrameAgrupadoCampos
        
    End Select
    
    Me.cmdCancelar(Opcion).Cancel = True

End Sub


Private Sub PonerFrameVisible(Fr As Frame)
    Fr.visible = True
    Fr.Top = 0
    Fr.Left = 120
    Me.Height = Fr.Height + 480
    Me.Width = Fr.Width + 320
End Sub



'----------------------------------
Private Sub CargaCampos2(ByVal SQ As String)

    On Error GoTo ecargaCampos
    Set miRsAux = New ADODB.Recordset
    
    Me.lw11.ListItems.Clear
    'Para no meter MUCHOS ariagro.tabla
    'Pongo @# y luego lo reemplazo por vparamaplic.Ariagro.
    SQL = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,rcampos.supsigpa"
    SQL = SQL & " from ((@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
    SQL = SQL & " inner join @#variedades on rcampos.codvarie = variedades.codvarie)"
    SQL = SQL & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
    'where socio
    If SQ <> "" Then SQL = SQL & " WHERE " & SQ
    
    SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw11.ListItems.Add()
        IT.Text = Format(miRsAux!codCampo, "0000000")
        IT.SubItems(1) = miRsAux!nomparti
        IT.SubItems(2) = miRsAux!nomvarie
        'Superficie
        
        IT.SubItems(3) = Format(DBLet(miRsAux!supsigpa, "N"), FormatoPrecio)
        
        If IsNull(miRsAux!codClien) Then
            IT.SubItems(4) = " "
        Else
            IT.SubItems(4) = Format(miRsAux!codClien, "00000")
        End If
        IT.SubItems(5) = Format(miRsAux!codsocio, "00000")
        IT.SubItems(6) = miRsAux!nomsocio
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
ecargaCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub





Private Sub CargaCamposAgr(ByVal SQ As String)

    On Error GoTo ecargaCampos
    Set miRsAux = New ADODB.Recordset
    
    Me.lwCamposAgrupados.ListItems.Clear
    'Para no meter MUCHOS ariagro.tabla
    'Pongo @# y luego lo reemplazo por vparamaplic.Ariagro.
    SQL = "select rcampos.nrocampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,"
    SQL = SQL & " sum(rcampos.supsigpa) supsigpa,nomclien,rcampos.codvarie"
    SQL = SQL & " from ((@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
    SQL = SQL & " inner join @#variedades on rcampos.codvarie = variedades.codvarie)"
    SQL = SQL & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
    SQL = SQL & " left join sclien on rcampos.codclien=sclien.codclien"
    'where socio
    If SQ <> "" Then SQL = SQL & " WHERE " & SQ
    
    SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
    
    SQL = SQL & " GROUP BY nrocampo,rcampos.codvarie"
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCamposAgrupados.ListItems.Add()
        IT.Text = Format(miRsAux!nrocampo, "0000000")
        IT.SubItems(1) = miRsAux!nomparti
        IT.SubItems(2) = miRsAux!nomvarie
        'Superficie
        
        IT.SubItems(3) = Format(DBLet(miRsAux!supsigpa, "N"), FormatoPrecio)
        
        If IsNull(miRsAux!codClien) Then
            IT.SubItems(4) = " "
        Else
            IT.SubItems(4) = Format(miRsAux!codClien, "00000")
        End If
        IT.SubItems(5) = Format(miRsAux!codsocio, "00000")
        If IsNull(miRsAux!NomClien) Then
            IT.SubItems(6) = miRsAux!nomsocio
        Else
            IT.SubItems(6) = miRsAux!NomClien
        End If
        IT.Tag = miRsAux!codvarie
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
ecargaCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub




Private Sub frmB_Selecionado(CadenaDevuelta As String)
    SQL = CadenaDevuelta
End Sub

Private Sub imgBusc_Click(index As Integer)
    
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    
    If index < 2 Then
        frmB.vCampos = "Codigo|" & vParamAplic.Ariagro & ".rsocios|codsocio|N|0000000|20·Nombre|" & vParamAplic.Ariagro & ".rsocios|nomsocio|T||70·"
        frmB.vTabla = vParamAplic.Ariagro & ".rsocios"
        frmB.vTitulo = "Socios"
    ElseIf index < 4 Then
        frmB.vCampos = "Codigo|sclien|codclien|N|0000000|20·Nombre|sclien|nomclien|T||70·"
        frmB.vTabla = "sclien"
        frmB.vTitulo = "Clientes"
    Else
        
        frmB.vCampos = "Codigo|" & vParamAplic.Ariagro & ".variedades|codvarie|N|0000000|20·Nombre|" & vParamAplic.Ariagro & ".variedades|nomvarie|T||70·"
        frmB.vTabla = vParamAplic.Ariagro & ".variedades"
        frmB.vTitulo = "Variedades"
    End If
    frmB.vSQL = ""
    
    '###A mano
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
    SQL = ""
    frmB.Show vbModal
 '   Dim i As Integer
 '   For i = 1 To Me.lw11.ColumnHeaders.Count
 '       Debug.Print lw11.ColumnHeaders(i).Text & " " & lw11.ColumnHeaders(i).Width
 '   Next i
    If SQL <> "" Then
        Text1(index).Text = RecuperaValor(SQL, 1)
        Text2(index).Text = RecuperaValor(SQL, 2)
        SQL = ""
        If index = 5 Then
            PonerFocoBtn Me.cmdBusqueda(0)
        Else
            PonerFoco Text1(index + 1)
        End If
    End If
End Sub

Private Sub imgCheck_Click(index As Integer)
    If index <= 1 Then
        For NumRegElim = 1 To lw11.ListItems.Count
           lw11.ListItems(NumRegElim).Checked = index = 1
        Next
    Else
            
        For NumRegElim = 1 To lwCamposAgrupados.ListItems.Count
             Me.lwCamposAgrupados.ListItems(NumRegElim).Checked = index = 2
        Next
    End If
End Sub

Private Sub lw11_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.index - 1 <> Me.lw11.SortKey Then
        lw11.SortKey = ColumnHeader.index - 1
        lw11.SortOrder = lvwAscending
    Else
        If lw11.SortOrder = lvwAscending Then
            lw11.SortOrder = lvwDescending
        Else
            lw11.SortOrder = lvwAscending
        End If
    End If
        
End Sub

Private Sub lw11_DblClick()
    cmdSelCampo_Click
End Sub

Private Sub lw11_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Set lw11.SelectedItem = Item
End Sub

Private Sub lwCamposAgrupados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.index - 1 <> Me.lwCamposAgrupados.SortKey Then
        lwCamposAgrupados.SortKey = ColumnHeader.index - 1
        lwCamposAgrupados.SortOrder = lvwAscending
    Else
        If lwCamposAgrupados.SortOrder = lvwAscending Then
            lwCamposAgrupados.SortOrder = lvwDescending
        Else
            lwCamposAgrupados.SortOrder = lvwAscending
        End If
    End If
End Sub

Private Sub Text1_GotFocus(index As Integer)
    ConseguirFoco Text1(index), 3
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(index As Integer)
    Text1(index).Text = Trim(Text1(index).Text)
    SQL = ""
    If Text1(index).Text <> "" Then
        If Not PonerFormatoEntero(Text1(index)) Then
            Text1(index).Text = ""
        
        Else
            If index < 2 Then
                'Socio
                SQL = DevuelveDesdeBD(conAri, "nomsocio", vParamAplic.Ariagro & ".rsocios", "codsocio", Text1(index).Text)
                If SQL = "" Then SQL = "NO existe el socio"
            ElseIf index < 4 Then
                SQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(index).Text)
                If SQL = "" Then SQL = "NO existe el cliente"
                    
            Else
                'Variedad
                SQL = DevuelveDesdeBD(conAri, "nomvarie", vParamAplic.Ariagro & ".variedades", "codvarie", Text1(index).Text)
                If SQL = "" Then SQL = "NO existe la variedad"
            End If
            
        End If
    End If
    Text2(index).Text = SQL
End Sub
