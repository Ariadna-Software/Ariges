VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEuler 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameVerDescri 
      Height          =   3495
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   7815
      Begin VB.CommandButton cmdAceptarFich 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   15
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcion 
         BackColor       =   &H80000018&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   7575
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6360
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   885
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1440
         Width           =   7575
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame FrameAnyadirArchivo 
      Height          =   4335
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtDescripcion 
         Height          =   885
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2280
         Width           =   7575
      End
      Begin VB.CommandButton cmdAceptarFich 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   2
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6600
         TabIndex        =   3
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   7575
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   405
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1440
         Width           =   7575
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Guardar como"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Image imgDir 
         Height          =   240
         Index           =   0
         Left            =   1800
         Picture         =   "frmEuler.frx":0000
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero a importar"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As Byte
    ' 0.- Insertar documento en oferta
    ' 1.- Modificar / er observaciones
    
Private Sub cmdAceptarFich_Click(index As Integer)


    If opcion = 0 Then
        If Text1(0).Text = "" Then Exit Sub
        If Trim(txtDescripcion(0).Text) = "" Then
            MsgBox "Ponga descripcion fichero", vbExclamation
            Exit Sub
        End If
        '                       origen                 guardarcomo                  descripcion
        CadenaDesdeOtroForm = Text1(0).Text & "|" & txtDescripcion(0).Text & "|" & txtDescripcion(1).Text & "|"
    Else
        If Me.txtDescripcion(2).Text = "" Then
            MsgBox "Ponga la descripcion", vbExclamation
            Exit Sub
        End If
        CadenaDesdeOtroForm = txtDescripcion(2).Text
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click(index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    limpiar Me
    FrameAnyadirArchivo.visible = False
    FrameVerDescri.visible = False
    Select Case opcion
    Case 0
        Caption = "insertar fichero"
        PonerFrameVisible FrameAnyadirArchivo
    Case 1
        Caption = "Ver descripcion"
        txtDescripcion(2).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        txtDescripcion(3).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        PonerFrameVisible FrameVerDescri
        CadenaDesdeOtroForm = ""
    End Select
    Me.cmdCancelar(opcion).Cancel = True
End Sub

Private Sub PonerFrameVisible(ByRef Fr As Frame)
    Fr.visible = True
    Fr.Top = 30
    Fr.Left = 30
    Me.Width = Fr.Width + 180
    Me.Height = Fr.Height + 520
End Sub

Private Sub imgDir_Click(index As Integer)
     cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.CancelError = False
    If index = 0 Then
        'cd1.Filter = "Adobe PDF (*.pdf)|*.pdf|MS Office WORD (*.doc)|*.doc|MS Office WORD 2007|*.docx"
        cd1.Filter = "Adobe PDF (*.pdf)|*.pdf"
        cd1.FilterIndex = 0
    End If
    cd1.ShowOpen
    If cd1.FileName = "" Then Exit Sub
    If UCase(Right(cd1.FileName, 4)) <> ".PDF" Then
        MsgBox "Solo PDFs", vbExclamation
        Exit Sub
    End If
    
    
    Text1(index).Text = cd1.FileName
    
    PonerRestoCamposInsertarFichero index
End Sub

Private Sub Text1_OLEDragDrop(index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Dim V
    NumRegElim = 0
    For Each V In Data.Files
        Debug.Print V
        Text1(index).Text = V
        NumRegElim = NumRegElim + 1
        
    Next V
    If NumRegElim > 1 Then MsgBox "Solo se contempla un archivo", vbExclamation
        
    PonerRestoCamposInsertarFichero index
End Sub

Private Sub PonerRestoCamposInsertarFichero(index As Integer)
    If UCase(Right(Text1(index).Text, 4)) <> ".PDF" Then
        MsgBox "Solo PDFs", vbExclamation
        Text1(index).Text = ""
    End If
    NumRegElim = InStrRev(Text1(index).Text, "\")
    If NumRegElim > 0 Then
        txtDescripcion(index).Text = Mid(Text1(index).Text, NumRegElim + 1)
        txtDescripcion(index).Text = Mid(txtDescripcion(index).Text, 1, Len(txtDescripcion(index).Text) - 4)

    End If
End Sub

Private Sub txtDescripcion_LostFocus(index As Integer)
    If index = 0 Then
        If txtDescripcion(1).Text = "" Then txtDescripcion(1).Text = txtDescripcion(0).Text
    End If
End Sub
