VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlmCatalogos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17100
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   17100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4560
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   6600
      Width           =   1155
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   9340
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label LabelDoc 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   6825
   End
End
Attribute VB_Name = "frmAlmCatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public desdeArticulos As Boolean  'Si no, clientes
Public Codigo As String

Dim Primeravez As Boolean
Dim Cambios As Boolean

Private Sub cmdAceptar_Click()
    If Cambios Then
        NumRegElim = Val(MsgBox("Desea guardar los cambios?", vbYesNoCancel + vbQuestion))
        If NumRegElim = vbCancel Then Exit Sub
        
        If NumRegElim = vbYes Then ACtualizaEnBD
            
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Primeravez Then
        Primeravez = False
        LabelDoc.Caption = "Leyendo BD..."
        CargaDatos
    End If
    LabelDoc.Caption = "Catalogos disponibles"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Primeravez = True
    Cambios = False
    Me.Icon = frmPpal.Icon
    Caption = "Catálogos"
    CargaColumnas
End Sub


Private Sub CargaColumnas()
Dim C As ColumnHeader

    
    lw1.ColumnHeaders.Clear
    If desdeArticulos Then
        'MOVIMIENTOS
        
        For NumRegElim = 1 To 2
            Set C = lw1.ColumnHeaders.Add()
            C.Text = RecuperaValor("Agrupacion|Titulo|", CInt(NumRegElim))
            C.Width = RecuperaValor("1500|5600|", CInt(NumRegElim))
            C.Alignment = Val(RecuperaValor("0|0|", CInt(NumRegElim)))
            
        Next NumRegElim
        Me.Height = 9000
        Me.Width = 8700
    Else
        For NumRegElim = 1 To 3
            Set C = lw1.ColumnHeaders.Add()
            C.Text = RecuperaValor("Agrupacion|Titulo|Dto|", CInt(NumRegElim))
            C.Width = RecuperaValor("1500|5600|1000|", CInt(NumRegElim))
            C.Alignment = Val(RecuperaValor("0|0|1|", CInt(NumRegElim)))
            
        Next NumRegElim
        Me.Height = 12000
        Me.Width = 9200
    End If
    
    Me.cmdCancelar.Top = Me.Height - Me.cmdCancelar.Height - 540
    Me.cmdAceptar.Top = cmdCancelar.Top
    Me.cmdCancelar.Left = Me.Width - Me.cmdCancelar.Width - 450
    Me.cmdAceptar.Left = cmdCancelar.Left - Me.cmdCancelar.Width - 120
    lw1.Height = cmdCancelar.Top - 240 - lw1.Top
    lw1.Width = Me.Width - lw1.Left - 240
    
End Sub

Private Sub lw1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Cambios = True
    
End Sub

Private Sub CargaDatos()
Dim YaEnCatalogos As String
Dim SQL As String
Dim It
    Set miRsAux = New ADODB.Recordset
    YaEnCatalogos = "|"
    If desdeArticulos Then
        SQL = "select codagrupa from sagrupaart where codartic= " & DBSet(Codigo, "T")
    Else
        SQL = "select codagrupa from sagrupacli where codclien= " & Codigo
    End If
    
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        YaEnCatalogos = YaEnCatalogos & miRsAux!codagrupa & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    SQL = "SELECT sagrupa.codagrupa,descagrupa,dto1,tipo FROM sagrupa  where tipo=" & DBSet(IIf(desdeArticulos, "T", "C"), "T") & " ORDER BY 1"
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set It = lw1.ListItems.Add()
        It.Text = miRsAux!codagrupa
        It.SubItems(1) = miRsAux!descagrupa
        
        If Not desdeArticulos Then It.SubItems(2) = Format(DBLet(miRsAux!Dto1, "T"), FormatoImporte)
   
        SQL = "|" & miRsAux!codagrupa & "|"
        If InStr(1, YaEnCatalogos, SQL) > 0 Then It.Checked = True
 
        
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
End Sub


Private Sub ACtualizaEnBD()
Dim SQL As String
    If desdeArticulos Then
        SQL = "DELETE FROM sagrupaart where codartic= " & DBSet(Codigo, "T")
    Else
        SQL = "DELETE FROM sagrupacli where codclien= " & Codigo
    End If
    ejecutar SQL, False
    
    SQL = ""
    For NumRegElim = 1 To lw1.ListItems.Count
        If lw1.ListItems(NumRegElim).Checked Then SQL = SQL & ", (" & DBSet(Codigo, IIf(desdeArticulos, "T", "N")) & "," & DBSet(lw1.ListItems(NumRegElim).Text, "T") & ")"
    Next
    If SQL <> "" Then
        SQL = Mid(SQL, 2)
        SQL = " VALUES " & SQL
        If desdeArticulos Then
            SQL = "INSERT INTO sagrupaart (codartic,codagrupa) " & SQL
        Else
            SQL = "INSERT INTO sagrupacli (codclien,codagrupa) " & SQL
        End If
        ejecutar SQL, False
    End If
    
End Sub
