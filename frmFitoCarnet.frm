VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFitoCarnet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carnet manipulador fitosanitarios"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
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
      Left            =   8520
      TabIndex        =   5
      Top             =   4560
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   9960
      TabIndex        =   4
      Top             =   4560
      Width           =   1155
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NIF"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   6210
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NºCarnet"
         Object.Width           =   2884
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "F. Caducidad"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tipo"
         Object.Width           =   2188
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
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
      Height          =   315
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   375
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   375
      Width           =   4845
   End
   Begin VB.Label label1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   2
      Top             =   375
      Width           =   900
   End
   Begin VB.Image imgBuscar 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   0
      Left            =   1215
      ToolTipText     =   "Buscar artículo"
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "frmFitoCarnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cliente As Long

Private WithEvents frmCli As frmBasico2
Attribute frmCli.VB_VarHelpID = -1

Dim Cad As String

Private Sub cmdAceptar_Click()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    
    'ManipuladorNumCarnet ManipuladorFecCaducidad ManipuladorNombre
    'ok, solo hay uno, regresamos el valor empipado
    With ListView1.SelectedItem
            CadenaDesdeOtroForm = .SubItems(2) & "|"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & .SubItems(3) & "|"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & .SubItems(1) & "|"
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & .SubItems(4) & "|"
                
    End With
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.Tag = 0 Then
        Me.Tag = 1
        Text1.Text = CStr(Cliente)
        Text1_LostFocus
        If ListView1.ListItems.Count = 0 Then
            PonerFocoBtn cmdCancelar
        Else
            ListView1.SetFocus
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
  '  Me.imgBuscar(0).Picture = frmPpal.ImageList3.ListImages(1).Picture
    Me.Tag = 0
    
    
    Me.imgBuscar(0).visible = False
    BloquearTxt Text1, True
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Cad = CadenaSeleccion
End Sub

Private Sub imgBuscar_Click(Index As Integer)
'    Cad = ""
'    Set frmCli = New frmFacClientes3
'    frmCli.DatosADevolverBusqueda = "0|1|"
'    frmCli.Show vbModal
'    Set frmCli = Nothing
'    If Cad <> "" Then
'        Text1.Text = RecuperaValor(Cad, 1)
'        Text1_LostFocus
'    End If
End Sub






Private Sub ListView1_DblClick()

  cmdAceptar_Click
    
End Sub

Private Sub Text1_GotFocus()
     ConseguirFoco Text1, 3
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub Text1_LostFocus()
Dim Continuar As Boolean
  

    ListView1.ListItems.Clear
    Text2.Text = ""
    
    If Text1.Text <> "" Then
        If Not EsNumerico(Text1.Text) Then
            Text1.Text = ""
            PonerFoco Text1
        End If
    End If
    
    If Text1.Text <> "" Then
    
        Cad = "Select nomclien elnom ,nifclien elnif ,clivario,ManipuladortipoCarnet as tipo,ManipuladorNumCarnet as nume,ManipuladorFecCaducidad as feccad"
        Cad = Cad & " from sclien where codclien =" & Text1.Text
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        NumRegElim = 0
        Continuar = False
        If miRsAux.EOF Then
            Cad = "No existe el cliente"
        Else
            If miRsAux!Clivario = 1 Then
                Cad = "Cliente de varios"
            Else
                'OK, vamos p'alla
                Text2.Text = miRsAux!elnom
                If DBLet(miRsAux!nume, "T") <> "" Then AnyadeItem True
                Continuar = True
                'If NumRegElim = 0 Then NumRegElim = 1 'Para que siga
            End If
        End If
        miRsAux.Close
        
        If Not Continuar Then
            MsgBox Cad, vbExclamation
           
            
        Else
            'Vemos si tiene AUTORIZADOS
            Cad = "SELECT cif elnif,nombre elnom,tipocarnet as tipo,numcarnet as nume,fcaducidad  as feccad FROM sclienmani "
            Cad = Cad & " WHERE codclien = " & Text1.Text
            miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                AnyadeItem False
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
        End If
        
        Set miRsAux = Nothing



    End If
End Sub

Private Sub AnyadeItem(DelPPal As Boolean)
    'nomclien,clivario,ManipuladortipoCarnet,ManipuladorNumCarnet,ManipuladorFecCaducidad "
    NumRegElim = NumRegElim + 1
    
    
    
    ListView1.ListItems.Add , , CStr(miRsAux!ElNif)
    If DelPPal Then
        ListView1.ListItems(NumRegElim).Bold = True
    Else
        If NumRegElim = 1 Then ListView1.ListItems(NumRegElim).Bold = True
    End If
    
    ListView1.ListItems(NumRegElim).SubItems(1) = miRsAux!elnom
    ListView1.ListItems(NumRegElim).SubItems(2) = DBLet(miRsAux!nume, "T")
    ListView1.ListItems(NumRegElim).SubItems(3) = Format(miRsAux!feccad, "dd/mm/yyyy")
    If Trim(ListView1.ListItems(NumRegElim).SubItems(3)) = "" Then
        ListView1.ListItems(NumRegElim).SubItems(3) = "01/01/1900"
        ListView1.ListItems(NumRegElim).ForeColor = vbRed
    End If
    ListView1.ListItems(NumRegElim).SubItems(4) = IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
    If Not DelPPal Then ListView1.ListItems(NumRegElim).SubItems(5) = "Autorizado"
    If NumRegElim = 1 Then Set ListView1.SelectedItem = ListView1.ListItems(NumRegElim)
        
        
    
End Sub
