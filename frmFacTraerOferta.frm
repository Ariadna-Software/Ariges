VERSION 5.00
Begin VB.Form frmFacTraerOferta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traer Lineas de Oferta"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ClipControls    =   0   'False
   Icon            =   "frmFacTraerOferta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FrameNuevo 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3000
      TabIndex        =   12
      Top             =   240
      Width           =   6975
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         ItemData        =   "frmFacTraerOferta.frx":000C
         Left            =   5880
         List            =   "frmFacTraerOferta.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "frmFacTraerOferta.frx":0028
         Left            =   3360
         List            =   "frmFacTraerOferta.frx":0035
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "frmFacTraerOferta.frx":0044
         Left            =   1080
         List            =   "frmFacTraerOferta.frx":0051
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   4920
         TabIndex        =   16
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   2640
         TabIndex        =   15
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Aceptada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   0
         TabIndex        =   14
         Top             =   900
         Width           =   885
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1080
         ToolTipText     =   "Buscar Nº oferta"
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Oferta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Copiar observaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1650
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Datos carta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2920
      TabIndex        =   6
      Top             =   1650
      Width           =   1575
   End
   Begin VB.Frame FrameAntiguo 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   2055
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   0
         TabIndex        =   0
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Oferta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   900
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   960
         ToolTipText     =   "Buscar Nº oferta"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1320
         ToolTipText     =   "Buscar Nº oferta en HCO."
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2160
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
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
      Left            =   240
      TabIndex        =   9
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmFacTraerOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Antiguo As Boolean  'Modo antiguo: oferta e historico       Nuevo. Todo en ofertas


Dim NombreTabla As String
Dim Ordenacion As String


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmOfe As frmBasico2
Attribute frmOfe.VB_VarHelpID = -1

Private Modo As Byte 'Solo utilizamos el Modo=4 -> Modificar

Dim kCampo As Integer

Public Event CargarOferta2(NumOfert As String)

'Private HaDevueltoDatos As Boolean
Private HanSeleccionadoHistorico As Byte

Private Sub cmdAceptar_Click()
Dim NumOfe As String
Dim cad As String

On Error GoTo Error1

        cad = ""
        If Antiguo Then
        
            'Lo de antes
            If Text1.Text = "" Then Exit Sub
        
            'Comprobar que la oferta existe
            'Si habian seleccionado desde manteimiento o no habian selecionado nada
            NumOfe = ""
            If HanSeleccionadoHistorico <> 2 Then NumOfe = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", Text1.Text, "N")
            
            'Si habian seleccionado desde manteimiento o no habian selecionado nada
            cad = ""
            If HanSeleccionadoHistorico <> 1 Then cad = DevuelveDesdeBDNew(conAri, "schpre", "numofert", "numofert", Text1.Text, "N")
            
            
            
            
            If cad <> "" Then
                If NumOfe <> "" Then
                    'Existe la misma ofertam, tanto en historico como en opedido. Preguntar
                    cad = "Existe el mismo codigo de oferta en mantimiento como en el pedido."
                    cad = cad & vbCrLf & "¿Desea cojer la del historico?"
                    kCampo = MsgBox(cad, vbQuestion + vbYesNoCancel)
                    If CByte(kCampo) = vbCancel Then Exit Sub
                    cad = ""
                    If CByte(kCampo) = vbYes Then cad = "H"
                    cad = cad & Text1.Text
                Else
                    cad = "H" & Text1.Text
                End If
            Else
                cad = NumOfe
            End If
        
        Else
            'NUEVO
            
            If Text2.Text = "" Then Exit Sub
            
            cad = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", Text2.Text, "N")
        
        
        End If
        If cad = "" Then
            MsgBox "Ninguna oferta con ese código", vbExclamation
            Exit Sub
        End If
    NumOfe = cad & "|" & Check1(0).Value & "|" & Check1(1).Value & "|"
    Unload Me
    RaiseEvent CargarOferta2(NumOfe)

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Unload Me
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

    If Antiguo Then
        Me.imgBuscar(0).Picture = frmPpal.imgListComun.ListImages(19).Picture
        Me.imgBuscar(1).Picture = frmPpal.imgListComun.ListImages(19).Picture
        Me.Width = 4995
        cmdAceptar.Left = 2280
        cmdCancelar.Left = 3600
    Else
        Me.imgBuscar(2).Picture = frmPpal.imgListComun.ListImages(1).Picture
        Me.Width = 7530
        cmdAceptar.Left = 4800
        cmdCancelar.Left = 6120
        FrameNuevo.Left = FrameAntiguo.Left
        Me.Combo1(0).ListIndex = 0
        Me.Combo1(1).ListIndex = 1
        Me.Combo1(2).ListIndex = 0
    End If
    FrameAntiguo.visible = Antiguo
    Me.FrameNuevo.visible = Not Antiguo
    
    
    
    PonerModo 0
    HanSeleccionadoHistorico = 0 'Dira si ha pulsado sobre el boton de Historico
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    
    Text1.Text = RecuperaValor(CadenaDevuelta, 1)
    If Val(imgBuscar(0).Tag) = 1 Then
        HanSeleccionadoHistorico = 2 'Desde el hco
    Else
        HanSeleccionadoHistorico = 1 'Desde normal
    End If
    Text1_LostFocus
End Sub

Private Sub frmOfe_DatoSeleccionado(CadenaSeleccion As String)
    Text2.Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim C As String

    'Antiguo
    If Index < 2 Then
        imgBuscar(0).Tag = Index
        Set frmB = New frmBuscaGrid
        frmB.vCampos = "Nº Oferta|scapre|numofert|N|0000000|13·Fecha Ofer.|scapre|fecofert|F|dd/mm/yyyy|15·Cliente|scapre|codclien|N|000000|12·" & _
            "Nombre Cliente|scapre|nomclien|T||45·Importe||sum(importel) importel|T|#,##0.00|13·"
        If Index = 0 Then
            frmB.vTabla = "scapre,slipre"
            frmB.vTitulo = "Ofertas"
        Else
            frmB.vTabla = "schpre scapre,slhpre slipre"
            frmB.vTitulo = "HISTORICO Ofertas"
        End If
        frmB.vSQL = "scapre.numofert=slipre.numofert group by numofert"
        
        
 
        frmB.vDevuelve = "0|"
        
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri
        frmB.Show vbModal



    Else
    
        'Llamaremos al basico2
            C = ""
            If Me.Combo1(0).ListIndex > 0 Then C = C & " AND aceptado =" & IIf(Combo1(0).ListIndex = 2, 0, 1)
            If Me.Combo1(1).ListIndex > 0 Then C = C & " AND numpedcl " & IIf(Combo1(1).ListIndex = 1, " > 0 ", " is null ")
            If Me.Combo1(2).ListIndex > 0 Then C = C & " AND motivoTraspaso " & IIf(Combo1(0).ListIndex = 1, " > 0 ", " is null ")
            If C <> "" Then C = Mid(C, 5) 'quitamos el primer AND
            
            Set frmOfe = New frmBasico2
            AyudaOfertas frmOfe, , C, True
            Set frmOfe = Nothing
            
            
    
    
    
    End If


End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, Modo
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus()
Dim devuelve As String

    With Text1
        If .Text = "" Then
            HanSeleccionadoHistorico = 0 'Reseteo
            Exit Sub
        End If
        .Text = Format(.Text, "0000000")
'        'Comprobar que la oferta existe
'        Devuelve = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", .Text, "N")
'        If Devuelve = "" Then
'            MsgBox "No existe la Oferta: " & .Text, vbInformation
'            text1.Text = ""
'            PonerFoco text1
'        End If
    End With
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
       
    Modo = Kmodo
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = Format(Text2.Text, "0000000")
End Sub
