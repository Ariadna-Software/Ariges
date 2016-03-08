VERSION 5.00
Begin VB.Form frmFacTraerOferta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traer Lineas de Oferta"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ClipControls    =   0   'False
   Icon            =   "frmFacTraerOferta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Datos carta"
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Copiar observaciones"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1560
      ToolTipText     =   "Buscar Nº oferta en HCO."
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1080
      ToolTipText     =   "Buscar Nº oferta"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Oferta"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   855
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
      TabIndex        =   3
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

Dim NombreTabla As String
Dim Ordenacion As String


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Private Modo As Byte
'Solo utilizamos el Modo=4 -> Modificar

Dim kCampo As Integer

Public Event CargarOferta2(NumOfert As String)

Private HaDevueltoDatos As Boolean
Private HanSeleccionadoHistorico As Byte

Private Sub cmdAceptar_Click()
Dim NumOfe As String
Dim Cad As String

On Error GoTo Error1

    
        If Text1.Text = "" Then Exit Sub
    
        'Comprobar que la oferta existe
        'Si habian seleccionado desde manteimiento o no habian selecionado nada
        NumOfe = ""
        If HanSeleccionadoHistorico <> 2 Then NumOfe = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", Text1.Text, "N")
        
        'Si habian seleccionado desde manteimiento o no habian selecionado nada
        Cad = ""
        If HanSeleccionadoHistorico <> 1 Then Cad = DevuelveDesdeBDNew(conAri, "schpre", "numofert", "numofert", Text1.Text, "N")
        
        
        If Cad <> "" Then
            If NumOfe <> "" Then
                'Existe la misma ofertam, tanto en historico como en opedido. Preguntar
                Cad = "Existe el mismo codigo de oferta en mantimiento como en el pedido."
                Cad = Cad & vbCrLf & "¿Desea cojer la del historico?"
                kCampo = MsgBox(Cad, vbQuestion + vbYesNoCancel)
                If CByte(kCampo) = vbCancel Then Exit Sub
                Cad = ""
                If CByte(kCampo) = vbYes Then Cad = "H"
                Cad = Cad & Text1.Text
            Else
                Cad = "H" & Text1.Text
            End If
        Else
            Cad = NumOfe
        End If
        If Cad = "" Then
            MsgBox "Ninguna oferta con ese código", vbExclamation
            Exit Sub
        End If
    NumOfe = Cad & "|" & Check1(0).Value & "|" & Check1(1).Value & "|"
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
    Me.imgBuscar(0).Picture = frmPpal.imgListComun.ListImages(19).Picture
    Me.imgBuscar(1).Picture = frmPpal.imgListComun.ListImages(19).Picture
    
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

Private Sub imgBuscar_Click(Index As Integer)
Dim C As String

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
Dim Devuelve As String

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

