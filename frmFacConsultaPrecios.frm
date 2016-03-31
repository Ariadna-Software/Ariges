VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacConsultaPrecios2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8250
   ClientLeft      =   345
   ClientTop       =   2430
   ClientWidth     =   16740
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   16740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePDF 
      Height          =   8175
      Left            =   11760
      TabIndex        =   49
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "Ver PDF"
         Height          =   615
         Left            =   3720
         TabIndex        =   51
         Top             =   7440
         Visible         =   0   'False
         Width           =   975
      End
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   6615
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   4455
         _cx             =   5080
         _cy             =   5080
      End
   End
   Begin VB.Frame FrameMostrarDatos 
      Height          =   8175
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   11535
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkCtrolStock 
         Caption         =   "Check1"
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Limpiar"
         Height          =   330
         Index           =   2
         Left            =   9120
         TabIndex        =   3
         Top             =   7680
         Width           =   975
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   7680
         Width           =   1215
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   7665
         Width           =   615
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   7665
         Width           =   615
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   7665
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmFacConsultaPrecios.frx":0000
         Left            =   240
         List            =   "frmFacConsultaPrecios.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   4440
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListStock 
         Height          =   1575
         Left            =   5040
         TabIndex        =   33
         Top             =   2400
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2778
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Almacen"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Stock"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Ped.cli"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Ped Prov"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Disponible"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2040
         Width           =   1215
      End
      Begin MSComctlLib.ListView listTarifa 
         Height          =   1575
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tarifa"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Des. tarifa"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Precio"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtResultado 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   7
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   465
         Width           =   615
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   6
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   465
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   465
         Width           =   975
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   4
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtResultado 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtResultado 
         Height          =   285
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtResultado 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   330
         Index           =   1
         Left            =   10320
         TabIndex        =   4
         Top             =   7680
         Width           =   975
      End
      Begin MSComctlLib.ListView listDatos 
         Height          =   2895
         Left            =   2520
         TabIndex        =   34
         Top             =   4440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5106
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "T"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Documento"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   2152
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Dto1"
            Object.Width           =   1023
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Dto2"
            Object.Width           =   1023
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   1765
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Stock"
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   48
         Top             =   2055
         Width           =   975
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmFacConsultaPrecios.frx":003B
         ToolTipText     =   "Buscar artículo"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmFacConsultaPrecios.frx":013D
         ToolTipText     =   "Buscar cliente"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblSituacion 
         Caption         =   "Label3"
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
         Left            =   7680
         TabIndex        =   45
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   6840
         TabIndex        =   44
         Top             =   7710
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   4800
         TabIndex        =   43
         Top             =   7710
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dto2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   42
         Top             =   7710
         Width           =   465
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dto1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   41
         Top             =   7710
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   40
         Top             =   7710
         Width           =   975
      End
      Begin VB.Label lblIndicador 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Left            =   9240
         TabIndex        =   35
         Top             =   7440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Disponible"
         Height          =   255
         Index           =   14
         Left            =   9360
         TabIndex        =   32
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "P.V.P."
         Height          =   255
         Index           =   13
         Left            =   2760
         TabIndex        =   29
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "P.M.P"
         Height          =   255
         Index           =   12
         Left            =   5280
         TabIndex        =   27
         Top             =   2055
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "P.U.C"
         Height          =   255
         Index           =   11
         Left            =   600
         TabIndex        =   25
         Top             =   2055
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Descrp."
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Pendiente(€)"
         Height          =   255
         Index           =   8
         Left            =   9240
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Dto Pp."
         Height          =   255
         Index           =   7
         Left            =   7320
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Dto gral"
         Height          =   255
         Index           =   6
         Left            =   5640
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "F. pago"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Situación"
         Height          =   255
         Index           =   4
         Left            =   5640
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ofertas,pedidos, albaranes,facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   4080
         X2              =   11280
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Index           =   1
         X1              =   1080
         X2              =   11400
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Index           =   0
         X1              =   1320
         X2              =   11400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   120
         Top             =   7560
         Width           =   8775
      End
   End
End
Attribute VB_Name = "frmFacConsultaPrecios2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Fecha As String   'Podra ser "" o una fecha valida
Public ConsultaDesdeFrm As String
Private WithEvents frmA As frmAlmArticu2
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmC As frmFacClientes3
Attribute frmC.VB_VarHelpID = -1
Private frmAlb As frmFacEntAlbaranes2

Dim Cad As String
Dim IT As ListItem

Dim Valor As Currency
Dim AntiguoTxt As String
Dim PrimeraVez As Boolean



Private Sub LimpiarResultados()
Dim T As TextBox
Dim Index As Integer
    lblIndicador.Caption = ""
    For Each T In Me.txtResultado
        Index = T.Index
        If Index <> 0 And Index <> 1 And Index <> 8 And Index <> 9 Then T.Text = ""
    Next
    Me.listTarifa.ListItems.Clear
    Me.ListStock.ListItems.Clear
    Me.listDatos.ListItems.Clear
    label2(4).Caption = ""

    lblSituacion.Caption = ""
    
    
    
    
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 1 Then
        'PonerVisiblePedir True
        Unload Me
    ElseIf Index = 2 Then
        limpiar Me
        LimpiarResultados
    Else
        Unload Me
    End If
End Sub



Private Sub Combo1_Click()
    If PrimeraVez Then Exit Sub
    '------------------------------------------
    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "Leyendo BD"
    Set miRsAux = New ADODB.Recordset
    CargarDatosFacturacion
    lblIndicador.Caption = ""
    Set miRsAux = Nothing
    Me.lblIndicador.Caption = ""
    Screen.MousePointer = vbDefault
End Sub


Private Sub Command1_Click()
    If Me.AcroPDF1.visible Then
        frmEulerPDF.Tag = Me.AcroPDF1.Tag
        frmEulerPDF.Show vbModal
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Fecha = "" Then Fecha = Now
        Caption = Caption & "     (" & Format(Fecha, "dd/mm/yyyy") & ")"
        If ConsultaDesdeFrm <> "" Then
            txtCodigo(0).Text = RecuperaValor(ConsultaDesdeFrm, 1)
            
            txtCodigo_LostFocus 0
            txtCodigo(1).Text = RecuperaValor(ConsultaDesdeFrm, 2)
            txtCodigo_LostFocus 1
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    limpiar Me
    Caption = "Consulta precios"
    'ASigno estos iconos
    lblIndicador.Caption = ""
    lblSituacion.Caption = ""
    Me.listDatos.SmallIcons = frmPpal.ImgListPpal
    If ConsultaDesdeFrm <> "" Then Me.Combo1.ListIndex = 1
        
    
    If vParamAplic.NumeroInstalacion = 4 Then
        Me.Width = 16965
    Else
        Me.Width = 11775
    End If
    
End Sub





Private Sub Form_Unload(Cancel As Integer)
    Fecha = ""
    ConsultaDesdeFrm = ""
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(1).Text = RecuperaValor(CadenaSeleccion, 1)
    txtResultado(9).Text = RecuperaValor(CadenaSeleccion, 2)
    Cad = "O"
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtResultado(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Cad = "O"
End Sub

Private Sub imgBuscarG_Click(Index As Integer)
Dim KCargo As Integer
    KCargo = -1
    Cad = "" 'Para ver si devuelve datos
    If Index = 0 Then
        'Cliente
        
        Set frmC = New frmFacClientes3
        frmC.DatosADevolverBusqueda = "1|"
        frmC.Show vbModal
        Set frmC = Nothing
        'If Cad <> "" Then PonerFoco txtCodigo(1)
        If Cad <> "" Then
            'cmdBuscar_Click
            KCargo = 0
            If txtCodigo(1).Text <> "" Then KCargo = 2
        End If
    Else
        'Articulo
        Set frmA = New frmAlmArticu2
        'frmA.DeConsulta = True
        'frmA.DatosADevolverBusqueda3 = "@1@"
        frmA.DesdeTPV = False
        frmA.Show vbModal
        Set frmA = Nothing
        If Cad <> "" Then
            'cmdBuscar_Click
            KCargo = 1
            If txtCodigo(0).Text <> "" Then KCargo = 2
        End If
    End If
    Cad = ""
    Set miRsAux = New ADODB.Recordset
    If KCargo >= 0 Then CargarDatos CByte(KCargo)
    Set miRsAux = Nothing
End Sub


'
Private Sub listDatos_DblClick()
Dim Seleccionado As Long
Dim SQL As String

    If listDatos.ListItems.Count = 0 Then Exit Sub
    If listDatos.SelectedItem Is Nothing Then Exit Sub


    If ConsultaDesdeFrm <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta en proceso de generacion de oferta/pedido/albaran.", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case Combo1.ListIndex
    
    Case 0
            'Si lo primero no es una F no es una factura
            SQL = ""
            
            If Mid(Me.listDatos.SelectedItem.Tag, 1, 1) <> "F" Then
                MsgBox "No es una  factura (F)", vbExclamation
            Else
                SQL = Mid(listDatos.SelectedItem.Tag, 1, 3) & "|" & Mid(listDatos.SelectedItem.Tag, 4) & "|"
                SQL = SQL & listDatos.SelectedItem.SubItems(1) & "|"
            End If
            If SQL = "" Then Exit Sub
            With frmFacHcoFacturas2
                .DesdeFichaCliente = True
                .hcoCodMovim = RecuperaValor(SQL, 2)
                .hcoCodTipoM = RecuperaValor(SQL, 1)
                .hcoFechaMov = RecuperaValor(SQL, 3)
                .Show vbModal
        End With
    Case 1
              'ALBARANES
            'Si lo primero no es una A no es una factura
            SQL = ""
            
            If Mid(Me.listDatos.SelectedItem.Tag, 1, 1) <> "A" Then
                MsgBox "No es un albaran(A)", vbExclamation
            Else
                SQL = Mid(listDatos.SelectedItem.Tag, 1, 3) & "|" & Mid(listDatos.SelectedItem.Tag, 4) & "|"
            End If
            If SQL = "" Then Exit Sub
              
        If vParamAplic.TipoFormularioClientes = 0 Then
            
            frmFacEntAlbaranes2.hcoCodMovim = RecuperaValor(SQL, 2)
            frmFacEntAlbaranes2.hcoCodTipoM = RecuperaValor(SQL, 1)
            frmFacEntAlbaranes2.Show vbModal
         
            
        Else
         
            frmFacEntAlbSAIL.hcoCodMovim = listDatos.SelectedItem.SubItems(1)
            frmFacEntAlbSAIL.hcoCodTipoM = listDatos.SelectedItem.Text
            frmFacEntAlbSAIL.Show vbModal
     
                 
            
        End If
    Case 2
        
            'PEDIDO CLIENTE
            If vParamAplic.TipoFormularioClientes = 0 Then
            
                frmFacEntPedidos.DatosADevolverBusqueda2 = listDatos.SelectedItem.Tag
                frmFacEntPedidos.EsHistorico = False
                frmFacEntPedidos.Show vbModal
            Else
                frmFacEntPedSail.DatosADevolverBusqueda2 = listDatos.SelectedItem.Tag
                frmFacEntPedSail.EsHistorico = False
                frmFacEntPedSail.Show vbModal
            End If
    Case 3
        'ofertas
            If vParamAplic.TipoFormularioClientes = 0 Then
             
                frmFacEntOfertas2.DatosOferta = listDatos.SelectedItem.Tag
                frmFacEntOfertas2.Show vbModal
                
            Else
                frmFacEntOferSAIL.DatosOferta = listDatos.SelectedItem.Tag
                frmFacEntOferSAIL.Show vbModal
            End If
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    listDatos.SetFocus
    Seleccionado = listDatos.SelectedItem.Index
    
    Combo1_Click
    If Not listDatos.SelectedItem Is Nothing Then listDatos.SelectedItem.Selected = False
    Set listDatos.SelectedItem = Nothing
    If listDatos.ListItems.Count >= Seleccionado Then
            listDatos.ListItems(Seleccionado).Selected = True
            listDatos.ListItems(Seleccionado).EnsureVisible
    End If
    Screen.MousePointer = vbDefault


End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    AntiguoTxt = txtCodigo(Index).Text
    ConseguirFoco txtCodigo(Index), 3
End Sub



Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Opc As Byte
    txtCodigo(Index).Text = Trim(txtCodigo(Index))
    If AntiguoTxt = txtCodigo(Index).Text Then Exit Sub
    Cad = ""
    Opc = 100
    If Index = 0 Then
        lblIndicador.Caption = ""
        
        'Cliente
        If txtCodigo(Index).Text <> "" Then
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "Campo codigo cliente debe ser numérico", vbExclamation
                
            Else
                Cad = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCodigo(Index).Text, "N")
                If Cad = "" Then
                    MsgBox "No existe el cliente : " & txtCodigo(Index).Text, vbExclamation
                End If
            End If
        End If
        If Cad <> "" Then
            Opc = 0
            If txtCodigo(1).Text <> "" Then Opc = 2
            
        End If
    Else
        'articulo
        If txtCodigo(Index).Text <> "" Then
            Cad = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", txtCodigo(Index).Text, "T")
            If Cad = "" Then
                MsgBox "No existe el articulo: " & txtCodigo(Index).Text, vbExclamation
                PonerFoco txtCodigo(Index)
            End If
        End If
        If Cad <> "" Then
            Opc = 1
            If txtCodigo(0).Text <> "" Then Opc = 2
        End If
    End If
    'Me.txtNombre(Index).Text = Cad
    If Cad = "" Then
        txtCodigo(Index).Text = ""
        PonerFoco txtCodigo(Index)
    End If
    If Opc = 100 Then
        'Mal. Borramos los campos
        
        LimpiarlosCampos CByte(Index)
        
            
    Else
        Set miRsAux = New ADODB.Recordset
        CargarDatos Opc
        Set miRsAux = Nothing
    End If
End Sub


Private Sub CargaStock()
Dim I As Currency
Dim J As Integer

    Valor = 0
    ListStock.ListItems.Clear
    txtResultado(13).Text = ""
    Cad = "select salmac.codalmac,nomalmac,canstock   from salmac,salmpr where salmac.codalmac="
    Cad = Cad & "salmpr.codalmac AND  codartic=" & DBSet(txtCodigo(1).Text, "T") & " ORDER BY salmac.codalmac"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListStock.ListItems.Add()
        IT.Text = miRsAux!codAlmac
        IT.SubItems(1) = miRsAux!nomalmac
        I = DBLet(miRsAux!CanStock, "N")
        IT.SubItems(2) = Format(I, FormatoCantidad)
        IT.SubItems(3) = " ": IT.SubItems(4) = " "
        IT.SubItems(5) = IT.SubItems(2)
        Valor = Valor + I
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Stock
    txtResultado(13).Text = Format(Valor, FormatoCantidad)
    
    'Cargamos primero los de cliente
    'FALTA###
    'If chkCtrolStock.Value Then
    If True Then
        Cad = "select codalmac,sum(cantidad) as cuantos"
        Cad = Cad & " from sliped where codartic='"
        Cad = Cad & DevNombreSQL(txtCodigo(1).Text) & "' GROUP BY 1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
              For J = 1 To ListStock.ListItems.Count
                    If ListStock.ListItems(J).Text = CStr(miRsAux.Fields(0)) Then
                        'ES este
                        I = DBLet(miRsAux.Fields(1), "N")
                        If I <> 0 Then ListStock.ListItems(J).SubItems(3) = Format(I, FormatoCantidad)
                        Valor = Valor - I
                        
                        I = ImporteFormateado(ListStock.ListItems(J).SubItems(2)) - I
                        ListStock.ListItems(J).SubItems(5) = Format(I, FormatoCantidad)
                        Exit For
                    End If
            Next
            miRsAux.MoveNext
        Wend
        miRsAux.Close


'    'Cargamos los comprados
'    C = "select scappr.numpedpr,fecpedpr,codprove,nomprove,sum(cantidad) as cuantos"
'    C = C & " from scappr,slippr where scappr.numpedpr=slippr.numpedpr  and codartic='"
'    C = C & DevNombreSQL(Data1.Recordset!codArtic) & "' group by 1"

        Cad = "select codalmac,sum(cantidad) as cuantos"
        Cad = Cad & " from slippr where codartic='"
        Cad = Cad & DevNombreSQL(txtCodigo(1).Text) & "' GROUP BY 1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
              For J = 1 To ListStock.ListItems.Count
                    If ListStock.ListItems(J).Text = CStr(miRsAux.Fields(0)) Then
                        'ES este
                        I = DBLet(miRsAux.Fields(1), "N")
                        If I <> 0 Then ListStock.ListItems(J).SubItems(4) = Format(I, FormatoCantidad)
                        Valor = Valor + I
                        
                        I = ImporteFormateado(ListStock.ListItems(J).SubItems(2)) + I
                                'los pedidos clientes (reservas)
                        I = I - ImporteFormateado(Trim(ListStock.ListItems(J).SubItems(3)))
                        ListStock.ListItems(J).SubItems(5) = Format(I, FormatoCantidad)
          
                        Exit For
                    End If
            Next
            miRsAux.MoveNext
        Wend
        miRsAux.Close


    End If
    
    'Disponible
    txtResultado(0).Text = Format(Valor, FormatoCantidad)
    
End Sub


Private Sub CargaTarifas()
Dim F As Date

    
    
    
    listTarifa.ListItems.Clear
    Cad = "select slista.codlista,nomlista,precioac,fechanue,precionu from slista,starif where slista.codlista="
    Cad = Cad & "starif.codlista and codartic = " & DBSet(txtCodigo(1).Text, "T") & " ORDER BY slista.codlista"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = listTarifa.ListItems.Add()
        IT.Text = miRsAux!codlista
        IT.SubItems(1) = miRsAux!nomlista
        
        F = CDate("01/01/2111")
        If Not IsNull(miRsAux!fechanue) Then
            If Not IsNull(miRsAux!precionu) Then F = miRsAux!fechanue
        End If
        
        If CDate(Fecha) >= F Then
            'Coje el nuevo
            IT.SubItems(2) = Format(DBLet(miRsAux!precionu, "N"), FormatoPrecio) & " *"
        Else
            IT.SubItems(2) = Format(DBLet(miRsAux!precioac, "N"), FormatoPrecio)
        End If
        If miRsAux!codlista = listTarifa.Tag Then
            'Tarifa del cliente
            IT.Bold = True
            IT.ForeColor = vbBlue
            IT.ListSubItems(1).Bold = True
            IT.ListSubItems(1).ForeColor = vbBlue
    
            IT.ListSubItems(2).Bold = True
            IT.ListSubItems(2).ForeColor = vbBlue
            
        End If
                    
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If listTarifa.ListItems.Count > 6 Then
        listTarifa.ColumnHeaders.item(2).Width = 2600
    Else
        listTarifa.ColumnHeaders.item(2).Width = 2800
    End If
    
End Sub



'0. Cliente
'1.- Articulop
'2.los dos

Private Sub CargarDatos(Opcion As Byte)
Dim Familia As Integer
Dim marca As Integer

    On Error GoTo EC
    
    Cad = "OK"

    If Opcion <> 1 Then
        lblIndicador.Caption = "Datos cliente"
        lblIndicador.Refresh
        
        Cad = "select codclien ,nomclien ,dtoppago ,dtognral  ,codsitua ,codmacta,codforpa,codtarif  from sclien where codclien =" & Me.txtCodigo(0).Text
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        'Ponemos los campos
        '--------------------------------------------------------
        Me.txtCodigo(0).Text = miRsAux!codClien
        Me.txtResultado(1).Text = miRsAux!Nomclien
        Me.txtResultado(2).Text = miRsAux!codforpa
        Me.txtResultado(3).Text = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", miRsAux!codforpa, "N")
        Me.txtResultado(4).Text = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", miRsAux!codsitua, "N")
        Me.txtResultado(6).Text = Format(miRsAux!DtoGnral, FormatoDescuento)
        Me.txtResultado(7).Text = Format(miRsAux!DtoPPago, FormatoDescuento)
        
        'Cargo la cta contable
        Cad = DBLet(miRsAux!Codmacta, "T")
        
        'Cargo la tarifa
        Me.listTarifa.Tag = miRsAux!codTarif
        
        'Cerramos el RS
        miRsAux.Close
    
    
    
    
        lblIndicador.Caption = "Cobros pendientes"
        lblIndicador.Refresh
    
        PonerCobrosPendientes Cad
    
        txtResultado(5).Text = Format(Valor, FormatoImporte)
    
        DoEvents
    End If
    'Datos articulo
    If Opcion <> 0 Then
        lblIndicador.Caption = "Articulo"
        lblIndicador.Refresh
        
        Cad = "select codartic,nomartic,preciouc,preciomp,preciove,unicajas,codstatu,ctrstock,codfamia,codmarca  from sartic where codartic =" & DBSet(Me.txtCodigo(1).Text, "T")
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Me.txtCodigo(1).Text = miRsAux!codArtic
        Me.txtResultado(9).Text = miRsAux!NomArtic
        Me.txtResultado(10).Text = Format(DBLet(miRsAux!precioUC, "N"), FormatoPrecio)
        Me.txtResultado(11).Text = Format(DBLet(miRsAux!PrecioMP, "N"), FormatoPrecio)
        Me.txtResultado(12).Text = Format(DBLet(miRsAux!PrecioVe, "N"), FormatoPrecio)
        chkCtrolStock.Value = miRsAux!CtrStock  'guardare si lleva control de stock
        Me.txtResultado(9).Tag = miRsAux!unicajas
        Select Case miRsAux!codstatu
        'Abril 2014
        Case 1
            lblSituacion.Caption = "Obsoleto"
        Case 2
            lblSituacion.Caption = "Bloqueado"
        Case 3
            lblSituacion.Caption = "Caducado"
        Case Else
            lblSituacion.Caption = ""
        End Select
        Familia = miRsAux!Codfamia
        marca = miRsAux!codmarca
        miRsAux.Close
        
            
        
        
        lblIndicador.Caption = "Stock"
        lblIndicador.Refresh
        CargaStock
        
        
        lblIndicador.Caption = "Tarifas"
        lblIndicador.Refresh
        CargaTarifas
    
    
        If vParamAplic.NumeroInstalacion = 4 Then
            'Si la familia, marca tiene catalogo lo mostrara
            Cad = "Select * from eulerprecios  WHERE "
            Cad = Cad & "( codfamia =" & Familia & " AND codmarca =" & marca & ")"
            Cad = Cad & " OR ( codfamia =" & Familia & " AND codmarca is null )"
            Cad = Cad & " OR ( codfamia is NULL AND codmarca =" & marca & ")"
            miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            If Not miRsAux.EOF Then
                'OOOOOOK
                'Tiene un documento asociado
                Cad = miRsAux!Documento
                
            End If
            miRsAux.Close
            CargaArchivo Cad
            
        End If
        DoEvents
    
    End If
    If Opcion = 2 Then
        'Datos albaranes......
        CargarDatosFacturacion
    
    
        'Ponemos el precio fianl
        CalcularPrecioFinal
    
    End If
    
    'Insertamos log de consulta
    If txtCodigo(0).Text <> "" And txtCodigo(1).Text <> "" Then
        If ConsultaDesdeFrm = "" Then
            lblIndicador.Caption = "Ins. log"
            lblIndicador.Refresh
            Cad = "insert into `sconsulta` (`DiaHora`,`Usuario`,`codclien`,`nomclien`,"
            '----------                                       cogera la fecha del mysql
            Cad = Cad & "`codartic`,`nomartic`) values (" & "concat(curdate(),' ',curtime())" & ","
            Cad = Cad & DBSet(vUsu.Nombre, "T") & "," & txtCodigo(0).Text & "," & DBSet(txtResultado(1), "T")
            Cad = Cad & "," & DBSet(txtCodigo(1), "T") & "," & DBSet(txtResultado(9), "T") & ")"
            conn.Execute Cad
            Espera 0.3
            
        End If
    End If
    lblIndicador.Caption = ""
        
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
End Sub





Private Sub PonerCobrosPendientes(ByVal Codmacta As String)
    Valor = 0
    If Codmacta = "" Then Exit Sub
    'Obtener a partir de la cuenta del cliente si hay cobros pendientes en Contabilidad
    Cad = " WHERE scobro.codmacta = '" & Codmacta & "'"
    Cad = Cad & " AND fecvenci <= ' " & Format(Now, FormatoFecha) & "' "
    Cad = Cad & " AND (sforpa.tipforpa between 0 and 3)"
    
    If vParamAplic.ContabilidadNueva Then
        Cad = " cobros as scobro INNER JOIN formapago as sforpa ON scobro.codforpa=sforpa.codforpa " & Cad
    Else
        Cad = " scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa " & Cad
    End If
    Cad = "SELECT sum(impvenci + coalesce(gastos,0) - coalesce(impcobro,0)) FROM " & Cad
    'Lee de la Base de Datos de CONTABILIDAD
    miRsAux.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then Valor = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
End Sub



'Cargara los datos de las lineas
'de OFERTAS,PEDIDOS,ALBARANES,FACTURA
Private Sub CargarDatosFacturacion()

    Me.listDatos.ListItems.Clear
    
    If Me.txtCodigo(1).Text <> "" And txtCodigo(0).Text <> "" Then
        If Combo1.ListIndex >= 0 Then CargaDatosTablas CByte(Combo1.ListIndex)
    End If
    
    
End Sub



Private Sub CargaDatosTablas(Ktabla As Byte)
Dim Aux As String
Dim Ico As Integer
    Select Case Ktabla
    Case 3
        Ico = 5
        Cad = "slipre,scapre WHERE slipre.numofert=scapre.numofert"
        Aux = " '' as Primero,slipre.numofert as elnumero,fecofert as fecha"
        Me.lblIndicador.Caption = "Ofertas"
    Case 2
        Ico = 6
        Cad = "sliped,scaped where sliped.numpedcl=scaped.numpedcl"
        Aux = " '' as Primero,sliped.numpedcl as elnumero,fecpedcl  as fecha"
        Me.lblIndicador.Caption = "Pedidos"
    Case 1
        Ico = 7
        Cad = "slialb,scaalb where slialb.numalbar=scaalb.numalbar and slialb.codtipom=scaalb.codtipom"
        Aux = " slialb.codtipom as Primero,slialb.numalbar as elnumero,fechaalb as fecha"
        Me.lblIndicador.Caption = "Albaranes"
    Case 0
        Ico = 8
        Cad = " slifac,scafac where slifac.numfactu=scafac.numfactu and slifac.codtipom=scafac.codtipom and slifac.fecfactu=scafac.fecfactu"
        Aux = "slifac.codtipom as primero,slifac.numfactu as elnumero,slifac.fecfactu as fecha"
        Me.lblIndicador.Caption = "Facturas"
    End Select
    Me.lblIndicador.Refresh
    
    Aux = "Select " & Aux & ",Cantidad, precioar, dtoline1, dtoline2, ImporteL FROM " & Cad
    Cad = Aux & " AND codartic = " & DBSet(Me.txtCodigo(1).Text, "T")
    Cad = Cad & " AND codclien = " & txtCodigo(0).Text
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = Me.listDatos.ListItems.Add()
        Cad = miRsAux!primero & Format(miRsAux!elnumero, "000000")
        IT.Text = ""
        
        IT.SubItems(1) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        
        'Nuevo. El documento
        IT.SubItems(2) = Cad
        IT.SubItems(3) = Format(miRsAux!cantidad, FormatoCantidad)
        IT.SubItems(4) = Format(miRsAux!precioar, FormatoPrecio)
        IT.SubItems(5) = Format(miRsAux!dtoline1, FormatoDescuento)
        IT.SubItems(6) = Format(miRsAux!dtoline2, FormatoDescuento)
        IT.SubItems(7) = Format(miRsAux!ImporteL, FormatoImporte)
        IT.SmallIcon = Ico
        IT.Tag = Cad
        IT.ToolTipText = Me.lblIndicador.Caption & " " & Cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub





'Calculamos el precio que se le va a quedar
Private Sub CalcularPrecioFinal()
Dim CPrecioFact As CPreciosFact
Dim Precio As Currency
Dim cantidad As Currency
Dim PorCaja As Boolean
Dim NumCajas As Integer
                Set CPrecioFact = New CPreciosFact
                'Si se puede vender por cajas(devuelve>1) poner numero de cajas en una linea con el
                'precio de caja, y otra linea con el resto unidades un precio unidad
                'Cantidad = txtAux(Index).Text
                cantidad = 1
                NumCajas = CPrecioFact.ObtenerNumCajas(CStr(cantidad), CStr(txtResultado(9).Tag))
                'RestoUnid = CInt(ComprobarCero(Cantidad)) - NumCajas * CInt(devuelve)
                'Obtenemos la Tarifa del Cliente
                CPrecioFact.CodigoArtic = Me.txtCodigo(1).Text
                CPrecioFact.CodigoClien = Me.txtCodigo(0).Text
                CPrecioFact.FijarTarifaActividad
                CPrecioFact.CodigoLista2 = Me.listTarifa.Tag  'la tarifa del cliente
                    
                
               
                PorCaja = (NumCajas > 0)
                Precio = CPrecioFact.ObtenerPrecio(PorCaja, CStr(Fecha), Cad, "")
                    
                'En cad TENGO el origen del precio
                Select Case Cad
                    Case "P": label2(4).Caption = "Promoción"
                    Case "E": label2(4).Caption = "Precio Especial"
                    Case "T": label2(4).Caption = "Tarifa Artículo"
                    Case "A": label2(4).Caption = "Precio Artículo"
                    Case "M": label2(4).Caption = "Manual"
                End Select
                
                    txtResultado(14).Text = Precio
                    PonerFormatoDecimal txtResultado(14), 2
                    txtResultado(15).Text = CPrecioFact.Descuento1
                    PonerFormatoDecimal txtResultado(15), 4
                    txtResultado(16).Text = CPrecioFact.Descuento2
                    PonerFormatoDecimal txtResultado(16), 4
                    txtResultado(17).Text = CalcularImporte(CStr(cantidad), txtResultado(14).Text, txtResultado(15).Text, txtResultado(16).Text, vParamAplic.TipoDtos)
                    PonerFormatoDecimal txtResultado(17), 1
                Set CPrecioFact = Nothing
End Sub




'0  Artic   1 Cliente      2 los dos
Private Sub LimpiarlosCampos(Opcion As Byte)
Dim I As Integer
 
    If Opcion <> 1 Then
        'ARTICULOS
        For I = 1 To 7
            txtResultado(I).Text = ""
        Next I
        chkCtrolStock.Value = 0  'guardare si lleva control de stock
        
        
        
    End If
    If Opcion <> 0 Then
        'CLIENTE
         For I = 9 To 13
            txtResultado(I).Text = ""
        Next I
        txtResultado(0).Text = ""
        Me.listTarifa.ListItems.Clear
        Me.ListStock.ListItems.Clear
        CargaArchivo ""
    End If
    
    
    listDatos.ListItems.Clear
    For I = 14 To 17
            txtResultado(I).Text = ""
    Next I
    label2(4).Caption = ""
    lblSituacion.Caption = ""
End Sub



Private Function CargaArchivo(Archivo As String) As Boolean
    On Error GoTo eCargaArchivo
    
    If vParamAplic.NumeroInstalacion <> 4 Then Exit Function
    
    CargaArchivo = False
    
    If Archivo = "" Then
        AcroPDF1.visible = False
    Else
        AcroPDF1.LoadFile (Archivo)
        AcroPDF1.visible = True
    End If
    AcroPDF1.Tag = Archivo
    Me.Command1.visible = AcroPDF1.visible
    Screen.MousePointer = vbDefault
    
    
    CargaArchivo = True
    Exit Function
eCargaArchivo:
    MuestraError Err.Number, "Carga archivo PDF"
End Function
