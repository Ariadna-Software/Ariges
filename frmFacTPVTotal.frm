VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacTPVTotal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Total Venta"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8655
   Icon            =   "frmFacTPVTotal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCarnet 
      Caption         =   "Carnet manipulador fitosanitarios"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   360
      TabIndex        =   29
      Top             =   3360
      Width           =   7455
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   4245
      End
      Begin VB.TextBox txtManipulador 
         Height          =   315
         Index           =   2
         Left            =   3000
         TabIndex        =   32
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtManipulador 
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   31
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtManipulador 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   30
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   35
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   34
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Nº"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdMtoCampos 
      Height          =   375
      Index           =   1
      Left            =   10320
      Picture         =   "frmFacTPVTotal.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Eliminar campo"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdMtoCampos 
      Height          =   375
      Index           =   0
      Left            =   9720
      Picture         =   "frmFacTPVTotal.frx":0A0E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Añadir campo"
      Top             =   720
      Width           =   375
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4935
      Left            =   8640
      TabIndex        =   26
      Top             =   1200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8705
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
         Text            =   "Campo"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Partida"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Variedad"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.CommandButton cmdCampos 
      Caption         =   "Campos  Ctrl+F8"
      Height          =   615
      Left            =   360
      Picture         =   "frmFacTPVTotal.frx":7260
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5565
      Width           =   1695
   End
   Begin VB.CommandButton cmdQUitarCliVar 
      Height          =   375
      Left            =   7200
      Picture         =   "frmFacTPVTotal.frx":8CD2
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   23
      Text            =   "Text3"
      Top             =   1440
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   5
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   1905
      Width           =   4245
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   360
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   2880
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "C"
      Top             =   3000
      Width           =   3885
   End
   Begin VB.Frame FrameEfectivo 
      Height          =   1695
      Left            =   2520
      TabIndex        =   6
      Top             =   4560
      Width           =   5295
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   3
         Left            =   2160
         TabIndex        =   5
         Text            =   "0.0"
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   405
         Index           =   1
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label label2 
         Appearance      =   0  'Flat
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   405
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   4
         Left            =   2160
         TabIndex        =   13
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label label2 
         Caption         =   "CAMBIO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label label2 
         Caption         =   "ENTREGADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   2460
      Width           =   3885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2460
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1005
      Width           =   4485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1005
      Width           =   855
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   1164
      ButtonWidth     =   2302
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ticket  F5"
            Object.ToolTipText     =   "Generar Ticket"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Albaran Ctr+F6"
            Object.ToolTipText     =   "Generar Albaran"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Factura  F7"
            Object.ToolTipText     =   "Generar Factura"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cli varios F9"
            Object.ToolTipText     =   "Cliente varios"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Carnet F11"
            Object.ToolTipText     =   "Carnet manipulador"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.Label label1 
      Caption         =   "Campos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Index           =   5
      Left            =   8640
      TabIndex        =   36
      Top             =   840
      Width           =   960
   End
   Begin VB.Label label1 
      Caption         =   "Dpto."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   22
      Top             =   1905
      Width           =   975
   End
   Begin VB.Image imgBuscar 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   5
      Left            =   1560
      ToolTipText     =   "Buscar direc./dpto"
      Top             =   1905
      Width           =   360
   End
   Begin VB.Label label1 
      Caption         =   "Cheque regalo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   4648
      Width           =   1815
   End
   Begin VB.Image imgBuscar 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   2
      Left            =   2400
      ToolTipText     =   "Buscar artículo"
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label label1 
      Caption         =   "Operador "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   18
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image imgBuscar 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   0
      Left            =   1680
      ToolTipText     =   "Buscar cliente"
      Top             =   1005
      Width           =   360
   End
   Begin VB.Label LabelB 
      Caption         =   "F2 = Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   480
      TabIndex        =   16
      Top             =   1320
      Width           =   1320
   End
   Begin VB.Image imgBuscar 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   1
      Left            =   2400
      ToolTipText     =   "Buscar forma de pago"
      Top             =   2460
      Width           =   360
   End
   Begin VB.Label label1 
      Caption         =   "Forma de pago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   480
      TabIndex        =   9
      Top             =   2460
      Width           =   1815
   End
   Begin VB.Label label1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   1005
      Width           =   975
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnTicket 
         Caption         =   "&Ticket"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnAlbaran 
         Caption         =   "&Albaran"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnFactura 
         Caption         =   "&Factura"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnAsociarCampos 
         Caption         =   "Asociar campos"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnClienteVarios 
         Caption         =   "Cliente &varios"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnManipulador 
         Caption         =   "Carnet de manipulador"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacTPVTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cadSel As String 'cadena para seleccion de la venta a totalizar
Public ImporteInicial As String

Public LLevaArticulosFitosanitarios As Boolean   'Por si pide el


'Public NumTermi As Integer

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmClv As frmFacClientesV
Attribute frmClv.VB_VarHelpID = -1

Private PrimeraVez As Boolean
Private HaDevueltoDatos As Boolean


Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TipoForPa As Byte 'tipo forma de pago: efectivo, banco,...
Dim codAlmac As Integer 'cod. almacen
Dim NomTraba As String 'nombre trabajador

Dim RSVenta As ADODB.Recordset


Dim SQL As String
Dim cadImpresion As String

'--- Variables generales para nueva impresión ticket (RAFA/ALZIRA 05092006)
Dim vNumTicket As String
Dim vNumAlbTicket As String

'   Octubre 2009            David
'   Puede que en el TPV se hayan quedado tickets de otros dias y se retomen con fecha posterior
'   Por ello, para la gneeracion de tik,alb y facturas sera fecha de hoy
Dim miFechaTicket As Date



'Mayo 2014
'Alzira. Hay formas de pago con recargo financiero
Dim PorceRecFinan As Currency
Dim ImporteFinal As Currency

Private Sub cmdCampos_Click()
       
    
   
   
    CadenaDesdeOtroForm = ""
    frmADVvarios.Opcion = 0
    frmADVvarios.vCampos = Text1(0).Text
    frmADVvarios.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        
        
            
        MultiInsercionCampos
        
        'Cargamos GRID
        
    End If
    CargaDatosCampos
End Sub

Private Sub cmdMtoCampos_Click(Index As Integer)
    If Index = 0 Then
        'Añadir mas campos
        cmdCampos_Click
        
    Else
        SQL = ""
        If Me.ListView1.ListItems.Count > 0 Then
            If Not Me.ListView1.SelectedItem Is Nothing Then
                SQL = "Va a eliminar el campo: "
                SQL = SQL & vbCrLf & "Codigo : " & Me.ListView1.SelectedItem.Text
                SQL = SQL & vbCrLf & "Partida : " & Me.ListView1.SelectedItem.SubItems(1)
                SQL = SQL & vbCrLf & "Variedad : " & Me.ListView1.SelectedItem.SubItems(2)
                SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                    'El tag tiene codcampo
                    SQL = "DELETE FROM sliven2 WHERE  fecventa = " & DBSet(RSVenta!fecventa, "F")
                    SQL = SQL & " AND numtermi = " & RSVenta!NumTermi & " AND numventa = " & RSVenta!NumVenta
                    SQL = SQL & " AND codcampo  = " & CStr(Val(Me.ListView1.SelectedItem.Text))
                    conn.Execute SQL
                    
                    Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
                    
                    PonerWidth
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdQUitarCliVar_Click()
    SQL = "Seguro que desea quitar datos del cliente de varios?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Text3.Text = ""
    Text3.Tag = ""
    Limpiarmanipulador
    cmdQUitarCliVar.visible = False
    ActualizarCliVariosEnBD
End Sub

Private Sub Limpiarmanipulador()
    If vParamAplic.ManipuladorFitosanitarios2 Then
        Me.txtManipulador(0).Text = "": Me.txtManipulador(1).Text = "": Me.txtManipulador(2).Text = ""
        Text2(3).Text = ""
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        If vParamTPV.Rapida Then
            If vParamAplic.ForPagoChequeRegalo = "" Then
                PonerFoco Text1(3)
            Else
                PonerFoco Text1(4)
            End If
        Else
            PonerFoco Text1(1)
        End If
        PrimeraVez = False
        PonerVisibleAsociacionCampos
        
        If vParamAplic.ManipuladorFitosanitarios2 Then
            If FrameCarnet.visible Then
                'Significa que lleva fitosanitarios. Veamos si tiene
                '
                If Val(Text1(0).Text) <> Val(vParamTPV.Cliente) Then
                    SQL = DevuelveDesdeBD(conAri, "ManipuladorNumCarnet", "sclien", "codclien", Text1(0).Text)
                    If SQL = "" Then
                        'Veo si tiene autirzados
                        SQL = DevuelveDesdeBD(conAri, "numcarnet", "sclienmani", "codclien", Text1(0).Text)
                    End If
                    
                    If SQL = "" Then
                        MsgBox "El cliente no tiene carnet de manipulador ni autorizados", vbExclamation
                    Else
                        mnManipulador_Click
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim Cad As String


'    If cadSel = "" Then Unload Me

    '1305,071
    



     'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'Icono de busqueda
    Me.imgBuscar(0).Picture = frmPpal.ImgListPpal.ListImages(17).Picture
    Me.imgBuscar(1).Picture = frmPpal.ImgListPpal.ListImages(17).Picture
    Me.imgBuscar(2).Picture = frmPpal.ImgListPpal.ListImages(17).Picture
    Me.imgBuscar(5).Picture = frmPpal.ImgListPpal.ListImages(17).Picture
    
    ' ICONITOS DE LA BARRA
    
    PonerVisibleAsociacionCampos
    
    With Me.Toolbar1
     
        .ImageList = frmPpal.ImgListPpal
        .Buttons(2).Image = 18   'Generar Ticket
        .Buttons(4).Image = 7   'Generar Albaran
        .Buttons(6).Image = 8   'Generar Factura
        
        .Buttons(8).Image = 3   'Clientes
        
        .Buttons(10).Image = 38   'manipulador
        
        .Buttons(12).Image = 14  'Salir
    End With
    
    miFechaTicket = Now 'FECHA DE HOY
    
    PrimeraVez = True
    CodTipoMov = "FTI" 'factura ticket
    
    SQL = "SELECT * FROM scaven "
    If cadSel <> "" Then SQL = SQL & " WHERE " & cadSel
    Set RSVenta = New ADODB.Recordset
    RSVenta.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
    'Almacen por defecto el del trabajador
    If RSVenta!CodTraba <> "" Then
        NomTraba = "nomtraba"
        codAlmac = ComprobarCero(DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", CStr(RSVenta!CodTraba), "N", NomTraba))
        If codAlmac = 0 Then codAlmac = DevuelveDesdeBDNew(conAri, "salmpr", "min(codalmac)", "", "")
    Else
        codAlmac = DevuelveDesdeBDNew(conAri, "salmpr", "min(codalmac)", "", "")
    End If
            
    SQL = DBLet(RSVenta!nifvarios, "T")
    If SQL = "" Then
        Text3.Text = "": Text3.Tag = "": cmdQUitarCliVar.visible = False
    Else
        PonerDatosClienteVarios
        
    End If
    
        
    Me.FrameCarnet.visible = False
    Limpiarmanipulador
    If LLevaArticulosFitosanitarios Then
        FrameCarnet.visible = True
        
        Me.txtManipulador(0).Text = DBLet(RSVenta!ManipuladorNumCarnet, "T")
        
        If Not IsNull(RSVenta!ManipuladorFecCaducidad) Then Me.txtManipulador(1).Text = Format(RSVenta!ManipuladorFecCaducidad, "dd/mm/yyyy")
        If DBLet(RSVenta!TipoCarnet, "N") > 0 Then Me.txtManipulador(2).Text = IIf(RSVenta!TipoCarnet = 2, "Cualificado", "Básico")
        Text2(3).Text = DBLet(RSVenta!ManipuladorNombre, "T")
    Else
        
    End If
        
    If vParamTPV.OfertaImporteEntregado Then
            Text1(3).Text = ImporteFinal
            PonerFormatoDecimal Text1(3), 1
    End If
    
    
    'Trabajador conectado
    Text1(2).Text = Format(RSVenta!CodTraba, "0000")
    Text2(2).Text = NomTraba
    
    
    'Poner el cliente de la venta
    If Not IsNull(RSVenta!codClien) Then
        Cad = "codforpa"
        Text1(0).Text = Format(RSVenta!codClien, "000000")
        Text2(0).Text = DevuelveDesdeBDNew(conAri, "sclien", "nomclien", "codclien", Text1(0).Text, "N", Cad) '(RAFA/ALZIRA 31082006)
        
        
        'FORMA DE PAGO
        
        Text1(1).Text = Format(CLng(Cad), "000")
        PonerformaDePago
        
        
        
        'departamento del cliente
        If Not IsNull(RSVenta!CodDirec) Then
            Text1(5).Text = Format(DBLet(RSVenta!CodDirec, "N"), "000")
            PonerDptoEnCliente
        Else
            Text1(5).Text = ""
            Text2(5).Text = ""
        End If
        
    Else
        'Poner el cliente que hay por defecto en los parametros
        Text1(0).Text = Format(vParamTPV.Cliente, "000000")
        Text2(0).Text = vParamTPV.NomCliente
        'Forma de pago por defecto (RAFA/ALZIRA 31082006)
        Text1(1).Text = Format(vParamTPV.ForPago, "000")
        Text2(1).Text = vParamTPV.NomForPago
        TipoForPa = vParamTPV.TipoForPago
        
        
        
    End If
    
    
    
    
    
    If vParamTPV.Rapida Then
        Me.Label2(4).Caption = "0.00" 'cambio
        Me.Text1(3).Text = "0.00" 'Entregado
        
        
        
    Else
        Me.Label2(4).Caption = "" 'cambio
        Me.Text1(3).Text = "" 'Entregado
    End If
    Me.Text1(4).Text = "" 'Cheque regalo
    
    
    'Campos
    Me.Height = 6030
    Me.FrameEfectivo.Top = 3480
    Me.cmdCampos.Top = 4440
    If LLevaArticulosFitosanitarios Then
        Me.Height = 7080
        Me.FrameEfectivo.Top = Me.FrameEfectivo.Top + 1090
        Me.cmdCampos.Top = Me.cmdCampos.Top + 1000
    End If
    'Cheque
    Label1(3).Top = Me.FrameEfectivo.Top + 120
    Text1(4).Top = Label1(3).Top + 392
    
    Me.mnManipulador.visible = LLevaArticulosFitosanitarios
    Toolbar1.Buttons(10).visible = LLevaArticulosFitosanitarios
    
    If vParamAplic.Ariagro <> "" Then
        Cad = "fecventa = " & DBSet(RSVenta!fecventa, "F")
        Cad = Cad & " AND numtermi = " & RSVenta!NumTermi & " AND numventa"
        Cad = DevuelveDesdeBD(conAri, "count(*)", "sliven2", Cad, CStr(RSVenta!NumVenta), "N")
        If Cad <> "" Then
            If Val(Cad) > 0 Then CargaDatosCampos
        End If
    End If
    
    
    
    Me.FrameEfectivo.Enabled = (TipoForPa = 0)
End Sub


Private Sub PonerVisibleAsociacionCampos()
    cmdCampos.visible = False
    Me.mnAsociarCampos.visible = False
    
    If Val(Text1(0).Text) <> vParamTPV.Cliente Then
        If vParamAplic.Ariagro <> "" Then
            cmdCampos.visible = True
            Me.mnAsociarCampos.visible = True
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    RSVenta.Close
    Set RSVenta = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'para busquedas
Dim I As Byte

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        I = CInt(Me.imgBuscar(0).Tag)
        
        Text1(I).Text = RecuperaValor(CadenaDevuelta, 1)
'        If i <> 5 Then Text1_LostFocus (i)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmClv_DatoSeleccionado(CadenaSeleccion As String)
    Text3.Text = RecuperaValor(CadenaSeleccion, 2) & "    (" & RecuperaValor(CadenaSeleccion, 1) & ")"
    Text3.Tag = RecuperaValor(CadenaSeleccion, 1)
    
    If vParamAplic.ManipuladorFitosanitarios2 Then
        SQL = "concat(ManipuladortipoCarnet,'|',ManipuladorNumCarnet,'|',fcaducidad,'|')    "
        SQL = DevuelveDesdeBD(conAri, SQL, "sclvar", "nifclien", Text3.Tag, "T")
        If SQL = "" Then
            Limpiarmanipulador
        Else
            txtManipulador(0).Text = RecuperaValor(SQL, 2)
            txtManipulador(1).Text = Format(RecuperaValor(SQL, 3), "dd/mm/yyyy")
            SQL = RecuperaValor(SQL, 1)
            SQL = IIf(SQL = "1", "Básico", IIf(SQL = "2", "Cualificado", ""))
            txtManipulador(2).Text = SQL
            Text2(3).Text = Text2(0).Text
        End If
    
    End If
    Me.cmdQUitarCliVar.visible = True
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    If RSVenta.EOF Then Exit Sub
    imgBuscar(0).Tag = Index
    MandaBusquedaPrevia CStr(Index)
'    If Index = 5 Then Text1_LostFocus (5)
End Sub





Private Sub mnAlbaran_Click()
'Pasamos la venta a una albaran de venta generado a partir de un ticket
'en los campos del pedido almacenamos de que ticket viene
Dim NumAlbaran As String

    '---- comprobar datos correctos
    If Not DatosOk(1) Then Exit Sub
    
    If Me.Text1(4).Text <> "" Then
        MsgBox "El cheque regalo no se puede utilizar en Albaranes", vbInformation
        Exit Sub
    End If
    
    ' ---- [21/10/2009] [LAURA] : centro de coste por trabajado
    If Not ActualizarCentroCoste Then Exit Sub
    
    
    'Mayo 2013
    If HayArticuloFitosanitario_O_BloqFamilia(False) Then Exit Sub
    
    
    '---- Generar el albaran y eliminar la venta
    CodTipoMov = "ALV" 'factura ticket
    If GenerarAlbaran(NumAlbaran) Then
        '---- Imprimir el Albaran
        NumAlbaran = "Se ha generado correctamente el Albaran de venta: " & NumAlbaran & vbCrLf
        If MsgBox(NumAlbaran & "¿Desea imprimirlo?  ", vbQuestion + vbYesNo) = vbYes Then ImprimirAlbaran
        
        'cerrar ventana total y regresar a entrada de ventas
        'cadSel = "1"
        Volver_A_Cargar_Datos = True
        
'        Unload Me
        mnSalir_Click
    End If
End Sub


Private Sub mnAsociarCampos_Click()
    If vParamAplic.Ariagro = "" Then Exit Sub
    If Val(Text1(0).Text) = vParamTPV.Cliente Then
        MsgBox "Cliente de varios", vbExclamation
        Exit Sub
    End If
    cmdCampos_Click
End Sub

Private Sub mnClienteVarios_Click()


        If Val(Text1(0).Text) <> vParamTPV.Cliente Then
            MsgBox "No es cliente de varios", vbExclamation
            Text3.Tag = ""
            Text3.Text = ""
            cmdQUitarCliVar.visible = False
            Exit Sub
        End If
        
        LanzarClientesVarios

End Sub

Private Sub mnFactura_Click()
Dim NumFactura As String
Dim bSalir As Boolean

    
    If Not DatosOk(2) Then Exit Sub
    
    ' ---- [21/10/2009] [LAURA] : centro de coste por trabajado
    If Not ActualizarCentroCoste Then Exit Sub
    
    'Mayo 2013
    If HayArticuloFitosanitario_O_BloqFamilia(False) Then Exit Sub
    
    bSalir = False
    Me.Toolbar1.Buttons(6).Enabled = False
    Me.mnFactura.Enabled = False
    
    CodTipoMov = "FAV" 'factura ticket
    
    Screen.MousePointer = vbHourglass
    
    'Si hay que meter en
    
    
    
    
    If GenerarFactura(NumFactura) Then
        'Lo pidio ANNA
        'Si es factura, pero paga en efectivo, abrimos cajon tambien
        SQL = DevuelveDesdeBD(conAri, "tipforpa", "sforpa", "codforpa", Text1(1).Text, "N")
        If SQL = "0" Then
            If vParamTPV.AbreCajon Then ImprimePorLaCom ""
        End If
    
    
    
        Screen.MousePointer = vbDefault
        'Imprimir la factura
        NumFactura = "Se ha generado correctamente la Factura: " & NumFactura & vbCrLf
        If MsgBox(NumFactura & "¿Desea imprimirla?  ", vbQuestion + vbYesNo) = vbYes Then ImprimirFactura
        'cerrar ventana total y regresar a entrada de ventas
        'cadSel = "1"
        Volver_A_Cargar_Datos = True
        
'        Unload Me
        bSalir = True
    End If
    Screen.MousePointer = vbDefault
    Me.Toolbar1.Buttons(6).Enabled = True
    Me.mnFactura.Enabled = True
    
    If bSalir Then mnSalir_Click
End Sub

Private Sub mnManipulador_Click()
    If Not LLevaArticulosFitosanitarios Then Exit Sub
    If Val(Text1(0).Text) = vParamTPV.Cliente Then
        MsgBox "Cliente varios", vbExclamation
        Exit Sub
    End If
    CadenaDesdeOtroForm = ""
    frmFitoCarnet.Cliente = Val(Text1(0).Text)
    frmFitoCarnet.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
            
            Me.txtManipulador(0).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Me.txtManipulador(1).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            Me.txtManipulador(2).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
            Text2(3).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
            
            SQL = "Update scaven SET ManipuladorNumCarnet =" & DBSet(txtManipulador(0).Text, "T")
            SQL = SQL & ",ManipuladorFecCaducidad =" & DBSet(txtManipulador(1).Text, "F")
            SQL = SQL & ",ManipuladorNombre =" & DBSet(Text2(3).Text, "T")
            'IIf(miRsAux!Tipo = 2, "Cualificado", "Básico")
            SQL = SQL & ", TipoCarnet = " & IIf(UCase(txtManipulador(2).Text) = "CUALIFICADO", 2, 1)
            SQL = SQL & " WHERE " & cadSel
            conn.Execute SQL
            Espera 0.5
            Me.Refresh
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


Private Sub mnTicket_Click()
Dim Impr As Boolean
Dim curEntregado As Currency
Dim curCambio As Currency


    If Not DatosOk(0) Then Exit Sub
    
    '## LAURA 20/06/2008
    '-- Comprobar que si existe un articulo con registro fitosanitario
    '-- no se puede hacer un ticket y salimos
    If HayArticuloFitosanitario_O_BloqFamilia(True) Then Exit Sub
    
    '##
    
    cadImpresion = ""
    CodTipoMov = "FTI" 'factura ticket
    curEntregado = 0
    curCambio = 0
    
    'Si contabilizamos los tickets agrupados, entonces NO podra generar el ticket si
    ' el cliente no es cliente varios O NO TIENE LA FORPA de parametros
    If vParamAplic.ContabilizarTicketAgrupados Then
    
        If Val(Text1(1).Text) <> vParamTPV.ForPago Then
            cadImpresion = DevuelveDesdeBD(conAri, "tipforpa", "sforpa", "codforpa", Text1(1).Text, "N")
            
            'Ni 1.- trnasferencia    4.-Recibo bancario   5.-Confirming
            If cadImpresion = "1" Or cadImpresion = "4" Or cadImpresion = "5" Then
                cadImpresion = "Forma de pago no puede ser ni TRANS/RECIBO/CONFIRMING.   FP:" & cadImpresion
                MsgBox cadImpresion, vbExclamation
                Exit Sub
            End If
              
        End If
    End If
    cadImpresion = ""
    
    ' ---- [21/10/2009] [LAURA] : centro de coste por trabajado
    If Not ActualizarCentroCoste Then Exit Sub
    
    
    
    If GenerarTicket Then
        
        
        

        Impr = True
        If Not vParamTPV.ImprimiDirecto Then
            If MsgBox("¿Desea imprimir el ticket?", vbQuestion + vbYesNo) = vbNo Then Impr = False
        End If
        'If Impr Then ImprimirTicketDirecto vNumTicket, vNumAlbTicket, RSVenta!fecventa
        
        '# Modificado: LAURA (25/07/2008)
        If Text1(3).Text <> "" Then curEntregado = CCur(Text1(3).Text)
        If Label2(4).Caption <> "" Then curCambio = CCur(Label2(4).Caption)
        If Impr Then ImprimirTicketDirecto vNumTicket, miFechaTicket, curEntregado, curCambio
        '#
        
        'cerrar ventana total y regresar a entrada de ventas
        'cadSel = "1"
        Volver_A_Cargar_Datos = True
        
'        Unload Me
        mnSalir_Click
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then BotonBuscar (Index)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim ImpCheque As Currency
Dim devuelve As String

    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    Select Case Index
        Case 0 'cod cliente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
                If Text2(Index).Text = "" Then
                    PonerFoco Text1(Index)
                ElseIf ClienteOK(Text1(Index), RSVenta!codClien, True) Then
                    
                    If Val(Text1(Index)) <> Val(RSVenta!codClien) Then
                        'Actualizo scaven
                        SQL = "UPDATE scaven SET codclien= " & Text1(Index).Text
                        SQL = SQL & ", coddirec=NULL"
                            'Fuerzo un null
                        SQL = SQL & ",ManipuladorNumCarnet=NULL,ManipuladorFecCaducidad =NULL ,ManipuladorNombre =  NULL , TipoCarnet = NULL"
                        SQL = SQL & " WHERE " & cadSel
                        
                        ejecutar SQL, False
                    End If
                    If Text1(1).Text = "" Then
                        'recuperar la forma de pago del cliente
                        SQL = DevuelveDesdeBD(conAri, "codforpa", "sclien", "codclien", Text1(Index).Text, "N")
                        Text1(1).Text = SQL
                        Text1_LostFocus (1)
                    End If
                Else
                    Text1(Index).Text = RSVenta!codClien
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", "cliente", "N")
                    Text1(Index).Text = Format(Text1(Index).Text, "000000")
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            PonerVisibleAsociacionCampos
        Case 1 'cod forpa
            If PonerFormatoEntero(Text1(Index)) Then
                Text1(Index).Text = Format(Text1(Index).Text, "000")
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa", "codforpa", "Forma de pago", "N")

                'MAYO 2014
                PonerformaDePago

                
                If Text2(Index).Text = "" Then
                    PonerFoco Text1(Index)
                Else
                    Me.FrameEfectivo.Enabled = (TipoForPa = 0)
                    If TipoForPa <> 0 Then
                        Me.Label2(4).Caption = ""
                        Me.Text1(3).Text = ""
                    Else
                        'Forpa correcta. SI NO tiene checque regalo lo posicionamos
                        If vParamAplic.ForPagoChequeRegalo = "" Then PonerFoco Text1(3)
                        
                        
                    End If
                   
                End If
            Else
                If Text2(Index).Text <> "" Then
                    Text2(Index).Text = ""
                    PonerFoco Text1(1)
                Else
'                    PonerFoco Text1(1)
                End If
            End If
        
        Case 2 'cod trabajador
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba", "Operador", "N")
            Text1(Index).Text = Format(Text1(Index).Text, "0000")
            
        Case 3 'Entregado
            If PonerFormatoDecimal(Text1(Index), 1) Then
                'obtener el importe del cheque regalo si hay
                ImpCheque = CCur(ComprobarCero(Text1(4).Text))
                'Obtener el cambio= entregado + cheque_regalo - importe
                Label2(4).Caption = Format(CCur(Text1(Index).Text) + ImpCheque - CCur(ImporteFinal), FormatoImporte)
                If Text1(1).Text = "" Then PonerFoco Text1(1) 'Si no ha puesto la forma de pago.. que la ponga
            Else
                Label2(4).Caption = ""
            End If
            frmFacTPVEnt.EnviarVisorPuerto Label2(3).Caption, Label2(4).Caption, Label2(0).Caption, Label2(1).Caption
            
        Case 4 'cheque regalo
             If PonerFormatoDecimal(Text1(Index), 1) Then
                'If Me.Text1(3).Enabled = False Then PonerFoco Text1(1)
             End If
             
        Case 5 'DIREC./DPTO.
             If PonerFormatoEntero(Text1(Index)) Then
                Text1(Index).Text = Format(Text1(Index).Text, "000")
                'Comprobar que el cliente seleccionado tiene esa direccion
                If PonerDptoEnCliente Then
                    'Comprobar que el cliente tiene mantenimientos en esa direc/dpto
                    devuelve = DevuelveDesdeBDNew(conAri, "scaman", "nummante", "codclien", Text1(0).Text, "N", , "coddirec", Text1(5).Text, "N")
                    If devuelve <> "" Then MsgBox "El cliente tiene Mantenimientos.", vbInformation
                Else
                    Text1(Index).Text = ""
                    Text2(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
             Else
                Text2(Index).Text = ""
'                PonerFoco Text1(Index)
             End If
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 3, cerrar
    If KeyAscii = 27 Then cerrar = True
    If cerrar Then Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 'Generar ticket
            mnTicket_Click
        Case 4 'Generar Albaran
            mnAlbaran_Click
        Case 6 'Generar Factura
            mnFactura_Click
        
        Case 8
            mnClienteVarios_Click
        Case 10
            mnManipulador_Click
        Case 12 'Salir
            mnSalir_Click
    End Select
End Sub



Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, ByRef Rs As ADODB.Recordset) As Boolean
'On Error Resume Next
On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.HoraMov = CStr(Rs!horventa)
    vCStock.codArtic = Rs!codArtic
    vCStock.codAlmac = codAlmac
    vCStock.cantidad = CSng(Rs!cantidad)
    '16 Mayo 08
    '----------
    ' El importe de la linea esta en una columna de la BD
    'vCStock.Importe = CCur(RS!Cantidad) * CCur(RS!precioar)
    vCStock.Importe = Rs!implineareal
    vCStock.LineaDocu = CInt(Rs!numlinea)
        
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function


Private Function ObtenerContadorTicket(NumTicket As String) As Boolean
Dim vTipoMov As CTiposMov

    On Error Resume Next

    CodTipoMov = "FTI"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        NumTicket = vTipoMov.ConseguirContador(CodTipoMov)
        If NumTicket <> "-1" Then ObtenerContadorTicket = True
        
        vTipoMov.IncrementarContador (CodTipoMov)
    Else
        ObtenerContadorTicket = False
    End If
    Set vTipoMov = Nothing
    
    If Err.Number <> 0 Then ObtenerContadorTicket = False
End Function



'01/09/06 Laura
Private Function ObtenerContadorAlbTicket(NumAlbTicket As String) As Boolean
Dim vTipoMov As CTiposMov

    On Error Resume Next

    CodTipoMov = "ATI"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        NumAlbTicket = vTipoMov.ConseguirContador(CodTipoMov)
        If NumAlbTicket <> "-1" Then ObtenerContadorAlbTicket = True
        
        vTipoMov.IncrementarContador (CodTipoMov)
    Else
        ObtenerContadorAlbTicket = False
    End If
    Set vTipoMov = Nothing
    
    If Err.Number <> 0 Then ObtenerContadorAlbTicket = False
End Function



Private Function ObtenerContadorAlbaran(NumAlb As String) As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean

    On Error GoTo ErrConAlb

    CodTipoMov = "ALV"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Do
            NumAlb = vTipoMov.ConseguirContador(CodTipoMov)
            vTipoMov.IncrementarContador (CodTipoMov)
            SQL = "select count(*) from scaalb where codtipom='" & CodTipoMov & "' and numalbar=" & NumAlb
            Existe = (RegistrosAListar(SQL) > 0)
        Loop Until Existe = False
        ObtenerContadorAlbaran = True
    Else
        ObtenerContadorAlbaran = False
    End If
    Set vTipoMov = Nothing
    Exit Function
    
ErrConAlb:
    ObtenerContadorAlbaran = False
    MuestraError Err.Number, "Obtener contador albaran", Err.Description
End Function




Private Function ObtenerContadorFactura(NumFactu As String) As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean

    On Error GoTo ErrConFac

    CodTipoMov = "FAV"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Do
            NumFactu = vTipoMov.ConseguirContador(CodTipoMov)
            vTipoMov.IncrementarContador (CodTipoMov)
            SQL = "select count(*) from scafac where codtipom='" & CodTipoMov & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(miFechaTicket, "F")
            Existe = (RegistrosAListar(SQL) > 0)
        Loop Until Existe = False
        ObtenerContadorFactura = True
    Else
        ObtenerContadorFactura = False
    End If
    Set vTipoMov = Nothing
    Exit Function
    
ErrConFac:
    ObtenerContadorFactura = False
    MuestraError Err.Number, "Obtener contador factura", Err.Description
End Function



Private Function InsertarMovAlmacen(NumTicket As String) As Boolean
'PAra tickets, albaranes y facturas
Dim Rs As ADODB.Recordset
Dim vCStock As CStock
Dim b As Boolean
Dim ErroresEnStock As String

    On Error GoTo EInsMov
    
    'Para cada linea de venta insertar el movimiento e actualizar stocks
    Set Rs = New ADODB.Recordset
    Set vCStock = New CStock
    
    SQL = Replace(cadSel, "scaven", "sliven")
    SQL = "SELECT * FROM sliven WHERE " & SQL
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    '26 Abril 2011
    'El documento ira formateado con ceros. Como si fuera en la entrada de albaran
    'vCStock.Documento = NumTicket
    vCStock.Documento = Format(NumTicket, "0000000")
    vCStock.DetaMov = CodTipoMov
    vCStock.Trabajador = CLng(Text1(0).Text) 'sera el cliente
    vCStock.FechaMov = CStr(miFechaTicket)
    
    b = True
    ErroresEnStock = ""
    
    'En la funcion muevestock tiene los avisioss sobre las cantidades y sobre
    'maximos minimos y puntos de pedido
    While Not Rs.EOF And b
        If Not InicializarCStock(vCStock, "S", Rs) Then Exit Function
        
        'Para que compruebe las cantidades y eso
        If vParamTPV.CtrstockVenta Then
            If vCStock.MueveStock Then
                b = vCStock.MoverStock(False, True, False)    'True en actualizar DB
            Else
                b = True
            End If
        End If
        If Not vCStock.ActualizarStock(True, False) Then b = False
        Rs.MoveNext
    Wend
    Rs.Close
    
    'Ahora REstaremos en los lotes
    If vParamAplic.ManipuladorFitosanitarios2 Then
        SQL = Replace(cadSel, "scaven", "slivenlotes")
        SQL = "SELECT * FROM slivenlotes WHERE " & SQL
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            If Rs!cantidad < 0 Then
                SQL = "-"
            Else
                SQL = "+"
            End If
            SQL = "UPDATE slotes SET vendida=vendida " & SQL & DBSet(Abs(Rs!cantidad), "N")
            SQL = SQL & " WHERE numlotes= " & DBSet(Rs!numLote, "T")
            SQL = SQL & " AND codartic= " & DBSet(Rs!codArtic, "T")
            SQL = SQL & " AND fecentra= " & DBSet(Rs!fecentra, "F")
            conn.Execute SQL
            
            Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    Set vCStock = Nothing
    
    Set Rs = Nothing
    
EInsMov:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando movimientos de almacen.", Err.Description
        b = False
        Set vCStock = Nothing
        Rs.Close
        Set Rs = Nothing
    End If
    InsertarMovAlmacen = b
End Function



Private Function InsertarHistFactura(NumTicket As String, Optional NumFactu As String, Optional NumAlbTicket As String, Optional MenError As String) As Boolean
Dim b As Boolean
Dim vFactu As CFactura
Dim vClien As CCliente
Dim DatosDelClienteVarios As String
Dim RT As ADODB.Recordset
Dim UpdatesNumlotes As String
Dim I As Integer
    On Error GoTo EInsFac
    
    
    
    
    'Preparo los numeros de lote, ya que al crear la factura se borra la venta
    UpdatesNumlotes = ""
    
    Set RT = New ADODB.Recordset
    SQL = "Select * from sliven,slotes WHERE sliven.codartic =slotes.codartic AND " & Replace(cadSel, "scaven", "sliven")
    SQL = SQL & " ORDER BY numlinea,fecentra desc"
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""  'guardaremos la linea
    While Not RT.EOF
        If SQL <> RT!numlinea Then
            'Aticulo nuevo. La primera entrada es la que vale
            SQL = "UPDATE slifac SET numlote = " & DBSet(RT!numlotes, "T") & " WHERE numlinea = " & RT!numlinea
            UpdatesNumlotes = UpdatesNumlotes & SQL & "|"
            SQL = RT!numlinea
        End If
        
        
        
        RT.MoveNext
    Wend
    RT.Close
    Set RT = Nothing
    
    SQL = ""
    'Insertar la cabecera de Factura (scafac)
    Set vFactu = New CFactura
    If NumFactu = "" Then
        vFactu.codtipom = "FTI"
        vFactu.NumFactu = NumTicket
    Else
        vFactu.codtipom = "FAV"
        vFactu.NumFactu = NumFactu
    End If
    vFactu.FecFactu = Format(miFechaTicket, "dd/mm/yyyy")
    vFactu.NumTerminal = RSVenta!NumTermi
    vFactu.NumVenta = RSVenta!NumVenta
    
    
    'Guardamos los valores identificativos de la factura generada
    'para imprimirla posteriormente
    cadImpresion = "{scafac.codtipom}='" & vFactu.codtipom & "' and {scafac.numfactu}=" & vFactu.NumFactu

    vFactu.Cliente = Text1(0).Text
    vFactu.DirDpto = Text1(5).Text
    vFactu.NombreDirDpto = Text2(5).Text
    If vFactu.Cliente <> "" Then
        Set vClien = New CCliente
        If vClien.LeerDatos(vFactu.Cliente) Then
            
            'Si es cliente varios
            DatosDelClienteVarios = ""
            If Val(Text1(0).Text) = vParamTPV.Cliente Then
                'SI ha metido los datos del cliente de varios , o no
                If Text3.Tag <> "" Then DatosDelClienteVarios = LeerDesdeTablaClienteVarios(vFactu)
            End If
            'si es "" entonces o NO es de varios, o siendo de varios NO ha leido los datos
            If DatosDelClienteVarios = "" Then
                vFactu.NombreClien = vClien.Nombre
                vFactu.DomicilioClien = vClien.Domicilio
                vFactu.CPostal = vClien.CPostal
                vFactu.Poblacion = vClien.Poblacion
                vFactu.Provincia = vClien.Provincia
                vFactu.NIF = vClien.NIF
                vFactu.Telefono = vClien.TfnoClien
            End If
            
            vFactu.Agente = vClien.Agente
            vFactu.Banco = vClien.Banco
            vFactu.Sucursal = vClien.Sucursal
            vFactu.DigControl = vClien.DigControl
            vFactu.CuentaBan = vClien.CuentaBan
            'Actualizamos fecha ult. movim del cliente si es posterior
            b = vClien.ActualizaUltFecMovim(vFactu.FecFactu)
        Else
            InsertarHistFactura = False
            Exit Function
        End If
        Set vClien = Nothing
    End If
    'obtener letra serie de la factura (para el tipo de movimiento)
    SQL = DevuelveDesdeBDNew(conAri, "stipom", "letraser", "codtipom", vFactu.codtipom, "T")
    vFactu.LetraSerie = SQL
    
    vFactu.ForPago = Text1(1).Text
    vFactu.TipForPago = TipoForPa
    vFactu.TotalFac = ImporteFinal
     
    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = vParamTPV.CtaPrevCobro
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", vFactu.BancoPr, "N")
   
    If vFactu.CuentaPrev = "" Then
        b = False
        SQL = "La cuenta prevista de cobro no puede ser nula. Parámetos TPV."
    End If
    
    If b Then
        SQL = Text1(2).Text 'Trabajador
        b = b And vFactu.PasarTicketAFactura(cadSel, SQL, NumTicket, NumAlbTicket, Text1(4).Text)
    End If
    
    
    
    If b Then
        'Actualizamos si lleva articulos fitosanitarios
        While UpdatesNumlotes <> ""
            I = InStr(1, UpdatesNumlotes, "|")
            If I = 0 Then
                UpdatesNumlotes = ""
            Else
                SQL = Mid(UpdatesNumlotes, 1, I - 1)
                UpdatesNumlotes = Mid(UpdatesNumlotes, I + 1)
        
                'Hacemos el update
                SQL = SQL & " AND codtipom = '" & vFactu.codtipom & "' AND numfactu = " & vFactu.NumFactu
                SQL = SQL & " AND fecfactu = " & DBSet(miFechaTicket, "F")
                conn.Execute SQL
 
                
            End If
        Wend
    End If
    
    
    
'    If Not b Then MsgBox SQL, vbInformation
    If Not b Then MenError = SQL
    Set vFactu = Nothing
    
EInsFac:
    If Err.Number <> 0 Then
        'MuestraError Err.Number, "Insertando Histórico de Factura.", Err.Description
        MenError = "Insertando Histórico de Factura." & vbCrLf & Err.Description
        b = False
    End If
    InsertarHistFactura = b
End Function



Private Function InsertarAlbaran(NumAlb As String, NumTicket As String, menErr As String) As Boolean
Dim b As Boolean
Dim vClien As CCliente
Dim DatosDelClienteVarios As String
    On Error GoTo EInsAlb

    'Cabecera de albaran
    '----------------------------------
    SQL = "INSERT INTO scaalb (codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    SQL = SQL & "coddirec,nomdirec,referenc,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,"
    SQL = SQL & "observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa "
    'Octubre 2015
    SQL = SQL & ",ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet) "
    'Abril 2008
    'Pongo la marca de facturar a TRUE: 1
    SQL = SQL & " VALUES ('" & CodTipoMov & "'," & NumAlb & "," & DBSet(miFechaTicket, "F") & ",1," & Text1(0).Text & ","
    
    'Obtenemos los datos del cliente
    Set vClien = New CCliente
    If vClien.Existe(Text1(0).Text) Then
        If vClien.LeerDatos(Text1(0).Text) Then
        
            'Si es cliente varios
            DatosDelClienteVarios = ""
            If Val(Text1(0).Text) = vParamTPV.Cliente Then
                'SI ha metido los datos del cliente de varios , o no
                If Text3.Tag <> "" Then DatosDelClienteVarios = LeerDesdeTablaClienteVarios(Nothing)
            End If
            
            'Si no es clietne varios, o no ha metido los datos del cliente de varios
            If DatosDelClienteVarios = "" Then
                SQL = SQL & DBSet(vClien.Nombre, "T", "N") & ", " & DBSet(vClien.Domicilio, "T", "N") & ","
                SQL = SQL & DBSet(vClien.CPostal, "T", "N") & ", " & DBSet(vClien.Poblacion, "T", "N") & "," & DBSet(vClien.Provincia, "T", "N") & ","
                SQL = SQL & DBSet(vClien.NIF, "T", "N") & "," & DBSet(vClien.TfnoClien, "T")
            Else
                SQL = SQL & DatosDelClienteVarios
            End If
            SQL = SQL & "," & DBSet(Text1(5).Text, "N", "S") & "," & DBSet(Text2(5).Text, "T") & "," & ValorNulo & "," 'coddirec,nomdirec,referenc a nulo
            SQL = SQL & Text1(2).Text & "," & Text1(2).Text & "," & Text1(2).Text & "," 'trabajador
            SQL = SQL & vClien.Agente & "," & Text1(1).Text & "," & vClien.FEnvio & ",0,0," & vClien.TipoFactu & ","
            'observaciones
            'La primera observacion sera el campo de observaciones de la venta
            If IsNull(RSVenta!observa1) Then
                SQL = SQL & ValorNulo
            Else
                SQL = SQL & "'" & DevNombreSQL(RSVenta!observa1) & "'"
            End If
            SQL = SQL & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            'datos oferta: aqui guardamos nº venta
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
            'En los campos de datos del pedido guardamos los datos del ticket
            SQL = SQL & NumTicket & "," & DBSet(miFechaTicket, "F") & "," & ValorNulo & "," & ValorNulo & ",1," & DBSet(RSVenta!NumTermi, "N") & "," & DBSet(RSVenta!NumVenta, "N", "S")  'esticket=1, terminal
            
            'Octubre 2015
            If FrameCarnet.visible Then
                'Lleva FITO txtManipulador
                'ManipuladorNumCarnet,ManipuladorFecCaducidad,ManipuladorNombre,TipoCarnet "
                SQL = SQL & "," & DBSet(txtManipulador(0).Text, "T", "S") & "," & DBSet(txtManipulador(1).Text, "F", "S") & ","
                SQL = SQL & DBSet(Text2(3).Text, "T", "S") & "," & IIf(UCase(txtManipulador(2).Text) = "CUALIFICADO", 2, 1) & ")"
            Else
                SQL = SQL & ",NULL,NULL,NULL,NULL)"
            End If
            b = vClien.ActualizaUltFecMovim(CStr(miFechaTicket))
        Else
            b = False
        End If
    End If
    Set vClien = Nothing
    
    
    If b Then
        'Insertar Cabecera
'    MenError = "Error al insertar en la tabla Cabecera de Albaranes (scaalb )."
        conn.Execute SQL, , adCmdText
        
        'Lineas del albaran
        'Inserta en tabla "slialb" todas las lineas de venta
        SQL = "INSERT INTO slialb "
        SQL = SQL & "(codtipom, numalbar,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad,"
        ' -- [19/10/2009] [LAURA] : añadir centro de coste
        SQL = SQL & "precioar, dtoline1, dtoline2, importel, origpre,codprovex,codccost,numlote) "
        
        
        
        
        'Neuvo Abril 2008. David
        'Para llevar el codprove a la linea de albaran y que no ponga el 0
        SQL = SQL & " SELECT '" & CodTipoMov & "' as codtipom," & DBSet(NumAlb, "N") & " as numalbar," & "numlinea," & codAlmac & " as codalmac,"
        SQL = SQL & " sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,"
        If Not vParamTPV.CalculaIVAsobrePVP Then
            'COMO estaba ANTES
            SQL = SQL & "precioar "
        Else
            'Nuevo ENERO 2010
            'Calcula sobre PVP
            SQL = SQL & "impartalb"
        End If
        SQL = SQL & ",dto1 as dtoline1,dto2 as dtoline2,"
        'NUEVO###
        'David.    La linea puede llevar dtos, con lo cual hay un
        '           campo en sliven que lleva el importe real de la linea
        ' ANTES:  round(cantidad*precioar,2) as importel
        SQL = SQL & " implineareal as importel,'' as origpre ,codprove,codccost,"
        
        'Si es manipulador fitosantir grabaremos lo que hay
        'es decir. * si lleva lotes   NUll si no
        'Si no es fito grabaremos un NULL, ya que actualizamos luego
        If vParamAplic.ManipuladorFitosanitarios2 Then
            SQL = SQL & "numlote"
        Else
            SQL = SQL & "NULL"
        End If
        
        SQL = SQL & " FROM sliven,sartic WHERE sliven.codArtic = sartic.codArtic AND " & Replace(cadSel, "scaven", "sliven")
        
        
        conn.Execute SQL, , adCmdText
        
        
        
        
        'UPdateamos la columna de numlotes del albaran si lleva fitosanitarios
        Dim RT As ADODB.Recordset
        
        
        If Not vParamAplic.ManipuladorFitosanitarios2 Then
            'ESTO ES LO QUE HACIA ANTES
            'Es decir. U n numero de serie por linea, cogiod por orden del slotes
            Set RT = New ADODB.Recordset
            SQL = "Select * from sliven,slotes WHERE sliven.codartic =slotes.codartic AND " & Replace(cadSel, "scaven", "sliven")
            SQL = SQL & " ORDER BY numlinea,fecentra desc"
            RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""  'guardaremos la linea
            While Not RT.EOF
                If SQL <> RT!numlinea Then
                    'Aticulo nuevo. La primera entrada es la que vale
                    SQL = "UPDATE slialb SET numlote = " & DBSet(RT!numlotes, "T") & " WHERE codtipom= '" & CodTipoMov & "' AND numalbar= " & NumAlb
                    SQL = SQL & " AND numlinea = " & RT!numlinea
                    conn.Execute SQL
                    SQL = RT!numlinea
                End If
                
                
                
                RT.MoveNext
            Wend
            RT.Close
            Set RT = Nothing
        
       Else
            'Nuevo. Nueva tabla de slialblotes. Con lo cual, aquellos articulos que lleven lotes estaran en slialblotes
            'Con lo caul, en este punto no hay que hacer nada
            'Ya que en slialb , la columna numotes esta marcada con un * para indicar que SI tiene lotes
       
       
       
       End If
        
        
        
        
        
        
    End If

     
    'Lineas de los campos
     If b Then
        If vParamAplic.Ariagro <> "" Then
            'pasamps los campos asignados
            SQL = "INSERT INTO slialbcampos(codtipom, numalbar,numlinea, codcampo) "
            SQL = SQL & " SELECT '" & CodTipoMov & "' as codtipom," & DBSet(NumAlb, "N") & " as numalbar," & "numlinea,codcampo"
            SQL = SQL & " FROM sliven2 WHERE " & Replace(cadSel, "scaven", "sliven2")
            conn.Execute SQL
            
        End If
        
        If vParamAplic.ManipuladorFitosanitarios2 Then
            'Llevamos tambien "control"lotes
            SQL = "INSERT INTO slialblotes(codtipom,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic)"
            SQL = SQL & " SELECT '" & CodTipoMov & "' as codtipom," & DBSet(NumAlb, "N") & " as numalbar," & "numlinea"
            SQL = SQL & ",sublinea,cantidad,numlote,fecentra,codartic"
            SQL = SQL & " FROM slivenlotes WHERE " & Replace(cadSel, "scaven", "slivenlotes")
            conn.Execute SQL
        End If
    End If
     
    'Eliminar las ventas que se han pasado a albaranes
    If b Then b = EliminarVenta(cadSel)
    
    'Guardamos los valores identificativos de la factura generada
    'para imprimirla posteriormente
    If b Then cadImpresion = "{scaalb.codtipom}='" & CodTipoMov & "' and {scaalb.numalbar}=" & DBSet(NumAlb, "N")

EInsAlb:
    If Err.Number <> 0 Then
        menErr = "Insertando el Albaran: " & vbCrLf & Err.Description
        b = False
    End If
    InsertarAlbaran = b
End Function


'0: tiket   1: Albaran    2:Factura
Private Function DatosOk(Destino As Byte) As Boolean
Dim b As Boolean
Dim I As Byte
Dim Cad As String
Dim Rs As ADODB.Recordset

    On Error GoTo EDatosOK
    b = True
    
    'Comprobaciones
    '------------------
    
    'comprobar que los campos tienen valor
    For I = 0 To 2
        If Trim(Me.Text1(I).Text) = "" Then
            If I = 0 Then
                Cad = "Cliente"
            ElseIf I = 1 Then
                Cad = "Forma de pago"
            ElseIf I = 2 Then
                Cad = "Operador"
            End If
            MsgBox "El campo " & Cad & " debe tener valor.", vbInformation
            b = False
            Exit For
        End If
    Next I
    
    'comprobar que el trabajador existe
    If b Then
        If DevuelveDesdeBDNew(conAri, "straba", "codtraba", "codtraba", Text1(2).Text, "N") = "" Then
            b = False
            MsgBox "No existe el trabajador " & Text1(2).Text, vbExclamation
        End If
    End If
    
    
    '--- Laura: 11/04/2007
    '--- comprobar q el cliente no esta bloqueado y q si se ha cambiado sea de la mista
    '--- tarifa q para el q se insertaron las lineas
    If b Then
        b = ClienteOK(Text1(0), RSVenta!codClien, False)
        If Not b Then Text1(0).Text = RSVenta!codClien
    End If
    '---
    
    
    '--- Laura: 12/04/2007
    '--- comprobar q si es cliente contado el tipo de forma de pago sea efectivo
    If b Then
        'obtenemos tipoforpa correcta por si acaso
        Cad = DevuelveDesdeBD(conAri, "tipforpa", "sforpa", "codforpa", Text1(1).Text, "N")
        If Cad = "" Then
            b = False
            MsgBox "No existe la forma de pago.", vbExclamation
        Else
            TipoForPa = CByte(Cad)
        
            'si se ha definido un cliente como contado en parametros del TPV
            If vParamTPV.Cliente <> "" Then
                If CLng(Text1(0).Text) = CLng(vParamTPV.Cliente) Then 'si es cliente definido como CONTADO
                    'Aceptamos EFECTIVO y  TARJETA DE CREDITO
                    If TipoForPa <> 0 And TipoForPa <> 6 Then 'tiene q tener tipo forpa EFECTIVO or TARJ CREDIT
                        b = False
                        MsgBox "El cliente '" & Text2(0).Text & "' debe tener una Forma de Pago de tipo EFECTIVO.", vbExclamation
                    End If
                End If
            End If
            
            aqui aqui aqui
            
            'Mayo 2014
            'Si hay recargo financiero, y es la empresa ALZIRA, entonces
            'Comprobaremos si las suma de lo que nos debe(con las formas de pago con recfinan)
            If vParamTPV.FormaDePagoConRegargoFinanciero(CInt(Text1(1).Text)) Then
            
                'Las que tengan forpa de pago con recargo financiero NO se pueden hacer
                If Destino = 0 Then
                    'NO se puede hacer tickets con formas de pago ..
                    MsgBox "No se pueden hacer tickets con formas de pago con recargo financiero", vbExclamation
                    b = False
                Else
                    If Not ComprobarTotalPendienteFormasPagoRecFinan(CLng(Me.Text1(0).Text), vParamTPV.FormaDePagoConRecFinan_SQL, ImporteFinal) Then b = False
                End If
            End If
        End If
    End If
    '---
    
    
    If b Then
        If TipoForPa = 0 Then 'Contado
            If Me.Text1(3).Text = "" Then
                MsgBox "Debe introducir la cantidad a pagar.", vbInformation
                b = False
            Else
                If Not vParamTPV.Rapida Then
            
                    If (CCur(Me.Text1(3).Text) + CCur(ComprobarCero(Text1(4).Text))) < CCur(Me.Label2(1).Caption) Then
                        MsgBox "La cantidad entregada debe ser igual o superior al importe total.", vbInformation
                        b = False
                    End If
                End If
            End If
        ElseIf TipoForPa = 4 Then 'Recibo
            'comprueba que el cliente tenga cuenta bancaria OK sino
            'muestra aviso pero deja pasar
            ComprobarCtaBanCliente
        End If
    End If
    
    
    '--- Laura: 18/12/2006
    'direc./dpto del cliente
    If b And Text1(5).Text <> "" Then
        'comprobar q existe el dpto para el cliente
        b = PonerDptoEnCliente
    End If
    
    
    '--- Laura: 01/12/2006
    'si hay cheque regalo
    If b Then
        If Me.Text1(4).Text <> "" Then
            'comprobar q en parametros de la aplicacion el campo codforpa tiene valor
            If vParamAplic.ForPagoChequeRegalo = CCur(Me.Label2(1).Caption) Then
                MsgBox "No se ha introducido la forma de pago del cheque regalo." & vbCrLf & "Configurar parámetros aplicación.", vbInformation, "Comprobar datos"
                b = False
            End If
            'comprobar que el importe del cheque sea >= q total factura
            If CCur(Me.Text1(4).Text) > CCur(Me.Label2(1).Caption) Then
                MsgBox "El importe del cheque regalo no puede ser superior al TOTAL.", vbExclamation
                b = False
            End If
        End If
    End If
    
    
    ' ---- [21/10/2009] [LAURA] : añadir centro de costes para contab. analitica
    'Modifica DAVID. Si anal=1 or 2(proyecto)
    If b And vEmpresa.TieneAnalitica And vParamAplic.ModoAnalitica = 1 Then

        'si hay analitica  por familia=1, si es por trabajador=0 se comprueba en form de total TPV
        'comprobar q las lineas de venta tienen centro de coste
        Cad = "SELECT codartic FROM sliven WHERE " & Replace(cadSel, "scaven", "sliven")
        Cad = Cad & " and isnull(codccost)"
        
        Set Rs = New ADODB.Recordset
        Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs.EOF Then
            b = False
            MsgBox "La familia del artículo " & DBLet(Rs!codArtic, "T") & " no tiene asignado centro de coste.", vbExclamation
        End If
        Rs.Close
        Set Rs = Nothing
    End If
    
    If b Then
        'Va para tiket.  Si tienen el aviso y el cliente es distinto de
        If vParamTPV.AvisoGeneraFactura Then
            Cad = ""
            If Destino = 0 Then
                If Val(Text1(0).Text) <> Val(vParamTPV.Cliente) Then Cad = "Va a realizar un ticket a un cliente"
                
            Else
                'Fra o albaran
                If Val(Text1(0).Text) = Val(vParamTPV.Cliente) Then Cad = "Va a realizar un factura/albaran a un cliente genérico"
            End If
            If Cad <> "" Then
                Cad = Cad & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then b = False
            End If
        End If
    End If
    DatosOk = b
    Exit Function
    
EDatosOK:
    MuestraError Err.Number, "Comprobando datos.", Err.Description
    DatosOk = False
End Function



Private Function GenerarTicket() As Boolean
Dim b As Boolean
Dim NumTicket As String
'01/09/06
Dim NumAlbTicket As String
Dim MenError As String

    On Error GoTo ETicket
    
    conn.BeginTrans
    'si el tipo de forma de pago no es efectivo habrá que insertar
    'en la tabla de contabilidad conta.scobro
'    If TipoForPa <> 0 Then
    ConnConta.BeginTrans
    
    
    
        
        
    'Obtener el contador de ticket (FTI).
    b = ObtenerContadorTicket(NumTicket)
    
    'Obtener el contador albaran de ticket (ATI).
    If b Then b = ObtenerContadorAlbTicket(NumAlbTicket)
    
    If b Then
        'Actualizar los stocks de todos los articulos comprados
        'Insertar movimiento en smoval
        b = InsertarMovAlmacen(NumAlbTicket)
        If Not b Then MenError = "Control stock"
        
        'Insertar en el historico de facturas: scafac, scafac1,slifac
        'en el campo scafac1.numalbar guardamos el nº de ticket
        If b Then
            b = InsertarHistFactura(NumTicket, , NumAlbTicket, MenError)
            
        End If
    End If
    vNumTicket = NumTicket ' (RAFA/ALZIRA 05092006)
    vNumAlbTicket = NumAlbTicket ' (RAFA/ALZIRA 05092006)
    
ETicket:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    If b Then
        conn.CommitTrans
        'If TipoForPa <> 0 Then
        ConnConta.CommitTrans
    Else
        conn.RollbackTrans
        'If TipoForPa <> 0 Then
        ConnConta.RollbackTrans
        MsgBox "ERROR: " & vbCrLf & MenError, vbExclamation, "Generar Ticket"
    End If
    GenerarTicket = b
    TerminaBloquear
    Espera 0.2
End Function



Private Function GenerarAlbaran(NumAlb As String) As Boolean
'La venta se combierte en un albaran.
Dim b As Boolean
Dim NumTicket As String
Dim MenError As String

    On Error GoTo EAlbar
    conn.BeginTrans
   
    'Obtener el contador de ticket (FTI).
    b = ObtenerContadorAlbTicket(NumTicket)
    
    If b Then
        'Obtener el contador de Albaran (ALV).
        b = ObtenerContadorAlbaran(NumAlb)
        
        If b Then
            If PorceRecFinan > 0 Then InsertarEnSliven
        
        
            'Actualizar los stocks de todos los articulos comprados
            'Insertar movimiento en smoval
            b = InsertarMovAlmacen(NumAlb)
            If Not b Then MenError = "Control stock"
            'Insertar en las tablas de Albaranes: scaalb, slialb
            'en el campo scafac1.numalbar guardamos el nº de ticket
            If b Then b = InsertarAlbaran(NumAlb, NumTicket, MenError)
        End If
    End If
    
EAlbar:
    If Err.Number <> 0 Then
        MenError = MenError & vbCrLf & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
        MsgBox MenError, vbExclamation, "Generar Albaran"
    End If
    GenerarAlbaran = b
    Espera 0.2
End Function



Private Function GenerarFactura(NumFactu As String) As Boolean
Dim b As Boolean
Dim NumTicket As String
Dim MenError As String
    
    On Error GoTo EGenFac
    
    conn.BeginTrans
    'si el tipo de forma de pago no es efectivo habrá que insertar
    'en la tabla de contabilidad conta.scobro
    '---- Laura: 10/10/2006 siempre se inserta en la scobro aunque sea efectivo
'    If TipoForPa <> 0 Then ConnConta.BeginTrans
    ConnConta.BeginTrans
    
    'Obtener el contador de ticket (ATI).
    b = ObtenerContadorAlbTicket(NumTicket)
    
    If b Then b = ObtenerContadorFactura(NumFactu)
    
    If b Then
        'Si lleva recfinan, meto una linea en la sliven
        'Meto la linea en la sliven
        If PorceRecFinan > 0 Then InsertarEnSliven
    
    
        'Actualizar los stocks de todos los articulos comprados
        'Insertar movimiento en smoval
        CodTipoMov = "ATI"
        b = InsertarMovAlmacen(NumTicket)
        If Not b Then MenError = "Control stock"
        
        'Insertar en el historico de facturas: scafac, scafac1,slifac
        'en el campo scafac1.numalbar guardamos el nº de ticket
        If b Then
            CodTipoMov = "FAV"
            b = InsertarHistFactura(NumTicket, NumFactu, , MenError)
        End If
    End If
    
EGenFac:
    If Err.Number <> 0 Then
        MenError = Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        '---- Laura 10/10/2006: siempre se inserta en la conta.scobro aunque sea efectivo
        'If TipoForPa <> 0 Then ConnConta.CommitTrans
        ConnConta.CommitTrans
    Else
        conn.RollbackTrans
        '---- Laura 10/10/2006: siempre se inserta en la conta.scobro aunque sea efectivo
        'If TipoForPa <> 0 Then ConnConta.RollbackTrans
        ConnConta.RollbackTrans
        MsgBox "ERROR: " & MenError & vbCrLf, vbExclamation, "Generar Factura"
    End If
    GenerarFactura = b
    Espera 0.2
End Function



Private Sub ImprimirTicket_old()
Dim MIPATH As String
'Dim NomImpre As String

    On Error GoTo EImpTick
    
    SQL = cadImpresion & " and {scafac.fecfactu}=" & DBSet(miFechaTicket, "F")
    If Not HayRegParaInforme("scafac", SQL) Then Exit Sub


    MIPATH = App.Path & "\Informes\"
    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(miFechaTicket) & "," & Month(miFechaTicket) & "," & Day(miFechaTicket) & ")"
    
'    'Establecemos la impresora de ticket
'    If vParamTPV.NomImpresora <> "" Then
'        If Printer.DeviceName <> vParamTPV.NomImpresora Then
'            'guardamos la impresora que habia
'            NomImpre = Printer.DeviceName
'            'establecemos la de ticket
'            EstablecerImpresora vParamTPV.NomImpresora
'        End If
'    End If

    With frmVisReport
        .FormulaSeleccion = cadImpresion
        .SoloImprimir = True
        .OtrosParametros = ""
        .NumeroParametros = 0
        .MostrarTree = False
        .Informe = MIPATH & "rTPVTicket.rpt"
        .ConSubInforme = False
        .Opcion = 93
        .ExportarPDF = False
        .Show vbModal
    End With
    
    'volver la impresora a la predeterminada
'    EstablecerImpresora NomImpre
    
EImpTick:
    If Err.Number <> 0 Then MuestraError Err.Number, "Imprimir ticket.", Err.Description
End Sub



Private Sub ImprimirFactura()
Dim MIPATH As String
Dim CadParam As String, nomDocu As String
Dim numParam As Byte
Dim ImprimeDirecto As Boolean

    On Error Resume Next

    SQL = cadImpresion & " and {scafac.fecfactu}=" & DBSet(miFechaTicket, "F")
    If Not HayRegParaInforme("scafac", SQL) Then Exit Sub
     

    
    
    '===================================================
    '============ PARAMETROS ===========================
    
    If Not PonerParamRPT2(18, CadParam, numParam, nomDocu, ImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
        
    '===================================================
    If ImprimeDirecto Then
        cadImpresion = cadImpresion & " and {scafac.fecfactu}='" & Format(miFechaTicket, FormatoFecha) & "'"
        SQL = cadImpresion
        SQL = Replace(SQL, "{", "")
        SQL = Replace(SQL, "}", "")
        ImprimirDirectoFact SQL
    Else
        MIPATH = App.Path & "\Informes\"
        cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(miFechaTicket) & "," & Month(miFechaTicket) & "," & Day(miFechaTicket) & ")"
    
    
         With frmVisReport
            .FormulaSeleccion = cadImpresion
            .SoloImprimir = True ' (RAFA/ALZIRA 31082006)
            'No lleva multiminforme
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .MostrarTree = False
            .Informe = MIPATH & nomDocu
            .ConSubInforme = True
            .Opcion = 53
            .ExportarPDF = False
            .NumCopias = 2 ' (RAFA/ALZIRA 31082006)
            .Show vbModal
        End With
    End If
    If Err.Number <> 0 Then MuestraError Err.Number, "Imprimir Factura.", Err.Description
End Sub




Private Sub ImprimirAlbaran()
Dim MIPATH As String
Dim CadParam As String, nomDocu As String
Dim numParam As Byte
Dim ImprimeDirecto As Boolean

    SQL = cadImpresion '& " and {scafac.fecfactu}=" & DBSet(RSVenta!fecventa, "F")
    If Not HayRegParaInforme("scaalb", SQL) Then Exit Sub
     
    MIPATH = App.Path & "\Informes\"
    
    '===================================================
    '============ PARAMETROS ===========================
    
    If Not PonerParamRPT2(10, CadParam, numParam, nomDocu, ImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
    
    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    CadParam = CadParam & "pCodUsu=" & vUsu.codigo & "|"
    numParam = numParam + 1
    
    
    '=========================================================================
    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    SQL = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(0).Text, "N")
    If SQL <> "" Then
        CadParam = CadParam & "pTipoIVA=" & SQL & "|"
        numParam = numParam + 1
    End If
    
    
    '===================================================
    If ImprimeDirecto Then
        SQL = cadImpresion
        SQL = Replace(SQL, "{", "")
        SQL = Replace(SQL, "}", "")
        ImprimirDirectoAlb SQL
    
    Else
    
         With frmVisReport
            .FormulaSeleccion = cadImpresion
            .SoloImprimir = True ' (RAFA/ALZIRA 31082006)
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .MostrarTree = False
            .Informe = MIPATH & nomDocu
            .ConSubInforme = True
            .Opcion = 45
            .ExportarPDF = False
            .NumCopias = 2 ' (RAFA/ALZIRA 31082006)
            .Show vbModal
        End With
    End If
End Sub




Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String
Dim devuelve As String
Dim I As Byte

    'Llamamos a al form
    '##A mano
    Cad = ""
    
    Select Case cadB
        Case "0" 'Cliente
            tabla = "sclien"
            Titulo = "Clientes"
            devuelve = "0|1|2|3|"
            Cad = Cad & "Cod.Cli.|sclien|codclien|N|000000|13·"
            Cad = Cad & "Nom. Cliente|sclien|nomclien|T||47·"
            Cad = Cad & "Nom. Comer|sclien|nomcomer|T||25·"
            Cad = Cad & "NIF|sclien|nifclien|T||15·"
            cadB = ""
            
        Case "1" 'Forma pago
            tabla = "sforpa inner join stippa on sforpa.tipforpa=stippa.tipforpa "
            Titulo = "Formas de Pago"
            devuelve = "0|1|2|"
            Cad = Cad & "Cod.For.|sforpa|codforpa|N|000|14·"
            Cad = Cad & "Nom. Forma pago|sforpa|nomforpa|T||50·"
            Cad = Cad & "Tipo|sforpa|tipforpa|N||12·"
            Cad = Cad & "Desc Tip.|stippa|destippa|T||23·"
            'cad = cad & "Desc Tip.|sforpa|case tipforpa when 0 then ""Efectivo"" when 1 then ""Transferencia""  when 2 then ""Talón"" when 3 then ""Pagaré"" when 4 then ""Recibo bancario"" when 5 then ""Confirming"" end as desctipo|T||23·"
            cadB = ""
             
        Case "2" 'Trabajador
            tabla = "straba"
            Titulo = "Operadores"
            devuelve = "0|1|2|"
            Cad = Cad & "Cod.Op.|straba|codtraba|N|0000|25·"
            Cad = Cad & "Nom. Operador.|straba|nomtraba|T||55·"
            Cad = Cad & "NIF|straba|niftraba|T||15·"
            cadB = ""
             
        Case "5" 'direc./dpto del cliente
            If vParamAplic.HayDeparNuevo = 1 Then
                Titulo = "Dptos Cliente: "
                Desc = "Dpto."
            ElseIf vParamAplic.HayDeparNuevo = 0 Then
                Titulo = "Direc. Cliente: "
                Desc = "Direc."
            Else
                Titulo = "Obra Cliente: "
                Desc = "Obra"
            End If
            Titulo = Titulo & Text1(0).Text & " - " & Text2(0).Text
            Cad = Cad & "Cod. " & Desc & "|sdirec|coddirec|N|000|15·"
            Cad = Cad & "Desc. " & Desc & "|sdirec|nomdirec|T||65·"
            tabla = "sdirec"
            devuelve = "0|1|"
            cadB = "codclien=" & Text1(0).Text
    End Select
   
   
    
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges
        '#
        If tabla = "sdirec" Then frmB.Label1.FontSize = 11
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
        I = CInt(Me.imgBuscar(0).Tag)
        Text1_LostFocus (I)
        PonerFoco Text1(I)


    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonBuscar(Ind As Integer)
    imgBuscar_Click (Ind)
End Sub







Private Function PonerDptoEnCliente() As Boolean
Dim vClien As CCliente
Dim NomDpto As String

    Set vClien = New CCliente
    vClien.codigo = Text1(0).Text
    
    'si existe el departamento para el cliente
    If vClien.DptoCliente(Text1(5).Text, NomDpto) Then
        Text2(5).Text = NomDpto
        PonerDptoEnCliente = True
    Else
        PonerDptoEnCliente = False
    End If
    Set vClien = Nothing
End Function



Private Function ComprobarCtaBanCliente() As Boolean
Dim cCli As CCliente
Dim MenError As String

    Set cCli = New CCliente
    cCli.codigo = Text1(0).Text
    
    If cCli.LeerDatos(Text1(0).Text) Then
        ComprobarCtaBanCliente = cCli.ComprobarCtaBancaria(MenError)
        If Not ComprobarCtaBanCliente Then MsgBox MenError & vbCrLf & "Contacte con Administración.", vbInformation
    End If
    
    Set cCli = Nothing
End Function

'ParaTiket = true -> hacer un ticket
'            false -> albaran factura
Private Function HayArticuloFitosanitario_O_BloqFamilia(ParaTiket As Boolean) As Boolean
'comprueba si entre las lineas de venta insertadas hay algun articulo
'q tiene registro fitosanitario (en ese caso no se puede crear Ticket)
'(OUT) -> true si encuentra algun articulo fitosanitario
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim I As Integer
Dim Clivario As Boolean
    On Error GoTo ErrFito
    
    SQL = "SELECT distinct sliven.nomartic,numserie FROM sliven"
    SQL = SQL & " inner join sartic on sliven.codartic=sartic.codartic"
    SQL = SQL & " WHERE " & Replace(cadSel, "scaven", "sliven")
    SQL = SQL & " and not isnull(numserie) and trim(numserie)<>''"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Rs.EOF Then
        'no hay articulos con registro fitosanitario
        HayArticuloFitosanitario_O_BloqFamilia = False
    Else
        'hay articulos fitosanitarios
        HayArticuloFitosanitario_O_BloqFamilia = True
        
        '- seleccionamos algunos articulos para mostrar en el mensaje
        I = 1
        SQL = ""
        While Not Rs.EOF And I < 3
            If SQL <> "" Then SQL = SQL & vbCrLf
            SQL = SQL & DBLet(Rs!NomArtic, "T") & " (" & DBLet(Rs!numSerie, "T") & ")"
            
            I = I + 1
            Rs.MoveNext
        Wend
        If I >= 3 And Not Rs.EOF Then SQL = SQL & vbCrLf & "..."
        
        '- mostramos mensaje de error
        If ParaTiket Then
            SQL = "NO se puede crear un Ticket ya que existen articulos con registro Fitosanitario. " & vbCrLf & SQL
            MsgBox SQL, vbExclamation
            
        Else
            'Vamos a hacer un albaran/factura.
            'Si hay fitosanitarios, el cliente debe estar identificado
            'Es decir, o no es de varios o si es de varios lleva el NIF
            Clivario = False
            SQL = "0"
            If Val(Text1(0).Text) = vParamTPV.Cliente Then
                SQL = "1"
            Else
                SQL = DevuelveDesdeBD(conAri, "clivario", "sclien", "codclien", Text1(0).Text)
            End If
            Clivario = SQL = "1"
            Dim Salir As Boolean
            
            Salir = False
            If Clivario Then
                If vParamTPV.ProhibirFitosantiarios_a_Varios Then
                    MsgBox "Ventas de fitosanitarios prohibidas a clientes varios", vbExclamation
                    Salir = True
                    'HayArticuloFitosanitario_O_BloqFamilia = False
                Else
                    'CLIENTE DE VARIOS. LO ha identificado?
                    If Text3.Tag = "" Then
                        MsgBox "Debe identificar al cliente para realizar una venta con productos fitosanitarios", vbExclamation
                        Salir = True
                    Else
                        ' Para los que  no tengan el modulo de fitosanitarios/manipulador
                        ' dejo pasar
                        
                        
                    End If
                End If
                
            Else
                ' Para los que  no tengan el modulo de fitosanitarios/manipulador
                ' dejo pasar
                'If Not vParamAplic.ManipuladorFitosanitarios2 Then HayArticuloFitosanitario_O_BloqFamilia = False                            'de momento dejamos pasar
              
            End If
            
            'Marzo 2015.
            If Salir Then
                HayArticuloFitosanitario_O_BloqFamilia = True
                Exit Function
            End If
            
            
            If HayArticuloFitosanitario_O_BloqFamilia Then
                'Veremos el carnet de manipulador
                'Veremos si el TITULAR tiene el carnet de manipulador
                If Val(Text1(0).Text) <> vParamTPV.Cliente Then
                
                    If Not vParamAplic.ManipuladorFitosanitarios2 Then
                
                        'Dejamos pasar
                
                    Else
                        'Llevamos control fitosanitarios.
                        'Debe bhaber seleccionado un carnet de la lista
                        If txtManipulador(0).Text = "" Or Text2(3).Text = "" Then
                            MsgBox "Seleccione un carnet de manipulador de fitosantiarios", vbExclamation
                        Else
                            HayArticuloFitosanitario_O_BloqFamilia = False
                            'para que vea si bloqueamos por familis
                        End If
                    End If
                Else
                    'Para varios tambien debe indicar carnet si lleva parametro
                    If vParamAplic.ManipuladorFitosanitarios2 Then
                        If txtManipulador(0).Text = "" Or Text2(3).Text = "" Then
                            MsgBox "No tiene carnet de manipulador de fitosantiarios", vbExclamation
                            HayArticuloFitosanitario_O_BloqFamilia = True
                        Else
                            HayArticuloFitosanitario_O_BloqFamilia = False
                        End If
                    Else
                        HayArticuloFitosanitario_O_BloqFamilia = False 'Dejo pasar. Cliente de varios IDENTIFICADO
                    End If
                End If
            End If
        End If
    End If
    
    Rs.Close
    
    
    If Not HayArticuloFitosanitario_O_BloqFamilia Then
        'LAS FAMIlias las bloqueaban para que hiceran albaran factura
        'Voy a comprobar si las familias de los articulos estan bloqueadas
        '------------------------------------------------------------------
        If ParaTiket Then
            SQL = "select sliven.nomartic,cantidad,codfamia from sliven,sartic where sliven.codartic=sartic.codartic "
            SQL = SQL & " AND " & Replace(cadSel, "scaven", "sliven")
            SQL = SQL & " AND codfamia in (select codfamia from sfamia where bloqEnTPV=1)"
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            While Not Rs.EOF
                SQL = SQL & vbCrLf & "- " & Rs!NomArtic & "   (" & Rs!cantidad & ")   Fam: " & Rs!Codfamia
                Rs.MoveNext
            Wend
            Rs.Close
            
            If SQL <> "" Then
                SQL = "No se puede vender por TICKET los articulos siguientes:" & SQL & vbCrLf & vbCrLf & "    DEBE HACER ALBARAN / FACTURA"
                MsgBox SQL, vbExclamation
                HayArticuloFitosanitario_O_BloqFamilia = True
            End If
        End If
    End If
    
    Set Rs = Nothing
    Exit Function
    
ErrFito:
    MuestraError Err.Number, "Comprobar articulos fitosanitarios", Err.Description
    'Pongo un true para que no siga
    HayArticuloFitosanitario_O_BloqFamilia = True
    Set Rs = Nothing
End Function


Private Function ActualizarCentroCoste() As Boolean
Dim SQL As String
Dim ccoste As String
Dim Rs As ADODB.Recordset

    On Error GoTo ErrActCC
    
    ' ---- [21/10/2009] [LAURA]: añadir campo centro de coste trabajador
    'Modif David.     10/11/2009
    '
    '   Si es por familia, ver las familias, si no por trabajador
    If vEmpresa.TieneAnalitica Then
        If vParamAplic.ModoAnalitica = 1 Then
            'Por FAMILIA
            Set Rs = New ADODB.Recordset
            SQL = "Select sliven.codartic,sfamia.codfamia,sfamia.codccost from sliven,sartic,sfamia"
            SQL = SQL & " WHERE sliven.codartic=sartic.codartic AND sartic.codfamia=sfamia.codfamia"
            SQL = SQL & " AND " & Replace(cadSel, "scaven", "sliven") & " ORDER BY codccost"
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                If IsNull(Rs!CodCCost) Then
                    SQL = ""
                    While Not Rs.EOF
                        SQL = SQL & Rs!Codfamia & "      -         " & DBLet(Rs!CodCCost, "T") & vbCrLf
                        Rs.MoveNext
                        
                        
                    Wend
                    Rs.Close
                    Set Rs = Nothing
                    SQL = "Familia      Centro de coste" & vbCrLf & String(20, "=") & vbCrLf & SQL
                    MsgBox SQL, vbExclamation
                    Exit Function
                Else
                    'OK. Actualizamos el CC
                    SQL = "UPDATE sliven set codccost=" & DBSet(Rs!CodCCost, "T") & " WHERE codartic="
                    SQL = SQL & DBSet(Rs!codArtic, "T") & " AND " & Replace(cadSel, "scaven", "sliven")
                    conn.Execute SQL
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            ActualizarCentroCoste = True
        Else
            
            
            'si contab. analitica por trabajador traer su centro de coste
            ccoste = DevuelveDesdeBDNew(conAri, "straba", "codccost", "codtraba", Text1(2).Text, "N")
            If ccoste = "" Then
                ActualizarCentroCoste = False
                MsgBox "El operador no tiene asignado un centro de coste.", vbInformation
            Else
                SQL = "UPDATE sliven SET codccost=" & DBSet(ccoste, "T", "S")
                SQL = SQL & " WHERE " & Replace(cadSel, "scaven", "sliven")
                conn.Execute SQL
                ActualizarCentroCoste = True
            End If
        End If
    
    Else
        ActualizarCentroCoste = True
    End If

    Exit Function
    
ErrActCC:
    ActualizarCentroCoste = False
    MuestraError Err.Number, "Centro de coste por trabajador.", Err.Description
End Function



Private Function ClienteOK(newCli As String, oldCli As String, Optional mostrarObs As Boolean) As Boolean
'(IN) newCli: cliente nuevo q queremos poner
'(IN) oldCli: cliente guardado actualmente si existe
Dim cCli As CCliente
Dim devuelve As String

    On Error GoTo ErrCliOK
    ClienteOK = False
    
    If newCli <> "" Then newCli = CStr(Val(newCli))
    Set cCli = New CCliente
    If cCli.LeerDatos(newCli) Then
        '-- Si el cliente esta bloqueado no permitimos este cliente para la venta
        If cCli.ClienteBloqueado Then
            Set cCli = Nothing
            Exit Function
        End If
        
        '-- si se ha modificado el cliente y si hay lineas de articulos:
        '   comprobar q el nuevo cliente tiene la misma tarifa q el cliente anterior
        '   sino no permitimos el nuevo cliente para la venta
        If (oldCli <> "") And (newCli <> oldCli) Then
            'obtener la tarifa del cliente actual
            devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", oldCli, "N")
            If devuelve <> CStr(cCli.Tarifa) Then
                devuelve = "No se puede seleccionar el cliente " & newCli & " "
                devuelve = devuelve & "ya que tiene distinta tarifa de precios." & vbCrLf
                devuelve = devuelve & "Seleccione un cliente de la misma tarifa o elimine la venta."
                MsgBox devuelve, vbExclamation, "Comprobar cliente"
                Set cCli = Nothing
                Exit Function
            End If
        End If
        ClienteOK = True
        
        'mostrar las observaciones del cliente
        If mostrarObs Then cCli.MostrarObservaciones
        
        
        'Si hay fitosanitarios, limpio el campode carnet...
       If newCli <> oldCli Then Limpiarmanipulador
    End If
    
    Set cCli = Nothing
    Exit Function
    
ErrCliOK:
    MuestraError Err.Number, "Comprobar cliente correcto.", Err.Description
End Function

Private Sub LanzarClientesVarios()

    Set frmClv = New frmFacClientesV
    frmClv.DatosADevolverBusqueda = "0|1|"
    frmClv.vNif = Text3.Tag
    frmClv.Show vbModal
    Set frmClv = Nothing
    ActualizarCliVariosEnBD
End Sub




Private Sub ActualizarCliVariosEnBD()
        
    If Text3.Tag = "" Then
        SQL = "NULL"
    Else
        SQL = DBSet(Text3.Tag, "T")
    End If
    SQL = "UPDATE scaven SET nifvarios = " & SQL
    
    If vParamAplic.ManipuladorFitosanitarios2 Then
        
        SQL = SQL & ", ManipuladorNumCarnet =" & DBSet(txtManipulador(0).Text, "T", "S")
        SQL = SQL & ", ManipuladorFecCaducidad =" & DBSet(txtManipulador(1).Text, "F", "S")
        SQL = SQL & ", ManipuladorNombre =" & DBSet(Text2(3).Text, "T", "S")
        SQL = SQL & ", TipoCarnet = "
        If txtManipulador(0).Text = "" Then
            SQL = SQL & "NULL"
        Else
            SQL = SQL & IIf(UCase(txtManipulador(2).Text) = "CUALIFICADO", 2, 1)
        End If
    End If
    
    If cadSel <> "" Then
        SQL = SQL & " WHERE " & cadSel
        On Error Resume Next
        conn.Execute SQL
        If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    End If
End Sub



Private Sub PonerDatosClienteVarios()
    SQL = DevuelveDesdeBD(conAri, "nomclien", "sclvar", "nifclien", CStr(RSVenta!nifvarios), "T")
    If SQL = "" Then SQL = "Sin nombre cliente varios"
    SQL = RSVenta!nifvarios & "|" & SQL & "|"
    frmClv_DatoSeleccionado SQL
    SQL = ""
End Sub
        
'Si CF= nothing --> Albaran
'   si no es una factura
Private Function LeerDesdeTablaClienteVarios(ByRef CF As CFactura) As String
Dim RN As ADODB.Recordset
    Set RN = New ADODB.Recordset
    LeerDesdeTablaClienteVarios = ""
    RN.Open "Select * from sclvar WHERE nifclien = " & DBSet(Text3.Tag, "T"), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RN.EOF Then
        'Orden para generar el albaran
        ' nomclien domclien codpobla pobclien proclien nifclien telclien
        If CF Is Nothing Then
            LeerDesdeTablaClienteVarios = DBSet(RN!Nomclien, "T", "N") & ", " & DBSet(RN!domclien, "T", "N") & ","
            LeerDesdeTablaClienteVarios = LeerDesdeTablaClienteVarios & DBSet(RN!codpobla, "T", "N") & ", " & DBSet(RN!pobclien, "T", "N") & "," & DBSet(RN!proclien, "T", "N") & ","
            LeerDesdeTablaClienteVarios = LeerDesdeTablaClienteVarios & DBSet(RN!nifClien, "T", "N") & "," & DBSet(RN!telclien, "T")
        Else
                'ASignamos a la factura
                CF.NombreClien = RN!Nomclien
                CF.DomicilioClien = DBLet(RN!domclien, "T")
                CF.CPostal = DBLet(RN!codpobla, "T")
                CF.Poblacion = DBLet(RN!pobclien, "T")
                CF.Provincia = DBLet(RN!proclien, "T")
                CF.NIF = DBLet(RN!nifClien, "T")
                CF.Telefono = DBLet(RN!telclien, "T")
            
            
                LeerDesdeTablaClienteVarios = "OK"
        End If
        
        
    End If
    RN.Close
    Set RN = Nothing
End Function



Private Sub MultiInsercionCampos()
Dim I As Integer
Dim VariedadPartida As String

        'Quito el indicador # de multi campo
        If InStr(1, CadenaDesdeOtroForm, 1) > 0 Then CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 2)
        
        SQL = "fecventa = " & DBSet(RSVenta!fecventa, "F")
        SQL = SQL & " AND numtermi = " & RSVenta!NumTermi & " AND numventa"
        SQL = DevuelveDesdeBD(conAri, "max(numlinea)", "sliven2", SQL, CStr(RSVenta!NumVenta), "N")
        NumRegElim = 0
        If SQL <> "" Then NumRegElim = Val(SQL)
        NumRegElim = NumRegElim + 1
        SQL = ""
        While CadenaDesdeOtroForm <> ""
            I = InStr(1, CadenaDesdeOtroForm, "·#")
            
            If I = 0 Then
                CadenaDesdeOtroForm = ""
            Else
                cadImpresion = Mid(CadenaDesdeOtroForm, 1, I - 1)
                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, I + 2)
                
                VariedadPartida = "," & DBSet(RecuperaValor(cadImpresion, 2), "T", "S") & "," & DBSet(RecuperaValor(cadImpresion, 3), "T", "S")
                cadImpresion = RecuperaValor(cadImpresion, 1) 'cdocampo
                
                For I = 1 To Me.ListView1.ListItems.Count
                    'Si no lo ha insertado YA
                    If Val(Me.ListView1.ListItems(I).Text) = Val(cadImpresion) Then
                        cadImpresion = ""
                        Exit For
                    End If
                
                Next I
                
                If cadImpresion <> "" Then
                    
                        '  ' numtermi,numventa,fecventa,numlinea,codcampo
                        SQL = SQL & ", (" & RSVenta!NumTermi & "," & RSVenta!NumVenta
                        SQL = SQL & "," & DBSet(RSVenta!fecventa, "F") & "," & NumRegElim & "," & cadImpresion
                        SQL = SQL & VariedadPartida & ")"
                        NumRegElim = NumRegElim + 1
                End If
            End If
        Wend
        If SQL <> "" Then
            SQL = Mid(SQL, 2) 'quito la primera coma
            ' numtermi,numventa,fecventa,
            SQL = "INSERT INTO sliven2(numtermi,numventa,fecventa,numlinea,codcampo,nomvarie,nompartida) VALUES " & SQL
            If ejecutar(SQL, False) Then
                'Hay que refrescar y boton anyadir
        
            End If
        End If
        
        cadImpresion = ""
        SQL = ""
        
        '
        
End Sub





Private Sub CargaDatosCampos()
Dim IT
    'Para no meter MUCHOS ariagro.ss
    'Pongo @# y luego lo reemplazo por vparamaplic.Ariagro.
'    SQL = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie"
'    SQL = SQL & " from (@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
'    SQL = SQL & " inner join @#variedades on rcampos.codvarie = variedades.codvarie"
'    'where socio
'    SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
'
    
    SQL = "select rcampos.codcampo, rpartida.nomparti, variedades.nomvarie,rcampos.codclien,rsocios.codsocio,rsocios.nomsocio,rcampos.codsitua"
    SQL = SQL & " from ((@#rcampos inner join @#rpartida on rcampos.codparti = rpartida.codparti)"
    SQL = SQL & " inner join @#variedades on rcampos.codvarie = variedades.codvarie)"
    SQL = SQL & " inner join @#rsocios on rsocios.codsocio=rcampos.codsocio"
    
    SQL = Replace(SQL, "@#", vParamAplic.Ariagro & ".")
    
    SQL = SQL & " WHERE codcampo IN "
    SQL = SQL & "(select codcampo from sliven2 where numtermi=" & RSVenta!NumTermi
    SQL = SQL & " AND numventa=" & RSVenta!NumVenta & " AND fecventa = " & DBSet(RSVenta!fecventa, "F")
    SQL = SQL & ")"
    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        IT.Text = Format(miRsAux!codCampo, "000000")
        IT.SubItems(1) = DBLet(miRsAux!nomparti, "T")
        IT.SubItems(2) = DBLet(miRsAux!nomvarie, "T")
        IT.Tag = miRsAux!codCampo
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    PonerWidth
  
End Sub

Private Sub PonerWidth()
  If Me.ListView1.ListItems.Count > 0 Then
        
        Me.Width = 14715
        Me.cmdCampos.visible = False
        If Me.Left > 4000 Then Me.Left = 4100
    Else
        Me.Width = 8625
        Me.cmdCampos.visible = True
    End If
End Sub

Private Sub PonerformaDePago()
    On Error GoTo ePonerformaDePago

    ImporteFinal = CCur(ImporteInicial)


    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select codforpa,nomforpa,tipforpa,porgasfi from sforpa WHERE codforpa=" & Text1(1).Text, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    PorceRecFinan = 0
    
    If Not miRsAux.EOF Then
        Text1(1).Text = miRsAux!codforpa
        Text2(1).Text = miRsAux!nomforpa
        TipoForPa = miRsAux!tipforpa
        'ALZIRA
        'FORMA DE PAGO CON recargo financiero
        If vParamAplic.NumeroInstalacion = 1 Then PorceRecFinan = DBLet(miRsAux!porgasfi, "N")
            
            
    Else
        Text1(1).Text = ""
        Text2(1).Text = ""
    End If
    miRsAux.Close
    
    'ImporteFinal = CCur(ImporteInicial)  esta al ppio del SUB
    If PorceRecFinan > 0 Then
        
        
        ImporteFinal = ImporteFinal + Round2(((ImporteFinal * PorceRecFinan) / 100), 2)
    
        
    End If
    
    
    
ePonerformaDePago:
    If Err.Number <> 0 Then
        MuestraError Err.Number
        Text1(1).Text = ""
        Text2(1).Text = ""
    End If
    Set miRsAux = Nothing
    Me.Label2(1).Caption = Format(ImporteFinal, FormatoImporte)
End Sub


Private Sub InsertarEnSliven()
Dim Aux As Currency
    Aux = ImporteFinal - ImporteInicial
    'Esta dentro de una transaccion. No protegemos con errorers
    SQL = Replace(cadSel, "scaven", "sliven")
    SQL = "select numtermi,numventa,fecventa,horventa,numlinea from sliven where " & SQL
    
    SQL = SQL & " order by numlinea desc"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not miRsAux.EOF Then
        'numtermi,numventa,fecventa,horventa,numlinea,codartic,nomartic,cantidad,precioiv,importel,precioar,codigiva,dto1,dto2,implineareal,codccost
        cadImpresion = "codigiva"
        SQL = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArticuloRecargoFinanciero, "T", cadImpresion)
        SQL = miRsAux!numlinea + 1 & "," & DBSet(vParamAplic.ArticuloRecargoFinanciero, "T") & "," & DBSet(SQL, "T") & ","
        SQL = " VALUES (" & miRsAux!NumTermi & "," & miRsAux!NumVenta & "," & DBSet(miRsAux!fecventa, "F") & "," & DBSet(miRsAux!horventa, "FH") & "," & SQL
        SQL = SQL & "1," & DBSet(Aux, "N") & "," & DBSet(Aux, "N") & "," & DBSet(Aux, "N") & "," & cadImpresion & ",0,0," & DBSet(Aux, "N") & ")"
        
        SQL = "INSERT INTO sliven(numtermi,numventa,fecventa,horventa,numlinea,codartic,nomartic,cantidad,precioiv,importel,precioar,codigiva,dto1,dto2,implineareal) " & SQL
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If SQL <> "" Then conn.Execute SQL
    cadImpresion = ""
    
End Sub

