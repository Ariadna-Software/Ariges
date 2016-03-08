VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVarios3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   13785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameVerificarCCCAriadna 
      Height          =   8535
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   13095
      Begin VB.CheckBox chkActualizarNIF 
         Caption         =   "Actualizar NIF"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   8040
         Width           =   1335
      End
      Begin VB.CommandButton cmdActualErroresCCC 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   8040
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimirErrores 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   10200
         TabIndex        =   11
         Top             =   8040
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   7695
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   13573
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Codmacta"
            Object.Width           =   2116
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "C.C.C"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "NIF"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Aplic"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "BD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Campo1"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   11520
         TabIndex        =   9
         Top             =   8040
         Width           =   1095
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmVarios3.frx":0000
         ToolTipText     =   "Puntear al haber"
         Top             =   8040
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmVarios3.frx":014A
         ToolTipText     =   "Quitar al haber"
         Top             =   8040
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   13
         Top             =   8160
         Width           =   105
      End
   End
   Begin VB.Frame FramePresuElim 
      Height          =   8655
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CommandButton cmdEliminarFAZ 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   8760
         TabIndex        =   29
         Top             =   8160
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   5
         Left            =   10080
         TabIndex        =   27
         Top             =   8160
         Width           =   1095
      End
      Begin MSComctlLib.ListView lwPresuElim 
         Height          =   7695
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   13573
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   1940
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   30
         Top             =   8280
         Width           =   105
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   240
         Picture         =   "frmVarios3.frx":0294
         ToolTipText     =   "Quitar al haber"
         Top             =   8280
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   600
         Picture         =   "frmVarios3.frx":03DE
         ToolTipText     =   "Puntear al haber"
         Top             =   8280
         Width           =   240
      End
   End
   Begin VB.Frame FrameNuevaFamiliaAgrupado 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmdVentasAgrupadas 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cboTipoFra 
         Height          =   315
         ItemData        =   "frmVarios3.frx":0528
         Left            =   1440
         List            =   "frmVarios3.frx":0535
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtFamia 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtDescFamia 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   96
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   600
      End
      Begin VB.Image imgFamilia 
         Height          =   240
         Index           =   0
         Left            =   1200
         Picture         =   "frmVarios3.frx":0566
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame FrameCutoaTfno 
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtimporte 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdCuotaTfno 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtTextoPlano 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   720
         Width           =   3975
      End
      Begin VB.Image imgCuota 
         Height          =   240
         Left            =   960
         Picture         =   "frmVarios3.frx":0668
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   4560
         TabIndex        =   21
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame FrameTelefonosSinConsumo 
      Height          =   6495
      Left            =   360
      TabIndex        =   22
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton cmdEliminarFrasSinConsumo 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   6360
         TabIndex        =   25
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   3
         Left            =   7560
         TabIndex        =   24
         Top             =   5880
         Width           =   1095
      End
      Begin MSComctlLib.ListView lwTelefoDe 
         Height          =   5535
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9763
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Telefono"
            Object.Width           =   2452
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Inact."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Error Oper."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   480
         Picture         =   "frmVarios3.frx":076A
         ToolTipText     =   "Puntear al haber"
         Top             =   6000
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   120
         Picture         =   "frmVarios3.frx":08B4
         ToolTipText     =   "Quitar al haber"
         Top             =   6000
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVarios3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0.- Insertar familias en listado vtentas agrupado ALZIRA
    '1.- Comprobacion cuentas erroneas en aplicaciones ARIADNA
    '2.- Cuotas propias de telefonia
    '3.- Telefonos BOLBAITE sin consumo ni cutoas ni na de na
        
    '4.- Cutoas porpias para insercion masiva
    '5.- Eliminar facturas FAZ (presupuestos) herbelca
    
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
    
    
Dim miSQL As String
Dim PrimVez As Boolean


Dim PulsadoCerrar As Boolean
Dim It As ListItem

Private Sub cboTipoFra_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cmdActualErroresCCC_Click()
Dim Aux As String
Dim LinkaPorCodmacta As Byte   'LinkaPorCodmacta    '0. Codmacta   1.- Codclien
    miSQL = ""
    For NumRegElim = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Checked Then
                                            'OK, es de las de actualizar
            If ListView1.ListItems(NumRegElim).Tag > 0 Then miSQL = miSQL & "1"
        End If
    Next
    
    If miSQL = "" Then
        MsgBox "Ningun dato seleccionado para actualizar", vbExclamation
        Exit Sub
    End If
    
    miSQL = Len(miSQL)
    miSQL = "Va a actualizar " & miSQL & " registros."
    
    
    If Me.chkActualizarNIF.Value = 1 Then _
        miSQL = miSQL & vbCrLf & vbCrLf & vbCrLf & "*****    Va actualizar tambien el N.I.F.   **************" & vbCrLf & vbCrLf
    
    miSQL = miSQL & vbCrLf & vbCrLf & " ¿Desea continuar?"
    If MsgBox(miSQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    'Grabaremos LOG
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    If Me.chkActualizarNIF.Value = 1 Then CadenaDesdeOtroForm = "[NIF]"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "·" 'Para el LOG
    Me.FrameVerificarCCCAriadna.Enabled = False
    
    Set LOG = New cLOG
    
    miSQL = DevuelveDesdeBD(conAri, "LinkaPorCodmacta", "spara2", "1", "1")
    If miSQL = "" Then miSQL = "0"
    LinkaPorCodmacta = CByte(Val(miSQL))
    
    
    For NumRegElim = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(NumRegElim).Checked Then
                                            'OK, es de las de actualizar
            If ListView1.ListItems(NumRegElim).Tag > 0 Then
                
                davidNumalbar = ListView1.ListItems(NumRegElim).Tag
                miSQL = Me.ListView1.ListItems(NumRegElim).SubItems(6)
                Label4(1).Caption = Me.ListView1.ListItems(davidNumalbar).SubItems(1) & " -> " & miSQL
                Label4(1).Refresh
                
                miSQL = MontaSQlUpdateErrorCta(davidNumalbar, LinkaPorCodmacta)
                If miSQL <> "" Then
                    'PARA EL LOG
                    If InStr(1, CadenaDesdeOtroForm, "·" & ListView1.ListItems(davidNumalbar).Text & "·") = 0 Then
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "·" & ListView1.ListItems(davidNumalbar).Text & "·"
                        

                    End If
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ListView1.ListItems(NumRegElim).SubItems(6) & ","
                    If Len(CadenaDesdeOtroForm) > 240 Then
                        LOG.Insertar 22, vUsu, CadenaDesdeOtroForm
                        Espera 0.5
                        CadenaDesdeOtroForm = ""
                        If Me.chkActualizarNIF.Value = 1 Then CadenaDesdeOtroForm = "[NIF]"
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "·" 'Para el LOG
                    End If
                    conn.Execute miSQL
                    miSQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo
                    miSQL = miSQL & " AND codigo1=" & ListView1.ListItems(davidNumalbar)
                    miSQL = miSQL & " AND campo1=" & ListView1.ListItems(NumRegElim).SubItems(7)
                    conn.Execute miSQL
                End If
            End If
        End If
    Next
    If Len(CadenaDesdeOtroForm) > 1 Then LOG.Insertar 22, vUsu, CadenaDesdeOtroForm
    
    Label4(1).Caption = "Cargar datos"
    CargaItemsErroresCCC
    
    Label4(1).Caption = "" 'indicador
    davidNumalbar = 0   'reutilizada
    
    FrameVerificarCCCAriadna.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub



'LinkaPorCodmacta    '0. Codmacta   1.- Codclien
Private Function MontaSQlUpdateErrorCta(Padre As Long, LinkaPorCodmacta As Byte) As String
Dim UpdateaAlgo As Byte '0 Nada    1 CCC    2 NIF    3 AMBOS
Dim C1 As String
Dim LaCuenta As Boolean

'Cuando linka por CODIGOCLIENTE entonces el update a la conta NO va por el codmacta de ariges.
' va por el codmacta de rsocios,secciones ya que para buscar el socio no lo hace por codmacta
'Con lo cual proveedores(este ya lo hacia) como clientes de ariagro tienen que ir por codmacta indicada
Dim CodmactaIndicada As Boolean

        UpdateaAlgo = 0
          
        'SI UPDATEAMOS C.C.C.
        'Es decir, si la columna CCC no tiene valor salimos
        If Trim(Me.ListView1.ListItems(NumRegElim).SubItems(3)) <> "" Then UpdateaAlgo = 1
            
        If Me.chkActualizarNIF.Value = 1 Then
            If Trim(Me.ListView1.ListItems(NumRegElim).SubItems(4)) <> "" Then
                If Trim(Me.ListView1.ListItems(NumRegElim).SubItems(4)) <> "VACIO" Then UpdateaAlgo = UpdateaAlgo + 2
            End If
        End If
        
        If UpdateaAlgo = 0 Then Exit Function
    
        'codbanco|codsucur|digcontr|cuentaba|notabla|codmacta|"
        LaCuenta = True
        Select Case ListView1.ListItems(NumRegElim).SubItems(7)
        Case 2
            'miSQL = "arigasol"
            C1 = "iban|codbanco|codsucur|digcontr|cuentaba|##ssocio|codmacta|nifsocio|"
        Case 1, 3, 5, 6
             ' "contaariges" "contagasol"  "agrocli" "agroprov"
             C1 = "iban|entidad|oficina|CC|cuentaba|##cuentas|codmacta|nifdatos|"
             
             If ListView1.ListItems(NumRegElim).SubItems(7) = 6 Then
                CodmactaIndicada = True
             Else
                If ListView1.ListItems(NumRegElim).SubItems(7) = 5 Then CodmactaIndicada = (LinkaPorCodmacta = 1)
             End If
        Case 4
             ' "socios"
             C1 = "iban|codbanco|codsucur|digcontr|cuentaba|##rsocios"
             C1 = C1 & " inner join ##rsocios_seccion on rsocios_seccion.codsocio=rsocios.codsocio |"
             
             If LinkaPorCodmacta = 0 Then
                C1 = C1 & "codmaccli"
             Else
                C1 = C1 & "rsocios.codsocio"
                LaCuenta = False
             End If
             C1 = C1 & "|nifsocio|"
             
        Case Else
            Stop
            C1 = ""
        End Select


        ''0. Codmacta   1.- Codclien
        With ListView1.ListItems(Padre)
            'Actualizamos CCC
            If UpdateaAlgo <> 2 Then
                MontaSQlUpdateErrorCta = RecuperaValor(C1, 1) & " = '" & Mid(.SubItems(3), 1, 4) & "',"
                MontaSQlUpdateErrorCta = MontaSQlUpdateErrorCta & RecuperaValor(C1, 2) & " = '" & Mid(.SubItems(3), 5, 4) & "',"
                MontaSQlUpdateErrorCta = MontaSQlUpdateErrorCta & RecuperaValor(C1, 3) & " = '" & Mid(.SubItems(3), 9, 4) & "', "
                MontaSQlUpdateErrorCta = MontaSQlUpdateErrorCta & RecuperaValor(C1, 4) & " = '" & Mid(.SubItems(3), 13, 2) & "', "
                MontaSQlUpdateErrorCta = MontaSQlUpdateErrorCta & RecuperaValor(C1, 5) & " = '" & Trim(Mid(.SubItems(3), 15)) & "' "
            End If
            If UpdateaAlgo <> 1 Then
                If MontaSQlUpdateErrorCta <> "" Then MontaSQlUpdateErrorCta = MontaSQlUpdateErrorCta & ", "
                MontaSQlUpdateErrorCta = RecuperaValor(C1, 8) & " = '" & Trim(.SubItems(4)) & "'"
            End If
            

        End With
        
        MontaSQlUpdateErrorCta = "UPDATE " & RecuperaValor(C1, 6) & " SET " & MontaSQlUpdateErrorCta
        MontaSQlUpdateErrorCta = MontaSQlUpdateErrorCta & " WHERE " & RecuperaValor(C1, 7) & " = '"
        'Vemos en que BD reemplazando ## por la BD
        MontaSQlUpdateErrorCta = Replace(MontaSQlUpdateErrorCta, "##", Me.ListView1.ListItems(NumRegElim).SubItems(6) & ".")
        
        'If ListView1.ListItems(NumRegElim).SubItems(7) = 6 Then
        If CodmactaIndicada Then
            'Cuenta PROVEEDOR
            'esta en la columna de la cuenta, pero no la del padre, la seleccionada. Es cta prov 400 o 4100
            MontaSQlUpdateErrorCta = MontaSQlUpdateErrorCta & Trim(Me.ListView1.ListItems(NumRegElim).SubItems(2)) & "'"
        Else
            'Para todos los demas
            If LaCuenta Then
                MontaSQlUpdateErrorCta = MontaSQlUpdateErrorCta & Trim(Me.ListView1.ListItems(Padre).SubItems(2)) & "'"
            Else
                MontaSQlUpdateErrorCta = MontaSQlUpdateErrorCta & Trim(Me.ListView1.ListItems(Padre).Text) & "'"
            End If
        End If
        
End Function


Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 0 Then CadenaDesdeOtroForm = "" 'por si acaso
    
    PulsadoCerrar = True
    
    'Tfnos sin consumos ni cutoas
    'Si da a cancelar es que no quiere borrar ninguno
    'Eso significa borrar toos los de la tabla
    If Index = 3 Then ejecutar "DELETE from tmpnseries WHERE codusu = " & vUsu.codigo, False
    
    'Opcion 4
    If Opcion = 4 Then CadenaDesdeOtroForm = "" 'por si las moscas
    
    Unload Me
End Sub

Private Sub cmdCuotaTfno_Click()
     
    Me.txtTextoPlano(0).Text = Trim(txtTextoPlano(0).Text)
    Me.txtimporte(0).Text = Trim(txtimporte(0).Text)
    If Me.txtTextoPlano(0).Text = "" Or Trim(txtimporte(0).Text) = "" Then Exit Sub
    
    ' numlinea
    If ImporteFormateado(Me.txtimporte(0).Text) = 0 Then
        MsgBox "Importe debe ser > 0", vbExclamation
        Exit Sub
    End If
    
    
    
    If Opcion = 4 Then
    
        CadenaDesdeOtroForm = txtTextoPlano(0).Tag & "|" & txtTextoPlano(0).Text & "|" & txtimporte(0).Text & "|"
        Unload Me
    Else
        'Insertando cuota para un socio/cliente
        'CadenaDesdeOtroForm     telefono|numlinea|desc|precio|   ya habremos calculado el nuemero mayor
        
        miSQL = RecuperaValor(CadenaDesdeOtroForm, 3)
        If miSQL = "" Then
            'NUEVO. EL id de la cutoa de coperativa es la que insertare
            'No la max que viene desde alli
            ' sclientfnoCuotas(IdTelefono numlinea  descripcion precio
            
            miSQL = "'" & RecuperaValor(CadenaDesdeOtroForm, 1) & "'," & txtTextoPlano(0).Tag
            'ANTES  miSQL = "'" & RecuperaValor(CadenaDesdeOtroForm, 1) & "'," & RecuperaValor(CadenaDesdeOtroForm, 2)
            
            miSQL = miSQL & "," & DBSet(Me.txtTextoPlano(0).Text, "T") & "," & DBSet(Me.txtimporte(0), "N")
            miSQL = "INSERT INTO sclientfnoCuotas(IdTelefono, numlinea ,descripcion ,precio) VALUES (" & miSQL & ")"
            
        Else
            'UPDATE
            miSQL = "UPDATE sclientfnoCuotas set descripcion = " & DBSet(Me.txtTextoPlano(0).Text, "T")
            miSQL = miSQL & ", precio =" & DBSet(Me.txtimporte(0), "N")
            miSQL = miSQL & " WHERE IdTelefono = '" & RecuperaValor(CadenaDesdeOtroForm, 1) & "' AND numlinea =" & RecuperaValor(CadenaDesdeOtroForm, 2)
            ' sclientfnoCuotas(IdTelefono numlinea  descripcion precio
            
            
        End If
        
        If ejecutar(miSQL, False) Then Unload Me
    
    End If
    
End Sub

Private Sub cmdEliminarFAZ_Click()
    miSQL = ""
    For NumRegElim = 1 To Me.lwPresuElim.ListItems.Count
        If Me.lwPresuElim.ListItems(NumRegElim).Checked Then miSQL = miSQL & "X"
    Next
    
    If miSQL = "" Then
        MsgBox "Seleccione alguna factura", vbExclamation
        Exit Sub
    End If
    
    
    miSQL = "Va a eliminar " & Len(miSQL) & " factura(s). " & vbCrLf & vbCrLf & "Introduzca el password para continuar"
    miSQL = InputBox(miSQL)
    If UCase(miSQL) <> "ARIADNA" Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    Set LOG = New cLOG
    
        
    
    
    CadenaDesdeOtroForm = "" 'PARA EL LOG
    For NumRegElim = 1 To Me.lwPresuElim.ListItems.Count
        
        If Me.lwPresuElim.ListItems(NumRegElim).Checked Then
            Label4(4).Caption = lwPresuElim.ListItems(NumRegElim).Text & " - " & lwPresuElim.ListItems(NumRegElim).SubItems(3)
            Label4(4).Refresh
                    
                    
            EliminarFacturaFAZ
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & lwPresuElim.ListItems(NumRegElim).Text & " " & lwPresuElim.ListItems(NumRegElim).SubItems(1) & ";   "
            If Len(CadenaDesdeOtroForm) > 220 Then
                Label4(4).Caption = "Actualizar registros"
                Label4(4).Refresh
                CadenaDesdeOtroForm = "[PRESU] " & CadenaDesdeOtroForm
                LOG.Insertar 1, vUsu, CadenaDesdeOtroForm
                CadenaDesdeOtroForm = ""
                Espera 0.6
            End If
            Me.Refresh
        End If
    Next
    If CadenaDesdeOtroForm <> "" Then
        CadenaDesdeOtroForm = "[PRESU] " & CadenaDesdeOtroForm
        LOG.Insertar 1, vUsu, CadenaDesdeOtroForm
    End If
    
    Set LOG = Nothing
    Unload Me
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub cmdEliminarFrasSinConsumo_Click()
    miSQL = ""
    For NumRegElim = 1 To Me.lwTelefoDe.ListItems.Count
        If Me.lwTelefoDe.ListItems(NumRegElim).Checked Then
            If Trim(lwTelefoDe.ListItems(NumRegElim).Text) <> "" Then miSQL = miSQL & "X"
        End If
    Next
    
    If miSQL = "" Then
        MsgBox "Seleccione algun teléfono", vbExclamation
    Else
        miSQL = CStr(Len(miSQL))
        miSQL = "Va a eliminar " & miSQL & " telefono(s) de la facturacion" & vbCrLf & vbCrLf
        miSQL = miSQL & "¿Continuar?"
        If MsgBox(miSQL, vbQuestion + vbYesNo) <> vbYes Then miSQL = ""
    End If
    
    If miSQL <> "" Then
        'Vamos a eliminar telefonos
        'Para saber que telefonos tiene que eliminar el proceso, realmente borramos los que NO
        'tiene que eliminar
        '---> Los que no
        miSQL = ""
        For NumRegElim = 1 To Me.lwTelefoDe.ListItems.Count
            If Not Me.lwTelefoDe.ListItems(NumRegElim).Checked Then
                If Trim(lwTelefoDe.ListItems(NumRegElim).Text) <> "" Then miSQL = miSQL & ", '" & Me.lwTelefoDe.ListItems(NumRegElim).Text & "'"
            End If
        Next
        If miSQL <> "" Then
            miSQL = Mid(miSQL, 2)
            miSQL = "DELETE from tmpnseries WHERE codartic in (" & miSQL & ")"
            conn.Execute miSQL
        End If
        PulsadoCerrar = True
        Unload Me
    End If
End Sub

Private Sub cmdImprimirErrores_Click()
    With frmImprimir
        .FormulaSeleccion = "{tmpcrmclien.codusu} = " & vUsu.codigo & " AND {tmpInformes.campo1}>0"
        .OtrosParametros = "|pEmpresa=""" & vParam.NombreEmpresa & """|" & Me.cmdImprimirErrores.Tag
        .NumeroParametros = 2 'numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 3002
        .Titulo = "Errores C.C.C."
        .NombreRpt = "rComprobarCCC.rpt"
        .ConSubInforme = False
        .Show vbModal
    End With
End Sub

Private Sub cmdVentasAgrupadas_Click()
    If Me.txtFamia(0).Text = "" Then Exit Sub
    CadenaDesdeOtroForm = txtFamia(0).Text & "|" & Format(txtFamia(0).Text, "0000") & " - " & Me.txtDescFamia(0).Text & "|" & Me.cboTipoFra.ListIndex & "|"
    Unload Me
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        PulsadoCerrar = True
        Screen.MousePointer = vbHourglass
        
        If Opcion = 1 Then CargaItemsErroresCCC
        
        If Opcion = 2 Then DatosCuotasPropiasTelefonia
            
        If Opcion = 3 Then
            PulsadoCerrar = False
            CargaTelefonosSinConsumo
        End If
        
        If Opcion = 5 Then CargalwPresuElim
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim indice As Integer
    Me.Icon = frmPpal.Icon
    FrameNuevaFamiliaAgrupado.visible = False
    FrameVerificarCCCAriadna.visible = False
    FrameCutoaTfno.visible = False
    FrameTelefonosSinConsumo.visible = False
    FramePresuElim.visible = False
    limpiar Me
    PrimVez = True
    indice = Opcion
    Select Case Opcion
    Case 0
        PonerFrameVisible FrameNuevaFamiliaAgrupado
        If CadenaDesdeOtroForm = "" Then
            'Nuevo
            cboTipoFra.ListIndex = 0
        Else
            Me.txtFamia(0).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Me.txtDescFamia(0).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            CadenaDesdeOtroForm = RecuperaValor(CadenaDesdeOtroForm, 3)
            Me.cboTipoFra.ListIndex = Val(CadenaDesdeOtroForm)
            
            
        End If
        Me.txtFamia(0).Enabled = CadenaDesdeOtroForm = ""
        Me.txtDescFamia(0).Enabled = CadenaDesdeOtroForm = ""
        Me.imgFamilia(0).visible = CadenaDesdeOtroForm = ""
        CadenaDesdeOtroForm = "" 'La pongo a ""
    Case 1
    
        'En cadenadesdeotroform viene los desde hasta para la impresion
    
        PonerFrameVisible FrameVerificarCCCAriadna
        Caption = "Verificar datos ariadna"
        cmdActualErroresCCC.visible = vUsu.Nivel <= 1
        cmdImprimirErrores.Tag = CadenaDesdeOtroForm
      
    Case 2, 4
        
        PonerFrameVisible FrameCutoaTfno
        Caption = "Cuotas propias telefonía"
        If Opcion = 4 Then indice = 2
    Case 3
        
        PonerFrameVisible FrameTelefonosSinConsumo
        Caption = "Telefonos sin consumo"
        
    Case 5
        
        PonerFrameVisible FramePresuElim
        Caption = "Eliminar presupuestos"
        Label4(4).Caption = ""
    End Select
    
    Me.cmdCancelar(indice).Cancel = True
    
End Sub



Private Sub PonerFrameVisible(ByRef F As Frame)
    F.Top = 0
    F.Left = 120
    F.visible = True
    Me.Height = F.Height + 480
    Me.Width = F.Width + 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Opcion = 3 Then
        If Not PulsadoCerrar Then Cancel = 1
    End If
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    miSQL = CadenaDevuelta
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If Index < 2 Then
    
        For NumRegElim = 1 To Me.ListView1.ListItems.Count
            ListView1.ListItems(NumRegElim).Checked = Index = 1
        Next

    ElseIf Index < 4 Then
        For NumRegElim = 1 To Me.lwTelefoDe.ListItems.Count
            lwTelefoDe.ListItems(NumRegElim).Checked = Index = 3
        Next
    Else
        For NumRegElim = 1 To Me.lwPresuElim.ListItems.Count
            lwPresuElim.ListItems(NumRegElim).Checked = Index = 4
        Next
    End If
End Sub

Private Sub imgCuota_Click()
    LanzaBuscaGrid 1
    If miSQL <> "" Then
        Me.txtTextoPlano(0).Text = RecuperaValor(miSQL, 2)
        Me.txtimporte(0).Text = RecuperaValor(miSQL, 3)
                
        miSQL = RecuperaValor(miSQL, 1)  'idcuota
        txtTextoPlano(0).Tag = miSQL
        
        If Opcion = 2 Then
            miSQL = "numlinea = " & miSQL & " AND idtelefono"
            miSQL = DevuelveDesdeBD(conAri, "descripcion", "sclientfnocuotas", miSQL, RecuperaValor(CadenaDesdeOtroForm, 1), "T")
            If miSQL <> "" Then
                MsgBox "Ya tienen asignada la cuota como: " & miSQL, vbExclamation
                txtTextoPlano(0).Text = "": txtimporte(0).Text = ""
            End If
        End If
        PonerFoco txtimporte(0)
                
        miSQL = ""
    End If
End Sub

Private Sub imgFamilia_Click(Index As Integer)
    LanzaBuscaGrid 0
    If miSQL <> "" Then
        
        Me.txtFamia(Index).Text = RecuperaValor(miSQL, 2)
        Me.txtDescFamia(Index).Text = RecuperaValor(miSQL, 3)
        PonerFoco txtFamia(Index)
        miSQL = ""
    End If
End Sub

'0.- Familia, 1. Cuotas operador
Private Sub LanzaBuscaGrid(LOpcion As Byte)
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    
    If LOpcion = 0 Then
    
        frmB.vTitulo = "Familia"
        miSQL = "Codigo|sfamia|Codfamia|N||30·"
        miSQL = miSQL & "descripcion|sfamia|nomfamia|T||65·"
        frmB.vTabla = "sfamia"
        frmB.vDevuelve = "0|1|"
        frmB.vSQL = ""
    Else
        frmB.vTitulo = "Cuotas propias"
        miSQL = "Codigo|stfnocuotaspropias|codigoCuota|N||30·"
        miSQL = miSQL & "descripcion|stfnocuotaspropias|nombre|T||50·"
        miSQL = miSQL & "Importe|stfnocuotaspropias|Importe|N||10·"
        frmB.vTabla = "stfnocuotaspropias"
        If Opcion = 2 Then
            frmB.vSQL = " operadora = " & RecuperaValor(CadenaDesdeOtroForm, 5)
        Else
            frmB.vSQL = ""
         End If
        frmB.vDevuelve = "0|1|2|"
        'operadora         CadenaDesdeOtroForm
    End If
    frmB.vCampos = miSQL
    frmB.vCargaFrame = False
    
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Ariges
    
    miSQL = ""
    frmB.Show vbModal
   
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
    'misql tiene el valor devuelto
End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Tag = 0 Then
        'Ha pinchado sobre el nodo "padre"
        NumRegElim = Item.Index + 1
        Do
            If NumRegElim > ListView1.ListItems.Count Then
                NumRegElim = 0
            Else
                If ListView1.ListItems(NumRegElim).Tag = 0 Then
                    NumRegElim = 0
                Else
                    ListView1.ListItems(NumRegElim).Checked = Item.Checked
                    NumRegElim = NumRegElim + 1
                End If
            End If
        Loop Until NumRegElim = 0
    End If
End Sub

Private Sub lwTelefoDe_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index - 1 = lwTelefoDe.SortKey Then
        If lwTelefoDe.SortOrder = lvwAscending Then
            lwTelefoDe.SortOrder = lvwDescending
        Else
            lwTelefoDe.SortOrder = lvwAscending
        End If
    Else
        lwTelefoDe.SortOrder = lvwAscending
        lwTelefoDe.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub txtFamia_GotFocus(Index As Integer)
    ConseguirFoco txtFamia(Index), 3
End Sub

Private Sub txtFamia_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFamia_LostFocus(Index As Integer)
    txtFamia(Index).Text = Trim(txtFamia(Index).Text)
    miSQL = ""
    If txtFamia(Index).Text <> "" Then
        If PonerFormatoEntero(txtFamia(Index)) Then
            miSQL = DevuelveDesdeBD(conAri, "nomfamia", "sfamia", "codfamia", txtFamia(Index).Text, "N")
            If miSQL = "" Then
                MsgBox "El codigo no pertence a ningun familia", vbExclamation
                txtFamia(Index).Text = ""
            End If
        Else
            txtFamia(Index).Text = ""
        End If
    End If
     
    Me.txtDescFamia(Index).Text = miSQL
    If txtFamia(Index).Text = "" Then PonerFoco txtFamia(Index)
    
End Sub



Private Sub CargaItemsErroresCCC()
Dim N As Integer

    ListView1.ListItems.Clear
    Label4(1).Caption = "Leyendo BD" 'indicador
    Label4(1).Refresh
    miSQL = "select tmpcrmclien.*,tmpinformes.*,nomclien"
    miSQL = miSQL & " from tmpcrmclien inner join sclien on tmpcrmclien.codclien=sclien.codclien  "
    miSQL = miSQL & " left join tmpinformes on tmpcrmclien.CodUsu = tmpinformes.CodUsu And "
    miSQL = miSQL & " tmpcrmclien.codclien = tmpinformes.Codigo1"
    miSQL = miSQL & " where tmpcrmclien.codusu=" & vUsu.codigo
    miSQL = miSQL & " and campo1>0 order by tmpcrmclien.codclien,campo1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = -1
    While Not miRsAux.EOF
        
        If miRsAux!codclien <> NumRegElim Then
            'Cliente nuevo
            Set It = ListView1.ListItems.Add
            It.Text = Format(miRsAux!codclien, "0000")
            It.SubItems(1) = miRsAux!Nomclien
            'Los 10 primeros (relleando a blancos sera el codmacta, los siguientes desde el 11 el NIF
            It.SubItems(2) = Mid(miRsAux!nomforpa, 1, 10)
            
            It.SubItems(3) = DBLet(miRsAux!nomactiv, "T")
            It.SubItems(4) = Mid(miRsAux!nomforpa, 11)
            It.SubItems(5) = " "
            It.Tag = 0
            NumRegElim = miRsAux!codclien
            Label4(1).Caption = It.Text
            Label4(1).Refresh
            davidNumalbar = It.Index
        End If
        
        
        
        Set It = ListView1.ListItems.Add
        It.Text = " "
        It.SubItems(1) = ""
        
        
        It.SubItems(2) = DBLet(miRsAux!nombre1, "T") & " "
        It.SubItems(3) = DBLet(miRsAux!nombre2, "T") & " "
        It.SubItems(4) = DBLet(miRsAux!nombre3, "T") & " "
    
        
        
        N = 0
        If miRsAux!campo1 = 2 Then
            miSQL = "arigasol"
        ElseIf miRsAux!campo1 = 3 Then
            miSQL = "contagasol"
        ElseIf miRsAux!campo1 = 4 Then
            miSQL = "socios"
        ElseIf miRsAux!campo1 = 5 Then
            miSQL = "contagrocli"
            N = 1
        ElseIf miRsAux!campo1 = 6 Then
            miSQL = "contagroprov"
            N = 1
        Else
            miSQL = "contaariges"
        End If
        If N = 1 Then miSQL = miSQL & Mid(miRsAux!obser, 6)
        It.SubItems(5) = miSQL
        
        It.Tag = davidNumalbar 'Indice ande esta el "padre" , los datos de cuenta cabecera
        It.SubItems(6) = DBLet(miRsAux!obser, "N")
        It.SubItems(7) = miRsAux!campo1
        It.ToolTipText = It.SubItems(6)
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Label4(1).Caption = "" 'indicador
    davidNumalbar = 0   'reutilizada
    

    
End Sub
 


Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtimporte(Index), 3
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
    txtimporte(Index).Text = Trim(txtimporte(Index).Text)
    If txtimporte(Index).Text = "" Then Exit Sub
'    Select Case Index
'    Case 0
    
        PonerFormatoDecimal txtimporte(Index), 2   'decimal 10,4  en formato decimal
'    Case 1
'        'El uno es obligado el campo
'        If Not PonerFormatoDecimal(txtimporte(Index), 3) Then txtimporte(Index).Text = ""   'importe
'
'    Case 2
'        PonerFormatoDecimal txtimporte(Index), 1   '2 decimales
'    End Select
End Sub



Private Sub txtTextoPlano_GotFocus(Index As Integer)
    ConseguirFoco txtTextoPlano(Index), 3
End Sub

Private Sub txtTextoPlano_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub





Private Sub DatosCuotasPropiasTelefonia()
    
    'CadenaDesdeOtroForm     telefono|numlinea|desc|precio|
    miSQL = RecuperaValor(CadenaDesdeOtroForm, 3)
    Me.txtimporte(0).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
        
    Me.txtTextoPlano(0).Text = miSQL
    Me.txtTextoPlano(0).Enabled = miSQL <> ""
    If Me.txtTextoPlano(0).Text = "" Then
        PonerFoco Me.txtTextoPlano(0)
    Else
        PonerFoco Me.txtimporte(0)
    End If

    
End Sub

Private Sub CargaTelefonosSinConsumo()



Dim N As Integer

    lwTelefoDe.SortKey = 0
    lwTelefoDe.SortOrder = lvwAscending
    lwTelefoDe.Sorted = True
    lwTelefoDe.ListItems.Clear
    Label4(1).Caption = "Leyendo BD" 'indicador
    Label4(1).Refresh
    miSQL = "select tmpnseries.codartic,sclientfno.codclien,nomclien,numlinea from tmpnseries left join sclientfno on sclientfno.idtelefono=tmpnseries.codartic left join sclien on sclientfno.codclien=sclien.codclien "
    miSQL = miSQL & " WHERE tmpnseries.codusu=" & vUsu.codigo & " order by 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
            
        Set It = lwTelefoDe.ListItems.Add
        It.Text = miRsAux!codArtic
        It.SubItems(1) = " "
        It.SubItems(2) = " "
        If Not IsNull(miRsAux!codclien) Then It.SubItems(1) = Format(miRsAux!codclien, "0000")
        If Not IsNull(miRsAux!Nomclien) Then It.SubItems(2) = miRsAux!Nomclien
        
        
        'Febrero 2014
        '1  Sin consumo
        '2  Inactivo
        '5+ Otra compañia
        It.SubItems(4) = " "
        It.SubItems(3) = " "
        
        If miRsAux!numlinea > 0 Then
            'inactivo   otra compañia
            If miRsAux!numlinea = 2 Then
                'INACTIVO
                It.SubItems(3) = "SI"
                It.ForeColor = vbRed
            ElseIf miRsAux!numlinea >= 5 Then
                It.SubItems(4) = "*"
                It.ForeColor = vbBlue
            End If
        End If
        

        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub CargalwPresuElim()
    
    lwPresuElim.ListItems.Clear
    
    
    miSQL = "select numfactu,fecfactu,codclien,nomclien,totalfac from scafac WHERE codtipom ='FAZ'"
    miSQL = miSQL & " AND  " & CadenaDesdeOtroForm & " ORDER BY fecfactu,numfactu"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
            
        Set It = lwPresuElim.ListItems.Add
        It.Text = Format(miRsAux!NumFactu, "0000")
        It.SubItems(1) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        It.SubItems(2) = miRsAux!codclien
        It.SubItems(3) = miRsAux!Nomclien
        It.SubItems(4) = Format(miRsAux!TotalFac, FormatoImporte)

        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub



Private Sub EliminarFacturaFAZ()
    miSQL = " WHERE codtipom ='FAZ' and numfactu =" & lwPresuElim.ListItems(NumRegElim).Text & " AND "
    miSQL = miSQL & " fecfactu = " & DBSet(lwPresuElim.ListItems(NumRegElim).SubItems(1), "F")
    
    conn.Execute "DELETE FROM svenci " & miSQL
    conn.Execute "DELETE FROM slifac " & miSQL
    conn.Execute "DELETE FROM scafac1 " & miSQL
    conn.Execute "DELETE FROM scafac " & miSQL
    
End Sub
