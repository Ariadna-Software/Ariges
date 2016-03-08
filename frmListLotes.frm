VERSION 5.00
Begin VB.Form frmListLotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro Trazabilidad"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Index           =   1
      Left            =   2520
      MaxLength       =   15
      TabIndex        =   2
      Tag             =   "Lote Hasta|T|S|||slifac1|numlote||N|"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Index           =   0
      Left            =   2520
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "Lote desde|T|S|||slifac1|numlote||N|"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "fecfactu|F|N|||scafac1|fecfactu||N|"
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha proceso"
      Height          =   255
      Index           =   15
      Left            =   960
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      Caption         =   "Lote"
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
      Left            =   3120
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   2400
      Picture         =   "frmListLotes.frx":0000
      ToolTipText     =   "Buscar fecha"
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Registro Trazabilidad Lotes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmListLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public DatosADevolverBusqueda2 As String
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Private WithEvents frmT As frmAdmTrabajadores
'Private WithEvents frmAl As frmAlmAlPropios
'Private WithEvents frmO As frmTallEntrada
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1



Private cadB As String
Private Cad As String
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim PrimeraVez As Boolean
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private Modo As Byte


Private Sub cmdAceptar_Click()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim cadAux As String
Dim bol As Boolean
    InicializarVbles
    cadNomRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "50", "N")
    '------ > Listado 50 = FonTraza.rpt
    
    
    If Not PonerFormulaYParametrosInf9() Then Exit Sub
        'comprobar que hay datos para mostrar en el Informe
    cadAux = "{slifac}"
    If Not HayRegParaInforme(cadAux, cadSelect) Then Exit Sub
    conSubRPT = False
   
    
    
    
    LlamarImprimir
    Unload Me
   
    
End Sub

Private Sub cmdCancelar_Click()
txtCodigo(0).Text = ""
txtCodigo(1).Text = ""
txtCodigo(2).Text = Date

PonerFoco txtCodigo(2)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
Dim encontrado As String
Dim encontrado2 As String
      If PrimeraVez Then
        PrimeraVez = False
        encontrado = DevuelveDesdeBD(conAri, "codalmac", "straba", "login", vUsu.Login, "T")
        encontrado2 = DevuelveDesdeBD(conAri, "nomalmac", "salmpr", "codalmac", encontrado, "T")
        txtCodigo(0).Text = ""
        txtCodigo(1).Text = ""
        txtCodigo(2).Text = Date
        PonerFoco txtCodigo(2)
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    cadTitulo = ""
    cadNomRPT = ""
    cadTitulo = "REGISTRO TRAZABILIDAD LOTES FINALES"
    conSubRPT = True
    Modo = 0
End Sub

Private Sub frmAl_DatoSeleccionado(CadenaSeleccion As String)
    cadB = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
cadB = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
   Cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmO_DatoSeleccionado(CadenaSeleccion As String)
    cadB = CadenaSeleccion
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    cadB = CadenaSeleccion
End Sub



Private Sub imgFecha_Click(Index As Integer)
    Cad = ""
    Select Case Index
        Case 0 'fecha envio
            Set frmC = New frmCal
            frmC.Fecha = Now
            If txtCodigo(2).Text <> "" Then frmC.Fecha = CDate(txtCodigo(2).Text)
            Cad = ""
            frmC.Show vbModal
            Set frmC = Nothing
            If Cad <> "" Then txtCodigo(2).Text = Cad
            PonerFoco txtCodigo(2)
    End Select
End Sub
Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
    conSubRPT = False
End Sub
Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If Tipo <> "F" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            'parametro Fecha
        
        End If
        
        PonerDesdeHasta = True
    End If
End Function
Private Function DatosOk() As Boolean
Dim b As Boolean

    b = False
    
    ' comprobamos que haya lote
    If txtCodigo(2) <> "" Then
        b = True
    Else
        MsgBox "Debe introducir la fecha del listado", vbExclamation
        PonerFoco txtCodigo(2)
    End If
    
    If txtCodigo(1) >= txtCodigo(0) Then
         b = True
    Else
        MsgBox "Desde N.Lote no puede ser > que Hasta N.Lote", vbExclamation
        PonerFoco txtCodigo(1)
        b = False
    End If

    
    
    DatosOk = b
End Function
Private Function PonerFormulaYParametrosInf9() As Boolean
Dim Cad As String
Dim devuelve As String
Dim i As Byte
Dim posicion As Integer
Dim Status As String

    PonerFormulaYParametrosInf9 = False
    InicializarVbles
    
    'Parametro EMPRESA
     cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
     numParam = 1
       
         
    'Cadena para seleccion Desde y Hasta NUMERO DE LOTE
     If txtCodigo(0).Text <> "" Or txtCodigo(1).Text <> "" Then
        Codigo = "{slifac.numlote}"
        devuelve = "pDHLotes=""Lotes: "
        If Not PonerDesdeHasta(Codigo, "T", 0, 1, devuelve) Then Exit Function
    
    End If
    
    'Parametro FECHA
     cadParam = cadParam & "|pFECHA=""" & txtCodigo(2) & """|"
     numParam = numParam + 1
    
    
    'Parametro DESDE LOTE
     cadParam = cadParam & "|pDESDEL=""" & txtCodigo(0) & """|"
     numParam = numParam + 1
     
    
    'Parametro HASTA LOTE
     cadParam = cadParam & "|pHASTAL=""" & txtCodigo(1) & """|"
     numParam = numParam + 1
     
          
        
    PonerFormulaYParametrosInf9 = True
    
End Function
Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 3002
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub
Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    
     If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
     End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
    End If
    
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function
Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
With txtCodigo(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim encontrado As String
Select Case Index
    Case 0, 1 'lote
        If txtCodigo(Index).Text <> "" Then
           txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        End If
        
    Case 2 'fecha
        If txtCodigo(Index).Text <> "" Then
            PonerFormatoFecha txtCodigo(Index)
        End If
 
End Select
End Sub
