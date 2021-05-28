VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInformesNew6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe "
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11910
   Icon            =   "frmInformesNew6.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameListRiesgoVopt 
      Height          =   4455
      Left            =   6360
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin MSComctlLib.ListView lw1 
         Height          =   3495
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Situacion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1275
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   4200
         Picture         =   "frmInformesNew6.frx":000C
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4680
         Picture         =   "frmInformesNew6.frx":0156
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame FrameListRiesgoVsel 
      Height          =   3375
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   7455
      Begin VB.ComboBox cboTipoASeg 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Height          =   360
         Index           =   1
         Left            =   1290
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
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
         Height          =   360
         Index           =   1
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   1320
         Width           =   4905
      End
      Begin VB.TextBox txtNombre 
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
         Height          =   360
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   795
         Width           =   4905
      End
      Begin VB.TextBox txtCodigo 
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
         Height          =   360
         Index           =   0
         Left            =   1290
         MaxLength       =   4
         TabIndex        =   0
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo crédito"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   1
         Left            =   1005
         ToolTipText     =   "Buscar familia"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   240
         TabIndex        =   22
         Top             =   810
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   23
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   915
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   0
         Left            =   1005
         ToolTipText     =   "Buscar familia"
         Top             =   795
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Aceptar"
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
      Index           =   1
      Left            =   5880
      TabIndex        =   17
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   7320
      TabIndex        =   16
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Imprimir"
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
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame FrameTipoSalida 
      Caption         =   "Tipo de salida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   2280
      TabIndex        =   4
      Top             =   5400
      Width           =   7515
      Begin VB.CommandButton PushButtonImpr 
         Caption         =   "Propiedades"
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
         Left            =   5760
         TabIndex        =   14
         Top             =   585
         Width           =   1635
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   7170
         TabIndex        =   13
         Top             =   1545
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   7170
         TabIndex        =   12
         Top             =   1065
         Width           =   255
      End
      Begin VB.TextBox txtTipoSalida 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1545
         Width           =   5265
      End
      Begin VB.TextBox txtTipoSalida 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1065
         Width           =   5265
      End
      Begin VB.TextBox txtTipoSalida 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   585
         Width           =   3825
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "eMail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   2025
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "PDF"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1545
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Archivo csv"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1065
         Width           =   1515
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   585
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2640
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmInformesNew6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor DAVID  +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Integer

'12.  Familias de Artículos
'15.  Clientes Varios

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    

Private HaDevueltoDatos As Boolean


'Private WithEvents frmMar As frmAlmMarcas 'marcas
'Private WithEvents frmAlm As frmAlmAlPropios 'almacenes propios
'Private WithEvents frmTArt As frmAlmTipoArticulo 'tipo de articulos
'Private WithEvents frmTUni As frmAlmTipoUnidad 'tipo de unidad
'Private WithEvents frmUbi As frmAlmUbicaciones 'ubicaciones
'Private WithEvents frmCat As frmAlmCategorias 'categorias
'Private WithEvents frmFam As frmBasico2 'familias
'Private WithEvents frmProv As frmBasico2 'Proveedores
Private WithEvents frmMtoAgente As frmBasico2 'Basico2
Attribute frmMtoAgente.VB_VarHelpID = -1
'Private WithEvents frmMtoClientes As frmBasico2

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1


Dim miSQL As String




 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean



'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe

Dim NombreRPT As String


'Los reports
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadNombreRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private ConSubInforme As Boolean 'Si el informe tiene subreports

Private cadPDFrpt As String 'Nombre del informe a enviar por email
Private vMostrarTree As Boolean
Private ExportarPDF As Boolean
Private SoloImprimir As Boolean

Private HaPulsadoImprimir As Boolean


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Function MontaSQL() As Boolean
Dim i As Integer
Dim SQL As String
Dim Sql2 As String
Dim Rc As String
Dim RC2 As String

Dim cDesde As String
Dim cHasta As String

    MontaSQL = False
    
    If Not DatosOk Then Exit Function
    
    
    Select Case OpcionListado
    
    Case 1
    
            If Not PonerDesdeHasta2("{sclien.codagent}", "N", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHCodigo=""Agente: ") Then Exit Function
            Sql2 = ""
            SQL = ""
            If Me.cboTipoASeg.ListIndex > 0 Then
                SQL = "Tipo credito: " & Me.cboTipoASeg.Text
                Sql2 = " {sclien.credipriv} = " & cboTipoASeg.ItemData(cboTipoASeg.ListIndex)
                
                AnyadirAFormula cadFormula, Sql2
                AnyadirAFormula cadSelect, Replace(Replace(Sql2, "{", ""), "}", "")
                
                
            End If
            
            Rc = ""
            RC2 = ""
            NumRegElim = 0
            For i = 1 To Me.lw1.ListItems.Count
                If lw1.ListItems(i).Checked Then
                    Rc = Rc & "   -" & lw1.ListItems(i).SubItems(1)
                    RC2 = RC2 & ", " & lw1.ListItems(i).Text
                    NumRegElim = NumRegElim + 1
                End If
            Next
            If NumRegElim = lw1.ListItems.Count Then Rc = "TODAS"
            SQL = Trim(SQL & "        Situacion: " & Rc)
            cadParam = cadParam & "pAgen= """ & SQL & """|"
            numParam = numParam + 1
            
            RC2 = Mid(RC2, 2)
            
            AnyadirAFormula cadFormula, "{sclien.codsitua} IN [" & RC2 & "]"
            AnyadirAFormula cadSelect, "sclien.codsitua IN (" & RC2 & ")"
            
            'NO de varios, NI forpa efectivo y tarjeta
            AnyadirAFormula cadFormula, "{sclien.clivario} = 0"
            AnyadirAFormula cadSelect, "sclien.clivario = 0"
            
            '
            RC2 = "({sforpa.tipforpa}<>0 and   {sforpa.tipforpa}<>6 )"
            AnyadirAFormula cadFormula, RC2
            RC2 = "(tipforpa<>0 and   tipforpa<>6 )"
            AnyadirAFormula cadSelect, RC2
        
        
            'En este caso, riesgo , la tabal es clien forpa
            tabla = "sclien,sforpa"
            cadSelect = " sclien.codforpa = sforpa.codforpa AND " & cadSelect
            
            
    End Select
    
   
        
    Select Case OpcionListado
'        Case 12 'familias/descuentos
'            If chkVarios(1).Value = 1 Then
'                'Hacemos el select este y cargamos tmpinformes
'                If Not CargarDatosFamiliasDtoEnTmp Then Exit Function
'
'
'                cadFormula = "{tmpcommand.codusu} = " & vUsu.Codigo
'                cadSelect = "tmpcommand.codusu = " & vUsu.Codigo
'
'
'                cadParam = cadParam & "Particulares=" & chkVarios(2).Value & "|"
'                numParam = numParam + 1
'
'                tabla = "tmpcommand"
'            Else
'                tabla = "sfamia"
'            End If
'
    End Select
    
    
    MontaSQL = True
    
End Function



Private Function DatosOk() As Boolean
Dim SQL As String
Dim B As Boolean
Dim i As Integer

    B = True
    
    Select Case OpcionListado
        Case 1
            SQL = ""
            For i = 1 To Me.lw1.ListItems.Count
                If lw1.ListItems(i).Checked Then SQL = "OK"
            Next
            
            If SQL = "" Then
                MsgBox "Seleccione alguna situacion", vbExclamation
                B = False
            End If
    End Select
    
    DatosOk = B

End Function


Private Sub cboTipoASeg_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAccion_Click(Index As Integer)

    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV
    
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal
    End If

End Sub

Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    Select Case OpcionListado
        Case 1
              
              cadTitulo = "Listado riesgo"
              cadNombreRPT = "rRiesgo.rpt"
              cadPDFrpt = cadNombreRPT
                vMostrarTree = True
    End Select
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook OpcionListado
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Sub AccionesCSV()
Dim SQL As String

    'Monto el SQL
    Select Case OpcionListado
        Case 1 'riesgo
            
            SQL = " SELECT `sclien`.`codclien`, `sclien`.`nomclien`, `sagent`.`codagent`, `sagent`.`nomagent`, `ssitua`.`codsitua`, `ssitua`.`nomsitua`, "
            SQL = SQL & " `sclien`.`credisol`, `sclien`.`FechaSol`, `sclien`.`codaseg`, `sclien`.`CreditoConcedido`, `sclien`.`limcredi`, "
            SQL = SQL & " `sclien`.`UtFecrecal`, `sclien`.`riesgoact`, `stipocredito`.`nomTipoCredito`"
            SQL = SQL & " FROM   `sclien` `sclien` INNER JOIN `ariges3`.`ssitua` `ssitua` ON `sclien`.`codsitua`=`ssitua`.`codsitua` "
            SQL = SQL & " INNER JOIN `sagent` `sagent` ON `sclien`.`codagent`=`sagent`.`codagent` "
            SQL = SQL & " LEFT OUTER JOIN `stipocredito` `stipocredito` ON `sclien`.`credipriv`=`stipocredito`.`codTipoCredito` "
        
            If cadSelect <> "" Then SQL = SQL & " WHERE " & cadSelect
            SQL = SQL & " ORDER BY `sagent`.`codagent`, `ssitua`.`codsitua`,sclien.codclien"
    
        
    End Select
    
    
    
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 1
            
        End Select
            
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim N As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me
    Me.Icon = frmPpal.Icon
    
   'IMAGES para busqueda
    For N = 0 To 1
        imgBuscarG(N).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next
    
     
    
        
    '------------------------
    'En DISEÑO -visible deberian estar a false
    Me.FrameListRiesgoVsel.visible = False
    Me.FrameListRiesgoVsel.visible = False
    
    
    
    'En DISEÑO -visible deberian estar a false
    '------------------------
    
    Select Case OpcionListado
        Case 1
        
            tabla = "sclien"
            Me.Caption = "Listado riesgos"
            
            PonerFrameVisibles Me.FrameListRiesgoVsel, FrameListRiesgoVopt
            lw1.Height = FrameListRiesgoVopt.Height - 720
            CargaListView lw1
            
            
            
            CargarCombo_Tabla cboTipoASeg, "stipocredito", "codTipoCredito", "nomTipoCredito", , True
            
    End Select
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    Me.cmdCancel.Cancel = True
    
    
    
    
    
End Sub






Private Sub frmMtoAgente_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub imgBuscarG_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1
    
        miSQL = ""
        Set frmMtoAgente = New frmBasico2
        AyudaAgentesComerciales frmMtoAgente, txtCodigo(Index), , True
        Set frmMtoAgente = Nothing
        If miSQL <> "" Then
            txtCodigo(Index).Text = RecuperaValor(miSQL, 1)
            txtNombre(Index).Text = RecuperaValor(miSQL, 2)
            
        End If
    End Select
    miSQL = ""
    PonerFoco txtCodigo(Index)
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim chk As Boolean
    
    If Index < 2 Then
        chk = Index = 1
        For NumRegElim = 1 To Me.lw1.ListItems.Count
            lw1.ListItems(NumRegElim).Checked = chk
        Next
    End If
    
End Sub

Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
         frmPpal.CommonDialog1.Filter = "*.csv|*.csv"
         
    Else
        frmPpal.CommonDialog1.Filter = "*.pdf|*.pdf"
    End If
    frmPpal.CommonDialog1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmPpal.CommonDialog1.FilterIndex = 1
    frmPpal.CommonDialog1.ShowSave
    If frmPpal.CommonDialog1.FileTitle <> "" Then
        If Dir(frmPpal.CommonDialog1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmPpal.CommonDialog1.FileName
    End If

End Sub





'Private Sub AbrirFrmTArticulos(Indice As Integer)
'    indCodigo = Indice
'    Set frmTArt = New frmAlmTipoArticulo
'    frmTArt.DatosADevolverBusqueda = "0|1|"
'    frmTArt.Show vbModal
'    Set frmTArt = Nothing
'End Sub



Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

'#################################################
'###########    A Ñ A D I D O     ################  DE  NUEVA CONTA DE DAVID
'#################################################
Private Sub PonerDatosPorDefectoImpresion(ByRef formu As Form, SoloImpresora As Boolean, Optional NombreArchivoEx As String)
On Error Resume Next
'        AbiertoOtroFormEnListado = False
        
        formu.txtTipoSalida(0).Text = Printer.DeviceName
        If Err.Number <> 0 Then
            formu.txtTipoSalida(0).Text = "No hay impresora instalada"
            Err.Clear
        End If
        If SoloImpresora Then Exit Sub
        
      
        
        formu.txtTipoSalida(1).Text = App.Path & "\Exportar\" & NombreArchivoEx & ".csv"
        formu.txtTipoSalida(2).Text = App.Path & "\Exportar\" & NombreArchivoEx & ".pdf"
        
        If Err.Number <> 0 Then Err.Clear
    
End Sub


'PDF=true   CSV=false
Private Function EliminarDocum(PDF As Boolean) As Boolean
    On Error Resume Next
    If PDF Then
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    Else
        If Dir(App.Path & "\docum.csv", vbArchive) <> "" Then Kill App.Path & "\docum.csv"
    End If
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Err.Clear
        EliminarDocum = False
    Else
        EliminarDocum = True
    End If
End Function


Private Sub ponerLabelBotonImpresion(ByRef BotonAcept As CommandButton, ByRef BotonImpr As CommandButton, SelectorImpresion As Integer)
    On Error GoTo eponerLabelBotonImpresion
    If SelectorImpresion = 0 Then
        BotonAcept.Caption = "&Vista previa"
    Else
        BotonAcept.Caption = "&Aceptar"
    End If
    BotonImpr.visible = SelectorImpresion = 0
    
eponerLabelBotonImpresion:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function ImprimeGeneral() As Boolean
Dim cadPDFrpt As String

    Screen.MousePointer = vbHourglass


'    frmPpal.SkinFramework1.AutoApplyNewWindows = False
'    frmPpal.SkinFramework1.AutoApplyNewThreads = False

  
    HaPulsadoImprimir = False
    cadPDFrpt = cadNombreRPT
    With frmVisReport
        .Informe = App.Path & "\Informes\"
        If ExportarPDF Then
            'PDF
            .Informe = .Informe & cadPDFrpt
        Else
            'IMPRIMIR
            .Informe = .Informe & cadNombreRPT
        End If
        .FormulaSeleccion = cadFormula
        .SoloImprimir = False
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .ConSubInforme = ConSubInforme

        .NumCopias = 1

        .SoloImprimir = SoloImprimir
        .ExportarPDF = ExportarPDF
        .MostrarTree = vMostrarTree
        
        .Show vbModal
        HaPulsadoImprimir = .EstaImpreso
        
      End With
    
    
'     'DAVID
'     frmPpal.SkinFramework1.AutoApplyNewWindows = True
'     frmPpal.SkinFramework1.AutoApplyNewThreads = True
    
End Function

Private Function CopiarFicheroASalida(csv As Boolean, Salida As String, Optional SinMensaje As Boolean) As Boolean
    CopiarFicheroASalida = False
    If Dir(Salida, vbArchive) <> "" Then
        If Not SinMensaje Then
            If Not csv Then
                If MsgBox("Fichero ya existe. ¿Reemplazar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
            End If
        End If
    End If
    
   
    If csv Then
        FileCopy App.Path & "\docum.csv", Salida
    Else
        FileCopy App.Path & "\docum.pdf", Salida
    End If
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Copiando " & Salida
    Else
        If Not SinMensaje Then
            MsgBox "Fichero:  " & Salida & vbCrLf & "Generado con éxito.", vbInformation
        End If
        CopiarFicheroASalida = True
    End If
End Function


Private Sub LanzaProgramaAbrirOutlook(outTipoDocumento As Integer, Optional emailDestinatario As String)
Dim NombrePDF As String
Dim Aux As String
Dim Lanza As String

    On Error GoTo ELanzaProgramaAbrirOutlook

    If Not PrepararCarpetasEnvioMail(True) Then Exit Sub
    
    If Not ExisteARIMAILGES Then Exit Sub

    'Primer tema. Copiar el docum.pdf con otro nombre mas significatiov
    Select Case outTipoDocumento
    Case 1 To 100
        'Marcas
        Aux = "|||||||||||Familias-Descuentos|"
        Aux = Aux & "||ClientesVarios||||||||||"
            
        Aux = RecuperaValor(Aux, outTipoDocumento) & ".pdf"
             
        
    End Select
    NombrePDF = App.Path & "\temp\" & Aux
    If Dir(NombrePDF, vbArchive) <> "" Then Kill NombrePDF
    FileCopy App.Path & "\docum.pdf", NombrePDF
    
    Aux = FijaDireccionEmail(outTipoDocumento)
    If Aux = "" And emailDestinatario <> "" Then Aux = emailDestinatario
    Lanza = Aux & "|"
    Aux = ""
    Select Case outTipoDocumento
        
    Case 1 To 100
        Aux = "|||||||||||FamiliasDescuentos|"
        Aux = Aux & "||ClientesVarios||||||||||"
        '--------------------------------------------------
        Aux = RecuperaValor(Aux, outTipoDocumento)
        
        
    End Select
    Aux = vEmpresa.nomresum & ". " & Aux
    
    Lanza = Lanza & Aux & "|"
    
    'Aqui pondremos lo del texto del BODY
    Aux = ""
    Lanza = Lanza & Aux & "|"
    
    
    'Envio o mostrar
    Lanza = Lanza & "0"   '0. Display   1.  send
    
    'Campos reservados para el futuro
    Lanza = Lanza & "||||"
    
    'El/los adjuntos
    Lanza = Lanza & NombrePDF & "|"
    
    Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Aux, vbNormalFocus
    
    Exit Sub
ELanzaProgramaAbrirOutlook:
    MuestraError Err.Number, Err.Description
End Sub

Private Function FijaDireccionEmail(outTipoDocumento As Integer) As String
Dim campoemail As String
Dim otromail As String


    FijaDireccionEmail = ""
    campoemail = ""
    
'    If outTipoDocumento < 50 Then
''            'Para provedores
''            If outTipoDocumento = 51 Or outTipoDocumento = 52 Or outTipoDocumento = 53 Then
''                campoemail = "maiprov1"
''                otromail = "maiprov2"
''            Else
''                campoemail = "maiprov2"
''                otromail = "maiprov1"
''            End If
''            campoemail = DevuelveDesdeBDNew(cpconta, "proveedor", campoemail, "codprove", Me.outCodigoCliProv, "N", otromail)
'            If campoemail = "" Then campoemail = otromail
'        Else
'            'Para Socios
'            If outTipoDocumento >= 100 Then
'                campoemail = "maisocio"
'                otromail = "maisocio"
'            Else
'                campoemail = "maisocio"
'                otromail = "maisocio"
'            End If
''            campoemail = DevuelveDesdeBDNew(cAgro, "rsocios", campoemail, "codsocio", Me.outCodigoCliProv, "N") ' , otromail)
'            If campoemail = "" Then campoemail = otromail
'        End If
'    End If
    FijaDireccionEmail = campoemail
End Function


Private Function GeneraFicheroCSV(cadSQL As String, Salida As String, Optional OcultarMensajeCreacionCorrecta As Boolean) As Boolean
Dim NF As Integer
Dim i  As Integer

    On Error GoTo eGeneraFicheroCSV
    GeneraFicheroCSV = False
    
    
    If Dir(Salida, vbArchive) <> "" Then
        If MsgBox("El fichero ya existe. ¿Sobreescribir?", vbQuestion + vbYesNo) <> vbYes Then Exit Function
    End If
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Ningun dato generado", vbExclamation
        cadSQL = ""
    Else
        NF = FreeFile
        Open App.Path & "\docum.csv" For Output As #NF
        'Cabecera
        cadSQL = ""
        For i = 0 To miRsAux.Fields.Count - 1
            cadSQL = cadSQL & ";""" & miRsAux.Fields(i).Name & """"
        Next i
        Print #NF, Mid(cadSQL, 2)
    
    
        'Lineas
        While Not miRsAux.EOF
            cadSQL = ""
            For i = 0 To miRsAux.Fields.Count - 1
                cadSQL = cadSQL & ";""" & DBLet(miRsAux.Fields(i).Value, "T") & """"
            Next i
            Print #NF, Mid(cadSQL, 2)
            
            
            
            miRsAux.MoveNext
        Wend
        cadSQL = "OK"
    End If
    miRsAux.Close
    Close #NF

    If cadSQL = "OK" Then
        If CopiarFicheroASalida(True, Salida, OcultarMensajeCreacionCorrecta) Then GeneraFicheroCSV = True
    End If
    
    Exit Function
eGeneraFicheroCSV:
    MuestraError Err.Number, Err.Description
End Function


Private Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadSelect = ""
    cadParam = "|"
    numParam = 0
    cadNombreRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub



Private Function PonerDesdeHasta2(campo As String, Tipo As String, ByRef Desde As TextBox, ByRef DesD As TextBox, ByRef Hasta As TextBox, ByRef DesH As TextBox, param As String) As Boolean
Dim devuelve As String
Dim Cad As String
Dim Subtipo As String 'F: fecha   N: numero   T: texto  H: HORA



    PonerDesdeHasta2 = False
    
    Select Case Tipo
    Case "F", "FEC"
        'Campos fecha
        Subtipo = "F"
    
    Case "N"
        'concepto
        Subtipo = "N"
        
    Case "T"
        Subtipo = "T"
        
    End Select
    
    devuelve = CadenaDesdeHasta(CStr(Desde), CStr(Hasta), campo, Subtipo)
    If devuelve = "Error" Then
        PonerFoco Desde
        Exit Function
    End If
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If devuelve = "" Then
        PonerDesdeHasta2 = True
        Exit Function
    End If
    
    'QUITO LAS LLAVES
    devuelve = Replace(devuelve, "{", "")
    devuelve = Replace(devuelve, "}", "")
    
    If Subtipo <> "F" And Subtipo <> "FH" Then
        'Fecha para Crystal Report

        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(Desde.Text, Hasta.Text, campo, Subtipo)
        Cad = Replace(Cad, "{", "")
        Cad = Replace(Cad, "}", "")
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH2(param, Desde, Hasta, DesD, DesH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta2 = True
    End If
End Function


Private Function AnyadirParametroDH2(Cad As String, ByRef TextoDESDE As TextBox, TextoHasta As TextBox, ByRef TD As TextBox, ByRef TH As TextBox) As String
On Error Resume Next
    
    
    If Not TextoDESDE Is Nothing Then
         If TextoDESDE.Text <> "" Then
            Cad = Cad & "desde " & TextoDESDE.Text
            If TD.Text <> "" Then Cad = Cad & " - " & TD.Text
        End If
    End If
    If Not TextoHasta Is Nothing Then
        If TextoHasta.Text <> "" Then
            Cad = Cad & "  hasta " & TextoHasta.Text
            If TH.Text <> "" Then Cad = Cad & " - " & TH.Text
        End If
    End If
    
    AnyadirParametroDH2 = Cad
    If Err.Number <> 0 Then Err.Clear
End Function




Private Function ExisteARIMAILGES()
Dim SQL As String

    If Dir(App.Path & "\ArimailGes.exe") = "" Then
        MsgBox "No existe el programa ArimailGes.exe. Llame a Ariadna.", vbExclamation
        ExisteARIMAILGES = False
    Else
        ExisteARIMAILGES = True
    End If
End Function





'la suma de fra_select y fram_imporsar sera el alto del form
' la suma del fra_select + fr_opt(si existe) será el with
Private Sub PonerFrameVisibles(ByRef FrameSel As Frame, ByRef FrameOpt As Frame)
Dim Alt As Integer
    FrameSel.Top = 90
    FrameTipoSalida.Left = 90
    FrameSel.Left = FrameTipoSalida.Left
    FrameSel.Width = FrameTipoSalida.Width
    FrameTipoSalida.Top = FrameSel.Top + FrameSel.Height + 120
    cmdAccion(0).Top = FrameTipoSalida.Top + FrameTipoSalida.Height + 90
    Me.Height = cmdAccion(0).Top + cmdAccion(0).Height + 540
    cmdAccion(1).Top = cmdAccion(0).Top
    cmdCancel.Top = cmdAccion(0).Top
    
    'ancho
    If Not FrameOpt Is Nothing Then
        FrameOpt.Top = FrameSel.Top
        FrameOpt.Left = FrameSel.Left + FrameSel.Width + 240
        FrameOpt.Height = FrameSel.Height + FrameTipoSalida.Height + 120
        Me.Width = FrameOpt.Left + FrameOpt.Width + 240
        FrameOpt.visible = True
    Else
        Me.Width = FrameSel.Width + 360
        
    End If
    
    cmdCancel.Left = Me.Width - cmdCancel.Width - 240
    cmdAccion(1).Left = cmdCancel.Left - 240 - cmdAccion(1).Width
    
    
    
    FrameSel.visible = True
End Sub




Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
        Case 0, 1
            KEYBusquedaG KeyAscii, Index 'agentes
            
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

'Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
'    KeyAscii = 0
'    imgBuscar_Click (Indice)
'End Sub

Private Sub KEYBusquedaG(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscarG_Click (Indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Cad = ""
    Select Case Index
        Case 0, 1
        
                If PonerFormatoEntero(txtCodigo(Index)) Then
                    Cad = PonerNombreDeCod(txtCodigo(Index), conAri, "sagent", "nomagent", "codagent", , "N")
                    txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
                Else
                    txtCodigo(Index).Text = ""
                End If
                txtNombre(Index).Text = Cad
                
'
'                Case 1 ' marcas
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "smarca", "nommarca", "codmarca", , "N")
'                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
'                Case 2 ' almacenes propios
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "salmpr", "nomalmac", "codalmac", , "N")
'                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
'                Case 3 ' tipos de unidad
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sunida", "nomunida", "codunida", , "N")
'                Case 4 ' tipos de articulos
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "stipar", "nomtipar", "codtipar", , "T")
'                Case 20 ' actividades
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sactiv", "nomactiv", "codactiv", , "N")
'                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
'                Case 21 ' zonas
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "szonas", "nomzonas", "codzonas", , "N")
'                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
'                Case 22 ' rutas
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "srutas", "nomrutas", "codrutas", , "N")
'                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
'                Case 23 ' categorias
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "scateg", "descateg", "codcateg", , "T")
'                Case 24 ' tarifas
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "starif", "nomlista", "codlista", , "N")
'                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
'                Case 27 ' situaciones
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "ssitua", "nomsitua", "codsitua", , "N")
'                    If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
'                Case 58 ' proveedores
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sprove", "nomprove", "codprove", , "N")
'                Case 110 ' ubicaciones
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "subica", "nomubica", "codubica", , "T")
'                Case 999 ' incidencias
'                    txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sincid", "nomincid", "codincid", , "T")
    End Select
        
        
  
  
End Sub



Private Sub CargaListView(lwX As ListView)
Dim IT
    On Error GoTo eCargaListView

    lwX.ListItems.Clear
    
    If OpcionListado = 1 Then
        miSQL = "select codsitua codigo ,nomsitua descripcion ,"
        miSQL = miSQL & " if(codsitua=" & vParamAplic.SituacionBloqueoOpAseg & ",0, "
        miSQL = miSQL & " if(codsitua=" & vParamAplic.SituacionBloqueoOpAsegSinbloq & ",1,3)) "
        miSQL = miSQL & ", if(codsitua=" & vParamAplic.SituacionBloqueoOpAseg & ",1, "
        miSQL = miSQL & " if(codsitua=" & vParamAplic.SituacionBloqueoOpAsegSinbloq & ",1,0)) marcar from ssitua order by 3,1 "
    End If
    
    Set miRsAux = New ADODB.Recordset
    
    'EL SQL llevará
    '    codigo --
    '    descripcion
    '    marcar
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        miSQL = miRsAux!Codigo
        Set IT = lwX.ListItems.Add(, , miSQL)
        IT.SubItems(1) = miRsAux!Descripcion
        IT.Checked = miRsAux!Marcar = 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
eCargaListView:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
    Set miRsAux = Nothing
End Sub
