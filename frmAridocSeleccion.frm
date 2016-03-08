VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAridocSeleccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración ARIDOC"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameArid 
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdCancelarProceso 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblIndic 
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   22
         Top             =   1200
         Width           =   4365
      End
      Begin VB.Label lblIndic 
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   4365
      End
   End
   Begin VB.Frame FramFacCli 
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   6795
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtClien 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   0
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtNombreCli 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtfecha 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtfecha 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   5
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   10
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdClientes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtClien 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNombreCli 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
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
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   26
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   25
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nºfactura"
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
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   1320
         TabIndex        =   18
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   18
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   585
      End
      Begin VB.Image imgBuscarCli 
         Height          =   240
         Index           =   0
         Left            =   1875
         Picture         =   "frmAridocSeleccion.frx":0000
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   21
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   30
         Left            =   1440
         TabIndex        =   15
         Top             =   4080
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   1440
         TabIndex        =   14
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Facturas de clientes"
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
         Index           =   8
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   6015
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   37
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   465
      End
      Begin VB.Image imgBuscarCli 
         Height          =   240
         Index           =   1
         Left            =   1875
         Picture         =   "frmAridocSeleccion.frx":0102
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmAridocSeleccion.frx":0204
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1920
         Picture         =   "frmAridocSeleccion.frx":028F
         Top             =   4080
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmAridocSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Public vOpcion As Byte    '1.- Facturas clientes


Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private EstadoProceso As Byte
        '0. Nada
        '1. En marcha
        '2. Tratando de cancelar
        
Private IntAri As CAridocIntegra
Private vTipoDocumento As Byte


Dim SQL As String
Dim H As Integer
Dim W As Integer

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
'Private cadTitulo As String 'Titulo para el frmImprimir
'Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private nomDocu As String
Private OpcionListado As Integer









Private Sub cmdCancel_Click(Index As Integer)
    If EstadoProceso > 0 Then Exit Sub
    
    Unload Me
End Sub

Private Sub cmdCancelarProceso_Click()
    EstadoProceso = 2  'Canlando
End Sub

Private Sub cmdClientes_Click()
    If Not VaciarTemporal Then Exit Sub
    
    'Hariamos comprobaciones
    If Not comprobarFacturas Then Exit Sub
    
    
    If ComprobarAridoc Then
        
            If BloqueoManual("ARIDOC", "1") Then
                Me.lblIndic(0).Caption = "Generando documentos  (I)"
                Me.lblIndic(1).Caption = ""
                EstadoProceso = 1
                Me.Hide
                frmPpal.Hide
                'Ajustamos tamño ventana
                Me.FramFacCli.visible = False
                PonerFrameVisible FrameArid, H, W
                
                
                Me.Show
                    
                
                DoEvents
                Set miRsAux = New ADODB.Recordset
                HacerIntegracion
                
                Set miRsAux = Nothing
                
                Me.Hide
                frmPpal.Show
                FrameArid.visible = False
                PonerFrameVisible Me.FramFacCli, H, W
                Me.Show
                
            End If
            DesBloqueoManual "ARIDOC"
        Me.SetFocus
        If EstadoProceso = 2 Then Unload Me
    End If
    EstadoProceso = 0
End Sub






Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 0, False
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon
    limpiar Me
    FramFacCli.visible = False
    Set IntAri = New CAridocIntegra
    IntAri.Leer vOpcion   'Si no puede leer: intari.codigo =0
    Select Case vOpcion
    Case 1
        'Facturas clientes
        PonerFrameVisible Me.FramFacCli, H, W
        Me.cmdClientes.Enabled = IntAri.codigo > 0
        CargaCombo
    End Select
    cmdCancel(vOpcion).Cancel = True
End Sub

Private Sub PonerFrameVisible(ByRef Fr As Frame, HE As Integer, WI As Integer)
    Fr.visible = True
    HE = Fr.Height
    WI = Fr.Width
    Me.Height = H + 480
    Me.Width = W + 240
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set IntAri = Nothing
End Sub

Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub imgBuscarCli_Click(Index As Integer)
    Set frmCli = New frmFacClientes
    SQL = ""
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.Show vbModal
    Set frmCli = Nothing
    If SQL <> "" Then
        txtClien(Index).Text = RecuperaValor(SQL, 1)
        txtNombreCli(Index).Text = RecuperaValor(SQL, 2)
    End If
    
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    SQL = ""
    If txtfecha(Index).Text <> "" Then frmC.Fecha = CDate(txtfecha(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    If SQL <> "" Then txtfecha(Index).Text = SQL
End Sub

Private Sub txtClien_GotFocus(Index As Integer)
    ConseguirFoco txtClien(Index), 2
End Sub

Private Sub txtClien_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 0, False
End Sub

Private Sub txtClien_LostFocus(Index As Integer)
    txtClien(Index).Text = Trim(txtClien(Index).Text)
    SQL = ""
    If txtClien(Index).Text <> "" Then
        If EsNumerico(txtClien(Index).Text) Then
           SQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtClien(Index).Text)
           If SQL = "" Then SQL = "NO existe el cliente"
        Else
            'Le pongo el foco
            PonerFoco txtClien(Index)
        End If
    End If
    txtNombreCli(Index).Text = SQL
End Sub


Private Sub txtNumero_GotFocus(Index As Integer)
        ConseguirFoco txtNumero(Index), 2
End Sub


Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 0, False
End Sub

Private Sub txtNumero_LostFocus(Index As Integer)
    txtNumero(Index).Text = Trim(txtNumero(Index).Text)
    If txtNumero(Index).Text <> "" Then
        If Not IsNumeric(txtNumero(Index).Text) Then
            MsgBox "Campo numerico: " & txtNumero(Index).Text, vbExclamation
            txtNumero(Index).Text = ""
            PonerFoco txtNumero(Index)
        End If
    End If
End Sub


Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtClien(Index), 2
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 0, False
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)

    txtfecha(Index).Text = Trim(txtfecha(Index).Text)
    If txtfecha(Index).Text <> "" Then
        SQL = txtfecha(Index).Text
        If EsFechaOK(SQL) Then
            txtfecha(Index).Text = SQL
        Else
            MsgBox "Fecha con formato incorrecto: " & txtfecha(Index).Text, vbExclamation
            txtfecha(Index).Text = ""
            PonerFoco txtfecha(Index)
        End If
    End If
End Sub



Private Sub CargaCombo()
    
    Combo1.Clear
    
    SQL = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom LIKE 'F%'"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 1
    
    'El primero lo meto a mano
    Combo1.AddItem "TODAS"
    
    While Not miRsAux.EOF
        SQL = miRsAux!nomtipom
        SQL = Replace(SQL, "Factura", "")
        Combo1.AddItem miRsAux!codtipom & "-" & SQL
        Combo1.ItemData(Combo1.NewIndex) = NumRegElim
        NumRegElim = NumRegElim + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Combo1.ListIndex = 0  'Pongo por defecto al 0
End Sub





'Comun. Comprobacion de parametros ARIDOC
'-------------------------------------------------------
Private Function ComprobarAridoc() As Boolean
    ' Cad = "Select     "
    ComprobarAridoc = False
   
   
    If Conexion_Aridoc_(True) Then
        If IntAri.EstablecerValoresARidoc(ConnConta) Then ComprobarAridoc = True
    End If
    Conexion_Aridoc_ False
      
   
   
   
   
End Function



Private Function VaciarTemporal() As Boolean
Dim MiNombre  As String
    On Error GoTo EMiNombre
    
    MiNombre = Dir(App.Path & "\temp\*.*", vbArchive)   ' Recupera la primera entrada.
    Do While MiNombre <> ""   ' Inicia el bucle.
       ' Ignora el directorio actual y el que lo abarca.
       If MiNombre <> "." And MiNombre <> ".." Then
          Kill App.Path & "\temp\" & MiNombre
       End If
       MiNombre = Dir   ' Obtiene siguiente entrada.
    Loop
    
    VaciarTemporal = True
    Exit Function
EMiNombre:
    VaciarTemporal = False
    MuestraError Err.Number, Err.Description
End Function




Private Sub HacerIntegracion()
Dim Col  As Collection
Dim CA As CAridoc
Dim T1 As Single
Dim indRPT As Byte

    
    
    MontaSQL False
    vTipoDocumento = 0
    Set Col = New Collection
    PB.Value = 0
    miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
    
        'Veremos si hay que cargar los datos del DOCUMENTO
        '-   En el rS deben ir los datos para saber que informe tengo que mostrar
        indRPT = DevuelveIndiceInformeCrystalReport
        If vTipoDocumento <> indRPT Then
            'TEngo que recargar los valores
            InicializarVbles
            If Not PonerParamRPT2(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, 0) Then
                miRsAux.Close
                Exit Sub
            End If
            vTipoDocumento = indRPT  'Para que no recarge cada vez
            
            'Mayo 2012
            'Para Aridoc siempre trabajara con el del pdf
            nomDocu = pPdfRpt
        End If
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
        PB = PB + 1
        T1 = Timer
        Set CA = New CAridoc
        Me.lblIndic(1).Caption = miRsAux.Fields(0) & " " & miRsAux.Fields(1) & "  " & miRsAux.Fields(2) & "     (" & PB.Value & " de " & PB.Max & ")"
        Me.lblIndic(1).Refresh
        
        If GeneraPDFIntegracion(miRsAux, CA) Then
        
            Col.Add CA
        
        
        End If
        Set CA = Nothing
        Me.Refresh
        DoEvents
        Do
        Loop Until Timer - T1 > 1.25 'Un segundo 25 por cada documento. Para que no de
                    
        'Si Ha cancelado
        If EstadoProceso = 2 Then
            miRsAux.MoveLast  'ASI EL SIGUIENTE DA SALIR
        End If
        
        miRsAux.MoveNext
        
        
            
        
    Wend
    miRsAux.Close
    
    'Si Ha cancelado
    If EstadoProceso = 2 Then Exit Sub
    
    If Col.Count = 0 Then
        MsgBox "Ningun datos generado", vbExclamation
        Exit Sub
    End If
    'SEGUIMOS. Tenemos los archivos y en COL los objetos para insertar
    '----------------------------------------------------------
    Me.lblIndic(0).Caption = "Insertando ARIDOC (y II)"
    Me.lblIndic(1).Caption = ""
    
    If Conexion_Aridoc_(True) Then
        InsertarEnAridoc Col
    End If
    Conexion_Aridoc_ False
    
    Set Col = Nothing
End Sub




'-----------------------------------------------------------------
'-----------------------------------------------------------------
'-----------------------------------------------------------------
'-----------------------------------------------------------------
'
'
'      F A C T U R A S          C L I E N T E S
'
'
'-----------------------------------------------------------------
'-----------------------------------------------------------------
'-----------------------------------------------------------------
'-----------------------------------------------------------------
Private Function comprobarFacturas() As Boolean

    comprobarFacturas = False
    
    MontaSQL True
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    If NumRegElim > 0 Then
        comprobarFacturas = True
        PB.Max = NumRegElim
    Else
        MsgBox "Ningun dato a traspasar", vbExclamation
    End If
    Set miRsAux = Nothing
End Function


Private Sub MontaSQL(Comprobacion As Boolean)
    If Comprobacion Then
        SQL = "count(*)"
    Else
        SQL = IntAri.DevuelveSQL
        'Los campos clave que necesite
        'El resto
        SQL = SQL & ",totalfac"
        SQL = SQL & ",numfactu,scafac.codtipom,fecfactu"
    End If
    Select Case vOpcion
    Case 1
        'Facturas
        SQL = "SELECT " & SQL & " from scafac,stipom WHERE "
        'Link de stipoma scafac
        SQL = SQL & "  scafac.codtipom=stipom.codtipom AND "
        'QUe no esten pasadas a ARIDOC
        SQL = SQL & " (aridoc is null) "
        
        'SI lleva codclien
        If txtClien(0).Text <> "" Then SQL = SQL & " AND codclien >= " & txtClien(0).Text
        If txtClien(1).Text <> "" Then SQL = SQL & " AND codclien <= " & txtClien(1).Text
        'Si lleva numero factura
        If txtNumero(0).Text <> "" Then SQL = SQL & " AND numfactu >= " & txtNumero(0).Text
        If txtNumero(1).Text <> "" Then SQL = SQL & " AND numfactu <= " & txtNumero(1).Text
        'Si lleva fecha factura
        If txtfecha(0).Text <> "" Then SQL = SQL & " AND fecfactu >= '" & Format(txtfecha(0).Text, FormatoFecha) & "'"
        If txtfecha(1).Text <> "" Then SQL = SQL & " AND fecfactu <= '" & Format(txtfecha(1).Text, FormatoFecha) & "'"
        'Si lleva en combo
        If Combo1.ListIndex > 0 Then
            'HA seleccinado un tipo de movimiento
            SQL = SQL & " AND scafac.codtipom = '" & Mid(Combo1.Text, 1, 3) & "'"
        End If
        
        
    End Select
End Sub



'-----------------------------------------------------------------------------
' Integracion

Private Function GeneraPDFIntegracion(ByRef RS As ADODB.Recordset, ByRef vAri As CAridoc) As Boolean

    On Error GoTo EGeneraPDFIntegracion
    GeneraPDFIntegracion = False

    'Los 4 primeros campos son los de ARIDOC
    vAri.campo1 = DBLet(RS.Fields(0))
    vAri.campo2 = DBLet(RS.Fields(1))
    vAri.campo3 = DBLet(RS.Fields(2))
    vAri.campo4 = DBLet(RS.Fields(3))
    
    'Las fechas
    vAri.fecha1 = DBLet(RS.Fields(4)) 'Fecha1
    
    
    'El importe1
    vAri.Importe1 = DBLet(RS!TotalFac)
    
    'Segun sea la opcion (facturas o lo que sea, la fecha 2 tb podria ser utilizada
    If vOpcion = 1 Then
        'SQL para la impresion
        
        If RS!codtipom = "FAT" Then
            SQL = DevuelveDesdeBDNew(conAri, "stipom", "letraser", "codtipom", RS!codtipom, "T")  'LEtra de serie
            cadFormula = ""
            SQL = "{tel_cab_factura.Serie} ='" & SQL & "' and {tel_cab_factura.Ano} =" & Year(RS!FecFactu) & " and {tel_cab_factura.NumFact} =" & RS!NumFactu
            
        Else
        
            cadFormula = "{scafac.codtipom}='" & RS!codtipom & "' AND {scafac.numfactu}=" & RS!NumFactu
        
        
            SQL = "{scafac.fecfactu}= Date(" & Year(RS!FecFactu) & "," & Month(RS!FecFactu) & "," & Day(RS!FecFactu) & ")"
        
        End If
        If Not AnyadirAFormula(cadFormula, SQL) Then Exit Function
        
        'Cadwhere para el update de scafac
        vAri.cadUPDATE = "codtipom='" & RS!codtipom & "' AND numfactu=" & RS!NumFactu & " AND fecfactu = '" & Format(RS!FecFactu, FormatoFecha) & "'"
    End If
    'AHora mandamos a generar el PDF
    With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = True
            .opcion = OpcionListado
            .NombreRPT = nomDocu
            .Show vbModal
    End With
    
    If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then
        'Copiamos
        vAri.codigo = PB.Value   'identificador
        FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & vAri.codigo
    End If
    GeneraPDFIntegracion = True
    Exit Function
EGeneraPDFIntegracion:
    MuestraError Err.Number, "GeneraPDFIntegracion"
End Function



Private Function DevuelveIndiceInformeCrystalReport() As Byte


    'En mirsaux tengo los datos necesarios
    Select Case vOpcion
    Case 1
        '----------------- FACTURAS CLIENTES
        OpcionListado = 53  'Impresion de facturas
        Select Case miRsAux!codtipom
        Case "FAZ"
            'Factura B
            DevuelveIndiceInformeCrystalReport = 30
        
        Case "FAV"
            'martin Abril 2012
            ' Quiere que las facturas del TPV lleven el mismo formato
            DevuelveIndiceInformeCrystalReport = 18 'Facturas Clientes TPV
            DevuelveIndiceInformeCrystalReport = 12 'Peticion Martin
        Case "FAT"
        
            DevuelveIndiceInformeCrystalReport = 63  'FAT de telefonia
            
        Case Else
            DevuelveIndiceInformeCrystalReport = 12 'Facturas Clientes
            
            
        End Select
        
    End Select
End Function


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
    
    nomDocu = ""
End Sub





Private Sub InsertarEnAridoc(ByRef Cole As Collection)
Dim I As Integer
Dim vA As CAridoc

    PB.Value = 0
    PB.Max = Cole.Count
    
    lblIndic(0).Caption = "INSERTANDO"
    Me.Refresh
    For I = Cole.Count To 1 Step -1
    
        If (PB.Value Mod 25) = 0 Then
            'Leo de la BD. Para no se
            NumRegElim = IntAri.DevMaxTimagen(ConnConta)
        End If
        NumRegElim = NumRegElim + 1
        PB.Value = PB.Value + 1
        DoEvents
        If EstadoProceso = 2 Then Exit For
            
            
        'Los labels y eso
        
        'Si copia el archivo:
        Set vA = Cole.item(I)
        InsertaUnArchivoAridoc vA, NumRegElim
    Next I
    
End Sub

Private Function InsertaUnArchivoAridoc(ByRef CA As CAridoc, CodigoParaAridoc As Long) As Boolean
Dim Onde As Byte
Dim Tam As Currency



    On Error GoTo EI
    InsertaUnArchivoAridoc = False
    Onde = 0
    SQL = ""
    
    Tam = FileLen(App.Path & "\temp\" & CA.codigo)
    Tam = Round((Tam / 1024), 2)
    
    lblIndic(1).Caption = CA.campo1 & " " & Format(Tam, FormatoImporte) & " Kb"
    lblIndic(1).Refresh
    
    
    SQL = IntAri.GeneraSQLTimagen(CA, NumRegElim, Tam)
    
    ConnConta.Execute SQL
    Onde = 1 'Ejecuta SQL con exito
    
    FileCopy App.Path & "\temp\" & CA.codigo, IntAri.RutaAlmacen & "\" & CStr(CodigoParaAridoc)
    
    
    'Updateamos la tabla scafac
    SQL = "UPDATE scafac set aridoc=" & NumRegElim & " WHERE " & CA.cadUPDATE
    conn.Execute SQL
    InsertaUnArchivoAridoc = True
    Exit Function
EI:
    MuestraError Err.Number
    'Si esta ejecutado el sql , pero NO se copia el archvio
    If Onde = 1 Then
        SQL = "DELETE FROM timagen where codigo = " & NumRegElim
        ConnConta.Execute SQL
    End If
End Function
