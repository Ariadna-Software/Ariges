VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCRMImprimir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresion CRM"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmCRMImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCobros 
      Height          =   425
      Left            =   120
      Picture         =   "frmCRMImprimir.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   615
   End
   Begin VB.CommandButton cmdFecha 
      Height          =   375
      Left            =   840
      Picture         =   "frmCRMImprimir.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cambiar fecha desde"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin MSComctlLib.TreeView TV1 
      Height          =   5295
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9340
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "932"
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tv2 
      Height          =   5295
      Left            =   5640
      TabIndex        =   8
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9340
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tv3 
      Height          =   5295
      Left            =   8760
      TabIndex        =   10
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9340
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   3
      Left            =   7320
      Picture         =   "frmCRMImprimir.frx":0B20
      ToolTipText     =   "seleccionar todos"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   2
      Left            =   6960
      Picture         =   "frmCRMImprimir.frx":0C6A
      ToolTipText     =   "Quitar seleccion"
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Impresion cobros"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   15
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   9720
      Picture         =   "frmCRMImprimir.frx":0DB4
      ToolTipText     =   "Quitar seleccion"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   10080
      Picture         =   "frmCRMImprimir.frx":0EFE
      ToolTipText     =   "seleccionar todos"
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblInd 
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   8760
      TabIndex        =   12
      Top             =   810
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Salen en CRM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   11
      Top             =   810
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Informe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   810
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   960
      Picture         =   "frmCRMImprimir.frx":1048
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmCRMImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public N 'As Node
Private PrimeraVez As Boolean
Private WithEvents frmC2 As frmFacClientes3
Attribute frmC2.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private GuardarConfig As Boolean

Dim J As Integer
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Donde As String




Private vCRM As cCRM
Private HayAlgunDato As Boolean
Private cadParam2 As String   'Para pasarle los parametros al rpt

Dim DatosGuardados As Collection

'Configuracion en el equipo



Private Sub CargaTreeView()

    'EN EL TAG llevara los valores para la cadparam
    ' parametrovisible|parametrofecha|    el de fecha es optativo
    
    'Losw parametros son:
    'Para las fechas
    ' {pDesdeAlbarabSat} {pDesdeAlbaran} {pDesdeAnyo} {pDesdeAvisos} {pDesdeEmail}
    '{pDesdeLlamada} {pDesdeOferta} {pDesdepedido} {pDesdeReclamas} {pDesdeRepara}
    'Para los visibles
    '{pVisAccionesComer} {pVisAlbaranes} {pVisAlbSat} {pVisAvisos} {pVisCobrPdte}
    '{pVisEmails} {pVisFreq} {pVisLlamadas} {pVisMtos} {pVisOfertas}{pVisPedido}
    '{pVisReclamas} {pVisReparas} {pVisVolVenta}
    Configuracion True
    
    CargaAdmon
    CargaComercial
    CargaSAT
End Sub

Private Sub CargaAdmon()

    'Losw parametros son:
    'Para las fechas
    ' {pDesdeAlbarabSat} {pDesdeAlbaran} {} {pDesdeAvisos} {pDesdeEmail}
    '{pDesdeLlamada} {pDesdeOferta} {pDesdepedido} {pDesdeReclamas} {pDesdeRepara}
    'Para los visibles
    '{pVisAccionesComer} {pVisAlbaranes} {pVisAlbSat} {pVisAvisos} {pVisCobrPdte}
    '{pVisEmails} {pVisFreq} {pVisLlamadas} {pVisMtos} {pVisOfertas}{pVisPedido}
    '{pVisReclamas} {pVisReparas} {}
    
    'Departamento administradcio
    Set N = TV1.Nodes.Add(, , "ADM")
    N.Text = "Datos dpto de administración"
    N.Bold = True
    N.Checked = NodoPadreCheckeado(N.Index)    '
    
    FijarNodo3 N, "ADM", "adm1", True, True, "Volumen facturación"
    N.Tag = "pVisVolVenta|pDesdeAnyo|"

    FijarNodo3 N, "ADM", "adm2", False, True, "Facturas pendientes de cobro"
    N.Tag = "pVisCobrPdte||"
    
    
    FijarNodo3 N, "ADM", "adm3", True, False, "Detalle reclamaciones de cobros efectuadas"
    N.Tag = "pVisReclamas|pDesdeReclamas|"

    FijarNodo3 N, "ADM", "adm4", False, False, "Detalle mantenimiento"
    N.Tag = "pVisMtos||"

End Sub

Private Function NodoPadreCheckeado(Indice As Integer) As Boolean
    
    NodoPadreCheckeado = True
    If Not DatosGuardados Is Nothing Then
        If DatosGuardados.Count >= Indice Then NodoPadreCheckeado = RecuperaValor(DatosGuardados(Indice), 1) = "1"
    End If
End Function
Private Sub CargaComercial()
        'Para las fechas
    ' {pDesdeAlbarabSat} {pDesdeAlbaran} {pDesdeAnyo} {pDesdeAvisos} {pDesdeEmail}
    '{pDesdeLlamada} {pDesdeOferta} {pDesdepedido} {pDesdeReclamas} {pDesdeRepara}
    'Para los visibles
    '{pVisAccionesComer} {pVisAlbaranes} {pVisAlbSat} {pVisAvisos} {pVisCobrPdte}
    '{pVisEmails} {pVisFreq} {pVisLlamadas} {pVisMtos} {pVisOfertas}{pVisPedido}
    '{pVisReclamas} {pVisReparas} {pVisVolVenta}
    
    'Departamento administradcio
    Set N = TV1.Nodes.Add(, , "COM")
    N.Text = "Datos dpto de comercial"
    N.Tag = "||"
    N.Bold = True
    N.Checked = NodoPadreCheckeado(N.Index)
     
    FijarNodo3 N, "COM", "com1", True, False, "Detalle ofertas pendientes"
    N.Tag = "pVisOfertas|pDesdeOferta|"
    
    
    
    FijarNodo3 N, "COM", "com2", True, False, "Detalle pedidos pendientes de entregar"
    N.Tag = "pVisPedido|pDesdepedido|"
    
    FijarNodo3 N, "COM", "com3", True, False, "Detalle albaranes pendientes de facturar"
    N.Tag = "pVisAlbaranes|pDesdeAlbaran|"
    
    'Acciones comerciales.
    'Herbelca NO tiene
    If vParamAplic.NumeroInstalacion <> 2 Then
        FijarNodo3 N, "COM", "com6", True, False, "Acciones comerciales "
        N.Tag = "pVisAccionesComer|pDesdeAccComer|"
    End If
    
    
    
    
    
    FijarNodo3 N, "COM", "com4", True, False, "Detalle llamadas"
    N.Tag = "pVisLlamadas|pDesdeLlamada|"
    
            FijarNodo3 N, "com4", "com41", False, False, "Recibidas"
            FijarNodo3 N, "com4", "com42", False, False, "Realizadas"
    
    
    'Herbelca NO tiene
    If vParamAplic.NumeroInstalacion <> 2 Then
    
        FijarNodo3 N, "COM", "com5", True, False, "Detalle correos(eMail)"
        N.Tag = "pVisEmails|pDesdeEmail|"
        
        
                FijarNodo3 N, "com5", "com51", False, False, "Recibidos"
                FijarNodo3 N, "com5", "com52", False, False, "Enviados"
    End If
End Sub



Private Sub CargaSAT()
    
    If Not vParamAplic.Reparaciones Then Exit Sub
    
        'Para las fechas
    ' {pDesdeAlbarabSat} {pDesdeAlbaran} {pDesdeAnyo} {pDesdeAvisos} {pDesdeEmail}
    '{pDesdeLlamada} {pDesdeOferta} {pDesdepedido} {pDesdeReclamas} {pDesdeRepara}
    'Para los visibles
    '{pVisAccionesComer} {pVisAlbaranes} {pVisAlbSat} {pVisAvisos} {pVisCobrPdte}
    '{pVisEmails} {pVisFreq} {pVisLlamadas} {pVisMtos} {pVisOfertas}{pVisPedido}
    '{pVisReclamas} {pVisReparas} {pVisVolVenta}
    

    'Departamento administradcio
    Set N = TV1.Nodes.Add(, , "SAT")
    N.Text = "Datos dpto de S.A.T."
    N.Bold = True
    N.Checked = NodoPadreCheckeado(N.Index)
    
    
 
 
    FijarNodo3 N, "SAT", "sat1", False, False, "Frecuencias"
    N.Tag = "pVisFreq||"
    
    FijarNodo3 N, "SAT", "sat2", True, False, "Albaranes reparacion pendientes facturar"
    N.Tag = "pVisAlbSat|pDesdeAlbarabSat|"
    
    FijarNodo3 N, "SAT", "sat3", True, False, "Avisos pendientes de cerrar"
    N.Tag = "pVisAvisos|pDesdeAvisos|"
    
    
    FijarNodo3 N, "SAT", "sat4", True, False, "Equipos pendientes de reparar"
    N.Tag = "pVisReparas|pDesdeRepara|"
    

End Sub




Private Sub cmdCobros_Click()
    ImprimeSub False
End Sub

Private Sub cmdFecha_Click()
    If TV1.Nodes.Count = 0 Then Exit Sub
    If TV1.SelectedItem Is Nothing Then Exit Sub
    
    If Right(TV1.SelectedItem.Text, 1) <> "]" Then
        MsgBox "NO se le asigna fecha a esta opcion", vbExclamation
    Else
        SQL = ""
        J = InStr(1, TV1.SelectedItem, "[")
        If J = 0 Then
            MsgBox "No se ha encotrado la marca de fecha", vbExclamation
        Else
            Donde = Mid(TV1.SelectedItem.Text, J + 1)
            Donde = Mid(Donde, 1, Len(Donde) - 1)
            If Len(Donde) = 4 Then
                'Es AÑO
                J = 0
                Donde = "01/01/" & Donde
            Else
                'Es fecha
                J = 1
            End If
            SQL = ""
            Set frmC = New frmCal
            frmC.Fecha = CDate(Donde)
            frmC.Show vbModal
            Set frmC = Nothing
            If SQL <> "" Then
                
                            'Solo quiero el año
                If J = 0 Then SQL = Year(SQL)
                
                J = InStr(TV1.SelectedItem.Text, "[")
                If J = 0 Then
                    MsgBox ""
            
                Else
                    'Ha retornado dato
                    GuardarConfig = True
                    Donde = Mid(TV1.SelectedItem.Text, 1, J)
                    Donde = Donde & SQL & "]"
                    TV1.SelectedItem.Text = Donde
                    
                    
                    'Si es de ofertas o de albaranes CARGAMOS los datos de nuevo
                    SQL = TV1.SelectedItem.Key
                    If SQL = "com1" Or SQL = "com3" Then CargaDatosAux
                    
                    
                End If
                Donde = ""
                SQL = ""
            
                End If
            End If
        End If
End Sub

Private Sub cmdImprimir_Click()
    ImprimeSub True
End Sub


Private Sub ImprimeSub(Normal As Boolean)

    'el unico control de errores esta aqui
On Error GoTo EcmdImprimir
    
    If Text1.Text = "" Then
        MsgBox "Ponga el cliente", vbExclamation
        PonerFoco Text1
        Exit Sub
    End If
    
    'A ver si esta configurada
    If Normal Then
        NumRegElim = 46
    Else
        NumRegElim = 67
    End If
    pPdfRpt = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", CStr(NumRegElim))
    If pPdfRpt = "" Then
        MsgBox "Falta configurar en informes " & NumRegElim, vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'En estas cargaremos los albaranes, ofertas y facturas seleccionadas
    ejecutar "DELETE FROM tmpinformes WHERE codusu = " & vUsu.codigo, False
    NumRegElim = 0 'contador para tmp con los ofe/ped/alb
    
    
    Set Rs = New ADODB.Recordset
    Set vCRM = New cCRM
    HayAlgunDato = False
    cadParam2 = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    
    HayAlgunDato = True

    
    If Normal Then
        GenerarDatosInformes
    Else
        GenerarDatosCobros
    End If
    
    
    If HayAlgunDato Then
        InsertaDatosBasicos
        LlamarImprimir2 Not Normal
        
        If Normal Then ImprimirDocumentosAuxiliares
        
    End If
        
    
EcmdImprimir:
    If Err.Number <> 0 Then MuestraError Err.Number, Donde & vbCrLf & Err.Description
    Set Rs = Nothing
    Set vCRM = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        Screen.MousePointer = vbHourglass
        PrimeraVez = False
        CargaDatosAux
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    CargaTreeView
    For J = 1 To TV1.Nodes.Count
        'TV1.Nodes(J).Checked = True
        TV1.Nodes(J).EnsureVisible
    Next J

    GuardarConfig = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GuardarConfig Then Configuracion False
End Sub



Private Sub frmC_Selec(vFecha As Date)
    SQL = CStr(vFecha)
End Sub

Private Sub frmC2_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub Image1_Click()
    SQL = ""
    Set frmC2 = New frmFacClientes3
    frmC2.DatosADevolverBusqueda = "0|1|"
    frmC2.Show vbModal
    Set frmC2 = Nothing
    If SQL <> "" Then
        Me.Text1.Text = RecuperaValor(SQL, 1)
        Me.Text2.Text = RecuperaValor(SQL, 2)
        CargaDatosAux
    End If
End Sub

Private Sub imgCheck_Click(Index As Integer)
    If Index < 2 Then

        If tv3.Nodes.Count = 0 Then Exit Sub
        For J = 1 To tv3.Nodes.Count
            tv3.Nodes(J).Checked = Index = 1
        Next J

    Else
        If tv2.Nodes.Count = 0 Then Exit Sub
        For J = 1 To tv2.Nodes.Count
            tv2.Nodes(J).Checked = Index = 1
        Next J
    
    End If

End Sub

Private Sub Text1_GotFocus()
    ConseguirFoco Text1, 3
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus()
    SQL = ""
    Text1.Text = Trim(Text1.Text)
    If Text1.Text <> "" Then
        If Not IsNumeric(Text1.Text) Then
            MsgBox "Codigo cliente numérico: " & Text1.Text, vbExclamation
            Text1.Text = ""
            PonerFoco Text1
        Else
            SQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1.Text)
            If SQL = "" Then
                MsgBox "no existe cliente: " & Text1.Text, vbExclamation
                PonerFoco Text1
    
            End If
        End If
    End If
    Text2.Text = SQL
    CargaDatosAux
    
End Sub

Private Sub TV1_DblClick()
    If TV1.Nodes.Count = 0 Then Exit Sub
    If TV1.SelectedItem Is Nothing Then Exit Sub
    
    If InStr(TV1.SelectedItem.Text, "[") = 0 Then Exit Sub
    
    cmdFecha_Click
End Sub

Private Sub Tv1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim CH As Boolean

    If PrimeraVez Then Exit Sub
    
    
   

    
    If Node.Checked Then
        If Not Node.Parent Is Nothing Then Node.Parent.Checked = True
    End If
    
    CH = Node.Checked
    CheckSubNodo Node, CH, False
    GuardarConfig = True
End Sub


Private Sub CheckSubNodo(ByRef N, Checkar As Boolean, EsElTV2 As Boolean)
Dim no
    
    Set no = N
    no.Checked = Checkar
    If EsElTV2 Then CheckeaTambienEnElTv3 no.Index, Checkar
    Set no = N.Child
    While Not no Is Nothing
        CheckSubNodo no, Checkar, EsElTV2
        Set no = no.Next
    Wend
    
    
    
End Sub


Private Sub CheckeaTambienEnElTv3(Indice As Integer, ChK)
    On Error Resume Next
    tv3.Nodes(Indice).Checked = ChK
    Err.Clear
End Sub


Private Sub LlamarImprimir2(SoloCobros As Boolean)
Dim K As Integer

    With frmImprimir
        .FormulaSeleccion = "{tmpcrmclien.codusu} = " & vUsu.codigo
        
        'Cuantos parametros envio
        NumRegElim = 0
        J = 2
        Do
           K = InStr(J, cadParam2, "|")
           If K > 0 Then
                NumRegElim = NumRegElim + 1
                J = K + 1
            End If
        Loop Until K = 0
        .OtrosParametros = cadParam2
        .NumeroParametros = NumRegElim

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 2018
        .Titulo = "CRM"
        If SoloCobros Then .Titulo = "CRM Cobros"
        .NombreRPT = pPdfRpt
        .NombrePDF = ""
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub



'Generad Datos
Private Sub InsertaDatosBasicos()
Dim Aux As String

    'Si habian metido algun dato...
    SQL = "insert into `tmpcrmclien` (`codusu`,`codclien`,`saldopdte`,saldototal,`nomactiv`,`nomforpa`) values ("
    SQL = SQL & vUsu.codigo & "," & Text1.Text & ","
    
    'Saldo pdte (a fecha NOW
    Aux = "Imp"
    ComprobarCobrosCliente Text1.Text, Now, Aux
    If Aux = "" Or Aux = "Imp" Then Aux = "0"
    SQL = SQL & DBSet(Aux, "N") & ","
    'saldo totoal A fecha 31/12/2222"
    Aux = "Imp"
    ComprobarCobrosCliente Text1.Text, CDate("31/12/2222"), Aux
    If Aux = "" Or Aux = "Imp" Then Aux = "0"
    SQL = SQL & DBSet(Aux, "N") & ","
    
    
    
    Aux = DevuelveDesdeBD(conAri, "nomactiv", "sclien,sactiv", "sclien.codactiv=sactiv.codactiv and codclien", Text1.Text)
    SQL = SQL & DBSet(Aux, "T") & ","
    Aux = DevuelveDesdeBD(conAri, "nomforpa", "sclien,sforpa", "sclien.codforpa=sforpa.codforpa and codclien", Text1.Text)
    SQL = SQL & DBSet(Aux, "T") & ")"
    conn.Execute SQL
End Sub



Private Sub GenerarDatosInformes()

    vCRM.BorrarTemporales
    vCRM.codClien = CLng(Text1.Text)
    vCRM.Codmacta = DevuelveDesdeBD(conAri, "codmacta", "sclien", "codclien", Text1.Text)

    
    
    J = DevuelveIndiceNodo("ADM")
    If Me.TV1.Nodes(J).Checked Then
        GenerarDatosAdmon
    Else
        'PONGO TODOS LOS SUBPARAMETROS A FALSE
        PonerparametrosVisiblesFalse
    End If
    
    
    'Para saber si tiene datos cada secccion
    J = DevuelveIndiceNodo("COM")
    If Me.TV1.Nodes(J).Checked Then
        GenerarDatosComer
    Else
        'PONGO TODOS LOS SUBPARAMETROS A FALSE
        PonerparametrosVisiblesFalse
    End If


    'Para saber si tiene datos cada secccion
    If vParamAplic.Reparaciones Then
        J = DevuelveIndiceNodo("SAT")
        If Me.TV1.Nodes(J).Checked Then
            GenerarDatosSAT
        Else
            'PONGO TODOS LOS SUBPARAMETROS A FALSE
            PonerparametrosVisiblesFalse
        End If
    Else
        cadParam2 = cadParam2 & "pVisFreq=0|pVisAlbSat=0|pVisAvisos=0|pVisReparas=0|"
    End If
    
End Sub




Private Sub GenerarDatosCobros()

    vCRM.BorrarTemporales
    vCRM.codClien = CLng(Text1.Text)
    vCRM.Codmacta = DevuelveDesdeBD(conAri, "codmacta", "sclien", "codclien", Text1.Text)

    'Truco. Marcamos todos los check del primero y  lo volvemos a dejar como estaba
    MarcaDescmarcaAdministracion True
    
    GenerarDatosAdmon
    
    
    MarcaDescmarcaAdministracion False
   
    
End Sub


Private Sub MarcaDescmarcaAdministracion(Leer As Boolean)
Dim N As Node


    
    Set N = TV1.Nodes(1).Child 'ADMON
    If Leer Then
        
        CadenaDesdeOtroForm = CInt(TV1.Nodes(1).Checked) & "|"
        TV1.Nodes(1).Checked = True
        
        While Not (N Is Nothing)
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & CInt(N.Checked) & "|"
            N.Checked = True
            Set N = N.Next
        Wend

    Else
        'Tengo en cadenasql lo que habia leido antes de marcarlo
        NumRegElim = 1
        SQL = RecuperaValor(CadenaDesdeOtroForm, CInt(NumRegElim))
        TV1.Nodes(1).Checked = SQL = "-1"
        While Not (N Is Nothing)
            NumRegElim = NumRegElim + 1
            SQL = RecuperaValor(CadenaDesdeOtroForm, CInt(NumRegElim))
            N.Checked = SQL = "-1"
            
            Set N = N.Next
        Wend
        
    End If
End Sub


Private Sub PonerparametrosVisiblesFalse()
Dim N As Node
    'en TV1(j) tengo el NODO padre
    'Con lo cual, recorrro todos sus hijos, obteneido la cadena param de visible y poneindola a cero
    Set N = TV1.Nodes(J).Child '
    While Not (N Is Nothing)
        SQL = RecuperaValor(N.Tag, 1)
        If SQL <> "" Then cadParam2 = cadParam2 & SQL & "=0|"
        Set N = N.Next
    Wend
End Sub



Private Function DevuelveIndiceNodo(Clave As String) As Integer
Dim i As Integer
    
    For i = 1 To TV1.Nodes.Count
        If TV1.Nodes(i).Key = Clave Then
            DevuelveIndiceNodo = i
            Exit Function
        End If
    Next
    
    'Si llega aqui generaremos un erro
    Err.Raise 512, , "NO se encuentra NODO : " & Clave
End Function


'COmercia
'---------------------------
Private Sub GenerarDatosComer()
Dim cad As String
Dim Contador As Long
Dim F As Date
Dim Procesar As Boolean

    Donde = "Comercial"
    'Volumen facturacion
    J = DevuelveIndiceNodo("com1")
    If HayKprocesarNodo(J, F) Then
        Donde = "Ofertas pendientes"
        
    
    End If
    
    
    J = DevuelveIndiceNodo("com2")
    If HayKprocesarNodo(J, F) Then
        Donde = "Pedidos pendientes"
        
    End If
    
    
    J = DevuelveIndiceNodo("com3")
    If HayKprocesarNodo(J, F) Then
        Donde = "Albaranes pdtes"
        
    End If
    
    
    'Acciones comerciales
    If vParamAplic.NumeroInstalacion <> 2 Then
        J = DevuelveIndiceNodo("com6")
        If HayKprocesarNodo(J, F) Then
            Donde = "Acciones comerciales"
        End If
        
    End If
    
    Contador = 0
    
    J = DevuelveIndiceNodo("com4")
    If HayKprocesarNodo(J, F) Then
        Donde = "Llamadas"

        
        'Si no quiere las recibidas
        J = DevuelveIndiceNodo("com41")
        If HayKprocesarNodo(J, F) Then
            'insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,`trabajador`,`adjuntos`) values ( '1','0','','',NULL,NULL,NULL,NULL)
            SQL = "select feholla,usuario,nomllama1,observac,codtraba,nomtraba from sllama,sllama1 "
            SQL = SQL & "  where sllama.codllama1 = sllama1.codllama1"
            SQL = SQL & " and codclien=" & vCRM.codClien
            SQL = SQL & " AND feholla>=" & DBSet(F, "F")
            
            
            
            
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                NumRegElim = NumRegElim + 1
                SQL = "insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,"
                SQL = SQL & "`trabajador`,`adjuntos`) values ( " & vUsu.codigo & "," & NumRegElim & ",0,"
                SQL = SQL & DBSet(Rs!feholla, "FH") & ","
                'En sllama siempre son RECIBIDAS
                SQL = SQL & "'Recibida',"
                cad = DBLetMemo(Rs!observac)
                cad = Replace(cad, vbCrLf, " ")
                SQL = SQL & DBSet(cad, "T", "S") & ","
                'Trabajador
                SQL = SQL & DBSet(Rs!NomTraba, "T") & ","
                'En adjuntos guardare el tipop llamada
                SQL = SQL & DBSet(Rs!nomllama1, "T") & ")"
                
                conn.Execute SQL
                Rs.MoveNext
            Wend
            Rs.Close
            'Ha metido algun dato
           ' If NumRegElim > 0 Then comer(4) = True   'tiene datos
            Contador = NumRegElim
        End If
            
        'Si no quiere las realizadas
        J = DevuelveIndiceNodo("com42")
        If HayKprocesarNodo(J, F) Then
            'insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,`trabajador`,`adjuntos`) values ( '1','0','','',NULL,NULL,NULL,NULL)
            SQL = "select fechora ,usuario,nomtraba ,observaciones from"
            SQL = SQL & " scrmacciones left join straba on scrmacciones.codtraba=straba.codtraba "
            SQL = SQL & " WHERE scrmacciones.tipo=1  and codclien= " & vCRM.codClien
            SQL = SQL & " AND fechora>=" & DBSet(F, "F")
            
            
            
            
            
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                NumRegElim = NumRegElim + 1
                SQL = "insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,"
                SQL = SQL & "`trabajador`,`adjuntos`) values ( " & vUsu.codigo & "," & NumRegElim & ",0,"
                SQL = SQL & DBSet(Rs!fechora, "FH") & ","
                'En sllama siempre son RECIBIDAS
                SQL = SQL & "'Realizada',"
                cad = DBLetMemo(Rs!Observaciones)
                cad = Replace(cad, vbCrLf, " ")
                SQL = SQL & DBSet(cad, "T", "S") & ","
                'Trabajador
                SQL = SQL & DBSet(Rs!NomTraba, "T") & ","
                'En adjuntos guardare el tipop llamada
                SQL = SQL & "NULL)"
                
                conn.Execute SQL
                Rs.MoveNext
            Wend
            Rs.Close
            'Ha metido algun dato
            'If NumRegElim > Contador Then comer(4) = True   'tiene datos
            Contador = NumRegElim
        End If
        
    End If
    
    
    
    If vParamAplic.NumeroInstalacion <> 2 Then
        J = DevuelveIndiceNodo("com5")
        If HayKprocesarNodo(J, F) Then
            Donde = "Emails"
            
            
            'Si no quiere las recibidas
            NumRegElim = 0
            J = DevuelveIndiceNodo("com51")
            If TV1.Nodes(J).Checked Then NumRegElim = 1
            
            J = DevuelveIndiceNodo("com51")
            If TV1.Nodes(J).Checked Then NumRegElim = NumRegElim + 2
            
            If NumRegElim > 0 Then
                    'Ha selecionado alguno de los dos, o los dos
                    
                    'insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,`trabajador`,`adjuntos`) values ( '1','0','','',NULL,NULL,NULL,NULL)
                    SQL = "select fechahora,enviado,email,asunto,adjuntos from scrmmail"
                    SQL = SQL & " WHERE codclien=" & vCRM.codClien
                     SQL = SQL & " AND fechahora>=" & DBSet(F, "F")
                    If NumRegElim = 1 Or NumRegElim = 2 Then
                        cad = "1"
                        If NumRegElim = 2 Then cad = "0"
                        'Ha selecionado solo una de las dos
                        SQL = SQL & " AND enviado = " & cad
                    End If
                    NumRegElim = Contador
                    
                
                
                
                    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not Rs.EOF
                        NumRegElim = NumRegElim + 1
                        SQL = "insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,"
                        SQL = SQL & "`trabajador`,`adjuntos`) values ( " & vUsu.codigo & "," & NumRegElim & ",1,"  '1.email
                        SQL = SQL & DBSet(Rs!fechahora, "FH") & ","
                        'En sllama siempre son RECIBIDAS
                        If Val(Rs!Enviado) = 1 Then
                            SQL = SQL & "'Enviado',"
                        Else
                            SQL = SQL & "'Recibido',"
                        End If
                        cad = DBLetMemo(Rs!asunto)
                        cad = Replace(cad, vbCrLf, " ")
                        SQL = SQL & DBSet(cad, "T", "S") & ","
                        'Trabajador
                        SQL = SQL & DBSet(Rs!email, "T", "S") & ","
                        'En adjuntos guardare el tipop llamada
                        cad = "'*'"
                        If DBLet(Rs!adjuntos, "T") = "" Then cad = "NULL"
                        SQL = SQL & cad & ")"
                        
                        conn.Execute SQL
                        Rs.MoveNext
                    Wend
                    Rs.Close
                    'Ha metido algun dato
                    'If NumRegElim > Contador Then comer(5) = True   'tiene datos
                    Contador = NumRegElim
            End If
                
            
    
            
        End If ' de procesarnodo
    End If  'de numinstal
    
End Sub










Private Sub GenerarDatosAdmon()
Dim Impor1 As Currency
Dim Base As Currency
Dim cad As String
Dim Aux As String
Dim F As Date

    Donde = "Administracion"
    'Volumen facturacion
    J = DevuelveIndiceNodo("adm1")
    If HayKprocesarNodo(J, F) Then
        Donde = "Volumen fact."
        
        'Volumen facturacion
        SQL = "select year(fecfactu) anyo,sum(totalfac) totalfac from scafac "
        'SEPTIEMBE 2011. Quito FRT del select
        'SQL = SQL & " where codclien=" & Text1.Text & " and codtipom <>'FAZ' and codtipom<>'FRT' "
        SQL = SQL & " where codclien=" & Text1.Text & " and codtipom <>'FAZ'"
        SQL = SQL & " AND fecfactu>='" & Format(F, FormatoFecha) & "'"
        'Aqui va lo de ultimos años
        SQL = SQL & " group by 1 order by 1,2"
        
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        
        While Not Rs.EOF
            cad = ""
        
            NumRegElim = NumRegElim + 1
            Impor1 = DBLet(Rs!TotalFac, "N")
            
            SQL = "insert into `tmpcrmtesor` (`codusu`,`codigo`,`importe`,`anyotxt`,`variacion`)"
            SQL = SQL & " values (" & vUsu.codigo & "," & NumRegElim & "," & TransformaComasPuntos(CStr(Impor1)) & ",'"
            
            If Val(Rs!Anyo) = Year(Now) Then
                'Valor actual.
                SQL = SQL & "actual',"
                'Cambio la base para comprar con el mismo periodo del actual
                
                'Cad = "codtipom <>'FAZ' and codtipom<>'FRT' and "
                cad = "codtipom <>'FAZ' and "
                cad = cad & " fecfactu>='" & Year(Now) - 1 & "-01-01' and "
                cad = cad & " fecfactu<='" & Year(Now) - 1 & "-" & Format(Now, "mm-dd") & "' AND codclien "
                cad = DevuelveDesdeBD(conAri, "sum(totalfac)", "scafac", cad, Text1.Text)
                If cad = "" Then cad = "0"
                Base = CCur(cad)
                If NumRegElim > 1 And Base <> 0 Then
                    Impor1 = CStr(((100 * Impor1) / Base) - 100)
                    cad = Format(Impor1, FormatoPorcen) & "% sobre misma fecha año anterior"
                Else
                    cad = ""
                End If
            Else
                'Otro año cualquiera
                 SQL = SQL & Rs!Anyo & "',"
                If NumRegElim > 1 And Base <> 0 Then
                    Impor1 = CStr(((100 * Impor1) / Base) - 100)
                    cad = Format(Impor1, FormatoPorcen) & "%"
                End If
                 
            End If
            Base = DBLet(Rs!TotalFac, "N")
            SQL = SQL & "'" & cad & "')"
          

            conn.Execute SQL
            Rs.MoveNext
        Wend
        Rs.Close
        'If NumRegElim > 0 Then admon(1) = True
    
    
    End If
    
    
    J = DevuelveIndiceNodo("adm2")
    If HayKprocesarNodo(J, F) Then
        Donde = "Cobros pendientes"
        'insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ( '1','0','0','','','',NULL,NULL)
        SQL = "SELECT scobro.*,nomforpa FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        SQL = SQL & " WHERE scobro.codmacta = '" & vCRM.Codmacta & "'"
        'JUNIO 2010
        'PONGO Toooodos los vtos es decir, comento la linea inferior
        'SQL = SQL & " AND fecvenci <= ' " & Format(Now, FormatoFecha) & "' "
        'SQL = SQL & " AND (sforpa.tipforpa between 0 and 3) ORDER BY fecvenci desc"
        SQL = SQL & "  AND recedocu=0 ORDER BY fecvenci desc"
        
        NumRegElim = 0
        Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        Base = 0
        Impor1 = 0
        
        While Not Rs.EOF
              'trozo copiado d ela funcion de ver cobros pdtes
          If DBLet(Rs!Devuelto, "N") = 1 Then
                'SALE SEGURO (si no esta girado otra vez ¿no?
                'Si esta girado otra vez tendra impcobro, con lo cual NO tendra diferencia de importes
                Impor1 = Rs!ImpVenci + DBLet(Rs!gastos, "N") - DBLet(Rs!impcobro, "N")
                
            Else
                'Si esta recibido NO lo saco
                If Val(Rs!recedocu) = 1 Then
                    Impor1 = 0
                Else
                    'NO esta recibido. Si tiene diferencia
                    Impor1 = Rs!ImpVenci + DBLet(Rs!gastos, "N") - DBLet(Rs!impcobro, "N")
            
                End If
          End If
          If Impor1 <> 0 Then
                NumRegElim = NumRegElim + 1
                SQL = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,"
                SQL = SQL & "`importe`,`observa`,forpa) values ( "
                SQL = SQL & vUsu.codigo & "," & NumRegElim & ",0,'"
                SQL = SQL & Rs!numSerie & Format(Rs!Codfaccl, "000000")
                If Rs!FecVenci < Now Then SQL = SQL & " *"
                SQL = SQL & "','" & Format(Rs!fecfaccl, FormatoFecha)
                SQL = SQL & "','" & Format(Rs!FecVenci, FormatoFecha) & "'," & TransformaComasPuntos(CStr(Impor1)) & ","
                'Antes la observa era NULL, ahora llevare el Depto
                If IsNull(Rs!Departamento) Then
                    Aux = "NULL"
                Else
                    Aux = "codclien = " & vCRM.codClien & " AND coddirec  "
                    Aux = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Aux, CStr(Rs!Departamento))
                    If Aux = "" Then Aux = Rs!Departamento
                    Aux = "'" & DevNombreSQL(Aux) & "'"
                    
                End If
                SQL = SQL & Aux
                'Mayo 2010
                'Con forma de pago
                SQL = SQL & ",'" & Format(Rs!codforpa, "000") & " - " & DevNombreSQL(Rs!nomforpa) & "')"
                conn.Execute SQL
          End If
          Rs.MoveNext

            
        
        Wend
        Rs.Close
        
        
        'Marzo 2011
        'Tambien sacare el riesgo. Habra que configurar el rpt de cada uno
        '----------------------------------------------------------------
        Donde = "Riesgo tesoreria"
        'insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ( '1','0','0','','','',NULL,NULL)
        SQL = "SELECT scobro.*,nomforpa FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        SQL = SQL & " WHERE scobro.codmacta = '" & vCRM.Codmacta & "'"
        SQL = SQL & " AND (sforpa.tipforpa between 2 and 5) "
        SQL = SQL & " AND impcobro<>0 ORDER BY fecvenci desc"   'ANTES era >0
        
        J = CInt(NumRegElim) 'pk puede que haya metidos de cobros. NO reseteo Numregelim
        Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        Base = 0
        Impor1 = 0
        
        While Not Rs.EOF
        'trozo copiado d ela funcion de ver cobros pdtes
          
                'NO esta recibido. Si tiene diferencia
                Impor1 = DBLet(Rs!impcobro, "N")
                NumRegElim = NumRegElim + 1
                SQL = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,"
                SQL = SQL & "`importe`,`observa`,forpa) values ( "
                SQL = SQL & vUsu.codigo & "," & NumRegElim & ",2,'"    '2.  El 2 es RIESGO para el rpt
                SQL = SQL & Rs!numSerie & Format(Rs!Codfaccl, "000000")
                If Rs!FecVenci < Now Then SQL = SQL & " *"
                SQL = SQL & "','" & Format(Rs!fecfaccl, FormatoFecha)
                SQL = SQL & "','" & Format(Rs!FecVenci, FormatoFecha) & "'," & TransformaComasPuntos(CStr(Impor1)) & ","
                'Antes la observa era NULL, ahora llevare el Depto
                If IsNull(Rs!Departamento) Then
                    Aux = "NULL"
                Else
                    Aux = "codclien = " & vCRM.codClien & " AND coddirec  "
                    Aux = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Aux, CStr(Rs!Departamento))
                    If Aux = "" Then Aux = Rs!Departamento
                    Aux = "'" & DevNombreSQL(Aux) & "'"
                    
                End If
                SQL = SQL & Aux
                'Mayo 2010
                'Con forma de pago
                SQL = SQL & ",'" & Format(Rs!codforpa, "000") & " - " & DevNombreSQL(Rs!nomforpa) & "')"
                conn.Execute SQL
                Rs.MoveNext

            
        
        Wend
        Rs.Close
        
       
        
        
        
         
        
    End If
    
    J = DevuelveIndiceNodo("adm3")
    If HayKprocesarNodo(J, F) Then
        Donde = "Hco reclamas"
        SQL = "SELECT codigo,numserie,codfaccl,fecfaccl,fecreclama,impvenci,codmacta,observaciones from shcocob "
        SQL = SQL & " WHERE codmacta = '" & vCRM.Codmacta & "'"
        SQL = SQL & " AND fecreclama >= '" & Format(F, FormatoFecha) & "' "
        'SQL = SQL & " AND (sforpa.tipforpa between 0 and 3) ORDER BY fecvenci desc"
        J = CInt(NumRegElim) 'pk puede que haya metidos de cobros. NO reseteo Numregelim
        
        Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            NumRegElim = NumRegElim + 1
            SQL = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ("
            SQL = SQL & vUsu.codigo & "," & NumRegElim & ",1,'"
            SQL = SQL & DBLet(Rs!numSerie, "T") & Format(DBLet(Rs!Codfaccl, "N"), "000000") & "','"
            SQL = SQL & Format(Rs!fecfaccl, FormatoFecha) & "','" & Format(Rs!fecreclama, FormatoFecha) & "',"
            SQL = SQL & TransformaComasPuntos(Rs!ImpVenci) & ",'"
            cad = DBLetMemo(Rs!Observaciones)
            cad = Replace(cad, vbCrLf, " ")
            SQL = SQL & DevNombreSQL(cad) & "')"
            conn.Execute SQL
            Rs.MoveNext
        Wend
        Rs.Close
        
        'Ha metido algun dato
        'If NumRegElim > J Then admon(3) = True   'tiene datos
    End If
    
    
    'Vere si teiene manteinimeots para mostrar/o no en el rpt
    J = DevuelveIndiceNodo("adm4")
    If HayKprocesarNodo(J, F) Then
        SQL = "Select count(*) from scaman where codclien = " & Text1.Text
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        If Not Rs.EOF Then NumRegElim = DBLet(Rs.Fields(0), "N")
        Rs.Close
        'If NumRegElim > 0 Then admon(4) = True
    End If
End Sub


Private Sub GenerarDatosSAT()
Dim cad As String
Dim Contador As Long
Dim F As Date

   

    Donde = "SAT"
    'Volumen facturacion
    J = DevuelveIndiceNodo("sat1")
    If HayKprocesarNodo(J, F) Then
        Donde = "Frecuencias"
        
    
    End If
    
    
    J = DevuelveIndiceNodo("sat2")
    If HayKprocesarNodo(J, F) Then
        Donde = "Albaranes reparacion"
        
    End If
    
    J = DevuelveIndiceNodo("sat3")
    If HayKprocesarNodo(J, F) Then
        Donde = "Avisos pdtes de cerrar"
        
    End If
    
    J = DevuelveIndiceNodo("sat4")
    If HayKprocesarNodo(J, F) Then
        Donde = "Equipos pendientes reparar"
        
    End If
    
End Sub



Private Sub FijarNodo3(ByRef Nod, Padre As String, Clave As String, LlevaFecha As Boolean, Anyo As Boolean, texto As String)
Dim Aux As String
Dim Fecha As Date
Dim Leido As Boolean

    'Primero AÑADO EL NODO
    Set Nod = TV1.Nodes.Add(Padre, tvwChild, Clave)
    Nod.Text = texto
    
    'Veo si estan leido los datos de preselccion
    Leido = False
    If Not DatosGuardados Is Nothing Then
        If DatosGuardados.Count > 0 Then Leido = True
    End If
        
    If Leido Then
        If Nod.Index > DatosGuardados.Count Then
            Leido = False
        End If
    End If
    
    
    If Not Leido Then
        Nod.Checked = True
        
    Else
        Nod.Checked = RecuperaValor(DatosGuardados(Nod.Index), 1) = "1"
        'Debug.Print Nod.Text & " " & Nod.Checked
    End If
    
    If LlevaFecha Then
        If Not Leido Then
            Fecha = "01/01/2010"
        Else
            Aux = RecuperaValor(DatosGuardados(Nod.Index), 2)
            If Aux = "" Then
                Aux = "01/10/2010"
            Else
                If Not IsDate(Aux) Then Aux = "01/01/2010"
            End If
            Fecha = Aux
            
        End If
        
        Aux = Nod.Text & "   ["
        If Anyo Then
            Aux = Aux & Year(Fecha)
        Else
            Aux = Aux & Format(Fecha, "dd/mm/yyyy")
        End If
        Aux = Aux & "]"
        Nod.Text = Aux
    End If
End Sub



''''''Private Sub FijarNodoConFecha(ByRef Nod, Anyo As Boolean)
''''''Dim Aux As String
''''''Dim Fecha As Date
''''''
''''''    'Leeriamos de datos guardados
''''''    If False Then
''''''
''''''    Else
''''''        Fecha = "01/01/2010"
''''''    End If
''''''
''''''
''''''
''''''
''''''    Aux = Nod.Text & "   ["
''''''    If Anyo Then
''''''        Aux = Aux & Year(Fecha)
''''''    Else
''''''        Aux = Aux & Format(Fecha, "dd/mm/yyyy")
''''''    End If
''''''    Aux = Aux & "]"
''''''    Nod.Text = Aux
''''''End Sub





'Dado un NODO
Private Function HayKprocesarNodo(Indice As Integer, ByRef Fecha As Date) As Boolean
Dim i As Integer
Dim Valor As String
Dim TieneFecha As Boolean
Dim CadenaFecha As String
Dim CadenaVisible As String
Dim Aux As String
Dim NodoOfertaPedidoAlbaran As Boolean


    Fecha = CDate("01/01/2007")
    i = InStr(1, TV1.Nodes(Indice).Text, "[")
    TieneFecha = i > 0
    
    
    If TieneFecha Then
        Valor = Mid(TV1.Nodes(Indice).Text, i + 1)
        Valor = Mid(Valor, 1, Len(Valor) - 1)
    End If
    
    'Sabremos si esta marcado o no
    HayKprocesarNodo = TV1.Nodes(Indice).Checked
    
    
    'Si es un NODO padre no leo mas, ya que no hay campos visibles para ellos
    If TV1.Nodes(Indice).Parent Is Nothing Then Exit Function
    

    NodoOfertaPedidoAlbaran = False
    If Indice = 7 Or Indice = 8 Or Indice = 9 Then NodoOfertaPedidoAlbaran = True
        
    If NodoOfertaPedidoAlbaran Then
        CadenaVisible = RecuperaValor(TV1.Nodes(Indice).Tag, 1)
        If CadenaVisible <> "" Then
            'El nodo esta marcado para imprimir
            If Not CadenaOfePedAlb(Indice, Aux) Then
                CadenaVisible = ""  'para qe no imprima

            End If
        End If
        
    Else
        CadenaVisible = RecuperaValor(TV1.Nodes(Indice).Tag, 1)
    End If  'para los nodos de ofer,ped alb y el resto
    
    
    If CadenaVisible <> "" Then
        cadParam2 = cadParam2 & CadenaVisible & "=" & Val(Abs(TV1.Nodes(Indice).Checked)) & "|"
    Else
       ' MsgBox "No hay campo visible en el rpt", vbInformation
    End If
    CadenaFecha = RecuperaValor(TV1.Nodes(Indice).Tag, 2)
    'FECHA
    'Si hay fecha
    If CadenaFecha <> "" Then
        If Len(Valor) = 4 Then
            'Es solo el año
            cadParam2 = cadParam2 & CadenaFecha & "=" & Valor
            Fecha = CDate("01/01/" & Valor)
        Else
            cadParam2 = cadParam2 & CadenaFecha & "=" & "Date(" & Year(Valor) & ", " & Month(Valor) & ", " & Day(Valor) & ")"
            Fecha = CDate(Valor)
        End If
        cadParam2 = cadParam2 & "|"
    Else
        If Valor <> "" Then MsgBox "Hay fecha y no hay campo en el rpt para indicarla", vbInformation
    End If
             
        
    
        
End Function

Private Sub Configuracion(Leer As Boolean)
    SQL = App.Path & "\crmdef.dat"
    If Leer Then
        If Dir(SQL, vbArchive) <> "" Then
            'Lo cargo todo
            If Not ProcFicheroConfig(True) Then Set DatosGuardados = Nothing
        End If
    Else
        ProcFicheroConfig False
    
    End If
End Sub



Private Function ProcFicheroConfig(Leer As Boolean) As Boolean
Dim TieneF As Boolean
Dim i As Integer
Dim Aux As String
Dim NF As Integer

    On Error GoTo eLeerFicheroConfig
    ProcFicheroConfig = False
    NF = FreeFile
    If Leer Then
        Open SQL For Input As #NF
        
        Set DatosGuardados = New Collection
        SQL = ""
        While Not EOF(NF)
            Line Input #NF, SQL
            DatosGuardados.Add SQL
        Wend
        Close #NF
        
    Else
    
        Open SQL For Output As #NF
        For J = 1 To TV1.Nodes.Count
            i = InStr(1, TV1.Nodes(J), "[")
            TieneF = i > 0
            
            SQL = Abs(TV1.Nodes(J).Checked) & "|"
            If TieneF Then
                Aux = Mid(TV1.Nodes(J).Text, i + 1)
                Aux = Mid(Aux, 1, Len(Aux) - 1)
                If Len(Aux) = 4 Then Aux = "01/01/" & Aux
                
            Else
                Aux = ""
            End If
            SQL = SQL & Aux & "|"
            Print #NF, SQL
        Next J
        Close #NF
    End If
    
    ProcFicheroConfig = True
    
    Exit Function
eLeerFicheroConfig:
    MuestraError Err.Number, "LeerFicheroConfig"
    TrataCerrarFichero NF
End Function

Private Sub TrataCerrarFichero(ByRef NFF As Integer)
    On Error Resume Next
    Close #NFF
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaDatosAux()
Dim C As Byte
    C = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    tv2.Nodes.Clear
    tv3.Nodes.Clear
    If Text1.Text <> "" Then
        Set Rs = New ADODB.Recordset
        lblInd.Caption = ""
        CargaImpresionAuxiliar
        lblInd.Caption = ""
        Set Rs = Nothing
    End If
    Screen.MousePointer = C
End Sub

Private Sub CargaImpresionAuxiliar()
Dim PpalInsertado As Boolean
Dim N

    
        
    '***********************************************************************
    'OFERTAS
    lblInd.Caption = "OFERTAS"
    lblInd.Refresh
    SQL = "Select numofert,fecofert from scapre where codclien =" & Text1.Text & " AND "
    SQL = SQL & DevFecha(7, "fecofert")
    SQL = SQL & " ORDER BY fecofert"
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PpalInsertado = False
    While Not Rs.EOF
        If Not PpalInsertado Then
            Set N = tv2.Nodes.Add(, , "OFE")
            N.Text = "OFERTAS"
            N.Bold = True
            N.Checked = True
            
            Set N = tv3.Nodes.Add(, , "OFE")
            N.Text = "OFERTAS"
            N.Bold = True
            N.Checked = True
            PpalInsertado = True
        End If
        
        SQL = Format(Rs!NumOfert, "000000") & "  -  " & Format(Rs!fecofert, "dd/mm/yyyy")
        Set N = tv2.Nodes.Add("OFE", tvwChild)
        N.Text = SQL
        N.Checked = True
        Set N = tv3.Nodes.Add("OFE", tvwChild)
        N.Text = SQL
        N.Checked = True
        Rs.MoveNext
    Wend
    Rs.Close
    If PpalInsertado Then
        tv2.Nodes(N.Index).EnsureVisible
        tv3.Nodes(N.Index).EnsureVisible
    End If
    
    
    
    '***********************************************************************
    'PEDIDO
    lblInd.Caption = "PEDIDOS"
    lblInd.Refresh
    SQL = "Select numpedcl,fecpedcl from scaped where codclien =" & Text1.Text & " AND "
    SQL = SQL & DevFecha(8, "fecpedcl")
    SQL = SQL & " ORDER BY fecpedcl"
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PpalInsertado = False
    While Not Rs.EOF
        If Not PpalInsertado Then
            Set N = tv2.Nodes.Add(, , "PED")
            N.Text = "PEDIDOS"
            N.Bold = True
            N.Checked = True
            N.ForeColor = &H4000&
            Set N = tv3.Nodes.Add(, , "PED")
            N.Text = "PEDIDOS"
            N.Bold = True
            N.Checked = True
            PpalInsertado = True
            N.ForeColor = &H4000&
        End If
        
        SQL = Format(Rs!numpedcl, "000000") & "  -  " & Format(Rs!fecpedcl, "dd/mm/yyyy")
        Set N = tv2.Nodes.Add("PED", tvwChild)
        N.Text = SQL
        N.Checked = True
        Set N = tv3.Nodes.Add("PED", tvwChild)
        N.Text = SQL
        N.Checked = True
        Rs.MoveNext
    Wend
    Rs.Close
    If PpalInsertado Then
        tv2.Nodes(N.Index).EnsureVisible
        tv3.Nodes(N.Index).EnsureVisible
    End If
    
    
    '***********************************************************************
    'ALBARANES
    lblInd.Caption = "ALBARANES"
    lblInd.Refresh
    SQL = "Select codtipom,numalbar,fechaalb from scaalb where "
    SQL = SQL & DevFecha(9, "fechaalb")
    SQL = SQL & " AND codtipom <>'ALZ' and codtipom<>'ALR' and "
    SQL = SQL & " codClien = " & Text1.Text
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PpalInsertado = False
    While Not Rs.EOF
        If Not PpalInsertado Then
            Set N = tv2.Nodes.Add(, , "ALB")
            N.Text = "ALBARANES"
            N.Bold = True
            N.Checked = True
            N.ForeColor = &H80&
            Set N = tv3.Nodes.Add(, , "ALB")
            N.Text = "ALBARANES"
            N.Bold = True
            N.Checked = True
            N.ForeColor = &H80&
            PpalInsertado = True
        End If
        
        SQL = Rs!codtipom & Format(Rs!NumAlbar, "000000") & "  -  " & Format(Rs!FechaAlb, "dd/mm/yy")
        Set N = tv2.Nodes.Add("ALB", tvwChild)
        N.Checked = True
        N.Text = SQL
        Set N = tv3.Nodes.Add("ALB", tvwChild)
        N.Text = SQL
        N.Checked = True
        
        Rs.MoveNext
    Wend
    Rs.Close
    If PpalInsertado Then
        tv2.Nodes(N.Index).EnsureVisible
        tv3.Nodes(N.Index).EnsureVisible
    End If
    
End Sub


Private Function DevFecha(Indice As Integer, CampoBD As String) As String
Dim i As Integer
Dim F As String
    F = CDate("01/01/1900")
    i = InStr(1, TV1.Nodes(Indice).Text, "[")
    If i > 0 Then F = Mid(TV1.Nodes(Indice), i + 1, 10)
    DevFecha = CampoBD & " >= '" & Format(F, FormatoFecha) & "'"
End Function

Private Sub tv2_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    If PrimeraVez Then Exit Sub
    
    'Pong el nodo en el tv3 chcec(unche
    tv3.Nodes(Node.Index).Checked = Node.Checked
    
    Dim CH As Boolean
    
    If Node.Checked Then
        If Not Node.Parent Is Nothing Then Node.Parent.Checked = True
    End If
    CH = Node.Checked
    CheckSubNodo Node, CH, True
    
    
    Err.Clear
End Sub

Private Sub tv3_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim CH As Boolean
    If PrimeraVez Then Exit Sub
    
    If Node.Checked Then
        If Not Node.Parent Is Nothing Then Node.Parent.Checked = True
    End If
    
    
    CH = Node.Checked
    CheckSubNodo Node, CH, False
    
    
End Sub


Private Function CadenaOfePedAlb(Index As Integer, CadenaSQL_ As String) As Boolean
Dim J As Integer
Dim N As Node
Dim Pad As Node
Dim C2 As String

    CadenaOfePedAlb = False
    CadenaSQL_ = "-1"
    If tv2.Nodes.Count <= 1 Then Exit Function  'si no hay modos, nos piaramos
    
    Set Pad = tv2.Nodes(1)
    
    Select Case Index
    Case 7
        'OFERTAS
        If Pad.Key <> "OFE" Then Exit Function
        Set N = Pad.Child
        CadenaSQL_ = ""
        While Not N Is Nothing
            
            If N.Checked Then
                J = InStr(1, N.Text, "-")
                If J > 0 Then CadenaSQL_ = CadenaSQL_ & ", " & Trim(Mid(N.Text, 1, J - 1))
            End If
            Set N = N.Next
       Wend
 
        
    Case 8
        J = 0
        While J = 0
            If Pad.Key = "PED" Then
                J = 1
            Else
                Set Pad = Pad.Next
                If Pad Is Nothing Then J = 1
            End If
        Wend
        
        If Pad Is Nothing Then Exit Function
        Set N = Pad.Child
        CadenaSQL_ = ""
        While Not N Is Nothing
            If N.Checked Then
                J = InStr(1, N.Text, "-")
                If J > 0 Then CadenaSQL_ = CadenaSQL_ & ", " & Trim(Mid(N.Text, 1, J - 1))
            End If
            Set N = N.Next
        Wend

       
       
       
    Case 9
         'ALBARANES
         J = 0
         While J = 0
             If Pad.Key = "ALB" Then
                 J = 1
             Else
                 Set Pad = Pad.Next
                 If Pad Is Nothing Then J = 1
             End If
         Wend
         
         If Pad Is Nothing Then Exit Function
         Set N = Pad.Child
         CadenaSQL_ = ""
         While Not N Is Nothing
             If N.Checked Then
                J = InStr(1, N.Text, "-")
                If J > 0 Then
                    C2 = Trim(Mid(N.Text, 1, J - 1))
                    CadenaSQL_ = CadenaSQL_ & ", ('" & Mid(C2, 1, 3) & "'," & Mid(C2, 4) & ")"
                End If
             End If
             Set N = N.Next
         Wend

    End Select
    
          'Ninguno seleccionado
       If InStr(1, CadenaSQL_, ",") = 0 Then
            CadenaOfePedAlb = False
            CadenaSQL_ = "-1"
       Else
            CadenaSQL_ = Mid(CadenaSQL_, 2)
            CadenaOfePedAlb = True
       
            InsertarEnTmpsOfePedAlb Index, CadenaSQL_
       
       
       
       
       
       
       
       End If
    
End Function




Private Sub InsertarEnTmpsOfePedAlb(Indice As Integer, ByRef Conjunto As String)
Dim C As String
Dim C2 As String
    Select Case Indice
    Case 7
        C = "Select * from scapre where numofert in (" & Conjunto & ") ORDER by fecofert asc"
        Rs.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            '                               secuencial   ofe/ped/alb  iden     dpto     vacio
            NumRegElim = NumRegElim + 1
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`fecha1`,`fecha2`)"
            
            'ANTES MAYO2010
'            C = C & " VALUES (" & vUsu.Codigo & "," & NumRegElim & ",1,"
'            'identificador
'            C = C & Format(RS!NumOfert, "000000") & ","
'
            'AHORA
            C = C & " VALUES (" & vUsu.codigo & "," & Rs!NumOfert & ",1,"
            'identificador
            C = C & Format(NumRegElim, "000000") & ","
                        
            If IsNull(Rs!CodDirec) Then
                C2 = "NULL"
            Else
                C2 = "'" & Rs!CodDirec & "   " & DevNombreSQL(DBLet(Rs!nomdirec, "T")) & "'"
            End If
            '               vacio de momento
            C = C & C2 & ",NULL,"
            C2 = DevuelveDesdeBD(conAri, "sum(importel)", "slipre", "numofert", Rs!NumOfert, "N")
            If C2 = "" Then C2 = "0"
            C = C & TransformaComasPuntos(C2)
            C = C & "," & DBSet(Rs!fecofert, "F") & "," & DBSet(Rs!FecEntre, "F") & ")"
            conn.Execute C
            Rs.MoveNext
        Wend
        Rs.Close
        
    Case 8
        C = "Select * from scaped where numpedcl IN (" & Conjunto & ") ORDER by fecpedcl asc"
        Rs.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            '                               secuencial   ofe/ped/alb  iden     dpto     vacio
            NumRegElim = NumRegElim + 1
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`fecha1`,`fecha2`,obser)"
            C = C & " VALUES (" & vUsu.codigo & "," & NumRegElim & ",2,"  '2 de pedido
            'identificador
            C = C & Format(Rs!numpedcl, "000000") & ","
            If IsNull(Rs!CodDirec) Then
                C2 = "NULL"
            Else
                C2 = "'" & Rs!CodDirec & "   " & DBLet(Rs!nomdirec, "T") & "'"
            End If
            '               vacio de momento
            C = C & C2 & ",NULL,"
            C2 = DevuelveDesdeBD(conAri, "sum(importel)", "sliped", "numpedcl", Rs!numpedcl, "N")
            If C2 = "" Then C2 = "0"
            C = C & TransformaComasPuntos(C2)
            C = C & "," & DBSet(Rs!fecpedcl, "F") & "," & DBSet(Rs!FecEntre, "F") & "," & DBSet(Rs!observacrm, "T", "S") & ")"
            conn.Execute C
            Rs.MoveNext
        Wend
        Rs.Close
    
    Case 9
        C = "Select * from scaalb where (codtipom,numalbar)  IN (" & Conjunto & ") ORDER by fechaalb,codtipom asc"
        Rs.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            '                               secuencial   ofe/ped/alb  iden     dpto     vacio
            NumRegElim = NumRegElim + 1
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`fecha1`,`fecha2`,obser)"
            C = C & " VALUES (" & vUsu.codigo & "," & NumRegElim & ",3,"  '3 de alb
            'identificador
            C = C & "'" & Rs!codtipom & Format(Rs!NumAlbar, "000000") & "',"
            If IsNull(Rs!CodDirec) Then
                C2 = "NULL"
            Else
                C2 = "'" & Rs!CodDirec & "   " & DBLet(Rs!nomdirec, "T") & "'"
            End If
            '               vacio de momento
            C = C & C2 & ",NULL,"
            C2 = DevuelveDesdeBD(conAri, "sum(importel)", "slialb", "codtipom = '" & Rs!codtipom & "' AND numalbar", Rs!NumAlbar, "N")
            If C2 = "" Then C2 = "0"
            C = C & TransformaComasPuntos(C2)
            C = C & "," & DBSet(Rs!FechaAlb, "F") & ",NULL" & "," & DBSet(Rs!observacrm, "T", "S") & ")"
            conn.Execute C
            Rs.MoveNext
        Wend
        Rs.Close
    
    End Select
        
End Sub

Private Sub ImprimirDocumentosAuxiliares()
Dim Cuantos As Integer
Dim N As Node

    If tv3.Nodes.Count = 0 Then Exit Sub
    
    
    Set N = tv3.Nodes(1)
    SQL = ""
    For J = 1 To tv3.Nodes.Count
        If tv3.Nodes(J).Checked Then
            If Not tv3.Nodes(J).Parent Is Nothing Then
                SQL = "OK"   'Si es nodo hijo
                Exit For
            End If
        End If
    Next
    
    If SQL = "" Then
      '  MsgBox "Ningun datos seleccionado", vbExclamation
        J = 0
    Else
        J = 1
        SQL = "Va a imprimir las ofertas/pedidos/albaranes seleccionados" & vbCrLf & vbCrLf
        SQL = SQL & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then J = 0
    End If
    If J = 0 Then Exit Sub
    
    Set N = tv3.Nodes(1)
    While Not N Is Nothing
        ImprimirReports N
        
        Set N = N.Next
    Wend
    
End Sub


'       0- Ofertas   1-Pedidos   2-Albaranes
Private Sub ImprimirReports(ByRef NodoPadre As Node)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim devuelve As String, campo As String
Dim OpcRPT As Integer
Dim numParam As Byte
Dim cadFormula As String
Dim N As Node
Dim AntiguoTipmov As String

'Dim campo1 As String, campo2 As String, campo3 As String
    
    J = 0
    Set N = NodoPadre.Child
    While Not (N Is Nothing)
        If N.Checked Then J = 1
        Set N = N.Next
    Wend
    
    If J = 0 Then Exit Sub 'No hay ninguno
  
    '===================================================
    '============ PARAMETROS ===========================
    Select Case NodoPadre.Key
        Case "PED"
            indRPT = 7 '7: Pedidos de Clientes
            OpcRPT = 38  'impreison pedidos
            
        Case "OFE"
            indRPT = 5
            OpcRPT = 31
        Case Else
            'NodoPadre .key ="ALB"
            indRPT = 10
            OpcRPT = 45
    End Select
    numParam = 0
    cadParam2 = ""
    If Not PonerParamRPT2(indRPT, cadParam2, numParam, Donde, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
     
    
    
        'Añadimos a los parametros el tipo de IVA que se aplica a ese cliente (para saber si esta exento o no de IVA)
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1.Text, "N")
        If devuelve <> "" Then
            cadParam2 = cadParam2 & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
        
        'PORTES
        cadParam2 = cadParam2 & "vPortes=""" & vParamAplic.ArtPortesN & """|"
        numParam = numParam + 1
    

    cadFormula = ""
    SQL = ""
    AntiguoTipmov = ""
    Set N = NodoPadre.Child
    While Not (N Is Nothing)
        If N.Checked Then
            
            Select Case NodoPadre.Key
            Case "PED"
                If SQL = "" Then SQL = "{scaped.codclien} = " & Text1.Text & " AND {scaped.numpedcl} IN "
                J = InStr(1, N.Text, "-")
                cadFormula = cadFormula & ", " & Trim(Mid(N.Text, 1, J - 1))
                
            Case "OFE"
                'Añado el parametro de carta NO
                If cadFormula = "" Then
                    'Es la 1era vez k entra aqui
                    cadParam2 = cadParam2 & "pCodCarta=0|"
                    numParam = numParam + 1
                    SQL = "{scapre.codclien} = " & Text1.Text & " AND {scapre.numofert} IN "
                End If
                 J = InStr(1, N.Text, "-")
                 cadFormula = cadFormula & ", " & Trim(Mid(N.Text, 1, J - 1))
                 
                 
            Case Else
                If Mid(N.Text, 1, 3) <> AntiguoTipmov Then
                    If AntiguoTipmov <> "" Then Imprime cadFormula, OpcRPT, cadParam2, numParam
                    cadFormula = ""
                    AntiguoTipmov = Mid(N.Text, 1, 3)
                End If
                'ALBARANES
                '{scaalb.codtipom}='ALV' AND ({scaalb.numalbar}=14)
                If cadFormula = "" Then
                    'Es la 1era vez k entra aqui
                    'PUNTO VERDE
                    cadParam2 = cadParam2 & "PuntoVerde=""" & vParamAplic.ArtReciclado & """|"
                    numParam = numParam + 1
                    
                    'Si se imprimen importes y/o
                    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Text1.Text, "N")
                    If devuelve = "" Then devuelve = "0"
                    ' 0 "Todo"
                    ' 1 "Cantidad y Precio"
                    ' 2 "Cantidad"
                    cadParam2 = cadParam2 & "Albarcon=" & devuelve & "|"
                    numParam = numParam + 1
                    
                    SQL = "{scaalb.codclien} = " & Text1.Text & " AND {scaalb.codtipom}= '" & AntiguoTipmov & "' AND {scaalb.numalbar} IN "
                    
                End If
                 J = InStr(1, N.Text, "-")
                 cadFormula = cadFormula & ", " & Trim(Mid(N.Text, 4, J - 4))
                
                
            End Select
            
            
        End If
        Set N = N.Next
        
    Wend
    
    Imprime cadFormula, OpcRPT, cadParam2, numParam
            
            
       
            
    
End Sub





Private Sub Imprime(cadFormula As String, OpcRPT As Integer, cadParam As String, numParam As Byte)
        cadFormula = Mid(cadFormula, 2) 'quito la primera coma
        cadFormula = "[" & cadFormula & "]"
        cadFormula = SQL & cadFormula
    
         With frmImprimir
                    
                    .outTipoDocumento = 0
            '        If DatosEnvioMail <> "" Then
            '            .outTipoDocumento = RecuperaValor(DatosEnvioMail, 1)
            '            .outCodigoCliProv = RecuperaValor(DatosEnvioMail, 2)
            '            .outClaveNombreArchiv = RecuperaValor(DatosEnvioMail, 3)
            '        End If
                    .FormulaSeleccion = cadFormula
                    .OtrosParametros = cadParam2
                    .NumeroParametros = numParam
                    .SoloImprimir = True
                    .EnvioEMail = False
                    .Opcion = OpcRPT
                    .Titulo = "Datos auxiliares desde CRM"
                    .SeleccionaRPTCodigo = pRptvMultiInforme
                    If OpcRPT = 31 Then
                        .Titulo = .Titulo & "(OFERTAS)"
                    ElseIf OpcRPT = 38 Then
                        .Titulo = .Titulo & "(PEDIDOS)"
                    Else
                        .Titulo = .Titulo & "(ALBARANES)"
                    End If
                    .NombreRPT = Donde  'tendra el nomrtp
                    'If PonerNombrePDF Then .NombrePDF = cadPDFrpt
                    .ConSubInforme = True
                    .Show vbModal
                End With
                Me.Refresh
                DoEvents
                Screen.MousePointer = vbHourglass
                Espera 0.4
End Sub
