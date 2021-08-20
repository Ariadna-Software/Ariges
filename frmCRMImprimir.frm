VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
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
Private WithEvents frmC2 As frmBasico2
Attribute frmC2.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private GuardarConfig As Boolean

Dim J As Integer
Dim RS As ADODB.Recordset
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
    Set N = tv1.Nodes.Add(, , "ADM")
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
    Set N = tv1.Nodes.Add(, , "COM")
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

   
     FijarNodo3 N, "COM", "com6", True, False, "Acciones comerciales "
     N.Tag = "pVisAccionesComer|pDesdeAccComer|"

    
    FijarNodo3 N, "COM", "com7", True, False, "Historial"
    N.Tag = "pHistorial|pDesdHistorial|"
    
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
    Set N = tv1.Nodes.Add(, , "SAT")
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
    If tv1.Nodes.Count = 0 Then Exit Sub
    If tv1.SelectedItem Is Nothing Then Exit Sub
    
    If Right(tv1.SelectedItem.Text, 1) <> "]" Then
        MsgBox "NO se le asigna fecha a esta opcion", vbExclamation
    Else
        SQL = ""
        J = InStr(1, tv1.SelectedItem, "[")
        If J = 0 Then
            MsgBox "No se ha encotrado la marca de fecha", vbExclamation
        Else
            Donde = Mid(tv1.SelectedItem.Text, J + 1)
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
                
                J = InStr(tv1.SelectedItem.Text, "[")
                If J = 0 Then
                    MsgBox ""
            
                Else
                    'Ha retornado dato
                    GuardarConfig = True
                    Donde = Mid(tv1.SelectedItem.Text, 1, J)
                    Donde = Donde & SQL & "]"
                    tv1.SelectedItem.Text = Donde
                    
                    
                    'Si es de ofertas o de albaranes CARGAMOS los datos de nuevo
                    SQL = tv1.SelectedItem.Key
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
    ejecutar "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo, False
    NumRegElim = 0 'contador para tmp con los ofe/ped/alb
    
    
    Set RS = New ADODB.Recordset
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
    Set RS = Nothing
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
    For J = 1 To tv1.Nodes.Count
        'TV1.Nodes(J).Checked = True
        tv1.Nodes(J).EnsureVisible
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
    Set frmC2 = New frmBasico2
    AyudaClientes frmC2, Text1.Text
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
            tv2.Nodes(J).Checked = Index = 3
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
    If tv1.Nodes.Count = 0 Then Exit Sub
    If tv1.SelectedItem Is Nothing Then Exit Sub
    
    If InStr(tv1.SelectedItem.Text, "[") = 0 Then Exit Sub
    
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
Dim NO
    
    Set NO = N
    NO.Checked = Checkar
    If EsElTV2 Then CheckeaTambienEnElTv3 NO.Index, Checkar
    Set NO = N.Child
    While Not NO Is Nothing
        CheckSubNodo NO, Checkar, EsElTV2
        Set NO = NO.Next
    Wend
    
    
    
End Sub


Private Sub CheckeaTambienEnElTv3(Indice As Integer, chk)
    On Error Resume Next
    tv3.Nodes(Indice).Checked = chk
    Err.Clear
End Sub


Private Sub LlamarImprimir2(SoloCobros As Boolean)
Dim K As Integer

    With frmImprimir
        .FormulaSeleccion = "{tmpcrmclien.codusu} = " & vUsu.Codigo
        
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
    SQL = SQL & vUsu.Codigo & "," & Text1.Text & ","
    
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
    If Me.tv1.Nodes(J).Checked Then
        GenerarDatosAdmon
    Else
        'PONGO TODOS LOS SUBPARAMETROS A FALSE
        PonerparametrosVisiblesFalse
    End If
    
    
    'Para saber si tiene datos cada secccion
    J = DevuelveIndiceNodo("COM")
    If Me.tv1.Nodes(J).Checked Then
        GenerarDatosComer
    Else
        'PONGO TODOS LOS SUBPARAMETROS A FALSE
        PonerparametrosVisiblesFalse
    End If


    'Para saber si tiene datos cada secccion
    If vParamAplic.Reparaciones Then
        J = DevuelveIndiceNodo("SAT")
        If Me.tv1.Nodes(J).Checked Then
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


    
    Set N = tv1.Nodes(1).Child 'ADMON
    If Leer Then
        
        CadenaDesdeOtroForm = CInt(tv1.Nodes(1).Checked) & "|"
        tv1.Nodes(1).Checked = True
        
        While Not (N Is Nothing)
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & CInt(N.Checked) & "|"
            N.Checked = True
            Set N = N.Next
        Wend

    Else
        'Tengo en cadenasql lo que habia leido antes de marcarlo
        NumRegElim = 1
        SQL = RecuperaValor(CadenaDesdeOtroForm, CInt(NumRegElim))
        tv1.Nodes(1).Checked = SQL = "-1"
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
    Set N = tv1.Nodes(J).Child '
    While Not (N Is Nothing)
        SQL = RecuperaValor(N.Tag, 1)
        If SQL <> "" Then cadParam2 = cadParam2 & SQL & "=0|"
        Set N = N.Next
    Wend
End Sub



Private Function DevuelveIndiceNodo(Clave As String) As Integer
Dim i As Integer
    
    For i = 1 To tv1.Nodes.Count
        If tv1.Nodes(i).Key = Clave Then
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
Dim Cad As String
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
    
        J = DevuelveIndiceNodo("com6")
        If HayKprocesarNodo(J, F) Then
            Donde = "Acciones comerciales"
        End If
        


    J = DevuelveIndiceNodo("com7")
        If HayKprocesarNodo(J, F) Then
            Donde = "Historial"
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
            
            
            
            
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                NumRegElim = NumRegElim + 1
                SQL = "insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,"
                SQL = SQL & "`trabajador`,`adjuntos`) values ( " & vUsu.Codigo & "," & NumRegElim & ",0,"
                SQL = SQL & DBSet(RS!feholla, "FH") & ","
                'En sllama siempre son RECIBIDAS
                SQL = SQL & "'Recibida',"
                Cad = DBLetMemo(RS!observac)
                Cad = Replace(Cad, vbCrLf, " ")
                SQL = SQL & DBSet(Cad, "T", "S") & ","
                'Trabajador
                SQL = SQL & DBSet(RS!NomTraba, "T") & ","
                'En adjuntos guardare el tipop llamada
                SQL = SQL & DBSet(RS!nomllama1, "T") & ")"
                
                conn.Execute SQL
                RS.MoveNext
            Wend
            RS.Close
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
            
            
            
            
            
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                NumRegElim = NumRegElim + 1
                SQL = "insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,"
                SQL = SQL & "`trabajador`,`adjuntos`) values ( " & vUsu.Codigo & "," & NumRegElim & ",0,"
                SQL = SQL & DBSet(RS!fechora, "FH") & ","
                'En sllama siempre son RECIBIDAS
                SQL = SQL & "'Realizada',"
                Cad = DBLetMemo(RS!Observaciones)
                Cad = Replace(Cad, vbCrLf, " ")
                SQL = SQL & DBSet(Cad, "T", "S") & ","
                'Trabajador
                SQL = SQL & DBSet(RS!NomTraba, "T") & ","
                'En adjuntos guardare el tipop llamada
                SQL = SQL & "NULL)"
                
                conn.Execute SQL
                RS.MoveNext
            Wend
            RS.Close
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
            If tv1.Nodes(J).Checked Then NumRegElim = 1
            
            J = DevuelveIndiceNodo("com51")
            If tv1.Nodes(J).Checked Then NumRegElim = NumRegElim + 2
            
            If NumRegElim > 0 Then
                    'Ha selecionado alguno de los dos, o los dos
                    
                    'insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,`trabajador`,`adjuntos`) values ( '1','0','','',NULL,NULL,NULL,NULL)
                    SQL = "select fechahora,enviado,email,asunto,adjuntos from scrmmail"
                    SQL = SQL & " WHERE codclien=" & vCRM.codClien
                     SQL = SQL & " AND fechahora>=" & DBSet(F, "F")
                    If NumRegElim = 1 Or NumRegElim = 2 Then
                        Cad = "1"
                        If NumRegElim = 2 Then Cad = "0"
                        'Ha selecionado solo una de las dos
                        SQL = SQL & " AND enviado = " & Cad
                    End If
                    NumRegElim = Contador
                    
                
                
                
                    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not RS.EOF
                        NumRegElim = NumRegElim + 1
                        SQL = "insert into `tmpcrmmsg` (`codusu`,`codigo`,`tipo`,`fechahora`,`rec_env`,`asun_obs`,"
                        SQL = SQL & "`trabajador`,`adjuntos`) values ( " & vUsu.Codigo & "," & NumRegElim & ",1,"  '1.email
                        SQL = SQL & DBSet(RS!FechaHora, "FH") & ","
                        'En sllama siempre son RECIBIDAS
                        If Val(RS!Enviado) = 1 Then
                            SQL = SQL & "'Enviado',"
                        Else
                            SQL = SQL & "'Recibido',"
                        End If
                        Cad = DBLetMemo(RS!asunto)
                        Cad = Replace(Cad, vbCrLf, " ")
                        SQL = SQL & DBSet(Cad, "T", "S") & ","
                        'Trabajador
                        SQL = SQL & DBSet(RS!email, "T", "S") & ","
                        'En adjuntos guardare el tipop llamada
                        Cad = "'*'"
                        If DBLet(RS!adjuntos, "T") = "" Then Cad = "NULL"
                        SQL = SQL & Cad & ")"
                        
                        conn.Execute SQL
                        RS.MoveNext
                    Wend
                    RS.Close
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
Dim Cad As String
Dim Aux As String
Dim F As Date
Dim DiasRiesgo As Long
Dim N As Integer

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
        
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        
        While Not RS.EOF
            Cad = ""
        
            NumRegElim = NumRegElim + 1
            Impor1 = DBLet(RS!TotalFac, "N")
            
            SQL = "insert into `tmpcrmtesor` (`codusu`,`codigo`,`importe`,`anyotxt`,`variacion`)"
            SQL = SQL & " values (" & vUsu.Codigo & "," & NumRegElim & "," & TransformaComasPuntos(CStr(Impor1)) & ",'"
            
            If Val(RS!Anyo) = Year(Now) Then
                'Valor actual.
                SQL = SQL & "actual',"
                'Cambio la base para comprar con el mismo periodo del actual
                
                'Cad = "codtipom <>'FAZ' and codtipom<>'FRT' and "
                Cad = "codtipom <>'FAZ' and "
                Cad = Cad & " fecfactu>='" & Year(Now) - 1 & "-01-01' and "
                Cad = Cad & " fecfactu<='" & Year(Now) - 1 & "-" & Format(Now, "mm-dd") & "' AND codclien "
                Cad = DevuelveDesdeBD(conAri, "sum(totalfac)", "scafac", Cad, Text1.Text)
                If Cad = "" Then Cad = "0"
                Base = CCur(Cad)
                If NumRegElim > 1 And Base <> 0 Then
                    Impor1 = CStr(((100 * Impor1) / Base) - 100)
                    Cad = Format(Impor1, FormatoPorcen) & "% sobre misma fecha año anterior"
                Else
                    Cad = ""
                End If
            Else
                'Otro año cualquiera
                 SQL = SQL & RS!Anyo & "',"
                If NumRegElim > 1 And Base <> 0 Then
                    Impor1 = CStr(((100 * Impor1) / Base) - 100)
                    Cad = Format(Impor1, FormatoPorcen) & "%"
                End If
                 
            End If
            Base = DBLet(RS!TotalFac, "N")
            SQL = SQL & "'" & Cad & "')"
          

            conn.Execute SQL
            RS.MoveNext
        Wend
        RS.Close
        'If NumRegElim > 0 Then admon(1) = True
    
    
    End If
    
    
    J = DevuelveIndiceNodo("adm2")
    If HayKprocesarNodo(J, F) Then
        Donde = "Cobros pendientes"
        'insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ( '1','0','0','','','',NULL,NULL)
        If vParamAplic.ContabilidadNueva Then
            SQL = " cobros as scobro INNER JOIN formapago as sforpa ON scobro.codforpa=sforpa.codforpa "
        Else
            SQL = " scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        End If
         
        SQL = "SELECT scobro.*,nomforpa FROM " & SQL
        SQL = SQL & " WHERE scobro.codmacta = '" & vCRM.Codmacta & "'"
        
        SQL = SQL & "  AND recedocu=0 ORDER BY fecvenci desc"
        
        NumRegElim = 0
        RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        Base = 0
        Impor1 = 0
        
        While Not RS.EOF
              'trozo copiado d ela funcion de ver cobros pdtes
          If DBLet(RS!Devuelto, "N") = 1 Then
                'SALE SEGURO (si no esta girado otra vez ¿no?
                'Si esta girado otra vez tendra impcobro, con lo cual NO tendra diferencia de importes
                Impor1 = RS!ImpVenci + DBLet(RS!gastos, "N") - DBLet(RS!impcobro, "N")
                
            Else
                'Si esta recibido NO lo saco
                If Val(RS!recedocu) = 1 Then
                    Impor1 = 0
                Else
                    'NO esta recibido. Si tiene diferencia
                    Impor1 = RS!ImpVenci + DBLet(RS!gastos, "N") - DBLet(RS!impcobro, "N")
            
                End If
          End If
          If Impor1 <> 0 Then
                NumRegElim = NumRegElim + 1
                SQL = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,"
                SQL = SQL & "`importe`,`observa`,forpa) values ( "
                SQL = SQL & vUsu.Codigo & "," & NumRegElim & ",0,'"
                
                SQL = SQL & RS!numSerie
                If vParamAplic.ContabilidadNueva Then
                    SQL = SQL & Format(RS!Numfactu, "000000")
                Else
                    SQL = SQL & Format(RS!Codfaccl, "000000")
                End If
                If RS!FecVenci < Now Then SQL = SQL & " *"
                
                If vParamAplic.ContabilidadNueva Then
                    SQL = SQL & "','" & Format(RS!FecFactu, FormatoFecha)
                Else
                    SQL = SQL & "','" & Format(RS!fecfaccl, FormatoFecha)
                End If
                
                SQL = SQL & "','" & Format(RS!FecVenci, FormatoFecha) & "'," & TransformaComasPuntos(CStr(Impor1)) & ","
                'Antes la observa era NULL, ahora llevare el Depto
                If IsNull(RS!Departamento) Then
                    Aux = "NULL"
                Else
                    Aux = "codclien = " & vCRM.codClien & " AND coddirec  "
                    Aux = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Aux, CStr(RS!Departamento))
                    If Aux = "" Then Aux = RS!Departamento
                    Aux = "'" & DevNombreSQL(Aux) & "'"
                    
                End If
                SQL = SQL & Aux
                'Mayo 2010
                'Con forma de pago
                SQL = SQL & ",'" & Format(RS!codforpa, "000") & " - " & DevNombreSQL(RS!nomforpa) & "')"
                conn.Execute SQL
          End If
          RS.MoveNext

            
        
        Wend
        RS.Close
        
        
        'Marzo 2011
        'Tambien sacare el riesgo. Habra que configurar el rpt de cada uno
        '----------------------------------------------------------------
        Donde = "Riesgo tesoreria"
        'insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ( '1','0','0','','','',NULL,NULL)
        
        
        
        
        If vParamAplic.ContabilidadNueva Then
            SQL = "SELECT impvenci,impcobro,numfactu codfaccl,numserie,fecvenci,fecfactu fecfaccl,Departamento,scobro.codforpa,nomforpa,  talondias ,pagaredias ,remesadiasmenor,gastos "
            SQL = SQL & " ,fecultco,sforpa.tipforpa,codrem,tiporem,recedocu FROM cobros as scobro INNER JOIN formapago as sforpa ON scobro.codforpa=sforpa.codforpa "
            SQL = SQL & " LEFT JOIN bancos ON scobro.ctabanc1=bancos.codmacta "
        Else
            SQL = "SELECT impcobro,Codfaccl,numserie,fecvenci,fecfaccl,Departamento,scobro.codforpa,nomforpa,  talondias , pagaredias ,remesadiasmayor"
            SQL = SQL & ",fecultco, sforpa.tipforpa,recedocu FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
            SQL = SQL & " LEFT JOIN ctabancaria bancos ON scobro.ctabanc1=bancos.codmacta "
        End If
        
        
        SQL = SQL & " WHERE scobro.codmacta = '" & vCRM.Codmacta & "'"
        
        
        
        If Not vParamAplic.ContabilidadNueva Then
            SQL = SQL & " AND (sforpa.tipforpa between 2 and 5) "
            SQL = SQL & " AND fecultco> " & DBSet(DateAdd("yyyy", -1, Now), "F") 'Aqueloos que la remesa(normal-talpag) sea de una año hasta ahora
            SQL = SQL & " AND impcobro<>0 "
        Else
            
            
        End If
        
        SQL = SQL & " ORDER BY fecvenci desc"
        J = CInt(NumRegElim) 'pk puede que haya metidos de cobros. NO reseteo Numregelim
        RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        Base = 0
        Impor1 = 0
        
        While Not RS.EOF
        
                
                
                If Not vParamAplic.ContabilidadNueva Then
                    DiasRiesgo = 1  'lo que hacia. No tocamos nada
                        Impor1 = RS!impcobro
                Else
                    
                    If DBLet(RS!codrem, "N") = 0 Then
                        
                        If DBLet(RS!impcobro, "N") = 0 Then
                            DiasRiesgo = 0
                        Else
                            If DBLet(RS!recedocu, "N") = 1 Then
                                'Nuevo Febrero 2020
                                Impor1 = RS!ImpVenci + DBLet(RS!gastos, "N") 'QUitamos el cobrado para indicar que esta en recepcion documento
                            Else
                                'Lo que habia
                                Impor1 = RS!ImpVenci + DBLet(RS!gastos, "N") - DBLet(RS!impcobro, "N")
                            End If
                            DiasRiesgo = IIf(Impor1 <> 0, 1, 0)
                        End If
                    Else
                        N = 0
                        If RS!tiporem = 1 Then
                            N = DBLet(RS!remesadiasmenor, "N")
                        Else
                            If RS!tiporem = 2 Then
                                N = DBLet(RS!pagaredias, "N")
                            Else
                                N = DBLet(RS!talondias, "N")
                            End If
                        End If
                      
                        DiasRiesgo = 0 'no inserta
                        If N > 0 Then
                            F = DateAdd("d", N, RS!FecVenci)
                            If F > Now Then
                                Impor1 = RS!ImpVenci + DBLet(RS!gastos, "N")
                                DiasRiesgo = 1
                        
                            End If
                        End If
                
                    End If
                
                
                End If
                
                If DiasRiesgo > 0 Then
                    Impor1 = DBLet(Impor1, "N")
                    NumRegElim = NumRegElim + 1
                    SQL = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,"
                    SQL = SQL & "`importe`,`observa`,forpa) values ( "
                    SQL = SQL & vUsu.Codigo & "," & NumRegElim & ",2,'"    '2.  El 2 es RIESGO para el rpt
                    SQL = SQL & RS!numSerie & Format(RS!Codfaccl, "000000")
                    If RS!FecVenci < Now Then SQL = SQL & " *"
                    SQL = SQL & "','" & Format(RS!fecfaccl, FormatoFecha)
                    SQL = SQL & "','" & Format(RS!FecVenci, FormatoFecha) & "'," & TransformaComasPuntos(CStr(Impor1)) & ","
                    'Antes la observa era NULL, ahora llevare el Depto
                    If IsNull(RS!Departamento) Then
                        Aux = "NULL"
                    Else
                        Aux = "codclien = " & vCRM.codClien & " AND coddirec  "
                        Aux = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Aux, CStr(RS!Departamento))
                        If Aux = "" Then Aux = RS!Departamento
                        Aux = "'" & DevNombreSQL(Aux) & "'"
                        
                    End If
                    SQL = SQL & Aux
                    'Mayo 2010
                    'Con forma de pago
                    SQL = SQL & ",'" & Format(RS!codforpa, "000") & " - " & DevNombreSQL(RS!nomforpa) & "')"
                    conn.Execute SQL
                End If
                RS.MoveNext

                
        
        Wend
        RS.Close
        Impor1 = 0
       
        
        
        
         
        
    End If
    
    
    
    
    J = DevuelveIndiceNodo("adm3")
    If HayKprocesarNodo(J, F) Then
        Donde = "Hco reclamas"
        
        If vParamAplic.ContabilidadNueva Then
            SQL = "select reclama.codigo,numserie,numfactu codfaccl,fecfactu fecfaccl,fecreclama,impvenci,codmacta,observaciones,importes "
            SQL = SQL & " from reclama  INNER JOIN reclama_facturas  ON reclama.codigo=reclama_facturas.codigo"
        
        Else
            SQL = "SELECT codigo,numserie,codfaccl,fecfaccl,fecreclama,impvenci,codmacta,observaciones from shcocob "
        
        End If
        
        SQL = SQL & " WHERE codmacta = '" & vCRM.Codmacta & "'"
        SQL = SQL & " AND fecreclama >= '" & Format(F, FormatoFecha) & "' ORDER BY fecreclama"
        If vParamAplic.ContabilidadNueva Then SQL = SQL & ",reclama_facturas.codigo"
        
        
        
        J = CInt(NumRegElim) 'pk puede que haya metidos de cobros. NO reseteo Numregelim
        
        RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not RS.EOF
            NumRegElim = NumRegElim + 1
            SQL = "insert into `tmpcrmcobros` (`codusu`,`secuencial`,`tipo`,`numfac`,`fecfaccl`,`fecha2`,`importe`,`observa`) values ("
            SQL = SQL & vUsu.Codigo & "," & NumRegElim & ",1,'"
            SQL = SQL & DBLet(RS!numSerie, "T") & Format(DBLet(RS!Codfaccl, "N"), "000000") & "','"
            SQL = SQL & Format(RS!fecfaccl, FormatoFecha) & "','" & Format(RS!fecreclama, FormatoFecha) & "',"
            If vParamAplic.ContabilidadNueva Then
                If IsNull(RS!ImpVenci) Then
                    Aux = DBLet(RS!Importes)
                Else
                    Aux = RS!ImpVenci
                End If
            Else
                Aux = RS!ImpVenci
            End If
            SQL = SQL & TransformaComasPuntos(Aux) & ",'"
            Cad = DBLetMemo(RS!Observaciones)
            Cad = Replace(Cad, vbCrLf, " ")
            SQL = SQL & DevNombreSQL(Cad) & "')"
            conn.Execute SQL
            
            
            
            
            RS.MoveNext
        Wend
        RS.Close
        
        'Ha metido algun dato
        'If NumRegElim > J Then admon(3) = True   'tiene datos
    End If
    
    
    'Vere si teiene manteinimeots para mostrar/o no en el rpt
    J = DevuelveIndiceNodo("adm4")
    If HayKprocesarNodo(J, F) Then
        SQL = "Select count(*) from scaman where codclien = " & Text1.Text
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        If Not RS.EOF Then NumRegElim = DBLet(RS.Fields(0), "N")
        RS.Close
        'If NumRegElim > 0 Then admon(4) = True
    End If
End Sub


Private Sub GenerarDatosSAT()
Dim Cad As String
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
Dim leido As Boolean

    'Primero AÑADO EL NODO
    Set Nod = tv1.Nodes.Add(Padre, tvwChild, Clave)
    Nod.Text = texto
    
    'Veo si estan leido los datos de preselccion
    leido = False
    If Not DatosGuardados Is Nothing Then
        If DatosGuardados.Count > 0 Then leido = True
    End If
        
    If leido Then
        If Nod.Index > DatosGuardados.Count Then
            leido = False
        End If
    End If
    
    
    If Not leido Then
        Nod.Checked = True
        
    Else
        Nod.Checked = RecuperaValor(DatosGuardados(Nod.Index), 1) = "1"
        'Debug.Print Nod.Text & " " & Nod.Checked
    End If
    
    If LlevaFecha Then
        If Not leido Then
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
    i = InStr(1, tv1.Nodes(Indice).Text, "[")
    TieneFecha = i > 0
    
    
    If TieneFecha Then
        Valor = Mid(tv1.Nodes(Indice).Text, i + 1)
        Valor = Mid(Valor, 1, Len(Valor) - 1)
    End If
    
    'Sabremos si esta marcado o no
    HayKprocesarNodo = tv1.Nodes(Indice).Checked
    
    
    'Si es un NODO padre no leo mas, ya que no hay campos visibles para ellos
    If tv1.Nodes(Indice).Parent Is Nothing Then Exit Function
    

    NodoOfertaPedidoAlbaran = False
    If Indice = 7 Or Indice = 8 Or Indice = 9 Then NodoOfertaPedidoAlbaran = True
        
    If NodoOfertaPedidoAlbaran Then
        CadenaVisible = RecuperaValor(tv1.Nodes(Indice).Tag, 1)
        If CadenaVisible <> "" Then
            'El nodo esta marcado para imprimir
            If Not CadenaOfePedAlb(Indice, Aux) Then
                CadenaVisible = ""  'para qe no imprima

            End If
        End If
        
    Else
        CadenaVisible = RecuperaValor(tv1.Nodes(Indice).Tag, 1)
    End If  'para los nodos de ofer,ped alb y el resto
    
    
    If CadenaVisible <> "" Then
        cadParam2 = cadParam2 & CadenaVisible & "=" & Val(Abs(tv1.Nodes(Indice).Checked)) & "|"
    Else
       ' MsgBox "No hay campo visible en el rpt", vbInformation
    End If
    CadenaFecha = RecuperaValor(tv1.Nodes(Indice).Tag, 2)
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
        For J = 1 To tv1.Nodes.Count
            i = InStr(1, tv1.Nodes(J), "[")
            TieneF = i > 0
            
            SQL = Abs(tv1.Nodes(J).Checked) & "|"
            If TieneF Then
                Aux = Mid(tv1.Nodes(J).Text, i + 1)
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
        Set RS = New ADODB.Recordset
        lblInd.Caption = ""
        CargaImpresionAuxiliar
        lblInd.Caption = ""
        Set RS = Nothing
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
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PpalInsertado = False
    While Not RS.EOF
        If Not PpalInsertado Then
            Set N = tv2.Nodes.Add(, , "OFE")
            N.Text = "OFERTAS"
            N.Bold = True
            N.Checked = False
            
            Set N = tv3.Nodes.Add(, , "OFE")
            N.Text = "OFERTAS"
            N.Bold = True
            N.Checked = False
            PpalInsertado = True
        End If
        
        SQL = Format(RS!NumOfert, "000000") & "  -  " & Format(RS!fecofert, "dd/mm/yyyy")
        Set N = tv2.Nodes.Add("OFE", tvwChild)
        N.Text = SQL
        N.Checked = False 'True
        Set N = tv3.Nodes.Add("OFE", tvwChild)
        N.Text = SQL
        N.Checked = False
        RS.MoveNext
    Wend
    RS.Close
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
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PpalInsertado = False
    While Not RS.EOF
        If Not PpalInsertado Then
            Set N = tv2.Nodes.Add(, , "PED")
            N.Text = "PEDIDOS"
            N.Bold = True
            N.Checked = True
            N.ForeColor = &H4000&
            Set N = tv3.Nodes.Add(, , "PED")
            N.Text = "PEDIDOS"
            N.Bold = True
            N.Checked = False
            PpalInsertado = True
            N.ForeColor = &H4000&
        End If
        
        SQL = Format(RS!NumPedcl, "000000") & "  -  " & Format(RS!fecpedcl, "dd/mm/yyyy")
        Set N = tv2.Nodes.Add("PED", tvwChild)
        N.Text = SQL
        N.Checked = True
        Set N = tv3.Nodes.Add("PED", tvwChild)
        N.Text = SQL
        N.Checked = False
        RS.MoveNext
    Wend
    RS.Close
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
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PpalInsertado = False
    While Not RS.EOF
        If Not PpalInsertado Then
            Set N = tv2.Nodes.Add(, , "ALB")
            N.Text = "ALBARANES"
            N.Bold = True
            N.Checked = True
            N.ForeColor = &H80&
            Set N = tv3.Nodes.Add(, , "ALB")
            N.Text = "ALBARANES"
            N.Bold = True
            N.Checked = False
            N.ForeColor = &H80&
            PpalInsertado = True
        End If
        
        SQL = RS!codtipom & Format(RS!Numalbar, "000000") & "  -  " & Format(RS!FechaAlb, "dd/mm/yy")
        Set N = tv2.Nodes.Add("ALB", tvwChild)
        N.Checked = True
        N.Text = SQL
        Set N = tv3.Nodes.Add("ALB", tvwChild)
        N.Text = SQL
        N.Checked = False
        
        RS.MoveNext
    Wend
    RS.Close
    If PpalInsertado Then
        tv2.Nodes(N.Index).EnsureVisible
        tv3.Nodes(N.Index).EnsureVisible
    End If
    
End Sub


Private Function DevFecha(Indice As Integer, CampoBD As String) As String
Dim i As Integer
Dim F As String
    F = CDate("01/01/1900")
    i = InStr(1, tv1.Nodes(Indice).Text, "[")
    If i > 0 Then F = Mid(tv1.Nodes(Indice), i + 1, 10)
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
        RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            '                               secuencial   ofe/ped/alb  iden     dpto     vacio
            NumRegElim = NumRegElim + 1
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`fecha1`,`fecha2`)"
            
            'ANTES MAYO2010
'            C = C & " VALUES (" & vUsu.Codigo & "," & NumRegElim & ",1,"
'            'identificador
'            C = C & Format(RS!NumOfert, "000000") & ","
'
            'AHORA
            C = C & " VALUES (" & vUsu.Codigo & "," & RS!NumOfert & ",1,"
            'identificador
            C = C & Format(NumRegElim, "000000") & ","
                        
            If IsNull(RS!CodDirec) Then
                C2 = "NULL"
            Else
                C2 = "'" & RS!CodDirec & "   " & DevNombreSQL(DBLet(RS!nomdirec, "T")) & "'"
            End If
            '               vacio de momento
            C = C & C2 & ",NULL,"
            C2 = DevuelveDesdeBD(conAri, "sum(importel)", "slipre", "numofert", RS!NumOfert, "N")
            If C2 = "" Then C2 = "0"
            C = C & TransformaComasPuntos(C2)
            C = C & "," & DBSet(RS!fecofert, "F") & "," & DBSet(RS!FecEntre, "F") & ")"
            conn.Execute C
            RS.MoveNext
        Wend
        RS.Close
        
    Case 8
        C = "Select * from scaped where numpedcl IN (" & Conjunto & ") ORDER by fecpedcl asc"
        RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            '                               secuencial   ofe/ped/alb  iden     dpto     vacio
            NumRegElim = NumRegElim + 1
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`fecha1`,`fecha2`,obser)"
            C = C & " VALUES (" & vUsu.Codigo & "," & NumRegElim & ",2,"  '2 de pedido
            'identificador
            C = C & Format(RS!NumPedcl, "000000") & ","
            If IsNull(RS!CodDirec) Then
                C2 = "NULL"
            Else
                C2 = "'" & RS!CodDirec & "   " & DBLet(RS!nomdirec, "T") & "'"
            End If
            '               vacio de momento
            C = C & C2 & ",NULL,"
            C2 = DevuelveDesdeBD(conAri, "sum(importel)", "sliped", "numpedcl", RS!NumPedcl, "N")
            If C2 = "" Then C2 = "0"
            C = C & TransformaComasPuntos(C2)
            C = C & "," & DBSet(RS!fecpedcl, "F") & "," & DBSet(RS!FecEntre, "F") & "," & DBSet(RS!observacrm, "T", "S") & ")"
            conn.Execute C
            RS.MoveNext
        Wend
        RS.Close
    
    Case 9
        C = "Select * from scaalb where (codtipom,numalbar)  IN (" & Conjunto & ") ORDER by fechaalb,codtipom asc"
        RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            '                               secuencial   ofe/ped/alb  iden     dpto     vacio
            NumRegElim = NumRegElim + 1
            C = "insert into `tmpinformes` (`codusu`,`codigo1`,`campo1`,`nombre1`,`nombre2`,`nombre3`,`importe1`,`fecha1`,`fecha2`,obser)"
            C = C & " VALUES (" & vUsu.Codigo & "," & NumRegElim & ",3,"  '3 de alb
            'identificador
            C = C & "'" & RS!codtipom & Format(RS!Numalbar, "000000") & "',"
            If IsNull(RS!CodDirec) Then
                C2 = "NULL"
            Else
                C2 = "'" & RS!CodDirec & "   " & DBLet(RS!nomdirec, "T") & "'"
            End If
            '               vacio de momento
            C = C & C2 & ",NULL,"
            C2 = DevuelveDesdeBD(conAri, "sum(importel)", "slialb", "codtipom = '" & RS!codtipom & "' AND numalbar", RS!Numalbar, "N")
            If C2 = "" Then C2 = "0"
            C = C & TransformaComasPuntos(C2)
            C = C & "," & DBSet(RS!FechaAlb, "F") & ",NULL" & "," & DBSet(RS!observacrm, "T", "S") & ")"
            conn.Execute C
            RS.MoveNext
        Wend
        RS.Close
    
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
