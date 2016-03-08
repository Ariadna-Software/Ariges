VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUtilidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilidades"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmUtilidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameCLI 
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   5055
      Begin VB.ComboBox cboTipoMov 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   1150
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1550
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   480
         Width           =   1150
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Inicio"
         Height          =   195
         Index           =   3
         Left            =   200
         TabIndex        =   15
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Fin"
         Height          =   195
         Index           =   2
         Left            =   1550
         TabIndex        =   14
         Top             =   240
         Width           =   810
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   3
         Left            =   1110
         Picture         =   "frmUtilidades.frx":000C
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   2
         Left            =   2460
         Picture         =   "frmUtilidades.frx":0097
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Mov."
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frameBus2 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   5055
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4035
      End
   End
   Begin VB.CommandButton cmdBus 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Can"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   5350
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   4980
      Width           =   4515
   End
   Begin VB.Label Label1 
      Caption         =   "Cuentas sin movimientos"
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
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6615
   End
End
Attribute VB_Name = "frmUtilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////////
'/*
'/*         Este formulario es para algunos puntos de utilidades.
'/*         Esta a parte pq vamos a poner el boton de parar busqueda
'/*         ya que es una simple busqueda



Public Opcion As Byte
    '0.-
    '1.-
    '2.-
    '3.-
    '4.-
    
    '5.- Facturas Clientes
    '6.- Facturas pendientes de contabilizar (Clientes)
    '7.- Facturas pendientes de contabilizar (Proveedores)
    
    
'Private WithEvents frmC As frmColCtas
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
    
Private Estado As Byte
    '0.- Antes de empezar a buscar
    '1.- Buscando
    '2.- Han parado la busqueda
    '3.- Ha terminado la busqueda y hay datos
Dim SQL As String
Dim RS As ADODB.Recordset
Dim NumCuentas As Long
Dim i As Long
Dim ItmX As ListItem
Dim HanPulsadoCancelar As Boolean
Dim PrimeraVez As Boolean

Dim FechaIni As String, FechaFin As String




Private Sub cboTipomov_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub cmdBus_Click()
    Select Case Estado
    Case 0
        ListView1.ListItems.Clear
        Select Case Opcion
            
        Case 5 'Buscar saltos en numero facturas
            Screen.MousePointer = vbHourglass
            BuscarFacturas
            Screen.MousePointer = vbDefault
            
        Case 6, 7
            Screen.MousePointer = vbHourglass
            BuscarFacturasPtesConta
            Screen.MousePointer = vbDefault

        End Select
    
    Case 2
        'Volvemos donde nos habiamos quedado
        PonerCampos 1
        HanPulsadoCancelar = False

    Case 3
        ListView1.ListItems.Clear
        PonerCampos 0
        
    End Select
End Sub


Private Sub cmdCancel_Click()
    Select Case Estado
    Case 0
        Unload Me
        
    Case 1
        HanPulsadoCancelar = True
            
    Case 2
        'Volvemos a poner una nueva busqueda
        IntentaCErrar
        PonerCampos 0
    Case 3
        Unload Me
    End Select
End Sub


Private Sub IntentaCErrar()
On Error Resume Next
    RS.Close
    Err.Clear
    Set RS = Nothing
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Not BloqueoManual("Busquedas", "1") Then
            MsgBox "Se esta realizando la busqueda desde otro PC", vbExclamation
            PrimeraVez = True
            Unload Me
        End If
        
        If Opcion >= 5 Or Opcion <= 7 Then
            PonerFoco Text1(3)
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    
    PrimeraVez = True
    Me.ListView1.Icons = frmPpal.ImageListB
    Me.ListView1.SmallIcons = frmPpal.ImageListB
    CargaEncabezado
    PonerCampos 0

    frameBus2.visible = False

    Me.frameCLI.visible = False
    Select Case Opcion

    Case 5     'Saltos en Facturas clientes
        Me.frameCLI.visible = True
        Label5.visible = Opcion = 5
        Me.cboTipoMov.visible = Opcion = 5
        CargaComboTipoMov
'        FechasEjercicioConta FechaIni, FechaFin
        Text1(2).Text = Format(Now, "dd/mm/yyyy")
        Text1(3).Text = Format(Now - 365, "dd/mm/yyyy")
        
        Label1.Caption = "Nº facturas CLIENTE incorrectos"
        
    Case 6, 7 '6: Facturas pendientes de contabilizar (Clientes)
              '7: Facturas pendientes de contabilizar (Proveedores)
        Me.frameCLI.visible = True
        Label5.visible = Opcion = 5
        Me.cboTipoMov.visible = False
        Text1(2).Text = Format(Now, "dd/mm/yyyy")
        Text1(3).Text = Format(Now - 365, "dd/mm/yyyy")
        
        If Opcion = 6 Then
            Label1.Caption = "Facturas CLIENTE pendientes contabilizar"
        Else
            Label1.Caption = "Facturas PROVEEDOR pendientes contabilizar"
        End If
        
    
    End Select
    
    'No puede eliminar cuentas
'    Command3.Enabled = vUsu.Nivel < 2
'    Me.cmdEliminarAgrup.Enabled = vUsu.Nivel < 2
'    Me.cmdNuevaAgrup.Enabled = Me.cmdEliminarAgrup.Enabled
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos(NuevoEstado As Byte)
    Select Case NuevoEstado
    Case 0
        Me.Label2.Caption = ""
        Me.pb1.visible = False
        Me.cmdCancel.Caption = "&Salir"
        Me.cmdBus.Caption = "&Iniciar"
    Case 1
        Me.cmdCancel.Caption = "&Parar"
        
    Case 2
        Me.cmdCancel.Caption = "&Cancelar"
        Me.cmdBus.Caption = "&Reanudar"
    Case 3
        Me.cmdBus.Caption = "&Reestablecer"
        Me.cmdCancel.Caption = "&Salir"
    End Select

    Me.cmdBus.Enabled = (NuevoEstado <> 1)
    cmdBus.visible = (Opcion < 2) Or Opcion >= 4 'Cuando es agrupacion no mostramos el inciar
    Estado = NuevoEstado
End Sub


Private Sub CargaEncabezado()
Dim clmX As ColumnHeader

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    
    Select Case Opcion

    Case 5 'Saltos en numeros de Facturas CLIENTES
        Me.ListView1.Checkboxes = False
        '* Facturas
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "T.Mov."
        clmX.Width = 800
        i = 3400
        SQL = "Codigo"
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = SQL
        clmX.Width = 1300
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Año"
        clmX.Width = 800
        'Clave2 ...
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Comentario"
        clmX.Width = i
        
    Case 6, 7 '6: Facturas ptes de contabilizar (Clientes)
              '7: Facturas ptes de contabilizar (Proveedores)
         Me.ListView1.Checkboxes = False
         
        If Opcion = 6 Then
            Set clmX = ListView1.ColumnHeaders.Add()
            clmX.Text = "T.Mov."
            clmX.Width = 800
        End If
'        i = 3400
'        SQL = "Nº Factura"
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Nº Factura"
        clmX.Width = 1200
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Fecha"
        clmX.Width = 1200
        
        If Opcion = 7 Then
            Set clmX = ListView1.ColumnHeaders.Add()
            clmX.Text = "Cod.Prov."
            clmX.Width = 1000
        End If
        
        Set clmX = ListView1.ColumnHeaders.Add()
        If Opcion = 6 Then
            SQL = "Cliente"
        Else
            SQL = "Proveedor"
        End If
        clmX.Text = SQL
        clmX.Width = 3000
           
         
    End Select
End Sub


Private Sub MontarBusqueda()
'    SQL = "DELETE FROM tmpbussinmov"
'    Conn.Execute SQL
'    SQL = "INSERT INTO tmpbussinmov SELECT codmacta,nommacta from cuentas where apudirec='S'"
'    Conn.Execute SQL
End Sub





Private Sub Form_Unload(Cancel As Integer)
    If Not PrimeraVez Then
        DesBloqueoManual "Busquedas"
'        BloqueoManual False, "Busquedas", ""
        IntentaCErrar
    End If
    
End Sub



Private Sub frmF_Selec(vFecha As Date)
    Text1(i).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imfech_Click(Index As Integer)
    i = Index
    Set frmF = New frmCal
    SQL = Now
    If Text1(i).Text <> "" Then
        If IsDate(Text1(i).Text) Then SQL = Text1(i).Text
    End If
    frmF.Fecha = CDate(SQL)
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 4
'    Text1(Index).SelStart = 0
'    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'   If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{tab}"
'    End If
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    PonerFormatoFecha Text1(Index)
End Sub



Private Sub BuscarFacturas()
Dim cad As String
Dim Aux As String
Dim Anyo As Integer
Dim J As Integer

    On Error GoTo EBuscarFacturas
        
    
    SQL = "codtipom,year(fecfactu) as anyo, fecfactu,numfactu as codigo"
    cad = "fecfactu"
    
    
    SQL = SQL & " FROM scafac"
'    If Opcion = 6 Then SQL = SQL & "prov"
    
    'Si hay fecha inicio
    Aux = CadenaDesdeHastaBD(Text1(3).Text, Text1(2).Text, cad, "F")
'    Aux = ""
'    If Text1(3).Text <> "" Then Aux = cad & " >= " & DBSet(Text1(3).Text, "F")
'    If Text1(2).Text <> "" Then
'        If Aux <> "" Then Aux = Aux & " AND "
'        Aux = Aux & cad & " <= " & DBSet(Text1(2).Text, "F")
'    End If
    
    If Me.cboTipoMov.List(Me.cboTipoMov.ListIndex) <> "" Then
'    If txtCLI.Text <> "" Then
        If Aux <> "" Then Aux = Aux & " AND "
        Aux = Aux & " codtipom = '" & Mid(Me.cboTipoMov.List(Me.cboTipoMov.ListIndex), 1, 3) & "'"
    End If
    
    
    If Aux <> "" Then SQL = SQL & " WHERE " & Aux
    SQL = SQL & " ORDER BY "
    SQL = SQL & "codtipom,year(fecfactu) , numfactu "
    
    
       
    'Vale. Ya tenemos montado el SQL
    Set RS = New ADODB.Recordset
    RS.Open "SELECT " & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Serie
    Aux = ""
    Anyo = 0
    While Not RS.EOF
'        If Opcion = 5 Then
        If RS!codTipoM <> Aux Then
            'Nueva SERIE
            Aux = RS!codTipoM
            Anyo = RS!Anyo
            i = RS!Codigo
        End If
'        End If

        If Anyo <> RS!Anyo Then
            'AÑO DISTINTO
            i = RS!Codigo
            Anyo = RS!Anyo
        End If
        
        'Para cada numero de factura
        If i = RS!Codigo Then
            i = i + 1
            'no hacemos nada mas
        Else
            'Si si que es mayor. Hay salto o hueco

            If RS!Codigo - i > 2 Then
                'SALTO
                cad = Format(RS!Codigo - 1, "0000000")
               '--- Para clientes
                Set ItmX = ListView1.ListItems.Add(, , RS!codTipoM)
                ItmX.SubItems(1) = cad
                J = 2
                '----
                ItmX.SubItems(J) = Anyo
                ItmX.SubItems(J + 1) = "Salto desde codigo: " & Format(i, "0000000")
                
                
            Else
                'HUECO
                cad = Format(i, "0000000")
                '--- Para clientes
                Set ItmX = ListView1.ListItems.Add(, , RS!codTipoM)
                ItmX.SubItems(1) = cad
                J = 2
                '---
                ItmX.SubItems(J) = Anyo
                'i = RS!Codigo + 1
            End If
            ItmX.SmallIcon = 9
            i = RS!Codigo + 1
        End If
        'Movemos siguiente
        RS.MoveNext
        
    Wend
    RS.Close
    Set RS = Nothing
    If ListView1.ListItems.Count = 0 Then MsgBox "Proceso finalizado", vbInformation
    
    Exit Sub
EBuscarFacturas:
    MuestraError Err.Number, Err.Description
End Sub



Private Sub CargaComboTipoMov()
'Dim RS As ADODB.Recordset
'Dim SQL As String
Dim i As Byte
    
    On Error Resume Next
    
    Me.cboTipoMov.Clear
    
    SQL = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom LIKE 'F%'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = RS!nomtipom
        SQL = Replace(SQL, "Factura", "Fact.")
        cboTipoMov.AddItem RS!codTipoM & " - " & SQL
        cboTipoMov.ItemData(cboTipoMov.NewIndex) = i
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'Situamos el combo en las facturas de venta que son mas comunes
    For i = 0 To Me.cboTipoMov.ListCount - 1
        If InStr(1, Me.cboTipoMov.List(i), "FAV") Then
            Me.cboTipoMov.ListIndex = i
            Exit For
        End If
    Next i
    
    If Err.Number = 0 Then Err.Clear
End Sub



Private Sub BuscarFacturasPtesConta()
'Buscar facturas pendientes de contabilizar para clientes o Proveedores
Dim J As Integer
Dim cad As String, Aux As String

    On Error GoTo EBuscarConta
    
    
    'Si hay fecha inicio
    cad = "fecfactu"
    Aux = ""
    Aux = CadenaDesdeHastaBD(Text1(3).Text, Text1(2).Text, cad, "F")
'    Aux = ""
'    If Text1(3).Text <> "" Then Aux = cad & " >= " & DBSet(Text1(3).Text, "F")
'    If Text1(2).Text <> "" Then
'        If Aux <> "" Then Aux = Aux & " AND "
'        Aux = Aux & cad & " <= " & DBSet(Text1(2).Text, "F")
'    End If
    
    
    If Opcion = 6 Then 'clientes
        SQL = "select codtipom,numfactu,fecfactu,codclien,nomclien from scafac "
    Else 'proveedores
        SQL = "select codprove,numfactu,fecfactu,nomprove from scafpc "
    End If
    If Aux <> "" Then SQL = SQL & " WHERE " & Aux
    SQL = SQL & " and intconta = 0 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        If Opcion = 6 Then
            '--- Para clientes
            Set ItmX = ListView1.ListItems.Add(, , RS!codTipoM)
            ItmX.SubItems(1) = Format(RS!NumFactu, "0000000")
            J = 2
            '----
        Else
            Set ItmX = ListView1.ListItems.Add(, , Format(RS!NumFactu, "0000000"))
            J = 1
        End If
        
        ItmX.SubItems(J) = RS!FecFactu
'        ItmX.SubItems(J + 1) = "Salto desde codigo: " & Format(i, "0000000")
        If Opcion = 6 Then
            ItmX.SubItems(J + 1) = RS!nomclien
        Else
            ItmX.SubItems(J + 1) = Format(RS!codProve, "000000")
            ItmX.SubItems(J + 2) = RS!nomprove
        End If
        ItmX.SmallIcon = 9
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    If ListView1.ListItems.Count = 0 Then MsgBox "Proceso finalizado", vbInformation
    
EBuscarConta:
    If Err.Number <> 0 Then MuestraError Err.Number, "Buscar Facturas pendientes contabilizar." & Err.Description
End Sub
