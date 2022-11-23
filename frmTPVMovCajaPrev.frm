VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacTPVMovCajaPrev 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos Caja"
   ClientHeight    =   8640
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   15315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameMovimientosCaja 
      Height          =   8610
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15330
      Begin VB.Frame Frame3 
         Height          =   915
         Left            =   11070
         TabIndex        =   8
         Top             =   180
         Width           =   4110
         Begin VB.TextBox Text1 
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
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "Text3"
            Top             =   450
            Width           =   1680
         End
         Begin VB.TextBox Text1 
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
            Left            =   315
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "Text3"
            Top             =   450
            Width           =   1680
         End
         Begin VB.Label Label1 
            Caption         =   "Ingresos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   14
            Left            =   2070
            TabIndex        =   12
            Top             =   180
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Gastos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   315
            TabIndex        =   11
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Height          =   525
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   7740
         Width           =   2175
         Begin VB.Label lblIndicador 
            Alignment       =   2  'Center
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   6
            Top             =   180
            Width           =   1755
         End
      End
      Begin VB.Frame FrameBotonGnral 
         Height          =   810
         Left            =   225
         TabIndex        =   3
         Top             =   225
         Width           =   2850
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   240
            TabIndex        =   4
            Top             =   180
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
                  Object.Tag             =   "2"
                  Object.Width           =   1e-4
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Ver todos"
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6450
         Left            =   225
         TabIndex        =   1
         Top             =   1260
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   11377
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Cargando datos...."
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
         Left            =   2625
         TabIndex        =   7
         Top             =   7920
         Visible         =   0   'False
         Width           =   6390
      End
      Begin VB.Label Label1 
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   1
         Left            =   3465
         TabIndex        =   2
         Top             =   405
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmFacTPVMovCajaPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public NumTermi As Long

Dim FechaPrimeraCaja As String 'primera fecha de caja
Dim PrimeraVez As Boolean
Dim Todos As Boolean


Private Sub cmdRegresar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub

Private Sub PonFrameVisible(ByRef Fr As Frame, ByRef Wi As Integer, ByRef He As Integer)
    Fr.visible = True
    Wi = Fr.Width
    He = Fr.Height + 300
End Sub

Private Sub Form_Load()
Dim W As Integer, H As Integer
    
    PrimeraVez = True
    Me.Icon = frmPpal.Icon

    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM2
        .DisabledImageList = frmPpal.imgListComun_BN2
        .ImageList = frmPpal.ImgListComun2
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 2
    End With
    FechaPrimeraCaja = ""
    Todos = False
    CargarMovimientosCajas
    Me.lblIndicador.Caption = PonerContRegistrosLw(ListView1, ListView1.SelectedItem)


    Label1(1).Caption = DevuelveDesdeBD(conAri, "destermi", "spatpvt", "numtermi", CStr(NumTermi))
    Label1(1).Caption = "Terminal: " & NumTermi & "   " & Label1(1).Caption
    
    
End Sub





Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblIndicador.Caption = PonerContRegistrosLw(ListView1, Item)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim LineaAnt As Long
Dim CajaCerrada As Boolean


    CajaCerrada = False
    If Button.Index = 2 Or Button.Index = 3 Then
        If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
        
        If CDate(ListView1.SelectedItem.SubItems(1)) < CDate(FechaPrimeraCaja) Then CajaCerrada = True
            
    End If


    Select Case Button.Index
        Case 1 'insertar
            frmFacTPVMovCaja.linea = 0
            frmFacTPVMovCaja.Show vbModal
            CargarMovimientosCajas
            Me.lblIndicador.Caption = PonerContRegistrosLw(ListView1, ListView1.SelectedItem)
            
        Case 2 'modificar
            LineaAnt = ListView1.SelectedItem.Text
            frmFacTPVMovCaja.SoloVer = CajaCerrada
            frmFacTPVMovCaja.linea = LineaAnt
            frmFacTPVMovCaja.Show vbModal
            
            CargarMovimientosCajas
            
            SituarListview CStr(LineaAnt)
            Me.lblIndicador.Caption = PonerContRegistrosLw(ListView1, ListView1.SelectedItem)
            
        Case 3 'eliminar
            LineaAnt = ListView1.SelectedItem.Index
            BotonEliminar
            CargarMovimientosCajas
            If ListView1.ListItems.Count < LineaAnt Then LineaAnt = LineaAnt - 1
            
            If LineaAnt > 1 Then SituarListview2 CStr(LineaAnt)
            
            Me.lblIndicador.Caption = PonerContRegistrosLw(ListView1, ListView1.SelectedItem)
            
        Case 5
            Todos = True
            CargarMovimientosCajas
            
        Case Else
    
    End Select
    
End Sub

Private Sub SituarListview(vValor As String)
Dim i As Long
    
    
    On Error Resume Next
    
    If vValor = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If CLng(ComprobarCero(ListView1.ListItems(i).Text)) = CLng(vValor) Then
            ListView1.ListItems(i).Selected = True
            Set ListView1.SelectedItem = ListView1.ListItems(i)
            Exit For
        End If
    Next i
    If Not ListView1.SelectedItem Is Nothing Then ListView1.SelectedItem.EnsureVisible
    PonerFocoLw ListView1

End Sub

Private Sub SituarListview2(vPosicion As String)
Dim i As Long
    
    
    On Error Resume Next
    
    If vPosicion = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If CLng(ComprobarCero(ListView1.ListItems(i).Index)) = CLng(vPosicion) Then
            ListView1.ListItems(i).Selected = True
            Set ListView1.SelectedItem = ListView1.ListItems(i)
            Exit For
        End If
    Next i
    If Not ListView1.SelectedItem Is Nothing Then ListView1.SelectedItem.EnsureVisible
    PonerFocoLw ListView1

End Sub



Private Sub BotonEliminar()
Dim Sql As String
Dim Cad As String
    On Error GoTo EEliminar

    If ListView1.SelectedItem Is Nothing Then Exit Sub

    
    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar el Movimiento?"
    Cad = Cad & vbCrLf & "Fecha: " & Format(ListView1.SelectedItem.SubItems(1), "dd/mm/yyyy hh:mm:ss")
    ' **************************************************************************
    
    'borrem
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = ListView1.SelectedItem.Text
        
        Sql = "delete from stpventradassalidas where numtermi = " & NumTermi & " and idlin = " & ListView1.SelectedItem.Text
        conn.Execute Sql
    End If
    
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Movimiento", Err.Description
End Sub


Private Sub CargarMovimientosCajas()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String

Dim Sql As String
Dim ItmX As ListItem

Dim TGastos As Currency
Dim TIngresos As Currency
    

    On Error GoTo ECarga
    
    If ListView1.ColumnHeaders.Count <= 2 Then
        ListView1.ColumnHeaders.Clear
        
        ListView1.ColumnHeaders.Add , , "Id", 0
        ListView1.ColumnHeaders.Add , , "Fecha", 2400
        ListView1.ColumnHeaders.Add , , "Concepto", 2000
        ListView1.ColumnHeaders.Add , , "Descripción", 3300
        ListView1.ColumnHeaders.Add , , "Código", 900, 0
        ListView1.ColumnHeaders.Add , , "Trabajador", 2600, 0
        ListView1.ColumnHeaders.Add , , "Gastos", 1700, 1
        ListView1.ColumnHeaders.Add , , "Ingresos", 1700, 1
    End If
    
    
    Set miRsAux = New ADODB.Recordset
    
    If FechaPrimeraCaja = "" Then
    
        Sql = "select fecha from stpvdiacaja where numtermi = " & DBSet(NumTermi, "N") & " and diacierre is null "
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            FechaPrimeraCaja = DBLet(miRsAux.Fields(0).Value, "F")  'Solo es para saber que hay registros que mostrar
        Else
            MsgBox "No existe registro de apertura de caja. Llame a Ariadna.", vbExclamation
            Exit Sub
        End If
        miRsAux.Close
    
    End If
    
    
    Cad = "select idLin,diahora,Entrada,importe,descrip, cc.observa, aa.codtraba, bb.nomtraba "
    Cad = Cad & " from (stpventradassalidas aa inner join straba bb on aa.codtraba = bb.codtraba) "
    Cad = Cad & " inner join stpvtipogastos cc on aa.concepto = cc.id "
    Cad = Cad & " where numtermi = " & NumTermi
    If Not Todos Then Cad = Cad & " and date(diahora) >= " & DBSet(FechaPrimeraCaja, "F")
    Cad = Cad & " order by idlin "
    
    TGastos = 0
    TIngresos = 0
    
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView1.ListItems.Clear
    
    While Not miRsAux.EOF
        Set ItmX = ListView1.ListItems.Add
        
        ItmX.Text = Format(miRsAux!idlin, "########0")
        ItmX.SubItems(1) = Format(miRsAux!diahora, "dd/mm/yyyy hh:mm:ss")
        ItmX.SubItems(2) = DBLet(miRsAux!observa, "T")
        ItmX.SubItems(3) = DBLet(miRsAux!Descrip, "T")
        ItmX.SubItems(4) = miRsAux!CodTraba
        ItmX.SubItems(5) = miRsAux!NomTraba
        
        If DBLet(miRsAux!Entrada, "N") = 0 Then
            TGastos = TGastos + DBLet(miRsAux!Importe, "N")
            ItmX.SubItems(6) = Format(miRsAux!Importe, "###,###,##0.00")
            ItmX.SubItems(7) = " "
        Else
            TIngresos = TIngresos + DBLet(miRsAux!Importe, "N")
            ItmX.SubItems(6) = " "
            ItmX.SubItems(7) = Format(miRsAux!Importe, "###,###,##0.00")
        End If
        
        miRsAux.MoveNext
    Wend
    
    ' cargamos los totales
    Text1(0).Text = Format(TGastos, "###,###,##0.00")
    Text1(1).Text = Format(TIngresos, "###,###,##0.00")
    
    miRsAux.Close
    Set miRsAux = Nothing
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
    Set miRsAux = Nothing
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'Private Sub PonerModo(Kmodo As Byte)
'Dim B As Boolean
'Dim NumReg As Byte
'
'    Modo = Kmodo
'    PonerIndicador lblIndicador, Modo
'
'    '--------------------------------------------------
'    'Modo 2. Hay datos y estamos visualizandolos
'    B = (Kmodo = 2)
'    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = B
'    Else
'        cmdRegresar.visible = False
'    End If
'    If cmdRegresar.visible Then
'        cmdRegresar.Cancel = True
'    Else
'        Me.cmdCancelar.Cancel = True
'    End If
'
'    '-----------------------------------------------------
'    'Modo insertar o modificar
'    B = (Kmodo >= 3) Or Modo = 1 '-->Luego not b sera kmodo<3
'    cmdAceptar.visible = B
'    cmdCancelar.visible = B
'    If cmdCancelar.visible Then
'        cmdCancelar.Cancel = True
'    Else
'       ' cmdCancelar.Cancel = False
'    End If
'
''    PonerModoOpcionesMenu 'Activar opciones de menu según modo
''    PonerOpcionesMenu   'Activar opciones de menu según nivel
''                        'de permisos del usuario
'End Sub
'
'
