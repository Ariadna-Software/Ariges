VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEulerReloj 
   Caption         =   "Reloj"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12960
   Icon            =   "frmEulerReloj.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAlbaran 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frmEulerReloj.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Generar albaran"
      Top             =   7440
      Width           =   615
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Imprimir listado"
      Top             =   7320
      Width           =   615
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3855
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   7215
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3120
         Width           =   6735
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   6855
      End
      Begin VB.ComboBox cboTipoTrabajo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "frmEulerReloj.frx":0894
         Left            =   0
         List            =   "frmEulerReloj.frx":0896
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Tarea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   2640
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Albar�n / Orden producci�n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   6975
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   120
         Width           =   6735
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   7335
      Left            =   7560
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   12938
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4110
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   1234
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Numero"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tarea"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Incio"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fin"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cliente/PRod"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Id Registro"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   56.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   6240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Leyendo ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   6240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmEulerReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1


Dim Cad As String
Dim T1 As Date

Dim HayQueCerrarNodo As Integer  'Marcara el IT que tiene que actualizar como HORAFIn, 0=Inicio



Private Sub cmdAlbaran_Click(Index As Integer)
    Cad = ""
    If Combo1.ListIndex < 0 Then Cad = "Seleccione el trabajador "
        
    If Me.cboTipoTrabajo.ListIndex <= 0 Then Cad = "Seleccione el tipo de albaran"
    If Me.cboTipoTrabajo.ListIndex = 4 Then Cad = "No puede seleccionar produccion"
        
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If

    
    
    
    Cad = "  " & UCase(Me.cboTipoTrabajo.Text)
        
    Cad = "Desea crear el albar�n del tipo " & Cad & "?"
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    CrearAlbaran
End Sub

Private Sub cmdImprimir_Click()
    'VERSION RELOJ
    'frmListado2.opcion = 46
    'frmListado2.Show vbModal
End Sub

Private Sub Combo1_Click()

     If Me.Command1.Tag = 1 Then Exit Sub

    BuscarNodo
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If Me.Command1.Tag = 1 Then Exit Sub
    
    If Combo1.ListIndex < 0 Then Exit Sub

    
    BuscarNodo
End Sub

Private Sub BuscarNodo()
Dim K As Integer

    HayQueCerrarNodo = 0
    For K = ListView2.ListItems.Count To 1 Step -1
        ListView2.ListItems(K).ForeColor = vbBlack
        ListView2.ListItems(K).Bold = False
        If Combo1.ListIndex >= 0 Then
            If Val(ListView2.ListItems(K).Text) = Val(Combo1.ItemData(Combo1.ListIndex)) Then
                
                ListView2.ListItems(K).Bold = True
                If Trim(ListView2.ListItems(K).SubItems(6)) = "" Then
                    HayQueCerrarNodo = K
                    ListView2.ListItems(K).ForeColor = vbBlue
                End If
            End If
        End If
    Next
        
    PonerFrames2
    
End Sub


Private Sub cboTipoTrabajo_Click()
Dim OrdenProduccion As Boolean
    OrdenProduccion = False
    
    
    If cboTipoTrabajo.ListIndex = 0 Then
        Combo3.Clear
        Combo4.Clear
        Exit Sub
    End If
    
    Cad = "ALR"
    If cboTipoTrabajo.ListIndex = 2 Then
        Cad = "ALE"
    ElseIf cboTipoTrabajo.ListIndex = 3 Then
        Cad = "ALO"
        
        
    ElseIf cboTipoTrabajo.ListIndex = 4 Then
        'Orden produccion
        OrdenProduccion = True
    End If
    
    If OrdenProduccion Then
    
        
            Cad = Format(DateAdd("yyyy", -1, Now), FormatoFecha)
            CargarCombo_Tabla Me.Combo3, "sordprod", "concat(right(concat(""000000"",codigo),6),' - ',coalesce(descripcion,feccreacion))", "codigo", "feccreacion>='" & Cad & "'"
    
    Else
    
            Cad = "codtipom = '" & Cad & "' AND  (origdat is null or origdat<>2)"
            CargarCombo_Tabla Me.Combo3, "scaalb", "concat(numalbar,' - ',nomclien)", "NumAlbar", Cad
                
    End If
            
    Cad = "R"
    If cboTipoTrabajo.ListIndex = 2 Then
        Cad = "E"
    ElseIf cboTipoTrabajo.ListIndex = 3 Then
        Cad = "O"
    ElseIf cboTipoTrabajo.ListIndex = 4 Then
        Cad = "L" 'FALTA######
    End If
    Cad = "codtipor like '" & Cad & "_'"
    CargarCombo_Tabla Me.Combo4, "stipor", "concat(nomtipor,' [',codtipor,']')", "1", Cad

    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()

Dim linea As Integer
Dim Horas As Currency
Dim C As String
    'Por si las moscas
    
    On Error GoTo EC
    
    If Combo1.ListIndex < 0 Then Exit Sub
    If cboTipoTrabajo.ListIndex < 0 Then Exit Sub

    
    
    If cboTipoTrabajo.ListIndex < 0 Then Exit Sub
 
         
    If cboTipoTrabajo.ListIndex = 0 Then
      'Hay que cerrar nodo.
        If HayQueCerrarNodo = 0 Then
            MsgBox "Ninguna tarea iniciada para el trabajador", vbExclamation
            Exit Sub
        End If
        
       
    
    Else
        'No ha asiganado tarea
        If Combo4.ListIndex < 0 Then Exit Sub
        
    End If          'de cerrar tarea
        
           
    'Hay que cerrar nodo
    If HayQueCerrarNodo > 0 Then
        If HayQueCerrarNodo > ListView2.ListItems.Count Then
            MsgBox "Error situando datos trabajador", vbExclamation
            Exit Sub
        End If
        
        'Horas
        Horas = DateDiff("n", CDate(ListView2.ListItems(HayQueCerrarNodo).Tag), CDate(Label1(1).Caption))
    
        If Horas < 0 Then
            MsgBox "Calculos negativos!!!", vbExclamation
            Horas = 0
        End If
        'En horas, llevamos los minutos. Ahora lo pasamos a horas en decimal
        linea = Horas Mod 60  'los minutos que exende de la hora
        Horas = Horas \ 60
        Horas = Horas + Round((linea / 60), 2)
        
       ' Fecha codtraba HoraInicio
        Cad = "UPDATE sreloj SET HoraFin =" & DBSet(Label1(1).Caption, "H")
        Cad = Cad & " ,calculadas=" & DBSet(Horas, "N")
        Cad = Cad & " WHERE codtraba = " & Combo1.ItemData(Combo1.ListIndex) & " AND fecha = " & DBSet(Label1(0).Caption, "F")
        Cad = Cad & " AND HoraInicio = " & DBSet(ListView2.ListItems(HayQueCerrarNodo).Tag, "H")
                        
        conn.Execute Cad
    End If
    
    'Insertamios la nueva, si es que hay que insertarla
    If Me.cboTipoTrabajo.ListIndex > 0 Then
        C = Combo4.Text
        NumRegElim = InStr(1, C, "[")
        If NumRegElim = 0 Then Err.Raise 513, , "MAl"
            
        C = Mid(C, NumRegElim + 1)
        C = Mid(C, 1, Len(C) - 1)
        
        Cad = ",'" & C & "')"
       
        C = Combo3.Text
        NumRegElim = InStr(1, C, " - ")
        If NumRegElim = 0 Then Err.Raise 513, , "MAl 2"
        C = Trim(Mid(C, 1, NumRegElim))
        
        Cad = "," & C & Cad 'Concatenamos
        C = "'ALR'"
        If cboTipoTrabajo.ListIndex = 2 Then
            C = "'ALE'"
        ElseIf cboTipoTrabajo.ListIndex = 3 Then
            C = "'ALO'"
        ElseIf cboTipoTrabajo.ListIndex = 4 Then
            C = "null"
        End If
        Cad = "," & C & Cad
        
        Cad = DBSet(Label1(0).Caption, "F") & "," & Combo1.ItemData(Combo1.ListIndex) & "," & DBSet(Label1(1).Caption, "H") & ",null,0" & Cad
        C = DevuelveDesdeBD(conAri, "max(id)", "sreloj", "1", "1")
        C = Str(Val(C) + 1)
        Cad = C & "," & Cad
        Cad = "INSERT INTO sreloj(ID,Fecha,codtraba,HoraInicio,HoraFin,Calculadas,codtipom,numalbar,codtipor) VALUES (" & Cad
       
        conn.Execute Cad
            
    End If
    
    
    Combo3.Clear
    Combo4.Clear
    Combo1.ListIndex = -1
    cboTipoTrabajo.ListIndex = -1
    CagarMarcajes
    
    
EC:
    If Err.Number <> 0 Then MuestraError Err.Number
End Sub

Private Sub Form_Activate()
    
    If Me.Command1.Tag = 1 Then
    
        Me.Command1.Tag = 0
        
        
        
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open "Select now()", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        T1 = miRsAux.Fields(0)
        miRsAux.Close
        Labels
        
        
        miRsAux.Open "Select distinct fecha from sreloj where horafin is null", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        HayQueCerrarNodo = 0
        While Not miRsAux.EOF
            HayQueCerrarNodo = HayQueCerrarNodo + 1
            Cad = Cad & Format(miRsAux!Fecha, "dd/mm/yyyy") & "     "
            If (HayQueCerrarNodo Mod 3) = 0 Then Cad = Cad & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If Cad <> "" Then MsgBox "Dias por finalizar tareas: " & vbCrLf & Cad, vbExclamation
            
        Labels
        
        
        
        'Cargo listitiem
        HayQueCerrarNodo = 0
        CagarMarcajes

        
        
        PonerFrames2
       
        
        Timer1.Enabled = True
        
        
        
    End If
End Sub




Private Sub Form_Load()
    'Me.Icon = frmPpal.Icon   formulario compartido de proyecto
    
    'Cargar Trabajadores
    Cad = "fechabaj is null "
    CargarCombo_Tabla Me.Combo1, "straba", "codtraba", "nomtraba", Cad
    
    'Carga combo trabajos
    'ALR|ALE|ALO|
    cboTipoTrabajo.AddItem "** Salida empresa **"
    cboTipoTrabajo.AddItem "Reparaci�n"
    cboTipoTrabajo.AddItem "T. exterior"
    cboTipoTrabajo.AddItem "Orden de trabajo"
    '
    'cboTipoTrabajo.AddItem "Producci�n"   'orden de produccion

    
    Me.Command1.Tag = 1
    
    
    
    
    CargaImangenBtn
End Sub


Private Sub CargaImangenBtn()

'Version normal

    Me.cmdImprimir.Picture = frmPpal.imgListComun.ListImages(16).Picture
    Me.cmdAlbaran(0).Picture = frmPpal.ImgListPpal.ListImages(10).Picture
    ' Me.cmdAlbaran(1).Picture = frmPpal.ImgListPpal.ListImages(7).Picture


'Version SOLORELOJ. Comentar en NORMAL
'    Me.cmdImprimir.visible = False

End Sub

Private Sub Form_Resize()
Dim H As Long
    If WindowState = 1 Then Exit Sub ' ha pulsado minimizar
    
    H = Me.Width - (ListView2.Left + 600)
    If H < 0 Then H = 6375
    ListView2.Width = H
    
    'H = ListView2.Width - 6500
    'If H < 0 Then H = 1440
    'ListView2.ColumnHeaders(6).Width = H
    
    
    H = Me.Height - 120 - 1200
    If H < 0 Then H = 6375
    ListView2.Height = H
    
    
    
    
    
    Command1.Top = Me.Height - 1840
    cmdAceptar.Top = Command1.Top
    cmdImprimir.Top = Command1.Top
    Me.cmdAlbaran(0).Top = Command1.Top
'    Me.cmdAlbaran(1).Top = Command1.Top
'
'    If Me.Width > 9540 Then
'        'Command1.Left = Me.Width - 1440
'        Command1.Left = cmdAceptar.Left + cmdAceptar.Width + 1200
        
    
End Sub

Private Sub frmAc_DatoSeleccionado(CadenaSeleccion As String)
    Cad = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Cad = CadenaDevuelta
End Sub

Private Sub ListView2_DblClick()
Dim I As Integer

    If ListView2.SelectedItem Is Nothing Then Exit Sub
    
    
    For I = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(I) = Val(ListView2.SelectedItem.Text) Then
            'Este es el trabajador
            Combo1.ListIndex = I
            Exit For
        End If
    Next
    
       
    Select Case UCase(Trim(ListView2.SelectedItem.SubItems(2)))
    Case "ALE"
        I = 2
    Case "ALO"
        I = 3
    Case Else
        I = 1
    End Select
    cboTipoTrabajo.ListIndex = I
    DoEvents
    Espera 0.1
    
    For I = 0 To Combo3.ListCount - 1
        Cad = Combo3.List(I)
        Cad = Mid(Cad, 1, InStr(1, Cad, "-") - 1)
        If Val(Cad) = Val(ListView2.SelectedItem.SubItems(3)) Then
            'Este es el trabajador
            Combo3.ListIndex = I
            Exit For
        End If
    Next
    
    
 
    ListView2.Tag = Trim(Mid(ListView2.SelectedItem.SubItems(4), 1, InStr(2, ListView2.SelectedItem.SubItems(4), " ")))
    For I = 0 To Combo4.ListCount - 1
        Cad = Trim(Combo4.List(I))
        
        Cad = Mid(Cad, InStr(1, Cad, "[") + 1)  'quitamos primer corchete
        Cad = Mid(Cad, 1, Len(Cad) - 1)  'quitamos segundo corchete
        
        If Cad = ListView2.Tag Then
            'Este es el trabajador
            Combo4.ListIndex = I
            Exit For
        End If
    Next
    ListView2.Tag = ""
    
End Sub


Private Sub Timer1_Timer()
    T1 = DateAdd("s", 1, T1)
    Labels
End Sub

Private Sub Labels()
    Me.Label1(0).Caption = Format(T1, "dd/mm/yyyy")
    Me.Label1(1).Caption = Format(T1, "hh:mm:ss")
    
    
    
End Sub


Private Sub CagarMarcajes()
Dim IT As ListItem
Dim parte As String
Dim Repar As String
    limpiar Me
    Combo1.ListIndex = -1
    Set miRsAux = New ADODB.Recordset
    Me.ListView2.ListItems.Clear
    
    'CAMBIAR
    
    Cad = "select sreloj.*,nomtraba,nomclien,nomtipor from sreloj inner join straba on sreloj.codtraba=straba.codtraba"
    Cad = Cad & " LEFT JOIN scaalb ON scaalb.codtipom = sreloj.codtipom AND sreloj.numalbar=scaalb.numalbar"
    Cad = Cad & " LEFT JOIN stipor ON sreloj.codtipor = stipor.codtipor"
    Cad = Cad & " where fecha=" & DBSet(Label1(0).Caption, "F")
    Cad = Cad & " ORDER BY HoraInicio,HoraFin"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Repar = ""
    parte = ""
    While Not miRsAux.EOF
        
        Set IT = ListView2.ListItems.Add()
        IT.Text = miRsAux!CodTraba
        IT.SubItems(1) = miRsAux!NomTraba
        
        IT.SubItems(2) = DBLet(miRsAux!codtipom, "T") & " "
        IT.SubItems(3) = DBLet(miRsAux!NumAlbar, "T") & " "
        
        If IsNull(miRsAux!NomTipor) Then
            Cad = "--- ** No encotrado"
        Else
            Cad = miRsAux!NomTipor
        End If
        Cad = miRsAux!codtipor & " " & Cad
        IT.SubItems(4) = Cad
        IT.SubItems(5) = Format(miRsAux!HoraInicio, "hh:mm")
        IT.Tag = Format(miRsAux!HoraInicio, "hh:mm:ss")
        
        If IsNull(miRsAux!HoraFin) Then
            IT.SubItems(6) = " "
        Else
            IT.SubItems(6) = Format(miRsAux!HoraFin, "hh:mm")
        End If
        
        
        If IsNull(miRsAux!codtipom) Then
            'ORden de produccion
            IT.SubItems(7) = "Prod. " & Format(miRsAux!NumAlbar, "000000")
        Else
        
            IT.SubItems(7) = DBLet(miRsAux!Nomclien, "T") & " "
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If Repar <> "" Then
    
    End If
    
    If parte <> "" Then
    
    End If
    
End Sub


Private Sub PonerFrames2()
Dim b As Boolean

    


    
    If Combo1.ListIndex < 0 Then
        b = False
    Else
        b = HayQueCerrarNodo = 0
    End If
    
    Frame4.visible = True
    
    
    
End Sub


Private Sub CrearAlbaran()
Dim vC As CTiposMov
Dim vCli As CCliente
Dim HaCambiadoContador As Boolean
    On Error GoTo eCrearAlbaran
    
    Set vC = New CTiposMov
    Set vCli = New CCliente
    
    
    conn.BeginTrans

    
    'VErsion normal
    Cad = DBSet(vParam.CifEmpresa, "T") & " ORDER BY codclien"
    

    
    Cad = DevuelveDesdeBD(conAri, "codclien", "sclien", "nifclien", Cad)
    If Cad = "" Then Err.Raise 513, , "Obteniendo cliente EULER"
    
    If Not vCli.LeerDatos(Cad) Then Err.Raise 513, , "Obteniendo datos cliente EULER " & Cad
    
    'Febrero 2016
    'Si es EUSKADI o VALENCIA para los albaranes de reparacion cogera el CAR o el ALR
    
    Select Case Me.cboTipoTrabajo.ListIndex
    Case 2
        Cad = "ALE"
    Case 3
        Cad = "ALO"
    Case Else
            
            Cad = DevuelveDesdeBD(conAri, "codalmac", "straba", "codtraba", CStr(Combo1.ItemData(Combo1.ListIndex)))
            If Cad = "10" Then
                Cad = "CAR"
            Else
                Cad = "ALR"
            End If
    End Select
    
    vC.ConseguirContador Cad
    
    
    Cad = InputBox("N� " & vC.NombreMovimiento, , CStr(vC.Contador + 1))
    If Cad = "" Then Err.Raise 513, , "Proceso cancelado por el usuario"
    
    HaCambiadoContador = False
    If Val(Cad) <> vC.Contador Then
        vC.Contador = Val(Cad) - 1   '
        HaCambiadoContador = True
    End If
    Cad = "INSERT INTO scaalb(codtipom,numalbar,fechaalb,factursn,codclien,"
    Cad = Cad & "nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
    Cad = Cad & "facturkm,codtraba,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,tipofact,esticket) VALUES ('"
    If Me.cboTipoTrabajo.ListIndex = 1 Then
        'ALR
         Cad = Cad & "ALR"
    Else
        Cad = Cad & vC.TipoMovimiento
    End If
    Cad = Cad & "'," & vC.Contador + 1 & "," & DBSet(Label1(0).Caption, "F") & ",1," & vCli.codigo
    Cad = Cad & "," & DBSet(vCli.Nombre, "T") & "," & DBSet(vCli.Domicilio, "T") & "," & DBSet(vCli.CPostal, "T") & "," & DBSet(vCli.Poblacion, "T")
    Cad = Cad & "," & DBSet(vCli.Provincia, "T") & "," & DBSet(vCli.NIF, "T") & "," & DBSet(vCli.TfnoClien, "T", "S") & ",0," & Val(Combo1.ItemData(Combo1.ListIndex)) & "," & Val(Combo1.ItemData(Combo1.ListIndex))
    Cad = Cad & "," & DBSet(vCli.Agente, "T") & "," & DBSet(vCli.ForPago, "T") & "," & DBSet(vCli.FEnvio, "T") & ",0,0,0,0)"
    conn.Execute Cad
    
    If Not HaCambiadoContador Then
        vC.IncrementarContador vC.TipoMovimiento
    Else
        vC.Contador = vC.Contador + 1
    End If
    
eCrearAlbaran:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        conn.RollbackTrans
    Else
        conn.CommitTrans
        
        Cad = "Se ha generado el albaran:   " & IIf(vC.TipoMovimiento = "CAR", "ALR", vC.TipoMovimiento) & " " & Format(vC.Contador, "000000") & vbCrLf
        Cad = Cad & "--> " & vC.NombreMovimiento
        MsgBox Cad, vbInformation
        
        'Cargamos el combo de albaranes
        Cad = IIf(vC.TipoMovimiento = "CAR", "ALR", vC.TipoMovimiento)
        Cad = "codtipom = '" & Cad & "'"
        CargarCombo_Tabla Me.Combo3, "scaalb", "concat(numalbar,' - ',nomclien)", "NumAlbar", Cad
        
        
        'Situamos el
        For NumRegElim = 0 To Combo3.ListCount - 1
            Cad = Mid(Combo3.List(NumRegElim), 1, InStr(1, Combo3.List(NumRegElim), "-") - 1)
            If Val(Cad) = vC.Contador Then
                
                Combo3.ListIndex = NumRegElim
            End If
            
        Next
        
        
    End If

    Set vC = New CTiposMov
    Set vCli = New CCliente
    
End Sub
