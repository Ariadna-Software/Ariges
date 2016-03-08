VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmWH_ExpedientesPot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expedientes"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdVerPDF 
         Height          =   495
         Left            =   120
         Picture         =   "frmWH_ExpedientesPot.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   6015
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         _cx             =   5080
         _cy             =   5080
      End
      Begin VB.Label lblTituloDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   15
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.TextBox Text1 
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
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "José Luis Fernández Rodriguéz"
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "961398959"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      CausesValidation=   0   'False
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
      Index           =   2
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "654649836"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "002341"
      Top             =   480
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Tag             =   "Prop. comercial"
      Top             =   2400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3413
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
         Text            =   "F. Presentacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contestacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Acep."
         Object.Width           =   2011
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Tag             =   "Contrato"
      Top             =   4920
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4048
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
         Text            =   "F. Presentacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contestacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Acep."
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   1
      Left            =   1560
      Picture         =   "frmWH_ExpedientesPot.frx":1272
      ToolTipText     =   "Contrato aceptado"
      Top             =   4680
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   0
      Left            =   2520
      Picture         =   "frmWH_ExpedientesPot.frx":1C74
      ToolTipText     =   "Propuesta aceptada"
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contrato"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   5
      Left            =   1200
      Picture         =   "frmWH_ExpedientesPot.frx":2676
      ToolTipText     =   "Rechazar"
      Top             =   4680
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   4
      Left            =   840
      Picture         =   "frmWH_ExpedientesPot.frx":3078
      ToolTipText     =   "Nueva contrato"
      Top             =   4680
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Propuesta comercial"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1680
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   3
      Left            =   2160
      Picture         =   "frmWH_ExpedientesPot.frx":3A7A
      ToolTipText     =   "Rechazar"
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   2
      Left            =   1800
      Picture         =   "frmWH_ExpedientesPot.frx":447C
      ToolTipText     =   "Nueva propuesta"
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   7
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Télefono"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Télefono"
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   6
      Left            =   840
      Picture         =   "frmWH_ExpedientesPot.frx":4E7E
      Top             =   4680
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmWH_ExpedientesPot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ClientePot As Long

Dim Cad As String


Private Sub cmdVerPDF_Click()
    LanzaVisorMimeDocumento Me.hWnd, AcroPDF1.Tag
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPPalWhose.Icon
    limpiar Me
    Me.Tag = "select nomclien,telclie1 ,telclie2,whoExpedientePot.* from whoExpedientePot,sclipot where sclipot.codclien=whoExpedientePot.codclien AND whoExpedientePot.codclien =" & ClientePot
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Me.Tag, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "ERROR GRAVE", vbCritical
    Else
        Text1(3).Text = Format(miRsAux!codclien, "000000")
        Text1(0).Text = miRsAux!NomClien
        Text1(1).Text = DBLet(miRsAux!telclie1, "T")
        Text1(2).Text = DBLet(miRsAux!telclie2, "T")
        
        CargaListviewWHOSE ListView1, True, True, ClientePot, False
        CargaListviewWHOSE ListView2, True, False, ClientePot, False
        
        NumRegElim = 0
        If Not IsNull(miRsAux!fecAceptPropComer) Then
            NumRegElim = 1
            'El primer list lleva la aceptacion
            Me.ListView1.ListItems(1).SubItems(2) = Format(miRsAux!fecAceptPropComer, "dd/mm/yyyy")
        End If
        
            
        If NumRegElim = 1 Then
            If Not IsNull(miRsAux!fecAceptContrato) Then
                NumRegElim = 2
                Me.ListView2.ListItems(1).SubItems(2) = Format(miRsAux!fecAceptContrato, "dd/mm/yyyy")
            End If
        End If

        IconosVisibles CByte(NumRegElim)
        
    End If
End Sub

'0.-Propuestas comertcial    1.- Contrato    2.-SOLO el de pasar a cliente
Private Sub IconosVisibles(Cual As Byte)
    Image3(3).visible = Cual = 0
    Image3(2).visible = Cual = 0
    Image3(0).visible = Cual = 0
    
    Image3(5).visible = Cual = 1
    Image3(4).visible = Cual = 1
    Image3(1).visible = Cual = 1
    
    Image3(6).visible = Cual = 2
    
End Sub


Private Sub Image1_Click(Index As Integer)

End Sub

Private Sub Image3_Click(Index As Integer)
        
        
    Select Case Index
    '*******************************
    'OFERTAS
    Case 0, 2, 3
    
        If Index = 2 Then
            'NUEVA OFERTA
            If ListView1.ListItems.Count > 0 Then
                'Si el item de arriba NO esta cerrado NO dejo meter otro
                If Trim(ListView1.ListItems(1).SubItems(1)) = "" Then
                    MsgBox "Rechace la oferta anterior", vbExclamation
                    Exit Sub
                                
                End If
            End If
            'Opcion para abrir
            frmWH_Varios.Opcion = 0
        Else
            'Rechazar o ACEPTAR
            Cad = ""
            If Me.ListView1.ListItems.Count = 0 Then
                Cad = "Ninguna propuesta disponible"
            Else
                'Si el primer item(que es el que podemos rechazar, ESTA ya rechazado no podemos continuar
                If Trim(ListView1.ListItems(1).SubItems(1)) <> "" Then Cad = "Rechazada"
            End If
            If Cad <> "" Then
                MsgBox Cad, vbExclamation
                Exit Sub
            End If
             
            If Index = 3 Then
                frmWH_Varios.Opcion = 1 'Abrir pantalla pedir FECHA rechazo
            Else
                frmWH_Varios.Opcion = 2 'Abrir pantalla pedir FECHA aceptacion
            End If
        End If
        'PROPUESTA COMERCIAL
        CadenaDesdeOtroForm = "01/01/2000"
        If Me.ListView1.ListItems.Count > 0 Then CadenaDesdeOtroForm = Me.ListView1.ListItems(1).Text
        frmWH_Varios.ExtraData2 = ClientePot & "|" & CadenaDesdeOtroForm & "|1|"
        
        CadenaDesdeOtroForm = ""
        frmWH_Varios.Show vbModal
        
        If Index = 3 Then
            'MEter RECHAZO con la fecha seleccionada
            If CadenaDesdeOtroForm <> "" Then
                Cad = "UPDATE whoexpedientepotprocomer set f_rechazoprop =" & DBSet(CadenaDesdeOtroForm, "F")
                Cad = Cad & " WHERE codclien=" & ClientePot & " and idpropcomer=  " & Mid(ListView1.ListItems(1).Key, 2)
                If Ejecutar(Cad, False) Then ListView1.ListItems(1).SubItems(1) = CadenaDesdeOtroForm
                    
            End If
        ElseIf Index = 0 Then
            'Propuesta comercial aprobada
            If CadenaDesdeOtroForm <> "" Then
                Cad = "UPDATE whoExpedientePot set fecAceptPropComer =" & DBSet(CadenaDesdeOtroForm, "F") & " WHERE codclien= " & ClientePot
                If Ejecutar(Cad, False) Then
                    ListView1.ListItems(1).SubItems(2) = CadenaDesdeOtroForm
                    IconosVisibles 1  'Los de contrato
                End If
            End If
        Else
            CargaListviewWHOSE ListView1, True, True, ClientePot, False
        End If
        
            
    
    Case 1, 4, 5
        
        'CONTRATO CONTRATO
        If Index = 4 Then
            'NUEVA contrato
            If ListView2.ListItems.Count > 0 Then
                'Si el item de arriba NO esta cerrado NO dejo meter otro
                If Trim(ListView2.ListItems(1).SubItems(1)) = "" Then
                    MsgBox "Rechace el contrato anterior", vbExclamation
                    Exit Sub
                                
                End If
            End If
            'Opcion para abrir
            frmWH_Varios.Opcion = 0
        Else
            'Rechazar o ACEPTAR
            Cad = ""
            If Me.ListView2.ListItems.Count = 0 Then
                Cad = "Ningún contrato disponible"
            Else
                'Si el primer item(que es el que podemos rechazar, ESTA ya rechazado no podemos continuar
                If Trim(ListView2.ListItems(1).SubItems(1)) <> "" Then Cad = "Rechazado"
            End If
            If Cad <> "" Then
                MsgBox Cad, vbExclamation
                Exit Sub
            End If
             
            If Index = 5 Then
                frmWH_Varios.Opcion = 1 'Abrir pantalla pedir FECHA rechazo
            Else
                frmWH_Varios.Opcion = 2 'Abrir pantalla pedir FECHA aceptacion
            End If
        End If
        'CONTRATO
        CadenaDesdeOtroForm = "01/01/2000"
        If Me.ListView2.ListItems.Count > 0 Then CadenaDesdeOtroForm = Me.ListView2.ListItems(1).Text
        frmWH_Varios.ExtraData2 = ClientePot & "|" & CadenaDesdeOtroForm & "|2|" '1 prop 2 Contrato
        
        CadenaDesdeOtroForm = ""
        frmWH_Varios.Show vbModal
        
        If Index = 5 Then
            'whoexpedientepotcontrato idcontrato f_rechazocon
            'MEter RECHAZO con la fecha seleccionada
            If CadenaDesdeOtroForm <> "" Then
                Cad = "UPDATE whoexpedientepotcontrato set f_rechazocon =" & DBSet(CadenaDesdeOtroForm, "F")
                Cad = Cad & " WHERE codclien=" & ClientePot & " and idcontrato=  " & Mid(ListView2.ListItems(1).Key, 2)
                If Ejecutar(Cad, False) Then ListView2.ListItems(1).SubItems(1) = CadenaDesdeOtroForm
                    
            End If
        ElseIf Index = 1 Then
            'contrato  aprobado
            If CadenaDesdeOtroForm <> "" Then
                Cad = "UPDATE whoExpedientePot set fecAceptContrato =" & DBSet(CadenaDesdeOtroForm, "F") & " WHERE codclien= " & ClientePot
                If Ejecutar(Cad, False) Then
                    ListView2.ListItems(1).SubItems(2) = CadenaDesdeOtroForm
                    IconosVisibles 2  'NINGUNO
                    
                    Cad = "El contrato ha sido aceptado. " & vbCrLf & vbCrLf & "Quiere asignar los datos de cliente definitivo en este momento?"
                    Cad = Cad & vbCrLf & vbCrLf & "Podra realizar el proceso en otro momento si lo desea"
                    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                        Image3_Click 6
                        Exit Sub
                    End If
                    
                End If
                
            End If
        Else
            CargaListviewWHOSE ListView2, True, False, ClientePot, False
        End If
        
    
    Case 6
        'de momento
        frmWH_Varios.ExtraData2 = CStr(ClientePot)
        frmWH_Varios.Opcion = 3
        frmWH_Varios.Show vbModal
    
        Cad = DevuelveDesdeBD(conAri, "codclien", "sclipot", "codclien", CStr(ClientePot))
        If Cad = "" Then
            'YA no existe el cliente potencial. Se ha pasado a CLIENTES
            Unload Me
        End If
    End Select
    
        
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEM(ListView1.SelectedItem, True, ClientePot)
    LanzaVisorMimeDocumento Me.hWnd, Cad
    
    
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PonerPDF ListView1, True
End Sub

Private Sub ListView2_DblClick()
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEM(ListView2.SelectedItem, False, ClientePot)
    LanzaVisorMimeDocumento Me.hWnd, Cad

End Sub


Private Sub PonerPDF(ByRef LISTV As ListView, PropComer As Boolean)
    
    If LISTV.SelectedItem Is Nothing Then Exit Sub
        
    If LCase(Right(LISTV.SelectedItem.Tag, 3)) <> "pdf" Then Exit Sub
    
    Cad = DevuelveNombreArhivoITEM(LISTV.SelectedItem, PropComer, ClientePot)
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Me.lblTituloDoc.Caption = LISTV.Tag & " " & LISTV.SelectedItem.Tag
        Me.lblTituloDoc.Refresh
        
        cmdVerPDF.visible = False
        If Not CargaArchivo Then
            Me.lblTituloDoc.Caption = "ERROR " & Me.lblTituloDoc.Caption
        Else
            cmdVerPDF.visible = True
        End If
        
        Screen.MousePointer = vbDefault
    End If
    
    
    
    
    
    
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PonerPDF ListView2, False
End Sub

Private Function CargaArchivo() As Boolean
    
    On Error GoTo eCargaArchivo
    CargaArchivo = False
    
    
    AcroPDF1.LoadFile (Cad)
    AcroPDF1.Tag = Cad
    AcroPDF1.visible = True
    Screen.MousePointer = vbDefault
    
    
    CargaArchivo = True
    Exit Function
eCargaArchivo:
    MuestraError Err.Number, "Carga archivo PDF"
End Function

