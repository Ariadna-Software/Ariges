VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelematImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar fichero telematel"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameProc 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   13455
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   9840
         TabIndex        =   11
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.CheckBox chkMultiProveedor 
      Alignment       =   1  'Right Justify
      Caption         =   "Multiproveedor"
      Height          =   375
      Left            =   10560
      TabIndex        =   13
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtProv 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox txtProv 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   6960
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdActualizar 
      Height          =   375
      Left            =   12480
      Picture         =   "frmTelematImportar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Actualizar datos"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   13080
      Picture         =   "frmTelematImportar.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   240
      Width           =   375
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   6015
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   10610
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod.tel"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   8017
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "EAN"
         Object.Width           =   2593
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Precio"
         Object.Width           =   2265
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ref. prov."
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Ud"
         Object.Width           =   1614
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fec. cambio"
         Object.Width           =   2011
      EndProperty
   End
   Begin VB.CommandButton cmdImportar 
      Height          =   375
      Left            =   6360
      Picture         =   "frmTelematImportar.frx":0F8C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Importar fichero"
      Top             =   240
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   -120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.CheckBox chkExcel 
      Alignment       =   1  'Right Justify
      Caption         =   "Formato excel"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   0
      Width           =   1695
   End
   Begin VB.CheckBox chkCabel 
      Alignment       =   1  'Right Justify
      Caption         =   "CABEL"
      Height          =   375
      Left            =   8640
      TabIndex        =   14
      Top             =   0
      Width           =   975
   End
   Begin VB.Image imgayuda 
      Height          =   240
      Index           =   1
      Left            =   9720
      ToolTipText     =   "Buscar cliente"
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgBuscarG 
      Height          =   240
      Index           =   0
      Left            =   7800
      Picture         =   "frmTelematImportar.frx":198E
      ToolTipText     =   "Buscar cliente"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgBuscarG 
      Height          =   240
      Index           =   1
      Left            =   840
      Picture         =   "frmTelematImportar.frx":1A90
      ToolTipText     =   "Buscar cliente"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmTelematImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmP As frmComProveedores
Attribute frmP.VB_VarHelpID = -1

Dim cad As String
Dim NF As Integer

Dim RArt As ADODB.Recordset  'para pre cargar los articulos del proveedor



Private Sub cmdActualizar_Click()



    cad = ""
    If lw1.ListItems.Count = 0 Then
        cad = "Ningun dato"
    Else
        If Me.chkCabel.Value And Me.chkMultiProveedor.Value Then
           ' Cad = "Debe indicar una de las dos opciones: Cabel / Multiproveedor "
        Else
            If Me.chkCabel.Value Then
                If txtProv(0).Text <> "" Then cad = "No debe indicar proveedor"
            Else
                If txtProv(0).Text = "" Or txtProv(1).Text = "" Then cad = "Falta proveedor"
            End If
        End If
    End If
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Sub
    End If
    
    If Me.chkCabel.Value Then
        cad = "Procesar los datos del fichero CABEL?"
    Else
        cad = "Seguro que desea actualizar los datos para el proveedor: " & vbCrLf & Format(txtProv(0).Text, "0000") & " -  " & txtProv(1).Text
        If Me.chkMultiProveedor.Value = 1 Then cad = cad & vbCrLf & vbCrLf & "Fichero multiproveedor"
    End If
    If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Me.FrameProc.visible = True
    Me.Label2(0).Caption = "Comienzo proceso. Leyendo BD"
    Me.Label2(1).Caption = ""
    Me.Refresh
    Actualizar
    Me.FrameProc.visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Actualizar()
Dim Aux As String
Dim TieneError As Boolean

Dim CodTelemYaInsertados As String
Dim HayQueInsertarTelematel As Boolean

    'OK
    Set miRsAux = New ADODB.Recordset
    Set RArt = New ADODB.Recordset
    Espera 0.2
    Me.Label2(0).Caption = "Abriendo RS"
    Me.Label2(0).Refresh
    
   
    
    If Me.chkCabel.Value Then
      
      '  Cad = "Select * from stelem ORDER BY codtelem,referprov"
      '  miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        'Cad = "Select codartic,referprov,codprove FROM sartic where codfamia IN (select codfamia from sfamia where nomfamia like '%cabel%') ORDER BY codartic,codprove"
        'RArt.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Else
    
        cad = "Select * from stelem where codprove = " & txtProv(0).Text & " ORDER BY codtelem,referprov"
        miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
        'Como antes
        cad = "Select codartic,referprov from sartic where codprove = " & txtProv(0).Text & " ORDER BY codtelem"
        RArt.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
    End If
    
    
    
    

    TieneError = False
    CodTelemYaInsertados = ""
    For NF = 1 To lw1.ListItems.Count
        Me.Label2(0).Caption = lw1.ListItems(NF).SubItems(1)
        Me.Label2(1).Caption = NF & " de " & lw1.ListItems.Count
        Me.Label2(0).Refresh
        Me.Label2(1).Refresh
        If (NF Mod 50) = 0 Then DoEvents
        'BUscamos codartic
        Aux = BuscarCodartic(NF)
        If Aux <> "" Then Debug.Print lw1.ListItems(NF).Text
        'Vemos si existe en telematel
        If ExisteEnStelem(NF) Then
            'UPDATEAMOS
             cad = "update `stelem` set nombre=" & DBSet(lw1.ListItems(NF).SubItems(1), "T")
             cad = cad & ",`codean`=" & DBSet(lw1.ListItems(NF).SubItems(2), "T", "S")
             cad = cad & ",`precio`=" & DBSet(lw1.ListItems(NF).SubItems(3), "N")
             cad = cad & ",`referprov`=" & DBSet(lw1.ListItems(NF).SubItems(4), "T")
             cad = cad & ",`uniprec`=" & DBSet(lw1.ListItems(NF).SubItems(5), "N")
             cad = cad & ",`fechacambio`=" & DBSet(lw1.ListItems(NF).SubItems(6), "F")
             
             If Aux <> "" Then
                'EXISTE EL ARTICULO
                cad = cad & ",`codartic`=" & DBSet(Aux, "T")
                
             Else
                'Agosto2011
                'Puede que no exista o puede que la referencia la hayan cambiado
                'y no lo encuentre por eso
                CodTelemYaInsertados = CodTelemYaInsertados & ", " & DBSet(lw1.ListItems(NF).Text, "T")
             End If
                
                
                
             
            cad = cad & " where `codtelem`=" & DBSet(lw1.ListItems(NF).Text, "N")
            
        Else
            'INSERTAMOS
            HayQueInsertarTelematel = False
            
            If chkMultiProveedor.Value = 0 Then HayQueInsertarTelematel = True
           
            
            If HayQueInsertarTelematel Then


                cad = "insert into `stelem` (`codtelem`,`nombre`,`codean`,`precio`,`referprov`,`uniprec`,`fechacambio`,"
                cad = cad & "`codartic`,`codprove`) values (" & lw1.ListItems(NF).Text & "," & DBSet(lw1.ListItems(NF).SubItems(1), "T")
                'ean precio
                cad = cad & "," & DBSet(lw1.ListItems(NF).SubItems(2), "T", "S") & "," & DBSet(lw1.ListItems(NF).SubItems(3), "N")
                'refprov uni
                cad = cad & "," & DBSet(lw1.ListItems(NF).SubItems(4), "T") & "," & DBSet(lw1.ListItems(NF).SubItems(5), "N") & "," & DBSet(lw1.ListItems(NF).SubItems(6), "F")
                cad = cad & "," & DBSet(Aux, "T", "S") & ","
                
                If Me.chkCabel.Value = 1 Then
                    cad = cad & "NULL)"
                Else
                    cad = cad & txtProv(0) & ")"
                End If
            Else
                'SI esta multiproveedor, si no la encuentra, NO la inserta
                cad = ""
                
            End If
        End If
        
        If cad <> "" Then
            If Not ejecutar(cad, False) Then TieneError = True
        End If
        
        
        
        
        
        
    Next NF
     
     
    miRsAux.Close
    RArt.Close
    
    'Junio16--> ERROR CABEL
    'Agosto 2011
    'Puede que haya cambiado la referencia, pero que el artiulo sea EL mismo
    'Para aquellos telem que ya existian comprobaremos si la tienen un codartic asignado
    'No los miraremos en todos, solo en aquellos que ya exisiteran y NO haya encontrado el codigo
    'esos telem me los guardo en CodTelemYaInsertados
    If CodTelemYaInsertados <> "" Then
        Me.Label2(0).Caption = "Comprobacion referencias capturadas"
        Me.Label2(1).Caption = ""
        Me.Label2(0).Refresh
        Me.Label2(1).Refresh
        CodTelemYaInsertados = Mid(CodTelemYaInsertados, 2)
        
        
        
        cad = "Select * from stelem where "
        'ANÑADO lo de txtpro="" el 25 ABril
        If txtProv(0).Text = "" Then
            cad = cad & " codprove is null"
        Else
            cad = cad & " codprove=" & txtProv(0).Text
        End If
        cad = cad & " AND codartic<>"""""
        cad = cad & " AND codtelem in (" & CodTelemYaInsertados & ")"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Me.Label2(1).Caption = DBLet(miRsAux!Nombre, "T")
            Me.Label2(1).Refresh
            cad = "UPDATE sartic set "
            cad = cad & " `referprov`=" & DBSet(miRsAux!referprov, "T")
            
            cad = cad & " WHERE `codartic`=" & DBSet(miRsAux!codArtic, "T")
            ejecutar cad, False
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
         Me.Label2(1).Caption = ""
    End If
    Set miRsAux = Nothing
    Set RArt = Nothing

    
    
    If Not TieneError Then
        MsgBox "Proceso finalizado con exito", vbExclamation
        lw1.ListItems.Clear
        Text1.Text = ""
        txtProv(0).Text = ""
        txtProv(1).Text = ""
    End If
End Sub

Private Function ExisteEnStelem(Ind As Integer) As Boolean
    ExisteEnStelem = False
    
    If Me.chkCabel.Value Then
        If NF > 1 Then miRsAux.Close
        miRsAux.Open "Select * from stelem WHERE codtelem = " & lw1.ListItems(Ind).Text, conn, adOpenKeyset, adLockPessimistic, adCmdText
    Else
        miRsAux.Find "codtelem = " & lw1.ListItems(Ind).Text, , adSearchForward, 1
    End If
    If Not miRsAux.EOF Then
        'Compruebo tb la referencia del proveedor?
        
        
        '
        ExisteEnStelem = True
    End If
End Function



Private Function BuscarCodartic(Ind As Integer) As String
    
       
    If Me.chkCabel.Value Then
        If NF > 1 Then RArt.Close
        
        RArt.Open "Select codartic,referprov,codprove FROM sartic where referprov = " & DBSet(lw1.ListItems(Ind).SubItems(4), "T"), conn, adOpenKeyset, adLockPessimistic, adCmdText
        
    Else
        RArt.Find "referprov = " & DBSet(lw1.ListItems(Ind).SubItems(4), "T"), , adSearchForward, 1
    End If
    If RArt.EOF Then
        BuscarCodartic = ""
    Else
        
        BuscarCodartic = RArt!codArtic
        
    End If
End Function






Private Sub cmdImportar_Click()

   

    cad = ""
    If Text1.Text = "" Then
        cad = "Seleccione el fichero para importar"
    Else
        If Dir(Text1.Text, vbArchive) = "" Then
            cad = "No existe el fichero"
        Else
            If Me.chkExcel.Value = 1 Then
                'EXCEL
                If UCase(Right(Text1.Text, 3)) <> "XLS" Then
                    cad = "Extension invalida. XLS"
                    
                Else
                    'existe el fichero y es una excel
                    'Veremos el fichero TXT que se genera.
                    cad = EliminarFicheroTXTdeEXCEL
                    
                End If
            Else
                If UCase(Right(Text1.Text, 3)) <> "TXT" Then cad = "Extension invalida."
            End If
            
        End If
    End If
    
    
    
    
    
    
    
    
    
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Sub
    End If
        
    'Por si acaso ya existen datos
    If lw1.ListItems.Count > 0 Then
        cad = "Ya existen datos. Desea importar el fichero?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
        
        
    Screen.MousePointer = vbHourglass
    lw1.ListItems.Clear
    If Me.chkExcel.Value = 1 Then
        'Procesar fichero excel
        If Not ProcesarFicheroExcel Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'Si ha ido bien cambio el text1
        Text1.Tag = Text1.Text
        Text1.Text = FicheroExcelConvertido
        
    End If
    ProcesarFichero
    
    If Me.chkExcel.Value = 1 Then Text1.Text = Text1.Tag
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()

'    Dim I As Integer
'    Cad = ""
'    For I = 1 To lw1.ColumnHeaders.Count
'        Cad = Cad & lw1.ColumnHeaders(I).Text & ": " & lw1.ColumnHeaders(I).Width & vbCrLf
'    Next
'    MsgBox Cad
    If lw1.ListItems.Count > 0 Then
        If MsgBox("Seguro que desea salir?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
    'NO existe el exe
    If Dir(App.Path & "\aTelemat.exe", vbArchive) = "" Then
        Me.chkExcel.Value = 0
        Me.chkExcel.visible = False
    Else
        Me.chkExcel.Value = 1
        Me.chkExcel.visible = True
    End If
    CargaIconosAyuda
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtProv(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtProv(1).Text = RecuperaValor(CadenaSeleccion, 2)
    
End Sub

Private Sub imgayuda_Click(index As Integer)
    MsgBox "Al importar los datos en telematel, los inserta sin asignarle codigo de proveedor", vbExclamation
End Sub

Private Sub imgBuscarG_Click(index As Integer)
    If index = 0 Then
        Set frmP = New frmComProveedores
        frmP.DatosADevolverBusqueda = "0|1|"
        frmP.Show vbModal
        Set frmP = Nothing
    
    Else
        If Me.chkExcel.Value = 0 Then
            'TELEMATEL
            cd1.Filter = "Texto *.txt|*.txt"
        Else
            'DEsde un fichero EXCEL.
            'Llamara al programa "atelemat.exe" que de la excel genera
            'el mismo nombre de fichero  pero txt y con formato telematel
            cd1.Filter = "Excel *.xls|*.xls"
        End If
        cd1.ShowOpen
        If cd1.FileName <> "" Then
            If Dir(cd1.FileName, vbArchive) <> "" Then Text1.Text = cd1.FileName
        End If
    End If
End Sub




'-----------------------------------------------------------------------------
Private Sub ProcesarFichero()
Dim OK As Boolean
Dim IT As ListItem
On Error GoTo EprocesarFichero
    NF = FreeFile
    Open Text1.Text For Input As #NF
    OK = EOF(NF)
    NumRegElim = 0
    CadenaDesdeOtroForm = ""
    While Not OK
        Line Input #NF, cad
        Set IT = lw1.ListItems.Add()
        If Not ProcesarLinea(IT) Then
            lw1.ListItems.Remove IT.index
            '¿Continuamos?
            'if msgbox("¿Continuar
            OK = True
        Else
            OK = EOF(NF)
        End If
    Wend
    NumRegElim = Val(CadenaDesdeOtroForm)
    lw1.ListItems(NumRegElim).EnsureVisible
    Set lw1.SelectedItem = lw1.ListItems(NumRegElim)
    lw1.SetFocus
EprocesarFichero:
    If Err.Number <> 0 Then MuestraError Err.Number
    On Error Resume Next
    Close #NF
End Sub



Private Function ProcesarLinea(ByRef IT As ListItem) As Boolean
Dim J As Integer
Dim Inicio As Integer
Dim I As Integer
Dim Aux As String
    
    
    ProcesarLinea = False
    Inicio = 1
    For I = 1 To 7
        J = InStr(Inicio, cad, ";")
        If J = 0 Then
            MsgBox "No se ha encontrado el separador " & J & ". Campo: " & I, vbExclamation
            Exit Function
        Else
            Aux = Mid(cad, Inicio, J - Inicio)
            Aux = Trim(Aux)
            If I = 1 Then
                IT.Text = Aux
            Else
                If I = 6 Then Aux = Val(Aux)
                
                IT.SubItems(I - 1) = Aux
            End If
            Inicio = J + 1
            
            
            
            If I = 2 Then
                If Len(Aux) > NumRegElim Then
                    NumRegElim = Len(Aux)
                    CadenaDesdeOtroForm = IT.index
                End If
            End If
        End If
    Next
    ProcesarLinea = True
End Function


Private Sub txtProv_GotFocus(index As Integer)
    ConseguirFoco txtProv(index), 3
End Sub

Private Sub txtProv_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtProv_LostFocus(index As Integer)
    If index = 0 Then
        cad = ""
        txtProv(0).Text = Trim(txtProv(0).Text)
        If txtProv(0).Text <> "" Then
            If PonerFormatoEntero(txtProv(0)) Then
                cad = DevuelveDesdeBD(conAri, "nomprove", "sprove", "codprove", txtProv(0).Text)
                If cad = "" Then
                    MsgBox "No existe el proveedor: " & txtProv(0).Text, vbExclamation
                    txtProv(0).Text = ""
                    PonerFoco txtProv(0)
                End If
            Else
                txtProv(0).Text = ""
            End If
        End If
        txtProv(1).Text = cad
    End If
End Sub

Private Function EliminarFicheroTXTdeEXCEL() As String
Dim Aux As String

    On Error GoTo eEliminarFicheroTXTdeEXCEL

    EliminarFicheroTXTdeEXCEL = ""
    Aux = FicheroExcelConvertido
    If Dir(Aux, vbArchive) <> "" Then Kill Aux
       

    Exit Function
eEliminarFicheroTXTdeEXCEL:
    EliminarFicheroTXTdeEXCEL = Err.Description
End Function


Private Function FicheroExcelConvertido() As String
Dim I As Integer


    I = InStrRev(Text1.Text, ".")
    
    FicheroExcelConvertido = Mid(Text1.Text, 1, I) & "txt"



End Function

Private Function ProcesarFicheroExcel() As Boolean
Dim Aux As String
Dim I As Integer
    
    On Error GoTo eProcesarFicheroExcel:

    ProcesarFicheroExcel = False
    
    Aux = App.Path & "\aTelemat.exe  /" & Text1.Text
    Shell Aux, vbNormalFocus
    Aux = FicheroExcelConvertido
    I = 0
    'Como mucho un minuto
    Caption = "     ******  Procesando fichero XLS   ******"
    Do
       
       If Dir(Aux, vbArchive) <> "" Then
            'OK, ya se ha creado el archivo
            I = 61
            ProcesarFicheroExcel = True
        Else
            I = I + 1
            Me.Refresh
            Screen.MousePointer = vbHourglass
            DoEvents
            Espera 0.9
        End If
       
    Loop Until I > 60
    
eProcesarFicheroExcel:
    If Err.Number <> 0 Then MuestraError Err.Number
    Screen.MousePointer = vbDefault
    Caption = "Importar fichero telematel"
End Function
    
Private Sub CargaIconosAyuda()
Dim Ima As Image
    On Error Resume Next 'mejor que no diera errores, pero bien, tampoco vamos a enfadarnos
    For Each Ima In Me.imgayuda
        Ima.Picture = frmPpal.imgListComun.ListImages(46).Picture
    Next
    Err.Clear
End Sub

