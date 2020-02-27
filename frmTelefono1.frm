VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelefono1 
   Caption         =   "Utilidades de telefonía"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   Icon            =   "frmTelefono1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameVerdatos 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmdAriadna 
         Caption         =   "Ariadna"
         Height          =   495
         Left            =   5640
         TabIndex        =   35
         Top             =   6600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminarDatosFracion 
         Caption         =   "Elim. fich."
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   6600
         Width           =   975
      End
      Begin VB.CheckBox chkMostrarBase 
         Caption         =   "Base imponible"
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   6780
         Width           =   1575
      End
      Begin VB.ComboBox cboFichero 
         Height          =   315
         Index           =   0
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   400
         Width           =   3255
      End
      Begin VB.CommandButton cmdFacturar 
         Caption         =   "Facturar"
         Height          =   495
         Left            =   1200
         TabIndex        =   14
         Top             =   6600
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Index           =   0
         Left            =   6840
         TabIndex        =   11
         Top             =   6600
         Width           =   1215
      End
      Begin MSComctlLib.ListView lwT 
         Height          =   5415
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Telefono"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Plz"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "OrdenTotal"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblI 
         Alignment       =   1  'Right Justify
         Caption         =   "Datos fichero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4200
         TabIndex        =   36
         Top             =   6240
         Width           =   3630
      End
      Begin VB.Label Label1 
         Caption         =   "Ficheros disponibles:"
         Height          =   195
         Index           =   5
         Left            =   3960
         TabIndex        =   17
         Top             =   6780
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Ficheros disponibles:"
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Datos pre-factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   2385
      End
   End
   Begin VB.Frame FrameImportacion 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8415
      Begin VB.ComboBox cboCompanyia2 
         Height          =   315
         ItemData        =   "frmTelefono1.frx":6852
         Left            =   120
         List            =   "frmTelefono1.frx":6859
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   6600
         TabIndex        =   7
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1140
         Width           =   6015
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   6360
         TabIndex        =   4
         Top             =   1140
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Compañia"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   3480
         Picture         =   "frmTelefono1.frx":6867
         ToolTipText     =   "Buscar ruta"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label lblinf 
         Alignment       =   2  'Center
         Caption         =   "Información de proceso"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Importación ficheros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Index           =   2
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   2850
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Fecha emisión facturas importadas:"
         Height          =   435
         Index           =   0
         Left            =   6360
         TabIndex        =   9
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre cualificado del fichero de importación"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.Frame FrameDtosTelefonia 
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdListadoDto 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   3360
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cboFichero 
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Index           =   2
         Left            =   4920
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   6090
      End
   End
   Begin VB.Frame FrameCatadau 
      Height          =   2895
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtCatadau2 
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   25
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton cmdCSV 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Index           =   5
         Left            =   3960
         TabIndex        =   27
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtCatadau2 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label lblInf2_ANTIGUO 
         Caption         =   "Información de proceso"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Dígito factura"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Importación CSV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Index           =   8
         Left            =   1080
         TabIndex        =   29
         Top             =   360
         Width           =   2970
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmTelefono1.frx":6969
         ToolTipText     =   "Buscar ruta"
         Top             =   960
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cmmDia 
      Left            =   0
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTelefono1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    ' 0.- Importar=FALS
    ' 1.- Importar
    
    ' 2.- Listado descuentos comprataiivo copera
    ' 3.- Rsumen fracion
    ' 4.- Datos face

    
    ' 5.- Importacion CATADAU

    ' 6.- Datos importados

Dim cad As String
Dim I As Integer


Dim IVA_standard As Currency




Private Sub cboCompanyia2_KeyPress(KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub cboFichero_Click(index As Integer)
    If index = 0 Then
        If Me.cboFichero(index).ListIndex < 0 Then Exit Sub
        Screen.MousePointer = vbHourglass
        CargarListView cboFichero(index).List(cboFichero(index).ListIndex)
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdAriadna_Click()
Dim Normales As Boolean
Dim idBanco As Integer
Dim SQL As String

        Normales = True
        If MsgBox("Normales", vbQuestion + vbYesNo) <> vbNo Then Normales = False

        SQL = InputBox("banco", "", "1")
        If SQL = "" Then Exit Sub
        idBanco = Val(SQL)
        
        SQL = InputBox("Fichero", "", "")
        If SQL = "" Then Exit Sub
        
        EstableceValoresFacturaTelefoniaROOT SQL
        
        'cboFichero(0).List(cboFichero(0).ListIndex), Label1(5), CInt(CadenaDesdeOtroForm))
        '(cboFichero(0).ListIndex), Label1(5), CInt(CadenaDesdeOtroForm))
         GenerarFacturasTelefonia idBanco, Label1(5), Normales, False

End Sub

Private Sub cmdCSV_Click()

    MsgBox "Aqui no deberia entrar"
End Sub

Private Sub HacerCoarval()

    
    
    'ANTES
    'cad = ""
    'If Trim(Me.txtCatadau(0).Text) = "" Or Me.txtCatadau(1).Text = "" Then
    '    cad = "Campos obligatorios"
    'Else
    '    If Not IsNumeric(Me.txtCatadau(1).Text) Then
    '        cad = "Dígito incorrecto"
    '    Else
    '        If Dir(Me.txtCatadau(0).Text, vbArchive) = "" Then cad = "No existe el archivo"
    '    End If
    'End If
    
    'AHORA
    
    I = -1
    cad = DevuelveDesdeBD(conAri, "DigitoCoarval", "spara2", "1", "1")
    If cad <> "" Then I = Val(cad)
    
    If I < 0 Then
        cad = "-No esta establecido el digito de coarval"
    Else
        cad = ""
    End If
    
    If Text1.Text = "" Then
        cad = "-Falta fichero"
    Else
        If Dir(Text1.Text, vbArchive) = "" Then cad = "-No existe el archivo:" & Text1.Text
    End If
    
    If Text2(0).Text <> "" Then
        cad = "-La fecha de las facturas la lleva el fichero. NO indique ninguna" & vbCrLf & cad
    End If
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Sub
    End If
    
            
    cad = DevuelveDesdeBD(conAri, "count(*)", "scaalb", "codtipom", "ALT", "T")
    If cad = "" Then cad = "0"
    
    If Val(cad) > 0 Then
        MsgBox "Albaranes telefonia pendientes de facturar. Avise soporte técnico", vbExclamation
        Exit Sub
    End If
    
    cad = PonerTrabajadorConectado("")
    If cad = "" Then
        MsgBox "Error obteniendo datos trabajador conectado", vbExclamation
        Exit Sub
    End If
            
            
    cad = "Continuar con la generacion de facturas de telefonía con el digito " & I & "?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    If GenerarImportacionCatadau(I) Then InsertarFacturasTelefonoCoarval
    
    Me.lblInf.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEliminarDatosFracion_Click()
Dim mGen2 As TelGenerador

    If cboFichero(0).ListCount = 0 Then Exit Sub
    If Me.lwT.ListItems.Count = 0 Then Exit Sub
    If vUsu.Nivel > 1 Then Exit Sub
    
    cad = "Desea eliminar el fichero de facturacion: " & cboFichero(0).Text & "?"
    If MsgBox(cad, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    If MsgBox("Seguro que desea eliminar el fichero?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Set mGen2 = New TelGenerador

'mGen2.EliminarTodoElFichero "CI0915168016", Label1(5)


    mGen2.EliminarTodoElFichero cboFichero(0).Text, Label1(5)
    
    Set mGen2 = Nothing
    lwT.ListItems.Clear
    
    cad = String(20, "*")
    cad = vbCrLf & vbCrLf & cad & cad & vbCrLf & vbCrLf
    cad = cad & "    REVISE LOS CONTADORES DE FACTURA    " & cad
    MsgBox cad, vbCritical
    Screen.MousePointer = vbDefault
    
    
    
    Unload Me
End Sub

Private Sub cmdFacturar_Click()
    'Algun dato a traspasar
    If Me.lwT.ListItems.Count = 0 Then Exit Sub
    
    
    
    
    Screen.MousePointer = vbHourglass
    Label1(5).Caption = "Comprobar"
    Label1(5).Refresh

    HacerFacturacionTelefonia
    
    Label1(5).Caption = ""
    Screen.MousePointer = vbDefault


End Sub

Private Sub HacerFacturacionTelefonia()
Dim b As Boolean
Dim J As Byte
Dim Col As Collection
Dim F As Date
Dim CambiaArticuloLineasFactura As Boolean

    'Primera comprobacion
    'No puede haber ningun albaran "ALT" pendiente de facturar
    
    'Ocubre 2013
    'Se facturan desde tel_cabfactura y los albaranes asociados al numero de telefono, SIN MIRAR fechas
    'Solo comprobare que estan marcados para facturar
    'Ademas comprobaremos que los albaranes tienen el numero correcto de socio/telefono/departamento
    On Error GoTo eHacerFacturacionTelefonia
    
    

    
    cad = ""
    For NumRegElim = 1 To Me.lwT.ListItems.Count
        If lwT.ListItems(NumRegElim).Bold Then
            If lwT.ListItems(NumRegElim).ForeColor = vbRed Then cad = cad & "-" & Me.lwT.ListItems(NumRegElim).Text & " " & lwT.ListItems(NumRegElim).SubItems(1) & vbCrLf
        End If
    Next NumRegElim
    If cad <> "" Then
        cad = "Estos telefonos tienen albaranes pero no estan marcados para facturar: " & vbCrLf & cad
        cad = cad & vbCrLf & "*** ¿Seguro que desea continuar?"
    Else
        'NUEVO oCT 2013
        cad = "factursn=0 AND codtipom"
        cad = DevuelveDesdeBD(conAri, "count(*)", "scaalb", cad, "ALT", "T")
        If cad = "" Then cad = "0"
        I = 0
        If Val(cad) > 0 Then
            cad = "Existen albaranes sin marca de facturar " & vbCrLf
            cad = cad & vbCrLf & "***    ¿Continuar?      ****"
        Else
            cad = ""
        End If
    End If
       
    If cad <> "" Then
        If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    'Si todos los telefonos esta asociados a un telefono/cliente ARIGES
    cad = " tel_cab_factura left join sclientfno on IdTelefono=Telefono"
    cad = DevuelveDesdeBD(conAri, "count(*)", cad, "IdTelefono is null AND fichero", cboFichero(0).List(cboFichero(0).ListIndex), "T")
    If cad = "" Then cad = "0"
    If Val(cad) > 0 Then
        MsgBox "Telefonos sin asignar a clientes ARIGES", vbExclamation
        Exit Sub
    End If
    
    Set miRsAux = New ADODB.Recordset
    
    'Comprobaremos que todos los albaranes que YA estan en telefonia, tienen correctos los clientes/departame
    Label1(5).Caption = "Comprobacion alb telefonia"
    Label1(5).Refresh
    
    cad = "Select * from scaalb where codtipom='ALT' AND factursn=1"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
            
            '1º  Que el numero de telefono lo tengo
           
                'select  concat(codclien,'|',if(coddirec is null,'',coddirec),'|') from sclientfno where IdTelefono ='625780666'
            cad = "concat(codclien,'|',if(coddirec is null,'',coddirec),'|')"
            cad = DevuelveDesdeBD(conAri, cad, "sclientfno", "IdTelefono", miRsAux!referenc, "T")
            If cad = "" Then
                Err.Raise 513, , "No se encuentra referencia: " & DBLet(miRsAux!referenc, "T")
            Else
                'Mismo cliente
                If Val(RecuperaValor(cad, 1)) = miRsAux!codClien Then
                    cad = RecuperaValor(cad, 2)
                   
                    
                    If DBLet(miRsAux!CodDirec, "T") <> cad Then Err.Raise 513, , "Coddirec incorrectas el nº de telefono: " & DBLet(miRsAux!referenc, "T")
                        
                    
                Else
                    Err.Raise 513, , "Distinto cliente: " & DBLet(miRsAux!referenc, "T")
                End If
            End If
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    

    cad = DevuelveDesdeBD(conAri, "fecha", "tel_cab_factura", "fichero", cboFichero(0).List(cboFichero(0).ListIndex), "T")
    If cad = "" Then
        MsgBox "Error obteniendo fecha factura", vbExclamation
        Exit Sub
    End If
    F = CDate(cad)
    
    cad = PonerTrabajadorConectado("")
    If cad = "" Then
        MsgBox "Error obteniendo datos trabajador conectado", vbExclamation
        Exit Sub
    End If
    
    'Veremos si se solapan las facturas. Obtenemos la fecha
    Label1(5).Caption = "Solapara nº factura"
    Label1(5).Refresh
    
    
    'Veamos series y (min) n1factura
    
    
    cad = "select serie,min(numfact) minim ,max(numfact) maxi from tel_cab_factura WHERE fichero=" & DBSet(cboFichero(0).List(cboFichero(0).ListIndex), "T") & " group by 1"
    
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    
    'SERIE|minimo|max]
    While Not miRsAux.EOF
        cad = DevuelveDesdeBD(conAri, "codtipom", "stipom", "letraser", miRsAux!Serie, "T")
        If cad = "" Then Err.Raise 513, "No se encuentra letraser=" & miRsAux!Serie
        
        cad = cad & "|" & miRsAux!Serie & "|"
        cad = cad & DBLet(miRsAux!minim, "N") & "|"
        cad = cad & DBLet(miRsAux!maxi, "N") & "|"
        Col.Add cad
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Col.Count = 0 Then Err.Raise 513, , "Ningun valor devuelto"
    
    
    For J = 1 To Col.Count
            
        CadenaDesdeOtroForm = RecuperaValor(Col.Item(J), 1)
        
        If CadenaDesdeOtroForm = "ALT" Then 'o es FAT o es FAI
        
            'VEremos la fecha de importacion
            NumRegElim = Year(F)
            cad = "fecfactu between '" & NumRegElim & "-01-01' and '" & NumRegElim & "-12-31' "
            cad = cad & " AND numfactu between " & RecuperaValor(Col.Item(J), 3) & " AND " & RecuperaValor(Col.Item(J), 4) & " AND codtipom"
            cad = DevuelveDesdeBD(conAri, "max(fecfactu)", "scafac", cad, CadenaDesdeOtroForm, "T")
            If cad <> "" Then
                If CDate(cad) > F Then Err.Raise 513, , "Fecha facturada mayor que fecha factura telefonia"
            End If
            
            
            'Veamos si se solapan numeros de factura
            cad = "fecfactu between '" & NumRegElim & "-01-01' and '" & NumRegElim & "-12-31' "
            cad = cad & " AND numfactu between " & RecuperaValor(Col.Item(J), 3) & " AND " & RecuperaValor(Col.Item(J), 4) & " AND codtipom"
            cad = DevuelveDesdeBD(conAri, "count(*)", "scafac", cad, CadenaDesdeOtroForm, "T")
            If cad = "" Then cad = "0"
            If Val(cad) > 0 Then Err.Raise 513, , "Se solapan " & cad & " factura(s)"
            
    
    
    
            'Salto de factura. Veremos cual es la ultima fra trasapsada
            cad = "fecfactu between '" & NumRegElim & "-01-01' and '" & NumRegElim & "-12-31' "
            cad = DevuelveDesdeBD(conAri, "max(numfactu)", "scafac", cad, CadenaDesdeOtroForm, "T")
            If cad <> "" Then
                NumRegElim = Val(RecuperaValor(Col.Item(J), 4)) - Val(RecuperaValor(Col.Item(J), 4))
                If NumRegElim > 1 Then Err.Raise 513, , "Salto factura"
            End If
        
    
        Else
            'FARA FAI comprobaremos que no existe NINGUN albaran en scaalb
            'con un numero igaul al de la factura. NO deberia ya que en su momento cogio de scaalb
            
            
            cad = " numalbar between " & RecuperaValor(Col.Item(J), 3) & " AND " & RecuperaValor(Col.Item(J), 4) & " AND codtipom"
            cad = DevuelveDesdeBD(conAri, "count(*)", "scaalb", cad, "ALI", "T")
            If cad = "" Then cad = "0"
            'If Val(Cad) > 0 Then Err.Raise 513, , "Se solapan albaranes internos"
    
    
        End If
    Next J
    
    'Ok , pues adelante
    '
    
    
    Screen.MousePointer = vbDefault
    Label1(5).Caption = ""
    CadenaDesdeOtroForm = ""
    frmListado3.Opcion = 36
    frmListado3.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        
        'Obtenemos la compañia que vamos a facturar
        CambiaArticuloLineasFactura = False
        If vParamAplic.TieneTelefonia2 = 3 Then
            cad = DevuelveDesdeBD(conAri, "distinct(companyia)", "tel_cab_factura", "Fichero", cboFichero(0).List(cboFichero(0).ListIndex), "T")
            If cad = "ORA" Then
                'ORANGE
                cad = DevuelveDesdeBD(conAri, "artiTelefNorORAN", "spara2", "1", "1")
                If cad <> "" Then
                    If cad <> vParamAplic.ArtiTelefonia Then
                        CambiaArticuloLineasFactura = True
                        vParamAplic.ArtiTelefonia = cad
                    End If
                End If
            Else
                If cad = "VOD" Then
                    'VODAFONE
                    cad = DevuelveDesdeBD(conAri, "artiTelefNorVOD", "spara2", "1", "1")
                    If cad <> "" Then
                        If cad <> vParamAplic.ArtiTelefonia Then
                            CambiaArticuloLineasFactura = True
                            vParamAplic.ArtiTelefonia = cad
                        End If
                    End If
                End If
            End If
        End If
            
    'Reestablecemos el articulo de telefonia
    

        
        b = traspasofacturasTelefonia(cboFichero(0).List(cboFichero(0).ListIndex), Label1(5), CInt(CadenaDesdeOtroForm))
        
        
        If CambiaArticuloLineasFactura Then
            'Sea como sea, dejo el articulo de telefonia como estaba
            cad = DevuelveDesdeBD(conAri, "codartictel", "spara1", "1", "1")
            vParamAplic.ArtiTelefonia = cad
        End If
        
        If b Then
            ACtualizarPuntosTelefonia
            Unload Me
        End If
        
    End If
    
    
eHacerFacturacionTelefonia:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set Col = Nothing
    CadenaDesdeOtroForm = ""
End Sub





Private Sub ACtualizarPuntosTelefonia()
    DoEvents
    Label1(5).Caption = "Ajuste puntos"
    Label1(5).Refresh
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    cad = "select Telefono,BaseImponible,base_exenta from tel_cab_factura WHERE fichero= '" & cboFichero(0).List(cboFichero(0).ListIndex) & "' ORDER BY Telefono"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF

            'FALTA### deberiamos parametrizar
            'Los puntos
            ' Email Martin: 15/05/13   - 1 punto por cada 0,20 € de Base Imponible.
            
            
                            
            I = CInt((miRsAux!BaseImponible + DBLet(miRsAux!base_exenta, "N")) / 0.2)
            cad = "UPDATE sclientfno SET puntos = puntos + " & CStr(I)
            cad = cad & " WHERE IdTelefono = " & DBSet(miRsAux!Telefono, "T")
            conn.Execute cad
            miRsAux.MoveNext
    Wend
    miRsAux.Close

End Sub

Private Sub cmdListadoDto_Click()

    If Me.cboFichero(1).ListIndex < 0 Then Exit Sub
    '
    I = 65
    If Opcion = 3 Then I = 69
    If Opcion = 4 Then
        I = 70
        HacerAccionesDelJOinDeRafa
    
    End If
    
    
    If Opcion = 6 Then
        frmListado4.vCadena = cboFichero(1).Text
        frmListado4.Opcion = 8
        frmListado4.Show vbModal
        Exit Sub
    End If
    
    cad = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", CStr(I), "N")
    If cad = "" Then
        MsgBox "Error obtener informe: " & I, vbExclamation
    Else
        
        With frmImprimir
            'Comun
            .OtrosParametros = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
            .NumeroParametros = 1
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 5
            .ConSubInforme = True
            .NombreRPT = cad
            Select Case Opcion
            Case 3
                .Titulo = "Resumen facturacion soporte"
                'frmVisReport.FicheroInforme = "C:\Telefonia" & "\Informes\" & "detalle_facturas_soporte.rpt"
                .FormulaSeleccion = "{tel_cab_factura.Fichero} = '" & cboFichero(1).Text & "'"
            Case 4
                .Titulo = "Factura resumen"
                '.FormulaSeleccion = "{tel_cab_factura.Fichero} = '" & cboFichero(1).Text & "'"
                .FormulaSeleccion = "{tmpcrmcobros.codusu} = " & vUsu.Codigo
        
            Case Else
                'DOS
                .FormulaSeleccion = "{tmp_inf_descuentos.Fichero} = '" & cboFichero(1).Text & "'"
                .Titulo = "Estudio descuentos telefonía"
            
            End Select
            .Show vbModal
        End With
    End If
End Sub

Private Sub HacerAccionesDelJOinDeRafa()

    'RAfa tenia un subreporte con un command y dentro un UNION
    ' Como los commands no se pueden enlazar tenemos que cargar
    'en una tmp
    conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    
    'Cojo el JOIN que habia en el rpt y lo meto aqui
    cad = ""
    cad = cad & "select " & vUsu.Codigo & ",0,0, a.CodCuota as Codigo, a.DescCuota  as Nombre, sum(a.importe), b.PorcentajeOperador as Porc, (sum(a.importe) * b.PorcentajeOperador)/100"
    cad = cad & " , b.Porcentaje, (sum(a.importe) * b.Porcentaje)/100 from tel_lin_factura_cuotas as a,"
    cad = cad & " tel_desc_cuotas as b, tel_cab_factura As C where A.serie = C.serie and a.NumFact = c.Numfact"
    cad = cad & " and a.Ano = c.Ano and a.CodCuota = b.CodCuota and fichero='" & cboFichero(1).Text & "'"
    ''CI0544330498'
    cad = cad & " group by c.Fichero, a.CodCuota UNION "
    cad = cad & " select " & vUsu.Codigo & ",0,0,a.CodTipoTrafico as Codigo, a.DescTipoTrafico as Nombre, sum(a.importe)"
    cad = cad & " , b.PorcentajeOperador as Porc, (sum(a.importe) * b.PorcentajeOperador)/100 "
    cad = cad & " , b.Porcentaje, (sum(a.importe) * b.Porcentaje)/100 from tel_lin_factura_consumos as a,"
    cad = cad & " tel_desc_consumos as b,tel_cab_factura As C where A.serie = C.serie and a.NumFact = c.Numfact"
    cad = cad & " and a.Ano = c.Ano and a.CodTipoTrafico = b.CodTipoTrafico and fichero='" & cboFichero(1).Text & "'"
    cad = cad & " group by c.Fichero, a.CodTipoTrafico"
    
    'Lo metemos en tmp
    cad = "INSERT INTO tmpinformes(codusu,campo1,campo2,nombre1,nombre2,importe1,porcen1,importe2,porcen2,importe3) " & cad
    conn.Execute cad
    
    'Para que solo coja un registro
    conn.Execute "DELETE FROM tmpcrmcobros WHERE codusu = " & vUsu.Codigo
    cad = "INSERT INTO tmpcrmcobros(codusu,secuencial,forpa) VALUES (" & vUsu.Codigo
    cad = cad & ",1,'" & cboFichero(1).Text & "')"
    conn.Execute cad
    
'ORDER BY Porc, Nombre;
End Sub

Private Sub cmdSalir_Click(index As Integer)
    Unload Me
End Sub

Private Sub Command1_Click()

            
    If cboCompanyia2.ItemData(cboCompanyia2.ListIndex) > 4 Then
        MsgBox "Proceso no desarrollado", vbExclamation
        Exit Sub
    End If

    If cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 4 Then
        'COARVAL COARVAL
        HacerCoarval
    Else
        'Resto operadores
        
        Screen.MousePointer = vbHourglass
        Me.Command1.Enabled = False
        HacerImportacion
        Me.Command1.Enabled = True
        Screen.MousePointer = vbDefault

    End If
End Sub






Private Sub HacerImportacion()
    Dim mGen2 As TelGenerador
    Dim resultado As Boolean
    Dim Mens As String
    Dim FicheroOrange As String
    
    
    
    
    Mens = ""
    If Text1.Text = "" Then
        Mens = "Fichero vacio"
    Else
        If Dir(Text1.Text) = "" Then
            Mens = "No existe fichero"
        Else
            'Comprobaremos si el fichero NO ha sido PROCESADO del todo, es decir, metido tb en scafac,slifac...
            'FALTA###
                    
                    
            Mens = Text1.Text
            
            
            If Len(Mens) > 12 Then Mens = Right(Mens, 12)
            

            '1era cosa a tener en cuenta. El fichero no puede estar procesado
            Mens = DevuelveDesdeBD(conAri, "fecha", "tel_fichtraspasados", "Fichero", Mens, "T")
            If Mens <> "" Then Mens = "El fichero se traspaso el dia " & Mens & vbCrLf
                
            
        End If
    End If
        
    
    '-- Controlamos que la fecha de emisión de facturas sea mas o menos correcta
    If Not IsDate(Text2(0)) Then
        Mens = Mens & vbCrLf & "Ha de introducir una fecha de emisión de facturas correcta"
    Else
        If vEmpresa.FechaIni > CDate(Text2(0).Text) Or DateAdd("yyyy", 1, vEmpresa.FechaFin) < CDate(Text2(0).Text) Then _
            Mens = Mens & vbCrLf & "Fuera de ejercicios contables"
            
    End If
    
    If Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 3 Then
        If Text1.Text <> "" Then
            If InStr(1, Text1.Text, ".") > 0 Then Mens = Mens & vbCrLf & "El fichero de VODAFONE no debe llevar extension"
        End If
    End If
            
    
    If Mens <> "" Then
        Mens = "Campos obligados" & vbCrLf & vbCrLf & Mens
        MsgBox Mens, vbExclamation
        Exit Sub
    End If
    
    '***********************************************************************
    '***********************************************************************
    '***********************************************************************
    '
    ' En referencia de las FAT grabaremos el NUMERO de telefono
    ' Una vez este en tal_cab_Factura entonces cuando vayamos a factura
    ' de tel_cab veremos su hay algun ALT nuevo que se ha generado desde
    ' albaranes de telefonia, con lo cual los contadores tienen que estar
    ' separados (FAT y ALT) ya que el proceso de facturacion desde el num
    ' de factura lo mete en albaranes y de ahi a scafac, entoces si
    ' hubiera algun albaran ALT asociado a ese numereo de telefono(refern)
    ' entonces lo tendria que meter como facturacion colectiva e irian
    ' los dos juntos(o 3 o cuatro...)
    '
    '
    '***********************************************************************
    '***********************************************************************
    '***********************************************************************
    'Solo dejamos UN fichero en 'proceso' de facturacion
    resultado = False
    Mens = Text1.Text
    If Len(Text1) > 12 Then Mens = Right(Mens, 12)
    
    
    
    'Para el proceso de ORANGE, como el nombre del fichero puede variar "demasiado"
    'preprocesaremos el fichero para obtener el numero de factura que
    'esta en la segunda linea. Si el numero de factura ya ha sido procesado entonces
    'daremos el mensaje
    'número de factura:;A10020017274-0813  --> A1002017274
    If Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 2 Then
        'ES ORANGE
        Set mGen2 = New TelGenerador
        FicheroOrange = mGen2.DevuelveNombreFicheroOrange(Text1.Text)
        Set mGen2 = Nothing
        If FicheroOrange = "" Then
            MsgBox "Imposible localizar datos factura en fichero Orange", vbExclamation
            Exit Sub
        End If
        
        
        
        'EN ORANGE tenemos que comprobar que el fichero NO ha siado traspasado
        Mens = DevuelveDesdeBD(conAri, "fecha", "tel_fichtraspasados", "Fichero", FicheroOrange, "T")
        If Mens <> "" Then
            MsgBox "El fichero se traspaso el dia " & Mens & vbCrLf, vbExclamation
            Exit Sub
        End If
        
        
        'Para la utilizacion posterior
        Mens = FicheroOrange
        
    End If
    
    
    
    
    
    
    
    'Mens = "select distinct(Fichero) from tel_cab_factura where not Fichero in (select Fichero from tel_fichtraspasados)"
    cad = " not Fichero in (select Fichero from tel_fichtraspasados) AND 1"
    cad = DevuelveDesdeBD(conAri, "distinct(Fichero)", "tel_cab_factura", cad, "1")
    If cad <> "" Then
        'Si el fichero que falta NO es el que estamos intentando pasar
        If cad <> Mens Then
            MsgBox "Falta procesar el archivo: " & cad, vbExclamation
            
            Exit Sub
            
        Else
            cad = "Volver a cargar los datos del fichero: " & Mens & "?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            resultado = True
        End If
    End If
    
    
    
    
    
    
    
    
    'DAVID###
    'EL proceso se divide en:
    '   -proceso antiguo(todo lo que hacia Rafa)
    '   -y luego metemos en la slialb para pasarles el proceso de facturacion NORMAL de
    ' ariges
    'Un fichero puede ser importado muchas veces. Siempre borra los datos etc.
    'Al final, cuando pulse el boton de llevar a scafac, una vez haga esto, YA no puede volver a importar
    ' el fichero
    '-- Por si nos pasan ruta completa modificamos el nombre de fichero
    
    cad = String(40, "*") & vbCrLf
    cad = cad & cad & vbCrLf & vbCrLf
    Mens = cad & "Va a importar el fichero de telefonía:"
    Mens = Mens & vbCrLf & vbCrLf & "Compañia: " & Me.cboCompanyia2.Text
    Mens = Mens & vbCrLf & vbCrLf & "FECHA: " & Text2(0).Text & vbCrLf & vbCrLf & cad
    
    
    'En resultado tenemos si ya ha hecho la pregunta de procesar, para que no la vuelva a hacer
    If Not resultado Then
        If MsgBox(Mens, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    resultado = False
    
    
    Set mGen2 = New TelGenerador
    If mGen2.LeerParametrosFacturacionTelefonica(cboCompanyia2.ItemData(cboCompanyia2.ListIndex)) Then
    
        Screen.MousePointer = vbHourglass
         resultado = False
        If Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 1 Then
            resultado = mGen2.cargarBaseDatosMOVISTAR(Text1, lblInf)
            'resultado = True
        ElseIf Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) = 2 Then
           
            
            'Le paso el fichero fisico, el nombre estara en: FicheroOrange
            resultado = mGen2.cargarBaseDatosOrange(Text1.Text, lblInf)
            
        Else
            'VODAFONE
            resultado = mGen2.cargarBaseDatosVODAFONE(Text1.Text, lblInf)
        End If
        lblInf.Caption = ""
        Screen.MousePointer = vbDefault
        If Not resultado Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        
        
        'LLamadas entre coooperativistas
        lblInf.Caption = "Acciones cooperativa"
        lblInf.Refresh
        
        
        'Si hay conceptos o cuotas nuevas las mete en tmpinformes para listarlas luego
        'tmpinformes(codusu,codigo1,campo1,nombre1,nombre2)
        conn.Execute "DELETE from tmpinformes WHERE codusu = " & vUsu.Codigo
        
        If vParamAplic.TieneTelefonia2 = 3 Then
                mGen2.RecalcularImporteLlamadasCoperativa Me.lblInf, cboCompanyia2.ItemData(cboCompanyia2.ListIndex)
        End If
        
        
         'VEmos cuotas
         resultado = mGen2.AjusteCuotasNuevas2(cboCompanyia2.ItemData(cboCompanyia2.ListIndex), Right(Text1, 12), Me.lblInf)
         
         'refacturamos
         If resultado Then mGen2.ComprobarConceptosFacturacion
        
                
        
        
        If Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex) <> 2 Then FicheroOrange = Text1.Text 'Para movistar dejo el nombre del fichero
        
        
        
        
        
        
        
        
        CadenaDesdeOtroForm = ""
        Screen.MousePointer = vbDefault
        If Not resultado Then Exit Sub
        
        Screen.MousePointer = vbDefault
        DoEvents
        
        resultado = mGen2.EmitirFacturas_(FicheroOrange, Text2(0), Me.lblInf, CByte(Me.cboCompanyia2.ItemData(cboCompanyia2.ListIndex)))
        'TRUE=Error
        If resultado Then
            Mens = "Se ha producido incidencias durante el proceso de generación. " & _
                    "Estas incidencias se han guardado en el fichero " & App.Path & "\emitefac.log" & vbCrLf & _
                    "¿Desea ver el contenido de este fichero?"
            If MsgBox(Mens, vbYesNo + vbQuestion) = vbYes Then
                Shell "notepad " & App.Path & "\emitefac.log", vbMaximizedFocus
            End If
        End If

        mGen2.calculaInformeDescuentos Mens
    
    
    
    End If
    
    
    
    Set mGen2 = Nothing
    lblInf.Caption = ""
     
    
    If vParamAplic.TieneTelefonia2 = 3 Then NuevasCuotasConceptos
    
    
    CadenaDesdeOtroForm = "SI"
    Unload Me
  
    
    
End Sub



Private Sub NuevasCuotasConceptos()
    
    On Error GoTo eNuevasCuotasConceptos
    cad = "Select * from tmpinformes where codusu=" & vUsu.Codigo & " ORDER BY campo2,nombre1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not miRsAux.EOF Then
        I = FreeFile
        cad = App.Path & "\NC" & Format(Now, "yymmddhhnn") & ".txt"
        CadenaDesdeOtroForm = cad
        Open cad For Output As #I
        While Not miRsAux.EOF
            
            If miRsAux!Codigo1 = 2 Then
                cad = "Varios "
            Else
                cad = IIf(miRsAux!campo1 = 0, "Cuota ", "Conce ")
            End If
            cad = "    " & cad & miRsAux!nombre1 & " :  " & miRsAux!nombre2 & vbCrLf
            Print #I, cad
            miRsAux.MoveNext
        Wend
        Close #I
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If I > 0 Then
        LanzaVisorMimeDocumento Me.hwnd, CadenaDesdeOtroForm
        CadenaDesdeOtroForm = ""
    End If
    
eNuevasCuotasConceptos:
   If Err.Number <> 0 Then MsgBox "Error leyendo cuotas nuevas creadas", vbExclamation
End Sub



Private Sub Form_Load()

    
    
    Me.FrameImportacion.visible = False
    Me.FrameVerdatos.visible = False
    FrameCatadau.visible = False
    I = Opcion
    Select Case Opcion
    Case 0
        lblI.Caption = ""
        Me.lwT.ColumnHeaders(3).Width = IIf(vParamAplic.TelefoniaVtaPlazos, 600, 0)
        PonerFrameVisible Me.FrameVerdatos
        CargaCombo Me.cboFichero(0), True
    Case 1
        PonerFrameVisible Me.FrameImportacion
        
        
        CargaComboCompanyia
        
        
    Case 2, 3, 4, 6
        I = 2 'cancelar(2)
        If Opcion = 3 Then
            Label1(6).Caption = "Facturación por soporte"
        ElseIf Opcion = 4 Then
            Label1(6).Caption = "Resumen por soporte"
            
        ElseIf Opcion = 6 Then
            Label1(6).Caption = "Datos importados fichero"
            
        Else
            '2
            Label1(6).Caption = "Estudio descuentos telefonía"
        End If
        
        PonerFrameVisible Me.FrameDtosTelefonia
        CargaCombo Me.cboFichero(1), False
        
'    Case 5
'        PonerFrameVisible FrameCatadau
'        lblInf.Caption = ""
    End Select
    
    cmdSalir(I).Cancel = True
    lblInf.Caption = ""
    
    Label1(5).Caption = ""
    Screen.MousePointer = vbDefault

End Sub

Private Sub PonerFrameVisible(ByRef Fr As Frame)
    Fr.Left = 120
    Fr.Top = 0
    Me.Height = Fr.Height + 510
    Me.Width = Fr.Width + 360
    Fr.visible = True
End Sub

Private Sub CargarListView(Fich As String)
Dim IT As ListItem
Dim Rc As ADODB.Recordset
Dim ImpoAux As Currency

    Set miRsAux = New ADODB.Recordset
    Set Rc = New ADODB.Recordset
   
    IVA_standard = -1
    If vParamAplic.TelefoniaVtaPlazos Then
    
        cad = "Select IdTelefono,PlazosMeses,ArtPlazos,ImportePlazo from sclientfno where PlazosMeses > 0 "
        Rc.Open cad, conn, adOpenKeyset, adCmdText
        If Not Rc.EOF Then
            cad = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", DBLet(Rc!artplazos, "T"), "T")
            If cad <> "" Then
                cad = DevuelveDesdeBD(conConta, "(porceiva+coalesce(porcerec,0))", "tiposiva", "codigiva", cad, "N")
                If cad <> "" Then IVA_standard = CCur(cad)
            End If
        End If
    Else
        If vParamAplic.ArtiTelefonia <> "" Then
            IVA_standard = 0
           cad = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", DBLet(vParamAplic.ArtiTelefonia, "T"), "T")
            If cad <> "" Then
                cad = DevuelveDesdeBD(conConta, "(porceiva+coalesce(porcerec,0))", "tiposiva", "codigiva", cad, "N")
                If cad <> "" Then IVA_standard = CCur(cad)
            End If
        End If
    End If
    
    cad = "select telefono,apellido1, apellido2,nombre,"
    cad = cad & " BaseImponible,Cuota,total,Serie ,Ano ,NumFact"
    cad = cad & "  from tel_cab_factura where fichero='" & Fich & "' order by telefono"
   

    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
    lwT.ListItems.Clear
    
    If Me.chkMostrarBase.Value = 1 Then
        Me.lwT.ColumnHeaders(4).Text = "B.Imp."
    Else
        Me.lwT.ColumnHeaders(4).Text = "Total"
    End If
    While Not miRsAux.EOF
    
        'If miRsAux!Telefono = "687751251" Then St op
    
    
        cad = ""
        If Not IsNull(miRsAux!apellido1) Then cad = miRsAux!apellido1
        If Not IsNull(miRsAux!apellido2) Then cad = Trim(cad & " " & miRsAux!apellido2)
        If Not IsNull(miRsAux!Nombre) Then
            If cad <> "" Then cad = cad & ","
            cad = Trim(cad & " " & miRsAux!Nombre)
        End If
   
        Set IT = lwT.ListItems.Add()
        IT.Text = miRsAux!Telefono
        IT.SubItems(1) = cad
        
        If Me.chkMostrarBase.Value = 1 Then
            IT.SubItems(3) = Format(miRsAux!BaseImponible, "#,##0.00")
        Else
            IT.SubItems(3) = Format(miRsAux!total, "#,##0.00")
        End If
        IT.SubItems(4) = Format(miRsAux!total * 100, "0000000")
        
        IT.SubItems(2) = " "
        If vParamAplic.TelefoniaVtaPlazos Then
            
            cad = "idtelefono='" & miRsAux!Telefono & "'"
            Rc.Find cad, , adSearchForward, 1
            If Not Rc.EOF Then
                IT.ListSubItems(3).ForeColor = vbBlue
                IT.ListSubItems(3).Bold = True
                IT.SubItems(2) = "S"
                
                                
                ImpoAux = 0
                If Me.chkMostrarBase.Value = 0 Then ImpoAux = IVA_standard
                 
                ImpoAux = Round2(DBLet(Rc!ImportePlazo, "N") * ((100 + ImpoAux) / 100), 2)
                
                ImpoAux = ImporteFormateado(IT.SubItems(3)) + ImpoAux
                                        
                IT.SubItems(3) = Format(ImpoAux, FormatoImporte)
                IT.SubItems(4) = Format(ImpoAux * 100, "0000000")
                           
                                
                                
                
            End If
        End If
        
        'Para el WHERE
        IT.Tag = "Serie = '" & miRsAux!Serie & "' AND Ano =" & miRsAux!Ano & " AND NumFact =" & miRsAux!NumFact
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    ComprobarAlbaranesPendientes
    
    
    'Por si se queda la facturacion a medias
    'if False Then
    '   For i = Me.lwT.ListItems.Count To 1 Step -1
    '        If InStr(1, "617716340v629378165v636153242v646603031v661078672v674845196v", lwT.ListItems(i).Text) = 0 Then
    '            lwT.ListItems.Remove i
    '        End If
    '    Next
    'End If
    
    
    lblI.Caption = "Telefonos a facturar: " & lwT.ListItems.Count
    
    
    
     Set miRsAux = Nothing
    Set Rc = Nothing
End Sub

Private Sub ComprobarAlbaranesPendientes()
Dim Impaux As Currency

    On Error GoTo eComprobarAlbaranesPendientes
    
    
    cad = "select scaalb.numalbar,referenc,codclien,nomclien,factursn,sum(importel) base from scaalb left join slialb "
    cad = cad & " on scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar"
    cad = cad & " Where scaalb.codtipom='ALT' group by scaalb.numalbar"
    
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not miRsAux.EOF        'QUe hago....
    
        'If miRsAux!referenc = "687751251" Then St op
    
    
        'Buscamos por el LW e telefono
        For I = 1 To Me.lwT.ListItems.Count
            If Me.lwT.ListItems(I).Text = miRsAux!referenc Then
                'Este es el numero de telefono
                    
                Exit For
            End If
        Next
        
        
        'Si NO lo ha encotrado, lo añado a CAD
        If I > lwT.ListItems.Count Then
            cad = cad & vbCrLf & miRsAux!codClien & " " & miRsAux!NomClien & " -> " & miRsAux!referenc
        Else
            Me.lwT.ListItems(I).Bold = True
            If miRsAux!factursn = 0 Then
                Me.lwT.ListItems(I).ForeColor = vbRed
            Else
                Me.lwT.ListItems(I).ForeColor = vbBlue
                'El total
                
            
                Impaux = 0
                If Me.chkMostrarBase.Value = 0 Then Impaux = IVA_standard
                 
                Impaux = Round2(DBLet(miRsAux!Base, "N") * ((100 + Impaux) / 100), 2)
                
                Impaux = ImporteFormateado(lwT.ListItems(I).SubItems(3)) + Impaux
                                        
                lwT.ListItems(I).SubItems(3) = Format(Impaux, FormatoImporte)
                lwT.ListItems(I).SubItems(4) = Format(Impaux * 100, "0000000")
            
                lwT.ListItems(I).ListSubItems(4).Bold = True
                lwT.ListItems(I).ListSubItems(4).ForeColor = vbBlue
            End If
        End If
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    If cad <> "" Then
        cad = "Albaranes que no se facturarán: " & vbCrLf & cad
        MsgBox cad, vbExclamation
    End If
        
    Exit Sub
eComprobarAlbaranesPendientes:
    MsgBox "Avise soporte tecnico. Albaran pdte   : " & miRsAux!Numalbar & "-" & I & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub imgBuscar_Click(index As Integer)
    cmmDia.ShowOpen
    If index = 0 Then
        Text1.Text = cmmDia.FileName
    Else
       ' Me.txtCatadau(0).Text = cmmDia.FileName
    End If
End Sub

Private Sub ListView1_DblClick()

End Sub




Private Sub CargaCombo(ByRef CBO As ComboBox, FaltaProcesar As Boolean)

    Set miRsAux = New ADODB.Recordset
    
    CBO.Clear
    cad = "Select distinct(fichero) from tel_cab_factura"
    If FaltaProcesar Then
        cad = cad & " WHERE not fichero in (select fichero from tel_FichTraspasados) ORDER BY fecha desc"
    Else
        cad = cad & " WHERE fecha > " & DBSet(DateAdd("yyyy", -4, Now), "F") & " ORDER BY fecha desc"
    End If
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not miRsAux.EOF
        cad = cad & "1"
        CBO.AddItem miRsAux!Fichero
        miRsAux.MoveNext
    Wend
    miRsAux.Close
 
    
    
End Sub


Private Sub Label1_Click(index As Integer)
    If vUsu.Login = "root" Then
            cmdAriadna.visible = True
    End If
End Sub

Private Sub lwT_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    I = ColumnHeader.index - 1
    If ColumnHeader.index = 4 Then I = 4
    
    If I = lwT.SortKey Then
        If lwT.SortOrder = lvwAscending Then
            lwT.SortOrder = lvwDescending
        Else
            lwT.SortOrder = lvwAscending
        End If
    Else
        If ColumnHeader.index = 4 Then
            lwT.SortKey = 4
        Else
            lwT.SortKey = ColumnHeader.index - 1
        End If
        lwT.SortOrder = lvwAscending
    End If
End Sub

Private Sub lwT_DblClick()
    If lwT.ListItems.Count = 0 Then Exit Sub
    If lwT.SelectedItem Is Nothing Then Exit Sub
    frmTelefonoVerFra.TieneAlbaranes = lwT.SelectedItem.Bold
    frmTelefonoVerFra.Where2 = cboFichero(0).Text & "|" & lwT.SelectedItem.Tag & "|"
    frmTelefonoVerFra.Show vbModal
End Sub

Private Sub Text2_GotFocus(index As Integer)
    ConseguirFoco Text2(index), 3
End Sub

Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub Text2_LostFocus(index As Integer)
Dim T As String
Dim BorrarCampo As Boolean
    Text2(index).Text = Trim(Text2(index).Text)
    BorrarCampo = False
    If Text2(index).Text <> "" Then
        T = Text2(index).Text
        If EsFechaOK(T) Then
            If index = 0 Then
                If CDate(T) < vEmpresa.FechaIni Or CDate(T) > DateAdd("yyyy", 1, vEmpresa.FechaFin) Then
                    MsgBox "Fechas fuera de ejercicio", vbExclamation
                    BorrarCampo = True
                End If
            End If
            Text2(index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & Text2(index).Text, vbExclamation
            BorrarCampo = True
        End If
    End If
    If BorrarCampo Then
        Text2(index).Text = ""
        PonerFoco Text2(index)
    End If
End Sub




Private Function GenerarImportacionCatadau(ByVal DigitoCoarval As Integer) As Boolean
Dim Campos() As String
Dim NF As Integer
Dim fin As Boolean
Dim cControlFra As CControlFacturaContab
    
    On Error GoTo eGenerarImportacionCatadau
    
    GenerarImportacionCatadau = False
    
    'tmpinformes codigo1 campo1 campo2  nombre1 fecha1 importe1 importe2 importe3
    conn.Execute "Delete from tmpinformes  WHERE codusu = " & vUsu.Codigo
    
    
    NF = FreeFile
    'Open Me.txtCatadau(0).Text For Input As #NF
    Open Text1.Text For Input As #NF
    'En toeria tiene que haber datos
    cad = ""
    NumRegElim = -1 'Indciador de situacion de proceso
    Do
        
        Line Input #NF, cad
        cad = Trim(cad)

        If Len(cad) >= 5 Then
            If Mid(cad, 1, 5) = ";;;;;" Then cad = ""
        End If

        If cad <> "" Then
            If NumRegElim >= 0 Then 'la primera NO vale y en la primera es -1
                If Mid(cad, 1, 1) = ";" Then MsgBox "Empieza con ;", vbCritical
                Campos = Split(cad, ";")
                NumRegElim = 0  'De momento BIEN
                cad = cad & vbCrLf & vbCrLf
                'Comprobaciones
                If UBound(Campos) < 11 Then
                    cad = cad & "Numero columnas incorrecto. Debian haber 11 columnas"
                    NumRegElim = 1
                Else
                    'OK. Columnas correctas
                    If Not IsNumeric(Campos(5)) Then
                        cad = cad & "Codigo socio incorrecto " & vbCrLf
                        NumRegElim = 1
                    Else
                        'Es el codigo de socio.
                        'Pos si acaso sa ha vuelto loco el de las factuas y lo envia "decimal"
                        Campos(5) = Replace(Campos(5), ",00", "")
                        Campos(5) = Replace(Campos(5), ".", "")
                    End If
                    Campos(0) = Campos(6)
                    If Len(Campos(0)) < 6 Then
                        cad = cad & "Longitud fra incorrecta " & vbCrLf
                        NumRegElim = 1
                    Else
                        If Not IsNumeric(Right(Campos(0), 6)) Then
                            cad = cad & "Numero fra incorrecta " & vbCrLf
                            NumRegElim = 1
                        End If
                    End If
                    
                    For I = 9 To 11
                        If Not IsNumeric(Campos(I)) Then
                            cad = cad & "Importes incorrectos " & Campos(I) & vbCrLf
                            NumRegElim = 1
                        End If
                    Next I
                    'Fecha factura
                    If Not IsDate(Campos(8)) Then
                        cad = cad & "Fecha incorrecta " & vbCrLf
                        NumRegElim = 1
                    End If
                    
                    
                    'SI llega a aqui, y ha ido bien, INSERTARA en tmp
                    If NumRegElim = 0 Then
                        '
                        
                        'Cad = Me.txtCatadau(1).Text & Right(Campos(6), 6)
                        cad = DigitoCoarval & Right(Campos(6), 6)
                        
                        'codusu codigo1, campo1, campo2,  nombre1,
                        cad = vUsu.Codigo & "," & cad & "," & Campos(5) & "," & DBSet(Campos(4), "T")
                        ' fecha1, importe1, importe2, importe3
                        cad = cad & "," & DBSet(Campos(8), "F") & "," & DBSet(Campos(9), "N")
                        cad = cad & "," & DBSet(Campos(10), "N") & "," & DBSet(Campos(11), "N") & ")"
                        
                        cad = "INSERT INTO tmpinformes(codusu,codigo1, campo1,   nombre1, fecha1, importe1, importe2, importe3) VALUES (" & cad
                        conn.Execute cad
                    Else
                        MsgBox cad, vbExclamation
                        fin = True
                    End If
                End If 'numcols
            Else
                NumRegElim = 0  'para que empieze
                cad = ""
            End If
        
            
        Else
            cad = "OK"
        End If
        If NumRegElim = 1 Then
            MsgBox cad, vbExclamation
            fin = True
        End If
        If EOF(NF) Then fin = True
    Loop Until fin
    Close #NF
    
    If cad = "" Then
        MsgBox "NUmero registros incorrecto", vbExclamation
        NumRegElim = 1
    End If
    
    'OK. Ahora un par de comprobaciones mas
    If NumRegElim > 0 Then Exit Function
    
    
    
    
    'De momento va bien. Varias comporbaciones
    'Primer asunto. Codclien=0 NO lo procesamos. Serían internas
    cad = "DELETE from tmpinformes where codusu=" & vUsu.Codigo & " AND campo1=0"
    conn.Execute cad
    
    
    '1ª comprobacion
    cad = "Select distinct(fecha1) from tmpinformes where codusu =" & vUsu.Codigo
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        If I = 0 Then
            cad = miRsAux!fecha1
            Campos(0) = "01/01/" & Year(CDate(cad))
            Campos(1) = "31/12/" & Year(CDate(cad))
        End If
        I = I + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If I <> 1 Then
        MsgBox "No es fecha única", vbExclamation
        Exit Function
    End If
    
    Set cControlFra = New CControlFacturaContab
    
    cad = cControlFra.FechaCorrectaContabilizazion(ConnConta, CDate(cad))
    If cad <> "" Then
        MsgBox cad, vbExclamation
        NumRegElim = 1
    End If
    Set cControlFra = Nothing
    
    
    
    
    
    If NumRegElim = 1 Then Exit Function
    
    
    
    
    'UN par de comprobaciones mas
    cad = "select codusu,max(codigo1) elmaximo ,min(codigo1) elminimo from tmpinformes where codusu=" & vUsu.Codigo & " group by 1"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    'Comprobemos si NO existe el numero factura
    cad = " numfactu >= " & miRsAux!elminimo & " AND numfactu<=" & miRsAux!elmaximo
    cad = cad & " AND fecfactu>=" & DBSet(Campos(0), "F") & " AND fecfactu<=" & DBSet(Campos(1), "F")
    miRsAux.Close
    
    
    cad = "Select count(*) FROM scafac where codtipom='FAT' AND " & cad
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "0"
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then cad = miRsAux.Fields(0)
    End If
    If Val(cad) > 0 Then
        cad = "Se van a solapar " & cad & " registro(s) que se solaparán numeros de factura"
        cad = cad & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(cad, vbQuestion + vbYesNoCancel + vbDefaultButton3) <> vbYes Then Exit Function
        
    End If
        
    
    
    GenerarImportacionCatadau = True
    
    
    
    Exit Function
eGenerarImportacionCatadau:
    MuestraError Err.Number, Err.Description
    
End Function

Private Sub txtCatadau_KeyPress(index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub



Private Sub InsertarFacturasTelefonoCoarval()
    CadenaDesdeOtroForm = ""
    frmListado3.Opcion = 36
    frmListado3.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        traspasofacturasTelefoniaCOARVAL Me.lblInf, CInt(CadenaDesdeOtroForm)
    End If
End Sub













Private Sub CargaComboCompanyia()

'    cboCompanyia2.Clear
'    cboCompanyia2.AddItem "Movistar"
'    cboCompanyia2.ItemData(cboCompanyia2.NewIndex) = 1
'
'    cboCompanyia2.AddItem "Orange"
'    cboCompanyia2.ItemData(cboCompanyia2.NewIndex) = 2
'

    CargarCombo_Tabla cboCompanyia2, "stfnoOperador", "codoperador", "nombre"
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        cboCompanyia2.ListIndex = 3
    ElseIf vParamAplic.NumeroInstalacion <> vbAlzira Then
        cboCompanyia2.ListIndex = 2
    Else
        cboCompanyia2.ListIndex = 0
    End If
    
    
End Sub
