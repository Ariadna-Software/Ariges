VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelefonoVerFra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver datos prefactura"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   9
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   8
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   3240
      Width           =   855
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
      Height          =   375
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   2280
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   3240
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Height          =   375
      Index           =   1
      Left            =   3600
      Picture         =   "frmTelefonoVerFra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Suma datos seleccionados"
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdExport 
      Height          =   375
      Index           =   0
      Left            =   4080
      Picture         =   "frmTelefonoVerFra.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exportar csv"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   10440
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   3240
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2445
      Left            =   5640
      TabIndex        =   7
      Top             =   330
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4313
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Concepto"
         Object.Width           =   6853
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   3240
      Width           =   855
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
      Height          =   375
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   6135
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10821
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Destino"
         Object.Width           =   2141
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha-Hora"
         Object.Width           =   2170
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Codigo"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Descripcion"
         Object.Width           =   6242
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Duracion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unidad"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Importe"
         Object.Width           =   2141
      EndProperty
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   120
      Top             =   2910
      Width           =   9975
   End
   Begin VB.Label Label1 
      Caption         =   "Vta plazos"
      Height          =   255
      Index           =   12
      Left            =   3720
      TabIndex        =   27
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Albaranes"
      Height          =   255
      Index           =   11
      Left            =   2520
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Vta plazos"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Exento"
      Height          =   255
      Index           =   9
      Left            =   1320
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   1650
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Resumen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   13
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "TOTAL"
      Height          =   255
      Index           =   5
      Left            =   8400
      TabIndex        =   12
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "IVA"
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Imponible"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Teléfono"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmTelefonoVerFra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Where2 As String   'fichero|where|
Public TieneAlbaranes As Boolean


Private UnaVez As Boolean
Dim Cad As String



Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub cmdExport_Click(Index As Integer)
Dim SumaCantidad As Currency
Dim SumaImporte As Currency
Dim Seleccionados As Integer
    'Sumatorio selecccion
    If Index = 0 Then
        'Exportar
        ExportaCSV
    Else
        Seleccionados = 0
        For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Selected Then Seleccionados = Seleccionados + 1
        Next
        If Seleccionados = 0 Then
            MsgBox "Seleccione algún elemento", vbExclamation
            Exit Sub
        End If
        
        Cad = "|"
        For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Selected Then
                If InStr(1, Cad, "|" & Me.ListView2.ListItems(NumRegElim).SubItems(2) & "|") = 0 Then Cad = Cad & ListView2.ListItems(NumRegElim).SubItems(2) & "|"
            End If
        Next
        'Ya tengo los distintos conceptos
        'AAhora busco
        Cad = Mid(Cad, 2) 'quito el primer pipe
        CadenaDesdeOtroForm = ""
        Do
            davidNumalbar = InStr(1, Cad, "|")
            If davidNumalbar = 0 Then
                Cad = ""
            Else
                pPdfRpt = Mid(Cad, 1, davidNumalbar - 1)
                Cad = Mid(Cad, davidNumalbar + 1)
                pImprimeDirecto = True
                SumaCantidad = 0
                SumaImporte = 0
                davidNumalbar = 0
                For NumRegElim = 1 To ListView2.ListItems.Count
                    If ListView2.ListItems(NumRegElim).Selected Then
                        If Me.ListView2.ListItems(NumRegElim).SubItems(2) = pPdfRpt Then
                            If pImprimeDirecto Then
                                CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & Me.ListView2.ListItems(NumRegElim).SubItems(3)
                                pImprimeDirecto = False
                            End If
                            davidNumalbar = davidNumalbar + 1
                            SumaCantidad = SumaCantidad + CCur(Me.ListView2.ListItems(NumRegElim).SubItems(4))
                            If Me.ListView2.ListItems(NumRegElim).SubItems(6) <> "" Then SumaImporte = SumaImporte + CCur(Me.ListView2.ListItems(NumRegElim).SubItems(6))
                        End If
                    End If
                Next
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & Format(davidNumalbar, "000") & "      Cantidad:    " & Format(SumaCantidad, FormatoCantidad)
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "         Importe:     " & Format(SumaImporte, FormatoPrecio) & vbCrLf
            End If
        Loop Until Cad = ""
        pPdfRpt = "": pImprimeDirecto = False: davidNumalbar = 0
        CadenaDesdeOtroForm = "Resumen datos seleccionados (" & Seleccionados & ") " & vbCrLf & CadenaDesdeOtroForm
        MsgBox CadenaDesdeOtroForm, vbInformation
        CadenaDesdeOtroForm = ""
    End If
End Sub



Private Sub Form_Activate()
Dim Rs As ADODB.Recordset
    
    If UnaVez Then
        UnaVez = False
        CargarDatos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    UnaVez = True
    Me.Icon = frmPpal.Icon
    ListView2.Tag = 0
    
    Label1(12).visible = vParamAplic.TelefoniaVtaPlazos
    Text1(9).visible = vParamAplic.TelefoniaVtaPlazos
    Label1(10).visible = vParamAplic.TelefoniaVtaPlazos
    Text1(7).visible = vParamAplic.TelefoniaVtaPlazos

    limpiar Me
    Text1(1).Text = "          L e y e n d o    B.D."
    'Limpiar
End Sub

Private Sub CargarDatos()
Dim Importe1 As Currency
Dim IVA As Currency
Dim Aux As Currency
Dim IT As ListItem

    

    Set miRsAux = New ADODB.Recordset
    
    'Cabecera
    Cad = "Select * from tel_cab_factura WHERE  " & RecuperaValor(Where2, 2)
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUESER EOF
    Text1(0).Text = miRsAux!Telefono
         Cad = ""
         If Not IsNull(miRsAux!apellido1) Then Cad = miRsAux!apellido1
         If Not IsNull(miRsAux!apellido2) Then Cad = Trim(Cad & " " & miRsAux!apellido2)
         If Not IsNull(miRsAux!Nombre) Then
             If Cad <> "" Then Cad = Cad & ","
             Cad = Trim(Cad & " " & Trim(miRsAux!Nombre))
         End If
    Text1(1).Tag = Cad
    
    Text1(2).Text = miRsAux!Fecha
    Text1(3).Text = Format(miRsAux!BaseImponible, "#,##0.00") ' Cuota Total
    Text1(4).Text = Format(miRsAux!Cuota, "#,##0.00") ' Cuota Total
    Text1(5).Text = Format(miRsAux!total, "#,##0.00") ' Cuota Total
    Text1(6).Text = ""
    If DBLet(miRsAux!base_exenta, "N") <> 0 Then Text1(6).Text = Format(miRsAux!base_exenta, "#,##0.00") ' Cuota Total
    
    If vParamAplic.TieneTelefonia2 = 3 Then
        Cad = FormatoPrecio
    Else
        Cad = FormatoImporte
    End If
    CargaLwTelefonia Me.ListView1, miRsAux!Serie, miRsAux!Ano, miRsAux!NumFact, Cad
    
    miRsAux.Close
    
    Importe1 = 0
    IVA = 0
    
    Label1(8).Caption = ""
    If TieneAlbaranes Then
    
        Cad = "Select nomclien,scaalb.numalbar,slialb.codartic,slialb.nomartic,importel,codigiva from scaalb left join slialb on scaalb.codtipom=slialb.codtipom"
        Cad = Cad & " AND scaalb.numalbar = slialb.numalbar left join sartic on slialb.codartic = sartic.codartic  "
        
        Cad = Cad & " WHERE referenc = '" & Text1(0).Text & "' AND factursn=1"
        
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            MsgBox "Albaranes NO encotrados", vbExclamation
        Else
            If miRsAux!NomClien <> Text1(1).Tag Then
                '
                Text1(1).ForeColor = vbRed
                Label1(8).Caption = miRsAux!NomClien
            Else
                Text1(1).ForeColor = vbBlack
            End If
            While Not miRsAux.EOF
                Cad = "** " & Format(miRsAux!NumAlbar, "0000") & " - " & miRsAux!NomArtic & "(" & miRsAux!codArtic & ")"
                ListView1.ListItems.Add , , Cad
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = Format(miRsAux!ImporteL, "#,##0.00")
                ListView1.ListItems(ListView1.ListItems.Count).ForeColor = vbBlue
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems(1).ForeColor = vbBlue
                
                Cad = DevuelveDesdeBD(conConta, "(porceiva+coalesce(porcerec,0))", "tiposiva", "codigiva", DBLet(miRsAux!Codigiva, "N"))
                If Cad = "" Then Cad = "0"
                Aux = (miRsAux!ImporteL * CCur(Cad)) / 100
                IVA = IVA + Round2(Aux, 2)
                Importe1 = Importe1 + miRsAux!ImporteL
                miRsAux.MoveNext
            Wend
        End If
        miRsAux.Close
        
        
        Text1(8).Text = Format(Importe1, FormatoImporte)  'Solo la base
        
    End If
    
    
    DoEvents
    Where2 = RecuperaValor(Where2, 1)
    
    DetalleLlamada 1
    
    
    
    'Si lleva venta plazos
    If vParamAplic.TelefoniaVtaPlazos Then
        Cad = "select ArtPlazos,PlazosMeses,ImportePlazo,nomartic,codigiva from sclientfno left join sartic on artplazos=codartic "
        Cad = Cad & " where PlazosMeses>0 AND idtelefono = " & DBSet(Text1(0).Text, "T")
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
        If Not miRsAux.EOF Then
            Cad = miRsAux!artplazos & "   " & DBLet(miRsAux!NomArtic, "T")
            Text1(7).Text = Cad
            Cad = DevuelveDesdeBD(conConta, "(porceiva+coalesce(porcerec,0))", "tiposiva", "codigiva", DBLet(miRsAux!Codigiva, "N"))
            If Cad = "" Then Cad = "0"
            Aux = (miRsAux!ImportePlazo * CCur(Cad)) / 100
        
            IVA = IVA + Aux
            Text1(9).Text = Format(miRsAux!ImportePlazo, FormatoImporte)
        End If
    
    
    
    End If
    
    
    If IVA <> 0 Then
        Aux = ImporteFormateado(Text1(4).Text)
        Aux = Aux + IVA
        Text1(4).Text = Format(Aux, FormatoImporte)
    End If
    
    If IVA <> 0 Or Text1(2).Text <> "" Or Text1(3).Text <> "" Then
        Importe1 = ImporteFormateado(Text1(3).Text)
        Aux = ImporteFormateado(Text1(8).Text)
        Importe1 = Importe1 + Aux
    
        Aux = ImporteFormateado(Text1(9).Text)
        Importe1 = Importe1 + Aux
    
        'NUEVO IVA
        Aux = ImporteFormateado(Text1(4).Text)
        Importe1 = Importe1 + Aux
    
        Text1(5).Text = Format(Importe1, FormatoImporte)
    End If
    Text1(1).Text = Text1(1).Tag
    
    Set miRsAux = Nothing
End Sub



Private Sub DetalleLlamada(Orden As Byte)

    'Detalle llamada
    If ListView2.Tag = Orden Then Exit Sub
    
    Me.ListView2.ListItems.Clear
    ListView2.Tag = Orden
    'Cad = "select Numero_llamado,Fecha,Hora_inicio,Unidad_de_medida,Codigo_de_trafico,"
    Cad = "SELECT Numero_llamado,Fecha,Codigo_de_trafico, Tipo_de_trafico,Cantidad_medida_originada,"
    Cad = Cad & " Unidad_de_medida,Importe,Hora_inicio from telefono.detalle_de_llamadas"
    Cad = Cad & " where fichero='" & Where2 & "' and   Numero_de_telefono='" & Text1(0).Text & "' and fecha<>'0000'"
    'Cad = Cad & " ORDER BY fecha,Hora_inicio"
    Cad = Cad & " ORDER BY " & Orden + 1
    If Orden <> 1 Then
        Cad = Cad & ",2,8"
    Else
        Cad = Cad & ",8"
    End If
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        While Not miRsAux.EOF
            'Cad = "** " & Format(miRsAux!NumAlbar, "0000") & " - " & miRsAux!NomArtic & "(" & miRsAux!codArtic & ")"
            ListView2.ListItems.Add , , DBLet(miRsAux!Numero_llamado, "T") & " "
            With ListView2.ListItems(ListView2.ListItems.Count)
                    
                    .SubItems(1) = Mid(miRsAux!Fecha, 3, 2) & "/" & Mid(miRsAux!Fecha, 1, 2) & "  " & Mid(miRsAux!Hora_inicio, 1, 2) & ":" & Mid(miRsAux!Hora_inicio, 3, 2)
                    .SubItems(2) = miRsAux!Codigo_de_trafico
                    .SubItems(3) = miRsAux!Tipo_de_trafico
                    .SubItems(4) = Format(miRsAux!Cantidad_medida_originada, "#,##0.00")
                    .SubItems(5) = miRsAux!Unidad_de_medida
                    .SubItems(6) = Format(miRsAux!Importe, "#,##0.0000")
                
                
            End With
            miRsAux.MoveNext
        Wend

    miRsAux.Close
    

End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    DetalleLlamada ColumnHeader.Index - 1
End Sub



Private Sub ExportaCSV()
Dim NF As Integer

    On Error GoTo eExp
        NF = FreeFile
        Open App.Path & "\docum.csv" For Output As #NF
        
        'Cabecera
        Cad = ""
        For NumRegElim = 1 To ListView2.ColumnHeaders.Count
            Cad = Cad & ";""" & ListView2.ColumnHeaders(NumRegElim).Text & """"
        Next NumRegElim
        Print #NF, Mid(Cad, 2)
    
    
        'Lineas
        For NumRegElim = 1 To ListView2.ListItems.Count
            Cad = """" & Trim(ListView2.ListItems(NumRegElim)) & """"
            For davidNumalbar = 1 To ListView2.ColumnHeaders.Count - 1
                Cad = Cad & ";""" & Trim(ListView2.ListItems(NumRegElim).SubItems(davidNumalbar)) & """"
            Next davidNumalbar
            Print #NF, Cad
            
            
        Next NumRegElim

        
        Close #NF



    
        cd1.Filter = "Archivo csv|*.csv"
        'Ofertare un nombre
        cd1.FileName = ""
        'cd1.InitDir = "c:\"
        cd1.CancelError = False
        cd1.ShowSave
        If cd1.FileName = "" Then Exit Sub
        
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("Ya existe el fichero: " & cd1.FileName & vbCrLf & vbCrLf & "¿Reemplazar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
        FileCopy App.Path & "\docum.csv", cd1.FileName


    Exit Sub
eExp:
    MuestraError Err.Number, Err.Description
    
End Sub

