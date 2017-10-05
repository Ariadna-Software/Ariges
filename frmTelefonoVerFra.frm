VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelefonoVerFra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver datos prefactura"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   2280
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3840
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Height          =   375
      Index           =   1
      Left            =   4560
      Picture         =   "frmTelefonoVerFra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Suma datos seleccionados"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdExport 
      Height          =   375
      Index           =   0
      Left            =   5040
      Picture         =   "frmTelefonoVerFra.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exportar csv"
      Top             =   2280
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
      Height          =   285
      Index           =   5
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   2280
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   5640
      TabIndex        =   7
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2280
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
      Top             =   2880
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
   Begin VB.Label Label1 
      Caption         =   "Exento"
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   21
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   1560
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
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "RESUMEN"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "TOTAL"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   12
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "IVA"
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Imponible"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2040
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
Dim cad As String



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
        
        cad = "|"
        For NumRegElim = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(NumRegElim).Selected Then
                If InStr(1, cad, "|" & Me.ListView2.ListItems(NumRegElim).SubItems(2) & "|") = 0 Then cad = cad & ListView2.ListItems(NumRegElim).SubItems(2) & "|"
            End If
        Next
        'Ya tengo los distintos conceptos
        'AAhora busco
        cad = Mid(cad, 2) 'quito el primer pipe
        CadenaDesdeOtroForm = ""
        Do
            davidNumalbar = InStr(1, cad, "|")
            If davidNumalbar = 0 Then
                cad = ""
            Else
                pPdfRpt = Mid(cad, 1, davidNumalbar - 1)
                cad = Mid(cad, davidNumalbar + 1)
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
        Loop Until cad = ""
        pPdfRpt = "": pImprimeDirecto = False: davidNumalbar = 0
        CadenaDesdeOtroForm = "Resumen datos seleccionados (" & Seleccionados & ") " & vbCrLf & CadenaDesdeOtroForm
        MsgBox CadenaDesdeOtroForm, vbInformation
        CadenaDesdeOtroForm = ""
    End If
End Sub



Private Sub Form_Activate()
Dim RS As ADODB.Recordset
    
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

    'Limpiar
End Sub

Private Sub CargarDatos()


Dim IT As ListItem

    Set miRsAux = New ADODB.Recordset
    
    'Cabecera
    cad = "Select * from tel_cab_factura WHERE  " & RecuperaValor(Where2, 2)
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUESER EOF
    Text1(0).Text = miRsAux!Telefono
         cad = ""
         If Not IsNull(miRsAux!apellido1) Then cad = miRsAux!apellido1
         If Not IsNull(miRsAux!apellido2) Then cad = Trim(cad & " " & miRsAux!apellido2)
         If Not IsNull(miRsAux!Nombre) Then
             If cad <> "" Then cad = cad & ","
             cad = Trim(cad & " " & Trim(miRsAux!Nombre))
         End If
    
    Text1(1).Text = cad
    Text1(2).Text = miRsAux!Fecha
    Text1(3).Text = Format(miRsAux!BaseImponible, "#,##0.00") ' Cuota Total
    Text1(4).Text = Format(miRsAux!Cuota, "#,##0.00") ' Cuota Total
    Text1(5).Text = Format(miRsAux!total, "#,##0.00") ' Cuota Total
    Text1(6).Text = ""
    If DBLet(miRsAux!base_exenta, "N") <> 0 Then Text1(6).Text = Format(miRsAux!base_exenta, "#,##0.00") ' Cuota Total
    
    If vParamAplic.TieneTelefonia2 = 3 Then
        cad = FormatoPrecio
    Else
        cad = FormatoImporte
    End If
    CargaLwTelefonia Me.ListView1, miRsAux!Serie, miRsAux!Ano, miRsAux!NumFact, cad
    
    miRsAux.Close
    
     Label1(8).Caption = ""
    If TieneAlbaranes Then
        cad = "Select nomclien,scaalb.numalbar,codartic,nomartic,importel from scaalb left join slialb on scaalb.codtipom=slialb.codtipom"
        cad = cad & " AND scaalb.numalbar = slialb.numalbar "
        cad = cad & " WHERE referenc = '" & Text1(0).Text & "' AND factursn=1"
        
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            MsgBox "Albaranes NO encotrados", vbExclamation
        Else
            If miRsAux!NomClien <> Text1(1).Text Then
                '
                Text1(1).ForeColor = vbRed
                Label1(8).Caption = miRsAux!NomClien
            Else
                Text1(1).ForeColor = vbBlack
            End If
            While Not miRsAux.EOF
                cad = "** " & Format(miRsAux!NumAlbar, "0000") & " - " & miRsAux!NomArtic & "(" & miRsAux!codArtic & ")"
                ListView1.ListItems.Add , , cad
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = Format(miRsAux!ImporteL, "#,##0.00")
                ListView1.ListItems(ListView1.ListItems.Count).ForeColor = vbBlue
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems(1).ForeColor = vbBlue
                miRsAux.MoveNext
            Wend
        End If
        miRsAux.Close
    End If
    
    
    DoEvents
    Where2 = RecuperaValor(Where2, 1)
    
    DetalleLlamada 1
    
    Set miRsAux = Nothing
End Sub



Private Sub DetalleLlamada(Orden As Byte)

    'Detalle llamada
    If ListView2.Tag = Orden Then Exit Sub
    
    Me.ListView2.ListItems.Clear
    ListView2.Tag = Orden
    'Cad = "select Numero_llamado,Fecha,Hora_inicio,Unidad_de_medida,Codigo_de_trafico,"
    cad = "SELECT Numero_llamado,Fecha,Codigo_de_trafico, Tipo_de_trafico,Cantidad_medida_originada,"
    cad = cad & " Unidad_de_medida,Importe,Hora_inicio from telefono.detalle_de_llamadas"
    cad = cad & " where fichero='" & Where2 & "' and   Numero_de_telefono='" & Text1(0).Text & "' and fecha<>'0000'"
    'Cad = Cad & " ORDER BY fecha,Hora_inicio"
    cad = cad & " ORDER BY " & Orden + 1
    If Orden <> 1 Then
        cad = cad & ",2,8"
    Else
        cad = cad & ",8"
    End If
    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

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
        cad = ""
        For NumRegElim = 1 To ListView2.ColumnHeaders.Count
            cad = cad & ";""" & ListView2.ColumnHeaders(NumRegElim).Text & """"
        Next NumRegElim
        Print #NF, Mid(cad, 2)
    
    
        'Lineas
        For NumRegElim = 1 To ListView2.ListItems.Count
            cad = """" & Trim(ListView2.ListItems(NumRegElim)) & """"
            For davidNumalbar = 1 To ListView2.ColumnHeaders.Count - 1
                cad = cad & ";""" & Trim(ListView2.ListItems(NumRegElim).SubItems(davidNumalbar)) & """"
            Next davidNumalbar
            Print #NF, cad
            
            
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
