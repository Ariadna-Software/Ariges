VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFacTPVTraerVen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas de otros terminales"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   ClipControls    =   0   'False
   Icon            =   "frmFacTPVTraerVen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdVerLin 
      Caption         =   "Ver lineas"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4020
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7091
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   4320
      Width           =   1035
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmFacTPVTraerVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public parNumTermi As Integer 'nº de terminal en el que estamos conectados
                              '  -1: Reimpresion de tickets o ver lineas

Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
'Solo utilizamos el Modo=4 -> Modificar

Dim kCampo As Integer
Dim OrdenList As Integer

Public Event CargarVenta(cadSel As String, numVen As Long)

Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim cad As String

    On Error GoTo Error1

    'Comprobar que solo se ha seleccionado una venta
    For i = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(i).Checked Then
            cad = "(numtermi=" & Me.ListView1.ListItems(i).Text & " AND numventa=" & Me.ListView1.ListItems(i).ListSubItems(1).Text
            cad = cad & " AND fecventa=" & DBSet(Me.ListView1.ListItems(i).ListSubItems(2).Text, "F") & ")"
            
            RaiseEvent CargarVenta(cad, CLng(Me.ListView1.ListItems(i).ListSubItems(1).Text))
            Exit For
        End If
    Next i

    Screen.MousePointer = vbHourglass
'    NumOfe = Text1.Text
    Unload Me
    

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Unload Me
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdImprimir_Click()
    DevolverDatosAlFormulario 1
End Sub

Private Sub cmdVerLin_Click()
    DevolverDatosAlFormulario 2
End Sub


Private Sub DevolverDatosAlFormulario(LaOpcionSeleccionada As Integer)
Dim i As Integer
Dim cad As String

    On Error GoTo Error2
    cad = ""
    'Comprobar que solo se ha seleccionado una venta
    For i = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(i).Checked Then
                'Tipo documento
            cad = Me.ListView1.ListItems(i).SubItems(2) & "|" & Me.ListView1.ListItems(i).Tag
            
            RaiseEvent CargarVenta(cad, -1 * LaOpcionSeleccionada)
            Exit For
        End If
    Next i
    If cad = "" Then
        MsgBox "Seleccione algún dato", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
'    NumOfe = Text1.Text
    Unload Me
    Exit Sub
Error2:
    MuestraError Err.Number
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
        
    
    Me.cmdAceptar.visible = parNumTermi >= 0
    
    Me.cmdImprimir.visible = parNumTermi = -1
    cmdVerLin.visible = parNumTermi = -1
    
'    PonerModo 0
    If parNumTermi >= 0 Then
        
        CargarListView_VentasTPV
        Caption = "Ventas de otros terminales"
    Else
        'Caption
        Caption = "Ventas dia: " & Format(Now, "dd/mm/yyyy")
        
        CargarListView_VentasDia
    End If
    Screen.MousePointer = vbDefault
End Sub


'Private Sub Text1_GotFocus()
'    ConseguirFoco Text1, Modo
'End Sub

'Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'    KEYdown KeyCode
'End Sub

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub

'Private Sub Text1_LostFocus()
'Dim devuelve As String
'
'    With Text1
'        If .Text = "" Then Exit Sub
'        .Text = Format(.Text, "0000000")
'        'Comprobar que la oferta existe
'        devuelve = DevuelveDesdeBDNew(conAri, "scapre", "numofert", "numofert", .Text, "N")
'        If devuelve = "" Then
'            MsgBox "No existe la Oferta: " & .Text, vbInformation
'            Text1.Text = ""
'            PonerFoco Text1
'        End If
'    End With
'End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
       
    Modo = Kmodo
End Sub



Private Sub CargarListView_VentasTPV()
'Muestra en una lista las ventas de otros terminales
'para seleccionar la que queremos recuperar en mi terminal activo
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ErrCargar
     
    
    'limpiamos Los encabezados
    ListView1.ColumnHeaders.Clear
    
    'cargamos los encabezados
    ListView1.ColumnHeaders.Add , , "Termi.", 900
    ListView1.ColumnHeaders.Add , , "Nº Venta", 1000, 1
    ListView1.ColumnHeaders.Add , , "Fecha Venta", 1250, 2
    ListView1.ColumnHeaders.Add , , "Cliente", 900, 2
    ListView1.ColumnHeaders.Add , , "Nombre Cliente", 3510, 0
    ListView1.ColumnHeaders.Add , , "Total (€)", 1150, 1
    
    SQL = "SELECT numtermi,numventa,fecventa,scaven.codclien, nomclien,imptotal "
'    SQL = SQL & " FROM (scaven INNER JOIN sliven ON scaven.numtermi=sliven.numtermi AND scaven.numventa=sliven.numventa AND scaven.fecventa=sliven.fecventa) "
    SQL = SQL & " FROM scaven "
    SQL = SQL & " INNER JOIN sclien ON scaven.codclien=sclien.codclien "
    'seleccionamos de otros terminales distintos al q estoy conectado
    SQL = SQL & " WHERE numtermi <>" & Me.parNumTermi
    SQL = SQL & " ORDER BY numtermi,numventa "

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Format(RS.Fields(0).Value, "0000") 'Nº terminal
        ItmX.SubItems(1) = RS.Fields(1).Value 'Nº venta
        ItmX.SubItems(2) = RS.Fields(2).Value 'Fecha venta
        ItmX.SubItems(3) = Format(RS.Fields(3).Value, "000000") 'cliente
        ItmX.SubItems(4) = RS.Fields(4).Value 'nombre cliente
        ItmX.SubItems(5) = Format(RS.Fields(5).Value, FormatoImporte) 'importe total
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'si aparece la barra de desplazamiento ajustar el ancho
    If Me.ListView1.ListItems.Count > 13 Then Me.ListView1.ColumnHeaders(5).Width = 3250
    
    Exit Sub
    
ErrCargar:
    MuestraError Err.Number, "Cargar lista ventas de otros terminales.", Err.Description
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    'Ha clikado sobre la columna. Si esta en negativo es que es la misma columna
    If Abs(OrdenList) <> ColumnHeader.Index Then
        If OrdenList <> 0 Then
            'Estaba puesto el orden
                     
        End If
        OrdenList = ColumnHeader.Index
    Else
        OrdenList = -1 * OrdenList
    End If
    ListView1.Sorted = True
    ListView1.SortKey = Abs(OrdenList) - 1
    If OrdenList < 0 Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
    'cuando se selecciona una venta se quitar las anteriores que pudiera haber
    'solo se podrá seleccionar una venta

    For i = 1 To Me.ListView1.ListItems.Count
        If i <> Item.Index Then Me.ListView1.ListItems(i).Checked = False
                
    Next i

    Me.ListView1.ListItems(Item.Index).Selected = True
End Sub







Private Sub CargarListView_VentasDia()
'Muestra en una lista las ventas de otros terminales
'para seleccionar la que queremos recuperar en mi terminal activo
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ErrCargar
     
    
    
    'limpiamos Los encabezados
    ListView1.ColumnHeaders.Clear
    
    'cargamos los encabezados
    ListView1.ColumnHeaders.Add , , "Termi.", 900
    ListView1.ColumnHeaders.Add , , "Nº Venta", 1000, 1
    ListView1.ColumnHeaders.Add , , "Tipo", 850, 2
    ListView1.ColumnHeaders.Add , , "Cliente", 900, 2
    ListView1.ColumnHeaders.Add , , "Nombre Cliente", 3210, 0
    ListView1.ColumnHeaders.Add , , "Total (€)", 1150, 1
    
    

    SQL = "select numtermi,numventa,s.codtipom,s.codclien,nomclien,totalfac,s.numfactu,numalbar,codtipoa from scafac s,scafac1 s1 where s.codtipom=s1.codtipom and s.numfactu=s1.numfactu and s.fecfactu=s1.fecfactu"
    SQL = SQL & " and (s.codtipom='FTI'  or (s.codtipom='FAV' and numtermi>0)) and s.fecfactu='" & Format(Now, FormatoFecha) & "' order by numventa"
    







    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Format(miRsAux.Fields(0).Value, "0000") 'Nº terminal
        ItmX.SubItems(1) = miRsAux.Fields(1).Value 'Nº venta
        ItmX.SubItems(2) = miRsAux.Fields(2).Value  'tipo documento
        ItmX.SubItems(3) = Format(miRsAux.Fields(3).Value, "000000") 'cliente
        ItmX.SubItems(4) = miRsAux.Fields(4).Value 'nombre cliente
        ItmX.SubItems(5) = Format(miRsAux.Fields(5).Value, "#,###,###,##0.00") 'importe total
        ItmX.Tag = DBLet(miRsAux!NumFactu, "N") & "|" & miRsAux!NumAlbar & "|" & miRsAux!codtipoa
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    'si aparece la barra de desplazamiento ajustar el ancho
    If Me.ListView1.ListItems.Count > 13 Then Me.ListView1.ColumnHeaders(5).Width = 3250
    
    Exit Sub
    
ErrCargar:
    MuestraError Err.Number, "Cargar lista ventas de otros terminales.", Err.Description
    Set miRsAux = Nothing
End Sub

