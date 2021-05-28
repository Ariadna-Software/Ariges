Attribute VB_Name = "modTelefonia"
Option Explicit

Public Sub CargaLwTelefonia(ByRef LstV As ListView, Serie As String, Ano As Integer, Numfactu As Long, FormatoDelPrecio As String, SoloDistintoCero As Boolean)
Dim cad As String
Dim RS As ADODB.Recordset
Dim IT
Dim Where As String

    Set RS = New ADODB.Recordset
    
    Where = "Serie = '" & Serie & "' AND Ano =" & Ano & " AND NumFact =" & Numfactu
    If SoloDistintoCero Then Where = Where & " AND importe<>0"
    
    cad = "Select DescTipoTrafico ,importe  from tel_lin_factura_consumos WHERE  " & Where
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = LstV.ListItems.Add()
        IT.Text = RS!DescTipoTrafico
        IT.SubItems(1) = Format(RS!Importe, FormatoDelPrecio)
        'El icono
        
        RS.MoveNext
    Wend
    RS.Close
    
    cad = "Select  DescCuota , importe,fechainicio,fechafin  from tel_lin_factura_cuotas WHERE  " & Where
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = LstV.ListItems.Add()
        cad = ""
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            If Not IsNull(RS!fechainicio) Then cad = Format(RS!fechainicio, "ddmmyyyy")
            If Not IsNull(RS!FechaFin) Then cad = cad & "  " & Format(RS!FechaFin, "ddmmyyyy")
            If cad <> "" Then cad = " (" & cad & ")"
        End If
        
        IT.Text = RS!DescCuota & cad
        
        IT.SubItems(1) = Format(RS!Importe, FormatoDelPrecio)
        'El icono
        
        RS.MoveNext
    Wend
    RS.Close

    cad = "Select  Concepto ,importe  from tel_lin_factura_descuentos WHERE  " & Where
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = LstV.ListItems.Add()
        IT.Text = RS!Concepto
        IT.SubItems(1) = Format(-RS!Importe, FormatoDelPrecio)
        'El icono
        
        RS.MoveNext
    Wend
    RS.Close


    cad = "Select  Concepto ,importe  from tel_lin_factura_especial WHERE  " & Where
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = LstV.ListItems.Add()
        IT.Text = RS!Concepto
        IT.SubItems(1) = Format(RS!Importe, FormatoDelPrecio)
        'El icono
        
        RS.MoveNext
    Wend
    RS.Close

    
    Set RS = Nothing
End Sub

'Es igaul que el de arriba, por cada telefono que tenga la agrupacion
Public Sub CargaLwTelefoniaAgrupadoTfono(ByRef LstV As ListView, Serie As String, Ano As Integer, Numfactu As Long, FormatoDelPrecio As String, SoloDistintoCero As Boolean)
Dim cad As String
Dim RS As ADODB.Recordset
Dim IT
Dim Where As String
Dim i As Integer
Dim Importe As Currency
Dim CadenaTelefonoImporte As String   'cada item son 10carcar telefono 10 importe ....
    Set RS = New ADODB.Recordset
    
    LstV.ListItems.Clear
    
    Where = "Serie = '" & Serie & "' AND Ano =" & Ano & " AND NumFact =" & Numfactu
    
    
    'Importes
    Where = Where & " AND importe<>0"
    Where = "Select telefono ,importe,numlin   from ##tabla## WHERE  " & Where
    cad = ""
    For i = 1 To 4
        If i <> 1 Then cad = cad & " UNION "
        cad = cad & Replace(Where, "##tabla##", "tel_lin_factura_" & RecuperaValor("consumos|cuotas|descuentos|especial|", i))
    Next
    cad = cad & " ORDER BY 1"
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Where = ""
    cad = ""
    While Not RS.EOF
        If Where <> RS.Fields(0) Then
            'Otro telefono
            If Where <> "" Then
                Set IT = LstV.ListItems.Add()
                IT.Text = CStr(Where)
                IT.SubItems(1) = Format(Importe, FormatoDelPrecio)
                IT.Checked = True
                cad = cad & ", " & DBSet(Where, "T")
                
            End If
            Where = RS.Fields(0)
            Importe = 0
        End If
        Importe = Importe + RS.Fields(1)
        RS.MoveNext
    Wend
    RS.Close
    'El ultimo
    If Where <> "" Then
        Set IT = LstV.ListItems.Add()
        IT.Text = Where
        IT.SubItems(1) = Format(Importe, FormatoDelPrecio)
        IT.Checked = True
        cad = cad & ", " & DBSet(Where, "T")
    End If
    
    
    'Por si hay algun telefono agrupado que noesta aqui
    Where = Trim(cad)
    If Len(Where) > 0 Then Where = " AND not telefono in (" & Mid(Where, 2) & ")"        'quitmoa la primera coma
    Where = " WHERE Serie = '" & Serie & "' AND Ano =" & Ano & " AND NumFact =" & Numfactu & Where
    Where = "Select telefono   from tel_cab_factura_agr " & Where
    
    RS.Open Where, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = LstV.ListItems.Add()
        IT.Text = RS.Fields(0)
        IT.SubItems(1) = " "
        IT.Checked = True
        cad = cad & ", " & DBSet(Where, "T")
        RS.MoveNext
    Wend
    RS.Close
    
    If LstV.ListItems.Count = 0 Then
        MsgBox "Ningun telefono vinculado", vbExclamation
        cad = " NO"
    End If
    LstV.Tag = Mid(cad, 2)
    Set RS = Nothing
End Sub



