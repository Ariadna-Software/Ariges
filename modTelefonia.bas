Attribute VB_Name = "modTelefonia"
Option Explicit

Public Sub CargaLwTelefonia(ByRef LstV As ListView, Serie As String, Ano As Integer, Numfactu As Long, FormatoDelPrecio As String)
Dim cad As String
Dim Rs As ADODB.Recordset
Dim IT
Dim Where As String

    Set Rs = New ADODB.Recordset
    
    Where = "Serie = '" & Serie & "' AND Ano =" & Ano & " AND NumFact =" & Numfactu
    cad = "Select DescTipoTrafico ,importe  from tel_lin_factura_consumos WHERE  " & Where
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set IT = LstV.ListItems.Add()
        IT.Text = Rs!DescTipoTrafico
        IT.SubItems(1) = Format(Rs!Importe, FormatoDelPrecio)
        'El icono
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    cad = "Select  DescCuota , importe  from tel_lin_factura_cuotas WHERE  " & Where
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set IT = LstV.ListItems.Add()
        IT.Text = Rs!DescCuota
        IT.SubItems(1) = Format(Rs!Importe, FormatoDelPrecio)
        'El icono
        
        Rs.MoveNext
    Wend
    Rs.Close

    cad = "Select  Concepto ,importe  from tel_lin_factura_descuentos WHERE  " & Where
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set IT = LstV.ListItems.Add()
        IT.Text = Rs!Concepto
        IT.SubItems(1) = Format(-Rs!Importe, FormatoDelPrecio)
        'El icono
        
        Rs.MoveNext
    Wend
    Rs.Close


    cad = "Select  Concepto ,importe  from tel_lin_factura_especial WHERE  " & Where
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set IT = LstV.ListItems.Add()
        IT.Text = Rs!Concepto
        IT.SubItems(1) = Format(Rs!Importe, FormatoDelPrecio)
        'El icono
        
        Rs.MoveNext
    Wend
    Rs.Close

    
    Set Rs = Nothing
End Sub

