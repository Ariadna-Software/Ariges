Attribute VB_Name = "libPuntos"
Option Explicit


    '**********************************************************************
    '**********************************************************************
    
    ' Puntos en ventas para canjes posteriores
    
    
    '**********************************************************************
    '**********************************************************************





Public Function CalcularPuntosAlbaran(cadWhere As String, FechaAlbaran As Date) As Currency
Dim RS As ADODB.Recordset
Dim C As String
Dim Importe As Currency

    On Error GoTo eCalcularPuntosAlbaran
    CalcularPuntosAlbaran = 0
    
    If FechaAlbaran < vParamAplic.PtosFechaIncio Then Exit Function
        
    
    C = "select sum(importel) from slialb  WHERE "
    C = C & cadWhere
    'El canje suma tambien
    'C = C & " AND slialb.codartic<>" & DBSet(vParamAplic.PtosArticuloCanje, "T")
    
    Set RS = New ADODB.Recordset
    RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            Importe = RS.Fields(0) * vParamAplic.PtosAsignar
            CalcularPuntosAlbaran = Round2(Importe / vParamAplic.PtosImporteCalculo, 2)
        End If
    End If
    RS.Close
    
eCalcularPuntosAlbaran:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
    
End Function



Public Function CalcularPuntosAlbaranCABEL(cadWhere As String, FechaAlbaran As Date, ByRef ImporteLineasTotal As String, ByRef Comision As String) As Currency
Dim RS As ADODB.Recordset
Dim C As String
Dim Importe As Currency

    On Error GoTo eCalcularPuntosAlbaran
    CalcularPuntosAlbaranCABEL = 0
    ImporteLineasTotal = ""
    Comision = ""
    If FechaAlbaran < vParamAplic.PtosFechaIncio Then Exit Function
        
    C = "max"
    If InStr(1, cadWhere, "ART") > 0 Then C = "min"
    C = C & "(comisionagente)"
    C = "select sum(if( sfamia.PtosPermiteCanje =1,importel,0)),sum(importel)," & C & " from slialb,sartic,sfamia   WHERE slialb.codartic=sartic.codartic and sartic.codfamia="
    C = C & "sfamia.codfamia   AND "
    C = C & cadWhere
    C = C & " AND slialb.codartic<>" & DBSet(vParamAplic.PtosArticuloCanje, "T")
    
    
    Set RS = New ADODB.Recordset
    RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            Importe = RS.Fields(0) * vParamAplic.PtosAsignar
            CalcularPuntosAlbaranCABEL = Round2(Importe / vParamAplic.PtosImporteCalculo, 2)
            ImporteLineasTotal = DBLet(RS.Fields(0), "N")  'AGosto 2019. Ponia field(1) que es sum(importel)
            If Not IsNull(RS.Fields(2)) Then Comision = Format(RS.Fields(2), FormatoImporte)
                
        End If
    End If
    RS.Close
    
eCalcularPuntosAlbaran:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
        Set RS = Nothing
    
End Function

