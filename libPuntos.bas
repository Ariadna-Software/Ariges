Attribute VB_Name = "libPuntos"
Option Explicit


    '**********************************************************************
    '**********************************************************************
    
    ' Puntos en ventas para canjes posteriores
    
    
    '**********************************************************************
    '**********************************************************************





Public Function CalcularPuntosAlbaran(cadWhere As String, FechaAlbaran As Date) As Currency
Dim Rs As ADODB.Recordset
Dim C As String
Dim Importe As Currency

    On Error GoTo eCalcularPuntosAlbaran
    CalcularPuntosAlbaran = 0
    
    If FechaAlbaran < vParamAplic.PtosFechaIncio Then Exit Function
        
    
    C = "select sum(importel) from slialb  WHERE "
    C = C & cadWhere
    'El canje suma tambien
    'C = C & " AND slialb.codartic<>" & DBSet(vParamAplic.PtosArticuloCanje, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            Importe = Rs.Fields(0) * vParamAplic.PtosAsignar
            CalcularPuntosAlbaran = Round2(Importe / vParamAplic.PtosImporteCalculo, 2)
        End If
    End If
    Rs.Close
    
eCalcularPuntosAlbaran:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Rs = Nothing
    
End Function




Public Function CalcularPuntosAlbaranCABEL(cadWhere As String, FechaAlbaran As Date) As Currency
Dim Rs As ADODB.Recordset
Dim C As String
Dim Importe As Currency

    On Error GoTo eCalcularPuntosAlbaran
    CalcularPuntosAlbaranCABEL = 0
    
    If FechaAlbaran < vParamAplic.PtosFechaIncio Then Exit Function
        
    
    C = "select sum(importel) from slialb,sartic,sfamia   WHERE slialb.codartic=sartic.codartic and sartic.codfamia="
    C = C & "sfamia.codfamia AND sfamia.PtosPermiteCanje =1  AND "
    C = C & cadWhere
    'El canje suma tambien
    'C = C & " AND slialb.codartic<>" & DBSet(vParamAplic.PtosArticuloCanje, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            Importe = Rs.Fields(0) * vParamAplic.PtosAsignar
            CalcularPuntosAlbaranCABEL = Round2(Importe / vParamAplic.PtosImporteCalculo, 2)
        End If
    End If
    Rs.Close
    
eCalcularPuntosAlbaran:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Rs = Nothing
    
End Function

