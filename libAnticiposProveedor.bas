Attribute VB_Name = "libAnticiposProveedor"
Option Explicit





Public Function EstadoAnticipoEnContabilidad(idAnticipo As Long) As Boolean
Dim Sql As String
    
    EstadoAnticipoEnContabilidad = False
        
    Sql = "select idanticipo,sproveanticipo.codprove,nomprove,numdocum,fechaant,codmacta"
    Sql = Sql & " from sproveanticipo left join sprove on sproveanticipo.codprove=sprove.codprove where idanticipo=" & idAnticipo
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Sql = "SELECT * from pagos where codmacta='" & miRsAux!Codmacta & "' and numfactu=" & DBSet(miRsAux!numdocum, "T")
    Sql = Sql & "  and numserie=" & DBSet(vParamAplic.SerieAnticipoProveedor, "T") & " and numorden=1"
    Sql = Sql & " and fecfactu=" & DBSet(miRsAux!fechaant, "F")
    miRsAux.Close
    
    
    'Ahora abro en contabilidad , a ver si esta el cobro
    miRsAux.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    If miRsAux.EOF Then
        Sql = "No existe el anticipo en contabilidad"
    Else
        If DBLet(miRsAux!imppagad, "N") <> 0 Then Sql = "Tiene importe pagado en contabilidad: " & miRsAux!imppagad
    End If

    If Sql <> "" Then
        Sql = Sql & vbCrLf & vbCrLf & "Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then Sql = ""
    End If
    
     EstadoAnticipoEnContabilidad = Sql = ""

End Function



Public Function InsertarAnticipoEnContabilidad(idAnticipo As Long) As Boolean
Dim Sql As String
Dim Cta As String

    On Error GoTo eInsertarEnContabilidad
    
    
    
    Set miRsAux = New ADODB.Recordset
        
        
    'De momento, por no parametrizar
    Sql = "Select ctabanc1,numserie from pagos where fecfactu > 20200101 AND ctabanc1 like '572%' ORDER by numserie desc,ctabanc1"
    miRsAux.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cta = DBLet(miRsAux!ctabanc1, "T")
    miRsAux.Close
    If Cta = "" Then
        Stop
        
    End If
    
    Sql = "select idanticipo,sproveanticipo.codprove,fechaant,sproveanticipo.codforpa fp,numdocum,importe,sprove.*"
    Sql = Sql & " from sproveanticipo left join sprove on sproveanticipo.codprove=sprove.codprove where idanticipo=" & idAnticipo
    
    
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Sql = "INSERT INTO pagos (numserie,codmacta,numfactu,fecfactu,numorden,codforpa,fecefect,impefect,ctabanc1,"
    Sql = Sql & "text1csb,text2csb,referencia,nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais) VALUES ("
    'numserie,codmacta,numfactu,fecfactu,numorden,codforpa,fecefect,impefect,ctabanc1,
    Sql = Sql & DBSet(vParamAplic.SerieAnticipoProveedor, "T") & "," & DBSet(miRsAux!Codmacta, "T") & "," & DBSet(miRsAux!numdocum, "T") & "," & DBSet(miRsAux!fechaant, "F") & ",1,"
    Sql = Sql & miRsAux!fp & "," & DBSet(miRsAux!fechaant, "F") & "," & DBSet(miRsAux!Importe, "N") & ",'" & Cta & "',"
    
    'text1csb,text2csb,referencia,nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais
    Sql = Sql & "'Anticipo prov: " & miRsAux!Codprove & "   Id: " & Format(idAnticipo, "0000") & "'," & DBSet("Creado por " & vUsu.Login, "T")
    Sql = Sql & ",'" & Format(idAnticipo, "000000") & "'," & DBSet(miRsAux!nomprove, "T") & "," & DBSet(miRsAux!domprove, "T")
    Sql = Sql & "," & DBSet(miRsAux!pobprove, "T") & "," & DBSet(miRsAux!codpobla, "T") & "," & DBSet(miRsAux!proprove, "T")
    Sql = Sql & "," & DBSet(miRsAux!nifProve, "T") & "," & DBSet(miRsAux!codpais, "T") & ")"
    ConnConta.Execute Sql
    
    
eInsertarEnContabilidad:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
    End If
    Set miRsAux = Nothing
End Function


'                       DatosVto:   codmactaprov|numdcoum|fecdocum|
Public Function BorrarAnticipoEnContabilidad(DatosVto As String) As Boolean


Dim Sql As String


    On Error GoTo eBorrarAnticipoEnContabilidad
        BorrarAnticipoEnContabilidad = False
        Sql = "Delete from pagos where numserie =" & DBSet(vParamAplic.SerieAnticipoProveedor, "T")
        Sql = Sql & " AND codmacta = " & DBSet(RecuperaValor(DatosVto, 1), "T") & " AND numfactu = " & DBSet(RecuperaValor(DatosVto, 2), "T")
        Sql = Sql & " AND fecfactu = " & DBSet(RecuperaValor(DatosVto, 3), "F")
        ConnConta.Execute Sql
        BorrarAnticipoEnContabilidad = True
    
    
eBorrarAnticipoEnContabilidad:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
End Function
