Attribute VB_Name = "libImprmeUnCliente"
Option Explicit




'Eliminamos temporales
'ListadoFacturas :  llevará el (codtipom,numfactu,fecfactu) in ('XXX',120000,'2020-01-01'), ('XXX',120001,'2020-01-02') .....
Public Sub ImprimeFacturasCliente(Codigo As Long, ListadoFacturas As String, ByRef lblInd As Label)
Dim Cade As String
Dim Aux As String
Dim EsdesdeTelCabFact As Boolean
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim rpt As String


    On Error GoTo eImprimeFacturasCliente

    conn.Execute "DELETE from tmpnlotes where codusu =" & vUsu.Codigo
    
    
    Set miRsAux = New ADODB.Recordset
    
        
    Screen.MousePointer = vbHourglass
    lblInd.Caption = "Devolver registros"
    lblInd.Refresh
    
    
    'Vamos a meter todas las facturas en la tabla temporal
    Cade = "Select codtipom,numfactu,codclien,fecfactu,totalfac,coddirec from scafac where codclien = " & Codigo
    Cade = Cade & " AND  (codtipom,numfactu,fecfactu) IN  (" & ListadoFacturas & ")"
    'El orden vamos a hacerlo por: Tipo documento
    Cade = Cade & " ORDER BY codtipom"
    
    miRsAux.Open Cade, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Aux = ""
    Screen.MousePointer = vbHourglass
    While Not miRsAux.EOF
    
        
        lblInd.Caption = miRsAux!Numfactu
        lblInd.Refresh
        
        
        Cade = ", (" & vUsu.Codigo & ",'" & miRsAux!codtipom & "'," & miRsAux!codClien & "," & miRsAux!Numfactu & "," & CStr(miRsAux!Numfactu Mod 32000) & ",'" & Format(miRsAux!FecFactu, FormatoFecha)
        
        'El tipo de informe lo guardare en el ultimo campo
        'El report es el = 12
        Cade = Cade & "',12," & TransformaComasPuntos(CStr(DBLet(miRsAux!TotalFac, "N")))
        
        'Abril 2020. Grabaamos en numlotes el coddirec, para ver si tiene direcion email .
        'Si no pondremos la del cliente
        Cade = Cade & "," & DBSet(miRsAux!CodDirec, "T", "S") & ")"
        
        Aux = Aux & Cade
        
        NumRegElim = NumRegElim + 1
        If (NumRegElim Mod 50) = 0 Then DoEvents
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Aux <> "" Then
        Aux = Mid(Aux, 2) 'quiito la primera coma
        Cade = "insert into tmpnlotes (codusu,numalbar , codprove,codartic,numlinea,fechaalb,codalmac,cantidad,numlotes) VALUES "
        conn.Execute Cade & Aux
    End If
    

    If NumRegElim = 0 Then Err.Raise 513, , "Error leyendo facturas para impresion"
        
    '
    'Ahora cojemos las facturas que son FVA pero tienen numero terminal. COn el desde /hasta seleccionado
    'MIRAMOS en la tabla scafac1
    lblInd.Caption = "Leyendo fav "
    lblInd.Refresh
    'Compruebo si tiene codclien
    Cade = "select scafac1.* from scafac1 ,scafac where scafac1.codtipom=scafac.codtipom and scafac1.numfactu=scafac.numfactu and scafac1.fecfactu =scafac.fecfactu"
    'NomTabla = "Select codtipom,numfactu,fecfactu from scafac1   " & cadSelect
    Cade = Cade & " AND (scafac1.codtipom,scafac1.numfactu,scafac1.fecfactu) IN (" & ListadoFacturas & ")  AND numtermi>=0  "
    
    miRsAux.Open Cade, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        lblInd.Caption = DBLet(miRsAux!codtipom, "T") & ": " & miRsAux!Numfactu
        lblInd.Refresh
        Cade = "numalbar = '" & miRsAux!codtipom & "' AND fechaalb = '" & Format(miRsAux!FecFactu, FormatoFecha) & "' AND numlinea = " & CStr(miRsAux!Numfactu Mod 32000)
        'El tipo de informe lo guardare en el ultimo campo
        'El report es el = 12
        Cade = "UPDATE tmpnlotes SET codalmac = 18 WHERE codusu = " & vUsu.Codigo & " AND " & Cade
        conn.Execute Cade
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    'TROZO COPIADO DESDE envio facturas email / face
    lblInd.Caption = "Generando documentos"
    lblInd.Refresh
    
    
    'AHora las fras  FAT tienen otro report
    If vParamAplic.TieneTelefonia2 > 0 Then
        Cade = "UPDATE tmpnlotes SET codalmac = 63 WHERE codusu = " & vUsu.Codigo & " AND numalbar= 'FAT'"
        conn.Execute Cade
    End If
    'Los tikets=66
    Cade = "UPDATE tmpnlotes SET codalmac = 66 WHERE codusu = " & vUsu.Codigo & " AND numalbar= 'FTI'"
    conn.Execute Cade
    
    
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        'Facturas alvic
        Cade = "UPDATE tmpnlotes SET codalmac = 93 WHERE codusu = " & vUsu.Codigo & " AND numalbar in ('FA1','FA2','FAB','FAD')"
        conn.Execute Cade
    End If
    Espera 1
    
    Cade = "Select * from tmpnlotes where codusu =" & vUsu.Codigo & " ORDER BY codalmac,numalbar,codprove"
    miRsAux.Open Cade, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = -1
    While Not miRsAux.EOF
        Screen.MousePointer = vbHourglass
        lblInd.Caption = "Fra: " & miRsAux!Numalbar & miRsAux!codArtic
        lblInd.Refresh
        
        
        
        If NumRegElim <> miRsAux!codAlmac Then   'If CodClien <> RS!codTipoM Then
            'OTRO TIPO DE DOCUMENTO
            
            '''''If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
            
            If Not PonerParamRPT2(miRsAux!codAlmac, cadParam, numParam, rpt, False, "", 0) Then
                Err.Raise 513, , "Obteniendo rpt de impresion factura: " & miRsAux!codAlmac
            End If
            NumRegElim = miRsAux!codAlmac
                        
            If vParamAplic.NumeroInstalacion = vbTaxco Then
                cadParam = cadParam & "pDuplicado=1|"
                numParam = numParam + 1
            End If
            
            'PUNTO VERDE
            '--------------------------------------------------------------------------
            If vParamAplic.ArtReciclado <> "" Then
                cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
                numParam = numParam + 1
            End If
            
            
        End If
        
        
        EsdesdeTelCabFact = False
        
        
        'Ahora empezará a imprimir, sin abrir ni ositias
        EsdesdeTelCabFact = False
        If miRsAux!Numalbar = "FAT" Then
            If vParamAplic.TieneTelefonia2 = 1 Then EsdesdeTelCabFact = True
        End If
        'If Rs!NumAlbar = "FAT" Then
        'If EsdesdeTelCabFact Then
        '
        '    'Factura de telefonia. Lleva otro SELECT     serie
        '    cadFormula = "{tel_cab_factura.Serie} ='" & miRsAux!numlotes & "' and {tel_cab_factura.Ano} =" & Year(miRsAux!FechaAlb) & " and {tel_cab_factura.NumFact} =" & miRsAux!codArtic
        'Else
            'RESTO de facturas
            cadFormula = "({scafac.codtipom}='" & miRsAux!Numalbar & "') "
            cadFormula = cadFormula & " AND ({scafac.numfactu}=" & miRsAux!codArtic & ") "
            cadFormula = cadFormula & " AND ({scafac.fecfactu}= Date(" & Year(miRsAux!FechaAlb) & "," & Month(miRsAux!FechaAlb) & "," & Day(miRsAux!FechaAlb) & "))"
        'End If
        
        
    
        Cade = "scafac.codtipom='" & miRsAux!Numalbar & "' "
        Cade = Cade & " AND scafac.numfactu=" & miRsAux!codArtic & " "
        Cade = Cade & " AND scafac.fecfactu= " & DBSet(miRsAux!FechaAlb, "F") & " AND 1 "
                    
        Cade = DevuelveDesdeBD(conAri, "nomrpt", "scafac", Cade, "1")
        
        If Cade <> "" Then

            If Dir(App.Path & "\informes\" & Cade, vbArchive) = "" Then
                Cade = "No existe RPT con el que se imprimió: " & Cade
                Cade = Cade & vbCrLf & "¿Continuar?"
                If MsgBox(Cade, vbQuestion + vbYesNoCancel) <> vbYes Then Err.Raise 513, , "Cancelado"
                Cade = rpt
            
            End If
        Else
            Cade = rpt
        End If


          
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = True
            .EnvioEMail = False
            .NombreRPT = Cade
            .NombrePDF = Cade
            .Opcion = 53
            .Titulo = ""
        End With
        
        Load frmImprimir
   '     Unload frmImprimir
        Screen.MousePointer = vbHourglass
        Set frmImprimir = Nothing
        Espera 0.5
        Set frmImprimir = Nothing
        
        
        miRsAux.MoveNext
        
        
    Wend
    miRsAux.Close
    ListadoFacturas = "OK"
    
eImprimeFacturasCliente:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
    End If
    
    Set miRsAux = Nothing
End Sub
