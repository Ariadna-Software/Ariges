Attribute VB_Name = "libAgua"
Option Explicit





Public Function GenerarFicheroPresentacionAgua(Anyo As Integer, L As Label) As Boolean
Dim NF As Integer
Dim Cad As String
Dim Poblaciones As Collection
Dim J As Integer
Dim K As Integer
Dim Aux As String
Dim RT As ADODB.Recordset
Dim RR As ADODB.Recordset
Dim TotalFra As Integer
Dim TotalFraFRT As Integer
Dim ImpoFacturado As Currency
Dim ImpoAnulado As Currency
Dim Impor As Currency


    On Error GoTo eGenerarFicheroPresentacionAgua
    NF = -1
    davidCodtipom = App.Path & "\AguaPres.dat"
    If Dir(davidCodtipom, vbArchive) <> "" Then Kill davidCodtipom

    
    Set miRsAux = New ADODB.Recordset
    Set RR = New ADODB.Recordset
    Set Poblaciones = New Collection
    
    'Va ordenado por POBLACION del contador, solo deberia haber uno BOLBAIE
    L.Caption = "Leyendo facturas"
    L.Refresh
    Cad = "not referenc is null  AND codtipom='FAG' AND fecfactu between '" & Anyo & "-01-01' AND '" & Anyo & "-12-31'"
    Cad = "(select referenc from scafac1 where " & Cad & ")"
    Cad = "select cpconta,pobconta from aguacontadores where contador in " & Cad & " group by 1"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux!cpconta
        Poblaciones.Add Cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Las rectificativas
    Cad = miRsAux.Source
    Cad = Replace(Cad, "'FAG'", "'FRT'")
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        For J = 1 To Poblaciones.Count
            If Poblaciones(J) <> miRsAux!cpconta Then
                Cad = miRsAux!cpconta
                Poblaciones.Add Cad
            End If
        Next J
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    If Poblaciones.Count = 0 Then Err.Raise 513, , "No existen facturas"
    
    Set RT = New ADODB.Recordset
    NF = FreeFile
    Open davidCodtipom For Output As #NF
    
    'A)    Cabecera comun
    '-----------------------------------------------
    Cad = "Select CodEntSuministra from sparamagua"
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = "10" & CadenaFichero(DBLet(miRsAux!CodEntSuministra, "T"), 10)
    Cad = Cad & CadenaFichero(vParam.NombreEmpresa, 100)
    Cad = Cad & CadenaFichero(vParam.CifEmpresa, 9) & Format(Now, "ddmmyyyy")
    Cad = Cad & "FACTURACION_DETALLADA" & CStr(Anyo) & Space(188)
    Print #NF, Cad
    
    miRsAux.Close
    
    'REUTILIZO VARIABLES.  Contador total de lineas
    davidNumalbar = 2   'Llevo ya dos
    
    For J = 1 To Poblaciones.Count
        
        'B)  Registro cabceraq del municipio.   Reutilizo davidCodtipom
        pPdfRpt = DevuelveDesdeBD(conAri, "provincia", "scpostal", "cpostal", Poblaciones(J), "N")
        Cad = "20" & CadenaFichero(Poblaciones(J), 5) & CadenaFichero(DBLet(pPdfRpt, "T"), 30) & Space(293)
            
        L.Caption = "Leyendo " & pPdfRpt
        L.Refresh
        
        Aux = "not referenc is null  AND codtipom='FAG' AND fecfactu between '" & Anyo & "-01-01' AND '" & Anyo & "-12-31'"
        Aux = " from scafac1 where " & Aux
        Aux = "select codtipom,min(fecfactu),max(fecfactu) " & Aux & " group by 1"
        miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Aux = Format(miRsAux.Fields(1), "mmyyyy") & Format(miRsAux.Fields(2), "mmyyyy")
        Cad = Cad & Aux
        Print #NF, Cad
    
        
        TotalFra = 0
        TotalFraFRT = 0
        ImpoFacturado = 0
        ImpoAnulado = 0
        davidCodtipom = ""   'LLEVARE LOS ERRORES EN LOS CONTADORES
        davidNumalbar = davidNumalbar + 1

        miRsAux.Close
        
        
        'C)  Facturas emitidas del municipo
        Aux = " AND referenc in (select contador from aguacontadores where cpconta=" & Poblaciones(J) & ")"
        Aux = " AND scafac.codtipom='FAG' AND scafac.fecfactu between '" & Anyo & "-01-01' AND '" & Anyo & "-12-31' " & Aux
        Aux = " AND scafac.fecfactu = scafac1.fecfactu AND  scafac.numfactu = scafac1.numfactu" & Aux
        
        Aux = " FROM scafac,scafac1 where scafac.codtipom = scafac1.codtipom " & Aux
        Aux = "  pobclien, observa1,observa2,observa3,referenc  " & Aux
        Aux = " codclien,nomclien,nifclien,domclien,codpobla,   " & Aux
        Aux = "SELECT scafac.codtipom,scafac.numfactu,scafac.fecfactu, " & Aux
        
        
        miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not miRsAux.EOF
        
        
            L.Caption = miRsAux!codtipom & Format(miRsAux!NumFactu, "000000")
            L.Refresh
            
        
            Aux = "Select contador,codclien,codcalibre,TipoUso from aguacontadores where contador =" & DBSet(miRsAux!referenc, "T")
            'NO PUEDE SER EOF
            RT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If miRsAux!codClien <> RT!codClien Then
                
                Cad = RT!Contador & "      " & Format(RT!codClien, "0000") & "   -" & Format(miRsAux!NumFactu, "000000") & "   " & miRsAux!codClien
                davidCodtipom = davidCodtipom & Cad & vbCrLf
            End If
            
            
            Cad = "30" & CadenaFichero(Poblaciones(J), 5)    'Codigo rg y municipio
            
            
            
            Aux = CadenaFichero(Format(miRsAux!NumFactu, "0000000"), 20)
            Cad = Cad & Aux & Format(miRsAux!FecFactu, "ddmmyyyy")  'num y fecha factura
            
            Cad = Cad & CadenaFichero(miRsAux!Nomclien, 100)
            Cad = Cad & CadenaFichero(miRsAux!nifClien, 9)
            
            'EsDomestico = miRsAux!TipoUso = 0
            Aux = "D"
            If RT!TipoUso = 1 Then Aux = "I"
            Cad = Cad & Aux
            
            'Direccion del abonado, cp y municipio (s fiera del contador tendria que ir a RT ,,
            Cad = Cad & CadenaFichero(miRsAux!domclien, 50)
            Cad = Cad & CadenaFichero(miRsAux!codpobla, 5)
            Cad = Cad & CadenaFichero(miRsAux!pobclien, 30)
            
            'Nº poliza y contador
            Cad = Cad & CadenaFichero(RT!Contador, 15) & CadenaFichero(RT!Contador, 15)
            
    
            'Calibre. --> observa3  : "CALIBRE 20"
            K = InStr(5, miRsAux!observa3, " ")
            If K = 0 Then
                Err.Raise 513, , "Observacion3 sin el calibre " & miRsAux!observa3
                
            Else
                Aux = Trim(Mid(miRsAux!observa3, K + 1))
                If Len(Aux) <> 2 Then Err.Raise 513, , "Calibre de mas de 2 digitos" & miRsAux!observa3
            End If
            Cad = Cad & "0" & Aux
            
            'Coeficiente corrector
            Cad = Cad & "0000" 'entera y decimal eedd
            
            'Fecha inicio del consumo, observa1: 24/02/2014  1410    fec lect
            If Mid(miRsAux!observa1, 1, 3) = "--/" Then
                'No hay lectura anterior
                Cad = Cad & "0101" & Anyo
                NumRegElim = 0
            Else
                K = InStr(5, miRsAux!observa1, " ")
                Aux = Trim(Mid(miRsAux!observa1, 1, K - 1))
                Aux = Replace(Aux, "/", "") 'quito las /
                Cad = Cad & Aux
                Aux = Trim(Mid(miRsAux!observa1, K + 1))
                If Val(Aux) > 32500 Then Aux = "9999"
                NumRegElim = CInt(Aux)
            End If
            
            'Fecha fin del consumo, observa2: 24/02/2014  1420    fec lect
            K = InStr(5, Trim(miRsAux!observa2), " ")
            If K = 0 Then
                'NO hay lectura posterior
                Cad = Cad & "3112" & Anyo
                Aux = NumRegElim
            Else
                Aux = Trim(Mid(miRsAux!observa2, 1, K - 1))
                Aux = Replace(Aux, "/", "") 'quito las /
                Cad = Cad & Aux
                Aux = Trim(Mid(miRsAux!observa2, K + 1))
                If Val(Aux) > 32500 Then Aux = "9999"
            End If
            NumRegElim = CInt(Aux) - NumRegElim
            If NumRegElim < 0 Then Err.Raise 513, "Consumo negativo"
        
            'Conusmo
            Cad = Cad & Format(NumRegElim, "000000000")
            
            RT.Close
            
            Aux = " codtipom =" & DBSet(miRsAux!codtipom, "T") & " AND numfactu =" & miRsAux!NumFactu
            Aux = Aux & " AND fecfactu =" & DBSet(miRsAux!FecFactu, "F") & " AND numlinea IN (21,25)"
            Aux = "SELECT numlinea,precioar,importel from slifac WHERE " & Aux & " ORDER BY numlinea"
            RT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Impor = 0 'Cuota consumo facturada
            If RT.EOF Then
                'NO tiene ni conusmo ni servicio
                Aux = String(50, "0")
                
            Else
                Aux = ""
                
                If RT!numlinea = 21 Then
                    'Cuota de consumo aplicada
                    Aux = Format(RT!precioar, "00.000000")
                    Impor = RT!ImporteL
                    RT.MoveNext
                    
                Else
                    'NO hay cuota de consumo
                    Aux = String(8, "0")
                End If
                
                'Cuota de servicio facturada
                If RT.EOF Then
                    Aux = Aux & String(14, "0")
                Else
                    Aux = Aux & Format(RT!ImporteL, "000000000000.00")
                End If
                
                'Consumo facturado
                Aux = Aux & Format(Impor, "000000000000.00")
                
                
                'Canon facturado.
                If Not RT.EOF Then
                    If RT!numlinea = 25 Then Impor = Impor + RT!ImporteL
                End If
                
                Aux = Aux & Format(Impor, "000000000000.00")
                
               'Quitamos las comas
               Aux = Replace(Aux, ",", "")
            End If
            RT.Close
            Cad = Cad & Aux
            Print #NF, Cad
            If Len(Cad) <> 342 Then Err.Raise 513, , "Error  long. linea " & Cad
            
            davidNumalbar = davidNumalbar + 1
            ImpoFacturado = ImpoFacturado + Impor 'del canon facturado emitido
            TotalFra = TotalFra + 1
            
                    
                    
            'Siguiente
            miRsAux.MoveNext
            
            
            
            
        Wend
        miRsAux.Close
        
        
        '------------------------------
        'D)  Facturas anuladas del municipo
       
        
        Aux = miRsAux.Source
        Aux = Replace(Aux, "'FAG'", "'FRT'")
        RR.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RR.EOF
        
        
            L.Caption = RR!codtipom & Format(RR!NumFactu, "000000")
            L.Refresh
            
            
            K = InStr(1, DBLet(RR!observa1, "T"), "FACTURA:")
            If K > 0 Then
                'DEJO A PIÑON la A como letra serie, habira que mirarlo
                K = InStr(K, DBLet(RR!observa1, "T"), ",")
                If K > 0 Then
                    Aux = Mid(RR!observa1, K + 2)
                    K = InStr(1, Aux, ",")
                    If K > 0 Then
                        Cad = Trim(Mid(Aux, K + 1))
                        Aux = Mid(Aux, 1, K - 1)
                        
                        If Not IsDate(Cad) Then
                            K = 0
                        Else
                            Cad = DBSet(Cad, "F")
                            Cad = " AND scafac1.numfactu = " & Aux & " AND scafac1.fecfactu =" & Cad
                            
                            Aux = RR.Source
                            Aux = Replace(Aux, "'FRT'", "'FAG'")
                            
                            'reemplzao el betwenn  ej: between '2014-01-01' AND '2014-12-31' pr  '2001  y  2030
                            '" & Anyo & "-01-01' AND '" & Anyo & "-12-31'
                            Aux = Replace(Aux, "'" & Anyo & "-01-01'", "'2000-01-01'")
                            Aux = Replace(Aux, "'" & Anyo & "-12-31'", "'2050-12-31'")
                            
                            
                            
                            
                            Aux = Aux & Cad
                            
                            miRsAux.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            If miRsAux.EOF Then
                                K = 0
                                miRsAux.Close
                            Else
                                K = 1
                            End If
                        End If
                    End If
                End If
            End If
                    
            If K > 0 Then
                    Aux = "Select contador,codclien,codcalibre,TipoUso from aguacontadores where contador =" & DBSet(miRsAux!referenc, "T")
                    'NO PUEDE SER EOF
                    RT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If miRsAux!codClien <> RT!codClien Then
                        
                        Cad = RT!Contador & "      " & Format(RT!codClien, "0000") & "   *" & Format(miRsAux!NumFactu, "000000") & "   " & miRsAux!codClien
                        davidCodtipom = davidCodtipom & Cad & vbCrLf
                    End If
                    
                    
                    Cad = "40" & CadenaFichero(Poblaciones(J), 5)    'Codigo rg y municipio
                    
                    
                    
                    Aux = CadenaFichero(Format(miRsAux!NumFactu, "0000000"), 20)
                    Cad = Cad & Aux & Format(miRsAux!FecFactu, "ddmmyyyy")  'num y fecha factura
                    
                    Cad = Cad & CadenaFichero(miRsAux!Nomclien, 100)
                    Cad = Cad & CadenaFichero(miRsAux!nifClien, 9)
                    
                    'EsDomestico = miRsAux!TipoUso = 0
                    Aux = "D"
                    If RT!TipoUso = 1 Then Aux = "I"
                    Cad = Cad & Aux
                    
                    'Direccion del abonado, cp y municipio (s fiera del contador tendria que ir a RT ,,
                    Cad = Cad & CadenaFichero(miRsAux!domclien, 50)
                    Cad = Cad & CadenaFichero(miRsAux!codpobla, 5)
                    Cad = Cad & CadenaFichero(miRsAux!pobclien, 30)
                    
                    'Nº poliza y contador
                    Cad = Cad & CadenaFichero(RT!Contador, 15) & CadenaFichero(RT!Contador, 15)
                    
            
                    'Calibre. --> observa3  : "CALIBRE 20"
                    K = InStr(5, DBLet(miRsAux!observa3, "T"), " ")
                    If K = 0 Then
                        
                        Aux = DevuelveDesdeBD(conAri, "calibre", "aguacalibre", "codcalibre", RT!codcalibre)
                        If Aux = "" Then Aux = "00"
                    Else
                        Aux = Trim(Mid(miRsAux!observa3, K + 1))
                        
                    End If
                    If Len(Aux) <> 2 Then Err.Raise 513, , "Calibre de mas de 2 digitos" & miRsAux!observa3
                    Cad = Cad & "0" & Aux
                    
                    'Coeficiente corrector
                    Cad = Cad & "0000" 'entera y decimal eedd
                    
                    'Fecha inicio del consumo, observa1: 24/02/2014  1410    fec lect
                    If Mid(miRsAux!observa1, 1, 3) = "--/" Then
                        'No hay lectura anterior
                        Cad = Cad & "0101" & Anyo
                        NumRegElim = 0
                    Else
                        K = InStr(5, miRsAux!observa1, " ")
                        Aux = Trim(Mid(miRsAux!observa1, 1, K - 1))
                        Aux = Replace(Aux, "/", "") 'quito las /
                        Cad = Cad & Aux
                        Aux = Trim(Mid(miRsAux!observa1, K + 1))
                        If Val(Aux) > 32500 Then Aux = "9999"
                        NumRegElim = CInt(Aux)
                    End If
                    
                    'Fecha fin del consumo, observa2: 24/02/2014  1420    fec lect
                    K = InStr(5, Trim(miRsAux!observa2), " ")
                    If K = 0 Then
                        'NO hay lectura posterior
                        Cad = Cad & "3112" & Anyo
                        Aux = NumRegElim
                    Else
                        Aux = Trim(Mid(miRsAux!observa2, 1, K - 1))
                        Aux = Replace(Aux, "/", "") 'quito las /
                        Cad = Cad & Aux
                        Aux = Trim(Mid(miRsAux!observa2, K + 1))
                        If Val(Aux) > 32500 Then Aux = "9999"
                    End If
                    NumRegElim = CInt(Aux) - NumRegElim
                    If NumRegElim < 0 Then Err.Raise 513, "Consumo negativo"
                
                    'Conusmo
                    Cad = Cad & Format(NumRegElim, "000000000")
                    
                    RT.Close
                    
                    Aux = " codtipom =" & DBSet(miRsAux!codtipom, "T") & " AND numfactu =" & miRsAux!NumFactu
                    Aux = Aux & " AND fecfactu =" & DBSet(miRsAux!FecFactu, "F") & " AND numlinea IN (21,25)"
                    Aux = "SELECT numlinea,precioar,importel from slifac WHERE " & Aux & " ORDER BY numlinea"
                    RT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    Impor = 0 'Cuota consumo facturada
                    If RT.EOF Then
                        'NO tiene ni conusmo ni servicio
                        Aux = String(50, "0")
                        
                    Else
                        Aux = ""
                        
                        If RT!numlinea = 21 Then
                            'Cuota de consumo aplicada
                            Aux = Format(RT!precioar, "00.000000")
                            Impor = RT!ImporteL
                            RT.MoveNext
                            
                        Else
                            'NO hay cuota de consumo
                            Aux = String(8, "0")
                        End If
                        
                        'Cuota de servicio facturada
                        If RT.EOF Then
                            Aux = Aux & String(14, "0")
                        Else
                            Aux = Aux & Format(RT!ImporteL, "000000000000.00")
                        End If
                        
                        'Consumo facturado
                        Aux = Aux & Format(Impor, "000000000000.00")
                        
                        
                        'Canon facturado.
                        If Not RT.EOF Then
                            If RT!numlinea = 25 Then Impor = Impor + RT!ImporteL
                        End If
                        
                        Aux = Aux & Format(Impor, "000000000000.00")
                        
                       'Quitamos las comas
                       Aux = Replace(Aux, ",", "")
                    End If
                    RT.Close
                    Cad = Cad & Aux
                    Print #NF, Cad
                    If Len(Cad) <> 342 Then Err.Raise 513, , "Linea<>342" & Cad
                            
                    
                    
                    davidNumalbar = davidNumalbar + 1
                    ImpoAnulado = ImpoAnulado + Impor 'del canon facturado emitido
                    TotalFraFRT = TotalFraFRT + 1

                    miRsAux.Close   'Es de la rectificada
            End If
            
            'Siguiente
            RR.MoveNext
            
            
            
            
        Wend
        RR.Close
        
        
        If davidCodtipom <> "" Then
            'Errores en clientes
            Aux = String(50, "*")
            Aux = Aux & vbCrLf & "Contador     Cliente      Factura    Cliente " & vbCrLf & Aux & vbCrLf
            Aux = Aux & davidCodtipom
            davidCodtipom = Aux
            ClientesIncorrectos
        End If
        
        
        'E)  Registro TOTAL del municipio
        
        
        Cad = "50" & CadenaFichero(Poblaciones(J), 5) & CadenaFichero(DBLet(pPdfRpt, "T"), 30) & Space(257)
        'Registros
        Aux = String(8, "0")
        Cad = Cad & Format(TotalFra, Aux) & Format(TotalFraFRT, Aux)
        'Importes
        Aux = Format(ImpoFacturado, "000000000000.0000") & Format(ImpoAnulado, "000000000000.0000")
        Aux = Replace(Aux, ",", "")
        Cad = Cad & Aux
        Print #NF, Cad
        davidNumalbar = davidNumalbar + 1
        
        Cad = "Poblacion: " & pPdfRpt & "(" & J & " de " & Poblaciones.Count & ")" & vbCrLf & vbCrLf
        Cad = Cad & "Tipo           Facturas       Canon" & vbCrLf
        Cad = Cad & String(30, "=") & vbCrLf
        Cad = Cad & "Emitidas        " & Right(Format(TotalFra, "0000000"), 7) & "       " & Format(ImpoFacturado, FormatoImporte) & vbCrLf
        Cad = Cad & "Amuladas      " & Right(Format(TotalFraFRT, "0000000"), 7) & "       " & Format(ImpoAnulado, FormatoImporte) & vbCrLf
        Cad = Cad & String(30, "=") & vbCrLf
        MsgBox Cad, vbInformation
        
    Next
    
    
    'F)  Registro TOTAL del soporte
    davidNumalbar = davidNumalbar + 1
    Aux = DevuelveDesdeBD(conAri, "CodEntSuministra", "sparamagua", "1", "1")
    Aux = CadenaFichero(Aux, 10)
    Cad = "60" & Aux & CadenaFichero(vParam.NombreEmpresa, 100) & CadenaFichero(vParam.CifEmpresa, 9)
    Cad = Cad & String(213, " ") & Format(davidNumalbar, "00000000")
   
    Print #NF, Cad

    GenerarFicheroPresentacionAgua = True
    
    
    
    
    
    
    
    
eGenerarFicheroPresentacionAgua:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set RR = Nothing
    Set RT = Nothing
    If NF > 0 Then Close #NF
    'Reestablezco las reutilizadas
    davidCodtipom = ""
    davidNumalbar = 0
    NumRegElim = 0
End Function


Private Function CadenaFichero(Valor As String, Posiciones As Integer) As String
    CadenaFichero = Mid(Valor & Space(Posiciones), 1, Posiciones)
End Function



Private Sub ImprimeRegistroFactura()

End Sub


Private Sub ClientesIncorrectos()
Dim NF2 As Integer
    On Error GoTo eClientesIncorrectos
    NF2 = FreeFile
    
    Open App.Path & "\ClientesIncorrectosAgua.txt" For Output As #NF2
    Print #NF2, davidCodtipom
    Close #NF2
    LanzaVisorMimeDocumento frmPpal.hWnd, App.Path & "\ClientesIncorrectosAgua.txt"
    Exit Sub
eClientesIncorrectos:
    MuestraError Err.Number, "Creando fichero Codigos cliente incorrectos"
End Sub
