Attribute VB_Name = "libCoarval"
Option Explicit

Dim CadenaInsert As String
Dim Msg As String
'Dim SerieDeFactu  As String


Dim DeParametrosTiposIVA As String
Dim DeparametrosSeriesFactura As String
Dim DeparametrosArticuloNuevo As String  'codtipar,codmarca,codigiva,codfamia,codprove,codunida



Public Function ProcesaFicheroClientesCOARVAL(Fichero As String, ByRef LB As Label) As Byte
Dim NF As Integer
Dim OK As Boolean
Dim linea As String
Dim Seguir As Boolean
Dim RA As ADODB.Recordset
Dim PrimLinea As Boolean
Dim J As Integer
Dim PirmeraLineaEncabezados As Boolean
Dim SerieDeFactu As String
On Error GoTo eProcesaFicheroClientes

    ProcesaFicheroClientesCOARVAL = 2 'NADA no procesa nada

    linea = "select * from sparamcoarval "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open linea, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    DeParametrosTiposIVA = miRsAux!TiposIva
    DeparametrosSeriesFactura = miRsAux!SeriesFactura
    linea = miRsAux!DefectoArticulo    'codtipar,codmarca,codigiva,codfamia,codprove,codunida'
    'QUito el ulimo pipe
    If Right(linea, 1) = "|" Then linea = Mid(linea, 1, Len(linea) - 1)
    linea = Replace(linea, "|", "','")
    linea = "'" & linea & "'"
    DeparametrosArticuloNuevo = linea
    

    LB.Caption = "Leyendo csv"
    LB.Refresh
    
    'Preparamos tabla de insercion para ver cuantas facturas o si hay errores...
    conn.Execute "DELETE FROM tmpintegracoarval WHERE codusu = " & vUsu.Codigo
    conn.Execute "DELETE FROM tmpcrmclien WHERE codusu = " & vUsu.Codigo

    'insert into `tmpintegracoarval` (`codusu`,`numserie`,`numfactu`,`fechaalt`,`base`,`total`,`base_sr`,`iva_sr`,`re_sr`,`total_sr`,`base_red`,`iva_red`,`re_red`,`total_red`,`base_norm`,`iva_norm`,`re_nor`,`total_nor`,`codclien`,`nomclien`,`domclien`,`codpobla`,`pobclien`,`proclien`,`nifclien`,`forpa`,`codartic`,`nomartic`,`PorcenIVA`,`ampliaci`,`cantidad`,`precioar`,`dtoline1`,`importel`)

    
    CadenaInsert = ""
    NF = FreeFile
    OK = False
    Open Fichero For Input As #NF
    Seguir = Not EOF(NF)
    PrimLinea = True
    Msg = ""
    While Seguir
        Line Input #NF, linea
         
        
        If PrimLinea Then
            J = InStr(1, linea, ";")
            If J > 0 Then
                SerieDeFactu = Trim(Mid(linea, 1, J - 1))
                If SerieDeFactu = "" Then SerieDeFactu = "N"
                If Not IsNumeric(SerieDeFactu) Then
                   PirmeraLineaEncabezados = True
                Else
                    PirmeraLineaEncabezados = False
                End If
            End If
            If PirmeraLineaEncabezados Then
                    Msg = "N" 'para no procesar la linea
                    OK = True
            End If
            SerieDeFactu = ""
            PrimLinea = False
        End If
        If Msg = "" Then OK = ProcesarLineaAsiento(linea)
        If Not OK Then
            Seguir = False
        Else
            Seguir = Not EOF(NF)
            Msg = ""
        End If
    Wend
    Close (NF)
    If CadenaInsert <> "" Then
        CadenaInsert = Mid(CadenaInsert, 2)
        SerieDeFactu = DevuelevInsert
        CadenaInsert = SerieDeFactu & CadenaInsert
        conn.Execute CadenaInsert
    End If
    Espera 0.5
    
    
    
    
    If OK Then
    
        
    
        LB.Caption = "Comprobando datos IVAS"
        LB.Refresh
        
       
                        
                
        pRptvMultiInforme = 0  'para saber si hay erroeres, cuantos y se utiliza para el insert en errores
                

                
        linea = "select distinct porceniva from tmpintegracoarval where codusu = " & vUsu.Codigo & " ORDER BY 1"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open linea, conn, adOpenKeyset, adLockOptimistic, adCmdText
        
        NF = 0
        While Not miRsAux.EOF
           NF = NF + 1
           If NF > 4 Then InsertaError "Mas de 4 tipos de IVA"
                
           If InStr(1, DeParametrosTiposIVA, "|" & miRsAux.Fields(0) & "|") = 0 Then InsertaError "IVA no tratado "
            
                       
           miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        '**********************************************
        'Los IVAS, noraml y con recargo DEBEN estar configurados
        '*******************************************************
        
        
        LB.Caption = "Comprobando datos clientes"
        LB.Refresh
        
        linea = "select tmpintegracoarval.codclien, tmpintegracoarval.nifclien n1 , sclien.nifclien n2 from tmpintegracoarval,sclien where codusu = " & vUsu.Codigo & "  AND tmpintegracoarval.codclien = sclien.codclien"
        miRsAux.Open linea, conn, adOpenKeyset, adLockOptimistic, adCmdText
        NF = 0
        While Not miRsAux.EOF
            If DBLet(miRsAux!n1, "T") <> DBLet(miRsAux!N2, "T") Then
                'ERROR NIFS distintos
                'InsertaError "Cliente: " & miRsAux!codClien & " NIFs distintos :" & miRsAux!n1 & "  --  " & miRsAux!N2
                'FALTA###
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        
        
        linea = "select * FROM tmpintegracoarval where codusu = " & vUsu.Codigo & "  AND NOT tmpintegracoarval.codclien IN (select codclien from sclien) GROUP BY codclien"
        miRsAux.Open linea, conn, adOpenKeyset, adLockOptimistic, adCmdText
        NF = 0
        If Not miRsAux.EOF Then
            
            'Forma de pago defecto
            Msg = DevuelveDesdeBD(conAri, "codforpa", "sforpa", "1", "1 ORDER BY tipforpa,codforpa")
            'Cuenta por defecto select * from sclien where clivario=1
            SerieDeFactu = DevuelveDesdeBD(conAri, "codmacta", "sclien", "1", "1 order by clivario desc, codclien")
            
            
            While Not miRsAux.EOF
                linea = "INSERT INTO sclien (pasclien,nifclien,clivario,fechaalt,visitador,particular,periodof,referobl,promocio"
                linea = linea & ",codtarif,numrepet,albarcon,tipofact,codagent,EnvFraEmail,iban,Rentin_x_dpto,AplicaPortesFactura,enviocorreo,"
                linea = linea & "tasareciclado,tipoiva,cliabono,dtognral,dtoppago,diavtoat,diapago1,codmacta,"
                linea = linea & "cuentaba,digcontr,codsucur,codbanco,mesnogir,diapago2,diapago3,codforpa,"
                linea = linea & "codenvio,codactiv,codrutas,codzonas,perclie2,telclie2,faxclie2,maiclie2,maiclie1,"
                linea = linea & "faxclie1,telclie1,perclie1,observac,wwwclien,proclien,pobclien,codpobla,domclien,tipclien,"
                linea = linea & "kilometr,fechamov,codsitua,nomcomer,nomclien,codclien) VALUES ("
        
                
                'pasclien,nifclien,clivario,fechaalt,visitador,particular,periodof,referobl,promocio
                linea = linea & DBSet(miRsAux!nifClien, "T") & "," & DBSet(miRsAux!nifClien, "T") & ",0," & DBSet(Now, "F") & "," & vParamAplic.PorDefecto_Agente & ",0,0,0,0,"
                ',codtarif,numrepet,albarcon,tipofact,codagent,EnvFraEmail,
                linea = linea & vParamAplic.PorDefecto_Tarifa & ",1,1,1," & vParamAplic.PorDefecto_Agente & ",0"
                'iban,Rentin_x_dpto,AplicaPortesFactura,enviocorreo, tasareciclado,tipoiva,
                linea = linea & ",null,null,0,0,0,0,"
                'cliabono,dtognral,dtoppago,diavtoat,diapago1,codmacta,cuentaba,digcontr,codsucur,codbanco,mesnogir,diapago2,diapago3,
                linea = linea & "0,0,0,0,0," & DBSet(SerieDeFactu, "T") & ",null,null,null,null,null,null,null,"
                'codforpa,codenvio,codactiv,codrutas
                linea = linea & Msg & "," & vParamAplic.PorDefecto_Envio & "," & vParamAplic.PorDefecto_Activ & "," & vParamAplic.PorDefecto_Ruta
                ',codzonas,perclie2,telclie2,faxclie2,maiclie2,maiclie1,faxclie1,telclie1,perclie1,observac,wwwclien,
                linea = linea & "," & vParamAplic.PorDefecto_Zona & ",null,null,null,null,null,null,null,null,null,null,"
                'proclien,pobclien,codpobla
                linea = linea & DBSet(miRsAux!proclien, "T", "N") & "," & DBSet(miRsAux!pobclien, "T", "N") & "," & DBSet(miRsAux!codpobla, "T", "N") & ","
                ',domclien,tipclien,kilometr,fechamov,codsitua,nomcomer
                linea = linea & DBSet(miRsAux!domclien, "T", "N") & ",0,0,null,0," & DBSet(miRsAux!NomClien, "T", "N") & ","
                ',nomclien,codclien)'
                linea = linea & DBSet(miRsAux!NomClien, "T", "N") & "," & DBSet(miRsAux!codClien, "N") & ")"
                 
                conn.Execute linea
            
                miRsAux.MoveNext
            Wend
        End If
        miRsAux.Close
        
                
        LB.Caption = "Forma de pago"
        LB.Refresh
        linea = "select distinct forpa from tmpintegracoarval  where codusu = " & vUsu.Codigo
        miRsAux.Open linea, conn, adOpenKeyset, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
                NF = 1
                If DBLet(miRsAux!ForPa, "T") <> "" Then
                    linea = DevuelveDesdeBD(conAri, "codforpa", "sforpa", "nomforpa", miRsAux!ForPa, "T")
                    If linea <> "" Then NF = 0
                End If
                If NF = 1 Then
                    'ERROR NIFS distintos
                    InsertaError "Forma de pago: " & miRsAux!ForPa
                End If
           
                miRsAux.MoveNext
        Wend
        miRsAux.Close
        
               
        LB.Caption = "Serie factura"
        LB.Refresh
        linea = "select numserie from tmpintegracoarval  where codusu = " & vUsu.Codigo
        miRsAux.Open linea, conn, adOpenKeyset, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
                
                
                
                linea = "|" & miRsAux!numSerie & "#"
                NF = InStr(1, DeparametrosSeriesFactura, linea)
                If NF = 0 Then
                    InsertaError "Nº Serie incorrecto"
                Else
                     NumRegElim = InStr(NF + 1, DeparametrosSeriesFactura, "|")
                     If NumRegElim = 0 Then
                        InsertaError "Nº Serie incorrecto"
                     Else
                        NF = Len(miRsAux!numSerie) + NF + 2
                        linea = Mid(DeparametrosSeriesFactura, NF, NumRegElim - NF)
                        
                        linea = "UPDATE  tmpintegracoarval set numserie='" & linea & "' where codusu=" & vUsu.Codigo & " and numserie=" & DBSet(miRsAux!numSerie, "T")
                        conn.Execute linea
                    End If
                End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
               
        'Comprobamos articulos
        LB.Caption = "Comprobar articulos"
        LB.Refresh
        linea = "select  codartic,nomartic from tmpintegracoarval  where codusu = " & vUsu.Codigo
        miRsAux.Open linea, conn, adOpenKeyset, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
                
                NF = Len(Trim(miRsAux!codArtic))
                
                If NF = 0 Then
                    InsertaError "Articulo vacio"
                ElseIf NF > 16 Then
                    InsertaError "Longitud incorrecta articulo"
                Else
                    ComprobarArticulo
                End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    
        
               
               
               
               
               
                       
                
        'Si llega a aqui, vamos a generar las facturas
        Set RA = New ADODB.Recordset
        NumRegElim = 0
  
        LB.Caption = "Comprobar numero factura"
        LB.Refresh
        
        linea = "select numserie,numfactu,fechaalt from `tmpintegracoarval` where codusu=" & vUsu.Codigo & " group by 1,2,3"
        miRsAux.Open linea, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        linea = ""
        NF = 1
        While Not miRsAux.EOF
            linea = linea & ", (" & DBSet(miRsAux!numSerie, "T") & "," & miRsAux!Numfactu & "," & Year(miRsAux!fechaalt) & ")"
            miRsAux.MoveNext
            If miRsAux.EOF Then
                NF = 11
            Else
                NF = NF + 1
            End If
            
            If NF > 10 Then
                linea = "(" & Mid(linea, 2) & ")"
                Msg = "Select codtipom,numfactu,fecfactu from scafac where (codtipom,numfactu,year(fecfactu)) in " & linea
                RA.Open Msg, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RA.EOF
                    linea = "YA existe factura: " & RA!codtipom & " " & RA!Numfactu & " " & RA!FecFactu
                    InsertaError linea
                    RA.MoveNext
                Wend
                RA.Close
                linea = ""
                NF = 0
            End If
        Wend
        miRsAux.Close
        
        
        LB.Caption = "Comprobar totales"
        LB.Refresh
        
        linea = "select numserie,numfactu,fechaalt,sum(importel) from tmpintegracoarval where codusu =" & vUsu.Codigo & " group by numserie,numfactu,fechaalt"
        linea = linea & " ORDER BY numserie,numfactu"
        miRsAux.Open linea, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        linea = "select numserie,numfactu,fechaalt,round(base,2),total from tmpintegracoarval where codusu  =" & vUsu.Codigo & "  group by numserie,numfactu,fechaalt"
        linea = linea & " ORDER BY numserie,numfactu"
        RA.Open linea, conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        
        While Not miRsAux.EOF
            RA.MoveFirst
            NF = 0
            While NF = 0
                If RA!numSerie = miRsAux!numSerie Then
                    If miRsAux!Numfactu = RA!Numfactu Then
                        NF = 2
                        
                        If RA.Fields(3) <> miRsAux.Fields(3) Then
                            
                            If Abs(RA.Fields(3) - miRsAux.Fields(3)) > 1 Then
                                linea = "base imponible factura. " & RA.Fields(0) & RA.Fields(1) & " :" & RA.Fields(3) & "  // " & miRsAux.Fields(3)
                                InsertaError linea
                            End If
                        End If
                        
                    End If
                End If
                If NF = 0 Then
                    RA.MoveNext
                    If RA.EOF Then NF = 1
            
                End If
            Wend
            'If NF = 1 Then St op
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        RA.Close
    End If
    
            
    'para el año de la factura, NO existen ya en contabilidad
    LB.Caption = "Comprobando facturas"
    LB.Refresh
    If OK Then
        If pRptvMultiInforme = 0 Then
           ProcesaFicheroClientesCOARVAL = 0  'TODO BIEN
        Else
             ProcesaFicheroClientesCOARVAL = 1  'Duplicados
        End If
    Else
        ProcesaFicheroClientesCOARVAL = 2
    End If


    
        
    
eProcesaFicheroClientes:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
       
    End If
    pRptvMultiInforme = 0
    Set miRsAux = Nothing
    Set RA = Nothing
End Function






Private Function ProcesarLineaAsiento(linea As String) As Boolean
Dim strArray() As String
Dim Aux As String
Dim Cad As String
Dim N As Integer

    On Error GoTo EProcesarLineaAsientO
    
    ProcesarLineaAsiento = False
    
    linea = Replace(linea, """", "")
    strArray = Split(linea, ";")
            
            
    If UBound(strArray) + 1 <> 31 Then
        Aux = "Campos en fichero: " & UBound(strArray) + 1 & "       Campos para procesar: 31"
        MsgBox Aux, vbExclamation
        Exit Function
    End If
    
    'insert into `tmpintegracoarval` (`codusu`,`numserie`,`numfactu`,`fechaalt`,`base`,`total`,`base_sr`,`iva_sr`,`re_sr`,`total_sr`,`base_red`,`iva_red`,`re_red`,`total_red`,`base_norm`,`iva_norm`,`re_nor`,`total_nor`,`codclien`,`nomclien`,`domclien`,`codpobla`,`pobclien`,`proclien`,`nifclien`,`forpa`,`codartic`,`nomartic`,`PorcenIVA`,`ampliaci`,`cantidad`,`precioar`,`dtoline1`,`importel`)
    
    CadenaInsert = CadenaInsert & ", (" & vUsu.Codigo & ","
    
    'Numero factura  00010100023733  ---->> SERIE: 101   nº: 00023733
    
    Aux = Val(Mid(strArray(0), 4, 3))
    Cad = Val(Mid(strArray(0), 7))
    If Val(Aux) = 0 Or Val(Cad) = 0 Then Err.Raise 513, , "Error en serie-factura: " & strArray(0)
    CadenaInsert = CadenaInsert & DBSet(Aux, "T") & "," & DBSet(Cad, "N") & ","
    
    N = 1
    Aux = strArray(N)
    If Not IsDate(Aux) Then Err.Raise 513, , "Error en fecha: " & strArray(N)
    CadenaInsert = CadenaInsert & DBSet(Aux, "F") & ","
    
    ' `base`,`total`
    N = 2
    Aux = strArray(N)
    Cad = strArray(N + 1)
    If Not IsNumeric(Aux) Then Err.Raise 513, , "Error en campo numerico: " & strArray(N)
    If Not IsNumeric(Cad) Then Err.Raise 513, , "Error en campo numerico: " & strArray(N + 1)
    CadenaInsert = CadenaInsert & DBSet(CCur(Aux), "N") & "," & DBSet(CCur(Cad), "N") & ","
    
    'SUPER-REDUCIDO
    ',`base_sr`,`iva_sr`,`re_sr`,`total_sr`
    N = 4
    Aux = strArray(N)
    Cad = strArray(N + 1)
    If Aux = "" Xor Cad = "" Then Err.Raise 513, , "Error en iva S-R: " & strArray(N) & strArray(N + 1)
    If Cad <> "" Then
        If Not IsNumeric(Aux) Then Err.Raise 513, , "Error en campo numerico: " & strArray(N)
        If Not IsNumeric(Cad) Then Err.Raise 513, , "Error en campo numerico: " & strArray(N + 1)
        
        CadenaInsert = CadenaInsert & DBSet(CCur(Aux), "N") & "," & DBSet(CCur(Cad), "N") & ","
        N = 6
        Aux = strArray(N)
        Cad = strArray(N + 1)
        
        CadenaInsert = CadenaInsert & DBSet(CCur(Aux), "N", "S") & "," & DBSet(CCur(Cad), "N") & ","
    
    Else
        CadenaInsert = CadenaInsert & "null,null,null,null,"
    
    End If
    
    'REDUCIDO
    ',`base_red`,`iva_red`,`re_red`,`total_red`,
    N = 8
    Aux = strArray(N)
    Cad = strArray(N + 1)
    If Aux = "" Xor Cad = "" Then Err.Raise 513, , "Error en iva reducido: " & strArray(N) & strArray(N + 1)
    If Cad <> "" Then
        If Not IsNumeric(Aux) Then Err.Raise 513, , "Error en campo numerico: " & strArray(N)
        If Not IsNumeric(Cad) Then Err.Raise 513, , "Error en campo numerico: " & strArray(N + 1)
        
        CadenaInsert = CadenaInsert & DBSet(CCur(Aux), "N") & "," & DBSet(CCur(Cad), "N") & ","
        N = 10
        Aux = strArray(N)
        Cad = strArray(N + 1)
        
        CadenaInsert = CadenaInsert & DBSet(CCur(Aux), "N", "S") & "," & DBSet(CCur(Cad), "N") & ","
    
    Else
        CadenaInsert = CadenaInsert & "null,null,null,null,"
    
    End If
    
    'IVA NORMAL
    '`base_norm`,`iva_norm`,`re_nor`,`total_nor`
    N = 12
    Aux = strArray(N)
    Cad = strArray(N + 1)
    If Aux = "" Xor Cad = "" Then Err.Raise 513, , "Error en iva normal: " & strArray(N) & strArray(N + 1)
    If Cad <> "" Then
        If Not IsNumeric(Aux) Then Err.Raise 513, , "Error en campo numerico: " & strArray(N)
        If Not IsNumeric(Cad) Then Err.Raise 513, , "Error en campo numerico: " & strArray(N + 1)
        
        CadenaInsert = CadenaInsert & DBSet(CCur(Aux), "N") & "," & DBSet(CCur(Cad), "N") & ","
        N = 14
        Aux = strArray(N)
        Cad = strArray(N + 1)
        
        CadenaInsert = CadenaInsert & DBSet(CCur(Aux), "N", "S") & "," & DBSet(CCur(Cad), "N") & ","
    
    Else
        CadenaInsert = CadenaInsert & "null,null,null,null,"
    
    End If
    
    ',`codclien`,`nomclien`,`domclien`,`codpobla`,`pobclien`,`proclien`,`nifclien`
    For N = 16 To 22
        Cad = Trim(strArray(N))
        If N <= 17 Or N = 22 Then
            'CAMPO OBLIGADO
            If Cad = "" Then
                Err.Raise 513, , "Campo en cliente obligado(Codigo-Nombre-NIF)"
            Else
                If N = 16 Then If Not IsNumeric(Cad) Then Err.Raise 513, , "Campo en codigo cliente numerico: " & strArray(N)
            End If
        End If
        CadenaInsert = CadenaInsert & DBSet(Cad, "T", "S") & ","
    Next
    
    
    
    ',`forpa`,`codartic`,`nomartic,ampliaci`
    For N = 23 To 25
        Cad = Trim(strArray(N))
        If Cad = "" Then Err.Raise 513, , "Campo  obligado: " & RecuperaValor("Forma pago|Art|Desc.articulo|", N - 22)
        CadenaInsert = CadenaInsert & DBSet(Cad, "T", "S") & ","
    Next N
    
    
    ',`PorcenIVA`,`cantidad`,`precioar`,`dtoline1`,`importel`)
    For N = 26 To 30
        Cad = Trim(strArray(N))
        If Cad = "" Then
            Err.Raise 513, , "Campo  obligado: " & RecuperaValor("porcenIVa|cantidad|Precio|dto|importe lin|", N - 26)
        Else
            If Not IsNumeric(Cad) Then Err.Raise 513, , "Campo  numerico: " & RecuperaValor("porcenIVa|cantidad|Precio|dto|importe lin|", N - 26) & Cad
        End If
        CadenaInsert = CadenaInsert & DBSet(CCur(Cad), "N", "S")
        If N <> 30 Then CadenaInsert = CadenaInsert & ","
    Next N
    
    CadenaInsert = CadenaInsert & ")"
            
    If Len(CadenaInsert) > 8000 Then
        CadenaInsert = Mid(CadenaInsert, 2)
        Cad = DevuelevInsert
        Cad = Cad & CadenaInsert
        conn.Execute Cad
        CadenaInsert = ""
    End If

    
    ProcesarLineaAsiento = True
    Exit Function
EProcesarLineaAsientO:
    MuestraError Err.Number, Err.Description
End Function




Private Function DevuelevInsert() As String
        
        DevuelevInsert = "INSERT INTO `tmpintegracoarval` (`codusu`,`numserie`,`numfactu`,`fechaalt`,`base`,`total`"
        DevuelevInsert = DevuelevInsert & ",`base_sr`,`iva_sr`,`re_sr`,`total_sr`,`base_red`,`iva_red`,`re_red`,`total_red`,"
        DevuelevInsert = DevuelevInsert & "`base_norm`,`iva_norm`,`re_nor`,`total_nor`,"
        DevuelevInsert = DevuelevInsert & "`codclien`,`nomclien`,`nifclien`,`domclien`,`codpobla`,`pobclien`,`proclien`"
        DevuelevInsert = DevuelevInsert & ",`forpa`,`codartic`,`nomartic`,`PorcenIVA`,`cantidad`,`precioar`,`dtoline1`,`importel`) VALUES "
        
        
End Function


Private Function CrearFacturaClientes() As Boolean
    CrearFacturaClientes = False
End Function


Private Function ComprobarCuentasContables() As Boolean
ComprobarCuentasContables = False
End Function

Private Sub InsertaError(CadenaError As String)
    pRptvMultiInforme = pRptvMultiInforme + 1
    conn.Execute "INSERT INTO tmpcrmclien(codusu,codclien,auxiliar) VALUES (" & vUsu.Codigo & "," & pRptvMultiInforme & "," & DBSet(CadenaError, "T") & ")"
End Sub

Private Function ComprobarNumerosDeFactura() As Boolean
ComprobarNumerosDeFactura = False
End Function



Private Function ComprobarArticulo() As Boolean
    On Error GoTo eComprobarArticulo
    ComprobarArticulo = False
    Msg = DevuelveDesdeBD(conAri, "codartic", "sartic", "codartic", miRsAux!codArtic, "T")
    If Msg = "" Then
        Msg = DevuelveDesdeBD(conAri, "max(numorden)", "sartic", "1", "1")
        Msg = Val(Msg) + 1
        'NUEVO
        CadenaInsert = "INSERT INTO sartic (numorden,fecaltas,preciove,rotacion,mateprima,ctrstock,codstatu,garantia,unicajas,conjunto,artvario,codartic,nomartic,"
        CadenaInsert = CadenaInsert & "codtipar,codmarca,codigiva,codfamia,codprove,codunida) VALUES (" & Msg & "," & DBSet(Now, "F") & ","
        CadenaInsert = CadenaInsert & "0,0,0,0,0,0,1,0,0,"
        CadenaInsert = CadenaInsert & DBSet(miRsAux!codArtic, "T") & "," & DBSet(miRsAux!NomArtic, "T") & ","
        'DeparametrosArticuloNuevo --> codtipar,codmarca,codigiva,codfamia,codprove,codunida
        CadenaInsert = CadenaInsert & DeparametrosArticuloNuevo & ")"
        conn.Execute CadenaInsert
        
        'Y EN SALMAC
        CadenaInsert = "insert into salmac(codartic,codalmac,canstock,statusin) VALUES (" & DBSet(miRsAux!codArtic, "T") & ",1,0,0)"
        conn.Execute CadenaInsert
        
    End If
    ComprobarArticulo = True
    Exit Function
eComprobarArticulo:
    InsertaError Err.Description
End Function

