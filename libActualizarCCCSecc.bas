Attribute VB_Name = "libActualizarCCCSecc"
Option Explicit





'Siempre vendranlas cuentas que entidad tenga valor
Public Function ComprobarDatosProcesoCCC(DesdeHasta As String, ByRef L As Label, ContabilidadAriagro As Boolean) As Boolean
    ComprobarDatosProcesoCCC = False
    If FaseCargaDatosAProcesarCCC(DesdeHasta) Then
        Fase2 L, ContabilidadAriagro
        DesdeHasta = DevuelveDesdeBD(conAri, "count(*)", "tmpinformes", "codusu", vUsu.codigo)
        If Val(DesdeHasta) > 0 Then ComprobarDatosProcesoCCC = True
    End If
End Function



Private Function FaseCargaDatosAProcesarCCC(DesdeHasta As String) As Boolean   'podemos reutilizar DesdeHasta
Dim Cad As String

    conn.Execute "DELETE from tmpcrmclien WHERE codusu = " & vUsu.codigo
   

    Set miRsAux = New ADODB.Recordset
    Cad = "select codclien,codmacta,nifclien,iban,codbanco,codsucur,digcontr,cuentaba from sclien"
    Cad = Cad & " WHERE " & DesdeHasta
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    DesdeHasta = "INSERT INTO tmpcrmclien(codusu,codclien,nomforpa,nomactiv) VALUES "
    Cad = ""
    NumRegElim = 0
    
    While Not miRsAux.EOF
        NumRegElim = NumRegElim + 1
        'Los 10 primeros (relleando a blancos sera el codmacta, los siguientes desde el 11 el NIF
        Cad = Cad & ", (" & vUsu.codigo & "," & miRsAux!codClien & ",'"
        Cad = Cad & Mid(miRsAux!Codmacta & "          ", 1, 10) & DBLet(miRsAux!NIFClien, "T") & "','"
        'Formateamos la cadena del banco
        Cad = Cad & FormatearCadenaBanco2(3, miRsAux)
        Cad = Cad & "')"
        miRsAux.MoveNext
        
        If miRsAux.EOF Then NumRegElim = 101
        
        If NumRegElim > 100 Then
            Cad = Mid(Cad, 2)
            Cad = DesdeHasta & Cad
            conn.Execute Cad
            Cad = ""
            NumRegElim = 1
        End If
    Wend
    miRsAux.Close
        
    FaseCargaDatosAProcesarCCC = NumRegElim > 0
    
End Function

Private Sub Fase2(ByRef L As Label, EnlasContabilidadesDeAriagro As Boolean)
Dim Cad As String
Dim TieneArigasol As Boolean
Dim I As Integer
Dim J As Integer
Dim RN As ADODB.Recordset
Dim RProv As ADODB.Recordset
Dim CuentasAtratar As String
Dim ColAriagro As Collection
Dim K As Integer
Dim Aux2 As String
Dim ContaGasol As Integer
Dim ContaAgro As Integer
Dim H As Integer
Dim Ariagro As String

Dim VinculaPorCodmacta As Byte   '0. Codmacta   1.- Codclien
Dim CodigosSocios As String  'en CATADAU linka por codigo socio -- codigo cliente

Dim RN3 As ADODB.Recordset

    'Veremos cual es el ariagro y cual el arigasol si es que los procesa
    TieneArigasol = False
    VinculaPorCodmacta = 0
    Ariagro = ""
    If vParamAplic.ComprobarBancoRestoAplicaciones Then
        Cad = "Select NumAriagro,TieneArigasol,linkaPorCodmacta from spara2 "
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            Ariagro = DBLet(miRsAux!NumAriagro, "T")
            TieneArigasol = DBLet(miRsAux!TieneArigasol, "N")
            VinculaPorCodmacta = DBLet(miRsAux!LinkaPorCodmacta, "N")  '0. Codmacta   1.- Codclien
        End If
        miRsAux.Close
    
    End If
    
    'Hay I registros. Los dividiremmos en grupos de 200 maximos
    L.Caption = "Obteniendo errores"
    L.Refresh
    Cad = DevuelveDesdeBD(conAri, "count(*)", "tmpcrmclien", "codusu", vUsu.codigo)
    I = Val(Cad)
    J = (I \ 200) + 1
     conn.Execute "DELETE from tmpinformes WHERE codusu = " & vUsu.codigo


    'Si tiene ARIAGRO, veremos las distintas secciones con sus contabilidades
    'asociadas
    Set ColAriagro = New Collection
    If Ariagro <> "" Then
            Cad = "select codsecci,empresa_conta from " & Ariagro & ".rseccion order by 2,1"
            
     
            miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            I = 0
            Cad = ""
            While Not miRsAux.EOF
                If I <> miRsAux!empresa_conta Then
                    If I > 0 Then ColAriagro.Add Cad & "|"
                    Cad = miRsAux!empresa_conta & "|"
                    I = miRsAux!empresa_conta
                Else
                    Cad = Cad & ","
                End If
                Cad = Cad & miRsAux!codsecci
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            If Cad <> "" Then ColAriagro.Add Cad & "|"
            
        
    End If 'ariagro

    'Si tiene arigasol. Conta a comprobar
    'Pudiera ser que tuviera gasol y la conta fuera la misma que ariges(ejem. CASTELDUC)
    If TieneArigasol Then
        ContaGasol = 0
        Cad = DevuelveDesdeBD(conAri, "numconta ", "arigasol.sparam", "1", "1")
        ContaGasol = CInt(Cad)
    End If
    
    'Cuentas a comprobar
    For I = 1 To J
        L.Caption = "Comprobar (" & I & "/" & J & ")"
        L.Refresh
        Cad = "select codclien,trim(substring(nomforpa,1,10)) LaCta,substring(nomforpa,11) ElNif,nomactiv CtaDelBanco from tmpcrmclien"
        Cad = Cad & " WHERE codusu = " & vUsu.codigo
        Cad = Cad & " order by 1 limit " & (I - 1) * 200 & ",200"
        
        Set RN = New ADODB.Recordset
        RN.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
        If Not RN.EOF Then
            'Primer question. Cargamos las cuentas a tratar
            'Clientes a tratar
            CuentasAtratar = ""
            CodigosSocios = ""
            While Not RN.EOF
                CuentasAtratar = CuentasAtratar & ", '" & RN!lacta & "'"
                CodigosSocios = CodigosSocios & ", " & RN!codClien
                RN.MoveNext
            Wend
            CuentasAtratar = Mid(CuentasAtratar, 2)
            CodigosSocios = Mid(CodigosSocios, 2)
            RN.MoveFirst
            
            
            
            'Comprobar en ariconta de ariges
            '--------------------------------
            L.Caption = I & "/" & J & " - Contab. ariges"
            L.Refresh
            Cad = "Select codmacta,nifdatos,iban,entidad,oficina,CC,cuentaba from conta" & vParamAplic.NumeroConta & ".cuentas WHERE codmacta"
            Cad = Cad & " IN (" & CuentasAtratar & ")"
            miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not RN.EOF
                Cad = ComprobarAricontaAriges2(RN)
                If Cad <> "" Then
                    'HA habido algun error
                    '                   cliente,                   '
                    '                           aplic   conta noEXISTE niferro  ctaerr
                    'tmpinformes(codusu,codigo1,campo1,caampo2,nombre1,nombre2,nombre3)
                    'la funcion retorna los valores para nom1,nom2,nom3 ya en formato SQL
                    
                    Cad = vUsu.codigo & "," & RN!codClien & ",1,'conta" & vParamAplic.NumeroConta & "'," & Cad
                    Cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,obser,nombre1,nombre2,nombre3) VALUES (" & Cad & ")"
                    conn.Execute Cad
                End If
                RN.MoveNext
            Wend
            miRsAux.Close
            RN.MoveFirst
            
            
            Set RN3 = New ADODB.Recordset
            If vParamAplic.ComprobarBancoRestoAplicaciones Then
                
                
            
                'HACE ARIGASOL
                'Comprobar en ARIGASOL
                '--------------------------------
                If TieneArigasol Then
                    L.Caption = I & "/" & J & " - GASOL"
                    L.Refresh
                    
                   
                    Cad = "Select codmacta,nifsocio,iban,codbanco,codsucur,digcontr,cuentaba from arigasol.ssocio WHERE codmacta"
                    Cad = Cad & " IN (" & CuentasAtratar & ")"
                        
                    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not miRsAux.EOF
                        Cad = ComprobarArigasol2(RN)
                        If Cad <> "" Then
                            'HA habido algun error
                            '                   cliente,                   '
                            '                           aplic   noEXISTE niferro  ctaerr
                            'tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,nombre3)
                            'la funcion retorna los valores para nom1,nom2,nom3 ya en formato SQL
                            
                            Cad = vUsu.codigo & "," & RN!codClien & ",2,'arigasol'," & Cad
                            Cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,obser,nombre1,nombre2,nombre3) VALUES (" & Cad & ")"
                            conn.Execute Cad
                        End If
                        miRsAux.MoveNext
                    Wend
                    miRsAux.Close
                    RN.MoveFirst
                    
                    'ARIGASOL conta. Si la contabilidad del arigasol es distinta a la del ariges tendra que mirar tb
                    'Los errores iran con el 3               #####
                    'Cad = vUsu.Codigo & "," & RN!codclien & ",3," & Cad
                    If ContaGasol <> vParamAplic.NumeroConta Then
                    
                    
                        L.Caption = I & "/" & J & " - Contab. gasol"
                        L.Refresh
                        Cad = "Select codmacta,nifdatos,iban,entidad,oficina,CC,cuentaba from conta" & ContaGasol & ".cuentas WHERE codmacta"
                        Cad = Cad & " IN (" & CuentasAtratar & ")"
                        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        
                        While Not RN.EOF
                            Cad = ComprobarAricontaAriges2(RN)
                            If Cad <> "" Then
                                'HA habido algun error
                                '                   cliente,                   '
                                '                           aplic   conta noEXISTE niferro  ctaerr
                                'tmpinformes(codusu,codigo1,campo1,caampo2,nombre1,nombre2,nombre3)
                                'la funcion retorna los valores para nom1,nom2,nom3 ya en formato SQL
                                
                                Cad = vUsu.codigo & "," & RN!codClien & ",3,'conta" & ContaGasol & "'," & Cad
                                Cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,obser,nombre1,nombre2,nombre3) VALUES (" & Cad & ")"
                                conn.Execute Cad
                            End If
                            RN.MoveNext
                        Wend
                        miRsAux.Close
                        RN.MoveFirst
                        
                    End If
                End If 'de tiene arigasol
                
                'HACE ARIAGRO
                'Para cada seccion mirara las contas
                If Ariagro <> "" Then
                    For K = 1 To ColAriagro.Count
                        '----------------------------------
                        
                        Cad = ColAriagro.item(K)
                        Aux2 = RecuperaValor(Cad, 1) 'la conta
                        ContaAgro = CInt(Aux2)
                        Aux2 = RecuperaValor(Cad, 2) 'las seccciones con esa conta
                        
                        L.Caption = I & "/" & J & " - AriAGRO (" & K & " de " & ColAriagro.Count & ")"
                        L.Refresh
                        
                        Cad = "select distinct  codmaccli,codmacpro,nifsocio,iban,codbanco,codsucur,digcontr,cuentaba,rsocios.codsocio from "
                        Cad = Cad & Ariagro & ".rsocios_seccion," & Ariagro & ".rsocios "
                        Cad = Cad & " where rsocios_seccion.codsocio=rsocios.codsocio AND codsecci"
                        Cad = Cad & " in (" & Aux2 & ") and "
                        
                        If VinculaPorCodmacta = 0 Then
                            'por codmacta
                            Cad = Cad & " codmaccli IN (" & CuentasAtratar & ")"
                        Else
                            
                            'por codssocio
                            Cad = Cad & " rsocios.codsocio IN (" & CodigosSocios & ")"
                        End If
                        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        Aux2 = "" 'Para ir a buscar a la contabilidad del ariagro
    
                        Set RProv = New ADODB.Recordset
                        'While Not miRsAux.EOF
                        RN.MoveFirst
                        While Not RN.EOF
                            L.Caption = I & "/" & J & "." & K & " - AriAGRO (" & RN!codClien & ")"
                            L.Refresh
                            Cad = ComprobarArigroSocio2(RN, VinculaPorCodmacta)
                            If Cad <> "" Then
                                'HA habido algun error
                                '                   cliente,                   '
                                '                           aplic   noEXISTE niferro  ctaerr
                                'tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,nombre3)
                                'la funcion retorna los valores para nom1,nom2,nom3 ya en formato SQL
                                If K = 1 Then
                                    Cad = vUsu.codigo & "," & RN!codClien & ",4,'" & Ariagro & "'," & Cad
                                    Cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,obser,nombre1,nombre2,nombre3) VALUES (" & Cad & ")"
                                    conn.Execute Cad
                                End If
                            End If
                             
                            
                            'Cuenta contable
                            If EnlasContabilidadesDeAriagro Then
                                If Not RN.EOF Then
                                    If miRsAux.EOF Then
                                        Cad = ""
                                    Else
                                        Cad = DBLet(miRsAux!CodMacCli, "T")
                                    End If
                                    If Cad <> "" Then
                                        Cad = "WHERE codmacta = '" & Cad & "'"
                                        Cad = "Select codmacta,nifdatos,iban,entidad,oficina,CC,cuentaba from conta" & ContaAgro & ".cuentas " & Cad
                                        
                                        RN3.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                        Cad = ComprobarAricontaArialgo2(RN3, RN, VinculaPorCodmacta)
                                        If Cad <> "" Then
                                            'HA habido algun error
                                            '                   cliente,                   '
                                            '                           aplic   noEXISTE niferro  ctaerr
                                            'tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,nombre3)
                                            'la funcion retorna los valores para nom1,nom2,nom3 ya en formato SQL
                                            
                                            Cad = vUsu.codigo & "," & RN!codClien & ",5,'conta" & ContaAgro & "'," & Cad
                                            Cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,obser,nombre1,nombre2,nombre3) VALUES (" & Cad & ")"
                                            conn.Execute Cad
                                        Else
                                            'Stop
                                        End If
                                        RN3.Close
                                    End If
                                End If
                                
                                'Si tiene proveedor haremos la comprobacion AHORA
                                If Not RN.EOF Then
                                    If miRsAux.EOF Then
                                        Cad = ""
                                    Else
                                        Cad = DBLet(miRsAux!CodMacPro, "T")
                                    End If
                                    If Cad <> "" Then
                                    
                                    
                                        Cad = "WHERE codmacta = '" & Cad & "'"
                                        Cad = "Select codmacta,nifdatos,iban,entidad,oficina,CC,cuentaba from conta" & ContaAgro & ".cuentas " & Cad
                                        
                                        RProv.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                                        Cad = ComprobarAricontaProveedor(RN, RProv)
                                        
                                        If Cad <> "" Then
                                            'HA habido algun error
                                            '                   cliente,                   '
                                            '                           aplic   noEXISTE niferro  ctaerr
                                            'tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,nombre3)
                                            'la funcion retorna los valores para nom1,nom2,nom3 ya en formato SQL
                                            
                                            Cad = vUsu.codigo & "," & RN!codClien & ",6,'conta" & ContaAgro & "'," & Cad
                                            Cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,obser,nombre1,nombre2,nombre3) VALUES (" & Cad & ")"
                                            conn.Execute Cad
                                        End If
                                        RProv.Close
                                        
                                    End If
                                End If
                            End If
                            
                            
                            'miRsAux.MoveNext
                            RN.MoveNext
                        Wend
                        miRsAux.Close
                        
                        Set RProv = Nothing
                        
                        
                       
                    Next K  'siguiente contaseccion
                    
                End If 'de tiene arigaro
            
            
            
            
            
            End If
        End If
        RN.Close
        Set RN3 = Nothing
    Next
    
End Sub





'Dado un RS, normalmente miRsaux, a partir de la poscion indicers estaran los 4 campos
' entidad,oficina,
Private Function FormatearCadenaBanco2(IndiceRS As Integer, ByRef RS As ADODB.Recordset) As String

    'OCUTBRE 2014
    '---->  Ponemos el iban antes de codmacta

    'Cadena = Cadena & Format(miRsAux!entidad, "0000") & Format(DBLet(miRsAux!oficina, "N"), "0000")
    'Cadena = Cadena & Right("00" & DBLet(miRsAux!CC, "T"), 2) & Right(String(10, "0") & DBLet(miRsAux!cuentaba, "T"), 10) & "')"
    
    FormatearCadenaBanco2 = Right("    " & DBLet(RS.Fields(IndiceRS), "N"), 4) & Format(DBLet(RS.Fields(IndiceRS + 1), "N"), "0000") & Format(DBLet(RS.Fields(IndiceRS + 2), "N"), "0000")
    FormatearCadenaBanco2 = FormatearCadenaBanco2 & Right("00" & DBLet(RS.Fields(IndiceRS + 3), "T"), 2) & Right(String(10, "0") & DBLet(RS.Fields(IndiceRS + 4), "T"), 10)
    

End Function



Private Function ComprobarAricontaAriges2(ByRef RsOrigen As ADODB.Recordset) As String
Dim C As String
Dim Ok As Boolean
    
    C = " codmacta = '" & RsOrigen!lacta & "'"
    miRsAux.Find C, , adSearchForward, 1
    If miRsAux.EOF Then
        ComprobarAricontaAriges2 = "'NO existe la cuenta',NULL,NULL"
    Else
        Ok = True
        C = FormatearCadenaBanco2(2, miRsAux)
        If C <> RsOrigen!CtaDelBanco Then
            C = "NULL,'" & C & "',"
            Ok = False
        Else
            C = "NULL,NULL,"
        End If
        
        If DBLet(miRsAux!nifdatos, "T") <> RsOrigen!ElNif Then
            Ok = False
            C = C & "'" & DBLet(miRsAux!nifdatos, "T") & "'"
        Else
            C = C & "NULL"
        End If
        If Ok Then
            C = ""
        Else
            ComprobarAricontaAriges2 = C
        End If
    End If
    
End Function




Private Function ComprobarArigasol2(ByRef rsb As ADODB.Recordset) As String
Dim C As String
Dim Ok As Boolean
    
    
    
    C = " LaCta = '" & miRsAux!Codmacta & "'"
    rsb.Find C, , adSearchForward, 1
    If rsb.EOF Then
        ComprobarArigasol2 = "'NULL','NULL','NULL'"  'No deberia pasar
    Else
        Ok = True
        C = FormatearCadenaBanco2(2, miRsAux)
        If C <> rsb!CtaDelBanco Then
            C = "NULL,'" & C & "',"
            Ok = False
        Else
            C = "NULL,NULL,"
        End If
        
        If DBLet(miRsAux!nifsocio, "T") <> rsb!ElNif Then
            Ok = False
            C = C & "'" & DBLet(miRsAux!nifsocio, "T") & "'"
        Else
            C = C & "NULL"
        End If
        If Ok Then
            C = ""
        Else
            ComprobarArigasol2 = C
        End If
    End If
    
End Function




Private Function ComprobarArigroSocio2(ByRef rsb As ADODB.Recordset, ComoLinkaBD As Byte) As String
Dim C As String
Dim Ok As Boolean
    
    '0codmacta      1codcliensoc
    If ComoLinkaBD = 0 Then
        'codmacta
        C = " codmaccli = '" & rsb!lacta & "'"
    Else

        C = " codsocio= " & rsb!codClien
    End If
    'rsb.Find C, , adSearchForward, 1
    miRsAux.Find C, , adSearchForward, 1
    If miRsAux.EOF Then
        ComprobarArigroSocio2 = "'NOEXIS','NULL','NULL'"  'No deberia pasar
        Ok = False
    Else
        Ok = True
        C = FormatearCadenaBanco2(3, miRsAux)
        If C <> rsb!CtaDelBanco Then
            C = "NULL,'" & C & "',"
            Ok = False
        Else
            C = "NULL,NULL,"
        End If
        
        If DBLet(miRsAux!nifsocio, "T") <> rsb!ElNif Then
            Ok = False
            C = C & "'" & DBLet(miRsAux!nifsocio, "T") & "'"
        Else
            C = C & "NULL"
        End If
        If Ok Then
            C = ""
        Else
            ComprobarArigroSocio2 = C
        End If
    End If
    
End Function



Private Function ComprobarAricontaArialgo2(ByRef RsOrigen As ADODB.Recordset, ByRef RD As ADODB.Recordset, VinculaPorElCodmacta As Byte) As String
Dim C As String
Dim C2 As String
Dim Ok As Boolean
    
    'C = " laCta = '" & miRsAux!Codmacta & "'"   NO hace find pq el SQL es codmacta ='VALOR'
    'RsOrigen.Find C, , adSearchForward, 1
    If RsOrigen.EOF Then
        ComprobarAricontaArialgo2 = "'NO existe la cuenta',NULL,NULL"
    Else
        Ok = True
        
        '0. Codmacta   1.- Codclien Con lo cual, aqui me quiero grabar la cuenta para updatear
        If VinculaPorElCodmacta = 1 Then
            C2 = "'" & RsOrigen!Codmacta & "',"
        Else
            C2 = "NULL,"
        End If
        
        C = FormatearCadenaBanco2(2, RsOrigen)
        If C <> RD!CtaDelBanco Then
            C = C2 & "'" & C & "',"
            Ok = False
        Else
            C = C2 & "NULL,"
        End If
        
        If DBLet(RsOrigen!nifdatos, "T") <> RD!ElNif Then
            Ok = False
            If IsNull(RsOrigen!nifdatos) Then
                C = C & "'VACIO" & "'"
                '"
            Else
                C = C & "'" & DBLet(RsOrigen!nifdatos, "T") & "'"
            End If
        Else
            C = C & "NULL"
        End If
        If Ok Then
            C = ""
        Else
            ComprobarAricontaArialgo2 = C
        End If
    End If
    
End Function



Private Function ComprobarAricontaProveedor(ByRef RsOrigen As ADODB.Recordset, ByRef RDest As ADODB.Recordset) As String
Dim C As String
Dim Ok As Boolean
     
   
    If RDest.EOF Then
        'Si hubiera que mostrar el error, descomentar el trozo este
        'C = Mid(RDest.Source, InStr(RDest.Source, "codmacta = '") + 12)
        'C = Mid(C, 1, Len(C) - 1)
        'ComprobarAricontaProveedor = "'NO existe" & C & "',NULL,NULL"
        
        
        'NO existe la cta de proveedor.
        'Como NO la vamos a crear, no hace falta que mostremos el error
        ComprobarAricontaProveedor = ""
        
    Else
        Ok = True
        C = FormatearCadenaBanco2(2, RDest)
        If C <> RsOrigen!CtaDelBanco Then
            'La cuenta del proveedor la guardo
            C = "'" & C & "',"
            Ok = False
        Else
            C = "NULL,"
        End If
        C = "'" & RDest!Codmacta & "'," & C 'SIEMPRE metemos el codmactaprov
        If DBLet(RDest!nifdatos, "T") <> RsOrigen!ElNif Then
            Ok = False
            If IsNull(RDest!nifdatos) Then
                C = C & "'nulo'"
            Else
                C = C & "'" & RDest!nifdatos & "'"
            End If
        Else
            C = C & "NULL"
        End If
        If Ok Then
            C = ""
        Else
            ComprobarAricontaProveedor = C
        End If
    End If
    
End Function












