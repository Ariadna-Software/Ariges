Attribute VB_Name = "libGesSocial"
Option Explicit




Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim SQL As String

'rUdNegocio
'TEndra la entrada en la BD de unidades de negocio
'    CStr(empresa_conta)= para la conta y para ariges
'Solo para nuevos. La coje del formulario de asociados
Public Function TraspasaAsociadoAriges(IdAsoc As Long, ByRef rsUdNegocio As ADODB.Recordset, FechaDeAlta As Date) As Boolean
    Dim yaExiste_ As Boolean
    Dim Auxiliar2 As String
    Dim Codmacta As String
    Dim LaConta As String
    
        'Tendra las observaciones si no es nuevo
        'Si es nuevo tendra los valores por defecto de spara1 para envio, zoan.....
    
    
    TraspasaAsociadoAriges = False
    
    Set RS = New ADODB.Recordset
    
    '-- comprobamos si esta asociado ya existía como cliente al otro lado
    SQL = "select * from ariges" & rsUdNegocio!empresa_conta & ".sclien where codclien = " & IdAsoc
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    yaExiste_ = False
    If Not RS.EOF Then
        yaExiste_ = True
        Auxiliar2 = DBLet(RS!observac, "T")
    End If
    RS.Close


    'NO existe veo los valores por defecto para
    'defenvio,defzona,defruta,defagente,
    If Not yaExiste_ Then
        SQL = "Select defenvio,defzona,defruta,defagente from  ariges" & rsUdNegocio!empresa_conta & ".spara1"
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE SER EOF
        Auxiliar2 = RS!defenvio & "|" & RS!defzona & "|" & RS!defruta & "|" & RS!defagente & "|"
        RS.Close
            
    End If
        
    '-- Buscamos los datos del asociado
    SQL = "select * from asociados where IdAsoc = " & IdAsoc
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        
        'Establezco la cuenta para contabilidad
        If RS!essocio Then
            Codmacta = rsUdNegocio!raiz_cliente_socio & Format(RS!CodSocEuroagro, "00000")
        Else
            Codmacta = rsUdNegocio!raiz_cliente_asociado & Format(RS!IdAsoc, "00000")
        End If
        
        
        'MAYO 2014.
        'Observaciones de sclien. Las quito ya que en el gesoccial antiguo
        'las guardaba en una variable pero no las lleva a la BD
        
        If Not yaExiste_ Then
                'NUEVO
            SQL = "INSERT INTO ariges" & rsUdNegocio!empresa_conta & ".sclien (codclien, nomclien, nomcomer, domclien, codpobla"
            SQL = SQL & " ,pobclien, proclien, nifclien,   fechaalt, codactiv"
            SQL = SQL & " ,telclie1, faxclie1, maiclie1,  telclie2, faxclie2, "
            SQL = SQL & "  iban, codbanco, codsucur, digcontr, cuentaba, codmacta, codtarif "
            SQL = SQL & " ,codenvio, codzonas, codrutas, codagent,visitador, codforpa, diapago1"
            SQL = SQL & " ,clivario, tipoiva, tipofact, albarcon, periodof, numrepet,"
            SQL = SQL & " dtoppago, dtognral, promocio, codsitua, referobl,cliabono,pasclien,credipriv"
            SQL = SQL & ") values ("
            SQL = SQL & IdAsoc & ","
            SQL = SQL & DBSet(RS!nomlargo, "T") & ","
            SQL = SQL & DBSet(RS!nomlargo, "T") & ","
            SQL = SQL & DBSet(RS!Direccion, "T") & ","
            SQL = SQL & DBSet(RS!CodPostal, "N") & ","
            SQL = SQL & DBSet(RS!Poblacion, "T") & ","
            SQL = SQL & DBSet(RS!Provincia, "T") & ","
            SQL = SQL & DBSet(RS!NIF, "T") & ","
            
            'Nov 2014
            SQL = SQL & DBSet(FechaDeAlta, "F") & ","
            
            
            'Antes de junio 14
            '
            'Codigo actividad
            'If RS!essocio Then
            '    SQL = SQL & "1"
            'Else
            '    SQL = SQL & "2"
            'End If
            SQL = SQL & RS!tarifaprecio
            
            SQL = SQL & "," & DBSet(RS!Telefono1, "T") & ","
            SQL = SQL & DBSet(RS!Movil, "T") & ","
            SQL = SQL & DBSet(RS!mail, "T") & ","
            SQL = SQL & DBSet(RS!Telefono2, "T") & "," & DBSet(RS!Telefono3, "T") & ","
            
            'iban, codbanco, codsucur, digcontr, cuentaba,codmacta
            SQL = SQL & DBSet(RS!Iban, "T", "S") & ","
            SQL = SQL & DBLet(RS!entidad, "N") & ","
            SQL = SQL & DBLet(RS!Sucursal, "N") & ","
            SQL = SQL & DBSet(RS!DC, "T") & ","
            SQL = SQL & DBSet(RS!NumCC, "T") & ","
            'Codmacta
            SQL = SQL & DBSet(Codmacta, "T")
            
            'Junio 2014
            'Tarifaprecio es ACTIVIDAD
            ' codactiv = rs!tarifaprecio
            SQL = SQL & ",1,"
            
            
                            
            'Auxiliar2 = rs!defenvio & "|" & rs!defzona & "|" & rs!defruta & "|" & rs!defagente & "|"
            SQL = SQL & RecuperaValor(Auxiliar2, 1) & ","
            SQL = SQL & RecuperaValor(Auxiliar2, 2) & ","
            SQL = SQL & RecuperaValor(Auxiliar2, 3) & ","
            SQL = SQL & RecuperaValor(Auxiliar2, 4) & ","
            'Visitador. Lo mismo que codagent
            SQL = SQL & RecuperaValor(Auxiliar2, 4) & ","
        
        
            'Codforpa
            SQL = SQL & rsUdNegocio!ForPa & ","
            
            
            
            'Diapago1, clivario  tipoiva, tipofact, albarcon, periodof, numrepet
            SQL = SQL & "10,0,0,0,0,1,1,"
        
            'dtoppago, dtognral, promocio, codsitua, referobl,  cliabono pasclien"
            'tarifaprecio
            SQL = SQL & "0,0,1,0,0,"
            If RS!tarifaprecio = 1 Then
                SQL = SQL & "0"
            Else
                SQL = SQL & "1"
            End If
            
            
            SQL = SQL & "," & DBSet(RS!NIF, "T") & ",9"   '9. SIN asegurar (credipriv)
            SQL = SQL & ")"
            
        Else
            'MODIFICAR
            
            'codclien, nomclien, nomcomer, domclien, codpobla"
            'pobclien, proclien, nifclien,
            'telclie1, faxclie1, maiclie1,  telclie2, faxclie2, "
            ' iban, codbanco, codsucur, digcontr, cuentaba, codmacta, observac "
            
            
            SQL = "UPDATE ariges" & rsUdNegocio!empresa_conta & ".sclien SET "
            SQL = SQL & " nomclien = " & DBSet(RS!nomlargo, "T")
            SQL = SQL & ", nomcomer = " & DBSet(RS!nomlargo, "T")
            SQL = SQL & ", domclien = " & DBSet(RS!Direccion, "T")
            SQL = SQL & ", codpobla = " & DBSet(RS!CodPostal, "N")
            SQL = SQL & ", pobclien = " & DBSet(RS!Poblacion, "T")
            SQL = SQL & ", proclien = " & DBSet(RS!Provincia, "T")
            SQL = SQL & ", nifclien = " & DBSet(RS!NIF, "T")
            
            SQL = SQL & ", telclie1 = " & DBSet(RS!Telefono1, "T")
            SQL = SQL & ", faxclie1 = " & DBSet(RS!Movil, "T")
            SQL = SQL & ", maiclie1 = " & DBSet(RS!mail, "T")
            SQL = SQL & ", telclie2 = " & DBSet(RS!Telefono2, "T")
            SQL = SQL & ", faxclie2 = " & DBSet(RS!Telefono3, "T")

            SQL = SQL & ", iban = " & DBSet(RS!Iban, "T")
            SQL = SQL & ", codbanco = " & DBSet(RS!entidad, "N")
            SQL = SQL & ", codsucur = " & DBSet(RS!Sucursal, "N")
            SQL = SQL & ", digcontr = " & DBSet(RS!DC, "T")
            SQL = SQL & ", cuentaba = " & DBSet(RS!NumCC, "T")
                
            'Antes JUNIO 2014
            'SQL = SQL & ", codtarif = " & RS!tarifaprecio
            SQL = SQL & ", codactiv = " & RS!tarifaprecio
            
            'Cuent alternativa
            SQL = SQL & ", cliabono = "
            
            If RS!tarifaprecio = 1 Then
                SQL = SQL & "0"
            Else
                SQL = SQL & "1"
            End If
            
            
            SQL = SQL & " WHERE codclien =" & IdAsoc
            
        End If
        
            
        If ejecutar(SQL, False) Then
            TraspasaAsociadoAriges = True
        
            
            'Actualizamos datos en contabilidad
            
            LaConta = DevuelveDesdeBD(conAri, "empresa_conta", "unidadesnegocio", "IdUnidad", rsUdNegocio!IdUnidad)
            If vParamAplic.ContabilidadNueva Then
                SQL = DevuelveDesdeBD(conAri, "codmacta", "ariconta" & LaConta & ".cuentas", "codmacta", Codmacta)
            Else
                SQL = DevuelveDesdeBD(conAri, "codmacta", "conta" & LaConta & ".cuentas", "codmacta", Codmacta)
            End If
            If SQL = "" Then
                'No existe la cuenta. La creo
                ActualizarLaCuenta LaConta, Codmacta, RS, vParamAplic.ContabilidadNueva
                Espera 0.2
            End If
            
            'Para no pasar muchas variables , como lo que estoy enviando NO es arigasol
            'Le digo que arigasol es la 127 y ya esta
            ActualizaCuentasAsociado IdAsoc, rsUdNegocio!IdUnidad, 127
        End If
            
        
    End If
    
    Set RS = Nothing
End Function



'QueSeccion (o unidad de negocio)
'   0.- TODAS
'   'Cualquier otro sera su UD de negocio
Public Function ActualizaCuentasAsociado(IdAsoc As Long, QueSeccion As Byte, QueUDEsGasolinera As Byte) As Boolean
Dim rUd As ADODB.Recordset
Dim Codmacta As String

Dim UltimoNivel As Byte
Dim i As Byte

    Set RS = New ADODB.Recordset
    Set rUd = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    SQL = "select unidadesnegocio.* from asociados_unidadesnegocio,unidadesnegocio where "
    SQL = SQL & " asociados_unidadesnegocio.IdUnidad= unidadesnegocio.idunidad and idasoc=" & CStr(IdAsoc)
    If QueSeccion > 0 Then SQL = SQL & " AND unidadesnegocio.IdUnidad = " & QueSeccion
    
    rUd.Open SQL & " order by empresa_conta", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not rUd.EOF Then
        SQL = "select * from asociados where IdAsoc = " & CStr(IdAsoc)
        RS.Open SQL, conn, adOpenForwardOnly
        If Not RS.EOF Then
    
    
            While Not rUd.EOF
            'Datos asociado
                    
                    If vParamAplic.ContabilidadNueva Then
                        SQL = "Select * from ariconta" & rUd!empresa_conta & ".empresa"
                    Else
                        SQL = "Select * from conta" & rUd!empresa_conta & ".empresa"
                    End If
                    rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    'NO PUEDE SER EOF
                    i = rs2!numnivel
                    UltimoNivel = rs2.Fields("numdigi" & CStr(i))
                    rs2.Close
                    
                    
                    
                    'Para la gasolinera siempre cojera IdASOC
                    If rUd!IdUnidad = QueUDEsGasolinera Then
                        i = UltimoNivel - Len(rUd!raiz_cliente_asociado)
                        Codmacta = String(CLng(i), "0")
                         
                        Codmacta = rUd!raiz_cliente_asociado & Format(IdAsoc, Codmacta)
                        
                        ActualizarLaCuenta CStr(rUd!empresa_conta), Codmacta, RS, vParamAplic.ContabilidadNueva
                       
                        conn.Execute "update asociados set codmacta = '" & Codmacta & "' where IdAsoc = " & CStr(IdAsoc)
                    Else
                        'Pueden ser varias cuentas a actualizar
                        If rUd!raiz_cliente_socio <> "" And RS!essocio = 1 Then
                            '
                             i = UltimoNivel - Len(rUd!raiz_cliente_socio)
                             Codmacta = String(CLng(i), "0")
                             
                             Codmacta = rUd!raiz_cliente_socio & Format(RS!CodSocEuroagro, Codmacta)
                             
                             ActualizarLaCuenta CStr(rUd!empresa_conta), Codmacta, RS, vParamAplic.ContabilidadNueva
                                                          
                        End If
                                                
                        If rUd!raiz_cliente_asociado <> "" And RS!essocio = 0 Then
                            i = UltimoNivel - Len(rUd!raiz_cliente_asociado)
                            Codmacta = String(CLng(i), "0")
                             
                            Codmacta = rUd!raiz_cliente_asociado & Format(IdAsoc, Codmacta)
                            
                            ActualizarLaCuenta CStr(rUd!empresa_conta), Codmacta, RS, vParamAplic.ContabilidadNueva
                            
                            
                            
                        End If
                        
                End If
                        
                        
                If rUd!raiz_proveedor <> "" Then
                    i = UltimoNivel - Len(rUd!raiz_proveedor)
                    Codmacta = String(CLng(i), "0")
                    
                    If RS!essocio = 1 Then
                        Codmacta = rUd!raiz_proveedor & Format(RS!CodSocEuroagro, Codmacta)
                    Else
                        Codmacta = rUd!raiz_proveedor & Format(IdAsoc, Codmacta)
                    End If
                    ActualizarLaCuenta CStr(rUd!empresa_conta), Codmacta, RS, vParamAplic.ContabilidadNueva
                End If
                        

                    
                rUd.MoveNext
            Wend
            
        
        End If
        RS.Close
    End If
    rUd.Close
    Set rUd = Nothing
    Set RS = Nothing
    Set rs2 = Nothing
End Function


Private Sub ActualizarLaCuenta(Contabilidad As String, Cuenta As String, ByRef vRs As ADODB.Recordset, NuevaContabilidad As Boolean)
Dim SQL As String
Dim Iban As String

        If Not NuevaContabilidad Then
            'Conta antigua
            SQL = DevuelveDesdeBD(conAri, "codmacta", "conta" & Contabilidad & ".cuentas", "codmacta", Cuenta)
            If SQL = "" Then
                'NUEVO
                SQL = "INSERT INTO conta" & Contabilidad & ".cuentas(codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,"
                SQL = SQL & "desprovi,nifdatos,maidatos,obsdatos,pais,entidad,oficina,CC,cuentaba,iban) VALUES ('"
                SQL = SQL & Cuenta & "'," & DBSet(vRs!nomlargo, "T") & ",'S',1," & DBSet(vRs!nomlargo, "T") & ","
                SQL = SQL & DBSet(vRs!Direccion, "T") & "," & DBSet(vRs!CodPostal, "T") & "," & DBSet(vRs!Poblacion, "T") & ","
                SQL = SQL & DBSet(vRs!Provincia, "T") & "," & DBSet(vRs!NIF, "T") & "," & DBSet(vRs!mail, "T") & ","
                SQL = SQL & DBSet(vRs!Observaciones, "T") & ",'ESPAÑA'," & DBSet(vRs!entidad, "N") & "," & DBSet(vRs!Sucursal, "N")
                SQL = SQL & "," & DBSet(vRs!DC, "T") & "," & DBSet(vRs!NumCC, "T") & "," & DBSet(vRs!Iban, "T") & ") "
            
            
            Else
                            
                '(codmacta,nommacta razosoci,dirdatos,codposta,despobla,"
                'desprovi,nifdatos,maidatos,obsdatos,pais,entidad,oficina,CC,cuentaba,iban
                
                'UPDATEAR
                SQL = "UPDATE conta" & Contabilidad & ".cuentas SET  nommacta = " & DBSet(vRs!nomlargo, "T")
                SQL = SQL & ", razosoci = " & DBSet(vRs!nomlargo, "T")
                SQL = SQL & ", dirdatos = " & DBSet(vRs!Direccion, "T")
                SQL = SQL & ", codposta = " & DBSet(vRs!CodPostal, "N")
                SQL = SQL & ", despobla = " & DBSet(vRs!Poblacion, "T")
                SQL = SQL & ", desprovi = " & DBSet(vRs!Provincia, "T")
                SQL = SQL & ", nifdatos = " & DBSet(vRs!NIF, "T")
                
                SQL = SQL & ", maidatos = " & DBSet(vRs!Telefono1, "T")
                SQL = SQL & ", obsdatos = " & DBSet(vRs!Movil, "T")
    
                SQL = SQL & ", iban = " & DBSet(vRs!Iban, "T")
                SQL = SQL & ", entidad = " & DBSet(vRs!entidad, "N")
                SQL = SQL & ", oficina = " & DBSet(vRs!Sucursal, "N")
                SQL = SQL & ", CC = " & DBSet(vRs!DC, "T")
                SQL = SQL & ", cuentaba = " & DBSet(vRs!NumCC, "T")
                SQL = SQL & " WHERE codmacta = " & DBSet(Cuenta, "T")
             End If
        Else
            'Nueva contabulidada!!!!!
            'Conta antigua
            If DBLet(vRs!entidad, "N") = 0 Or DBLet(vRs!Sucursal) = 0 Then
                Iban = ""
            Else
                Iban = DBLet(vRs!Iban, "T") & Format(DBLet(vRs!entidad, "N"), "0000") & Format(DBLet(vRs!Sucursal, "N"), "0000")
                Iban = Iban & Right("00" & DBLet(vRs!DC, "T"), 2) & Right(String(10, "0") & DBLet(vRs!NumCC, "T"), 10)
            End If
            
            
            
            SQL = DevuelveDesdeBD(conAri, "codmacta", "ariconta" & Contabilidad & ".cuentas", "codmacta", Cuenta)
            If SQL = "" Then
                
            
            
                'NUEVO
                SQL = "INSERT INTO ariconta" & Contabilidad & ".cuentas(codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,"
                SQL = SQL & "desprovi,nifdatos,maidatos,obsdatos,codpais,iban) VALUES ('"
                SQL = SQL & Cuenta & "'," & DBSet(vRs!nomlargo, "T") & ",'S',1," & DBSet(vRs!nomlargo, "T") & ","
                SQL = SQL & DBSet(vRs!Direccion, "T") & "," & DBSet(vRs!CodPostal, "T") & "," & DBSet(vRs!Poblacion, "T") & ","
                SQL = SQL & DBSet(vRs!Provincia, "T") & "," & DBSet(vRs!NIF, "T") & "," & DBSet(vRs!mail, "T") & ","
                SQL = SQL & DBSet(vRs!Observaciones, "T") & ",'ES'," & DBSet(Iban, "T", "S") & ") "
            
            
            Else
                            
                '(codmacta,nommacta razosoci,dirdatos,codposta,despobla,"
                'desprovi,nifdatos,maidatos,obsdatos,pais,entidad,oficina,CC,cuentaba,iban
                
                'UPDATEAR
                SQL = "UPDATE ariconta" & Contabilidad & ".cuentas SET  nommacta = " & DBSet(vRs!nomlargo, "T")
                SQL = SQL & ", razosoci = " & DBSet(vRs!nomlargo, "T")
                SQL = SQL & ", dirdatos = " & DBSet(vRs!Direccion, "T")
                SQL = SQL & ", codposta = " & DBSet(vRs!CodPostal, "N")
                SQL = SQL & ", despobla = " & DBSet(vRs!Poblacion, "T")
                SQL = SQL & ", desprovi = " & DBSet(vRs!Provincia, "T")
                SQL = SQL & ", nifdatos = " & DBSet(vRs!NIF, "T")
                
                SQL = SQL & ", maidatos = " & DBSet(vRs!Telefono1, "T")
                SQL = SQL & ", obsdatos = " & DBSet(vRs!Movil, "T")
    
                SQL = SQL & ", iban = " & DBSet(Iban, "T")
                
                SQL = SQL & " WHERE codmacta = " & DBSet(Cuenta, "T")
             End If
         
         
        End If
        conn.Execute SQL

End Sub






Public Function ActualizaSocioAriagro(IdAsoc As Long) As Boolean
    '-- Montamos el bucle de lectura de todos los asociados / socios
    Dim i As Long
    Dim rs3 As ADODB.Recordset
    Dim CodMacCli As String
    Dim CodMacPro As String
    
    ActualizaSocioAriagro = False
    
    SQL = "select * from asociados"
    SQL = SQL & " where IdAsoc = " & CStr(IdAsoc)
    SQL = SQL & " and (fechabaja is null)"
    SQL = SQL & " and CodSocEuroagro < 10000"
    SQL = SQL & " and EsSocio = 1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        '-- Lo primero es montar el nuevo socio
        SQL = DevuelveDesdeBD(conAri, "codsocio", "ariagro.rsocios", "codsocio", RS!CodSocEuroagro)
        
        If SQL = "" Then
        

            'NUEVO  NUEVO    NUEVO
            '-- NO existe y se da de alta
            SQL = "insert into ariagro.rsocios (codsocio, nifsocio, nomsocio, dirsocio, pobsocio, prosocio," & _
                    "codpostal, fechanac, telsoci1, telsoci2, telsoci3, movsocio, maisocio," & _
                    "codcoope, IBAN,codbanco, codsucur, digcontr, cuentaba, observaciones," & _
                    "fechaalta, fechabaja, correo, codiva, tipoirpf, tipoprod, codsitua) VALUES ("
            'odsocio, nifsocio, nomsocio, dirsocio, pobsocio, prosocio
            SQL = SQL & DBSet(RS!CodSocEuroagro, "N") & ","
            SQL = SQL & DBSet(RS!NIF, "T") & ","
            SQL = SQL & DBSet(RS!nomlargo, "T") & ","
            SQL = SQL & DBSet(RS!Direccion, "T") & ","
            SQL = SQL & DBSet(RS!Poblacion, "T") & ","
            SQL = SQL & DBSet(RS!Provincia, "T") & ","
    
            'codpostal, fechanac, telsoci1, telsoci2, telsoci3, movsocio, maisocio
            SQL = SQL & DBSet(RS!CodPostal, "N") & ","
            SQL = SQL & DBSet(RS!FechaNac, "F") & ","
            SQL = SQL & DBSet(RS!Telefono1, "T") & ","
            SQL = SQL & DBSet(RS!Telefono2, "T") & ","
            SQL = SQL & DBSet(RS!Telefono3, "T") & ","
            SQL = SQL & DBSet(RS!Movil, "T") & ","
            SQL = SQL & DBSet(RS!mail, "T") & ","
            
            'codcoope, IBAN,codbanco, codsucur, digcontr, cuentaba, observaciones
            SQL = SQL & "1,"   '-- Ponemos la cooperativa 1 a capón (OJO)
            SQL = SQL & DBSet(RS!Iban, "T") & ","
            SQL = SQL & DBSet(RS!entidad, "N") & ","
            SQL = SQL & DBSet(RS!Sucursal, "N") & ","
            SQL = SQL & DBSet(RS!DC, "T") & ","
            SQL = SQL & DBSet(RS!NumCC, "T") & ","
            SQL = SQL & "'Gesocial: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & "',"
            
            'fechaalta, fechabaja, correo, codiva, tipoirpf, tipoprod, codsitua
            SQL = SQL & DBSet(RS!fechaalta, "F") & ","
            SQL = SQL & DBSet(RS!fechabaja, "F", "S") & ","
            SQL = SQL & DBSet(RS!Correo, "T") & ","
            'Enero 2019. Codigiva
            SQL = SQL & DBSet(RS!codiva, "N") & ","
            SQL = SQL & DBSet(RS!TipoIrpf, "N") & ","
            SQL = SQL & "0,1)"
                
                
            
            
        Else
            'ACUTALIZAR
        
           
            ' nifsocio, nomsocio, dirsocio, pobsocio, prosocio
            SQL = "UPDATE ariagro.rsocios SET "
            SQL = SQL & " nifsocio = " & DBSet(RS!NIF, "T") & ","
            SQL = SQL & " nomsocio = " & DBSet(RS!nomlargo, "T") & ","
            SQL = SQL & " dirsocio = " & DBSet(RS!Direccion, "T") & ","
            SQL = SQL & " pobsocio = " & DBSet(RS!Poblacion, "T") & ","
            SQL = SQL & " prosocio = " & DBSet(RS!Provincia, "T") & ","
            
            
            'codpostal, fechanac, telsoci1, telsoci2, telsoci3, movsocio, maisocio
            SQL = SQL & " codpostal = " & DBSet(RS!CodPostal, "N") & ","
            SQL = SQL & " fechanac = " & DBSet(RS!FechaNac, "F") & ","
            SQL = SQL & " telsoci1 = " & DBSet(RS!Telefono1, "T") & ","
            SQL = SQL & " telsoci2 = " & DBSet(RS!Telefono2, "T") & ","
            SQL = SQL & " telsoci3 = " & DBSet(RS!Telefono3, "T") & ","
            SQL = SQL & " movsocio = " & DBSet(RS!Movil, "T") & ","
            SQL = SQL & " maisocio = " & DBSet(RS!mail, "T") & ","
            
            ' IBAN,codbanco, codsucur, digcontr, cuentaba, observaciones
            SQL = SQL & " IBAN = " & DBSet(RS!Iban, "T") & ","
            SQL = SQL & " codbanco = " & DBSet(RS!entidad, "N") & ","
            SQL = SQL & " codsucur = " & DBSet(RS!Sucursal, "N") & ","
            SQL = SQL & " digcontr = " & DBSet(RS!DC, "T") & ","
            SQL = SQL & " cuentaba= " & DBSet(RS!NumCC, "T") & ","

            
            'fechaalta, fechabaja, correo, codiva,
            SQL = SQL & " fechaalta = " & DBSet(RS!fechaalta, "F") & ","
            'Sql = Sql & " fechabaja = " & DBSet(RS!fechabaja, "F", "S") & ","
            SQL = SQL & " correo = " & DBSet(RS!Correo, "N") & ","
            SQL = SQL & " codiva = " & DBSet(RS!codiva, "N") & ","
            SQL = SQL & " TipoIrpf = " & DBSet(RS!TipoIrpf, "N")
            
            SQL = SQL & " WHERE codsocio = " & DBSet(RS!CodSocEuroagro, "N")
            
        End If
            
        conn.Execute SQL
            
        'En esta unidad de
            
            
            
            
        
        '-- Leemos su relación con unidades de negocio
        'sql = "select * from unidadesnegocio where IdAsoc = " & rs!IdAsoc
        SQL = "select * from unidadesnegocio where idunidad = 3" 'ARIAGRO
        Set rs2 = New ADODB.Recordset
        Set rs3 = New ADODB.Recordset
        rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not rs2.EOF
    
            If vParamAplic.ContabilidadNueva Then
                SQL = "Select * from ariconta" & rs2!empresa_conta & ".empresa"
            Else
                SQL = "Select * from conta" & rs2!empresa_conta & ".empresa"
            End If
            rs3.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            'NO PUEDE SER EOF
            i = rs3!numnivel
            i = rs3.Fields("numdigi" & CStr(i))
            rs3.Close
           
            
            CodMacCli = ""
            If DBLet(rs2!raiz_cliente_asociado, "T") <> "" Then CodMacCli = rs2!raiz_cliente_asociado & Format(RS!CodSocEuroagro, String(i - Len(rs2!raiz_cliente_asociado), "0"))
            CodMacPro = ""
            If DBLet(rs2!raiz_proveedor, "T") <> "" Then CodMacPro = rs2!raiz_proveedor & Format(RS!CodSocEuroagro, String(i - Len(rs2!raiz_proveedor), "0"))

            
            SQL = " codsocio = " & RS!CodSocEuroagro & " and codsecci "
            SQL = DevuelveDesdeBD(conAri, "codsecci", "ariagro.rsocios_seccion", SQL, rs2!CodSeccEuroagro)
             '-- NO existe y se da de alta
            If SQL = "" Then
                SQL = "insert into ariagro.rsocios_seccion(codsocio, codsecci, fecalta, fecbaja," & _
                        "codmaccli, codmacpro, codiva) VALUES ("
                SQL = SQL & RS!CodSocEuroagro & "," & rs2!CodSeccEuroagro & ","
                SQL = SQL & DBSet(RS!fechaalta, "F") & ","
                SQL = SQL & DBSet(RS!fechabaja, "F", "S") & ","
                SQL = SQL & DBSet(CodMacCli, "T", "S") & ","
                SQL = SQL & DBSet(CodMacPro, "T", "S") & ","
                SQL = SQL & RS!codiva & ")"
            Else
                '-- Si existe y se modifica
                SQL = "update ariagro.rsocios_seccion set "
                SQL = SQL & "fecalta = " & DBSet(RS!fechaalta, "F") & ","
                SQL = SQL & "fecbaja = " & DBSet(RS!fechabaja, "F", "S") & ","
                SQL = SQL & "codmaccli = " & DBSet(CodMacCli, "T", "S") & ","
                SQL = SQL & "codmacpro = " & DBSet(CodMacPro, "T", "S")
                SQL = SQL & " where codsocio = " & RS!CodSocEuroagro
                SQL = SQL & " and codsecci = " & rs2!CodSeccEuroagro
            End If
            conn.Execute SQL
           
            'LAs cremos en contabilidad
            If CodMacCli <> "" Then ActualizarLaCuenta CStr(rs2!empresa_conta), CodMacCli, RS, vParamAplic.ContabilidadNueva
            If CodMacPro <> "" Then ActualizarLaCuenta CStr(rs2!empresa_conta), CodMacPro, RS, vParamAplic.ContabilidadNueva
            rs2.MoveNext
        Wend
        rs2.Close
        
        Set rs2 = Nothing
        Set rs3 = Nothing
        
        ActualizaSocioAriagro = True
        
    End If
    RS.Close
    Set RS = Nothing
End Function






'FechaAltaSeccion
'Solo se utiliza en el NUEVO.
'Es la que tiene en asociados_unidadesnegocio
Public Function ActGasolineraAsociadoSocio(IdAsoc As Long, IdEntidadCoop As Integer, FechaAltaSeccion As Date, ElArigaso As String) As Boolean
'Dim mAux As String
Dim TipoConta As Byte
    
    '-- Primero buscamos al asociado en GesSocial para obtener sus datos
    Set RS = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    SQL = "select * from asociados where IdAsoc = " & CStr(IdAsoc)
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
    
        ''Octubre 2014.
        'Tipconta=1 SI, y solo si, Essocio=1 ; FechaBaja= Nulo ; tarifa precio = 1
        TipoConta = 0
        If DBSet(RS!essocio, "N") = 1 Then
            If IsNull(RS!fechabaja) Then
                If DBLet(RS!tarifaprecio, "N") = 1 Then TipoConta = 1
            End If
        End If
    
    
        '-- Ahora miramos si el asociado ya existe en la aplicación de gasolinera
        SQL = "select * from " & ElArigaso & ".ssocio where codsocio = " & CStr(IdAsoc)
        rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If rs2.EOF Then
            '-- No existe y hay que darlo de alta
            '   Para darlo de alta hay que conocer la entidad a la que pertenece
            '   y comprobar que está dada de alta en la gasolinera.
            '
        
            
            '-- Y ahora ya podemos seguir dando de alta al socio
            SQL = "insert into " & ElArigaso & ".ssocio (codsocio,codcoope,nomsocio,domsocio,codposta," & _
                    "pobsocio,prosocio,nifsocio,telsocio,movsocio,maisocio," & _
                    "fechaalt,codtarif,codbanco,codsucur,digcontr,cuentaba,iban," & _
                    "impfactu,dtolitro,codforpa,tipsocio,bonifbas,codsitua,codmacta,obssocio,tipconta)"
            SQL = SQL & " values("
            SQL = SQL & IdAsoc & "," & IdEntidadCoop & ","
            SQL = SQL & DBSet(RS!nomlargo, "T") & ","
            SQL = SQL & DBSet(RS!Direccion, "T") & ","
            SQL = SQL & DBSet(RS!CodPostal, "T") & ","
            SQL = SQL & DBSet(RS!Poblacion, "T") & ","
            SQL = SQL & DBSet(RS!Provincia, "T") & ","
            SQL = SQL & DBSet(RS!NIF, "T") & ","
            SQL = SQL & DBSet(RS!Telefono1, "T") & ","
            SQL = SQL & DBSet(RS!Movil, "T") & ","
            SQL = SQL & DBSet(RS!mail, "T") & ","
            SQL = SQL & DBSet(FechaAltaSeccion, "F") & ","
            SQL = SQL & "1,"  ' por defecto la tarifa es la 1
            SQL = SQL & DBSet(Right("0000" & DBLet(RS!entidad, "T"), 4), "T") & ","
            SQL = SQL & DBSet(Right("0000" & DBLet(RS!Sucursal, "T"), 4), "T") & ","
            SQL = SQL & DBSet(RS!DC, "T") & ","
            SQL = SQL & DBSet(RS!NumCC, "T") & ","
            SQL = SQL & DBSet(RS!Iban, "T") & ","
            
            'SQL = SQL & gasDB.numero(0) & "," ' impfactu
            'SQL = SQL & gasDB.numero(0) & "," ' dtolitro
            'SQL = SQL & gasDB.numero(0) & "," ' codforpa
            'SQL = SQL & gasDB.numero(0) & "," ' tipsocio
            'SQL = SQL & gasDB.numero(0) & "," ' bonifbas
            'SQL = SQL & gasDB.numero(0) & "," ' codsitua
            SQL = SQL & "0,0,0,0,0,0,"


            SQL = SQL & DBSet(RS!Codmacta, "T") & ","
            SQL = SQL & DBSet(RS!Observaciones, "T") & ","
            
            
            SQL = SQL & TipoConta & ")"
            
        Else
          '  SQL = "select * from asociados_entidades where IdAsoc = " & IdAsoc
          '  Set Rs2 = gesDB.cursor(SQL)
          '  If Not Rs2.EOF Then
          '      IdEntidad = Rs2!IdEntidad
          '      ActGasolineraEntidadesColectivos IdEntidad, gesDB, gasDB
          '  End If
            '-- Ya exite, lo modificamos, aunque hay muchos campos que no se tocan
            SQL = "update " & ElArigaso & ".ssocio set "
            'FALTA###
            SQL = SQL & "codcoope=" & IdEntidadCoop & ","
            SQL = SQL & "nomsocio=" & DBSet(RS!nomlargo, "T") & ","
            SQL = SQL & "domsocio=" & DBSet(RS!Direccion, "T") & ","
            SQL = SQL & "codposta=" & DBSet(RS!CodPostal, "T") & ","
            SQL = SQL & "pobsocio=" & DBSet(RS!Poblacion, "T") & ","
            SQL = SQL & "prosocio=" & DBSet(RS!Provincia, "T") & ","
            
            SQL = SQL & "codbanco=" & DBSet(Right("0000" & DBLet(RS!entidad, "T"), 4), "T") & ","
            SQL = SQL & "codsucur=" & DBSet(Right("0000" & DBLet(RS!Sucursal, "T"), 4), "T") & ","
            SQL = SQL & "digcontr=" & DBSet(RS!DC, "T") & ","
            SQL = SQL & "cuentaba=" & DBSet(RS!NumCC, "T") & ","
            SQL = SQL & "nifsocio=" & DBSet(RS!NIF, "T") & ","
            SQL = SQL & "telsocio=" & DBSet(RS!Telefono1, "T") & ","
            SQL = SQL & "movsocio=" & DBSet(RS!Movil, "T") & ","
            SQL = SQL & "maisocio=" & DBSet(RS!mail, "T") & ","
            SQL = SQL & "codmacta=" & DBSet(RS!Codmacta, "T") & ","
            SQL = SQL & "obssocio=" & DBSet(RS!Observaciones, "T", "N")
            SQL = SQL & ", iban=" & DBSet(RS!Iban, "T")
            
            'Octubre 2014.
            'Tipconta=1 SI, y solo si, Essocio=1 ; FechaBaja= Nulo ; tarifa precio = 1
            
            
            SQL = SQL & ", tipconta=" & TipoConta
            
            
            SQL = SQL & " where codsocio =" & IdAsoc
        End If
        conn.Execute SQL

        ActGasolineraAsociadoSocio = True
    End If
    RS.Close
    
    
    
    
    
    Set RS = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    
    Exit Function
ErrActGAS:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical
    Set RS = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    
End Function




'Crear en Arifacelec
Public Function ActAriFacElec(IdAsoc As Long) As Boolean
    
    Set RS = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim Aux As String
    Dim i As Integer
    Dim idArifacelec As Long
    Dim Actualizar As Boolean
    Dim NombeyEmail As String
    Dim AltaNueva As Boolean
    Dim ACtualizaIdGesso As Boolean
    
    SQL = "select * from asociados where IdAsoc = " & CStr(IdAsoc)
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        AltaNueva = False
        ACtualizaIdGesso = False
        
        '-- Ahora miramos si el asociado ya existe en la aplicación
        SQL = "select * from facelec_ariadna.cliente where cod_gessoc = " & CStr(IdAsoc)
        rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If rs2.EOF Then
        
            'No eixste. Podria ser que con el mismo NIF estuviera dado de alta como ARIGES
            
            SQL = ""
            
            'Veremos si esta con codigo de ariges, o codigo ariges2
            rs2.Close
            
            SQL = "select * from facelec_ariadna.cliente where codclien_ariges =" & CStr(IdAsoc) & " OR cod_clien_ariges2= " & CStr(IdAsoc)
            SQL = SQL & " ORDER BY codclien_ariges,cod_clien_ariges2"
            rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If rs2.EOF Then
                AltaNueva = True
            End If
            ACtualizaIdGesso = True
        End If
        
        
        If AltaNueva Then

            '-- No existe y hay que darlo de alta
            '   Para darlo de alta hay que conocer la entidad a la que pertenece
            '   y comprobar que está dada de alta en la gasolinera.
        
        
            'No lleva la columna ID
            SQL = "insert into facelec_ariadna.cliente (nombre,  login,    f_nueva,  id_empresa,  email,  contrasena, cif ,codclien_ariges" & _
                    ",codclien__arigasol,  codclien_ariagro,  tiene_factura_p_r_o_v,  cod_socio_ariagro,  cod_clien_ariges2,cod_teletaxi,  cod_gessoc"
                    
            SQL = SQL & ") VALUES ("
            SQL = SQL & DBSet(RS!nomlargo, "T") & ","
            SQL = SQL & DBSet(RS!NIF, "T") & ",0,1,"   'nueva y id_empresa
            SQL = SQL & DBSet(RS!mail, "T", "T") & ","
            'contrasena, cif
            SQL = SQL & DBSet(RS!NIF, "T") & "," & DBSet(RS!NIF, "T") & ","
            'Codclien ariges codclien__arigasol ....
            SQL = SQL & "0,0,0,0,0,0,0,"
            SQL = SQL & IdAsoc & ")"
            conn.Execute SQL
            
            'PARA QUE NO MUESTRE MSGBOX abajo
            SQL = ""
            
        Else
          
            '-- Ya exite, Comprobamos el NIF, y comprobamos varias cosas
            '--
            SQL = ""
            Actualizar = True
            idArifacelec = rs2!i_d
            NombeyEmail = RS!nomlargo & "|" & DBLet(RS!mail, "T") & "|"
            If rs2!CIF <> RS!NIF Then
                SQL = SQL & "    -CIF:  " & vbCrLf
                Actualizar = False
            End If
            
            If rs2!Nombre <> RS!nomlargo Then SQL = SQL & "    -nombre:" & rs2!Nombre & vbCrLf
            If rs2!Login <> RS!NIF Then
                SQL = SQL & "    -Login ()" & vbCrLf
                Actualizar = False
            End If
            
            
            If rs2!codclien__arigasol > 0 Then
                If rs2!codclien__arigasol <> RS!IdAsoc Then
                    SQL = SQL & "-Facelec  arigasol" & vbCrLf
                    Actualizar = False
                End If
            End If
            If rs2!codclien_ariges > 0 Then
                If rs2!codclien_ariges <> RS!IdAsoc Then
                    SQL = SQL & "-Facelec  ariges(1)" & vbCrLf
                    Actualizar = False
                End If
            End If
            If rs2!cod_clien_ariges2 > 0 Then
                If rs2!cod_clien_ariges2 <> RS!IdAsoc Then
                    SQL = SQL & "-Facelec  ariges(2)" & vbCrLf
                    Actualizar = False
                End If
            End If
            
            
            If rs2!cod_socio_ariagro > 0 Then
                If rs2!cod_socio_ariagro <> RS!CodSocEuroagro Then
                    SQL = SQL & "-Facelec  socio euroagro" & vbCrLf
                    Actualizar = False
                End If
            End If
            
            
            
        End If
        
        'Busco si hay algun cif o login en facelec que sea el del cliente y no sea el codigo asociado
        rs2.Close
        
        'Comprobaremos si hay algun datao en facelec.clientes que tenga ya ese nif o ese login
        ' el codasco NO sea el de aqui
            'Octubre 2014. Metemos or cod_gessoc is null
            Aux = "select * from facelec_ariadna.cliente where (cod_gessoc<>" & RS!IdAsoc & " or cod_gessoc is null )"
            Aux = Aux & " and (login=" & DBSet(RS!NIF, "T") & " or cif=" & DBSet(RS!NIF, "T") & ") ORDER BY i_d"
            rs2.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Aux = ""
            i = 0
            While Not rs2.EOF
                i = i + 1
                Aux = Aux & "     " & Format(rs2!i_d, "00000")
                If i > 4 Then
                    Aux = Aux & vbCrLf
                    i = 0
                End If
                rs2.MoveNext
            Wend
            
            If Aux <> "" Then
                Aux = "Errores vinculados al NIF(login-CIF)" & vbCrLf & Aux
                If SQL <> "" Then SQL = SQL & vbCrLf & vbCrLf
                SQL = SQL & Aux
                Actualizar = False
            End If
            
            If SQL <> "" Then
                SQL = "Campos erroneos en Facturacion Electronica" & vbCrLf & vbCrLf & SQL
                If Actualizar Then
                    SQL = SQL & vbCrLf & " Desea continuar de igual modo?"
                    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Actualizar = False
                Else
                    MsgBox SQL, vbExclamation
                End If
                
            End If
        
            If Actualizar Then
                'Actualizaremos Nombre, email y nada mas
                SQL = "UPDATE facelec_ariadna.cliente SET nombre=" & DBSet(RecuperaValor(NombeyEmail, 1), "T")
                SQL = SQL & ", email=" & DBSet(RecuperaValor(NombeyEmail, 2), "T", "S")
                'Actualizamos ges_soc y ariges
                If ACtualizaIdGesso Then SQL = SQL & ", cod_gessoc =" & RS!IdAsoc
                SQL = SQL & " WHERE i_d =" & idArifacelec
                conn.Execute SQL
            End If
        
        rs2.Close

        ActAriFacElec = True
    End If
    RS.Close
    Set RS = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    
    
End Function



