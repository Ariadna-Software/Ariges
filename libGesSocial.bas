Attribute VB_Name = "libGesSocial"
Option Explicit




Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Sql As String

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
    Sql = "select * from ariges" & rsUdNegocio!empresa_conta & ".sclien where codclien = " & IdAsoc
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    yaExiste_ = False
    If Not RS.EOF Then
        yaExiste_ = True
        Auxiliar2 = DBLet(RS!observac, "T")
    End If
    RS.Close


    'NO existe veo los valores por defecto para
    'defenvio,defzona,defruta,defagente,
    If Not yaExiste_ Then
        Sql = "Select defenvio,defzona,defruta,defagente from  ariges" & rsUdNegocio!empresa_conta & ".spara1"
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE SER EOF
        Auxiliar2 = RS!defenvio & "|" & RS!defzona & "|" & RS!defruta & "|" & RS!defagente & "|"
        RS.Close
            
    End If
        
    '-- Buscamos los datos del asociado
    Sql = "select * from asociados where IdAsoc = " & IdAsoc
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            Sql = "INSERT INTO ariges" & rsUdNegocio!empresa_conta & ".sclien (codclien, nomclien, nomcomer, domclien, codpobla"
            Sql = Sql & " ,pobclien, proclien, nifclien,   fechaalt, codactiv"
            Sql = Sql & " ,telclie1, faxclie1, maiclie1,  telclie2, faxclie2, "
            Sql = Sql & "  iban, codbanco, codsucur, digcontr, cuentaba, codmacta, codtarif "
            Sql = Sql & " ,codenvio, codzonas, codrutas, codagent, codforpa, diapago1"
            Sql = Sql & " ,clivario, tipoiva, tipofact, albarcon, periodof, numrepet,"
            Sql = Sql & " dtoppago, dtognral, promocio, codsitua, referobl,cliabono,pasclien"
            Sql = Sql & ") values ("
            Sql = Sql & IdAsoc & ","
            Sql = Sql & DBSet(RS!nomlargo, "T") & ","
            Sql = Sql & DBSet(RS!nomlargo, "T") & ","
            Sql = Sql & DBSet(RS!Direccion, "T") & ","
            Sql = Sql & DBSet(RS!CodPostal, "N") & ","
            Sql = Sql & DBSet(RS!Poblacion, "T") & ","
            Sql = Sql & DBSet(RS!Provincia, "T") & ","
            Sql = Sql & DBSet(RS!NIF, "T") & ","
            
            'Nov 2014
            Sql = Sql & DBSet(FechaDeAlta, "F") & ","
            
            
            'Antes de junio 14
            '
            'Codigo actividad
            'If RS!essocio Then
            '    SQL = SQL & "1"
            'Else
            '    SQL = SQL & "2"
            'End If
            Sql = Sql & RS!tarifaprecio
            
            Sql = Sql & "," & DBSet(RS!Telefono1, "T") & ","
            Sql = Sql & DBSet(RS!Movil, "T") & ","
            Sql = Sql & DBSet(RS!mail, "T") & ","
            Sql = Sql & DBSet(RS!Telefono2, "T") & "," & DBSet(RS!Telefono3, "T") & ","
            
            'iban, codbanco, codsucur, digcontr, cuentaba,codmacta
            Sql = Sql & DBSet(RS!IBAN, "T", "S") & ","
            Sql = Sql & DBLet(RS!entidad, "N") & ","
            Sql = Sql & DBLet(RS!Sucursal, "N") & ","
            Sql = Sql & DBSet(RS!DC, "T") & ","
            Sql = Sql & DBSet(RS!NumCC, "T") & ","
            'Codmacta
            Sql = Sql & DBSet(Codmacta, "T")
            
            'Junio 2014
            'Tarifaprecio es ACTIVIDAD
            ' codactiv = rs!tarifaprecio
            Sql = Sql & ",1,"
            
            
            'SQL = SQL & "'Gesocial: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & "',"
                
            'Auxiliar2 = rs!defenvio & "|" & rs!defzona & "|" & rs!defruta & "|" & rs!defagente & "|"
            Sql = Sql & RecuperaValor(Auxiliar2, 1) & ","
            Sql = Sql & RecuperaValor(Auxiliar2, 2) & ","
            Sql = Sql & RecuperaValor(Auxiliar2, 3) & ","
            Sql = Sql & RecuperaValor(Auxiliar2, 4) & ","
        
            'Codforpa
            Sql = Sql & rsUdNegocio!ForPa & ","
            
            
            
            'Diapago1, clivario  tipoiva, tipofact, albarcon, periodof, numrepet
            Sql = Sql & "10,0,0,0,0,1,1,"
        
            'dtoppago, dtognral, promocio, codsitua, referobl,  cliabono pasclien"
            'tarifaprecio
            Sql = Sql & "0,0,1,0,0,"
            If RS!tarifaprecio = 1 Then
                Sql = Sql & "0"
            Else
                Sql = Sql & "1"
            End If
            
            
            Sql = Sql & "," & DBSet(RS!NIF, "T")
            Sql = Sql & ")"
            
        Else
            'MODIFICAR
            
            'codclien, nomclien, nomcomer, domclien, codpobla"
            'pobclien, proclien, nifclien,
            'telclie1, faxclie1, maiclie1,  telclie2, faxclie2, "
            ' iban, codbanco, codsucur, digcontr, cuentaba, codmacta, observac "
            
            
            Sql = "UPDATE ariges" & rsUdNegocio!empresa_conta & ".sclien SET "
            Sql = Sql & " nomclien = " & DBSet(RS!nomlargo, "T")
            Sql = Sql & ", nomcomer = " & DBSet(RS!nomlargo, "T")
            Sql = Sql & ", domclien = " & DBSet(RS!Direccion, "T")
            Sql = Sql & ", codpobla = " & DBSet(RS!CodPostal, "N")
            Sql = Sql & ", pobclien = " & DBSet(RS!Poblacion, "T")
            Sql = Sql & ", proclien = " & DBSet(RS!Provincia, "T")
            Sql = Sql & ", nifclien = " & DBSet(RS!NIF, "T")
            
            Sql = Sql & ", telclie1 = " & DBSet(RS!Telefono1, "T")
            Sql = Sql & ", faxclie1 = " & DBSet(RS!Movil, "T")
            Sql = Sql & ", maiclie1 = " & DBSet(RS!mail, "T")
            Sql = Sql & ", telclie2 = " & DBSet(RS!Telefono2, "T")
            Sql = Sql & ", faxclie2 = " & DBSet(RS!Telefono3, "T")

            Sql = Sql & ", iban = " & DBSet(RS!IBAN, "T")
            Sql = Sql & ", codbanco = " & DBSet(RS!entidad, "N")
            Sql = Sql & ", codsucur = " & DBSet(RS!Sucursal, "N")
            Sql = Sql & ", digcontr = " & DBSet(RS!DC, "T")
            Sql = Sql & ", cuentaba = " & DBSet(RS!NumCC, "T")
                
            'Antes JUNIO 2014
            'SQL = SQL & ", codtarif = " & RS!tarifaprecio
            Sql = Sql & ", codactiv = " & RS!tarifaprecio
            
            'Cuent alternativa
            Sql = Sql & ", cliabono = "
            
            If RS!tarifaprecio = 1 Then
                Sql = Sql & "0"
            Else
                Sql = Sql & "1"
            End If
            
            'Observaciones
            'If Auxiliar2 <> "" Then Auxiliar2 = vbCrLf & Auxiliar2
            'Auxiliar2 = "Actualizado gessocial " & Format(Now, "dd/mm/yyyy hh:mm:ss") & Auxiliar2
            'SQL = SQL & ", observac = " & DBSet(Auxiliar2, "T")
            'SQL = SQL & ", codmacta = " & DBSet(Codmacta, "T")
            
            Sql = Sql & " WHERE codclien =" & IdAsoc
            
        End If
        
            
        If ejecutar(Sql, False) Then
            TraspasaAsociadoAriges = True
        
            
            'Actualizamos datos en contabilidad
            
            LaConta = DevuelveDesdeBD(conAri, "empresa_conta", "unidadesnegocio", "IdUnidad", rsUdNegocio!IdUnidad)
            Sql = DevuelveDesdeBD(conAri, "codmacta", "conta" & LaConta & ".cuentas", "codmacta", Codmacta)
            If Sql = "" Then
                'No existe la cuenta. La creo
                ActualizarLaCuenta2 LaConta, Codmacta, RS
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
    Sql = "select unidadesnegocio.* from asociados_unidadesnegocio,unidadesnegocio where "
    Sql = Sql & " asociados_unidadesnegocio.IdUnidad= unidadesnegocio.idunidad and idasoc=" & CStr(IdAsoc)
    If QueSeccion > 0 Then Sql = Sql & " AND unidadesnegocio.IdUnidad = " & QueSeccion
    
    rUd.Open Sql & " order by empresa_conta", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not rUd.EOF Then
        Sql = "select * from asociados where IdAsoc = " & CStr(IdAsoc)
        RS.Open Sql, conn, adOpenForwardOnly
        If Not RS.EOF Then
    
    
            While Not rUd.EOF
            'Datos asociado
                    
                    
                    Sql = "Select * from conta" & rUd!empresa_conta & ".empresa"
                    rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    'NO PUEDE SER EOF
                    i = rs2!numnivel
                    UltimoNivel = rs2.Fields("numdigi" & CStr(i))
                    rs2.Close
                    
                    
                    
                    'Para la gasolinera siempre cojera IdASOC
                    If rUd!IdUnidad = QueUDEsGasolinera Then
                        i = UltimoNivel - Len(rUd!raiz_cliente_asociado)
                        Codmacta = String(CLng(i), "0")
                         
                        Codmacta = rUd!raiz_cliente_asociado & Format(IdAsoc, Codmacta)
                        
                        ActualizarLaCuenta2 CStr(rUd!empresa_conta), Codmacta, RS
                       
                        conn.Execute "update asociados set codmacta = '" & Codmacta & "' where IdAsoc = " & CStr(IdAsoc)
                    Else
                        'Pueden ser varias cuentas a actualizar
                        If rUd!raiz_cliente_socio <> "" And RS!essocio = 1 Then
                            '
                             i = UltimoNivel - Len(rUd!raiz_cliente_socio)
                             Codmacta = String(CLng(i), "0")
                             
                             Codmacta = rUd!raiz_cliente_socio & Format(RS!CodSocEuroagro, Codmacta)
                             
                             ActualizarLaCuenta2 CStr(rUd!empresa_conta), Codmacta, RS
                                                          
                        End If
                                                
                        If rUd!raiz_cliente_asociado <> "" And RS!essocio = 0 Then
                            i = UltimoNivel - Len(rUd!raiz_cliente_asociado)
                            Codmacta = String(CLng(i), "0")
                             
                            Codmacta = rUd!raiz_cliente_asociado & Format(IdAsoc, Codmacta)
                            
                            ActualizarLaCuenta2 CStr(rUd!empresa_conta), Codmacta, RS
                            
                            
                            
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
                    ActualizarLaCuenta2 CStr(rUd!empresa_conta), Codmacta, RS
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


Private Sub ActualizarLaCuenta2(Contabilidad As String, Cuenta As String, ByRef vRs As ADODB.Recordset)
Dim Sql As String
        
        
        Sql = DevuelveDesdeBD(conAri, "codmacta", "conta" & Contabilidad & ".cuentas", "codmacta", Cuenta)
        If Sql = "" Then
            'NUEVO
            Sql = "INSERT INTO conta" & Contabilidad & ".cuentas(codmacta,nommacta,apudirec,model347,razosoci,dirdatos,codposta,despobla,"
            Sql = Sql & "desprovi,nifdatos,maidatos,obsdatos,pais,entidad,oficina,CC,cuentaba,iban) VALUES ('"
            Sql = Sql & Cuenta & "'," & DBSet(vRs!nomlargo, "T") & ",'S',1," & DBSet(vRs!nomlargo, "T") & ","
            Sql = Sql & DBSet(vRs!Direccion, "T") & "," & DBSet(vRs!CodPostal, "T") & "," & DBSet(vRs!Poblacion, "T") & ","
            Sql = Sql & DBSet(vRs!Provincia, "T") & "," & DBSet(vRs!NIF, "T") & "," & DBSet(vRs!mail, "T") & ","
            Sql = Sql & DBSet(vRs!Observaciones, "T") & ",'ESPAÑA'," & DBSet(vRs!entidad, "N") & "," & DBSet(vRs!Sucursal, "N")
            Sql = Sql & "," & DBSet(vRs!DC, "T") & "," & DBSet(vRs!NumCC, "T") & "," & DBSet(vRs!IBAN, "T") & ") "
        
        
        Else
                        
            '(codmacta,nommacta razosoci,dirdatos,codposta,despobla,"
            'desprovi,nifdatos,maidatos,obsdatos,pais,entidad,oficina,CC,cuentaba,iban
            
            'UPDATEAR
            Sql = "UPDATE conta" & Contabilidad & ".cuentas SET  nommacta = " & DBSet(vRs!nomlargo, "T")
            Sql = Sql & ", razosoci = " & DBSet(vRs!nomlargo, "T")
            Sql = Sql & ", dirdatos = " & DBSet(vRs!Direccion, "T")
            Sql = Sql & ", codposta = " & DBSet(vRs!CodPostal, "N")
            Sql = Sql & ", despobla = " & DBSet(vRs!Poblacion, "T")
            Sql = Sql & ", desprovi = " & DBSet(vRs!Provincia, "T")
            Sql = Sql & ", nifdatos = " & DBSet(vRs!NIF, "T")
            
            Sql = Sql & ", maidatos = " & DBSet(vRs!Telefono1, "T")
            Sql = Sql & ", obsdatos = " & DBSet(vRs!Movil, "T")

            Sql = Sql & ", iban = " & DBSet(vRs!IBAN, "T")
            Sql = Sql & ", entidad = " & DBSet(vRs!entidad, "N")
            Sql = Sql & ", oficina = " & DBSet(vRs!Sucursal, "N")
            Sql = Sql & ", CC = " & DBSet(vRs!DC, "T")
            Sql = Sql & ", cuentaba = " & DBSet(vRs!NumCC, "T")
            Sql = Sql & " WHERE codmacta = " & DBSet(Cuenta, "T")
         End If
         conn.Execute Sql

End Sub






Public Function ActualizaSocioAriagro(IdAsoc As Long) As Boolean
    '-- Montamos el bucle de lectura de todos los asociados / socios
    Dim i As Long
    Dim rs3 As ADODB.Recordset
    Dim CodMacCli As String
    Dim CodMacPro As String
    
    ActualizaSocioAriagro = False
    
    Sql = "select * from asociados"
    Sql = Sql & " where IdAsoc = " & CStr(IdAsoc)
    Sql = Sql & " and (fechabaja is null)"
    Sql = Sql & " and CodSocEuroagro < 10000"
    Sql = Sql & " and EsSocio = 1"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        '-- Lo primero es montar el nuevo socio
        Sql = DevuelveDesdeBD(conAri, "codsocio", "ariagro.rsocios", "codsocio", RS!CodSocEuroagro)
        
        If Sql = "" Then
        

            'NUEVO  NUEVO    NUEVO
            '-- NO existe y se da de alta
            Sql = "insert into ariagro.rsocios (codsocio, nifsocio, nomsocio, dirsocio, pobsocio, prosocio," & _
                    "codpostal, fechanac, telsoci1, telsoci2, telsoci3, movsocio, maisocio," & _
                    "codcoope, IBAN,codbanco, codsucur, digcontr, cuentaba, observaciones," & _
                    "fechaalta, fechabaja, correo, codiva, tipoirpf, tipoprod, codsitua) VALUES ("
            'odsocio, nifsocio, nomsocio, dirsocio, pobsocio, prosocio
            Sql = Sql & DBSet(RS!CodSocEuroagro, "N") & ","
            Sql = Sql & DBSet(RS!NIF, "T") & ","
            Sql = Sql & DBSet(RS!nomlargo, "T") & ","
            Sql = Sql & DBSet(RS!Direccion, "T") & ","
            Sql = Sql & DBSet(RS!Poblacion, "T") & ","
            Sql = Sql & DBSet(RS!Provincia, "T") & ","
    
            'codpostal, fechanac, telsoci1, telsoci2, telsoci3, movsocio, maisocio
            Sql = Sql & DBSet(RS!CodPostal, "N") & ","
            Sql = Sql & DBSet(RS!FechaNac, "F") & ","
            Sql = Sql & DBSet(RS!Telefono1, "T") & ","
            Sql = Sql & DBSet(RS!Telefono2, "T") & ","
            Sql = Sql & DBSet(RS!Telefono3, "T") & ","
            Sql = Sql & DBSet(RS!Movil, "T") & ","
            Sql = Sql & DBSet(RS!mail, "T") & ","
            
            'codcoope, IBAN,codbanco, codsucur, digcontr, cuentaba, observaciones
            Sql = Sql & "1,"   '-- Ponemos la cooperativa 1 a capón (OJO)
            Sql = Sql & DBSet(RS!IBAN, "T") & ","
            Sql = Sql & DBSet(RS!entidad, "N") & ","
            Sql = Sql & DBSet(RS!Sucursal, "N") & ","
            Sql = Sql & DBSet(RS!DC, "T") & ","
            Sql = Sql & DBSet(RS!NumCC, "T") & ","
            Sql = Sql & "'Gesocial: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & "',"
            
            'fechaalta, fechabaja, correo, codiva, tipoirpf, tipoprod, codsitua
            Sql = Sql & DBSet(RS!fechaalta, "F") & ","
            Sql = Sql & DBSet(RS!fechabaja, "F", "S") & ","
            Sql = Sql & DBSet(RS!Correo, "T") & ","
            Sql = Sql & DBSet(RS!TipoIrpf, "N") & ","
            Sql = Sql & "0,1,1)"
                
                
            
            
        Else
            'ACUTALIZAR
        
           
            ' nifsocio, nomsocio, dirsocio, pobsocio, prosocio
            Sql = "UPDATE ariagro.rsocios SET "
            Sql = Sql & " nifsocio = " & DBSet(RS!NIF, "T") & ","
            Sql = Sql & " nomsocio = " & DBSet(RS!nomlargo, "T") & ","
            Sql = Sql & " dirsocio = " & DBSet(RS!Direccion, "T") & ","
            Sql = Sql & " pobsocio = " & DBSet(RS!Poblacion, "T") & ","
            Sql = Sql & " prosocio = " & DBSet(RS!Provincia, "T") & ","
            
            
            'codpostal, fechanac, telsoci1, telsoci2, telsoci3, movsocio, maisocio
            Sql = Sql & " codpostal = " & DBSet(RS!CodPostal, "N") & ","
            Sql = Sql & " fechanac = " & DBSet(RS!FechaNac, "F") & ","
            Sql = Sql & " telsoci1 = " & DBSet(RS!Telefono1, "T") & ","
            Sql = Sql & " telsoci2 = " & DBSet(RS!Telefono2, "T") & ","
            Sql = Sql & " telsoci3 = " & DBSet(RS!Telefono3, "T") & ","
            Sql = Sql & " movsocio = " & DBSet(RS!Movil, "T") & ","
            Sql = Sql & " maisocio = " & DBSet(RS!mail, "T") & ","
            
            ' IBAN,codbanco, codsucur, digcontr, cuentaba, observaciones
            Sql = Sql & " IBAN = " & DBSet(RS!IBAN, "T") & ","
            Sql = Sql & " codbanco = " & DBSet(RS!entidad, "N") & ","
            Sql = Sql & " codsucur = " & DBSet(RS!Sucursal, "N") & ","
            Sql = Sql & " digcontr = " & DBSet(RS!DC, "T") & ","
            Sql = Sql & " cuentaba= " & DBSet(RS!NumCC, "T") & ","

            
            'fechaalta, fechabaja, correo, codiva,
            Sql = Sql & " fechaalta = " & DBSet(RS!fechaalta, "F") & ","
            'Sql = Sql & " fechabaja = " & DBSet(RS!fechabaja, "F", "S") & ","
            Sql = Sql & " correo = " & DBSet(RS!Correo, "N") & ","
            Sql = Sql & " codiva = " & DBSet(RS!TipoIrpf, "N")
            
            Sql = Sql & " WHERE codsocio = " & DBSet(RS!CodSocEuroagro, "N")
            
        End If
            
        conn.Execute Sql
            
        'En esta unidad de
            
            
            
            
        
        '-- Leemos su relación con unidades de negocio
        'sql = "select * from unidadesnegocio where IdAsoc = " & rs!IdAsoc
        Sql = "select * from unidadesnegocio where idunidad = 3" 'ARIAGRO
        Set rs2 = New ADODB.Recordset
        Set rs3 = New ADODB.Recordset
        rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not rs2.EOF
            'RAFA creaba o updateaba la tabla rseccion de ariagro
            '   '-- primero buscamos si la sección existe
            'sql = "select * from unidadesnegocio where IdUnidad = " & Rs2!IdUnidad
            'Set rs3 = gesDB.cursor(sql)
            'If Not rs3.EOF Then
            '    Set seccion = New RSeccion
            '    seccion.CodSecci = DBLet(rs3!CodSeccEuroagro, "N")
            '    seccion.NomSecci = DBLet(rs3!Nombre)
            '    seccion.EmpresaConta = DBLet(rs3!empresa_conta, "N")
            '    seccion.RaizClienteAsociado = DBLet(rs3!raiz_cliente_asociado)
            '    seccion.RaizClienteSocio = DBLet(rs3!raiz_cliente_socio)
            '    seccion.RaizProveedor = DBLet(rs3!raiz_proveedor)
            '    seccion.Guardar
            'End If
            
            '-- Y ahora la relación secciones socios
            
           ' socio_seccion.CodSecci = DBSet(seccion.CodSecci, "N")
           ' socio_seccion.CodSocio = DBLet(socio.CodSocio, "N")
           ' socio_seccion.FecAlta = DBLet(rs!FechaAlta, "F")
           ' socio_seccion.FecBaja = DBLet(rs!FechaBaja, "F")
           ' socio_seccion.CodIva = rs!CodIva ' (OJO) está a capon
           ' If seccion.RaizClienteSocio <> "" Then
           '     socio_seccion.CodMacCli = seccion.RaizClienteSocio & Format(socio.CodSocio, String(5, "0")) ' (OJO) la longitud de la cuenta
           ' End If
           ' If seccion.RaizProveedor <> "" Then
           '     socio_seccion.CodMacPro = seccion.RaizProveedor & Format(socio.CodSocio, String(5, "0")) ' (OJO) la longitud de la cuenta
           ' End If
           ' socio_seccion.Guardar
           
            Sql = "Select * from conta" & rs2!empresa_conta & ".empresa"
            rs3.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            'NO PUEDE SER EOF
            i = rs3!numnivel
            i = rs3.Fields("numdigi" & CStr(i))
            rs3.Close
           
            
            CodMacCli = ""
            If DBLet(rs2!raiz_cliente_asociado, "T") <> "" Then CodMacCli = rs2!raiz_cliente_asociado & Format(RS!CodSocEuroagro, String(i - Len(rs2!raiz_cliente_asociado), "0"))
            CodMacPro = ""
            If DBLet(rs2!raiz_proveedor, "T") <> "" Then CodMacPro = rs2!raiz_proveedor & Format(RS!CodSocEuroagro, String(i - Len(rs2!raiz_proveedor), "0"))

            
            Sql = " codsocio = " & RS!CodSocEuroagro & " and codsecci "
            Sql = DevuelveDesdeBD(conAri, "codsecci", "ariagro.rsocios_seccion", Sql, rs2!CodSeccEuroagro)
             '-- NO existe y se da de alta
            If Sql = "" Then
                Sql = "insert into ariagro.rsocios_seccion(codsocio, codsecci, fecalta, fecbaja," & _
                        "codmaccli, codmacpro, codiva) VALUES ("
                Sql = Sql & RS!CodSocEuroagro & "," & rs2!CodSeccEuroagro & ","
                Sql = Sql & DBSet(RS!fechaalta, "F") & ","
                Sql = Sql & DBSet(RS!fechabaja, "F", "S") & ","
                Sql = Sql & DBSet(CodMacCli, "T", "S") & ","
                Sql = Sql & DBSet(CodMacPro, "T", "S") & ","
                Sql = Sql & RS!CodIva & ")"
            Else
                '-- Si existe y se modifica
                Sql = "update ariagro.rsocios_seccion set "
                Sql = Sql & "fecalta = " & DBSet(RS!fechaalta, "F") & ","
                Sql = Sql & "fecbaja = " & DBSet(RS!fechabaja, "F", "S") & ","
                Sql = Sql & "codmaccli = " & DBSet(CodMacCli, "T", "S") & ","
                Sql = Sql & "codmacpro = " & DBSet(CodMacPro, "T", "S")
                Sql = Sql & " where codsocio = " & RS!CodSocEuroagro
                Sql = Sql & " and codsecci = " & rs2!CodSeccEuroagro
            End If
            conn.Execute Sql
           
            'LAs cremos en contabilidad
            If CodMacCli <> "" Then ActualizarLaCuenta2 CStr(rs2!empresa_conta), CodMacCli, RS
            If CodMacPro <> "" Then ActualizarLaCuenta2 CStr(rs2!empresa_conta), CodMacPro, RS
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
    Sql = "select * from asociados where IdAsoc = " & CStr(IdAsoc)
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        Sql = "select * from " & ElArigaso & ".ssocio where codsocio = " & CStr(IdAsoc)
        rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If rs2.EOF Then
            '-- No existe y hay que darlo de alta
            '   Para darlo de alta hay que conocer la entidad a la que pertenece
            '   y comprobar que está dada de alta en la gasolinera.
            '
        
            
            '-- Y ahora ya podemos seguir dando de alta al socio
            Sql = "insert into " & ElArigaso & ".ssocio (codsocio,codcoope,nomsocio,domsocio,codposta," & _
                    "pobsocio,prosocio,nifsocio,telsocio,movsocio,maisocio," & _
                    "fechaalt,codtarif,codbanco,codsucur,digcontr,cuentaba,iban," & _
                    "impfactu,dtolitro,codforpa,tipsocio,bonifbas,codsitua,codmacta,obssocio,tipconta)"
            Sql = Sql & " values("
            Sql = Sql & IdAsoc & "," & IdEntidadCoop & ","
            Sql = Sql & DBSet(RS!nomlargo, "T") & ","
            Sql = Sql & DBSet(RS!Direccion, "T") & ","
            Sql = Sql & DBSet(RS!CodPostal, "T") & ","
            Sql = Sql & DBSet(RS!Poblacion, "T") & ","
            Sql = Sql & DBSet(RS!Provincia, "T") & ","
            Sql = Sql & DBSet(RS!NIF, "T") & ","
            Sql = Sql & DBSet(RS!Telefono1, "T") & ","
            Sql = Sql & DBSet(RS!Movil, "T") & ","
            Sql = Sql & DBSet(RS!mail, "T") & ","
            Sql = Sql & DBSet(FechaAltaSeccion, "F") & ","
            Sql = Sql & "1,"  ' por defecto la tarifa es la 1
            Sql = Sql & DBSet(Right("0000" & DBLet(RS!entidad, "T"), 4), "T") & ","
            Sql = Sql & DBSet(Right("0000" & DBLet(RS!Sucursal, "T"), 4), "T") & ","
            Sql = Sql & DBSet(RS!DC, "T") & ","
            Sql = Sql & DBSet(RS!NumCC, "T") & ","
            Sql = Sql & DBSet(RS!IBAN, "T") & ","
            
            'SQL = SQL & gasDB.numero(0) & "," ' impfactu
            'SQL = SQL & gasDB.numero(0) & "," ' dtolitro
            'SQL = SQL & gasDB.numero(0) & "," ' codforpa
            'SQL = SQL & gasDB.numero(0) & "," ' tipsocio
            'SQL = SQL & gasDB.numero(0) & "," ' bonifbas
            'SQL = SQL & gasDB.numero(0) & "," ' codsitua
            Sql = Sql & "0,0,0,0,0,0,"


            Sql = Sql & DBSet(RS!Codmacta, "T") & ","
            Sql = Sql & DBSet(RS!Observaciones, "T") & ","
            
            
            Sql = Sql & TipoConta & ")"
            
        Else
          '  SQL = "select * from asociados_entidades where IdAsoc = " & IdAsoc
          '  Set Rs2 = gesDB.cursor(SQL)
          '  If Not Rs2.EOF Then
          '      IdEntidad = Rs2!IdEntidad
          '      ActGasolineraEntidadesColectivos IdEntidad, gesDB, gasDB
          '  End If
            '-- Ya exite, lo modificamos, aunque hay muchos campos que no se tocan
            Sql = "update " & ElArigaso & ".ssocio set "
            'FALTA###
            Sql = Sql & "codcoope=" & IdEntidadCoop & ","
            Sql = Sql & "nomsocio=" & DBSet(RS!nomlargo, "T") & ","
            Sql = Sql & "domsocio=" & DBSet(RS!Direccion, "T") & ","
            Sql = Sql & "codposta=" & DBSet(RS!CodPostal, "T") & ","
            Sql = Sql & "pobsocio=" & DBSet(RS!Poblacion, "T") & ","
            Sql = Sql & "prosocio=" & DBSet(RS!Provincia, "T") & ","
            
            Sql = Sql & "codbanco=" & DBSet(Right("0000" & DBLet(RS!entidad, "T"), 4), "T") & ","
            Sql = Sql & "codsucur=" & DBSet(Right("0000" & DBLet(RS!Sucursal, "T"), 4), "T") & ","
            Sql = Sql & "digcontr=" & DBSet(RS!DC, "T") & ","
            Sql = Sql & "cuentaba=" & DBSet(RS!NumCC, "T") & ","
            Sql = Sql & "nifsocio=" & DBSet(RS!NIF, "T") & ","
            Sql = Sql & "telsocio=" & DBSet(RS!Telefono1, "T") & ","
            Sql = Sql & "movsocio=" & DBSet(RS!Movil, "T") & ","
            Sql = Sql & "maisocio=" & DBSet(RS!mail, "T") & ","
            Sql = Sql & "codmacta=" & DBSet(RS!Codmacta, "T") & ","
            Sql = Sql & "obssocio=" & DBSet(RS!Observaciones, "T", "N")
            Sql = Sql & ", iban=" & DBSet(RS!IBAN, "T")
            
            'Octubre 2014.
            'Tipconta=1 SI, y solo si, Essocio=1 ; FechaBaja= Nulo ; tarifa precio = 1
            
            
            Sql = Sql & ", tipconta=" & TipoConta
            
            
            Sql = Sql & " where codsocio =" & IdAsoc
        End If
        conn.Execute Sql

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
    
    Sql = "select * from asociados where IdAsoc = " & CStr(IdAsoc)
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        AltaNueva = False
        ACtualizaIdGesso = False
        
        '-- Ahora miramos si el asociado ya existe en la aplicación
        Sql = "select * from facelec_ariadna.cliente where cod_gessoc = " & CStr(IdAsoc)
        rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If rs2.EOF Then
        
            'No eixste. Podria ser que con el mismo NIF estuviera dado de alta como ARIGES
            
            Sql = ""
            
            'Veremos si esta con codigo de ariges, o codigo ariges2
            rs2.Close
            
            Sql = "select * from facelec_ariadna.cliente where codclien_ariges =" & CStr(IdAsoc) & " OR cod_clien_ariges2= " & CStr(IdAsoc)
            Sql = Sql & " ORDER BY codclien_ariges,cod_clien_ariges2"
            rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Sql = ""
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
            Sql = "insert into facelec_ariadna.cliente (nombre,  login,    f_nueva,  id_empresa,  email,  contrasena, cif ,codclien_ariges" & _
                    ",codclien__arigasol,  codclien_ariagro,  tiene_factura_p_r_o_v,  cod_socio_ariagro,  cod_clien_ariges2,cod_teletaxi,  cod_gessoc"
                    
            Sql = Sql & ") VALUES ("
            Sql = Sql & DBSet(RS!nomlargo, "T") & ","
            Sql = Sql & DBSet(RS!NIF, "T") & ",0,1,"   'nueva y id_empresa
            Sql = Sql & DBSet(RS!mail, "T", "T") & ","
            'contrasena, cif
            Sql = Sql & DBSet(RS!NIF, "T") & "," & DBSet(RS!NIF, "T") & ","
            'Codclien ariges codclien__arigasol ....
            Sql = Sql & "0,0,0,0,0,0,0,"
            Sql = Sql & IdAsoc & ")"
            conn.Execute Sql
            
            'PARA QUE NO MUESTRE MSGBOX abajo
            Sql = ""
            
        Else
          
            '-- Ya exite, Comprobamos el NIF, y comprobamos varias cosas
            '--
            Sql = ""
            Actualizar = True
            idArifacelec = rs2!i_d
            NombeyEmail = RS!nomlargo & "|" & DBLet(RS!mail, "T") & "|"
            If rs2!CIF <> RS!NIF Then
                Sql = Sql & "    -CIF:  " & vbCrLf
                Actualizar = False
            End If
            
            If rs2!Nombre <> RS!nomlargo Then Sql = Sql & "    -nombre:" & rs2!Nombre & vbCrLf
            If rs2!Login <> RS!NIF Then
                Sql = Sql & "    -Login ()" & vbCrLf
                Actualizar = False
            End If
            
            
            If rs2!codclien__arigasol > 0 Then
                If rs2!codclien__arigasol <> RS!IdAsoc Then
                    Sql = Sql & "-Facelec  arigasol" & vbCrLf
                    Actualizar = False
                End If
            End If
            If rs2!codclien_ariges > 0 Then
                If rs2!codclien_ariges <> RS!IdAsoc Then
                    Sql = Sql & "-Facelec  ariges(1)" & vbCrLf
                    Actualizar = False
                End If
            End If
            If rs2!cod_clien_ariges2 > 0 Then
                If rs2!cod_clien_ariges2 <> RS!IdAsoc Then
                    Sql = Sql & "-Facelec  ariges(2)" & vbCrLf
                    Actualizar = False
                End If
            End If
            
            
            If rs2!cod_socio_ariagro > 0 Then
                If rs2!cod_socio_ariagro <> RS!CodSocEuroagro Then
                    Sql = Sql & "-Facelec  socio euroagro" & vbCrLf
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
                If Sql <> "" Then Sql = Sql & vbCrLf & vbCrLf
                Sql = Sql & Aux
                Actualizar = False
            End If
            
            If Sql <> "" Then
                Sql = "Campos erroneos en Facturacion Electronica" & vbCrLf & vbCrLf & Sql
                If Actualizar Then
                    Sql = Sql & vbCrLf & " Desea continuar de igual modo?"
                    If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Actualizar = False
                Else
                    MsgBox Sql, vbExclamation
                End If
                
            End If
        
            If Actualizar Then
                'Actualizaremos Nombre, email y nada mas
                Sql = "UPDATE facelec_ariadna.cliente SET nombre=" & DBSet(RecuperaValor(NombeyEmail, 1), "T")
                Sql = Sql & ", email=" & DBSet(RecuperaValor(NombeyEmail, 2), "T", "S")
                'Actualizamos ges_soc y ariges
                If ACtualizaIdGesso Then Sql = Sql & ", cod_gessoc =" & RS!IdAsoc
                Sql = Sql & " WHERE i_d =" & idArifacelec
                conn.Execute Sql
            End If
        
        rs2.Close

        ActAriFacElec = True
    End If
    RS.Close
    Set RS = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    
    
End Function



