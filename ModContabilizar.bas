Attribute VB_Name = "ModContabilizar"
Option Explicit


'===================================================================================
'CONTABILIZAR FACTURAS:
'Modulo para el traspaso de registros de cabecera y lineas de tablas de FACTURACION
'A las tablas de FACTURACION de Contabilidad
'====================================================================================

Private DtoGnral As Currency
Private DtoPPago As Currency
Private BaseImp As Currency
Private TotalFac As Currency
Private CCoste2 As String

Private vCCos As Byte
    'Para cuando pasamos en la contabilizacion de las facturas
    'Sera 2:    tiene mas de un centro de coste. Habra que agrupar por CC
    '     1:  o solo es un trabajador o tienen el mismo CC, con lo cual no hace falta agrupar por CC
    '     0:  no habra CC.  Si vpara.. tieneanalitica = false

Private conCtaAlt As Boolean 'el cliente utiliza cuentas alternativas

'Para pasar a contabilidad facturas de proveedor
Private AnyoFacPr As Integer 'año factura proveedor, es el ano de fecha_recepcion

'Modificacion Centro de coste.
'La factura cogera el Centro de coste del trabajador del albaran




'llevara: codmacta_proveedor | impo_retencion |
Private DatosRetencion As String
Private DatosAportacion As String





Public Function CrearTMPFacturas(cadTabla As String, cadWhere As String) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    SQL = "CREATE TEMPORARY TABLE tmpFactu ( "
    If cadTabla = "scafac" Then
        SQL = SQL & "codtipom char(3) NOT NULL default '',"
        SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        SQL = SQL & "codprove int(6) unsigned NOT NULL default '0',"
        SQL = SQL & "numfactu varchar(10)  NOT NULL ,"
    End If
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00') "
    conn.Execute SQL
     
     
    If cadTabla = "scafac" Then
        SQL = "SELECT codtipom, numfactu, fecfactu"
    Else
        SQL = "SELECT codprove, numfactu, fecfactu"
    End If
    SQL = SQL & " FROM " & cadTabla
    SQL = SQL & " WHERE " & cadWhere
    
    'DAVID###
    'Si son de proveedores el orden es MUY importante para
    'que vayan ordenaditas por fecha recepcion
    'ademas, por si tiene mas de una por prove añado los dos campos
    If cadTabla <> "scafac" Then SQL = SQL & " ORDER BY fecrecep,codprove,numfactu"

    
    SQL = " INSERT INTO tmpFactu " & SQL
    conn.Execute SQL

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        MuestraError Err.Number, "", Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpFactu;"
        conn.Execute SQL
    End If
End Function


Public Sub BorrarTMPFacturas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpFactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub InsertarTMPErrFac(MenError As String, cadWhere As String)
Dim SQL As String

    On Error Resume Next
    SQL = "Insert into tmpErrFac(codprove,numfactu,fecfactu,error) "
    SQL = SQL & " Select *," & DBSet(Mid(MenError, 1, 200), "T") & " as error From tmpFactu "
    SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "tmpFactu")
    conn.Execute SQL
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function CrearTMPErrFact(cadTabla As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    SQL = "CREATE TEMPORARY TABLE tmpErrFac ( "
    If cadTabla = "scafac" Then
        SQL = SQL & "codtipom char(3) NOT NULL default '',"
        SQL = SQL & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Else
        SQL = SQL & "codprove int(6) unsigned NOT NULL default '0',"
        SQL = SQL & "numfactu varchar(10) NOT NULL ,"
    End If
    SQL = SQL & "fecfactu date NOT NULL default '0000-00-00', "
    SQL = SQL & "error varchar(200) NULL )"
    conn.Execute SQL
     
     CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpErrFac;"
        conn.Execute SQL
    End If
End Function


Public Sub BorrarTMPErrFact()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpErrFac;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function ComprobarLetraSerie(cadTabla As String) As Boolean
'Para Facturas VENTA a clientes
'Comprueba que la letra del serie del tipo de movimiento es  correcta
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim Cad As String, devuelve As String

On Error GoTo EComprobarLetra

    ComprobarLetraSerie = False
    
    'Comprobar que existe la letra de serie en contabilidad
    If cadTabla = "scafac" Then
        'cargamos el RSConta con la tabla contadores de BD: Contabilidad
        'donde estan todas las letra de serie que existen en la contabilidad
        SQL = "Select distinct tiporegi from contadores"
        Set RSconta = New ADODB.Recordset
        RSconta.Open SQL, ConnConta, adOpenDynamic, adLockPessimistic, adCmdText
        If RSconta.EOF Then
            RSconta.Close
            Set RSconta = Nothing
            Exit Function
        End If
            
    
        'obtenemos los distintos tipos de movimiento que vamos a contabilizar
        'de las facturas seleccionadas
        SQL = "select distinct scafac.codtipom from " & cadTabla
        SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
'        SQL = SQL & cadWHERE
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        B = True
        While Not RS.EOF And B
            'comprobar que todas las letras serie existen en Ariges
            SQL = "letraser"
            devuelve = DevuelveDesdeBDNew(conAri, "stipom", "codtipom", "codtipom", RS!codtipom, "T", SQL)
            If devuelve = "" Then
                B = False
                Cad = RS!codtipom & " en BD de Gestión."
            ElseIf SQL <> "" Then
                'comprobar que todas las letras serie existen en la contabilidad
                devuelve = "tiporegi= " & DBSet(SQL, "T")
                RSconta.MoveFirst
                RSconta.Find (devuelve), , adSearchForward
                If RSconta.EOF Then
                    'no encontrado
                    B = False
                    Cad = SQL & " en BD de Contabilidad."
                End If
            End If
            If B Then Cad = Cad & DBSet(RS!codtipom, "T") & ","
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        RSconta.Close
        Set RSconta = Nothing
        
        If Not B Then 'Hay algun movimiento que no existe
            devuelve = "No existe el tipo de movimiento: " & Cad & vbCrLf
            devuelve = devuelve & "Consulte con el administrador."
            MsgBox devuelve, vbExclamation
            Exit Function
        End If
        
        'Todos los Tipo de movimiento existen
        If Cad <> "" Then
            Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitamos ult. coma
        
            'miramos si hay algun movimiento de factura que la letra serie sea nulo
            SQL = "select count(*) from stipom "
            SQL = SQL & "where codtipom IN (" & Cad & ") and (isnull(letraser) or letraser='')"
            If RegistrosAListar(SQL) > 0 Then
                SQL = "Hay algun tipo de movimiento de Facturación que no tiene letra serie." & vbCrLf
                SQL = SQL & "Comprobar en la tabla de tipos de movimiento: " & Cad
                MsgBox SQL, vbExclamation
                Exit Function
            End If
        End If
        ComprobarLetraSerie = True
    Else
        ComprobarLetraSerie = True
    End If

EComprobarLetra:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Letra Serie", Err.Description
    End If
End Function

'###### ESTE YA NO SE UTILIZA
'Public Function ComprobarNumFacturas(cadTabla As String, cadWConta) As Boolean
''Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
''vamos a contabilizar
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
'Dim b As Boolean
'
'    On Error GoTo ECompFactu
'
'    ComprobarNumFacturas = False
'
'    SQL = "SELECT numserie,codfaccl,anofaccl FROM cabfact "
'    SQL = SQL & " WHERE " & cadWConta
'
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Not RSconta.EOF Then
'        'Seleccionamos las distintas facturas que vamos a facturar
'        SQL = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser,scafac.numfactu,scafac.fecfactu "
'        SQL = SQL & " FROM (" & cadTabla & " INNER JOIN stipom ON " & cadTabla & ".codtipom=stipom.codtipom) "
'        SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
''        SQL = SQL & " WHERE " & cadWHERE
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "(numserie= " & DBSet(RS!LetraSer, "T") & " AND codfaccl=" & DBSet(RS!NumFactu, "N") & " AND anofaccl=" & Year(RS!FecFactu) & ")"
'            If SituarRSetMULTI(RSconta, SQL) Then
'                b = False
'                SQL = "          Nº Fac.: " & Format(RS!NumFactu, "0000000") & vbCrLf
'                SQL = SQL & "          Fecha: " & RS!FecFactu
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If Not b Then
'            SQL = "Ya existe la factura: " & vbCrLf & SQL
'            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
'
'            MsgBox SQL, vbExclamation
'            ComprobarNumFacturas = False
'        Else
'            ComprobarNumFacturas = True
'        End If
'    Else
'        ComprobarNumFacturas = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
'
'ECompFactu:
'     If Err.Number <> 0 Then
'        MuestraError Err.Number, "Comprobar Nº Facturas", Err.Description
'    End If
'End Function


Public Function ComprobarNumFacturas_new(cadTabla As String, cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim SQL As String
Dim SQLconta As String
Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECompFactu

    ComprobarNumFacturas_new = False
    
    
        
    
    
        'aqui aqui aqui aqui
    
        If vParamAplic.ContabilidadNueva Then
            SQLconta = "SELECT count(*) FROM factcli WHERE "
        Else
            SQLconta = "SELECT count(*) FROM cabfact WHERE "
        End If

        'Seleccionamos las distintas facturas que vamos a facturar
        SQL = "SELECT DISTINCT " & cadTabla & ".codtipom,letraser,scafac.numfactu,scafac.fecfactu "
        SQL = SQL & " FROM (" & cadTabla & " INNER JOIN stipom ON " & cadTabla & ".codtipom=stipom.codtipom) "
        SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "

        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not RS.EOF And B
            If vParamAplic.ContabilidadNueva Then
                SQL = "(numserie= " & DBSet(RS!LetraSer, "T") & " AND numfactu=" & DBSet(RS!NumFactu, "N") & " AND anofactu=" & Year(RS!FecFactu) & ")"
            Else
                SQL = "(numserie= " & DBSet(RS!LetraSer, "T") & " AND codfaccl=" & DBSet(RS!NumFactu, "N") & " AND anofaccl=" & Year(RS!FecFactu) & ")"
            End If
'            If SituarRSetMULTI(RSconta, SQL) Then
            SQL = SQLconta & SQL
            If RegistrosAListar(SQL, conConta) Then
                B = False
                SQL = "          Letra Serie: " & DBSet(RS!LetraSer, "T") & vbCrLf
                SQL = SQL & "          Nº Fac.: " & Format(RS!NumFactu, "0000000") & vbCrLf
                SQL = SQL & "          Fecha: " & Format(RS!FecFactu, "dd/mm/yyyy")
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        If Not B Then
            SQL = "Ya existe la factura: " & vbCrLf & SQL
            SQL = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & SQL
            
            MsgBox SQL, vbExclamation
            ComprobarNumFacturas_new = False
        Else
            ComprobarNumFacturas_new = True
        End If
'    Else
'        ComprobarNumFacturas_new = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
    Exit Function
    
ECompFactu:
     If Err.Number <> 0 Then
        ComprobarNumFacturas_new = False
        MuestraError Err.Number, "Comprobar Nº Facturas", Err.Description
    End If
End Function




'###### ESTE YA NO SE UTILIZA
'Public Function ComprobarCtaContable(cadTabla As String, Opcion As Byte) As Boolean
''Comprobar que todas las ctas contables de los distintos clientes de las facturas
''que vamos a contabilizar existan en la contabilidad
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
'Dim b As Boolean
'Dim cadG As String
'
'    On Error GoTo ECompCta
'
'    ComprobarCtaContable = False
'
'    If Opcion = 3 Then 'si hay analitica comprobar que todas las cuentas
'                        'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
'        cadG = "grupovta"
'        SQL = DevuelveDesdeBDNew(conConta, "parametros", "grupogto", "", "", "", cadG)
'        If SQL <> "" And cadG <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
'        ElseIf SQL <> "" Then
'            SQL = " AND (codmacta like '" & SQL & "%')"
'        ElseIf cadG <> "" Then
'            SQL = " AND (codmacta like '" & cadG & "%')"
'        End If
'        cadG = SQL
'    End If
'
'    SQL = "SELECT codmacta FROM cuentas "
'    SQL = SQL & " WHERE apudirec='S'"
'    If cadG <> "" Then SQL = SQL & cadG
'
'    Set RSconta = New ADODB.Recordset
'    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText
'
'    If Not RSconta.EOF Then
'        If Opcion = 1 Then
'            If cadTabla = "scafac" Then
'                'Seleccionamos los distintos clientes,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafac.codclien, sclien.codmacta "
'                SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
'            Else
'                'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
'                SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
'                SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
'                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
'            End If
'
'        ElseIf Opcion = 2 Or Opcion = 3 Then
'            SQL = "SELECT distinct "
'            If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
'            If cadTabla = "scafac" Then
'                SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
'            Else
'                SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
'                SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
'                SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
'            End If
'            SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
'        End If
'
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        b = True
'        While Not RS.EOF And b
'            SQL = "codmacta= " & DBSet(RS!Codmacta, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
'                b = False 'no encontrado
'                If Opcion = 1 Then
'                    If cadTabla = "scafac" Then
'                        SQL = RS!Codmacta & " del Cliente " & Format(RS!CodClien, "000000")
'                    Else
'                        SQL = RS!Codmacta & " del Proveedor " & Format(RS!codProve, "000000")
'                    End If
'                ElseIf Opcion = 2 Then
'                    SQL = RS!Codmacta & " de la familia " & Format(RS!codfamia, "0000")
'                ElseIf Opcion = 3 Then
'                    SQL = RS!Codmacta
'                End If
'            End If
'
'            If Opcion = 2 Then
'                'Comprobar que ademas de existir la cuenta de ventas exista tambien
'                'la cuenta ABONO ventas
'                SQL = "codmacta= " & DBSet(RS!ctaabono, "T")
'                RSconta.MoveFirst
'                RSconta.Find (SQL), , adSearchForward
'                If RSconta.EOF Then
'                    b = False 'no encontrado
'                    SQL = RS!ctaabono & " de la familia " & Format(RS!codfamia, "0000")
'                End If
'            End If
'
'            'comprobar cuentas alternativas solo para facturacion a clientes
'            If cadTabla = "scafac" Then
'                If Opcion = 2 Then
'                    ' Comprobar cuenta venta alternativa
'                    If DBLet(RS!ctavent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!ctavent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctavent1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta venta alternativa."
'                    End If
'                End If
'                If Opcion = 2 Then
'                    ' Comprobar cuenta de abono alternativa
'                    If DBLet(RS!abovent1, "T") <> "" Then
'                        SQL = "codmacta= " & DBSet(RS!abovent1, "T")
'                        RSconta.MoveFirst
'                        RSconta.Find (SQL), , adSearchForward
'                        If RSconta.EOF Then
'                            b = False 'no encontrado
'                            SQL = RS!ctaabon1 & " de la familia " & Format(RS!codfamia, "0000")
'                        End If
'                    Else
'                        b = False
'                        SQL = " o la familia no tiene asignada cuenta abono alternativa."
'                    End If
'                End If
'            End If
'            RS.MoveNext
'        Wend
'        RS.Close
'        Set RS = Nothing
'
'        If Not b Then
'            If Opcion <> 3 Then
'                SQL = "No existe la cta contable " & SQL
'            Else
'                SQL = "La cuenta " & SQL & " no es del nivel correcto."
'            End If
'            SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL
'
'            MsgBox SQL, vbExclamation
'            ComprobarCtaContable = False
'        Else
'            ComprobarCtaContable = True
'        End If
'    Else
'        ComprobarCtaContable = True
'    End If
'    RSconta.Close
'    Set RSconta = Nothing
'
'ECompCta:
'     If Err.Number <> 0 Then
'        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
'    End If
'End Function






Public Function ComprobarCtaContable_new(cadTabla As String, Opcion As Byte) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad

'NEUVO MARZO 2009
'COmprobaremos que no esten bloqueadas
Dim cContabF As CControlFacturaContab
Dim QueCuentasSon As String
Dim CtaBloq As Collection

Dim SQL As String
Dim RS As ADODB.Recordset
Dim Ic As Integer
'Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim cadG As String
Dim SQLcuentas As String
    
    On Error GoTo ECompCta

    ComprobarCtaContable_new = False
    
    cadG = ""
    
    
    If Opcion = 3 Then
            'si hay analitica comprobar que todas las cuentas
            'empiezan por el digito que hay en conta.parametros.grupogto o .grupovta
            cadG = "grupovta"
            SQL = DevuelveDesdeBDNew(conConta, "parametros", "grupogto", "", "", "", cadG)
            If SQL <> "" And cadG <> "" Then
                SQL = " AND (codmacta like '" & SQL & "%' OR codmacta like '" & cadG & "%')"
            ElseIf SQL <> "" Then
                SQL = " AND (codmacta like '" & SQL & "%')"
            ElseIf cadG <> "" Then
                SQL = " AND (codmacta like '" & cadG & "%')"
            End If
            cadG = SQL
    End If
    
    
    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    
    If Opcion = 1 Then
        If cadTabla = "scafac" Then
            'Seleccionamos los distintos clientes,cuentas que vamos a facturar
            
            SQL = "SELECT DISTINCT scafac.codclien, sclien.codmacta "
            SQL = SQL & " FROM (scafac INNER JOIN sclien ON scafac.codclien=sclien.codclien) "
            SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
        Else
            'Seleccionamos los distintos proveedores,cuentas que vamos a facturar
            SQL = "SELECT DISTINCT scafpc.codprove, sprove.codmacta "
            SQL = SQL & " FROM (scafpc INNER JOIN sprove ON scafpc.codprove=sprove.codprove) "
            SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
        End If
    
    ElseIf Opcion = 2 Or Opcion = 3 Then
        SQL = "SELECT distinct "
        If Opcion = 2 Then SQL = SQL & " sartic.codfamia,"
        If cadTabla = "scafac" Then
            SQL = SQL & " sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1 from ((slifac "
            SQL = SQL & " INNER JOIN tmpFactu ON slifac.codtipom=tmpFactu.codtipom AND slifac.numfactu=tmpFactu.numfactu AND slifac.fecfactu=tmpFactu.fecfactu) "
            SQL = SQL & "INNER JOIN sartic ON slifac.codartic=sartic.codartic) "
        Else
            SQL = SQL & " sfamia.ctacompr as codmacta,sfamia.abocompr as ctaabono from ((slifpc "
            SQL = SQL & " INNER JOIN tmpFactu ON slifpc.codprove=tmpFactu.codprove AND slifpc.numfactu=tmpFactu.numfactu AND slifpc.fecfactu=tmpFactu.fecfactu) "
            SQL = SQL & "INNER JOIN sartic ON slifpc.codartic=sartic.codartic) "
        End If
        SQL = SQL & " LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia "
        
    ElseIf Opcion = 4 Then
        'opcion para la contabilizacion de tickets AGRUPADA  FTG
        
        
        Set RS = New ADODB.Recordset
      
        
        cadG = "select sfactik.* from sfactik ,scafac where sfactik.numfacFTG=scafac.numfactu and sfactik.fecfacftg=scafac.fecfactu AND  codtipom='FTG' "
        RS.Open cadG, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        cadG = ""
        Do
            cadG = cadG & "," & RS!NumFactu
            RS.MoveNext
        Loop Until RS.EOF
        RS.Close
        Set RS = Nothing
        cadG = Mid(cadG, 2)
        'Monto el SELECT , igual que el de arriba, pero partiendo de los FTIs
         SQL = "SELECT distinct  sartic.codfamia, sfamia.ctaventa as codmacta,sfamia.aboventa as ctaabono, sfamia.ctavent1,sfamia.abovent1"
         SQL = SQL & " from (slifac   INNER JOIN sartic ON slifac.codartic=sartic.codartic)  LEFT OUTER JOIN sfamia ON sartic.codfamia=sfamia.codfamia"
         SQL = SQL & " WHERE  codtipom='FTI' and numfactu IN (" & cadG & ")"
         cadG = ""
         'Fuerzo para que haga las mismas comprobaciones que si fuera la opcion 2
         Opcion = 2
         
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    B = True
    QueCuentasSon = ""

    While Not RS.EOF And B
        SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!Codmacta, "T")
        
        'Para comporbar si estan bloqueadas
        QueCuentasSon = QueCuentasSon & ", '" & RS!Codmacta & "'"
        
        
        If Not (RegistrosAListar(SQL, conConta) > 0) Then
        'si no lo encuentra
            B = False 'no encontrado
            If Opcion = 1 Then
                If cadTabla = "scafac" Then
                    SQL = RS!Codmacta & " del Cliente " & Format(RS!codClien, "000000")
                Else
                    SQL = RS!Codmacta & " del Proveedor " & Format(RS!Codprove, "000000")
                End If
            ElseIf Opcion = 2 Then
                SQL = RS!Codmacta & " de la familia " & Format(RS!Codfamia, "0000")
            ElseIf Opcion = 3 Then
                SQL = RS!Codmacta
            End If
        End If
        
        
        If Opcion = 2 Or Opcion = 3 Then
            'Comprobar que ademas de existir la cuenta de ventas exista tambien
            'la cuenta ABONO ventas (sfamia.aboventa)
            '---------------------------------------------
            SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!ctaabono, "T")
'            RSconta.MoveFirst
'            RSconta.Find (SQL), , adSearchForward
'            If RSconta.EOF Then
            If Not (RegistrosAListar(SQL, conConta) > 0) Then
                B = False 'no encontrado
                If Opcion = 2 Then
                    SQL = RS!ctaabono & " de la familia " & Format(RS!Codfamia, "0000")
                ElseIf Opcion = 3 Then
                    SQL = RS!ctaabono
                End If
            End If
            
            
            'comprobar cuentas alternativas solo para facturacion a CLIENTES
            '----------------------------------------------------------------
            If cadTabla = "scafac" Then
                ' Comprobar cuenta VENTA alternativa
                If DBLet(RS!ctavent1, "T") <> "" Then
                    SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!ctavent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(SQL, conConta) > 0) Then
                        B = False 'no encontrado
                        If Opcion = 2 Then
                            SQL = RS!ctavent1 & " de la familia " & Format(RS!Codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            SQL = RS!ctavent1
                        End If
                    End If
                Else
                    B = False
                    SQL = " o la familia no tiene asignada cuenta venta alternativa."
                End If
                
                ' Comprobar cuenta de ABONO alternativa
                If DBLet(RS!abovent1, "T") <> "" Then
                    SQL = SQLcuentas & " AND codmacta= " & DBSet(RS!abovent1, "T")
'                    RSconta.MoveFirst
'                    RSconta.Find (SQL), , adSearchForward
'                    If RSconta.EOF Then
                    If Not (RegistrosAListar(SQL, conConta) > 0) Then
                        B = False 'no encontrado
                        If Opcion = 2 Then
                            SQL = RS!abovent1 & " de la familia " & Format(RS!Codfamia, "0000")
                        ElseIf Opcion = 3 Then
                            SQL = RS!abovent1
                        End If
                    End If
                Else
                    B = False
                    SQL = " o la familia no tiene asignada cuenta abono alternativa."
                End If
            End If
            
        End If
        
        RS.MoveNext
    Wend
    
    
        
        
        
        If Not B Then
            If Opcion <> 3 Then
                SQL = "No existe la cta contable " & SQL
            Else
                SQL = "La cuenta " & SQL & " no es del nivel correcto. (Familias de artículos)."
            End If
            SQL = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & SQL
            
            MsgBox SQL, vbExclamation
            ComprobarCtaContable_new = False
        Else
        
            'MARZO 2010
            'Para ver si estanbloqueadas las cuentas
            SQL = ""
            If QueCuentasSon <> "" Then
                QueCuentasSon = Mid(QueCuentasSon, 2)
                Set cContabF = New CControlFacturaContab
                cContabF.CuentasBloqueadas ConnConta, QueCuentasSon, Now, CtaBloq
                If CtaBloq.Count > 0 Then
                    'EXISTEN CUENTAS BLOQUEADAS
                    For Ic = 1 To CtaBloq.Count
                        QueCuentasSon = CtaBloq.item(Ic)
                        SQL = SQL & RecuperaValor(QueCuentasSon, 1) & "   " & RecuperaValor(QueCuentasSon, 2) & vbCrLf
                    Next
                    SQL = "Cuentas bloqueadas en contabilidad: " & vbCrLf & String(30, "=") & vbCrLf & SQL
                    MsgBox SQL, vbExclamation
                Else
                    SQL = ""
                End If
                Set cContabF = Nothing
            End If
            If SQL = "" Then
                ComprobarCtaContable_new = True
            Else
                ComprobarCtaContable_new = False
            End If
        End If
        
        
        
        
    Exit Function
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function







Public Function ComprobarTiposIVA(cadTabla As String) As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim i As Byte
'Dim CodigIVA As String

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    SQL = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open SQL, ConnConta, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            If cadTabla = "scafac" Then
                SQL = "SELECT DISTINCT scafac.codigiv" & i
                SQL = SQL & " FROM scafac "
                SQL = SQL & " INNER JOIN tmpFactu ON scafac.codtipom=tmpFactu.codtipom AND scafac.numfactu=tmpFactu.numfactu AND scafac.fecfactu=tmpFactu.fecfactu "
                SQL = SQL & " WHERE not isnull(codigiv" & i & ")"
'                SQL = SQL & " WHERE " & " codigiv" & i & " <> 0 "
            Else
                SQL = "SELECT DISTINCT scafpc.tipoiva" & i
                SQL = SQL & " FROM " & cadTabla
                SQL = SQL & " INNER JOIN tmpFactu ON scafpc.codprove=tmpFactu.codprove AND scafpc.numfactu=tmpFactu.numfactu AND scafpc.fecfactu=tmpFactu.fecfactu "
                SQL = SQL & " WHERE not isnull(tipoiva" & i & ")"
'                SQL = SQL & " WHERE " & " tipoiva" & i & " <> 0 "
            End If
'            SQL = SQL & " WHERE " & cadWHERE & " AND codigiv" & i & " <> 0 "

            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            B = True
            While Not RS.EOF And B
                SQL = "codigiva= " & DBSet(RS.Fields(0), "N")
                RSconta.MoveFirst
                RSconta.Find (SQL), , adSearchForward
                If RSconta.EOF Then
                    B = False 'no encontrado
                    SQL = "Tipo de IVA: " & RS.Fields(0)
                End If
                RS.MoveNext
            Wend
            RS.Close
            Set RS = Nothing
        
            If Not B Then
                SQL = "No existe el " & SQL
                SQL = "Comprobando Tipos de IVA en contabilidad..." & vbCrLf & vbCrLf & SQL
            
                MsgBox SQL, vbExclamation
                ComprobarTiposIVA = False
                Exit For
            Else
                ComprobarTiposIVA = True
            End If
        Next i
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function

'La comprobacion del centro de coste ha cambiado
'El centro de coste lo cojera de CADA factura donde tiene
'un trabajador asignado. Luego ya no necesito cadCC
'Comprobaremos:
'           que todas las facturas el trabajador asignado tiene CC
'           y que es distintos, puesto que si es el mismo CC no hare la fiesta

'Noviembe 2009
'
'   Analitica
'    vParamAplic.ModoAnalitica modo analitica: 0=trabajador, 1=Familia, 2=Proyecto
'
'       Es decir, todas las lineas traen el centro de coste asociado, con lo cual,
'   la opcion de comprobar coste será:
'MAYO 2010
'
' Todas las lineas de factura llevan el CC.
Public Function ComprobarCCoste(cadSQL As String, Clientes As Boolean) As Byte
Dim SQL As String
Dim i As Integer
Dim C As String
Dim Errores As String
    On Error GoTo ECCoste

    ComprobarCCoste = 0
    Set miRsAux = New ADODB.Recordset
    
    
    'anterior a la introducion de modoanalitica
    '
    'If Clientes Then
    '    SQL = "select codccost from scafac , scafac1, straba "
    '    SQL = SQL & " WHERE scafac.codtipom=scafac1.codtipom and scafac.numfactu=scafac1.numfactu and"
    '    SQL = SQL & " scafac.fecfactu=scafac1.fecfactu  and scafac1.codtraba=straba.codtraba"
    '
    'Else
    '    'PROVEEDORES
    '    SQL = "select codccost from scafpc ,scafpa, straba WHERE"
    '    SQL = SQL & " scafpc.codProve = scafpa.codProve And scafpc.NumFactu = scafpa.NumFactu And"
    '    SQL = SQL & " scafpc.FecFactu = scafpa.FecFactu AND codtrab2=straba.codtraba"
    '
    '
    'End If
    
    
    'AHORA
    If Clientes Then
    
        
        SQL = "select codccost from slifac where (codtipom,numfactu,fecfactu) "
        SQL = SQL & " in ( select codtipom,numfactu,fecfactu from scafac "
        If cadSQL <> "" Then SQL = SQL & " WHERE " & cadSQL
        SQL = SQL & ") GROUP BY codccost"
    
    
    
    
    
    '    select distinct(slifac.numfactu) from slifac where codtipom='FTI' and codccost is null AND (numfactu,fecfactu,codtipom) in (
' select sfactik.numfactu,sfactik.fecfactu,'FTI' from sfactik,scafac where sfactik.numfacftg=scafac.numfactu and scafac.fecfactu=sfactik.fecfacftg)
    
    
    
    Else
        SQL = "select codccost from slifpc where (codprove,numfactu,fecfactu) in ("
        SQL = SQL & "select codprove,numfactu,fecfactu from scafpc "
        If cadSQL <> "" Then SQL = SQL & " WHERE " & cadSQL
        SQL = SQL & ") GROUP BY codccost"
    

        
    End If
    
    Errores = ""  'De momento NO HAY ERRORES
    miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    SQL = ""
    While Not miRsAux.EOF
        If IsNull(miRsAux.Fields(0)) Then
            'MAL MAL. NO puede ser NULO
            Errores = Errores & "  ***  Lineas sin centro de coste asginado" & vbCrLf & vbCrLf
        Else
            SQL = SQL & DevNombreSQL(miRsAux.Fields(0)) & "|"
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If SQL <> "" Then
        
    
            
            While SQL <> ""
                i = InStr(1, SQL, "|")
                If i = 0 Then
                    'MSGBOX ALGO HA PASADO
                    Errores = Errores & " Sin asignar | en contabilizacion con centros de coste  " & vbCrLf
                    SQL = ""
                Else
                    C = Mid(SQL, 1, i - 1)
                    If vParamAplic.ContabilidadNueva Then
                        C = DevuelveDesdeBD(conConta, "codccost", "ccoste", "codccost", C, "T")
                    Else
                        C = DevuelveDesdeBD(conConta, "codccost", "cabccost", "codccost", C, "T")
                    End If
                    If C = "" Then
                        'ERROR EN CC. NO EXISTE
                        Errores = Errores & " - " & Mid(SQL, 1, i - 1) & "       no existe  " & vbCrLf
                    End If
                    SQL = Mid(SQL, i + 1)
                End If
            Wend
    
            If Errores <> "" Then
                MsgBox Errores, vbExclamation
                
            Else
                ComprobarCCoste = 2
            End If
    Else
            ComprobarCCoste = 0
            If Errores <> "" Then
                Errores = "Errores en CC. No deberia continuar. " & vbCrLf & Errores & "¿Continuar?"
                If MsgBox(Errores, vbQuestion + vbYesNo) = vbYes Then ComprobarCCoste = 1
                

            End If
    End If
    
    
    'ANTES
    'If SQL <> "" Then
    '    If Len(SQL) = 1 Then
    '        'Todos los CEntros de coste son el mismo. Con lo cual NO hara falta agrupar por trabajador
    '        ComprobarCCoste = 1
    '    Else
    '        'Tiene CC distintos. SI agruparemos por Trabajador
    '        ComprobarCCoste = 2
    '    End If
    'End If
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Cento de Coste", Err.Description
    End If
    Set miRsAux = Nothing
End Function



'Comprobara que todas las lineas que van a psara tiene CC
Public Function ComprobarCCosteTikAgrupado(cadSQL As String) As Byte
Dim SQL As String
Dim i As Integer
Dim C As String
Dim Errores As String
Dim ListaFTG As Collection
    On Error GoTo ECCoste

    ComprobarCCosteTikAgrupado = 0
    Set miRsAux = New ADODB.Recordset
    
        
    
    Errores = ""  'De momento NO HAY ERRORES
    SQL = "Select numfactu,fecfactu from scafac where " & cadSQL
    Set ListaFTG = New Collection
    miRsAux.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        SQL = " numfacFTG = " & miRsAux!NumFactu & " AND fecfacFTG = '" & Format(miRsAux!FecFactu, FormatoFecha) & "' "
        ListaFTG.Add SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    SQL = "|"
    
    For i = 1 To ListaFTG.Count
    
        C = "select codccost,count(*) from slifac where codtipom='FTI' AND"
        C = C & " (numfactu,fecfactu,codtipom) in ( select sfactik.numfactu,sfactik.fecfactu,'FTI' from sfactik "
        C = C & " where " & ListaFTG.item(i) & ") GROUP BY codccost order by codccost desc"
        miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            If IsNull(miRsAux.Fields(0)) Then
                'MAL MAL. NO puede ser NULO
                Errores = Errores & "  ***  Lineas sin centro de coste asginado" & vbCrLf & vbCrLf
            Else
                C = DevNombreSQL(miRsAux.Fields(0))
                If InStr(1, SQL, "|" & C & "|") = 0 Then SQL = SQL & C & "|"
                
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next i
    
    
    If Len(SQL) > 1 Then
            SQL = Mid(SQL, 2) 'Quito el primer pipe
    
            
            While SQL <> ""
                i = InStr(1, SQL, "|")
                If i = 0 Then
                    'MSGBOX ALGO HA PASADO
                    Errores = Errores & " Sin asignar | en contabilizacion con centros de coste  " & vbCrLf
                    SQL = ""
                Else
                    C = Mid(SQL, 1, i - 1)
                    If vParamAplic.ContabilidadNueva Then
                        C = DevuelveDesdeBD(conConta, "codccost", "ccoste", "codccost", C, "T")
                    Else
                        C = DevuelveDesdeBD(conConta, "codccost", "cabccost", "codccost", C, "T")
                    End If
                    If C = "" Then
                        'ERROR EN CC. NO EXISTE
                        Errores = Errores & " - " & Mid(SQL, 1, i - 1) & "       no existe  " & vbCrLf
                    End If
                    SQL = Mid(SQL, i + 1)
                End If
            Wend
    
            If Errores <> "" Then
                MsgBox Errores, vbExclamation
                
            Else
                ComprobarCCosteTikAgrupado = 2
            End If
    Else
            ComprobarCCosteTikAgrupado = 0
            If Errores <> "" Then
                Errores = "Errores en CC. No deberia continuar. " & vbCrLf & Errores & "¿Continuar?"
                If MsgBox(Errores, vbQuestion + vbYesNo) = vbYes Then ComprobarCCosteTikAgrupado = 2
                

            End If
    End If
    
    

ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Cento de Coste TICKETS AGRUPADOS", Err.Description
    End If
    Set miRsAux = Nothing
    Set ListaFTG = Nothing
End Function


'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador
'
Public Function PasarFactura(cadWhere As String, CodCCost As Byte, EsContabilizacionAgrupadaTickets As Boolean, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada

'EsContabilizacionAgrupadaTickets:  La diferencia es en las lineas de la factura.
'                                   Si false: procedimeineto normal
'                                       true: Las lineas hare los select de otra forma
Dim B As Boolean
Dim cadMen As String
Dim SQL As String
Dim ErrorContab As String

    On Error GoTo EContab

    ConnConta.BeginTrans
    conn.BeginTrans
    
    
    
    'MARZO 2011
    'lAs facturas internas NO van al registro
    'Solo meten el apunte
    If InStr(1, cadWhere, "'FAI'") > 0 Then
        'Estamos contabilizando una factura FAI, INTERNA
        'No entra en el registro de IVA de la contabilidad, solo el apunte
        vCCos = CodCCost
        
        B = ContabilizaFAI(cadMen, cadWhere, vContaFra)
        cadMen = "Insertando Cab. Factura interna: " & cadMen
        
    
    
    Else
    
        'Insertar en la conta Cabecera Factura
        B = InsertarCabFact(cadWhere, cadMen, vContaFra)
        cadMen = "Insertando Cab. Factura: " & cadMen
        vCCos = CodCCost
        If B Then
     
            'Insertar lineas de Factura en la Conta
            If EsContabilizacionAgrupadaTickets Then
                'Tickets agrupados
                B = InsertarLinFact_TicketsAgrupados("scafac", cadWhere, cadMen, False)
            Else
                'Normal. Esta es la forma NORMAL NORMAL de hacerlo
                B = InsertarLinFact_NUEVO2("scafac", cadWhere, cadMen, False)
            End If
            cadMen = "Insertando Lin. Factura: " & cadMen
    
            If vContaFra.RealizarContabilizacion Then
                ErrorContab = vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
                vContaFra.AnyadeElError ErrorContab
            End If
        End If
    End If  'de FAI

    
    If B Then
        'Poner intconta=1 en ariges.scafac
        B = ActualizarCabFact("scafac", cadWhere, cadMen)
        cadMen = "Actualizando Factura: " & cadMen
    End If

    
    
EContab:
    
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    

    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFactura = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFactura = False
        'Inserto en errores, DESPUES del rollback. Si no no lo refleja, y al hacer el rollback
        'tira atras la insercion
        SQL = "Insert into tmpErrFac(codtipom,numfactu,fecfactu,error) "
        SQL = SQL & " Select *," & DBSet(cadMen, "T") & " as error From tmpFactu "
        SQL = SQL & " WHERE " & Replace(cadWhere, "scafac", "tmpFactu")
        conn.Execute SQL
        
    End If
        

        
    
End Function


Private Function InsertarCabFact(cadWhere As String, cadErr As String, ByRef vCF As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim FraRectifica As String
Dim i As Integer
    On Error GoTo EInsertar
    
    Set RS = New ADODB.Recordset
    
    FraRectifica = ""
    If InStr(1, cadWhere, "'FRT'") > 0 Then
        '¡Voy a intentar sacar le numero de factura a la que rectifica. Sera de laobservacion
        SQL = Replace(cadWhere, "scafac.", "scafac1.")
        Cad = "select observa1 from scafac1 where " & SQL
        RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS!observa1) Then
                'Tiene valor, Vere si el valor es del tipo:
                ' RECTIFICA A FACTURA: A, 2600007, 30/1/2006
                Cad = CStr(RS!observa1)
                i = InStr(1, Cad, "TURA:")  '.. FACTURA: A, 2600007 ....
                If i > 0 Then
                    Cad = Mid(Cad, i + 5)
                    i = InStr(1, Cad, ",")
                    If i > 0 Then
                        'La letra
                        SQL = Trim(Mid(Cad, 1, i - 1))
                        Cad = Mid(Cad, i + 1)
                        'Busco el NUMERO DE factura
                        i = InStr(1, Cad, ",")
                        If i > 0 Then
                            Cad = Mid(Cad, 1, i - 1)
                            If IsNumeric(Cad) Then
                                'Biennnnnnnnnnnnnnn
                                'Ya tengo el numero de factura
                                SQL = SQL & Cad
                            Else
                                SQL = ""
                            End If
                            FraRectifica = SQL
                        End If 'De buscando letra
                    End If 'De buscando nºfac
                End If 'RECTIFICA A FACTURA: A, 2600007, 30/1/2006
            End If
        End If
        RS.Close
        Cad = ""
        
    End If
    SQL = " SELECT stipom.letraser,numfactu,fecfactu, sclien.codmacta,sclien.cliabono,year(fecfactu) as anofaccl,"
    SQL = SQL & "scafac.dtoppago,scafac.dtognral,baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,"
    SQL = SQL & "totalfac,codigiv1,codigiv2,codigiv3,aportacion "
    
    'Cuando MIS facfuras llevan recargo equivalencia
    SQL = SQL & ",porciva1re,porciva2re,porciva3re,imporiv1re,imporiv2re,imporiv3re,tipoiva,scafac.codtipom"
    
    SQL = SQL & " FROM (" & "scafac inner join " & "stipom on scafac.codtipom=stipom.codtipom) "
    SQL = SQL & "INNER JOIN " & "sclien ON scafac.codclien=sclien.codclien "
    SQL = SQL & " WHERE " & cadWhere
    
    
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not RS.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        DtoPPago = RS!DtoPPago
        DtoGnral = RS!DtoGnral
        BaseImp = RS!baseimp1 + CCur(DBLet(RS!baseimp2, "N")) + CCur(DBLet(RS!baseimp3, "N"))
        '---- Laura 10/10/2006:  añadir el totalfac para utilizarlo en insertar lineas
        TotalFac = RS!TotalFac
        DatosAportacion = ""
        If RS!Aportacion > 0 Then
            'Deberia dar error si vparam.ctaaportacion=""
            DatosAportacion = RS!Codmacta & "|" & RS!Aportacion & "|"
        Else
            
        End If
        '----
        conCtaAlt = RS!cliAbono
        
        
        'Guardamos los valores de la factura que estoy integrando
        If vCF.RealizarContabilizacion Then vCF.FijarNumeroFactura RS!NumFactu, Year(RS!FecFactu), RS!LetraSer
        
        SQL = ""
        SQL = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & DBSet(RS!FecFactu, "F") & "," & DBSet(RS!Codmacta, "T") & "," & Year(RS!FecFactu) & ","
        
        'MAYO 2009
        'Si es una factura rectificativa, y hemos encontrado
        ' a k factura rectifica entonces meto esto, sino sigue como antes
        If FraRectifica = "" Then
            
            If RS!codtipom = "FAT" Then
                'Las facturas de telefonia, llevaran en la observacion el NUmero de telefono
                Dim Aux As String
                
                Aux = "Serie='" & RS!LetraSer & "' AND ano = " & Year(RS!FecFactu) & " AND NumFact"
                Aux = DevuelveDesdeBD(conAri, "Telefono", "tel_cab_factura", Aux, CStr(RS!NumFactu))
                If Aux = "" Then Aux = "N/D"
                SQL = SQL & "'TEL " & Aux & "'"
            Else
                
                Select Case vParamAplic.ObsFactura
                Case 0
                    'Vacio
                    SQL = SQL & ValorNulo
                Case 1
                    'Nº Factura
                    SQL = SQL & "'" & DevNombreSQL("N/Fra " & RS!NumFactu) & "'"
                Case 2
                    'Fecha integracion
                    SQL = SQL & "'" & Format(Now, FormatoFecha) & "'"
                End Select
            
            End If
        Else
            SQL = SQL & "'" & FraRectifica & "'"
        End If
        
        '## LAURA (25/07/2008)
        Nulo2 = "N"
        Nulo3 = "N"
        If DBLet(RS!codigiv2, "N") = 0 Then Nulo2 = "S"
        If DBLet(RS!codigiv3, "N") = 0 Then Nulo3 = "S"
        
        'Abril
        SQL = SQL & "," & DBSet(RS!baseimp1, "N") & "," & DBSetDavid(RS!baseimp2, "N", Nulo2) & "," & DBSetDavid(RS!baseimp3, "N", Nulo3) & ","
        
        
        'SQL = SQL & DBSet(RS!porciva1, "N") & "," & DBSet(RS!porciva2, "N", Nulo2) & "," & DBSet(RS!porciva3, "N", Nulo3)
        SQL = SQL & DBSet(RS!porciva1, "N") & "," & DBSetDavid(RS!porciva2, "N", Nulo2) & "," & DBSetDavid(RS!porciva3, "N", Nulo3)
        
        
        
        SQL = SQL & "," & DBSet(RS!porciva1re, "N", "S") & "," & DBSet(RS!porciva2re, "N", "S") & "," & DBSet(RS!porciva3re, "N", "S")
        
        
        SQL = SQL & "," & DBSet(RS!imporiv1, "N", "N") & "," & DBSetDavid(RS!imporiv2, "N", Nulo2) & "," & DBSetDavid(RS!imporiv3, "N", Nulo3)
        
        'ANTES
        'SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & "," & DBSet(RS!imporiv1re, "N", "S") & "," & DBSet(RS!imporiv2re, "N", "S") & "," & DBSet(RS!imporiv3re, "N", "S") & ","
        
        
        SQL = SQL & DBSet(RS!TotalFac, "N") & "," & DBSet(RS!codigiv1, "N") & "," & DBSet(RS!codigiv2, "N", Nulo2) & "," & DBSet(RS!codigiv3, "N", Nulo3) & ","
        
        'INTRACOM
        If RS!TipoIVA = 3 Then
            'Tipo de iva intrcomunitatro
            SQL = SQL & "1"
        Else
            SQL = SQL & "0"
        End If
        
        SQL = SQL & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        SQL = SQL & DBSet(RS!FecFactu, "F")
        Cad = Cad & "(" & SQL & ")"
'        RS.MoveNext
    End If
    RS.Close
    Set RS = Nothing
    
    
    'Insertar en la contabilidad
    SQL = "INSERT INTO cabfact (numserie,codfaccl,fecfaccl,codmacta,anofaccl,confaccl,ba1faccl,ba2faccl,ba3faccl,"
    SQL = SQL & "pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,"
    SQL = SQL & "totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,numdiari,fechaent,numasien,fecliqcl) "
    SQL = SQL & " VALUES " & Cad
    ConnConta.Execute SQL
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFact = False
        cadErr = Err.Description
    Else
        InsertarCabFact = True
    End If
End Function



'Private Function InsertarLinFact(cadTabla As String, cadWhere As String, cadErr As String, Optional numRegis As Long) As Boolean
''cadWHere: selecciona un registro de scafac
''codtipom=x and numfactu=y and fecfactu=z
'Dim SQL As String
'Dim SQLaux As String
'Dim SQL2 As String
'Dim RS As ADODB.Recordset
'Dim Cad As String, Aux As String
'Dim I As Byte
'Dim TotImp As Currency, ImpLinea As Currency
'
'    On Error GoTo EInLinea
'
'    If cadTabla = "scafac" Then
'        SQL = " SELECT stipom.letraser,slifac.codtipom,numfactu,fecfactu,sartic.codfamia,sfamia.ctaventa,sfamia.ctavent1,sfamia.aboventa,sfamia.abovent1,sum(importel) as importe "
'        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
'        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafac", "slifac")
'        SQL = SQL & " GROUP BY sfamia.codfamia "
'    Else
'        SQL = " SELECT slifpc.codprove,numfactu,fecfactu,sartic.codfamia,sfamia.ctacompr,sfamia.abocompr,sum(importel) as importe "
'        SQL = SQL & " FROM (slifpc  "
'        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
'        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
'        SQL = SQL & " WHERE " & Replace(cadWhere, "scafpc", "slifpc")
'        SQL = SQL & " GROUP BY sfamia.codfamia "
'    End If
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    Cad = ""
'    I = 1
'    TotImp = 0
'    SQLaux = ""
'    While Not RS.EOF
'        SQLaux = Cad
'        'calculamos la Base Imp del total del importe para cada cta cble ventas
'        '---- Laura: 10/10/2006
'        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
'        ImpLinea = RS!Importe - CalcularPorcentaje(RS!Importe, DtoPPago, 2)
'        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
'        ImpLinea = ImpLinea - CalcularPorcentaje(RS!Importe, DtoGnral, 2)
'        'ImpLinea = Round(ImpLinea, 2)
'        '----
'        TotImp = TotImp + ImpLinea
'
'        'concatenamos linea para insertar en la tabla de conta.linfact
'        SQL = ""
'        SQL2 = ""
'        If cadTabla = "scafac" Then
'            SQL = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & Year(RS!FecFactu) & "," & I & ","
'            If Not conCtaAlt Then 'cliente no tiene cuenta alternativa
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctaventa, "T")
'                Else
'                    SQL = SQL & DBSet(RS!aboventa, "T")
'                End If
'            Else
'                If ImpLinea >= 0 Then
'                    SQL = SQL & DBSet(RS!ctavent1, "T")
'                Else
'                    SQL = SQL & DBSet(RS!abovent1, "T")
'                End If
'            End If
'        Else
'            SQL = numRegis & "," & Year(RS!FecFactu) & "," & I & ","
'            If ImpLinea >= 0 Then
'                SQL = SQL & DBSet(RS!ctacompr, "T")
'            Else
'                SQL = SQL & DBSet(RS!abocompr, "T")
'            End If
'        End If
'        SQL2 = SQL & ","
'        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
'
'        If CCoste = "" Then
'            SQL = SQL & ValorNulo
'        Else
'            SQL = SQL & DBSet(CCoste, "T")
'        End If
'
'        Cad = Cad & "(" & SQL & ")" & ","
'
'        I = I + 1
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
'
'    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
'    'de la factura
'    If TotImp <> BaseImp Then
''        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
'        'en SQL esta la ult linea introducida
'        TotImp = BaseImp - TotImp
'        TotImp = ImpLinea + TotImp '(+- diferencia)
'        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
'        If CCoste = "" Then
'            SQL2 = SQL2 & ValorNulo
'        Else
'            SQL2 = SQL2 & DBSet(CCoste, "T")
'        End If
'        If SQLaux <> "" Then 'hay mas de una linea
'            Cad = SQLaux & "(" & SQL2 & ")" & ","
'        Else 'solo una linea
'            Cad = "(" & SQL2 & ")" & ","
'        End If
'
''        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
''        cad = Replace(cad, SQL, Aux)
'    End If
'
'
'    'Insertar en la contabilidad
'    If Cad <> "" Then
'        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
'        If cadTabla = "scafac" Then
'            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
'        Else
'            SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
'        End If
'        SQL = SQL & " VALUES " & Cad
'        ConnConta.Execute SQL
'    End If
'
'EInLinea:
'    If Err.Number <> 0 Then
'        InsertarLinFact = False
'        cadErr = Err.Description
'    Else
'        InsertarLinFact = True
'    End If
'End Function
'


'
'Si lleva retencion(FRAPRO) se añadiren dos lineas codprove contra ctareten

Private Function InsertarLinFact_NUEVO2(cadTabla As String, cadWhere As String, cadErr As String, vLlevaRetencion As Boolean, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQLaux As String
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim TotImp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim LineaCentroCoste As Boolean
    'Puede ser que teniendo analitica, la cuenta no sea del grupo 6 o 7 , con lo cual nodebe poner el CC
    'Por si acaso alguna linea no es del grupo venta o grupo compras, no

    On Error GoTo EInLinea
    

    '
    '   Habra que ver en funcion de CC que tenga si agrupo, o no, por  codtraba
    '
    If cadTabla = "scafac" Then 'VENTAS
         'comprobar si el cliente utiliza cuenta alternativa
        If conCtaAlt Then
            'utilizamos sfamia.ctavent1 o sfamia.abovent1
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctavent1"
            Else
                cadCampo = "sfamia.abovent1" 'si es negativa es un abono
            End If
        Else
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctaventa"
            Else
                cadCampo = "sfamia.aboventa"
            End If
        End If
        
        
        'PARA las FAS, si viene la cuenta de venta, leida desde advparametros, entonces pone esa
        If InStr(1, cadWhere, "'FAS'") > 0 Then
            If conCtaAlt Then
                cadCampo = "sfamia.ctavent1"
            Else
                cadCampo = "sfamia.ctavtaseralt"
            End If
        End If
        
        SQL = " SELECT stipom.letraser,slifac.codtipom,slifac.numfactu,slifac.fecfactu," & cadCampo & " as cuenta,sum(importel) as importe"
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",slifac.codccost"
        
        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        
        
        SQL = SQL & " WHERE "
        
        
        SQL = SQL & " " & Replace(cadWhere, "scafac", "slifac")
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos > 0 Then SQL = SQL & " codccost, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo
        
    Else 'COMPRAS
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            cadCampo = "sfamia.ctacompr"
        Else
            cadCampo = "sfamia.abocompr"
        End If
        
        SQL = "SELECT slifpc.codprove,slifpc.numfactu,slifpc.fecfactu," & cadCampo & " as cuenta, sum(importel) as importe  "
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & ",slifpc.codccost"
                
        
        SQL = SQL & " FROM (slifpc  "
        SQL = SQL & " inner join sartic on slifpc.codartic=sartic.codartic) "
        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        
        If vCCos > 0 Then SQL = SQL & ",scafpa "
        
        SQL = SQL & " WHERE "
        
        'si tiene analitica, enlazo por con scafpa
        If vCCos > 0 Then SQL = SQL & " slifpc.NumFactu = scafpa.NumFactu And slifpc.FecFactu = scafpa.FecFactu and slifpc.codprove=scafpa.codprove AND slifpc.numalbar=scafpa.numalbar AND "
            
        SQL = SQL & Replace(cadWhere, "scafpc", "slifpc")
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos = 2 Then SQL = SQL & " codccost, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo
        
        
        
    End If
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    TotImp = 0
    SQLaux = ""
    Aux = ""
    While Not RS.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        TotImp = TotImp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL2 = ""
        
        If cadTabla = "scafac" Then 'VENTAS a clientes
            'En aux guardaremos el trozo comun de las lineas (letra/numero/anño
            If Aux = "" Then Aux = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & Year(RS!FecFactu) & ","
            SQL = Aux & i & ","
            SQL = SQL & DBSet(RS!Cuenta, "T")

        Else 'COMPRAS
            'Laura 24/10/2006
            'SQL = numRegis & "," & Year(RS!FecFactu) & "," & i & ","
            SQL = numRegis & "," & AnyoFacPr & "," & i & ","
            
'            If ImpLinea >= 0 Then
                SQL = SQL & DBSet(RS!Cuenta, "T")
'            Else
'                SQL = SQL & DBSet(RS!abocompr, "T")
'            End If
        End If
        

        
        SQL2 = SQL & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        
        'CENTRO DE COSTE
        LineaCentroCoste = False
        If vCCos = 0 Then
            'NO NECESTIA CENTRO DE COSTE.. seguro
           ' SQL = SQL & ValorNulo
        Else
            LineaCentroCoste = CuentaNecesitaCentroCoste(CStr(RS!Cuenta))
            
        End If
        If LineaCentroCoste Then
            CCoste2 = RS!CodCCost
            SQL = SQL & DBSet(CCoste2, "T")
        Else
            SQL = SQL & ValorNulo
        End If
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close

    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If TotImp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        TotImp = BaseImp - TotImp
        TotImp = ImpLinea + TotImp '(+- diferencia)
        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
        If CCoste2 = "" Then
            SQL2 = SQL2 & ValorNulo
        Else
            SQL2 = SQL2 & DBSet(CCoste2, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            Cad = SQLaux & "(" & SQL2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & SQL2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If



    'Si lleva retencion, solo sera en caso de facturas proveedores, entonces metere dos lineas mas
    '
    If vLlevaRetencion Then
        'Cojere los datos del proveedor
        'Reutilizo total fac
        TotalFac = CCur(RecuperaValor(DatosRetencion, 2))
        i = i + 1
        SQL = "(" & numRegis & "," & AnyoFacPr & "," & i & ",'" & RecuperaValor(DatosRetencion, 1) & "'," & TransformaComasPuntos(CStr(-TotalFac)) & ",NULL)"
        Cad = Cad & SQL
        i = i + 1
        SQL = ",(" & numRegis & "," & AnyoFacPr & "," & i & ",'" & vParamAplic.CtaReten & "'," & TransformaComasPuntos(CStr(TotalFac)) & ",NULL),"
        Cad = Cad & SQL
        
    End If

    
    
    
    'Facturas clientes. Ver si lleva aportacion al terminal
    If cadTabla = "scafac" Then
        If DatosAportacion <> "" Then
            
            
            SQL = "(" & Aux & i & ",'" & RecuperaValor(DatosAportacion, 1) & "',"
            'Dejo en DatosAportacion solo el importe
            DatosAportacion = TransformaComasPuntos(RecuperaValor(DatosAportacion, 2))
            SQL = SQL & DatosAportacion & ",NULL),"
            Cad = Cad & SQL
            i = i + 1                                                                                   'Importe en negativo
            SQL = "(" & Aux & i & ",'" & vParamAplic.ctaAportacion & "',-" & DatosAportacion & ",NULL),"
            Cad = Cad & SQL
        
        
        
    
        End If
    End If

    Set RS = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        If cadTabla = "scafac" Then
            SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        Else
            SQL = "INSERT INTO linfactprov (numregis,anofacpr,numlinea,codtbase,impbaspr,codccost) "
        End If
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    End If




EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_NUEVO2 = False
        cadErr = Err.Description
    Else
        InsertarLinFact_NUEVO2 = True
    End If
End Function







Private Function ActualizarCabFact(cadTabla As String, cadWhere As String, cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    
    SQL = "UPDATE " & cadTabla & " SET intconta=1 "
    SQL = SQL & " WHERE " & cadWhere

    conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        cadErr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function



'----------------------------------------------------------------------
' FACTURAS PROVEEDOR
'----------------------------------------------------------------------
'Ccoste
'   0: No tendra analitica
'   1: Solo hay un CC que tratar. NO agruparemos por trabajador
'   2: Mas de un CC. Agruparemos por trabajador


'Ahora la retencion puede llevarla CUALQUIERA de las facturas.
'   0. Retencion NORMAL
'   1. Retencion SOCIOS

Public Function PasarFacturaProv(cadWhere As String, CodCCost As Byte, FechaFin As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura PROVEEDOR
' ariges.scafpc --> conta.cabfactprov
' ariges.slifpc --> conta.linfactprov
'Actualizar la tabla ariges.scafpc.inconta=1 para indicar que ya esta contabilizada

'Modificacion Enero2008.   Tipo cooperativas (Ej. Terrasana)
'                          Si lleva retencion la factura, y el preoveedore es tipo REA
'                          entonces  a la contabilidad
'                          El importe de la factura es totfac + retencion
'                          y a las lineas van dos lineas mas
'                          proveedor     -impret
'                          ctareten      +impret

'Abril 2015.  Inversion de sujeto pasivo


Dim B As Boolean
Dim cadMen As String
Dim SQL As String
Dim Mc As Contadores
Dim vLlevaRetencion As Boolean
Dim i As Integer
Dim FraIntraCom2 As String
Dim Actual_sig As Boolean
    On Error GoTo EContab

'Sep 2012
'Comento este trozo pq si no solo salia UNA factura en el listado de facturas contabilizadas
'''''''    ' Mosrtaremos para cada factura de PROVEEDOR
'''''''    SQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
'''''''    conn.Execute SQL


    ConnConta.BeginTrans
    conn.BeginTrans
        
    
    Set Mc = New Contadores
    vLlevaRetencion = False 'Si llevara retencion me lo devolvera la fucion insertar
    FraIntraCom2 = ""
    '---- Insertar en la conta Cabecera Factura
    B = InsertarCabFactProv(cadWhere, cadMen, Mc, FechaFin, vLlevaRetencion, vContaFra, FraIntraCom2, Actual_sig)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    If B Then
        
        'Veremos que opcion de CC es la que hay que pasar (agrupar o no agrupar)
        vCCos = CodCCost
        '---- Insertar lineas de Factura en la Conta
        B = InsertarLinFact_NUEVO2("scafpc", cadWhere, cadMen, vLlevaRetencion, Mc.Contador)
        cadMen = "Insertando Lin. Factura: " & cadMen

        
        If B Then
            If vContaFra.RealizarContabilizacion Then
                vContaFra.AnyadeElError vContaFra.IntegraLaFacturaProv(vContaFra.NumeroFactura, vContaFra.Anofac)
            End If
        End If
        
        If B Then
            '---- Poner intconta=1 en ariges.scafac
            B = ActualizarCabFact("scafpc", cadWhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
        

        
    End If
    
    
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarFacturaProv = True
        
        
        'FEBRERO 2011
        'Si es Factura Intracom
        If FraIntraCom2 <> "" Then
            If vParamAplic.CtaContabIntracom <> "" Then
                Espera 0.5
                
                
                FraIntraCom2 = Mc.Contador & "|" & FraIntraCom2 & "|"
                'Haremos el proceso de  la insercion de dos fras extra, a partir de la ya insertada
                ConnConta.BeginTrans
                Mc.ConseguirContador "1", Actual_sig, True 'Siguiente contador
                
                If InsertarFrasEXTRAdeIntracomunitarias(Mc.Contador, FraIntraCom2) Then
                    ConnConta.CommitTrans
                Else
                    ConnConta.RollbackTrans
                End If
            End If
        End If
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarFacturaProv = False

        InsertarTMPErrFac cadMen, cadWhere
        
        'Si es correcto entonces creo una entrada en tmp para luego listar los resultados de
        'la contabilizacion
         If Mc.Contador > 0 Then
            SQL = "DELETE from tmpinformes where codusu = " & vUsu.codigo & " AND codigo1= " & Mc.Contador
            conn.Execute SQL
        End If
    
    End If
End Function


'Si es intracomunitaria devolvera el anofacpr para la generacion de las "extras"
Private Function InsertarCabFactProv(cadWhere As String, cadErr As String, ByRef Mc As Contadores, FechaFin As String, ByRef LlevaRetencion As Boolean, ByRef vCF As cContabilizarFacturas, ByRef EsFacturaIntracom2 As String, ByRef ContadorContaActual As Boolean) As Boolean
'Insertando en tabla conta.cabfact
'(OUT) AnyoFacPr: aqui devolvemos el año de fecha recepcion para insertarlo en las lineas de factura de proveedor de la conta
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar
       
    
    SQL = SQL & " SELECT fecfactu,year(fecrecep) as anofacpr,fecrecep,numfactu,sprove.codmacta,"
    SQL = SQL & "scafpc.dtoppago,scafpc.dtognral,baseiva1,baseiva2,baseiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,"
    SQL = SQL & "totalfac,tipoiva1,tipoiva2,tipoiva3,tipprove,impret,scafpc.nomprove,scafpc.codprove,tiporet,PorRet,impret,InvSujPas "   'Modificacion facturas socios
    SQL = SQL & " FROM " & "scafpc"
    SQL = SQL & " INNER JOIN " & "sprove ON scafpc.codprove=sprove.codprove "
    SQL = SQL & " WHERE " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Cad = ""
    If Not RS.EOF Then
        ContadorContaActual = (RS!FecRecep <= CDate(FechaFin) - 365)
        If Mc.ConseguirContador("1", ContadorContaActual, True) = 0 Then
        
                
        
        
            'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
            DtoPPago = RS!DtoPPago
            DtoGnral = RS!DtoGnral
            BaseImp = RS!BaseIVA1 + CCur(DBLet(RS!BaseIVA2, "N")) + CCur(DBLet(RS!BaseIVA3, "N"))
            TotalFac = RS!TotalFac
            AnyoFacPr = RS!anofacpr
            
            'Para que contabilice las facturas automaticamente
            If vCF.RealizarContabilizacion Then vCF.FijarNumeroFactura Mc.Contador, AnyoFacPr, ""
            
            'SI es facutra socio y tiene retencion
            If RS!TipoRet = 1 Then 'FACTURA SOCIO, con retencion
                If DBLet(RS!ImpRet, "N") <> 0 Then
                    'El total factura es totafac+ retencion
                    DatosRetencion = RS!Codmacta & "|" & RS!ImpRet & "|"
                    TotalFac = TotalFac + RS!ImpRet  'Luego en las lineas va la resta de este importe
                    LlevaRetencion = True
                Else
                    DatosRetencion = ""
                End If
            End If
            
            Nulo2 = "N"
            Nulo3 = "N"
            If DBLet(RS!BaseIVA2, "N") = "0" Then Nulo2 = "S"
            If DBLet(RS!BaseIVA3, "N") = "0" Then Nulo3 = "S"
            SQL = ""
            SQL = Mc.Contador & "," & DBSet(RS!FecFactu, "F") & "," & RS!anofacpr & "," & DBSet(RS!FecRecep, "F") & "," & DBSet(RS!NumFactu, "T") & "," & DBSet(RS!Codmacta, "T") & ","
            
            Select Case vParamAplic.ObsFactura
            Case 0
                'Vacio
                SQL = SQL & ValorNulo
            Case 1
                'Nº Factura
                SQL = SQL & "'" & DevNombreSQL("S/Fra " & RS!NumFactu) & "'"
            Case 2
                'Fecha integracion
                SQL = SQL & "'" & Format(Now, FormatoFecha) & "'"
            End Select
            SQL = SQL & "," & DBSet(RS!BaseIVA1, "N") & "," & DBSet(RS!BaseIVA2, "N", "S") & "," & DBSet(RS!BaseIVA3, "N", "S") & ","
            SQL = SQL & DBSet(RS!porciva1, "N") & "," & DBSet(RS!porciva2, "N", Nulo2) & "," & DBSet(RS!porciva3, "N", Nulo3) & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(RS!impoiva1, "N") & "," & DBSet(RS!impoiva2, "N", Nulo2) & "," & DBSet(RS!impoiva3, "N", Nulo3) & ","
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            'ANTES era dbset de Rs!totalfac, ahora lo haremos de la variabele totalfac
            SQL = SQL & DBSet(TotalFac, "N") & "," & DBSet(RS!TipoIVA1, "N") & "," & DBSet(RS!TipoIVA2, "N", Nulo2) & "," & DBSet(RS!TipoIVA3, "N", Nulo3) & ","
            
            
            
            'Enero 2011
            'Si el proveedor es INTRACOM, marco la de extranjero
            Nulo2 = DBLet(RS!tipprove, "N")
            If Nulo2 <> "1" Then Nulo2 = "0"
            
            'Abril 2015. ISP
            If DBLet(RS!InvSujPas, "N") = 1 Then Nulo2 = "3"
            
            
            SQL = SQL & Nulo2 & ","
            EsFacturaIntracom2 = ""
            If Nulo2 = "1" Then
                'OK es intracomunitaria
                EsFacturaIntracom2 = CStr(RS!anofacpr)
            End If
            
            
            
            
            
            'RETENCION.   29 MAYO 2008
            ' retfacpr,trefacpr,cuereten              Las facturas pueden llevar retencion
            Nulo2 = ""
            If RS!TipoRet = 0 Then
                If Not IsNull(RS!PorRet) And Not IsNull(RS!ImpRet) Then Nulo2 = "O"
            End If
            If Nulo2 = "" Then
                'NULOS
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
            Else
                'TIene valor
                SQL = SQL & DBSet(RS!PorRet, "N") & "," & DBSet(RS!ImpRet, "N") & ",'" & vParamAplic.CtaReten & "',"
            End If
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(RS!FecRecep, "F") & ",0"
            Cad = Cad & "(" & SQL & ")"
            
            'Insertar en la contabilidad
            SQL = "INSERT INTO cabfactprov (numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,"
            SQL = SQL & "pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,"
            SQL = SQL & "totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,numdiari,fechaent,numasien,fecliqpr,nodeducible) "
            SQL = SQL & " VALUES " & Cad
            ConnConta.Execute SQL
            
            
            
            
            'Para saber el numreo de registro que le asigna a la factrua
            SQL = "INSERT INTO tmpinformes (codusu,codigo1,nombre1,nombre2,importe1) VALUES (" & vUsu.codigo & "," & Mc.Contador
            SQL = SQL & ",'" & DevNombreSQL(RS!NumFactu) & " @ " & Format(RS!FecFactu, "dd/mm/yyyy") & "','" & DevNombreSQL(RS!nomprove) & "'," & RS!Codprove & ")"
            conn.Execute SQL
        End If
    End If
    RS.Close
    Set RS = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactProv = False
        cadErr = Err.Description
    Else
        InsertarCabFactProv = True
    End If
End Function



Public Sub FechasEjercicioConta(FIni As String, FFin As String)
'Dim RS As ADODB.Recordset
'
'    On Error GoTo EFechas
'
'    FIni = "Select fechaini,fechafin From parametros"
'    Set RS = New ADODB.Recordset
'    RS.Open FIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        FIni = DBLet(RS!FechaIni, "F")
'        FFin = DBLet(RS!FechaFin, "F")
'    End If
'    RS.Close
'    Set RS = Nothing
'
'EFechas:
'    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function InsertarLinFact_TicketsAgrupados(cadTabla As String, cadWhere As String, cadErr As String, LlevaRetencion As Boolean, Optional numRegis As Long) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim SQL As String
Dim SQLaux As String
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim Cad As String, Aux As String
Dim i As Byte
Dim TotImp As Currency, ImpLinea As Currency
Dim cadCampo As String
Dim FechaFac As Date

    On Error GoTo EInLinea
    
        
    
            
            
         'comprobar si el cliente utiliza cuenta alternativa
        If conCtaAlt Then
            'utilizamos sfamia.ctavent1 o sfamia.abovent1
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctavent1"
            Else
                cadCampo = "sfamia.abovent1" 'si es negativa es un abono
            End If
        Else
            'utilizamos sfamia.ctaventa o sfamia.aboventa
            If TotalFac >= 0 Then
                cadCampo = "sfamia.ctaventa"
            Else
                cadCampo = "sfamia.aboventa"
            End If
        End If
        
        
        'Monto el WHERE buscando los tikets que estan asociados a este numfact FTG
        SQLaux = Replace(cadWhere, "scafac.", "")
        SQLaux = Replace(SQLaux, "numfactu", "numfacftg")
        SQLaux = Replace(SQLaux, "fecfactu", "fecfacftg")
        SQLaux = "select sfactik.* from sfactik ,scafac where sfactik.numfacFTG=scafac.numfactu and sfactik.fecfacftg=scafac.fecfactu AND " & SQLaux
    
    
    
    
        Set RS = New ADODB.Recordset
        RS.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = RS!numfacftg & " as numfactu ,'" & Format(RS!fecfacftg, FormatoFecha) & "' as fecfactu,"
        FechaFac = RS!fecfacftg
        'En aux guardare el codtraba
        Aux = RS!CodTraba
        SQLaux = ""
        Do
            SQLaux = SQLaux & "," & RS!NumFactu
            RS.MoveNext
        Loop Until RS.EOF
        RS.Close
        
        
        
        
        
        SQL = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FTG", "T")
        SQL = " SELECT '" & SQL & "' as LetraSer,slifac.codtipom," & Cad & cadCampo & " as cuenta,sum(importel) as importe"
        
        'Tiene analitica. Luego el codtraba tiene que aparecer
        If vCCos > 0 Then SQL = SQL & "," & Aux & " as CodTraba"
        
        SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
        SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
        SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
        
        'David.
        'Lleva anal. Necesitare el trabajador para obtener el CC
        If vCCos > 0 Then SQL = SQL & " ,scafac1 "
        
        SQL = SQL & " WHERE "
        
        'Si lleva analitica
        If vCCos > 0 Then
            'Linkamos la tabla
            SQL = SQL & " slifac.codTipoM = scafac1.codTipoM And slifac.NumFactu = scafac1.NumFactu And slifac.FecFactu = scafac1.FecFactu"
            SQL = SQL & " and slifac.codtipoa=scafac1.codtipoa and slifac.numalbar=scafac1.numalbar AND "
        End If
        


        
        
        
        SQLaux = Mid(SQLaux, 2)
        SQLaux = "   slifac.codtipom='FTI' AND slifac.numfactu IN (" & SQLaux & ")"
        
        'Marzo 2011
        'El año y el mes de los tikets DEBE SER del de la fecha de factura
        SQLaux = " year(slifac.fecfactu)=" & Year(FechaFac) & " and month(slifac.fecfactu)=" & Month(FechaFac) & " AND " & SQLaux
        
        
        SQL = SQL & SQLaux
        SQL = SQL & " GROUP BY "
        
        'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
        If vCCos = 2 Then SQL = SQL & " codtraba, "
                  
        'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
        SQL = SQL & cadCampo
        
    
    

    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Cad = ""
    i = 1
    TotImp = 0
    SQLaux = ""
    While Not RS.EOF
        SQLaux = Cad
        'calculamos la Base Imp del total del importe para cada cta cble ventas
        '---- Laura: 10/10/2006
        'ImpLinea = RS!Importe - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoPPago)))
        ImpLinea = RS!Importe - CCur(CalcularPorcentaje(RS!Importe, DtoPPago, 2))
        'ImpLinea = ImpLinea - CCur(CalcularDto(CStr(RS!Importe), CStr(DtoGnral)))
        ImpLinea = ImpLinea - CCur(CalcularPorcentaje(RS!Importe, DtoGnral, 2))
        'ImpLinea = Round(ImpLinea, 2)
        '----
        TotImp = TotImp + ImpLinea
        
        'concatenamos linea para insertar en la tabla de conta.linfact
        SQL = ""
        SQL2 = ""
        

        SQL = "'" & RS!LetraSer & "'," & RS!NumFactu & "," & Year(RS!FecFactu) & "," & i & ","
        SQL = SQL & DBSet(RS!Cuenta, "T")
        

        
        SQL2 = SQL & "," 'nos guardamos la linea sin el importe por si a la última hay q descontarle para q coincida con total factura
        SQL = SQL & "," & DBSet(ImpLinea, "N") & ","
        
        If vCCos = 0 Then
            SQL = SQL & ValorNulo
        Else
            'Obtendremos el centro de coste a partir del trabajador
            CCoste2 = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", RS!CodTraba)
            If CCoste2 = "" Then
                cadErr = "ERROR en el centro de coste del trabajador: " & RS!CodTraba
                'CIerro el rs y salgo por patas
                RS.Close
                Set RS = Nothing
    
            End If
            SQL = SQL & DBSet(CCoste2, "T")
        End If
        
        
        Cad = Cad & "(" & SQL & ")" & ","
        
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close

    
    'comprtobar que la suma de los importes de las lineas insertadas suman la BImponible
    'de la factura
    If TotImp <> BaseImp Then
'        MsgBox "FALTA cuadrar bases imponibles!!!!!!!!!"
        'en SQL esta la ult linea introducida
        TotImp = BaseImp - TotImp
        TotImp = ImpLinea + TotImp '(+- diferencia)
        SQL2 = SQL2 & DBSet(TotImp, "N") & ","
        If CCoste2 = "" Then
            SQL2 = SQL2 & ValorNulo
        Else
            SQL2 = SQL2 & DBSet(CCoste2, "T")
        End If
        If SQLaux <> "" Then 'hay mas de una linea
            Cad = SQLaux & "(" & SQL2 & ")" & ","
        Else 'solo una linea
            Cad = "(" & SQL2 & ")" & ","
        End If
        
'        Aux = Replace(SQL, DBSet(ImpLinea, "N"), DBSet(TotImp, "N"))
'        cad = Replace(cad, SQL, Aux)
    End If



    'Si lleva retencion, solo sera en caso de facturas proveedores, entonces metere dos lineas mas
    '
    If LlevaRetencion Then
        'Cojere los datos del proveedor
        'Reutilizo total fac
        TotalFac = CCur(RecuperaValor(DatosRetencion, 2))
        i = i + 1
        SQL = "(" & numRegis & "," & AnyoFacPr & "," & i & ",'" & RecuperaValor(DatosRetencion, 1) & "'," & TransformaComasPuntos(CStr(-TotalFac)) & ",NULL)"
        Cad = Cad & SQL
        SQL = ",(" & numRegis & "," & AnyoFacPr & "," & i + 1 & ",'" & vParamAplic.CtaReten & "'," & TransformaComasPuntos(CStr(TotalFac)) & ",NULL),"
        Cad = Cad & SQL
    End If





    Set RS = Nothing

    'Insertar en la contabilidad
    If Cad <> "" Then
        Cad = Mid(Cad, 1, Len(Cad) - 1) 'quitar la ult. coma
        SQL = "INSERT INTO linfact (numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost) "
        SQL = SQL & " VALUES " & Cad
        ConnConta.Execute SQL
    End If




EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFact_TicketsAgrupados = False
        cadErr = Err.Description
    Else
        InsertarLinFact_TicketsAgrupados = True
    End If
End Function







'=============================================================================
'==========     CENTROS DE COSTE
'=============================================================================
'LAURA
Public Function PonerNombreCCoste(ByRef txt As TextBox) As String
'Obtener el nombre de un centro de coste
Dim codCCoste As String
Dim Cad As String

    If txt.Text = "" Then
         PonerNombreCCoste = ""
         Exit Function
    End If
    
    codCCoste = Trim(txt.Text)
    Cad = "cabccost"
    If vParamAplic.ContabilidadNueva Then Cad = "ccoste"
    Cad = DevuelveDesdeBDNew(conConta, Cad, "nomccost", "codccost", codCCoste, "T")
    If Cad = "" Then
        If Not txt.Locked Then MsgBox "No existe el Centro de coste : " & codCCoste, vbExclamation
        PonerNombreCCoste = ""
        txt.Text = ""
    Else
        txt.Text = codCCoste
        PonerNombreCCoste = Cad
    End If
    
End Function




Private Function CuentaNecesitaCentroCoste(cta As String) As Boolean
Dim i As Integer
Dim C As String
    
    CuentaNecesitaCentroCoste = False
    
    'vEmpresa.RaizAnalitica    lleva: gripo gasto |grupo vta| otros grupo
    For i = 1 To 3
        C = RecuperaValor(vEmpresa.RaizAnalitica, i)
        If i < 3 Then
            'UN DIGITO
            If Mid(cta, 1, 1) = C Then
                CuentaNecesitaCentroCoste = True
                Exit Function
            End If
        Else
            'Subgrupo a tres digitos
            If Mid(cta, 1, 3) = C Then
                CuentaNecesitaCentroCoste = True
                Exit Function
            End If
        End If
    Next i
End Function

'Dtos orig:
Private Function InsertarFrasEXTRAdeIntracomunitarias(NuevoContaPro As Long, datosOrig As String) As Boolean
Dim vT As CTiposMov
Dim Aux As String
Dim RN As ADODB.Recordset
Dim Base As Currency
Dim PorI As Currency
Dim TotI As Currency


    On Error GoTo eInsertarFrasEXTRAdeIntracomunitarias
    InsertarFrasEXTRAdeIntracomunitarias = False

    Set vT = New CTiposMov
    vT.Leer "CFI"    'NO puede dar error
    vT.ConseguirContador vT.TipoMovimiento
    
    Set RN = New ADODB.Recordset
    
    Aux = "Select * from tiposiva where codigiva = " & vParamAplic.IvaIntracomAdicional
    RN.Open Aux, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    PorI = RN!PorceIVA
    RN.Close
    
    
    Aux = "Select * from cabfactprov where anofacpr= " & RecuperaValor(datosOrig, 2) & " AND numregis= " & RecuperaValor(datosOrig, 1)
    RN.Open Aux, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    Base = RN!ba1facpr 'SOLO DEBE TENER UNA BASE ya que es exento la original
    TotI = Round2((Base * PorI) / 100, 2)

    'Inserto la de clientes
    Aux = "INSERT INTO cabfact (Numserie , Codfaccl, fecfaccl, Codmacta, anofaccl, confaccl, ba1faccl, pi1faccl, ti1faccl, totfaccl, tp1faccl, intracom, fecliqcl, numasien)"
    Aux = Aux & " VALUES ('" & vT.LetraSerie & "'," & vT.Contador + 1 & "," & DBSet(RN!fecrecpr, "F") & ",'" & vParamAplic.CtaContabIntracom & "',"
    'anofaccl, confaccl, ba1faccl, pi1faccl, ti1faccl, totfaccl, tp1faccl, intracom, fecliqcl, numasien)
    Aux = Aux & RN!anofacpr & ",'AUTOFACTURA'," & DBSet(Base, "N") & "," & DBSet(PorI, "N") & "," & DBSet(TotI, "N") & ","
    Aux = Aux & DBSet(Base + TotI, "N") & "," & vParamAplic.IvaIntracomAdicional & ",0," & DBSet(RN!fecliqpr, "F") & ","   'NO ENTRA COMO intracom  08/04/2011
    'Si lleva la marca de contabilizada
    If vParamAplic.IntracomAdicionalContab Then
        Aux = Aux & "NULL"
    Else
        Aux = Aux & "0"  'pondre en numasien un cero
    End If
    Aux = Aux & ")"
    ConnConta.Execute Aux
    
    
    'Inserto la factura PROVEEDORES
    Aux = "INSERT INTO cabfactprov(numregis,numfacpr,anofacpr,fecfacpr,fecrecpr,codmacta,confacpr,ba1facpr,pi1facpr,ti1facpr,totfacpr,tp1facpr,extranje,fecliqpr,numasien)"
    Aux = Aux & " VALUES (" & NuevoContaPro & ",'" & vT.LetraSerie & Format(vT.Contador + 1, "0000000") & "'," & RN!anofacpr & ","
    'fecfacpr,fecrecpr,codmacta,
    Aux = Aux & DBSet(RN!fecfacpr, "F") & "," & DBSet(RN!fecrecpr, "F") & ",'" & vParamAplic.CtaContabIntracom & "',"
    Aux = Aux & "'AUTOFACTURA'," & DBSet(Base, "N") & "," & DBSet(PorI, "N") & "," & DBSet(TotI, "N") & ","
    Aux = Aux & DBSet(Base + TotI, "N") & "," & vParamAplic.IvaIntracomAdicional & ",0," & DBSet(RN!fecliqpr, "F") & ","   'NO ENTRA COMO intracom  08/04/2011
    If vParamAplic.IntracomAdicionalContab Then
        Aux = Aux & "NULL"
    Else
        Aux = Aux & "0"  'pondre en numasien un cero
    End If
    Aux = Aux & ")"
    ConnConta.Execute Aux
    
    
    
    
    
    'Las lineas de la fra prov y clie "EXTRAS"
    Aux = "INSERT INTO linfact(numserie,codfaccl,anofaccl,numlinea,codtbase,impbascl,codccost)"
    Aux = Aux & " Select '" & vT.LetraSerie & "'," & vT.Contador + 1 & ",anofacpr,numlinea,codtbase,impbaspr,codccost"
    Aux = Aux & " FROM linfactprov WHERE numregis = " & RN!numRegis & " AND anofacpr = " & RN!anofacpr
    ConnConta.Execute Aux
    
    Aux = "INSERT INTO linfactprov(numregis,anofacpr,numlinea,codtbase,impbaspr,codccost)"
    Aux = Aux & " Select " & NuevoContaPro & ",anofacpr,numlinea,codtbase,impbaspr,codccost"
    Aux = Aux & " FROM linfactprov WHERE numregis = " & RN!numRegis & " AND anofacpr = " & RN!anofacpr
    ConnConta.Execute Aux
    
    RN.Close

    'Incrementaremos el contador
    vT.IncrementarContador vT.TipoMovimiento
    InsertarFrasEXTRAdeIntracomunitarias = True
    
eInsertarFrasEXTRAdeIntracomunitarias:
    If Err.Number <> 0 Then
        Aux = Err.Description & vbCrLf & Aux & vbCrLf & "El proceso continuará"
        Aux = "Error FRAS INTRACOMUNITARIAS" & vbCrLf & Aux
        MsgBox Aux, vbExclamation
    End If
    Set vT = Nothing
    Set RN = Nothing
End Function




'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'
'   Contabilizacion de las facturas internas (FAI)
'
'   - No inserta en cabfact(ni linfact).  Mete un apunte
'       43000   contra las 70000 que deriven de las familias
'
Private Function ContabilizaFAI(cadError As String, cadWhere As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
Dim Rc As ADODB.Recordset
Dim RL As ADODB.Recordset
Dim SQL As String
Dim Aux As String

    On Error GoTo eContabilizaFAI
    ContabilizaFAI = False
    cadError = ""
    SQL = " SELECT stipom.letraser,numfactu,fecfactu, sclien.codmacta,sclien.cliabono,year(fecfactu) as anofaccl,"
    SQL = SQL & "scafac.dtoppago,scafac.dtognral,baseimp1,baseimp2,baseimp3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,"
    SQL = SQL & "totalfac,codigiv1,codigiv2,codigiv3,aportacion "
    
    'Cuando MIS facfuras llevan recargo equivalencia
    SQL = SQL & ",porciva1re,porciva2re,porciva3re,imporiv1re,imporiv2re,imporiv3re,tipoiva"
    
    SQL = SQL & " FROM (" & "scafac inner join " & "stipom on scafac.codtipom=stipom.codtipom) "
    SQL = SQL & "INNER JOIN " & "sclien ON scafac.codclien=sclien.codclien "
    SQL = SQL & " WHERE " & cadWhere
    Set Rc = New ADODB.Recordset
    Rc.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DtoPPago = Rc!DtoPPago
    DtoGnral = Rc!DtoGnral
    BaseImp = Rc!baseimp1
    TotalFac = Rc!TotalFac
    DatosAportacion = ""
    conCtaAlt = Rc!cliAbono
    
    
    
    'Para las lineas de factura
    '-------------------------------
    If conCtaAlt Then
        'utilizamos sfamia.ctavent1 o sfamia.abovent1
        If TotalFac >= 0 Then
            Aux = "sfamia.ctavent1"
        Else
            Aux = "sfamia.abovent1" 'si es negativa es un abono
        End If
    Else
        'utilizamos sfamia.ctaventa o sfamia.aboventa
        If TotalFac >= 0 Then
            Aux = "sfamia.ctaventa"
        Else
            Aux = "sfamia.aboventa"
        End If
    End If
    
    
    SQL = " SELECT stipom.letraser,slifac.codtipom,slifac.numfactu,slifac.fecfactu," & Aux & " as cuenta,sum(importel) as importe"
    'Tiene analitica. Luego el codtraba tiene que aparecer
    If vCCos > 0 Then SQL = SQL & ",slifac.codccost"
    SQL = SQL & " FROM ((slifac inner join stipom on slifac.codtipom=stipom.codtipom) "
    SQL = SQL & " inner join sartic on slifac.codartic=sartic.codartic) "
    SQL = SQL & " inner join sfamia on sartic.codfamia=sfamia.codfamia "
    SQL = SQL & " WHERE "
    SQL = SQL & " " & Replace(cadWhere, "scafac", "slifac")
    SQL = SQL & " GROUP BY "
    'Si tiene mas de una trabajador con ditintos CC agrupamos en 1er nivel por codtraba
    If vCCos > 0 Then SQL = SQL & " codccost, "
    'Agrupemos por trabajador o no, tambien agrupamos por la cuenta
    SQL = SQL & Aux
    Set RL = New ADODB.Recordset
    RL.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText


    cadError = vContaFra.IntegraLaFacturaClienteINTERNA(Rc, RL)

    
eContabilizaFAI:
    If Err.Number <> 0 Then
        cadError = Err.Description
        Err.Clear
    End If
    If cadError = "" Then ContabilizaFAI = True
    
End Function
