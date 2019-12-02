Attribute VB_Name = "ModTMP"
Option Explicit

'MODULO PARA LA CARGA Y DESCARGA DE TABLAS TEMPORALES

'LO tengo que crear "global"
Dim codtipom(4) As String  'Probamos por EULER


'================================================================================
'================================================================================

'================================================================================
'TMPnseries: Temporal para introducir los Nº de Serie de los Articulos en compras o en ventas
'USO: frmFacEntAlbaran, frmRepEntAlbaran
'================================================================================

Public Function CargarDatosTMPNumSeries(NomTabla As String, codArtic As String, Cant As Integer, NumLinAlb As String) As Boolean
'IN -> NomTabla: Nombre de la tabla temporal
'      CodArtic: Codigo Articulo del que se van a Introducir los Nº de Serie
'      Cant: cantidad de Articulo (tantas filas como articulos)
'      Mostrar: si true se cargar los Nº de serie sino en blanco
Dim SQL As String
Dim I As Integer
Dim numlinea As String, vWhere As String

    On Error GoTo ECargaDatosTMP

    'Insertar tantos registros como cantidad de Articulo Introducida
    vWhere = "(codusu=" & vUsu.Codigo & " and codartic=" & DBSet(codArtic, "T") & " and numlinealb=" & DBSet(NumLinAlb, "N") & ")"

    'insertamos tantos num.serie como cantidad
    For I = 0 To Cant - 1
        'Obtener Num Linea
        numlinea = SugerirCodigoSiguienteStr(NomTabla, "numlinea", vWhere)
        'Insertar en la temporal para Nº Series
        SQL = "INSERT INTO " & NomTabla & " (codusu, codartic, numlinealb, numlinea, numserie,nummante) VALUES ("
        SQL = SQL & vUsu.Codigo & ", " & DBSet(codArtic, "T") & ", " & NumLinAlb & ", " & numlinea & ", ' ',' ')"
        conn.Execute SQL
    Next I
 
ECargaDatosTMP:
    If Err.Number <> 0 Then
        CargarDatosTMPNumSeries = False
        MuestraError Err.Number, "Numeros Serie", Err.Description
    Else
        CargarDatosTMPNumSeries = True
    End If
End Function


Public Function DescargarDatosTMPNumSeries(NomTabla As String)
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

     '------------- AHORA
    SQL = "DELETE from " & NomTabla & " where codusu= " & vUsu.Codigo
    conn.Execute SQL
    
    Exit Function
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal (Nº Serie).", Err.Description
End Function



Public Function InsertarNSeries(codArtic As String, CadValuesI As String, cadValuesU As String, DeVenta As Boolean) As Boolean
'Insertar un registro en la tabla "sserie" por cada uno de los
'Nº de Serie introducidos en la Tabla Temporal
Dim RS As ADODB.Recordset
Dim SQL As String, devuelve As String
Dim codTipar As String, Numalbar As String
    
    On Error GoTo EInsertar

    'Seleccionar los nº de serie cargados en la temporal: tmpnseries
    SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo & " AND codartic=" & DBSet(codArtic, "T")
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then RS.MoveFirst
    
    While Not RS.EOF
        'Comprobar si existe en la tabla sserie
        If DeVenta Then
            Numalbar = "numalbar" 'Nº albaran de Venta
        Else
            Numalbar = "numalbpr" 'Nº albaran de Compras
        End If
        devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", RS!numSerie, "T", Numalbar, "codartic", RS!codArtic, "T")
        If devuelve <> "" Then 'Existe en tabla sserie
            If Numalbar = "" Then
                SQL = Trim(DBLet(RS!nummante, "T"))
                If SQL = "" Then
                    SQL = "nummante = NULL, tieneman = 0,"
                Else
                    SQL = "nummante = " & DBSet(SQL, "T") & " , tieneman = 1,"
                End If
            
            
            
                'David. Abril 2012.  El SQL se lo esta comiendo
                'SQL = "UPDATE sserie SET " & cadValuesU   'estaba este
                SQL = "UPDATE sserie SET " & SQL & cadValuesU
                
                '=== David 22/12/2011
                'Nummante viene de la tmp
                
                
                
                
                
                
                '=== Laura 17/01/2007
                SQL = SQL & " WHERE numserie=" & DBSet(RS!numSerie, "T") & " AND codartic=" & DBSet(RS!codArtic, "T")
                '===
            End If
        Else
            'Obtener el tipo de Articulo
            codTipar = DevuelveDesdeBDNew(conAri, "sartic", "codtipar", "codartic", RS!codArtic, "T")
        
            'Insertar en la tabla sserie
            SQL = "INSERT INTO sserie (numserie, codartic, codtipar, codclien, coddirec,tieneman, nummante, ultrepar, fingaran, "
            SQL = SQL & " codtipom, numfactu, fechavta, numalbar, numline1, codprove, numalbpr, fechacom, numline2) "
            SQL = SQL & " VALUES ( " & DBSet(RS!numSerie, "T") & ", " & DBSet(RS!codArtic, "T") & ", " & DBSet(codTipar, "T") & ","
            SQL = SQL & CadValuesI
            SQL = SQL & ") "
        End If
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
EInsertar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando Nº Serie", Err.Description
End Function





Public Sub PedirNSeriesGnral(ByRef RS As ADODB.Recordset, Men As Boolean)
Dim SQL As String
Dim b As Boolean

    On Error GoTo EPedirNSeries

        If Men Then
            SQL = "Hay artículos que tienen control de Nº de Serie." & vbCrLf & vbCrLf
            SQL = SQL & "Introduzca los Nº De Serie." & vbCrLf
            MsgBox SQL, vbInformation
        End If
        
        'Cargar la tabla temporal con tantas filas como cantidad de Articulo
        'Para introducir el Nº de Serie
        DescargarDatosTMPNumSeries ("tmpnseries")
        b = True
        
        While Not RS.EOF
            If Not CargarDatosTMPNumSeries("tmpnseries", RS!codArtic, RS!cantidad, RS!numlinea) Then
                b = False
            End If
            RS.MoveNext
        Wend
        
        'Visualizar en pantalla el Grid, y rellenar los Nº Serie
        If Not b Then MsgBox "No se han podido mostrar todos los Artículos con Nº de Serie.", vbInformation
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Public Function MostrarNSeriesGnral(ByRef RSLineas As ADODB.Recordset, vCampos As String, Optional Rectifica As Boolean) As String
'Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
'los Nº de serie de los articulos comprados y se seleccionan tantos como cantidad de la linea
'IN -> RSLineas: lineas del Albaran generado
'OUT -> vCampos: concatena la cantidad requerida de Nº series de cada articulo
'RETURN -> cadena SQL con la Select que se pasara para mostrar los Nº series
Dim SQL As String
Dim cadArtic As String
Dim Campos As String
Dim totArtic As Integer

    On Error GoTo EMostrar

    'Concatenamos los codigos de Articulo que tenemos que seleccionar se "sseries"
    cadArtic = ""
    totArtic = 0
    Campos = ""
    While Not RSLineas.EOF
        Campos = Campos & RSLineas!codArtic & "|" & RSLineas!cantidad & "·"
        If cadArtic = "" Then
            cadArtic = DBSet(RSLineas!codArtic, "T")
        Else
            cadArtic = cadArtic & ", " & DBSet(RSLineas!codArtic, "T")
        End If
        totArtic = totArtic + 1
        RSLineas.MoveNext
    Wend
    RSLineas.MoveFirst
    cadArtic = "(" & cadArtic & ")"
    vCampos = Campos
   
    'Se puede seleccionar todos los Nº de Serie que se necesitan
     'se introdujo los Nº de Serie en COMPRAS y ahora
    'mostramos los Nº de Serie para seleccionar cual vamos a
    'vender al Cliente
    Screen.MousePointer = vbDefault
    If Rectifica Then
        'viene de una factura rectificativa los nº de serie que seleccionemos sera para quitar
        SQL = "Hay Artículos que tienen control de Nº de Serie." & vbCrLf
        SQL = SQL & "Seleccione los nº de Serie que desea rectificar."
    Else
        If totArtic > 1 Then
            SQL = "Hay Artículos que tienen control de Nº de Serie." & vbCrLf
            SQL = SQL & vbCrLf & "Seleccione un Nº de Serie para cada Artículo."
        Else
            SQL = "El Artículo tienen control de Nº de Serie." & vbCrLf
            SQL = SQL & vbCrLf & "Seleccione los Nº de Serie para el Artículo."
        End If
    End If
    MsgBox SQL, vbInformation
    SQL = " WHERE sserie.codartic IN " & cadArtic
    MostrarNSeriesGnral = SQL
    
EMostrar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Mostrar Nº series", Err.Description
End Function


'================================================================================
'TMPStockFec: Temporal para obtener el Stock que habia en una determinada fecha
'(Para Listado de Almacenes: "Inf. Stock a una Fecha")
'USO: frmListados
'================================================================================
'PositivosNegativos: 0.-todos     1.- +     2.-  menos
Public Function CargarTMPStockFecha(vSQL As String, cadFecha As String, cadHora As String, PositivosNegativos As Byte, ByRef Lbl As Label) As Boolean
'Carga la tabla temporal con el Stock del almacen seleccionado
'de los articulos seleccionados que habia a una determinada FECHA, HORA
Dim RS As ADODB.Recordset
Dim vStock As Single
Dim cadSQL As String
Dim Insertar As Boolean
Dim Entradas As Currency
Dim Salidas As Currency

    On Error GoTo ECargarTMPStock

    CargarTMPStockFecha = False
    
    Lbl.Caption = "Leyendo registros"
    Lbl.Refresh
    Set RS = New ADODB.Recordset
    RS.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Lbl.Caption = "Art  " & RS!codArtic
        Lbl.Refresh
        'Para cada articulo obtener el stock en esa fecha e insertarlo en la temporal
        'Antes abril 2011
        'cadSQL = "SELECT sum(cantidad) FROM smoval WHERE "
        cadSQL = "SELECT tipomovi,sum(cantidad) FROM smoval WHERE "
   
        cadSQL = cadSQL & " codartic=" & DBSet(RS!codArtic, "T") & " AND codalmac=" & RS!codAlmac
        
            vStock = RS!CanStock
            '- Deshacer los movimientos entre esas fecha alreves
            If cadHora = "" Then
                cadSQL = cadSQL & " AND fechamov> '" & Format(cadFecha, FormatoFecha) & "' "
            Else
                cadSQL = cadSQL & " AND horamovi> '" & Format(cadFecha & " " & cadHora, FormatoFechaHora) & "' "
            End If
            
            
            'ANTES ABRIL 2011
            'Movimientos de ENTRADA (tipomovi=1)
            'vStock = vStock - TotMovimientosStock2(cadSQL, 1)
            'Movimientos de SALIDA (tipomovi=0)
            'vStock = vStock + TotMovimientosStock2(cadSQL, 0)
            'AHORA
            TotMovimientosStockAgrup cadSQL, Entradas, Salidas
            vStock = vStock - Entradas + Salidas
'        End If
        '##

        '-- Insertar en la Tabla TMP el stock en esa Fecha del codartic,codalmac
        Insertar = True
        If PositivosNegativos > 0 Then
            If PositivosNegativos = 1 Then
                If vStock < 0 Then Insertar = False
            Else
                If vStock > 0 Then Insertar = False
            End If
        End If
        If Insertar Then
            cadSQL = "INSERT INTO tmpstockfec (codusu,codartic,codalmac,stock)"
            cadSQL = cadSQL & " VALUES (" & vUsu.Codigo & ", " & DBSet(RS!codArtic, "T") & ", "
            cadSQL = cadSQL & RS!codAlmac & ", " & TransformaComasPuntos(CStr(vStock)) & ")"
            conn.Execute cadSQL
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    cadSQL = "Select count(*) from tmpstockfec where codusu =" & vUsu.Codigo
    RS.Open cadSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then cadSQL = ""
    End If
    RS.Close
    
    Set RS = Nothing
    
    If cadSQL <> "" Then
        MsgBox "No hay datos para mostrar con estos parametros", vbExclamation
    Else
        CargarTMPStockFecha = True
    End If
    
ECargarTMPStock:
    If Err.Number <> 0 Then
'        RS.Close
        Set RS = Nothing
        MsgBox " No se ha podido cargar la Tabla Temporal correctamente", vbInformation
    End If
    Lbl.Caption = ""
End Function



Public Function DescargarDatosTMPStockFecha()
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

    '------------- AHORA
    SQL = "DELETE from tmpstockfec" & " where codusu= " & vUsu.Codigo
    conn.Execute SQL
    Exit Function
    
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal (Stock a Fecha).", Err.Description
End Function



Private Function TotMovimientosStock2(cadSQL As String, vTipomovi As Byte) As Single
'Para un tipo de Movimiento vtipomovi(0=Salida, 1=Entrada) devolver
'la cantidad de stock para esos registros de la select
Dim RSmov As ADODB.Recordset
Dim Cad As String

        TotMovimientosStock2 = 0
        Cad = cadSQL & " AND tipomovi=" & vTipomovi
        
        Set RSmov = New ADODB.Recordset
        RSmov.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RSmov.EOF Then
            If Not IsNull(RSmov.Fields(0).Value) Then _
                TotMovimientosStock2 = RSmov.Fields(0).Value
        End If
        
        RSmov.Close
        Set RSmov = Nothing

End Function
Private Sub TotMovimientosStockAgrup(cadSQL As String, ByRef Entradas As Currency, ByRef Salidas As Currency)
'Para un tipo de Movimiento vtipomovi(0=Salida, 1=Entrada) devolver
'la cantidad de stock para esos registros de la select
Dim RSmov As ADODB.Recordset
Dim Cad As String

    
        Entradas = 0
        Salidas = 0
        Cad = cadSQL & " GROUP BY tipomovi"
        
        Set RSmov = New ADODB.Recordset
        RSmov.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        While Not RSmov.EOF
            'El primero
            If RSmov.Fields(0) = 1 Then
                'Entrada
                Entradas = DBLet(RSmov.Fields(1), "N")
            Else
                Salidas = DBLet(RSmov.Fields(1), "N")
            End If
            RSmov.MoveNext
        Wend
        RSmov.Close
        Set RSmov = Nothing

End Sub





'============ Temporales de INFORMES ====================================
'JULIO 2013
'Añadimos el poder mostrar que tipos de facturas ENTRAN
Public Function TempVentasClientes(PorAgente As Boolean, cadSel As String, cadSelPeriodo As String, cadSelAnte As String, ByRef Lbl As Label, TiposDeFacturas As String) As Boolean
'Inserta en la temporal TMPINFORMES
Dim SQL As String, SQL2 As String
Dim SQLinsert As String
Dim RS As ADODB.Recordset
Dim Cliente As String
Dim T1 As Currency
Dim T2 As Currency
Dim T3 As Currency
Dim T4 As Currency
Dim T5 As Currency
Dim total As String 'Total del periodo seleccionado
Dim totalAnt As String 'total del periodo anterior
Dim ColAgent As Collection
Dim J As Long




    On Error GoTo ETmpVentas
    
    
    'Vemos que facturas tratamos

    J = 0
    Cliente = ""
    totalAnt = "" 'para no declarar mas variables
    SQL = TiposDeFacturas
    While SQL <> ""
       J = InStr(1, SQL, "|")
       If J = 0 Then
            SQL = ""
        Else
           If Not PorAgente Then codtipom(Len(totalAnt)) = Mid(SQL, 1, J - 1)
           Cliente = Cliente & ",'" & Mid(SQL, 1, J - 1) & "'"
           SQL = Mid(SQL, J + 1)
           totalAnt = totalAnt & "X"
        End If
    Wend
    TiposDeFacturas = "(" & Mid(Cliente, 2) & ")"
    J = Len(totalAnt)
    
    While J < 4
        codtipom(J) = "XXX"
        J = J + 1
    Wend
    
    
    
    'Obtenemos el TOTAL de ventas en ese PERIODO, de todos los clientes.
    'para obtener el % de ventas de cada cliente
    '---------------------------------------------------------------------
    Lbl.Caption = "Obteniendo importes"
    Lbl.Refresh
    SQL = "select sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
    SQL = SQL & " FROM scafac,sclien "
    
    'FEBRERO 2011
    'Esto NO estabam, y obivamente falta
    SQL = SQL & "  WHERE scafac.codclien=sclien.codclien "
    If cadSelPeriodo <> "" Then SQL = SQL & " AND " & cadSelPeriodo
    
    
    
    
    
    'JULIO 2013
    'Marzo 2011
    'hay que quitar el B
    'SQL = SQL & "  AND scafac.codtipom <> 'FAZ' "
    SQL = SQL & "  AND scafac.codtipom IN  " & TiposDeFacturas
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        total = CStr(DBLet(RS.Fields(0), "N"))
    End If
    RS.Close
    Set RS = Nothing
     
    
    
    'Obtenemos el TOTAL de ventas en el PERIODO ANTERIOR, de todos los clientes.
    'para obtener el % de ventas de cada cliente
    '---------------------------------------------------------------------------
    If cadSelAnte <> "" Then
        DoEvents
        Lbl.Caption = "Obteniendo importes ant."
        Lbl.Refresh

        SQL = "select sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
        SQL = SQL & " FROM scafac,sclien "
        SQL = SQL & "  WHERE scafac.codclien=sclien.codclien "
        
        
        If cadSelAnte <> "" Then SQL = SQL & " AND " & cadSelAnte
        
        'JUL2013
        'Marzo 2011
        'hay que quitar el B
        'SQL = SQL & "  AND scafac.codtipom <> 'FAZ' "
        SQL = SQL & "  AND scafac.codtipom IN  " & TiposDeFacturas
        
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            totalAnt = CStr(DBLet(RS.Fields(0), "N"))
        End If
        RS.Close
        Set RS = Nothing
    End If
    
    
    'Seleccion del PERIODO seleccionado
    'ventas por cliente y tipo de movimiento
    '----------------------------------------
    DoEvents
    Lbl.Caption = "Obteniendo importes periodo"
    Lbl.Refresh
    
    SQL = "select codtipom,scafac.codclien,scafac.nomclien,sum(baseimp1),sum(baseimp2),sum(baseimp3),sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
    SQL = SQL & " FROM scafac,sclien "
    SQL = SQL & "  WHERE scafac.codclien=sclien.codclien "
    If cadSel <> "" Then SQL = SQL & " AND " & cadSel
    'JUL2013
    SQL = SQL & "  AND scafac.codtipom IN  " & TiposDeFacturas
    
    SQL = SQL & " group by codclien,codtipom "
    SQL = SQL & " order by codclien"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQLinsert = "INSERT INTO tmpinformes (codusu, codigo1,nombre1,importe1,importe2,importe3,importe4,importe5,porcen1,importeb1,importeb2,importeb3,importeb4,importeb5) "
    SQLinsert = SQLinsert & " VALUES "
    
    SQL = ""
    SQL2 = ""
    J = 0
    While Not RS.EOF
        If Cliente <> RS!codClien Then
            Lbl.Caption = "Reg. cliente: " & RS!codClien
            Lbl.Refresh
            If SQL <> "" Then
                
                If PorAgente Then T5 = T1 + T4
                
                SQL = SQL & DBSet(T1, "N") & "," & DBSet(T2, "N") & "," & DBSet(T3, "N") & "," & DBSet(T4, "N") & "," & DBSet(T5, "N") & ","
                
                If PorAgente Then
                    T1 = T1 + T4
                Else
                    T1 = T1 + T2 + T3 + T4 + T5
                End If
                '----
                'SQL = SQL & DBSet(T1, "N") & ","  AHOR el total lo caclula en el rpt
                '% sobre el total de ventas
                T1 = Round((T1 * 100) / CCur(total), 2)
                SQL = SQL & DBSet(T1, "N") & ","
                'Obtener ventas del cliente para el periodo anterior
                SQL = SQL & VentasPeriodoAnterior(PorAgente, Cliente, cadSelAnte, TiposDeFacturas) & ")"
                SQL2 = SQL2 & SQL & ","
            End If
            'Insertamos por bloques de 500
            If J = 30 Then
                'Insertamos en la tabla temporal
                If SQL2 <> "" Then
                    SQL2 = Mid(SQL2, 1, Len(SQL2) - 1)
                    SQL = SQLinsert & SQL2
                    conn.Execute SQL
                End If
                
                'Reiniciamos los valores
                SQL = ""
                SQL2 = ""
                J = 0
            End If
            
            'Empezamos el registro para el siguiente cliente
            SQL = "(" & vUsu.Codigo & "," & RS!codClien & "," & DBSet(RS!NomClien, "T") & ","
            T1 = 0
            T2 = 0
            T3 = 0
            T4 = 0
            T5 = 0
            J = J + 1
        End If
    
        'AHORA
        If PorAgente Then
            
            If RS!codtipom = "FRT" Then
                T4 = T4 + RS!BaseImp
            Else
                T1 = T1 + RS!BaseImp
            End If
        Else
            If RS!codtipom = codtipom(0) Then
                T1 = T1 + RS!BaseImp
            ElseIf RS!codtipom = codtipom(1) Then
                T2 = T2 + RS!BaseImp
            ElseIf RS!codtipom = codtipom(2) Then
                T3 = T3 + RS!BaseImp
            ElseIf RS!codtipom = codtipom(3) Then
                T4 = T4 + RS!BaseImp
            ElseIf RS!codtipom = codtipom(4) Then
                T5 = T5 + RS!BaseImp
            End If
            
        End If
        Cliente = RS!codClien
        RS.MoveNext
    Wend
    
    RS.Close
    
    
    If SQL <> "" Then 'para el ultimo registro
        
                
        If PorAgente Then T5 = T1 + T4
        
        SQL = SQL & DBSet(T1, "N") & "," & DBSet(T2, "N") & "," & DBSet(T3, "N") & "," & DBSet(T4, "N") & "," & DBSet(T5, "N") & ","
        
        If PorAgente Then
            T1 = T1 + T4
        Else
            T1 = T1 + T2 + T3 + T4 + T5
        End If
        
        T1 = CStr(Round((CCur(T1) * 100) / CCur(total), 2))
        SQL = SQL & DBSet(T1, "N") & ","
        'Obtener ventas del cliente para el periodo anterior
        SQL = SQL & VentasPeriodoAnterior(PorAgente, Cliente, cadSelAnte, TiposDeFacturas) & ")"
        SQL2 = SQL2 & SQL & ","
    End If
    
    If SQL2 <> "" Then
        SQL2 = Mid(SQL2, 1, Len(SQL2) - 1)
        SQL = SQLinsert & SQL2
        conn.Execute SQL
    End If
    
    
    
    'MAYO 2016
    ' Estamos viendo para cad cliente del periodo actual, lo que compro en el periodo anterior
    
    
    
    
    
    
    
    
    cadSelPeriodo = DBSet(total, "N")
    cadSelAnte = DBSet(totalAnt, "N")
    
    
    
    
    'Para no tocar la funcioin de arriba, ahora recorro tmoinformes en codigo y le pongo en campo1 el agente
    Lbl.Caption = "Updatear tmp"
    Lbl.Refresh
    SQL = "Select codagent from tmpinformes,sclien where tmpinformes.codigo1 = sclien.codclien and  codusu = " & vUsu.Codigo & " GROUP BY 1"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColAgent = New Collection
    While Not RS.EOF
        SQL = RS.Fields(0)
        ColAgent.Add SQL
        RS.MoveNext
    Wend
    RS.Close
    
    For J = 1 To ColAgent.Count
        Lbl.Caption = J & " de " & ColAgent.Count
        Lbl.Refresh
        SQL = "UPDATE tmpinformes,sclien SET tmpinformes.campo1=" & ColAgent.Item(J) & " where tmpinformes.codigo1 = sclien.codclien and  codusu = " & vUsu.Codigo & " AND codagent=" & ColAgent.Item(J)
        conn.Execute SQL
    Next J
  
  
  
  
  
  
  
  
  
    'CLIENTES VARIOS
    Lbl.Caption = "Updatear tmp clivar"
    Lbl.Refresh


    SQL = "Select codigo1,nombre1,nomclien from tmpinformes,sclien where tmpinformes.codigo1 = sclien.codclien"
    SQL = SQL & " and clivario=1"
    SQL = SQL & " and  codusu = " & vUsu.Codigo & " GROUP BY 1"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColAgent = New Collection
    While Not RS.EOF
        SQL = "nombre1= " & DBSet(RS!NomClien, "T") & " WHERE codusu = " & vUsu.Codigo & " AND codigo1=" & RS!Codigo1
        ColAgent.Add SQL
        RS.MoveNext
    Wend
    RS.Close
    
    For J = 1 To ColAgent.Count
        Lbl.Caption = J & " de " & ColAgent.Count
        Lbl.Refresh
        SQL = "UPDATE tmpinformes SET " & ColAgent.Item(J)
        conn.Execute SQL
    Next J
  
  
    
    
    
    
    
ETmpVentas:
    Set RS = Nothing
    If Err.Number <> 0 Then
        TempVentasClientes = False
        MuestraError Err.Number, "Ventas del periodo", Err.Description
    Else
        TempVentasClientes = True
    End If
    Lbl.Caption = ""
    Set ColAgent = Nothing
    Screen.MousePointer = vbDefault
End Function


Private Function VentasPeriodoAnterior(PorAgente As Boolean, Cliente, cadSel, TiposDeFacturaAListar As String) As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim T1 As Currency
Dim T2 As Currency
Dim T3 As Currency
Dim T4 As Currency
Dim T5 As Currency

    On Error GoTo EVentas
    
    T1 = 0
    T2 = 0
    T3 = 0
    T4 = 0
    T5 = 0
    If cadSel <> "" Then
        '---- Laura: Modificaco 26/09/2006
        'SQL = "select codclien,codtipom,sum(baseimp1)+sum(baseimp2)+sum(baseimp3) as BaseImp "
        SQL = "SELECT scafac.codclien,codtipom, sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
        '----
        SQL = SQL & " from scafac,sclien where scafac.codclien =sclien.codclien AND " & cadSel
        'JUL2013
        SQL = SQL & "  AND scafac.codtipom IN  " & TiposDeFacturaAListar
        
        
        If cadSel <> "" Then SQL = SQL & " AND "
        SQL = SQL & "(scafac.codclien = " & Cliente & ")"
        
        SQL = SQL & " group by codclien,codtipom "
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
             'Select Case RS!Codtipom
             '   Case "FAV", "FTI", "FAS", "FMO", "FAI": totVentas = totVentas + RS!BaseImp
             '   Case "FAM": totMante = RS!BaseImp
             '   Case "FAR": totRepar = RS!BaseImp
             '   Case "FRT": totRectif = RS!BaseImp
             '   Case Else
             '

             'End Select
            If PorAgente Then
                
                If RS!codtipom = "FRT" Then
                    T4 = T4 + RS!BaseImp
                Else
                    T1 = T1 + RS!BaseImp
                End If
            Else
                If RS!codtipom = codtipom(0) Then
                    T1 = T1 + RS!BaseImp
                ElseIf RS!codtipom = codtipom(1) Then
                    T2 = T2 + RS!BaseImp
                ElseIf RS!codtipom = codtipom(2) Then
                    T3 = T3 + RS!BaseImp
                ElseIf RS!codtipom = codtipom(3) Then
                    T4 = T4 + RS!BaseImp
                ElseIf RS!codtipom = codtipom(4) Then
                    T5 = T5 + RS!BaseImp
                End If
            End If
                         
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End If
    
    
    If PorAgente Then T5 = T1 + T4
       
    SQL = DBSet(T1, "N") & "," & DBSet(T2, "N") & "," & DBSet(T3, "N") & "," & DBSet(T4, "N") & "," & DBSet(T5, "N")
    
    VentasPeriodoAnterior = SQL
    
EVentas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Ventas periodo anterior", Err.Description
    End If
End Function





Public Function TempVentasMeses(cadSel As String, Anyo As String) As Boolean
'Inseta en la tabla temporal TMPINFORMES
Dim SQL As String
Dim RS As ADODB.Recordset

Dim Cliente As Long
Dim MesAnt As Integer
Dim I As Integer

Dim llis As Collection
Dim TotClien(12) As Currency
Dim TotAnyo(12) As Currency
Dim Porce As Single

Dim Izquierda As String
Dim Derecha As String


    On Error GoTo ETmpVentas
    
    Set llis = New Collection
    
    'Inicializamos las listas
    For I = 1 To 12
        TotClien(I) = 0
        TotAnyo(I) = 0
    Next I
    
   
    I = InStr(cadSel, "codclien")
    If I > 0 Then 'Se ha seleccionado un cliente
        SQL = "SELECT  codclien , year(fecfactu) AnyoFac,month(fecfactu) as MesFac, sum(baseimp1+if(isnull(baseimp2),0,baseimp2)+If(isnull(baseimp3),0,baseimp3)) as BaseImp "
        SQL = SQL & " FROM scafac "
        SQL = SQL & " WHERE " & cadSel '& " AND month(fecfactu)=1 "
        SQL = SQL & " GROUP BY codclien,year(fecfactu),month(fecfactu)"
        SQL = SQL & " order by codclien,month(fecfactu) asc,year(fecfactu) asc"
    Else
        'Se seleccionara el total del anyo anterior
        SQL = "SELECT  year(fecfactu) AnyoFac,month(fecfactu) as MesFac, sum(baseimp1+if(isnull(baseimp2),0,baseimp2)+If(isnull(baseimp3),0,baseimp3)) as BaseImp "
        SQL = SQL & " FROM scafac "
        SQL = SQL & " WHERE year(fecfactu)=" & Anyo - 1
        SQL = SQL & " GROUP BY year(fecfactu),month(fecfactu)"
        SQL = SQL & " order by month(fecfactu) asc,year(fecfactu) asc"
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del cliente o del anyo anterior
    If I > 0 Then Cliente = RS!codClien
    While Not RS.EOF
        I = CInt(RS!mesfac)
        TotClien(I) = RS!BaseImp
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'Obtener el total del AÑO solicitado
    '-------------------------------------------------------------------
    SQL = "SELECT   year(fecfactu) AnyoFac,month(fecfactu) as MesFac, sum(baseimp1+if(isnull(baseimp2),0,baseimp2)+If(isnull(baseimp3),0,baseimp3)) as BaseImp "
    SQL = SQL & " FROM scafac "
    SQL = SQL & " WHERE  year(scafac.fecfactu) = " & Anyo
    SQL = SQL & " GROUP BY year(fecfactu),month(fecfactu)"
    SQL = SQL & " order by year(fecfactu) asc,month(fecfactu) asc"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del Anyo solicitado
    While Not RS.EOF
        I = CInt(RS!mesfac)
        TotAnyo(I) = RS!BaseImp
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'Insertamos en la lista todos los registros que vamos a insertar en la temporal
    'Un registro para cada mes
    For I = 1 To 12
        If TotAnyo(I) <> 0 Then
            If Cliente <> 0 Then
                'porcentaje del cliente respecto al total del año (por mes)
                Porce = Round((TotClien(I) * 100) / TotAnyo(I), 2)
            Else
                'Incremento/decremento respecto al anyo anterior (por mes)
                'en TotClien en este caso se ha almacenado el total del año anterior de cada mes
                If TotClien(I) <> 0 Then
                    Porce = Round(((TotAnyo(I) - TotClien(I)) / TotClien(I)) * 100, 2)
                Else
                    Porce = 0
                End If
            End If
        Else
            Porce = 0
        End If
        Derecha = "(" & vUsu.Codigo & "," & Cliente & "," & Anyo & "," & I & "," & DBSet(TotClien(I), "N") & "," & DBSet(Porce, "N") & "," & DBSet(TotAnyo(I), "N") & ")"
        llis.Add Derecha
    Next I
    
    
    Izquierda = "INSERT INTO tmpinformes (codusu,codigo1,campo1,campo2,importe1,porcen1,importeb1) VALUES "
    
    
    'Insertamos en la temporal todos los registros insertados en la lista
    'recorremos toda las lista
    SQL = ""
    For I = 1 To llis.Count
        SQL = SQL & llis.Item(I) & ","
        MesAnt = MesAnt + 1
    Next I
    Set llis = Nothing
    
    
    SQL = Mid(SQL, 1, Len(SQL) - 1)
    SQL = Izquierda & SQL
    conn.Execute SQL
 
    
ETmpVentas:
    If Err.Number <> 0 Then
        TempVentasMeses = False
        MuestraError Err.Number, "Ventas por meses", Err.Description
    Else
        TempVentasMeses = True
    End If
End Function




Public Sub BorrarTempInformes()
Dim SQL As String

    On Error GoTo EBorrar
    
    SQL = "DELETE FROM tmpinformes WHERE codusu=" & vUsu.Codigo
    conn.Execute SQL
    
EBorrar:
    If Err.Number <> 0 Then Err.Clear
End Sub



'
'
'  COMPRAS COMPRAS.   Va por proveedor
Public Function TempComprasMeses(cadSel As String, Anyo As String) As Boolean
'Inseta en la tabla temporal TMPINFORMES
Dim SQL As String
Dim RS As ADODB.Recordset

Dim Proveedor As Long
Dim MesAnt As Integer
Dim I As Integer

Dim llis As Collection
Dim TotClien(12) As Currency
Dim TotAnyo(12) As Currency
Dim Porce As Single

Dim Izquierda As String
Dim Derecha As String


    On Error GoTo ETmpVentas
    TempComprasMeses = False
    Set llis = New Collection
    
    'Inicializamos las listas
    For I = 1 To 12
        TotClien(I) = 0
        TotAnyo(I) = 0
    Next I
    
   
    I = InStr(cadSel, "codprove")
    If I > 0 Then 'Se ha seleccionado un cliente fecrecep
        SQL = "SELECT  codprove , year(fecrecep) AnyoFac,month(fecrecep) as MesFac, sum(baseiva1+if(isnull(baseiva2),0,baseiva2)+If(isnull(baseiva3),0,baseiva3)) as BaseImp "
        SQL = SQL & " FROM scafpc "
        SQL = SQL & " WHERE " & cadSel '& " AND month(fecfactu)=1 "
        SQL = SQL & " GROUP BY codprove,year(fecrecep),month(fecrecep)"
        SQL = SQL & " order by codprove,month(fecrecep) asc,year(fecrecep) asc"
    Else
        'Se seleccionara el total del anyo anterior
        SQL = "SELECT  year(fecrecep) AnyoFac,month(fecrecep) as MesFac, sum(baseiva1+if(isnull(baseiva2),0,baseiva2)+If(isnull(baseiva3),0,baseiva3)) as BaseImp "
        SQL = SQL & " FROM scafpc "
        SQL = SQL & " WHERE year(fecrecep)=" & Anyo - 1
        SQL = SQL & " GROUP BY year(fecrecep),month(fecrecep)"
        SQL = SQL & " order by month(fecrecep) asc,year(fecrecep) asc"
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del cliente o del anyo anterior
    If I > 0 Then Proveedor = RS!Codprove
    While Not RS.EOF
        I = CInt(RS!mesfac)
        TotClien(I) = RS!BaseImp
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'Obtener el total del AÑO solicitado
    '-------------------------------------------------------------------
    SQL = "SELECT   year(fecrecep) AnyoFac,month(fecrecep) as MesFac, sum(baseiva1+if(isnull(baseiva2),0,baseiva2)+If(isnull(baseiva3),0,baseiva3)) as BaseImp "
    SQL = SQL & " FROM scafpc "
    SQL = SQL & " WHERE  year(scafpc.fecrecep) = " & Anyo
    SQL = SQL & " GROUP BY year(fecrecep),month(fecrecep)"
    SQL = SQL & " order by year(fecrecep) asc,month(fecrecep) asc"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del Anyo solicitado
    While Not RS.EOF
        I = CInt(RS!mesfac)
        TotAnyo(I) = RS!BaseImp
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'Insertamos en la lista todos los registros que vamos a insertar en la temporal
    'Un registro para cada mes
    For I = 1 To 12
        If TotAnyo(I) <> 0 Then
            If Proveedor <> 0 Then
                'porcentaje del cliente respecto al total del año (por mes)
                Porce = Round((TotClien(I) * 100) / TotAnyo(I), 2)
            Else
                'Incremento/decremento respecto al anyo anterior (por mes)
                'en TotClien en este caso se ha almacenado el total del año anterior de cada mes
                If TotClien(I) <> 0 Then
                    Porce = Round(((TotAnyo(I) - TotClien(I)) / TotClien(I)) * 100, 2)
                Else
                    Porce = 0
                End If
            End If
        Else
            Porce = 0
        End If
        Derecha = "(" & vUsu.Codigo & "," & Proveedor & "," & Anyo & "," & I & "," & DBSet(TotClien(I), "N") & "," & DBSet(Porce, "N") & "," & DBSet(TotAnyo(I), "N") & ")"
        llis.Add Derecha
    Next I
    
    
    Izquierda = "INSERT INTO tmpinformes (codusu,codigo1,campo1,campo2,importe1,porcen1,importeb1) VALUES "
    
    
    'Insertamos en la temporal todos los registros insertados en la lista
    'recorremos toda las lista
    SQL = ""
    For I = 1 To llis.Count
        SQL = SQL & llis.Item(I) & ","
        MesAnt = MesAnt + 1
    Next I
    Set llis = Nothing
    
    
    SQL = Mid(SQL, 1, Len(SQL) - 1)
    SQL = Izquierda & SQL
    conn.Execute SQL
 
    
ETmpVentas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Ventas por meses", Err.Description
    Else
        TempComprasMeses = True
    End If
End Function





'================================================================================
'================================================================================

'================================================================================
'TMPnlotes: Temporal para introducir los Nº de lote de los Articulos en compras
'USO:
'================================================================================


Public Function DescargarDatosTMPNumLotes(NomTabla As String, cadWhere As String)
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

     '------------- AHORA
    SQL = "DELETE from " & NomTabla & " where codusu= " & vUsu.Codigo
    If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
    conn.Execute SQL
    
    Exit Function
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal (Nº Lotes).", Err.Description
End Function



'Public Function CargarDatosTMPNumLotes(NomTabla As String, cadWhere As String) As Boolean
''IN -> NomTabla: Nombre de la tabla temporal
'Dim SQL As String
'Dim i As Integer
'Dim numlinea As String, vWhere As String
'
'    On Error GoTo ECargaDatosTMP
'
'    'Insertar tantos registros como cantidad de Articulo Introducida
''    vWhere = "(codusu=" & vUsu.Codigo & " and codartic=" & DBSet(codArtic, "T") & ")"
'
'
'
''    'insertamos tantos num.serie como cantidad
''    For i = 0 To cant - 1
''        'Obtener Num Linea
''        numlinea = SugerirCodigoSiguienteStr(NomTabla, "numlinea", vWhere)
''        'Insertar en la temporal para Nº Series
''        SQL = "INSERT INTO " & NomTabla & " (codusu, codartic, numlinealb, numlinea, numserie) VALUES ("
''        SQL = SQL & vUsu.Codigo & ", " & DBSet(codArtic, "T") & ", " & NumLinAlb & ", " & numlinea & ", ' ')"
''        Conn.Execute SQL
''    Next i
'
'ECargaDatosTMP:
'    If Err.Number <> 0 Then
'        CargarDatosTMPNumLotes = False
'        MuestraError Err.Number, "Numeros Serie", Err.Description
'    Else
'        CargarDatosTMPNumSeries = True
'    End If
'End Function



Public Function PedirNLotesGnral(ByRef RS As ADODB.Recordset, Men As Boolean) As Boolean
Dim SQL As String
'Dim b As Boolean

    On Error GoTo EPedirNLotes

    If Men Then
        SQL = "Hay artículos que tienen control de Nº de Lote." & vbCrLf & vbCrLf
        SQL = SQL & "Introduzca los Nº De Lote." & vbCrLf
        MsgBox SQL, vbInformation
    End If
    
    'Cargar la tabla temporal con tantas filas como cantidad de Articulos
    'Para introducir el Nº de lote
    SQL = "numalbar=" & DBSet(RS!Numalbar, "T") & " AND fechaalb=" & DBSet(RS!FechaAlb, "F") & " AND codprove=" & DBSet(RS!Codprove, "N")
    DescargarDatosTMPNumLotes "tmpnlotes", SQL
'    b = True
    
    While Not RS.EOF
'        If Not CargarDatosTMPNumSeries("tmpnseries", RS!codArtic, RS!Cantidad, RS!numlinea) Then
'            b = False
'        End If
        SQL = "INSERT INTO tmpnlotes (codusu, numalbar, fechaalb, codprove, numlinea, codartic, codalmac, nomartic, cantidad, numlotes) VALUES ("
        SQL = SQL & vUsu.Codigo & "," & DBSet(RS!Numalbar, "T") & "," & DBSet(RS!FechaAlb, "F") & "," & RS!Codprove & "," & RS!numlinea & "," & DBSet(RS!codArtic, "T")
        SQL = SQL & "," & DBSet(RS!codAlmac, "N") & "," & DBSet(RS!NomArtic, "T") & "," & DBSet(RS!cantidad, "N") & "," & DBSet(RS!numlotes, "T", "S") & ")"
        conn.Execute SQL
        RS.MoveNext
    Wend
    PedirNLotesGnral = True
    Exit Function
    'Visualizar en pantalla el Grid, y rellenar los Nº Serie
'    If Not b Then MsgBox "No se han podido mostrar todos los Artículos con Nº de Serie.", vbInformation
    
EPedirNLotes:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        PedirNLotesGnral = False
    End If
End Function




Public Function CargarTmpInformes_Compras_312(cadTabla As String, cadSel As String) As Boolean
'Insertar en la tabla temporal tmpInformes los albaranes sin facturar
'y los albaranes ya facturados
Dim SQL As String
        
        On Error GoTo ErrTmp
        CargarTmpInformes_Compras_312 = False
        
        'codigo1= codprove, nombre3= nomprove
        'nombre1= numalbar, nombre2= numfactu
        'fecha1= fechaalb, fecha2= fecfactu
        'campo1= codforpa
        'importe1= baseimpo
        If cadTabla = "scaalp" Then 'Insertar albaranes
            SQL = "INSERT INTO tmpinformes(codusu,codigo1,nombre3,nombre1,fecha1,campo1,importe1) "
            SQL = SQL & "SELECT " & vUsu.Codigo & ", scaalp.codprove,nomprove,scaalp.numalbar,scaalp.fechaalb,codforpa,sum(importel) as baseimp"
            SQL = SQL & " FROM " & cadTabla & " inner join slialp on scaalp.numalbar=slialp.numalbar"
            SQL = SQL & " and scaalp.fechaalb=slialp.fechaalb and scaalp.codprove=slialp.codprove"
            If cadSel <> "" Then SQL = SQL & " WHERE " & cadSel
            SQL = SQL & " group by scaalp.numalbar,scaalp.fechaalb,scaalp.codprove"
            
            conn.Execute SQL
            CargarTmpInformes_Compras_312 = True
            
        Else 'insertar facturas
            SQL = "INSERT INTO tmpinformes(codusu,codigo1,nombre3,nombre1,fecha1,nombre2,fecha2,campo1,importe1) "
            SQL = SQL & "SELECT " & vUsu.Codigo & ", scafpc.codprove,nomprove,scafpa.numalbar,scafpa.fechaalb,"
            SQL = SQL & "scafpc.numfactu,scafpc.fecfactu,codforpa,sum(importel) as baseimp"
            SQL = SQL & " from (scafpc inner join scafpa on scafpc.codprove=scafpa.codprove"
            SQL = SQL & " and scafpc.numfactu=scafpa.numfactu and scafpc.fecfactu=scafpa.fecfactu)"
            SQL = SQL & " inner join slifpc on scafpa.codprove=slifpc.codprove and scafpa.numfactu=slifpc.numfactu"
            SQL = SQL & " and scafpa.fecfactu=slifpc.fecfactu and scafpa.numalbar=slifpc.numalbar"
            If cadSel <> "" Then SQL = SQL & " WHERE " & cadSel
            SQL = SQL & " group by scafpc.codprove,scafpc.numfactu,scafpc.fecfactu, scafpa.numalbar"
            conn.Execute SQL
            CargarTmpInformes_Compras_312 = True
        End If

        Exit Function
ErrTmp:
    MuestraError Err.Number, "Insertar en tmpInformes.", Err.Description
End Function










'Dado un periodo, hara un compartivo del periodo, menos un año...
Public Function TempVentasClientesPeriodoComparativo(PorAgente As Boolean, cadSel As String, cadSelPeriodo As String, cadSelAnte As String, ByRef Lbl As Label, TiposDeFacturas As String, FIni As Date, Ffin As Date, CiclosPeriodo As Integer) As Boolean
'Inserta en la temporal TMPINFORMES
Dim SQL As String, SQL2 As String
Dim SQLinsert As String
Dim RS As ADODB.Recordset
Dim Cliente As String
Dim T1 As Currency
Dim T2 As Currency
Dim T3 As Currency
Dim T4 As Currency
Dim T5 As Currency
Dim Porcen As Currency

Dim total As String 'Total del periodo seleccionado
Dim totalAnt As String 'total del periodo anterior
Dim J As Long
Dim K As Integer
Dim ColAgent As Collection
Dim vPeriodo As Byte
Dim FinBucle As Boolean
Dim NuevoCliente As Boolean

    On Error GoTo ETmpVentas
    
    '            cliente  peirodo importe
    'tmpstockfec (codusu,codartic,codalmac,stock)
    conn.Execute "DELETE FROM tmpstockfec  WHERE codusu = " & vUsu.Codigo
    
    
    'Vemos que facturas tratamos

    J = 0
    Cliente = ""
    totalAnt = "" 'para no declarar mas variables
    SQL = TiposDeFacturas
    While SQL <> ""
       J = InStr(1, SQL, "|")
       If J = 0 Then
            SQL = ""
        Else
'           If Not PorAgente Then codtipom(Len(totalAnt)) = Mid(SQL, 1, J - 1)
           Cliente = Cliente & ",'" & Mid(SQL, 1, J - 1) & "'"
           SQL = Mid(SQL, J + 1)
           totalAnt = totalAnt & "X"
        End If
    Wend
    TiposDeFacturas = "(" & Mid(Cliente, 2) & ")"
    J = Len(totalAnt)
    
    While J < 4
        codtipom(J) = "XXX"
        J = J + 1
    Wend
    
    
    
    'Obtenemos el TOTAL de ventas en ese PERIODO, de todos los clientes.
    'para obtener el % de ventas de cada cliente
    '---------------------------------------------------------------------
 
    Lbl.Caption = "Obteniendo importes :  " & vPeriodo
    Lbl.Refresh
  
     
    
    'Seleccion del PERIODO seleccionado
    'ventas por cliente y tipo de movimiento
    '----------------------------------------
    DoEvents
    Screen.MousePointer = vbHourglass
    For vPeriodo = 1 To CiclosPeriodo
        Set RS = New ADODB.Recordset
    
        Lbl.Caption = "Obteniendo importes periodo" & vPeriodo
        Lbl.Refresh
        
        SQL = "INSERT INTO tmpstockfec (codusu,codartic,codalmac,stock) SELECT " & vUsu.Codigo
        SQL = SQL & ", scafac.codclien," & vPeriodo & ",sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
        SQL = SQL & " FROM scafac,sclien "
        SQL = SQL & "  WHERE scafac.codclien=sclien.codclien "
        SQL2 = cadSel
        If vPeriodo > 1 Then
            'Es la segunda o posterior iteracion
            SQLinsert = Format(FIni, FormatoFecha)
            K = InStr(1, SQL2, SQLinsert)
            If K = 0 Then Err.Raise 513, , "Error cambiando fecha inicial peridodo "
            Cliente = Format(DateAdd("yyyy", -(vPeriodo - 1), FIni), FormatoFecha)
            SQL2 = Replace(SQL2, SQLinsert, Cliente)
            
            If FIni <> Ffin Then
                SQLinsert = Format(Ffin, FormatoFecha)
                K = InStr(1, SQL2, SQLinsert)
                If K = 0 Then Err.Raise 513, , "Error cambiando fecha final peridodo "
                Cliente = Format(DateAdd("yyyy", -(vPeriodo - 1), Ffin), FormatoFecha)
                SQL2 = Replace(SQL2, SQLinsert, Cliente)
            End If
            
            
            
            
            
        End If
        SQL = SQL & " AND " & SQL2
        SQL = SQL & "  AND scafac.codtipom IN  " & TiposDeFacturas
        SQL = SQL & " group by codclien"
        SQL = SQL & " order by codclien"
    
    
        conn.Execute SQL
    
    
    Next
    
    
    Lbl.Caption = "Abriendo temporal"
    Lbl.Refresh
    
    SQLinsert = "INSERT INTO tmpinformes (codusu, codigo1,nombre1,importe1,importe2,importe3,importe4,importe5,porcen1,importeb5,importeb1,importeb2,importeb3,importeb4) "
    SQLinsert = SQLinsert & " VALUES "
    
    SQL = ""
    SQL2 = ""
    Cliente = ""
    J = 0
    SQL = "select codartic codclien,codalmac,stock BaseImp from  tmpstockfec where codusu =" & vUsu.Codigo & " order by codartic,codalmac"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    FinBucle = RS.EOF
    Do
        
            
            If RS.EOF Then
                NuevoCliente = True
                FinBucle = True
            Else
                NuevoCliente = Cliente <> RS!codClien
                'If RS!codClien = 925 Then Stop
            End If
            If NuevoCliente Then
                Lbl.Caption = "Reg. cliente: " & Cliente
                Lbl.Refresh
                If SQL <> "" Then
                    'If Cliente = 925 Then Stop
                    If PorAgente Then T5 = T1 + T4
                    'importe1,importe2,importe3,importe4,importe5
                    SQL = SQL & DBSet(T1, "N") & "," & DBSet(T2, "N") & "," & DBSet(T3, "N") & "," & DBSet(T4, "N") & "," & DBSet(T5, "N") & ","
                    
                    If PorAgente Then
                        Porcen = T1 + T4
                    Else
                        Porcen = T1 + T2 + T3 + T4 + T5
                    End If
                    ',porcen1,importeb5,
                    SQL = SQL & "0," & DBSet(Porcen, "N") & ","
                    
                    'Los porcentajes de incremento
                    If T2 = 0 Then
                        Porcen = 0
                    Else
                        Porcen = (T1 - T2) / T2 * 100
                    End If
                    SQL = SQL & DBSet(Porcen, "N") & ","
                    If T3 = 0 Then
                        Porcen = 0
                    Else
                        Porcen = (T2 - T3) / T3 * 100
                    End If
                    SQL = SQL & DBSet(Porcen, "N") & ","
                    If T4 = 0 Then
                        Porcen = 0
                    Else
                        Porcen = (T3 - T4) / T4 * 100
                    End If
                    SQL = SQL & DBSet(Porcen, "N") & ","
                    If T5 = 0 Then
                        Porcen = 0
                    Else
                        Porcen = (T4 - T5) / T5 * 100
                    End If
                    SQL = SQL & DBSet(Porcen, "N")
                    
                    
                    SQL2 = SQL2 & SQL & ") ,"
                End If
            
                'Empezamos el registro para el siguiente cliente
                If Not RS.EOF Then
                    SQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", RS!codClien)
                    SQL = "(" & vUsu.Codigo & "," & RS!codClien & "," & DBSet(SQL, "T") & ","
                    T1 = 0
                    T2 = 0
                    T3 = 0
                    T4 = 0
                    T5 = 0
                    J = J + 1
                    Cliente = RS!codClien
                Else
                    J = 31 'para forzar el insert
                End If
            
            End If

            
            
        
        
        
        If Not RS.EOF Then
            'AHORA
            If PorAgente Then
                T1 = T1 + RS!BaseImp
            Else
                If RS!codAlmac = 1 Then
                    T1 = T1 + RS!BaseImp
                ElseIf RS!codAlmac = 2 Then
                    T2 = T2 + RS!BaseImp
                ElseIf RS!codAlmac = 3 Then
                    T3 = T3 + RS!BaseImp
                ElseIf RS!codAlmac = 4 Then
                    T4 = T4 + RS!BaseImp
                Else
                    T5 = T5 + RS!BaseImp
                End If
            End If
            
       
        
            RS.MoveNext
       End If
        
        
        
        
        If J >= 30 Then
            'Insertamos en la tabla temporal
            If SQL2 <> "" Then
                SQL2 = Mid(SQL2, 1, Len(SQL2) - 1)
                SQL2 = SQLinsert & SQL2
                conn.Execute SQL2
            End If
            
            'Reiniciamos los valores
            SQL2 = ""
            J = 0
        End If
        
        
        
        
    Loop Until FinBucle
    
    RS.Close
    
    
'
'    'Para no tocar la funcioin de arriba, ahora recorro tmoinformes en codigo y le pongo en campo1 el agente
'    Lbl.Caption = "Updatear tmp"
'    Lbl.Refresh
'    SQL = "Select codagent from tmpinformes,sclien where tmpinformes.codigo1 = sclien.codclien and  codusu = " & vUsu.Codigo & " GROUP BY 1"
'    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Set ColAgent = New Collection
'    While Not RS.EOF
'        SQL = RS.Fields(0)
'        ColAgent.Add SQL
'        RS.MoveNext
'    Wend
'    RS.Close
'
'    For J = 1 To ColAgent.Count
'        Lbl.Caption = J & " de " & ColAgent.Count
'        Lbl.Refresh
'        SQL = "UPDATE tmpinformes,sclien SET tmpinformes.campo1=" & ColAgent.Item(J) & " where tmpinformes.codigo1 = sclien.codclien and  codusu = " & vUsu.Codigo & " AND codagent=" & ColAgent.Item(J)
'        conn.Execute SQL
'    Next J
'
  
  
  
  
  
  
  
  
    'CLIENTES VARIOS
    Lbl.Caption = "Updatear tmp clivar"
    Lbl.Refresh


    SQL = "Select codigo1,nombre1,nomclien from tmpinformes,sclien where tmpinformes.codigo1 = sclien.codclien"
    SQL = SQL & " and clivario=1"
    SQL = SQL & " and  codusu = " & vUsu.Codigo & " GROUP BY 1"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColAgent = New Collection
    While Not RS.EOF
        SQL = "nombre1= " & DBSet(RS!NomClien, "T") & " WHERE codusu = " & vUsu.Codigo & " AND codigo1=" & RS!Codigo1
        ColAgent.Add SQL
        RS.MoveNext
    Wend
    RS.Close
    
    For J = 1 To ColAgent.Count
        Lbl.Caption = J & " de " & ColAgent.Count
        Lbl.Refresh
        SQL = "UPDATE tmpinformes SET " & ColAgent.Item(J)
        conn.Execute SQL
    Next J
  
  
  
  
      Lbl.Caption = "Totales"
    Lbl.Refresh
  
    SQL = "SELECT sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
    SQL = SQL & " FROM scafac,sclien "
    SQL = SQL & "  WHERE scafac.codclien=sclien.codclien AND " & cadSelPeriodo
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "0"
    If Not RS.EOF Then
        SQL = DBLet(RS!BaseImp, "N")
    End If
    RS.Close
    cadSelPeriodo = Replace(SQL, ",", ".")
    
    SQL = "UPDATE tmpinformes SET porcen1=round((importe1 * 100)/" & cadSelPeriodo & ",3) WHERE codusu =" & vUsu.Codigo
    conn.Execute SQL
    
    
    SQL = "Select sum(importeb5) BaseImp FROM tmpinformes WHERE  codusu =" & vUsu.Codigo
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "0"
    If Not RS.EOF Then
        If DBLet(RS!BaseImp, "N") <> 0 Then
            SQL = RS!BaseImp
            SQL = Replace(SQL, ",", ".")
            SQL = "UPDATE tmpinformes SET porcen2=round((importeb5 * 100)/" & SQL & ",3) WHERE codusu =" & vUsu.Codigo
            conn.Execute SQL
        End If
    End If
    RS.Close
   
    
    
    
    
ETmpVentas:
    Set RS = Nothing
    If Err.Number <> 0 Then
        TempVentasClientesPeriodoComparativo = False
        MuestraError Err.Number, "Ventas del periodo" & Lbl.Caption, Err.Description
    Else
        TempVentasClientesPeriodoComparativo = True
    End If
    Lbl.Caption = ""
    Set ColAgent = Nothing
    Screen.MousePointer = vbDefault
End Function

