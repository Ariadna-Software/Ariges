Attribute VB_Name = "ModTMP"
Option Explicit

'MODULO PARA LA CARGA Y DESCARGA DE TABLAS TEMPORALES

'LO tengo que crear "global"
Dim codtipom(3) As String


'================================================================================
'================================================================================

'================================================================================
'TMPnseries: Temporal para introducir los N� de Serie de los Articulos en compras o en ventas
'USO: frmFacEntAlbaran, frmRepEntAlbaran
'================================================================================

Public Function CargarDatosTMPNumSeries(NomTabla As String, codArtic As String, Cant As Integer, NumLinAlb As String) As Boolean
'IN -> NomTabla: Nombre de la tabla temporal
'      CodArtic: Codigo Articulo del que se van a Introducir los N� de Serie
'      Cant: cantidad de Articulo (tantas filas como articulos)
'      Mostrar: si true se cargar los N� de serie sino en blanco
Dim SQL As String
Dim i As Integer
Dim numlinea As String, vWhere As String

    On Error GoTo ECargaDatosTMP

    'Insertar tantos registros como cantidad de Articulo Introducida
    vWhere = "(codusu=" & vUsu.codigo & " and codartic=" & DBSet(codArtic, "T") & " and numlinealb=" & DBSet(NumLinAlb, "N") & ")"

    'insertamos tantos num.serie como cantidad
    For i = 0 To Cant - 1
        'Obtener Num Linea
        numlinea = SugerirCodigoSiguienteStr(NomTabla, "numlinea", vWhere)
        'Insertar en la temporal para N� Series
        SQL = "INSERT INTO " & NomTabla & " (codusu, codartic, numlinealb, numlinea, numserie,nummante) VALUES ("
        SQL = SQL & vUsu.codigo & ", " & DBSet(codArtic, "T") & ", " & NumLinAlb & ", " & numlinea & ", ' ',' ')"
        conn.Execute SQL
    Next i
 
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
    SQL = "DELETE from " & NomTabla & " where codusu= " & vUsu.codigo
    conn.Execute SQL
    
    Exit Function
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal (N� Serie).", Err.Description
End Function



Public Function InsertarNSeries(codArtic As String, CadValuesI As String, cadValuesU As String, DeVenta As Boolean) As Boolean
'Insertar un registro en la tabla "sserie" por cada uno de los
'N� de Serie introducidos en la Tabla Temporal
Dim Rs As ADODB.Recordset
Dim SQL As String, devuelve As String
Dim codTipar As String, NumAlbar As String
    
    On Error GoTo EInsertar

    'Seleccionar los n� de serie cargados en la temporal: tmpnseries
    SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.codigo & " AND codartic=" & DBSet(codArtic, "T")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then Rs.MoveFirst
    
    While Not Rs.EOF
        'Comprobar si existe en la tabla sserie
        If DeVenta Then
            NumAlbar = "numalbar" 'N� albaran de Venta
        Else
            NumAlbar = "numalbpr" 'N� albaran de Compras
        End If
        devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", Rs!numSerie, "T", NumAlbar, "codartic", Rs!codArtic, "T")
        If devuelve <> "" Then 'Existe en tabla sserie
            If NumAlbar = "" Then
                SQL = Trim(DBLet(Rs!nummante, "T"))
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
                SQL = SQL & " WHERE numserie=" & DBSet(Rs!numSerie, "T") & " AND codartic=" & DBSet(Rs!codArtic, "T")
                '===
            End If
        Else
            'Obtener el tipo de Articulo
            codTipar = DevuelveDesdeBDNew(conAri, "sartic", "codtipar", "codartic", Rs!codArtic, "T")
        
            'Insertar en la tabla sserie
            SQL = "INSERT INTO sserie (numserie, codartic, codtipar, codclien, coddirec,tieneman, nummante, ultrepar, fingaran, "
            SQL = SQL & " codtipom, numfactu, fechavta, numalbar, numline1, codprove, numalbpr, fechacom, numline2) "
            SQL = SQL & " VALUES ( " & DBSet(Rs!numSerie, "T") & ", " & DBSet(Rs!codArtic, "T") & ", " & DBSet(codTipar, "T") & ","
            SQL = SQL & CadValuesI
            SQL = SQL & ") "
        End If
        conn.Execute SQL
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
EInsertar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando N� Serie", Err.Description
End Function





Public Sub PedirNSeriesGnral(ByRef Rs As ADODB.Recordset, Men As Boolean)
Dim SQL As String
Dim b As Boolean

    On Error GoTo EPedirNSeries

        If Men Then
            SQL = "Hay art�culos que tienen control de N� de Serie." & vbCrLf & vbCrLf
            SQL = SQL & "Introduzca los N� De Serie." & vbCrLf
            MsgBox SQL, vbInformation
        End If
        
        'Cargar la tabla temporal con tantas filas como cantidad de Articulo
        'Para introducir el N� de Serie
        DescargarDatosTMPNumSeries ("tmpnseries")
        b = True
        
        While Not Rs.EOF
            If Not CargarDatosTMPNumSeries("tmpnseries", Rs!codArtic, Rs!cantidad, Rs!numlinea) Then
                b = False
            End If
            Rs.MoveNext
        Wend
        
        'Visualizar en pantalla el Grid, y rellenar los N� Serie
        If Not b Then MsgBox "No se han podido mostrar todos los Art�culos con N� de Serie.", vbInformation
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Public Function MostrarNSeriesGnral(ByRef RSLineas As ADODB.Recordset, vCampos As String, Optional Rectifica As Boolean) As String
'Si los N� de serie se introdujeron en ALBARAN COMPRAS se muestran
'los N� de serie de los articulos comprados y se seleccionan tantos como cantidad de la linea
'IN -> RSLineas: lineas del Albaran generado
'OUT -> vCampos: concatena la cantidad requerida de N� series de cada articulo
'RETURN -> cadena SQL con la Select que se pasara para mostrar los N� series
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
        Campos = Campos & RSLineas!codArtic & "|" & RSLineas!cantidad & "�"
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
   
    'Se puede seleccionar todos los N� de Serie que se necesitan
     'se introdujo los N� de Serie en COMPRAS y ahora
    'mostramos los N� de Serie para seleccionar cual vamos a
    'vender al Cliente
    Screen.MousePointer = vbDefault
    If Rectifica Then
        'viene de una factura rectificativa los n� de serie que seleccionemos sera para quitar
        SQL = "Hay Art�culos que tienen control de N� de Serie." & vbCrLf
        SQL = SQL & "Seleccione los n� de Serie que desea rectificar."
    Else
        If totArtic > 1 Then
            SQL = "Hay Art�culos que tienen control de N� de Serie." & vbCrLf
            SQL = SQL & vbCrLf & "Seleccione un N� de Serie para cada Art�culo."
        Else
            SQL = "El Art�culo tienen control de N� de Serie." & vbCrLf
            SQL = SQL & vbCrLf & "Seleccione los N� de Serie para el Art�culo."
        End If
    End If
    MsgBox SQL, vbInformation
    SQL = " WHERE sserie.codartic IN " & cadArtic
    MostrarNSeriesGnral = SQL
    
EMostrar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Mostrar N� series", Err.Description
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
Dim Rs As ADODB.Recordset
Dim vStock As Single
Dim cadSQL As String
Dim Insertar As Boolean
Dim Entradas As Currency
Dim Salidas As Currency

    On Error GoTo ECargarTMPStock

    CargarTMPStockFecha = False
    
    Lbl.Caption = "Leyendo registros"
    Lbl.Refresh
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Lbl.Caption = "Art  " & Rs!codArtic
        Lbl.Refresh
        'Para cada articulo obtener el stock en esa fecha e insertarlo en la temporal
        'Antes abril 2011
        'cadSQL = "SELECT sum(cantidad) FROM smoval WHERE "
        cadSQL = "SELECT tipomovi,sum(cantidad) FROM smoval WHERE "
   
        cadSQL = cadSQL & " codartic=" & DBSet(Rs!codArtic, "T") & " AND codalmac=" & Rs!codAlmac
        
            vStock = Rs!CanStock
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
            cadSQL = cadSQL & " VALUES (" & vUsu.codigo & ", " & DBSet(Rs!codArtic, "T") & ", "
            cadSQL = cadSQL & Rs!codAlmac & ", " & TransformaComasPuntos(CStr(vStock)) & ")"
            conn.Execute cadSQL
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
    cadSQL = "Select count(*) from tmpstockfec where codusu =" & vUsu.codigo
    Rs.Open cadSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0), "N") > 0 Then cadSQL = ""
    End If
    Rs.Close
    
    Set Rs = Nothing
    
    If cadSQL <> "" Then
        MsgBox "No hay datos para mostrar con estos parametros", vbExclamation
    Else
        CargarTMPStockFecha = True
    End If
    
ECargarTMPStock:
    If Err.Number <> 0 Then
'        RS.Close
        Set Rs = Nothing
        MsgBox " No se ha podido cargar la Tabla Temporal correctamente", vbInformation
    End If
    Lbl.Caption = ""
End Function



Public Function DescargarDatosTMPStockFecha()
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

    '------------- AHORA
    SQL = "DELETE from tmpstockfec" & " where codusu= " & vUsu.codigo
    conn.Execute SQL
    Exit Function
    
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal (Stock a Fecha).", Err.Description
End Function



Private Function TotMovimientosStock2(cadSQL As String, vTipomovi As Byte) As Single
'Para un tipo de Movimiento vtipomovi(0=Salida, 1=Entrada) devolver
'la cantidad de stock para esos registros de la select
Dim RSmov As ADODB.Recordset
Dim cad As String

        TotMovimientosStock2 = 0
        cad = cadSQL & " AND tipomovi=" & vTipomovi
        
        Set RSmov = New ADODB.Recordset
        RSmov.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
Dim cad As String

    
        Entradas = 0
        Salidas = 0
        cad = cadSQL & " GROUP BY tipomovi"
        
        Set RSmov = New ADODB.Recordset
        RSmov.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
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
'A�adimos el poder mostrar que tipos de facturas ENTRAN
Public Function TempVentasClientes(PorAgente As Boolean, cadSel As String, cadSelPeriodo As String, cadSelAnte As String, ByRef Lbl As Label, TiposDeFacturas As String) As Boolean
'Inserta en la temporal TMPINFORMES
Dim SQL As String, SQL2 As String
Dim SQLinsert As String
Dim Rs As ADODB.Recordset
Dim Cliente As String
Dim T1 As Currency
Dim T2 As Currency
Dim T3 As Currency
Dim T4 As Currency
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
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        total = CStr(DBLet(Rs.Fields(0), "N"))
    End If
    Rs.Close
    Set Rs = Nothing
     
    
    
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
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            totalAnt = CStr(DBLet(Rs.Fields(0), "N"))
        End If
        Rs.Close
        Set Rs = Nothing
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQLinsert = "INSERT INTO tmpinformes (codusu, codigo1,nombre1,importe1,importe2,importe3,importe4,importe5,porcen1,importeb1,importeb2,importeb3,importeb4,importeb5) "
    SQLinsert = SQLinsert & " VALUES "
    
    SQL = ""
    SQL2 = ""
    J = 0
    While Not Rs.EOF
        If Cliente <> Rs!codClien Then
            Lbl.Caption = "Reg. cliente: " & Rs!codClien
            Lbl.Refresh
            If SQL <> "" Then
                SQL = SQL & DBSet(T1, "N") & "," & DBSet(T2, "N") & "," & DBSet(T3, "N") & "," & DBSet(T4, "N") & ","
                '---- Laura: modificado 26/09/2006
'                totVentas = CStr(CCur(ComprobarCero(totVentas)) + CCur(ComprobarCero(totMante)) + CCur(ComprobarCero(totRepar)) + CCur(ComprobarCero(totRectif)))
                'totVentas = totVentas + totMante + totRepar + totRectif + totServi
                T1 = T1 + T2 + T3 + T4
                '----
                SQL = SQL & DBSet(T1, "N") & ","
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
            SQL = "(" & vUsu.codigo & "," & Rs!codClien & "," & DBSet(Rs!NomClien, "T") & ","
            T1 = 0
            T2 = 0
            T3 = 0
            T4 = 0
   
            J = J + 1
        End If
    
        'AHORA
        If PorAgente Then
            
            If Rs!codtipom = "FRT" Then
                T4 = T4 + Rs!BaseImp
            Else
                T1 = T1 + Rs!BaseImp
            End If
        Else
            If Rs!codtipom = codtipom(0) Then
                T1 = T1 + Rs!BaseImp
            ElseIf Rs!codtipom = codtipom(1) Then
                T2 = T2 + Rs!BaseImp
            ElseIf Rs!codtipom = codtipom(2) Then
                T3 = T3 + Rs!BaseImp
            ElseIf Rs!codtipom = codtipom(3) Then
                T4 = T4 + Rs!BaseImp
            End If
            
        End If
        Cliente = Rs!codClien
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    
    If SQL <> "" Then 'para el ultimo registro
        SQL = SQL & DBSet(T1, "N") & "," & DBSet(T2, "N") & "," & DBSet(T3, "N") & "," & DBSet(T4, "N") & ","
        '---- Laura: Modificado 26/09/2006
        'totVentas = CStr(CCur(ComprobarCero(totVentas)) + CCur(ComprobarCero(totMante)) + CCur(ComprobarCero(totRepar)) + CCur(ComprobarCero(totRectif)))
        T1 = CStr(T1 + T2 + T3 + T4)
        '----
        SQL = SQL & DBSet(T1, "N") & ","
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
    SQL = "Select codagent from tmpinformes,sclien where tmpinformes.codigo1 = sclien.codclien and  codusu = " & vUsu.codigo & " GROUP BY 1"
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColAgent = New Collection
    While Not Rs.EOF
        SQL = Rs.Fields(0)
        ColAgent.Add SQL
        Rs.MoveNext
    Wend
    Rs.Close
    
    For J = 1 To ColAgent.Count
        Lbl.Caption = J & " de " & ColAgent.Count
        Lbl.Refresh
        SQL = "UPDATE tmpinformes,sclien SET tmpinformes.campo1=" & ColAgent.item(J) & " where tmpinformes.codigo1 = sclien.codclien and  codusu = " & vUsu.codigo & " AND codagent=" & ColAgent.item(J)
        conn.Execute SQL
    Next J
  
  
  
  
  
  
  
  
  
    'CLIENTES VARIOS
    Lbl.Caption = "Updatear tmp clivar"
    Lbl.Refresh


    SQL = "Select codigo1,nombre1,nomclien from tmpinformes,sclien where tmpinformes.codigo1 = sclien.codclien"
    SQL = SQL & " and clivario=1"
    SQL = SQL & " and  codusu = " & vUsu.codigo & " GROUP BY 1"
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColAgent = New Collection
    While Not Rs.EOF
        SQL = "nombre1= " & DBSet(Rs!NomClien, "T") & " WHERE codusu = " & vUsu.codigo & " AND codigo1=" & Rs!Codigo1
        ColAgent.Add SQL
        Rs.MoveNext
    Wend
    Rs.Close
    
    For J = 1 To ColAgent.Count
        Lbl.Caption = J & " de " & ColAgent.Count
        Lbl.Refresh
        SQL = "UPDATE tmpinformes SET " & ColAgent.item(J)
        conn.Execute SQL
    Next J
  
  
    
    
    
    
    
ETmpVentas:
    Set Rs = Nothing
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
Dim Rs As ADODB.Recordset
Dim T1 As Currency
Dim T2 As Currency
Dim T3 As Currency
Dim T4 As Currency


    On Error GoTo EVentas
    
    T1 = 0
    T2 = 0
    T3 = 0
    T4 = 0
    
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
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
             'Select Case RS!Codtipom
             '   Case "FAV", "FTI", "FAS", "FMO", "FAI": totVentas = totVentas + RS!BaseImp
             '   Case "FAM": totMante = RS!BaseImp
             '   Case "FAR": totRepar = RS!BaseImp
             '   Case "FRT": totRectif = RS!BaseImp
             '   Case Else
             '
             '     '  If RS!codtipom <> "FAZ" Then Stop
             'End Select
            If PorAgente Then
                
                If Rs!codtipom = "FRT" Then
                    T4 = T4 + Rs!BaseImp
                Else
                    T1 = T1 + Rs!BaseImp
                End If
            Else
                If Rs!codtipom = codtipom(0) Then
                    T1 = T1 + Rs!BaseImp
                ElseIf Rs!codtipom = codtipom(1) Then
                    T2 = T2 + Rs!BaseImp
                ElseIf Rs!codtipom = codtipom(2) Then
                    T3 = T3 + Rs!BaseImp
                ElseIf Rs!codtipom = codtipom(3) Then
                    T4 = T4 + Rs!BaseImp
                End If
            End If
                         
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    End If
    
    SQL = DBSet(T1, "N") & "," & DBSet(T2, "N") & "," & DBSet(T3, "N") & "," & DBSet(T4, "N") & ","
    '---- Laura: Modificado 26/09/2006
    'totVentas = CStr(CCur(ComprobarCero(totVentas)) + CCur(ComprobarCero(totMante)) + CCur(ComprobarCero(totRepar)) + CCur(ComprobarCero(totRectif)))
    T1 = T1 + T2 + T3 + T4
    '----
    SQL = SQL & DBSet(T1, "N")
    VentasPeriodoAnterior = SQL
    
EVentas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Ventas periodo anterior", Err.Description
    End If
End Function





Public Function TempVentasMeses(cadSel As String, Anyo As String) As Boolean
'Inseta en la tabla temporal TMPINFORMES
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim Cliente As Long
Dim MesAnt As Integer
Dim i As Integer

Dim llis As Collection
Dim TotClien(12) As Currency
Dim TotAnyo(12) As Currency
Dim Porce As Single

Dim Izquierda As String
Dim Derecha As String


    On Error GoTo ETmpVentas
    
    Set llis = New Collection
    
    'Inicializamos las listas
    For i = 1 To 12
        TotClien(i) = 0
        TotAnyo(i) = 0
    Next i
    
   
    i = InStr(cadSel, "codclien")
    If i > 0 Then 'Se ha seleccionado un cliente
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del cliente o del anyo anterior
    If i > 0 Then Cliente = Rs!codClien
    While Not Rs.EOF
        i = CInt(Rs!mesfac)
        TotClien(i) = Rs!BaseImp
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'Obtener el total del A�O solicitado
    '-------------------------------------------------------------------
    SQL = "SELECT   year(fecfactu) AnyoFac,month(fecfactu) as MesFac, sum(baseimp1+if(isnull(baseimp2),0,baseimp2)+If(isnull(baseimp3),0,baseimp3)) as BaseImp "
    SQL = SQL & " FROM scafac "
    SQL = SQL & " WHERE  year(scafac.fecfactu) = " & Anyo
    SQL = SQL & " GROUP BY year(fecfactu),month(fecfactu)"
    SQL = SQL & " order by year(fecfactu) asc,month(fecfactu) asc"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del Anyo solicitado
    While Not Rs.EOF
        i = CInt(Rs!mesfac)
        TotAnyo(i) = Rs!BaseImp
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertamos en la lista todos los registros que vamos a insertar en la temporal
    'Un registro para cada mes
    For i = 1 To 12
        If TotAnyo(i) <> 0 Then
            If Cliente <> 0 Then
                'porcentaje del cliente respecto al total del a�o (por mes)
                Porce = Round((TotClien(i) * 100) / TotAnyo(i), 2)
            Else
                'Incremento/decremento respecto al anyo anterior (por mes)
                'en TotClien en este caso se ha almacenado el total del a�o anterior de cada mes
                If TotClien(i) <> 0 Then
                    Porce = Round(((TotAnyo(i) - TotClien(i)) / TotClien(i)) * 100, 2)
                Else
                    Porce = 0
                End If
            End If
        Else
            Porce = 0
        End If
        Derecha = "(" & vUsu.codigo & "," & Cliente & "," & Anyo & "," & i & "," & DBSet(TotClien(i), "N") & "," & DBSet(Porce, "N") & "," & DBSet(TotAnyo(i), "N") & ")"
        llis.Add Derecha
    Next i
    
    
    Izquierda = "INSERT INTO tmpinformes (codusu,codigo1,campo1,campo2,importe1,porcen1,importeb1) VALUES "
    
    
    'Insertamos en la temporal todos los registros insertados en la lista
    'recorremos toda las lista
    SQL = ""
    For i = 1 To llis.Count
        SQL = SQL & llis.item(i) & ","
        MesAnt = MesAnt + 1
    Next i
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
    
    SQL = "DELETE FROM tmpinformes WHERE codusu=" & vUsu.codigo
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
Dim Rs As ADODB.Recordset

Dim Proveedor As Long
Dim MesAnt As Integer
Dim i As Integer

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
    For i = 1 To 12
        TotClien(i) = 0
        TotAnyo(i) = 0
    Next i
    
   
    i = InStr(cadSel, "codprove")
    If i > 0 Then 'Se ha seleccionado un cliente fecrecep
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del cliente o del anyo anterior
    If i > 0 Then Proveedor = Rs!Codprove
    While Not Rs.EOF
        i = CInt(Rs!mesfac)
        TotClien(i) = Rs!BaseImp
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'Obtener el total del A�O solicitado
    '-------------------------------------------------------------------
    SQL = "SELECT   year(fecrecep) AnyoFac,month(fecrecep) as MesFac, sum(baseiva1+if(isnull(baseiva2),0,baseiva2)+If(isnull(baseiva3),0,baseiva3)) as BaseImp "
    SQL = SQL & " FROM scafpc "
    SQL = SQL & " WHERE  year(scafpc.fecrecep) = " & Anyo
    SQL = SQL & " GROUP BY year(fecrecep),month(fecrecep)"
    SQL = SQL & " order by year(fecrecep) asc,month(fecrecep) asc"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del Anyo solicitado
    While Not Rs.EOF
        i = CInt(Rs!mesfac)
        TotAnyo(i) = Rs!BaseImp
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'Insertamos en la lista todos los registros que vamos a insertar en la temporal
    'Un registro para cada mes
    For i = 1 To 12
        If TotAnyo(i) <> 0 Then
            If Proveedor <> 0 Then
                'porcentaje del cliente respecto al total del a�o (por mes)
                Porce = Round((TotClien(i) * 100) / TotAnyo(i), 2)
            Else
                'Incremento/decremento respecto al anyo anterior (por mes)
                'en TotClien en este caso se ha almacenado el total del a�o anterior de cada mes
                If TotClien(i) <> 0 Then
                    Porce = Round(((TotAnyo(i) - TotClien(i)) / TotClien(i)) * 100, 2)
                Else
                    Porce = 0
                End If
            End If
        Else
            Porce = 0
        End If
        Derecha = "(" & vUsu.codigo & "," & Proveedor & "," & Anyo & "," & i & "," & DBSet(TotClien(i), "N") & "," & DBSet(Porce, "N") & "," & DBSet(TotAnyo(i), "N") & ")"
        llis.Add Derecha
    Next i
    
    
    Izquierda = "INSERT INTO tmpinformes (codusu,codigo1,campo1,campo2,importe1,porcen1,importeb1) VALUES "
    
    
    'Insertamos en la temporal todos los registros insertados en la lista
    'recorremos toda las lista
    SQL = ""
    For i = 1 To llis.Count
        SQL = SQL & llis.item(i) & ","
        MesAnt = MesAnt + 1
    Next i
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
'TMPnlotes: Temporal para introducir los N� de lote de los Articulos en compras
'USO:
'================================================================================


Public Function DescargarDatosTMPNumLotes(NomTabla As String, cadWhere As String)
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

     '------------- AHORA
    SQL = "DELETE from " & NomTabla & " where codusu= " & vUsu.codigo
    If cadWhere <> "" Then SQL = SQL & " AND " & cadWhere
    conn.Execute SQL
    
    Exit Function
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal (N� Lotes).", Err.Description
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
''        'Insertar en la temporal para N� Series
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



Public Function PedirNLotesGnral(ByRef Rs As ADODB.Recordset, Men As Boolean) As Boolean
Dim SQL As String
'Dim b As Boolean

    On Error GoTo EPedirNLotes

    If Men Then
        SQL = "Hay art�culos que tienen control de N� de Lote." & vbCrLf & vbCrLf
        SQL = SQL & "Introduzca los N� De Lote." & vbCrLf
        MsgBox SQL, vbInformation
    End If
    
    'Cargar la tabla temporal con tantas filas como cantidad de Articulos
    'Para introducir el N� de lote
    SQL = "numalbar=" & DBSet(Rs!NumAlbar, "T") & " AND fechaalb=" & DBSet(Rs!FechaAlb, "F") & " AND codprove=" & DBSet(Rs!Codprove, "N")
    DescargarDatosTMPNumLotes "tmpnlotes", SQL
'    b = True
    
    While Not Rs.EOF
'        If Not CargarDatosTMPNumSeries("tmpnseries", RS!codArtic, RS!Cantidad, RS!numlinea) Then
'            b = False
'        End If
        SQL = "INSERT INTO tmpnlotes (codusu, numalbar, fechaalb, codprove, numlinea, codartic, codalmac, nomartic, cantidad, numlotes) VALUES ("
        SQL = SQL & vUsu.codigo & "," & DBSet(Rs!NumAlbar, "T") & "," & DBSet(Rs!FechaAlb, "F") & "," & Rs!Codprove & "," & Rs!numlinea & "," & DBSet(Rs!codArtic, "T")
        SQL = SQL & "," & DBSet(Rs!codAlmac, "N") & "," & DBSet(Rs!NomArtic, "T") & "," & DBSet(Rs!cantidad, "N") & "," & DBSet(Rs!numlotes, "T", "S") & ")"
        conn.Execute SQL
        Rs.MoveNext
    Wend
    PedirNLotesGnral = True
    Exit Function
    'Visualizar en pantalla el Grid, y rellenar los N� Serie
'    If Not b Then MsgBox "No se han podido mostrar todos los Art�culos con N� de Serie.", vbInformation
    
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
            SQL = SQL & "SELECT " & vUsu.codigo & ", scaalp.codprove,nomprove,scaalp.numalbar,scaalp.fechaalb,codforpa,sum(importel) as baseimp"
            SQL = SQL & " FROM " & cadTabla & " inner join slialp on scaalp.numalbar=slialp.numalbar"
            SQL = SQL & " and scaalp.fechaalb=slialp.fechaalb and scaalp.codprove=slialp.codprove"
            If cadSel <> "" Then SQL = SQL & " WHERE " & cadSel
            SQL = SQL & " group by scaalp.numalbar,scaalp.fechaalb,scaalp.codprove"
            
            conn.Execute SQL
            CargarTmpInformes_Compras_312 = True
            
        Else 'insertar facturas
            SQL = "INSERT INTO tmpinformes(codusu,codigo1,nombre3,nombre1,fecha1,nombre2,fecha2,campo1,importe1) "
            SQL = SQL & "SELECT " & vUsu.codigo & ", scafpc.codprove,nomprove,scafpa.numalbar,scafpa.fechaalb,"
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
