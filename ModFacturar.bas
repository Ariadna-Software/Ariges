Attribute VB_Name = "ModFacturar"
Option Explicit

'===================================================================================
'Modulo para el traspaso de registros de cabecera y lineas de las tablas de ALBARAN
'A las tablas del FACTURACION
' o para pasar de las tablas de Mantenimientos a tablas de FACTURACION
'====================================================================================

'operador del albaran para facturas de Mantenimientos
Private OpeFactu As String
Private MesFactu As String 'mes a facturar para Mantenimientos
Private TipCoMan As String 'tipo de contrato del mantenimiento

'Variables comunes en Albaranes para la cabecera de la FACTURA
Private LetraSer As String

Private TipoAlb As String
Private TipoFac As String

'Variable con la WHERE que selecciona todos los Albaranes que forma parte de la Factura
Private cadW As String


Dim Errores As String
Dim ErroresAux As String


Public Function TraspasoAlbaranesFacturas(cadSQL As String, cadWhere As String, FechaFact As String, banPr As String, ByRef PBar1 As ProgressBar, ByRef LblBar As Label, ImprimeLasFacturasGeneradas As Boolean, ByRef vTipoM As String, TextosCSB As String, NumeroCopias As Byte, MostrarMsgOK As Boolean, EsTraspasoOfeFAZ As Boolean) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      BanPr: Cod. de Banco Propio
'      Pbar1:  Una progressbar. Se puede mandar un NOTHING, y no pasa nada. Si no se manda
'              es que estamos en un proceso corto o que no necesitabaos un pb1, con lo cual NO muestro el PB1
'      Imprime: Si despues de generarlo los imprime
'
'       vTipom:  Que tipo de albaran es, para luego la impresion saber que factura imprime
'      TextosCSB:  Si lleva llevara 3 lineas para meter ent tesoreria

'   Abril
'   EsTraspasoOfeFAZ    Traspasa directamente una OFE a un FAZ


'Desde Albaranes Genera las Facturas correspondientes
Dim RSalb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim SQL As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antClien As Long
Dim antDirec As Long
Dim antForpa As Integer
Dim antDtoPP As Single, antDtoGn As Single

'direc/dpto actual para controlar el valor nulo
Dim actDirec As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vFactu As CFactura
Dim Inc As Integer
Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

'Por si no mando una progressbar, que no de errores
Dim PgbVisible As Boolean


Dim HazPulsarAceptarEnFrmImprimir As Boolean

    On Error GoTo ETraspasoAlbFac

    TraspasoAlbaranesFacturas = False

    ListFactu = ""
        
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("VENFAC") 'facturas de venta
    If Not BloqueoManual("VENFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los albaranes que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    SQL = " (scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien ) INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
    If Not BloqueaRegistro(SQL, cadWhere) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        Exit Function
    End If
    
   
    'Inicializar la Progress Bar
    PgbVisible = False
    If Not (PBar1 Is Nothing) Then
        If PBar1.visible Then PgbVisible = True
    End If
    If PgbVisible Then
        If InStr(1, cadSQL, "sclien") Then
            SQL = Replace(cadSQL, "scaalb.*, sclien.periodof", "count(*)") 'si hay INNER JOIN con sclien
        Else
            SQL = Replace(cadSQL, "*", "count(*)") 'si NO hay INNER JOIN con sclien
        End If
        
        
        Set RSalb = New ADODB.Recordset
        RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RSalb.EOF Then
            CargarProgresNew PBar1, CInt(RSalb.Fields(0))
            LblBar.Caption = "Inicializando el proceso..."
            LblBar.Refresh
            
        End If
        RSalb.Close
        Set RSalb = Nothing
    End If
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactu = New CFactura
    vFactu.FecFactu = FechaFact 'Fecha para las Facturas

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", banPr, "N")
    
    'comprobar que la cuenta prevista de cobro tiene valor
    b = (vFactu.CuentaPrev <> "")
    If Not b Then
        Set vFactu = Nothing
        'Desbloqueamos ya no estamos facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
        MsgBox "La cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Function
    End If
    
       
        
    'Marcar Albaranes que se van a Facturar
    '----------------------------------------
    SQL = cadSQL & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec, codforpa, dtoppago, dtognral "
    Set RSalb = New ADODB.Recordset
    RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    'Agrupar los Albaranes posibles en una misma Factura
    'Calcular y Grabar Factura en la Tabla de Facturas
    'Albaran(scaalb, slialb) -> Factura (scafac,scafac1,slifac)
    '----------------------------------------------------
    'Se factura por cliente y departamento
    'Agrupar albaranes en 1 factura por : tipofact,codclien,coddirec,codforpa,dtoppago, dtognral
    antClien = 0 'cliente
    antDirec = 0 'direccion/departamento
    antForpa = 0 'forma de pago
    antDtoPP = 0 'dto pronto pago
    antDtoGn = 0 'dto general
    
    cadW = ""
    Errores = ""
    Inc = 0
    
    While Not RSalb.EOF
        TipoAlb = RSalb!codtipom
        Inc = Inc + 1
        If IsNull(RSalb!CodDirec) Then
            actDirec = -1
        Else
            actDirec = DBLet(RSalb!CodDirec, "N")
        End If
        
        If RSalb!TipoFact = 1 Then 'tipofact=1 "FACTURA x ALBARAN"
        '---------------------------------------------------------
            'frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas individuales"
            LblBar.Caption = "Facturando: Facturas individuales"
            LblBar.Refresh
            If cadW <> "" Then 'Facturacion pendiente
                cadW = cadW & ")) "
                If Not vFactu.PasarAlbaranesAFactura(TipoAlb, cadW, TextosCSB, ErroresAux, EsTraspasoOfeFAZ) Then
                    If b Then b = False
                    AnyadirAvisos ErroresAux
                Else 'a?adirlo a la lista de facturas a imprimir
                    If ListFactu = "" Then
                        ListFactu = vFactu.NumFactu
                    Else
                        ListFactu = ListFactu & "," & vFactu.NumFactu
                    End If
                End If
                If PgbVisible Then
                    IncrementarProgresNew PBar1, Inc - 1
                    LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                    LblBar.Refresh
                End If
                Espera 0.2
                'Empezamos una nueva Factura
                cadW = ""
            End If
            
            'Los Albaranes que tengan tipofact=1 "factura x Albaran" generar una factura
            'para cada uno de ellos
            cadW = " scaalb.codtipom='" & RSalb!codtipom & "' AND scaalb.numalbar=" & RSalb!NumAlbar
            
            'Generar una Factura nueva
            vFactu.Cliente = RSalb!codClien
            vFactu.NombreClien = RSalb!Nomclien
            vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
            vFactu.CPostal = DBLet(RSalb!codpobla, "T")
            vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
            vFactu.Provincia = DBLet(RSalb!proclien, "T")
            vFactu.NIF = DBLet(RSalb!nifClien, "T")
            vFactu.Telefono = DBLet(RSalb!telclien, "T")
            vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
            vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
            vFactu.Agente = RSalb!CodAgent
            vFactu.ForPago = RSalb!codforpa
            vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
            vFactu.DtoPPago = CCur(RSalb!DtoPPago)
            vFactu.DtoGnral = CCur(RSalb!DtoGnral)

                
                
            If Not vFactu.PasarAlbaranesAFactura(TipoAlb, cadW, TextosCSB, ErroresAux, EsTraspasoOfeFAZ) Then
                If b Then b = False
                AnyadirAvisos ErroresAux
            Else 'a?adirlo a la lista de facturas a imprimir
                If ListFactu = "" Then
                    ListFactu = vFactu.NumFactu
                Else
                    ListFactu = ListFactu & "," & vFactu.NumFactu
                End If
            End If
            If PgbVisible Then
                Inc = 1 '1 albaran x factura
                LblBar.Caption = "Cliente: " & Format(RSalb!codClien, "000000") & " - " & RSalb!Nomclien
                LblBar.Refresh
                IncrementarProgresNew PBar1, Inc
                Inc = 0
            End If
            Espera 0.2
                
            cadW = ""
            
        Else 'tipofac=0 "factura COLECTIVA"
        '----------------------------------------------------------
            'Seleccionar todos los Albaranes pertenecientes a un mismo Cliente,Departamento
            'Los que tengan tipofac=0 "factura colectiva" agruparlos en una misma factura
            'para la misma Forma de PAgo, mismo dtoppago y mismo dtognral
             
             '-- David.      Esta linea da error si no viene de frmlistadoped
             'frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas colectivas"
             LblBar.Caption = "Facturando: Facturas colectivas"
             LblBar.Refresh
             '---- Laura: 06/10/2006
             'Comprobar si es Departamento o Direccion (segun paramatro)
             'DAVID 05/07/2010    Direccion Departamento Obra.  Agrupa <>direccion
             If vParamAplic.HayDeparNuevo > 0 Then
                'agrupar tb por departamento
                condicion = (antClien <> RSalb!codClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             Else
                condicion = (antClien <> RSalb!codClien) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             End If
             
'             If (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral) Then
             If condicion Then
             '-----
                If cadW <> "" Then 'Facturacion PEndiente
                    cadW = cadW & ")) "
                    If Not vFactu.PasarAlbaranesAFactura(TipoAlb, cadW, TextosCSB, ErroresAux, EsTraspasoOfeFAZ) Then
                        If b Then b = False
                        AnyadirAvisos ErroresAux
                    Else 'a?adirlo a la lista de facturas a imprimir
                        If ListFactu = "" Then
                            ListFactu = vFactu.NumFactu
                        Else
                            ListFactu = ListFactu & "," & vFactu.NumFactu
                        End If
                    End If
                    If PgbVisible Then
                        LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                        LblBar.Refresh
                        IncrementarProgresNew PBar1, Inc
                        Inc = 0
                    End If
                    Espera 0.2
                    
                    'Empezamos una nueva Factura
                    cadW = ""
                End If
                'Generar una Factura nueva
                vFactu.Cliente = RSalb!codClien
                vFactu.NombreClien = RSalb!Nomclien
                vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
                vFactu.CPostal = DBLet(RSalb!codpobla, "T")
                vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
                vFactu.Provincia = DBLet(RSalb!proclien, "T")
                vFactu.NIF = DBLet(RSalb!nifClien, "T")
                vFactu.Telefono = DBLet(RSalb!telclien, "T")
                vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
                vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
                vFactu.Agente = RSalb!CodAgent
                vFactu.ForPago = RSalb!codforpa
                vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
                vFactu.DtoPPago = CCur(RSalb!DtoPPago)
                vFactu.DtoGnral = CCur(RSalb!DtoGnral)
                vFactu.Aportacion = 0
                If RSalb!codtipom = "ALM" Then vFactu.Aportacion = DBLet(RSalb!Aportacion, "N")
                cadW = " (scaalb.codtipom='" & RSalb!codtipom & "' AND scaalb.numalbar IN (" & RSalb!NumAlbar
            Else
                cadW = cadW & ", " & RSalb!NumAlbar
            End If
        
            'Guardamos datos del registro anterior
            antClien = RSalb!codClien
'            antDirec = DBLet(RSalb!CodDirec, "N")
            antDirec = actDirec
            antForpa = RSalb!codforpa
            antDtoPP = RSalb!DtoPPago
            antDtoGn = RSalb!DtoGnral
        End If
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
        
    'Facturar la ultima Factura generada del blucle
    If cadW <> "" Then
        cadW = cadW & "))"
        If PgbVisible Then LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
        
        If Not vFactu.PasarAlbaranesAFactura(TipoAlb, cadW, TextosCSB, ErroresAux, EsTraspasoOfeFAZ) Then
            If b Then b = False
            AnyadirAvisos "Error Facturando el Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien & vbCrLf & ErroresAux
        Else 'a?adirlo a la lista de facturas a imprimir
            If ListFactu = "" Then
                ListFactu = vFactu.NumFactu
            Else
                ListFactu = ListFactu & "," & vFactu.NumFactu
            End If
        End If
        If PgbVisible Then
'            LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
            IncrementarProgresNew PBar1, Inc
        End If
        Espera 0.2
    End If
    
    TipoFac = vFactu.codtipom
    Set vFactu = Nothing
    
    
    If b Then
        TraspasoAlbaranesFacturas = True
        LblBar.Caption = "Proceso finalizado correctamente."
        If MostrarMsgOK Then MsgBox "Las Facturas de los Albaranes seleccionados se generaron correctamente.", vbInformation
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        SQL = "ATENCI?N:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    Espera 0.2
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    
    
    If ImprimeLasFacturasGeneradas Then
        If ListFactu <> "" Then
            HazPulsarAceptarEnFrmImprimir = False
            If vTipoM = "ALM" And vParamAplic.EntradaRapidaFacturasMostrador Then HazPulsarAceptarEnFrmImprimir = True
            
            ImprimirFacturas ListFactu, FechaFact, , DevuelveTipoDocumentoFactura(vTipoM), NumeroCopias, False, HazPulsarAceptarEnFrmImprimir

        End If
    End If
    'Voy a imprimir la hoja con las observaciones de la facturacion
    'Es decir si el cliente tiene observaciones de facturacion las mostrara ahora
    If ListFactu <> "" Then InformeObservacionFacturacion_ ListFactu, FechaFact
    
    
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Albaranes", Err.Description
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
    End If
End Function




'#Laura: 14/11/2006 Recuperar facturas Alzira



Private Sub AnyadirAvisos(Donde As String)
    Errores = Errores & vbCrLf & vbCrLf & Donde & vbCrLf
End Sub


Private Sub MostrarAvisos()
    MostrarAvisosPantalla Errores
    'frmMensajes.OpcionMensaje = 13
    'frmMensajes.Show vbModal
End Sub



'========================================================

Public Function ComprobarFechaVenci(FechaVenci As Date, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim newFecha As Date
Dim b As Boolean

'=== Modificada Laura: 23/01/2007
    On Error GoTo ErrObtFec
    b = False
    
    '--- comprobar que tiene dias de pago para obtener nueva fecha
    If Not (Dia1 > 0 Or Dia2 > 0 Or Dia3 > 0) Then
        'si no tiene dias de pago la fecha es OK y fin
        ComprobarFechaVenci = FechaVenci
        Exit Function
    End If
        
    
    '--- Obtener nueva fecha del vencimiento
    newFecha = FechaVenci
    
    Do
        'si dia de la fecha vencimiento es uno de los 3 dias de pagos fecha es OK
        If Day(newFecha) = Dia1 Or Day(newFecha) = Dia2 Or Day(newFecha) = Dia3 Then
'            newFecha = CStr(newFecha)
            b = True
        Else
            'mientras esta en el mismo mes vamos aumentando dias hasta encontrar un dia de pago
            newFecha = DateAdd("d", 1, CDate(newFecha))
        End If
    Loop Until b = True Or Year(newFecha) = Year(FechaVenci) + 3
    
    ComprobarFechaVenci = newFecha
    Exit Function
    
ErrObtFec:
    MuestraError Err.Number, "Obtener Fecha vencimiento seg?n dias de pago.", Err.Description
End Function





Public Function ComprobarFechaVenci_old(FechaVenci As Date, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim fechaV As Date
'Dim cadDias As String
Dim F As String

    fechaV = FechaVenci
    If Dia1 <> 0 Or Dia2 <> 0 Or Dia3 <> 0 Then
        OrdenarDias Dia1, Dia2, Dia3
        If Dia1 >= Day(fechaV) Then
            fechaV = Format(Dia1 & "/" & Month(fechaV) & "/" & Year(fechaV), "dd/mm/yyyy")
        Else
            If Dia2 >= Day(fechaV) Then
                fechaV = Format(Dia2 & "/" & Month(fechaV) & "/" & Year(fechaV), "dd/mm/yyyy")
            Else
                If Dia3 >= Day(fechaV) Then
                    fechaV = Format(Dia3 & "/" & Month(fechaV) & "/" & Year(fechaV), "dd/mm/yyyy")
                
                Else
                    'coger el primero del mes siguiente
                    If Dia1 <> 0 Then
                        F = Dia1 & "/"
                        
                    ElseIf Dia2 <> 0 Then
                        F = Dia2 & "/"
'                        fechaV = Format(Dia2 & "/" & Month(fechaV) + 1 & "/" & Year(fechaV), "dd/mm/yyyy")
                    ElseIf Dia3 <> 0 Then
                        F = Dia3 & "/"
'                        fechaV = Format(Dia3 & "/" & Month(fechaV) + 1 & "/" & Year(fechaV), "dd/mm/yyyy")
                    End If
                    If Month(fechaV) + 1 < 13 Then
                        F = F & Month(fechaV) + 1 & "/" & Year(fechaV)
                    Else
                        F = F & "01/" & Year(fechaV) + 1
                    End If
                    fechaV = Format(F, "dd/mm/yyyy")
                End If
            End If
        End If

    End If
    ComprobarFechaVenci_old = fechaV
End Function





Private Sub OrdenarDias(Dia1 As Byte, Dia2 As Byte, Dia3 As Byte)
'Entran los dias desordenados: dia1=10, dia2=5, dia3=20
'devuelve los dias ordenados: dia1=5, dia2=10, dia3=20
Dim diaAux As Byte

    On Error GoTo EOrdenar

    If Dia1 < Dia2 And Dia1 < Dia3 Then
        'dia 1 es el menor
        If Dia2 > Dia3 Then
            diaAux = Dia2
            Dia2 = Dia3
            Dia3 = diaAux
        End If
    ElseIf Dia2 < Dia3 Then
        'dia2 es el menor
        diaAux = Dia1
        Dia1 = Dia2
        If diaAux < Dia3 Then
            Dia2 = diaAux
        Else
            Dia2 = Dia3
            Dia3 = diaAux
        End If
    Else
        'dia3 es el menor
        diaAux = Dia1
        Dia1 = Dia3
        If diaAux < Dia2 Then
            Dia3 = Dia2
            Dia2 = diaAux
        Else
            Dia3 = diaAux
        End If
    End If

EOrdenar:
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function ComprobarMesNoGira(FecVenci As Date, MesNG As Byte, DiaVtoAt As Byte, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim F As String
Dim diaPago As Byte

    If Month(FecVenci) = MesNG Then
        '### LAURA 14/08/2008
'        If DiaVtoAt > 0 Then
'            F = DiaVtoAt & "/"
'        Else
'            F = Day(FecVenci) & "/"
'        End If
        
'        If Month(FecVenci) + 1 < 13 Then
'            F = F & Month(FecVenci) + 1 & "/" & Year(FecVenci)
'        Else
'            F = F & "01/" & Year(FecVenci) + 1
'        End If

        If DiaVtoAt > 0 Then
            'si tiene dia de vto atrasado a ese dia del mes siguiente
            'al mes a no girar
            F = DiaVtoAt & "/"
            F = F & Month(FecVenci) & "/" & Year(FecVenci)
            F = DateAdd("m", 1, F)
        Else
            'si no tiene dia de vto atrasado el primer dia de pago
            'del mes siguiente si tiene o sino el siguiente mes del
            'vencimiento obtenido
            If Dia1 > 0 Or Dia2 > 0 Or Dia3 > 0 Then
                'tiene dias de pago: el menor dia del mes siguiente
                diaPago = Dia1
                If (diaPago = 0) Or ((Dia2 < diaPago) And Dia2 <> 0) Then diaPago = Dia2
                If (diaPago = 0) Or ((Dia3 < diaPago) And Dia3 <> 0) Then diaPago = Dia3
                
                F = diaPago & "/"
                F = F & Month(FecVenci) & "/" & Year(FecVenci)
            Else
                'no tiene dias de pago: al mes siguiente
                F = Day(FecVenci) & "/"
                F = F & Month(FecVenci) & "/" & Year(FecVenci)
            End If
            
            F = DateAdd("m", 1, F)
        End If
        '###
        
        FecVenci = Format(F, "dd/mm/yyyy")
    End If
    
    ComprobarMesNoGira = FecVenci
End Function





Public Sub ImprimirHojaExpedicion(OpcionListado As Byte, NumAlb As String, tipMov As String, Optional fecAlb As String)
Dim cadFormula As String
Dim CadParam As String
Dim cadSelect As String 'select para insertar en tabla temporal
Dim numParam As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim ImpresionDirecta As Boolean
'Dim codClien As String
'Dim EsHistorico As Boolean
Dim NombreTabla As String
Dim NomTablaLineas As String

    If NumAlb = "" Then
        MsgBox "Debe seleccionar un Albaran para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    
'    EsHistorico = (fecAlb <> "")
    
'    If EsHistorico <> "" Then 'es historico
        NombreTabla = "scaalb"
        NomTablaLineas = "slialb" 'Tabla lineas de Albaranes
'    Else
'        NombreTabla = "schalb"
'        NomTablaLineas = "slhalb"
'    End If
    
    
    '===================================================
    '============ PARAMETROS ===========================
'    If (OpcionListado = 45) Then
'        If EsInformePortes Then
            'Es el de portes
             indRPT = 34
'        Else
'            If hcoCodTipoM = "ALZ" Then
'                indRPT = 29   'Albaranes B
'            Else
'                If EsHistorico Then
'                    indRPT = 11 'Hist. Albaranes clientes
'                Else
'                    indRPT = 10 'Albaran Clientes
'                End If
'            End If
'        End If
'    End If
    
    If Not PonerParamRPT2(indRPT, CadParam, numParam, nomDocu, ImpresionDirecta, pPdfRpt, pRptvMultiInforme) Then Exit Sub
   
    'A?adir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    CadParam = CadParam & "pCodUsu=" & vUsu.codigo & "|"
    numParam = numParam + 1
    
    'PORTES
    CadParam = CadParam & "vPortes=""" & vParamAplic.ArtPortesN & """|"
    numParam = numParam + 1
    
'    'PUNTO VERDE
'    cadParam = cadParam & "PuntoVerde=""" & vParamAplic.ArtReciclado & """|"
'    numParam = numParam + 1
    
'    'Si se imprimen importes y/o
'    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", codClien, "N")
'    If devuelve = "" Then devuelve = "0"
'    ' 0 "Todo"
'    ' 1 "Cantidad y Precio"
'    ' 2 "Cantidad"
'    cadParam = cadParam & "Albarcon=" & devuelve & "|"
'    numParam = numParam + 1
    
    
    'Nombre fichero .rpt a Imprimir
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
    If Not ImpresionDirecta Then
        frmImprimir.NombreRPT = nomDocu
        frmImprimir.NombrePDF = pPdfRpt
    End If
        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N? de Albaran
    '---------------------------------------------------
    If NumAlb <> "" Then
        '- Cod Tipo Movimiento
        devuelve = "{" & NombreTabla & ".codtipom}='" & tipMov & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        '- N? Albaran
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(NumAlb)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
        
'        If EsHistorico <> "" Then 'historico
'            'El campo fecha tambien es clave primaria
'            devuelve = fecAlb
'            devuelve = "{" & NombreTabla & ".fechaalb}=Date(" & Year(devuelve) & "," & Month(devuelve) & "," & Day(devuelve) & ")"
'            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'
'            devuelve = "{" & NombreTabla & ".fechaalb}='" & Format(fecAlb, FormatoFecha) & "'"
'            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
'        End If
    End If
   
'    '=========================================================================
'    'Aqui sabemos que valor tiene CodClien y a?adimos a los parametros el tipo de IVA
'    'que se aplica a ese cliente
'    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", codClien, "N")
'    If devuelve <> "" Then
'        cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
'        numParam = numParam + 1
'    End If

        
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    devuelve = NombreTabla & " INNER JOIN " & NomTablaLineas & " ON "
    devuelve = devuelve & NombreTabla & ".codtipom=" & NomTablaLineas & ".codtipom AND " & NombreTabla & ".numalbar= " & NomTablaLineas & ".numalbar "
'    If EsHistorico Then devuelve = devuelve & " AND " & NombreTabla & ".fechaalb= " & NomTablaLineas & ".fechaalb "
    If Not HayRegParaInforme(devuelve, cadSelect) Then Exit Sub
    
    
    If ImpresionDirecta Then
        'Imrpimie directamente. Tipo 4tonda.  -----------
        If MsgBox("?Imprimir el albar?n?", vbQuestion + vbYesNo) = vbYes Then ImprimirDirectoAlb cadSelect
    Else
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = OpcionListado
            If indRPT = 34 Then
                .Titulo = "Portes albaran "
'            Else
'                .Titulo = "Albaran de Cliente"
            End If
            .ConSubInforme = True
            .Show vbModal
        End With
    End If
End Sub


'FormatoFactura:
'               0.- Normal
'               1.- TPV
'               2.- Factura "B"
'               3.- Factura telefonia FAT
'               EULER
'               4.- Orden de trabajo
'               5.- Trabajo exterior
Public Sub ImprimirFacturas(listaF As String, fechaF As String, Optional SQL As String, Optional FormatoFactura As Byte, Optional NumeroCopias As Byte, Optional OrdenadoPorCliente As Boolean, Optional HazPulsarAceptar As Boolean)
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NombreTabla As String
Dim ImprimeDirecto As Boolean
Dim RN As ADODB.Recordset
Dim ListaFacturasDefinitiva As String
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    NombreTabla = "scafac"

    
    'Mayo 2015
    'NO imprimiremos las que el cliente tenga la marca de enviar por email
    ' Soalucion. De ListaF que lleva las facturas, quetare las los clientes lleven la marca
    
    devuelve = " coalesce(EnvFraEmail,0)=0 AND " & NombreTabla & ".codtipom='" & TipoFac & "' AND "
    devuelve = devuelve & NombreTabla & ".numfactu IN (" & listaF & ")"
    devuelve = devuelve & " AND year(" & NombreTabla & ".fecfactu) = " & Year(fechaF)
    devuelve = "Select numfactu from scafac,sclien WHERE scafac.codclien =sclien.codclien AND " & devuelve
    Set RN = New ADODB.Recordset
    RN.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ListaFacturasDefinitiva = ""
    While Not RN.EOF
       ListaFacturasDefinitiva = ListaFacturasDefinitiva & ", " & RN!NumFactu
       RN.MoveNext
    Wend
    RN.Close
    If ListaFacturasDefinitiva = "" Then
        ListaFacturasDefinitiva = "-1"
    Else
        ListaFacturasDefinitiva = Mid(ListaFacturasDefinitiva, 2) 'quito la primera coma
    End If
     
    Set RN = Nothing
    
    CadParam = ""
    cadSelect = ""


    '===================================================
    '============ PARAMETROS ===========================
    If FormatoFactura = 0 Then
        indRPT = 12 'Facturas Clientes  NORMAL
    ElseIf FormatoFactura = 1 Then
        indRPT = 18 'FACTURAS TPV
    
    ElseIf FormatoFactura = 2 Then
        indRPT = 30 'FACTURAS "B"
    ElseIf FormatoFactura = 3 Then
        indRPT = 63 'Telefonia
        
    ElseIf FormatoFactura = 4 Then
        indRPT = 78  'orden de trabajo
    ElseIf FormatoFactura = 5 Then
        indRPT = 79     'trabjo exterior
    End If
    
    
    If OrdenadoPorCliente Then
        
        'VA ordenado por cliente
        If FormatoFactura = 0 Then indRPT = 72 'Facturas Clientes  NORMAL
    End If
    
    
    
    If Not PonerParamRPT2(indRPT, CadParam, numParam, nomDocu, ImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        Exit Sub
    End If



    'PUNTO VERDE
    '--------------------------------------------------------------------------
    CadParam = CadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
    numParam = numParam + 1
    

    'Nombre fichero .rpt a Imprimir
    If Not ImprimeDirecto Then
        frmImprimir.NombreRPT = nomDocu
        frmImprimir.NombrePDF = pPdfRpt
    End If

    If SQL <> "" Then
        'Llamo desde el menu de Reimprimir facturas y tengo construida la
        'cadena de seleccion D/H tipoMov, D/H NumFactu, D/H fecfactu
        cadSelect = SQL
        cadFormula = ListaFacturasDefinitiva
        CadParam = CadParam & fechaF
        numParam = numParam + 1
    Else
        'Llama desde PasarAlbaranes a  Facturas y al terminar las imprime
        '===================================================
        '================= FORMULA =========================
        'Cadena para seleccion N? de Factura
        '---------------------------------------------------
        'Cod Tipo Movimiento
        devuelve = "({" & NombreTabla & ".codtipom}='" & TipoFac & "') "
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'N? Factura
        devuelve = "({" & NombreTabla & ".numfactu} IN [" & ListaFacturasDefinitiva & "])"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'fecha factu
        devuelve = "(year({" & NombreTabla & ".fecfactu}) = " & Year(fechaF) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub

        

        cadSelect = cadFormula

    
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect, True) Then Exit Sub


     If ImprimeDirecto Then
         'Abrire un formulario por si acaso quieren cancelar la impresion. Ya que al ser
         'directa puede tardar mucho, haberse equivocado ......
        CadenaDesdeOtroForm = cadSelect
        frmVarios.Opcion = 0
        frmVarios.Show vbModal
        'Ha terminado la reimpresion
        
     Else
     
         With frmImprimir
                .NumeroCopias = NumeroCopias
                .FormulaSeleccion = cadFormula
                .OtrosParametros = CadParam
                .NumeroParametros = numParam
                .SoloImprimir = False
                .EnvioEMail = False
                .PulsaAceptar = HazPulsarAceptar
                .Opcion = 53
                .Titulo = ""
                .SeleccionaRPTCodigo = pRptvMultiInforme
                .Show vbModal
        End With
    End If
End Sub



Public Function TraspasoMtosAFacturas(cadSQL As String, cadSel As String, FechaFact As String, OpeFact As String, banPr As String, MesFact As String, ByRef Lbl As Label, CentroCoste As String) As Boolean      'Fecha de la factura, Operador
'IN -> cadSQL: cadena para seleccion de los mantenimientos que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      OpeFact: Operador Factura
'
'   CentroCoste.   Si tiene analitica y el modoanalitica es por poryecto, es un dato del formulario
'
'Desde Mantenimientos Genera las Facturas correspondientes
Dim RSmto As ADODB.Recordset 'Ordenados por: clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim SQL As String

Dim vClien As CCliente 'aqui cargamos los datos del cliente del mantenimiento para grabar en scafac
Dim vFactu As CFactura

Dim ListFactu As String
Dim Conta2 As Long

    On Error GoTo ETraspasoMtoFac


    TraspasoMtosAFacturas = False
    
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("VENFAC") 'facturas de mantenimiento
    If Not BloqueoManual("VENFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los mantenimientos que vamos a facturar (cabeceras y lineas)
'    SQL = " (scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien ) INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
    SQL = " scaman "
    
    If Not BloqueaRegistro(SQL, cadSel) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        Exit Function
    End If
    
    
    
    
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactu = New CFactura
    vFactu.FecFactu = FechaFact 'Fecha para las Facturas

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", banPr, "N")
    
    OpeFactu = OpeFact 'operador de la factura de mantenimiento
    MesFactu = MesFact 'mes a factura para los mantenimientos
    
    b = True
    
    'Marcar Mantenimientos que se van a Facturar
    '----------------------------------------
    
    SQL = cadSQL & " ORDER BY scaman.codclien, scaman.coddirec, scaman.nummante "
    Set RSmto = New ADODB.Recordset
    Conta2 = InStr(1, cadSQL, " FROM ")
    ListFactu = "Select count(*) " & Mid(cadSQL, Conta2)
    
    
    
    RSmto.Open ListFactu, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Lbl.Tag = RSmto.Fields(0)
    RSmto.Close
    
    
    
    Conta2 = 0
    ListFactu = ""
    RSmto.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Le pongo                KEYSET      pq quiero contar los registros
    'Cada MAntenimiento genera una factura
    'Calcular y Grabar Factura en las Tablas de Facturas
    '---    -------------------------------------------------
     While Not RSmto.EOF
            
           Conta2 = Conta2 + 1
           Lbl.Caption = Conta2 & " de " & Lbl.Tag
           Lbl.Refresh
            
            If (RSmto.RecordCount Mod 10) = 9 Then DoEvents
        'para cada mantenimiento de la tabla scaman seleccionado para facturar
        vFactu.BrutoFac = CCur(RSmto!Importe)
        'tipo de contrato del mantenimientos
        TipCoMan = RSmto!codtipco
        
        'Datos de la Cabecera: Insertar en scafac
        '-----------------------------------------
        Set vClien = New CCliente
        If vClien.LeerDatos(RSmto!codClien) Then
            'Datos cliente
            vFactu.Cliente = RSmto!codClien
            vFactu.NombreClien = vClien.Nombre
            vFactu.DomicilioClien = vClien.Domicilio
            vFactu.CPostal = vClien.CPostal
            vFactu.Poblacion = vClien.Poblacion
            vFactu.Provincia = vClien.Provincia
            vFactu.NIF = vClien.NIF
            vFactu.Telefono = vClien.TfnoClien
            vFactu.DirDpto = DBLet(RSmto!CodDirec, "T")
            vFactu.NombreDirDpto = DBLet(RSmto!nomdirec, "T")
            vFactu.Agente = vClien.Agente
            'forma de pago del mantenimiento
            vFactu.ForPago = RSmto!codforpa
            vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSmto!codforpa, "N")
            
            vFactu.DtoGnral = 0
            vFactu.DtoPPago = 0
            vFactu.Banco = DBLet(vClien.Banco, "N")
            vFactu.Sucursal = DBLet(vClien.Sucursal, "N")
            vFactu.DigControl = DBLet(vClien.DigControl, "T")
            vFactu.CuentaBan = DBLet(vClien.CuentaBan, "T")
            vFactu.IBAN = DBLet(vClien.IBAN, "T")
            
            vFactu.Observacion = DBLet(RSmto!concefac, "T")
                
            
            
            
            If Not vFactu.PasarMtosAFactura(TipCoMan, OpeFactu, MesFactu, RSmto!nummante, CentroCoste) Then
                If b Then b = False
            Else
                vClien.ActualizaUltFecMovim (FechaFact)
                
                
                'a?adirlo a la lista de facturas a imprimir
                If ListFactu = "" Then
                    ListFactu = vFactu.NumFactu
                Else
                    ListFactu = ListFactu & "," & vFactu.NumFactu
                End If
            End If
        End If
        Set vClien = Nothing
        RSmto.MoveNext
    Wend
    
    RSmto.Close
    Set RSmto = Nothing
    
    Set vFactu = Nothing
    Lbl.Caption = "Finalizando proceso"
    Lbl.Refresh
    If b Then
        MsgBox "Las Facturas de los Mantenimientos seleccionados se generaron correctamente.", vbInformation
    Else
        SQL = "ATENCI?N:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbInformation
    End If
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    
    If ListFactu <> "" Then
        Lbl.Caption = "Imprimiendo"
        Lbl.Refresh
        ImprimirFacturaMan 53, ListFactu, FechaFact
        
        TipoFac = "FAM"
        InformeObservacionFacturacion_ ListFactu, FechaFact
        
    End If
    
    
ETraspasoMtoFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Mantenimientos", Err.Description
    End If
End Function




Private Sub ImprimirFacturaMan(OpcionListado As Byte, ListFactu As String, FecFactu As String)
'Imprime una factura de Mantenimiento
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NombreTabla As String
    
    NombreTabla = "scafac"
    
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    pRptvMultiInforme = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 53) Then indRPT = 12 'Facturas Clientes
    If Not PonerParamRPT2(indRPT, CadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then
        Exit Sub
    End If
      
    frmImprimir.SeleccionaRPTCodigo = pRptvMultiInforme
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    frmImprimir.NombrePDF = pPdfRpt
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N? de Factura
    '---------------------------------------------------
    'Cod Tipo Movimiento
    devuelve = "{" & NombreTabla & ".codtipom}='FAM'"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    cadSelect = cadFormula
    
    'N? Factura
    devuelve = "{" & NombreTabla & ".numfactu} IN [" & ListFactu & "]"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "{" & NombreTabla & ".numfactu} IN (" & ListFactu & ")"
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    'Fecha Factura
    devuelve = "year({" & NombreTabla & ".fecfactu})=" & Year(FecFactu)
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Fecha Factura en cadSelect
'        devuelve = "{" & NombreTabla & ".fecfactu}= '" & Format(FecFactu, FormatoFecha) & "'"
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
   
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
     With frmImprimir
            .NumeroCopias = vParamAplic.NumCopiasFacturacion
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = OpcionListado
            .Titulo = ""
            .Show vbModal
    End With
End Sub






'Ventas de TICKET
'=================================================================
Public Function EliminarVenta(cadSQL As String) As Boolean
'Eliminamos de las tablas de ventas: scaven, sliven
Dim SQL As String

    On Error GoTo EElimVen

    EliminarVenta = False
    
    'Diciembre 2012
    'Se pueden asociar tikets a campos
    SQL = "DELETE FROM sliven2 "
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven2")
    conn.Execute SQL
    
    'ELiminar lineas venta
    SQL = "DELETE FROM sliven "
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    conn.Execute SQL
    
   
    SQL = "DELETE FROM slivenlotes "
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "slivenlotes")
    conn.Execute SQL

    
    'Eliminar Cabeceras venta
    SQL = "DELETE FROM scaven "
    SQL = SQL & " WHERE " & Replace(cadSQL, "sliven", "scaven")
    conn.Execute SQL
        
    EliminarVenta = True

EElimVen:
    If Err.Number <> 0 Then
        EliminarVenta = False
        Err.Raise Err.Number, "Eliminar venta." & Err.Description
        
    Else
        EliminarVenta = True
    End If
End Function




Private Function DevuelveTipoDocumentoFactura(ByRef TipoAlbaran As String) As Byte
    DevuelveTipoDocumentoFactura = 0
    If TipoAlbaran <> "" Then
        If TipoAlbaran = "ATI" Then
            'Factura de tickets
            TipoAlbaran = 1
            DevuelveTipoDocumentoFactura = 1
        Else
            If TipoAlbaran = "ALZ" Then
                TipoAlbaran = 2
                DevuelveTipoDocumentoFactura = 2
            Else
                If TipoAlbaran = "ALO" Then
                    DevuelveTipoDocumentoFactura = 4
                ElseIf TipoAlbaran = "ALE" Then
                    DevuelveTipoDocumentoFactura = 5
                End If
            End If
        End If
    End If
    
End Function






'*****************************************************************
'
'
'   Mayo 2012.  Se facturara por cliente, DEPARTAMENTO, dependiendo
'   del parametro del cliente.
'   Que vamos a hacer....
'       Llamaremos a la funcion facturar renting desde un RS que har? el select
'

Public Function FacturarRenting(cadSQL As String, Fecfact As String, OpeFact As String, banPr As String, ByRef Lbl As Label, CentroCoste As String, SoloID As String, PeriodoFacturar As Date) As Boolean
Dim ListFactu As String
Dim Aux As String
Dim R As ADODB.Recordset
Dim b As Boolean
Dim Ik As Boolean
Dim ColClientes As Collection
Dim idCliente As Long
Dim PorDep As String
Dim ElDepartamento As String
        'comprobamos que no haya nadie facturando
    DesBloqueoManual ("RENTFAC") 'facturas de mantenimiento
    If Not BloqueoManual("RENTFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando " & RentingLB, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    




    'Mayo 2012
    'Los clientes se puede facturar por departamentos
    'Esto sginifica que sera una factura por cada departamento
    Set R = New ADODB.Recordset
    ListFactu = ""
    If SoloID <> "" Then
        'Es solo UNO.
        'Si factura por departamento, la factura ira a ese departamento
        Aux = ""
        b = FacturarRentingCliDpto(cadSQL, Fecfact, OpeFact, banPr, Lbl, CentroCoste, SoloID, PeriodoFacturar, ListFactu, Aux)
        
    Else
        ' '----------------------------------------
        Aux = "SELECT sclienrenting.codclien,coddirec,Rentin_x_dpto FROM sclienrenting,sclien WHERE sclienrenting.codclien=sclien.codclien AND " & cadSQL
        Aux = Aux & " ORDER BY sclienrenting.codclien,coddirec"
        
        R.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        idCliente = -1
        Set ColClientes = New Collection
        While Not R.EOF
            If idCliente <> R!codClien Then
                idCliente = R!codClien
                ElDepartamento = "@"
                If DBLet(R!Rentin_x_dpto, "N") = 0 Then
                    'Factura UNICA por cliente
                    Aux = "| sclienrenting.codclien = " & idCliente
                    ColClientes.Add Aux & "|"
                End If
            End If
            
            If DBLet(R!Rentin_x_dpto, "N") = 1 Then
                    'Factura por cliente departamento
                    If DBLet(R!CodDirec, "T") <> ElDepartamento Then
                        Aux = " sclienrenting.codclien = " & idCliente
                        If IsNull(R!CodDirec) Then
                            Aux = Aux & " AND coddirec is null"
                            Aux = "|" & Aux
                        Else
                            Aux = Aux & " AND coddirec =" & R!CodDirec
                            Aux = R!CodDirec & "|" & Aux
                        End If
                        ColClientes.Add Aux & "|"
                        ElDepartamento = DBLet(R!CodDirec, "T")
                    End If
            End If
            
            R.MoveNext
        Wend
        R.Close
        Set R = Nothing
        
        b = True
        For idCliente = 1 To ColClientes.Count
            Aux = ColClientes.item(idCliente)
            
            PorDep = RecuperaValor(Aux, 1)  'Si es por cliente departamento
            
            Aux = cadSQL & " AND " & RecuperaValor(Aux, 2)
            Ik = FacturarRentingCliDpto(Aux, Fecfact, OpeFact, banPr, Lbl, CentroCoste, SoloID, PeriodoFacturar, ListFactu, PorDep)
            If b Then
                If Not Ik Then b = False
            End If
            Lbl.Caption = "Actualizando....."
            Lbl.Refresh
            Espera 0.5
        Next
    End If

    If b Then
        MsgBox "Las Facturas de alquiler/" & RentingLB & " seleccionados se generaron correctamente.", vbInformation
        
    Else
        Aux = "ATENCI?N:" & vbCrLf
        MsgBox Aux & "No todas las Facturas se generaron correctamente!!!.", vbCritical
    End If
    
    
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("RENTFAC")
    TerminaBloquear
    
    If ListFactu <> "" Then
        Lbl.Caption = "Imprimiendo"
        Lbl.Refresh
        ImprimirFacturaMan 53, ListFactu, Fecfact
        
        TipoFac = "FAM"
        InformeObservacionFacturacion_ ListFactu, Fecfact
        
        
    End If

End Function


Private Function FacturarRentingCliDpto(cadSQL As String, Fecfact As String, OpeFact As String, banPr As String, ByRef Lbl As Label, CentroCoste As String, SoloID As String, PeridoFacturar As Date, ByRef ListadoFacturas As String, PorDepartamento As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los renting que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      OpeFact: Operador Factura
'
'   CentroCoste.   Si tiene analitica y el modoanalitica es por poryecto, es un dato del formulario
'
'Desde Mantenimientos Genera las Facturas correspondientes
Dim RSmto As ADODB.Recordset 'Ordenados por:
Dim b As Boolean
Dim SQL As String

Dim vClien As CCliente 'aqui cargamos los datos del cliente del renting para grabar en scafac
Dim vFactu As CFactura


Dim Conta2 As Long
Dim I As Integer
Dim Aux2 As String
Dim TipoFacturacion As Byte  '1: mensual   3:trimestral   6:semestral   12:anual

    On Error GoTo ETraspasoMtoFac


    FacturarRentingCliDpto = False
    

    
        
    'EMPEZAMOS LA FACTURA
    Set vFactu = New CFactura
    vFactu.FecFactu = Fecfact 'Fecha para las Facturas

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", banPr, "N")
    
    OpeFactu = OpeFact 'operador de la factura de mantenimiento
    
    
    b = True
    
    'Marcar Mantenimientos que se van a Facturar
    '----------------------------------------
    cadSQL = " FROM sclienrenting,sclien WHERE sclienrenting.codclien=sclien.codclien AND " & cadSQL
    SQL = cadSQL & "  GROUP BY sclienrenting.codclien ORDER BY sclienrenting.codclien"
    Set RSmto = New ADODB.Recordset
    Conta2 = InStr(1, cadSQL, " FROM ")

    
    
    
    RSmto.Open "Select count(*) " & Mid(cadSQL, Conta2), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Lbl.Tag = RSmto.Fields(0)
    RSmto.Close
    
    
    
    Conta2 = 0
    
    SQL = "SELECT sclienrenting.codclien,codtipco,sum(importe) importe " & SQL
    RSmto.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Le pongo                KEYSET      pq quiero contar los registros
    'Cada MAntenimiento genera una factura
    'Calcular y Grabar Factura en las Tablas de Facturas
    '---    -------------------------------------------------
     While Not RSmto.EOF
            
        Conta2 = Conta2 + 1
        Lbl.Caption = Conta2 & " de " & Lbl.Tag
        Lbl.Refresh
         
        If (RSmto.RecordCount Mod 10) = 9 Then DoEvents
        'para cada mantenimiento de la tabla scaman seleccionado para facturar
        
        Aux2 = DevuelveDesdeBD(conAri, "TipoFraRenting", "sclien", "codclien", CStr(RSmto!codClien))
        If Aux2 = "" Then Aux2 = "1"
        TipoFacturacion = CByte(Aux2)
        
        If SoloID <> "" Then
            'Es un cobro parcial del mantenimiento que se contrata AHORA.
            'Se facturar desde la fecha de alta hasta el ultimo dia del mes actual
            I = DiasMes(Month(PeridoFacturar), Year(PeridoFacturar))
            vFactu.BrutoFac = CCur(RSmto!Importe) / I 'Vamos a calcular el importe DIA
            
            Aux2 = Mid(cadSQL, InStr(1, cadSQL, "AND id =") + 8)
            Aux2 = Mid(Aux2, 1, InStr(1, Aux2, " AND ") - 1)
            
            Aux2 = DevuelveDesdeBD(conAri, "fecalta", "sclienrenting", "codclien = " & RSmto!codClien & " AND id", Aux2, "N")
            
            
            
            I = DateDiff("d", CDate(Aux2), CDate(I & Format(PeridoFacturar, "/mm/yyyy")))
            If I > 0 Then
                 vFactu.BrutoFac = Round2(I * vFactu.BrutoFac, 2)
            
            
            Else
                vFactu.BrutoFac = 0
            
            End If
            
            
            
            
            
        Else
            vFactu.BrutoFac = CCur(RSmto!Importe) * TipoFacturacion
        End If
        'tipo de contrato del mantenimientos
        TipCoMan = RSmto!codtipco 'cojo uno de ellos, solo para poder ver los iVAS
        
        'Datos de la Cabecera: Insertar en scafac
        '-----------------------------------------
        Set vClien = New CCliente
        If vClien.LeerDatos(RSmto!codClien) Then
            'Datos cliente
            vFactu.Cliente = RSmto!codClien
            vFactu.NombreClien = vClien.Nombre
            vFactu.DomicilioClien = vClien.Domicilio
            vFactu.CPostal = vClien.CPostal
            vFactu.Poblacion = vClien.Poblacion
            vFactu.Provincia = vClien.Provincia
            vFactu.NIF = vClien.NIF
            vFactu.Telefono = vClien.TfnoClien
            vFactu.DirDpto = ""   'SE FACTURA A UN CLIENTE, no a al departamento
            vFactu.NombreDirDpto = ""
        
            vFactu.Agente = vClien.Agente
            'forma de pago del mantenimiento
            vFactu.ForPago = vClien.ForPago
            vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", vClien.ForPago, "N")
            
            vFactu.DtoGnral = 0
            vFactu.DtoPPago = 0
            vFactu.Banco = DBLet(vClien.Banco, "N")
            vFactu.Sucursal = DBLet(vClien.Sucursal, "N")
            vFactu.DigControl = DBLet(vClien.DigControl, "T")
            vFactu.CuentaBan = DBLet(vClien.CuentaBan, "T")
            vFactu.IBAN = DBLet(vClien.IBAN, "T")
            
            'Cliente /departamento
            If PorDepartamento <> "" Then
                Aux2 = "codclien = " & vFactu.Cliente & " AND coddirec"
                Aux2 = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Aux2, PorDepartamento)
                vFactu.DirDpto = PorDepartamento
                vFactu.NombreDirDpto = Aux2
           
            End If
            vFactu.Observacion = ""   'DBLet(RSmto!concefac, "T")
                
            
            
            Aux2 = Mid(cadSQL, InStr(1, cadSQL, "FROM") + 25) 'EL SQL para seleccionar los datos de las lineas
            If Not vFactu.PasarRentingAFactura(TipCoMan, OpeFactu, CentroCoste, Aux2, PeridoFacturar) Then
                If b Then b = False
            Else
                vClien.ActualizaUltFecMovim (Fecfact)
                
                
                'a?adirlo a la lista de facturas a imprimir
                If ListadoFacturas = "" Then
                    ListadoFacturas = vFactu.NumFactu
                Else
                    ListadoFacturas = ListadoFacturas & "," & vFactu.NumFactu
                End If
            End If
        End If
        Set vClien = Nothing
        RSmto.MoveNext
    Wend
    
    RSmto.Close
    
    Lbl.Caption = "Finalizando proceso"
    Lbl.Refresh
   
    FacturarRentingCliDpto = b
    
ETraspasoMtoFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Mantenimientos", Err.Description
    End If
    Set RSmto = Nothing
    Set vFactu = Nothing
End Function




Private Sub InformeObservacionFacturacion_(ListFactu As String, FechaFact As String)
Dim Aux As String
    
    
    If TipoAlb = "ALM" Then Exit Sub
    
    
    Aux = "DELETE FROM tmpcrmcobros WHERE codusu = " & vUsu.codigo
    conn.Execute Aux
    
    Aux = "insert into tmpcrmcobros(codusu,secuencial,fecfaccl,fecha2,tipo,importe)"
    Aux = Aux & " select " & vUsu.codigo & ",scafac.codclien,curdate()," & DBSet(FechaFact, "F")
    Aux = Aux & ", count(*) ,sum(totalfac) from scafac,sclien where scafac.codclien=sclien.codclien AND "
    Aux = Aux & " (scafac.codtipom='" & TipoFac & "')  AND ((scafac.numfactu IN (" & ListFactu & "))) AND ((year(scafac.fecfactu) = " & Year(FechaFact) & "))"
    Aux = Aux & " and obsfacturacion <>"""" group by codclien"
    conn.Execute Aux

    Espera 0.2
    Aux = DevuelveDesdeBD(conAri, "count(*)", "tmpcrmcobros", "codusu", vUsu.codigo)
    If Aux = "" Then Aux = "0"
    
    
    If Val(Aux) > 0 Then
    
        
        With frmImprimir
            .FormulaSeleccion = "{tmpcrmcobros.codusu} = " & vUsu.codigo
            .OtrosParametros = "pEmpresa=""" & vEmpresa.nomempre & """|"""
            .NumeroParametros = 1
    
            .SoloImprimir = False
            .EnvioEMail = False
            .Titulo = "Observaciones facturacion"
            .Opcion = 3000   'VAN TODOS EN ESTE SACO
            .NombrePDF = "rObservaFracion.rpt"
            .NombreRPT = .NombrePDF
            .ConSubInforme = False
            .MostrarTreeDesdeFuera = False
            .Show vbModal
        End With
    
    End If
End Sub






'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'
'   Facturacion telefonia
'   =====================
'   Tipo ALZIRA
'   Las facturas YA estan creadas en las tablas: tel_cab_factura tel_lin_factura_cuotas....
'   Con lo cual, ahora, desde esa tabla creamos el albaran.
'   El resumen de la linea va al articulo de telefonia.
'   El numero de factura, y la fecha SON las indicadas en la linea
Public Function traspasofacturasTelefonia(Fichero As String, ByRef L As Label, ByVal idBanco As Integer) As Boolean
Dim b As Boolean
    
    traspasofacturasTelefonia = False
    
    'Bloqueamos proceso
    If Not BloqueoManual("VENFAC", "1") Then
        'MsgBox "No se puede facturar TELEFONIA. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If



    

    'Proceso 1. MEter en slialb,scaalb
    'EL NUmero de albaran, sera el mismo que el numero de factura
    TipoFac = Fichero
    b = GenerarAlbaranesTelefonia(L)
    cadW = ""  'reutiolizadas
    Errores = "" 'reutilzadas
    TipoFac = ""
    LetraSer = ""
    
    'Proceso 2.  FACTURAR
    If b Then
        DoEvents
        
        'MEtemos el nombre del fichero en los datos traspasados
        cadW = "referenc='" & Fichero & "' AND codtipom"
        cadW = DevuelveDesdeBD(conAri, "count(*)", "scaalb", cadW, "ALT", "T")
        If cadW = "" Then cadW = "0"
        
        Errores = "referenc='" & Fichero & "' AND codtipom"
        Errores = DevuelveDesdeBD(conAri, "count(*)", "scaalb", Errores, "ALI", "T")
        If Errores = "" Then Errores = "0"
        
        cadW = " VALUES ('" & Fichero & "'," & DBSet(Now, "F") & "," & cadW & "," & Errores & ")"
        cadW = "INSERT INTO tel_fichtraspasados(Fichero,Fecha,FraNormal,FraInt)" & cadW
        conn.Execute cadW  'Fichero YA procesado
        
        'PARA LAS INTERNAS
        '
        'Generaremos la factura en scafac y con el numero que obtenemos
        'updatearemos tel_cabfactura....
        'Por si acaso diera algun fallo y no se renumerara(no deberia pasar)
        'lo guardamos en tmpstockfec para avisar
        conn.Execute "DELETE FROM tmpcrmmsg where codusu = " & vUsu.codigo
        
        
        'Generaremos las ALT, las normales
        'y 'las ALI que sean de telefonia
        TipoAlb = Fichero
        b = GenerarFacturasTelefonia(idBanco, L, True, False)
        If b Then traspasofacturasTelefonia = True
    End If
    
    
    If Not b Then
        
        Errores = DevuelveDesdeBD(conAri, "count(*)", "scaalb", "referenc", Fichero, "T")
        If Errores <> "" Then MsgBox "Se han quedado " & Errores & " albaranes. Consulte soporte t?cnico[tmpcrmmsg]", vbExclamation
            
    End If
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    

    
End Function

Public Function GenerarFacturasTelefonia(banPr As Integer, LblBar As Label, FrasNormales As Boolean, Coarval As Boolean) As Boolean
' Dos pasos. Primero las fras normales de telefonia
'            Segundo   las internas
Dim RSalb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim SQL As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antClien As Long


Dim vFactu As CFactura
Dim Inc As Integer

Dim J As Byte  'Dos veces el bucle. Primero las internas, despues el resto
Dim Fichero As String

Dim RTT As ADODB.Recordset


    antClien = 0 'cliente
    
    
    
    Errores = ""
    Inc = 0
    
    
    Set vFactu = New CFactura
    
    'NuevasNormasBancarias
    If vParamAplic.TieneTelefonia2 = 3 Then
        'BOLBAITE
        'No pondra en los textos csb lo que pone alzira
        
    Else
        cadW = DevuelveDesdeBD(conConta, "Norma19_34Nueva", "paramtesor", "1", "1")
        If cadW = "1" Then vFactu.PonerValorNuevasNormasBancarias True
    End If
    
    
    cadW = ""
    Fichero = TipoAlb

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", CStr(banPr), "N")
    b = True
    
    For J = 1 To 2
    
        If J = 1 Then
            SQL = "Select * from scaalb WHERE codtipom = 'ALI' and referenc='" & TipoAlb & "' ORDER BY numalbar"
            TipoAlb = ""
        Else
            SQL = "Select * from scaalb WHERE codtipom = 'ALT' "
            'Si tiene nombre de fichero es que no viene ficheros movistar, vodafone. Viene de COARVAL
            If Not Coarval Then SQL = SQL & " and referenc='" & Fichero & "'"
            SQL = SQL & " ORDER BY numalbar"
        End If
        Set RSalb = New ADODB.Recordset
        RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set RTT = New ADODB.Recordset
        If Not Coarval Then
            If J > 1 Then
                'NO internas. ALT
                SQL = "Select numalbar,codclien,coddirec,referenc from scaalb WHERE codtipom = 'ALT' AND factursn=1 and referenc<>'" & Fichero & "' ORDER BY codclien,coddirec"
                RTT.Open SQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
            End If
        End If
        
        While Not RSalb.EOF
                TipoAlb = RSalb!codtipom
    
                
                vFactu.Cliente = RSalb!codClien
                vFactu.NombreClien = Trim(RSalb!Nomclien)
                LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & Mid(vFactu.NombreClien, 1, 25)
                LblBar.Refresh
                
                Espera 0.15
                If Inc = 5 Then
                    DoEvents
                    Inc = 0
                End If
                
                'Los Albaranes que tengan tipofact=1 "factura x Albaran" generar una factura
                'para cada uno de ellos
                If J = 1 Then    'INTERNOS
                    SQL = " scaalb.codtipom='" & RSalb!codtipom & "' AND scaalb.numalbar=" & RSalb!NumAlbar
                Else
                    'FAT
                    SQL = RSalb!NumAlbar 'en la funcion de pasaralbfras YA montar? el SELECT 'FUERZA EL NUMERO
                    
                    If Not Coarval Then
                        'PUEDE ser que tengamos albaranes de telefonia introducidos "a mano"
                        'Habra que verlos
                        If Not RTT.EOF Then
                            While Not RTT.EOF
                                If RTT!codClien = RSalb!codClien Then
                                    If DBLet(RSalb!CodDirec, "T") = DBLet(RTT!CodDirec, "T") Then
                                    
                                        'Habra que ver si el telefono es el mismo que facturamos
                                        'en esta factura
                                        If Trim(RSalb!observa04) = Trim(RTT!referenc) Then
                                        
                                            'OK. Este tenemos que facturarlo aqui
                                        
                                            SQL = SQL & ", " & RTT!NumAlbar
                                        End If
                                    End If
                                End If
                                RTT.MoveNext
                            Wend
                            RTT.MoveFirst
                        End If
                    End If
                    
                    
                    
                End If
                
                
                    
                
                
                'Generar una Factura nueva
                vFactu.FecFactu = RSalb!FechaAlb
                vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
                vFactu.CPostal = DBLet(RSalb!codpobla, "T")
                vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
                vFactu.Provincia = DBLet(RSalb!proclien, "T")
                vFactu.NIF = DBLet(RSalb!nifClien, "T")
                vFactu.Telefono = DBLet(RSalb!telclien, "T")
                vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
                vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
                vFactu.Agente = RSalb!CodAgent
                vFactu.ForPago = RSalb!codforpa
                vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
                vFactu.DtoPPago = CCur(RSalb!DtoPPago)
                vFactu.DtoGnral = CCur(RSalb!DtoGnral)
    
                    
                    
                If Not vFactu.PasarAlbaranesAFactura(TipoAlb, SQL, "", ErroresAux, False) Then
                    b = False
                    AnyadirAvisos ErroresAux
                Else
                    'ACTUALIZAMOS LA FACTURA
                    If J = 1 Then
                        'LAS INTERNAS
                        cadW = "serie='" & vFactu.LetraSerie & "' and ano=" & Year(vFactu.FecFactu) & " and numfact=" & RSalb!NumAlbar
                        ActualizaEnCabFacTel Fichero, cadW, vFactu.NumFactu
                        
                        'Marzo 2014. Hay que actualizar numofert y serietfno de scafac1
                        cadW = "UPDATE scafac1 SET numofert=" & vFactu.NumFactu & ", serietfno='" & vFactu.LetraSerie & "'"
                        cadW = cadW & " WHERE codtipom='FAI' AND numfactu=" & vFactu.NumFactu & " AND fecfactu=" & DBSet(vFactu.FecFactu, "F")
                        ejecutar cadW, False
                        
                    End If
                End If
                
                
            
                Espera 0.1
                Inc = Inc + 1
                cadW = ""
                
           
            RSalb.MoveNext
        Wend
        RSalb.Close
        Set RSalb = Nothing
        
        If Not Coarval Then
            If J > 1 Then RTT.Close
        End If
        Set RTT = Nothing
    Next J
    
    Set vFactu = Nothing
    GenerarFacturasTelefonia = True
    
    If b Then
    
    
        'MARZO 2014
        'PARA las itnernas de telefonia hay que poner en scafac1
        'en serietfno la letra de serie de las fras internas
        cadW = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FAI", "T")
        cadW = "UPDATE scafac1 SET serietfno='" & cadW & "' WHERE"
        cadW = cadW & " codtipom='FAI' AND referenc=" & DBSet(Fichero, "T")
        If Not ejecutar(cadW, False) Then MsgBox "Error actualizando referencias INTERNAS(serietfno). El proceso continua. Avise soporte t?cnico", vbExclamation
            
        cadW = "codtipom = 'ALT' AND factursn"
        cadW = DevuelveDesdeBD(conAri, "count(*)", "scaalb", cadW, "1")
        If Val(cadW) > 0 Then
            cadW = vbCrLf & vbCrLf & "Esisten albaranes de telefonia pendientes de facturar  y estan marcados"
        Else
            cadW = ""
        End If
        LblBar.Caption = "Proceso finalizado correctamente."
        MsgBox "Las Facturas de los Albaranes seleccionados se generaron correctamente." & cadW, vbInformation
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        SQL = "ATENCI?N:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    Espera 0.2
    

    
    
    
ETraspasoAlbFac:
    
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando telefon?a", Err.Description
    End If
    Set RTT = Nothing
End Function

Private Sub ActualizaEnCabFacTel(Fichero As String, Where As String, Nuevo As Long)
Dim SQ As String
Dim J As Byte

    On Error GoTo eAct
    
    
    For J = 1 To 5
        SQ = RecuperaValor("tel_cab_factura|tel_lin_factura_consumos|tel_lin_factura_cuotas|tel_lin_factura_descuentos|tel_lin_factura_especial|", CInt(J))
        
        SQ = "UPDATE " & SQ & " SET NumFact =" & Nuevo & " WHERE " & Where
        If J = 1 Then SQ = SQ & " AND fichero = '" & Fichero & "'"
        conn.Execute SQ
    Next J
    
    Exit Sub
eAct:
    Err.Clear
    '-----------------
    'tmpcrmmsg(codusu,codigo,tipo,fechahora,asun_obs)
    SQ = DevNombreSQL(SQ)
    'SQ = "INSERT INTO tmpcrmmsg(codusu,codigo,tipo,fechahora,asun_obs)"
    SQ = vUsu.codigo & "," & Nuevo & ",0,now(),'" & SQ & "')"
    SQ = "INSERT INTO tmpcrmmsg(codusu,codigo,tipo,fechahora,asun_obs) VALUES (" & SQ
    ejecutar SQ, True
    
End Sub

Private Function GenerarAlbaranesTelefonia(ByRef L As Label) As Boolean
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim vCli As CCliente
Dim Tr As Integer
Dim Internas As String
Dim PeriodoFacturacion As String
Dim F1 As Date
Dim CodCCost As String

    On Error GoTo eGenerarAlbaranesTelefonia

    GenerarAlbaranesTelefonia = False
    Set Rs = New ADODB.Recordset

    L.Caption = "Calculando lineas"
    L.Refresh
    
    
    'El periodo de facturaqcion ira a la observacion 5, el numero de telefono en la 4
    'TipoFac: Tienen el noombre del fichero
    Cad = DevuelveDesdeBD(conAri, "max(Fecha_final_periodo)", "telefono.resumen_de_llamadas", "Fichero", TipoFac, "T")
    If Cad = "" Then
        MsgBox "Error obteniendo el periodo de facturacion. Proceso continua", vbExclamation
        Cad = Format(Now, "dd/mm/yyyy")
    End If
    F1 = CDate(Cad)
    PeriodoFacturacion = Format(Day(F1), "00") & " de " & Format(F1, "mmmm") & " de " & Format(Year(F1), "0000")
    F1 = DateAdd("m", -1, F1)
    F1 = DateAdd("d", 1, F1)
    PeriodoFacturacion = Format(Day(F1), "00") & " de " & Format(F1, "mmmm") & " de " & Format(Year(F1), "0000") & " a " & PeriodoFacturacion
    
    
    'Metemos en una tmp lo que luego sumado sera el valor de la linea
    'tmpstockfec(codusu,codartic,codalmac,stock)

    


    L.Caption = "Obteniendo registros"
    L.Refresh
    Cad = "Select * from  tel_cab_factura,sclientfno where IdTelefono=Telefono AND fichero= '" & TipoFac & "' ORDER BY numfact"
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Errores = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtiTelefonia, "T")
    
    
    
    
    cadW = "INSERT INTO scaalb(codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,"
    cadW = cadW & "proclien,nifclien,telclien,referenc,facturkm,codtraba,codtrab1,codtrab2,codagent,"
    cadW = cadW & "codforpa,codenvio,dtoppago,dtognral,tipofact,numpedcl,observa01,observa02,observa03,observa04,"
    cadW = cadW & "observa05,esticket,coddirec,nomdirec,numofert,sementre) VALUES "
    
    DoEvents
    Cad = PonerTrabajadorConectado("")
    Tr = Val(Cad)
    
    
    Internas = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FAI", "T")
    
    
    
    
    '  Centros de coste para los que lleven analitica
    '   Es del trabajador      o de la familia
    ' para ello en codccost llevaremos 2 valores. Para las ALT y las ALI
    CodCCost = "null|null|"
    If vEmpresa.TieneAnalitica Then
        ' 0=trabajador, 1=Familia
        If vParamAplic.ModoAnalitica = 0 Then
            Cad = DevuelveDesdeBD(conAri, "codccost", "straba", "codtraba", CStr(Tr))
            If Cad = "" Then Err.Raise 513, , "No se puede establecer centro de coste para el trabajador conectado(" & Tr & ")"
                
            Cad = DBSet(Cad, "T")
            CodCCost = Cad & "|" & Cad & "|"
        
        Else
            'select codccost from sartic,sfamia where sartic.codfamia=sfamia.codfamia
            Cad = "sartic.codfamia=sfamia.codfamia AND codartic"
            Cad = DevuelveDesdeBD(conAri, "codccost", "sartic,sfamia", Cad, vParamAplic.ArtiTelefonia, "T")
            If Cad = "" Then Err.Raise 513, , "No se puede establecer centro de coste para el articulo telefonia"
            Cad = DBSet(Cad, "T")
            CodCCost = Cad & "|"

            'INternas
            Cad = "sartic.codfamia=sfamia.codfamia AND codartic"
            Cad = DevuelveDesdeBD(conAri, "codccost", "sartic,sfamia", Cad, vParamAplic.ArtiTelefonia, "T")
            If Cad = "" Then Err.Raise 513, , "No se puede establecer centro de coste para el articulo telefonia exento"
            Cad = DBSet(Cad, "T")
            CodCCost = CodCCost & Cad & "|"
        End If
    End If
    
    Set vCli = New CCliente
    Cad = ""
    While Not Rs.EOF
        L.Caption = Rs!Telefono
        L.Refresh
        If Not vCli.LeerDatos(CStr(Rs!codClien)) Then Err.Raise 513, , "Error leyendo el cliente: " & Rs!codClien

        'YA tengo el cliente
        'Vamos p'alla
        '(codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,
        'proclien,nifclien,telclien,referenc,facturkm,codtraba,codtrab1,codtrab2,codagent,
        'codforpa,codenvio,dtoppago,dtognral,tipofact,observa01,observa02,observa03,observa04,observa04,esticket)
        If Rs!Serie = Internas Then
            Cad = "ALI"
        Else
            Cad = "ALT"
        End If
        Cad = "('" & Cad & "'," & Rs!NumFact & "," & DBSet(Rs!Fecha, "F") & ",1," & vCli.codigo & "," & DBSet(vCli.Nombre, "T")
        Cad = Cad & "," & DBSet(vCli.Domicilio, "T") & "," & DBSet(vCli.CPostal, "T")
        Cad = Cad & "," & DBSet(vCli.Poblacion, "T") & "," & DBSet(vCli.Provincia, "T")
        Cad = Cad & "," & DBSet(vCli.NIF, "T") & "," & DBSet(Rs!Telefono, "T") & ",'" & TipoFac & "',"
        
        Cad = Cad & "0," & Tr & "," & Tr & "," & Tr & "," & vCli.Agente
        Cad = Cad & "," & vCli.ForPago & "," & vParamAplic.PorDefecto_Envio & ",0,0,1"
        
        'FEBRERO 2014.
        'Para la reimpresion de facturas de telefonia.
        'Grabaremos en el campo numpedcl
        'un 1 si se imprime o un 0 si debe ir por email
        LetraSer = "0"
        If Rs!Factura = 0 Then LetraSer = "1"  'si sclientfno.factura=0 es que no quiere la factura-->email
        Cad = Cad & "," & LetraSer
        
        'En las observaciones podemos poner DATOS de la facturacion
        'Observa 01.  Nombre
        LetraSer = ""
        If Not IsNull(Rs!apellido1) Then LetraSer = Rs!apellido1
        If Not IsNull(Rs!apellido2) Then LetraSer = Trim(LetraSer & " " & Rs!apellido2)
        If Not IsNull(Rs!Nombre) Then
             If LetraSer <> "" Then LetraSer = LetraSer & ","
             LetraSer = Trim(LetraSer & " " & Rs!Nombre)
        End If
        Cad = Cad & "," & DBSet(LetraSer, "T")
        'Observa2
        LetraSer = Trim(DBLet(Rs!CodPostal, "T") & "  " & DBLet(Rs!Direccion, "T"))
        Cad = Cad & "," & DBSet(LetraSer, "T")
        'Obs3
        LetraSer = Trim(DBLet(Rs!Provincia, "T") & "  " & DBLet(Rs!Companyia, "T"))
        Cad = Cad & "," & DBSet(LetraSer, "T")
        'Octubre 2013
        'Observa3,4 y 5 ->> N? telefono y periodo facturacion
        'Cad = Cad & ",NULL,NULL,0,"
        Cad = Cad & "," & DBSet(Rs!idtelefono, "T") & "," & DBSet(PeriodoFacturacion, "T") & ",0,"
        
        
        
        'Abril 2013
        'coddirec, nommdirec
        If IsNull(Rs!CodDirec) Then
            Cad = Cad & "NULL,NULL"
        Else
            Cad = Cad & Rs!CodDirec & ",'"
            Cad = Cad & DevuelveDesdeBD(conAri, "nomdirec", "sdirec", "codclien = " & vCli.codigo & " AND coddirec  ", Rs!CodDirec) & "'"
        End If
        
        'Octrubre 2013
        Cad = Cad & "," & Rs!NumFact & "," & Rs!Ano
        Cad = Cad & ")"
        Cad = cadW & Cad
        conn.Execute Cad
        
        'La linea
        Cad = "INSERT INTO slialb (codtipom,numalbar,numlinea,codalmac,codartic,nomartic,"
        Cad = Cad & "ampliaci,cantidad,numbultos,precioar , dtoline1, dtoline2, ImporteL, origpre, codproveX,codccost) VALUES ("
        If Rs!Serie = Internas Then
            Cad = Cad & "'ALI'"
        Else
            Cad = Cad & "'ALT'"
        End If
        Cad = Cad & "," & Rs!NumFact & ",1,1," & DBSet(vParamAplic.ArtiTelefonia, "T") & ","
        Cad = Cad & DBSet(Errores, "T") & ",'" & TipoFac & "',1,0,"
        Cad = Cad & DBSet(Rs!BaseImponible, "N") & ",0,0," & DBSet(Rs!BaseImponible, "N") & ",'M',0,"
        'Septiembre 2014
        Cad = Cad & RecuperaValor(CodCCost, 1) & ")"
        conn.Execute Cad
        
        
        'MAYO 2013
        'Segunda linea
        'IVA exento. Que estara en un campo
        ' De momento solo esta para los ficheros VODAFONE
        If DBLet(Rs!base_exenta, "N") > 0 Then
            Cad = "INSERT INTO slialb (codtipom,numalbar,numlinea,codalmac,codartic,nomartic,"
            Cad = Cad & "ampliaci,cantidad,numbultos,precioar , dtoline1, dtoline2, ImporteL, origpre, codproveX,codccost) VALUES ("
            If Rs!Serie = Internas Then
                Cad = Cad & "'ALI'"
            Else
                Cad = Cad & "'ALT'"
            End If
            Cad = Cad & "," & Rs!NumFact & ",2,1," & DBSet(vParamAplic.ArtTfniaIvaExento, "T") & ","
            Cad = Cad & DBSet(Errores, "T") & ",'" & TipoFac & "',1,0,"
            Cad = Cad & DBSet(Rs!base_exenta, "N") & ",0,0," & DBSet(Rs!base_exenta, "N") & ",'M',0,"
            'Septiembre 2014
            Cad = Cad & RecuperaValor(CodCCost, 2) & ")"
            conn.Execute Cad
        End If
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
    
    
    If Cad <> "" Then GenerarAlbaranesTelefonia = True
    
    
    
eGenerarAlbaranesTelefonia:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set vCli = Nothing
    Set Rs = Nothing
End Function



Public Function traspasofacturasTelefoniaCOARVAL(ByRef L As Label, ByVal idBanco As Integer) As Boolean
Dim b As Boolean
    
    traspasofacturasTelefoniaCOARVAL = False
    
    'Bloqueamos proceso
    If Not BloqueoManual("VENFAC", "1") Then
        'MsgBox "No se puede facturar TELEFONIA. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If



    'Proceso 1. MEter en slialb,scaalb
    'EL NUmero de albaran, sera el mismo que el numero de factura
    b = GenerarAlbaranesTelefoniaCOARVAL(L)
    cadW = ""  'reutiolizadas
    Errores = "" 'reutilzadas
    TipoFac = ""
    LetraSer = ""
    
    'Proceso 2.  FACTURAR
    If b Then
        DoEvents
        
        'MEtemos el nombre del fichero en los datos traspasados
'        cadW = DevuelveDesdeBD(conAri, "count(*)", "scaalb", "codtipom", "ALT", "T")
'        If cadW = "" Then cadW = "0"
'
'        Errores = "referenc='" & Fichero & "' AND codtipom"
'        Errores = DevuelveDesdeBD(conAri, "count(*)", "scaalb", Errores, "ALI", "T")
'        If Errores = "" Then Errores = "0"
'
'        cadW = " VALUES ('" & Fichero & "'," & DBSet(Now, "F") & "," & cadW & "," & Errores & ")"
'        cadW = "INSERT INTO tel_fichtraspasados(Fichero,Fecha,FraNormal,FraInt)" & cadW
'        conn.Execute cadW  'Fichero YA procesado
'
        'PARA LAS INTERNAS
        '
        'Generaremos la factura en scafac y con el numero que obtenemos
        'updatearemos tel_cabfactura....
        'Por si acaso diera algun fallo y no se renumerara(no deberia pasar)
        'lo guardamos en tmpstockfec para avisar
        conn.Execute "DELETE FROM tmpcrmmsg where codusu = " & vUsu.codigo
        
        
        'Generaremos las ALT, las normales
        'y 'las ALI que sean de telefonia
        b = GenerarFacturasTelefonia(idBanco, L, True, True)
        If b Then traspasofacturasTelefoniaCOARVAL = True
    End If
    
    
    If Not b Then
        
        Errores = DevuelveDesdeBD(conAri, "count(*)", "scaalb", "codtipom", "ALT", "T")
        If Errores <> "" Then MsgBox "Se han quedado " & Errores & " albaranes. Consulte soporte t?cnico", vbExclamation
            
    End If
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    

    
End Function



Private Function GenerarAlbaranesTelefoniaCOARVAL(ByRef L As Label) As Boolean
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim vCli As CCliente
Dim Tr As Integer
Dim Internas As String

    On Error GoTo eGenerarAlbaranesTelefonia2

    GenerarAlbaranesTelefoniaCOARVAL = False
    Set Rs = New ADODB.Recordset

    L.Caption = "Calculando lineas"
    L.Refresh
    
    
    'Metemos en una tmp lo que luego sumado sera el valor de la linea
    'tmpstockfec(codusu,codartic,codalmac,stock)




    L.Caption = "Obteniendo registros"
    L.Refresh
    Cad = "Select * from tmpinformes WHERE codusu = " & vUsu.codigo & " ORDER BY codigo1"
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Errores = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArtiTelefonia, "T")
    
    
    
    
    cadW = "INSERT INTO scaalb(codtipom,numalbar,fechaalb,factursn,codclien,nomclien,domclien,codpobla,pobclien,"
    cadW = cadW & "proclien,nifclien,telclien,referenc,facturkm,codtraba,codtrab1,codtrab2,codagent,"
    cadW = cadW & "codforpa,codenvio,dtoppago,dtognral,tipofact,observa01,observa02,observa03,observa04,"
    cadW = cadW & "observa05,esticket,coddirec,nomdirec) VALUES "
    
    DoEvents
    Cad = PonerTrabajadorConectado("")
    Tr = Val(Cad)
    
    
    Internas = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", "FAI", "T")
    
    
    Set vCli = New CCliente
    Cad = ""
    While Not Rs.EOF
        L.Caption = DBLet(Rs!nombre1, "T")
        L.Refresh
        If Not vCli.LeerDatos(CStr(Rs!campo1)) Then Err.Raise 513, , "Error leyendo el cliente: " & Rs!campo1

        
        'If Rs!serie = Internas Then
        '    cad = "ALI"
        'Else
            Cad = "ALT"
        'End If
        Cad = "('" & Cad & "'," & Rs!Codigo1 & "," & DBSet(Rs!fecha1, "F") & ",1," & vCli.codigo & "," & DBSet(vCli.Nombre, "T")
        Cad = Cad & "," & DBSet(vCli.Domicilio, "T") & "," & DBSet(vCli.CPostal, "T")
        Cad = Cad & "," & DBSet(vCli.Poblacion, "T") & "," & DBSet(vCli.Provincia, "T")
        Cad = Cad & "," & DBSet(vCli.NIF, "T") & "," & DBSet(vCli.TfnoClien, "T") & ",'" & TipoFac & "',"
        
        Cad = Cad & "0," & Tr & "," & Tr & "," & Tr & "," & vCli.Agente
        Cad = Cad & "," & vCli.ForPago & "," & vParamAplic.PorDefecto_Envio & ",0,0,1"
        
        'Observa 01.  Nombre
        LetraSer = ""
        Cad = Cad & "," & DBSet(LetraSer, "T")
        'Observa2
        LetraSer = ""
        Cad = Cad & "," & DBSet(LetraSer, "T")
        'Obs3
        LetraSer = ""
        Cad = Cad & "," & DBSet(LetraSer, "T")
        'Observa3,4 y 5
        Cad = Cad & ",NULL,NULL,0,"
        
        
        'coddirec, nommdirec
        'If IsNull(Rs!CodDirec) Then
        If True Then
            Cad = Cad & "NULL,NULL"
        Else
            'cad = cad & Rs!CodDirec & ",'"
            'cad = cad & DevuelveDesdeBD(conAri, "nomdirec", "sdirec", "codclien = " & vCli.Codigo & " AND coddirec  ", Rs!CodDirec) & "'"
        End If
        Cad = Cad & ")"
        Cad = cadW & Cad
        conn.Execute Cad
        
        'La linea
        Cad = "INSERT INTO slialb (codtipom,numalbar,numlinea,codalmac,codartic,nomartic,"
        Cad = Cad & "ampliaci,cantidad,numbultos,precioar , dtoline1, dtoline2, ImporteL, origpre, codproveX) VALUES ("
        Cad = Cad & "'ALT'," & Rs!Codigo1 & ",1,1," & DBSet(vParamAplic.ArtiTelefonia, "T") & ","
        Cad = Cad & DBSet(Errores, "T") & ",NULL,1,0,"
        Cad = Cad & DBSet(Rs!Importe1, "N") & ",0,0," & DBSet(Rs!Importe1, "N") & ",'M',0)"
        conn.Execute Cad
        
        

        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
    If Cad <> "" Then GenerarAlbaranesTelefoniaCOARVAL = True
    
    
    
eGenerarAlbaranesTelefonia2:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set vCli = Nothing
    Set Rs = Nothing
End Function








'********************************************************************************
'********************************************************************************
'********************************************************************************
'
'   CONTADORES AGUA
'
'********************************************************************************
'********************************************************************************
'********************************************************************************
'
'   Coge todos los albaranes ALG que hayan y los factura
Public Function FacturarContadoresAgua(Fecfact As String, banPr As String, ByRef LblBar As Label, CentroCoste As String) As Boolean

Dim RSalb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim SQL As String
Dim Aux As String
Dim vFactu As CFactura
Dim Inc As Integer
Dim DatosOk_ As Boolean
Dim RTT As ADODB.Recordset


    
    
    
    
    Errores = ""
    Inc = 0
    
    
    Set vFactu = New CFactura
    TipoAlb = "ALG"
    

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", CStr(banPr), "N")
    b = True
    
    
    
    SQL = "Select * from scaalb WHERE codtipom = " & DBSet(TipoAlb, "T")
    SQL = SQL & " ORDER BY numalbar"

    Set RSalb = New ADODB.Recordset
    RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Set RTT = New ADODB.Recordset
    
    While Not RSalb.EOF
            

            
            vFactu.Cliente = RSalb!codClien
            vFactu.NombreClien = Trim(RSalb!Nomclien)
            LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & Mid(vFactu.NombreClien, 1, 25)
            LblBar.Refresh
            
            Espera 0.15
            If Inc = 5 Then
                DoEvents
                Inc = 0
            End If
            
            
            
            
            
            'Generar una Factura nueva
            vFactu.FecFactu = RSalb!FechaAlb
            vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
            vFactu.CPostal = DBLet(RSalb!codpobla, "T")
            vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
            vFactu.Provincia = DBLet(RSalb!proclien, "T")
            vFactu.NIF = DBLet(RSalb!nifClien, "T")
            vFactu.Telefono = DBLet(RSalb!telclien, "T")
            vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
            vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
            vFactu.Agente = RSalb!CodAgent
            vFactu.ForPago = RSalb!codforpa
            vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
            vFactu.DtoPPago = CCur(RSalb!DtoPPago)
            vFactu.DtoGnral = CCur(RSalb!DtoGnral)
                
            'En el agua, busco aqui la cuenta banco, si es recibo
            DatosOk_ = True
            If vFactu.TipForPago = 4 Then
                Aux = "Select iban,codbanco,codsucur,digcontr,cuentaba from aguacontadores where contador =" & DBSet(RSalb!referenc, "T")
                'No puede ser eof
                RTT.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RTT.EOF Then
                   AnyadirAvisos "Error leyendo datos contador: " & DBSet(RSalb!referenc, "T")
                   DatosOk_ = False
                Else
                    'Si tiene valor
                    If DBSet(RTT!IBAN, "T") <> "" Then
                        'OK, de aqui va a coger el banco
                        vFactu.Banco = DBLet(RTT!codbanco, "T")
                        vFactu.Sucursal = DBLet(RTT!codsucur, "T")
                        vFactu.DigControl = DBLet(RTT!digcontr, "T")
                        vFactu.CuentaBan = DBLet(RTT!cuentaba, "T")
                        vFactu.IBAN = DBLet(RTT!IBAN, "T")
                    Else
                        RTT.Close
                        Aux = "Select iban,codbanco,codsucur,digcontr,cuentaba from sclien where codclien =" & vFactu.Cliente
                        RTT.Open Aux, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                        vFactu.Banco = DBLet(RTT!codbanco, "T")
                        vFactu.Sucursal = DBLet(RTT!codsucur, "T")
                        vFactu.DigControl = DBLet(RTT!digcontr, "T")
                        vFactu.CuentaBan = DBLet(RTT!cuentaba, "T")
                        vFactu.IBAN = DBLet(RTT!IBAN, "T")
                    End If
                End If
                RTT.Close
            End If
            
            If DatosOk_ Then
                SQL = " scaalb.codtipom='" & RSalb!codtipom & "' AND scaalb.numalbar=" & RSalb!NumAlbar
                If Not vFactu.PasarAlbaranesAFactura(TipoAlb, SQL, "", ErroresAux, False) Then
                    b = False
                    AnyadirAvisos ErroresAux
                End If
            
            End If
        
            Espera 0.1
            Inc = Inc + 1
            cadW = ""
            
       
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
        
        
    
    
    Set vFactu = Nothing
    FacturarContadoresAgua = True
    
    If b Then
        LblBar.Caption = "Proceso finalizado correctamente."
        MsgBox "Las Facturas de los Albaranes seleccionados se generaron correctamente." & cadW, vbInformation
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        SQL = "ATENCI?N:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    Espera 0.2
    

    
    
    
ETraspasoAlbFac:
    
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando telefon?a", Err.Description
    End If
    Set RTT = Nothing
End Function







'*****************************************************************************************************
' SAIL // EULER
' Podriamos extrapolarlo a todos, ya que lo que hace es una serie de albranes
'  por (tipo, numero), (tipo,numero) (tipo,numero.....
' Generaremos un tipo de factura final
Public Function TraspasoAlbaranesFacturasCliente(cadSQL As String, cadWhere As String, FechaFact As String, banPr As String, ByRef PBar1 As ProgressBar, ByRef LblBar As Label, ImprimeLasFacturasGeneradas As Boolean, ByRef TipoDeFactura As String, TextosCSB As String, NumeroCopias As Byte, MostrarMsgOK As Boolean) As Boolean

'Desde Albaranes Genera las Facturas correspondientes
Dim RSalb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim b As Boolean
Dim SQL As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim antClien As Long
Dim antDirec As Long
Dim antForpa As Integer
Dim antDtoPP As Single, antDtoGn As Single

'direc/dpto actual para controlar el valor nulo
Dim actDirec As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vFactu As CFactura
Dim Inc As Integer
Dim condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

'Por si no mando una progressbar, que no de errores
Dim PgbVisible As Boolean


Dim HazPulsarAceptarEnFrmImprimir As Boolean

    On Error GoTo ETraspasoAlbFac

    TraspasoAlbaranesFacturasCliente = False

    ListFactu = ""
        
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("VENFAC") 'facturas de venta
    If Not BloqueoManual("VENFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los albaranes que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    SQL = " (scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien ) INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
    If Not BloqueaRegistro(SQL, cadWhere) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        Exit Function
    End If
    
   
    'Inicializar la Progress Bar
    PgbVisible = False
    If Not (PBar1 Is Nothing) Then
        If PBar1.visible Then PgbVisible = True
    End If
    If PgbVisible Then
        If InStr(1, cadSQL, "sclien") Then
            SQL = Replace(cadSQL, "scaalb.*, sclien.periodof", "count(*)") 'si hay INNER JOIN con sclien
        Else
            SQL = Replace(cadSQL, "*", "count(*)") 'si NO hay INNER JOIN con sclien
        End If
        
        
        Set RSalb = New ADODB.Recordset
        RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RSalb.EOF Then
            CargarProgresNew PBar1, CInt(RSalb.Fields(0))
            LblBar.Caption = "Inicializando el proceso..."
            LblBar.Refresh
            
        End If
        RSalb.Close
        Set RSalb = Nothing
    End If
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactu = New CFactura
    vFactu.FecFactu = FechaFact 'Fecha para las Facturas

    'Cuenta Prevista de Cobro de las Facturas
    vFactu.BancoPr = banPr
    vFactu.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", banPr, "N")
    
    'comprobar que la cuenta prevista de cobro tiene valor
    b = (vFactu.CuentaPrev <> "")
    If Not b Then
        Set vFactu = Nothing
        'Desbloqueamos ya no estamos facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
        MsgBox "La cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Function
    End If
    
       
        
    'Marcar Albaranes que se van a Facturar
    '----------------------------------------
    SQL = cadSQL & " ORDER BY scaalb.tipofact, scaalb.codclien, scaalb.coddirec, codforpa, dtoppago, dtognral "
    Set RSalb = New ADODB.Recordset
    RSalb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    'Agrupar los Albaranes posibles en una misma Factura
    'Calcular y Grabar Factura en la Tabla de Facturas
    'Albaran(scaalb, slialb) -> Factura (scafac,scafac1,slifac)
    '----------------------------------------------------
    'Se factura por cliente y departamento
    'Agrupar albaranes en 1 factura por : tipofact,codclien,coddirec,codforpa,dtoppago, dtognral
    antClien = 0 'cliente
    antDirec = 0 'direccion/departamento
    antForpa = 0 'forma de pago
    antDtoPP = 0 'dto pronto pago
    antDtoGn = 0 'dto general
    
    cadW = ""
    Errores = ""
    Inc = 0
    
    While Not RSalb.EOF
        TipoAlb = TipoDeFactura     'RSalb!codtipom   siempre el tipo final
        Inc = Inc + 1
        If IsNull(RSalb!CodDirec) Then
            actDirec = -1
        Else
            actDirec = DBLet(RSalb!CodDirec, "N")
        End If
        
        If RSalb!TipoFact = 1 Then 'tipofact=1 "FACTURA x ALBARAN"
        '---------------------------------------------------------
            'frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas individuales"
            LblBar.Caption = "Facturando: Facturas individuales"
            LblBar.Refresh
            If cadW <> "" Then 'Facturacion pendiente
                cadW = cadW & ") "
                If Not vFactu.PasarAlbaranesAFactura(TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                    If b Then b = False
                    AnyadirAvisos ErroresAux
                Else 'a?adirlo a la lista de facturas a imprimir
                                   
                    ListFactu = ListFactu & "," & vFactu.NumFactu
                End If
                If PgbVisible Then
                    IncrementarProgresNew PBar1, Inc - 1
                    LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                    LblBar.Refresh
                End If
                Espera 0.2
                'Empezamos una nueva Factura
                cadW = ""
            End If
            
            'Los Albaranes que tengan tipofact=1 "factura x Albaran" generar una factura
            'para cada uno de ellos
            cadW = " scaalb.codtipom='" & RSalb!codtipom & "' AND scaalb.numalbar=" & RSalb!NumAlbar
            
            'Generar una Factura nueva
            vFactu.Cliente = RSalb!codClien
            vFactu.NombreClien = RSalb!Nomclien
            vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
            vFactu.CPostal = DBLet(RSalb!codpobla, "T")
            vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
            vFactu.Provincia = DBLet(RSalb!proclien, "T")
            vFactu.NIF = DBLet(RSalb!nifClien, "T")
            vFactu.Telefono = DBLet(RSalb!telclien, "T")
            vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
            vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
            vFactu.Agente = RSalb!CodAgent
            vFactu.ForPago = RSalb!codforpa
            vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
            vFactu.DtoPPago = CCur(RSalb!DtoPPago)
            vFactu.DtoGnral = CCur(RSalb!DtoGnral)

                
                
            If Not vFactu.PasarAlbaranesAFactura(TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                If b Then b = False
                AnyadirAvisos ErroresAux
            Else 'a?adirlo a la lista de facturas a imprimir

                ListFactu = ListFactu & "," & vFactu.NumFactu
                
            End If
            If PgbVisible Then
                Inc = 1 '1 albaran x factura
                LblBar.Caption = "Cliente: " & Format(RSalb!codClien, "000000") & " - " & RSalb!Nomclien
                LblBar.Refresh
                IncrementarProgresNew PBar1, Inc
                Inc = 0
            End If
            Espera 0.2
                
            cadW = ""
            
        Else 'tipofac=0 "factura COLECTIVA"
        '----------------------------------------------------------
            'Seleccionar todos los Albaranes pertenecientes a un mismo Cliente,Departamento
            'Los que tengan tipofac=0 "factura colectiva" agruparlos en una misma factura
            'para la misma Forma de PAgo, mismo dtoppago y mismo dtognral
             
             '-- David.      Esta linea da error si no viene de frmlistadoped
             'frmListadoPed.lblProgess(0).Caption = "Facturando: Facturas colectivas"
             LblBar.Caption = "Facturando: Facturas colectivas"
             LblBar.Refresh
             '---- Laura: 06/10/2006
             'Comprobar si es Departamento o Direccion (segun paramatro)
             'DAVID 05/07/2010    Direccion Departamento Obra.  Agrupa <>direccion
             If vParamAplic.HayDeparNuevo > 0 Then
                'agrupar tb por departamento
                condicion = (antClien <> RSalb!codClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             Else
                condicion = (antClien <> RSalb!codClien) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral)
             End If
             
'             If (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral) Then
             If condicion Then
             '-----
                If cadW <> "" Then 'Facturacion PEndiente
                    cadW = cadW & ") "
                    If Not vFactu.PasarAlbaranesAFactura(TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
                        If b Then b = False
                        AnyadirAvisos ErroresAux
                    Else 'a?adirlo a la lista de facturas a imprimir
                        'If ListFactu = "" Then
                        '    ListFactu = vFactu.NumFactu
                        'Else
                            ListFactu = ListFactu & "," & vFactu.NumFactu
                        'End If
                    End If
                    If PgbVisible Then
                        LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien
                        LblBar.Refresh
                        IncrementarProgresNew PBar1, Inc
                        Inc = 0
                    End If
                    Espera 0.2
                    
                    'Empezamos una nueva Factura
                    cadW = ""
                End If
                'Generar una Factura nueva
                vFactu.Cliente = RSalb!codClien
                vFactu.NombreClien = RSalb!Nomclien
                vFactu.DomicilioClien = DBLet(RSalb!domclien, "T")
                vFactu.CPostal = DBLet(RSalb!codpobla, "T")
                vFactu.Poblacion = DBLet(RSalb!pobclien, "T")
                vFactu.Provincia = DBLet(RSalb!proclien, "T")
                vFactu.NIF = DBLet(RSalb!nifClien, "T")
                vFactu.Telefono = DBLet(RSalb!telclien, "T")
                vFactu.DirDpto = DBLet(RSalb!CodDirec, "T")
                vFactu.NombreDirDpto = DBLet(RSalb!nomdirec, "T")
                vFactu.Agente = RSalb!CodAgent
                vFactu.ForPago = RSalb!codforpa
                vFactu.TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", RSalb!codforpa, "N")
                vFactu.DtoPPago = CCur(RSalb!DtoPPago)
                vFactu.DtoGnral = CCur(RSalb!DtoGnral)
                vFactu.Aportacion = 0
                If RSalb!codtipom = "ALM" Then vFactu.Aportacion = DBLet(RSalb!Aportacion, "N")
                cadW = " (scaalb.codtipom,scaalb.numalbar) IN (('" & RSalb!codtipom & "'," & RSalb!NumAlbar & ")"
            Else
                cadW = cadW & ",  ('" & RSalb!codtipom & "'," & RSalb!NumAlbar & ")"
            End If
        
            'Guardamos datos del registro anterior
            antClien = RSalb!codClien
'            antDirec = DBLet(RSalb!CodDirec, "N")
            antDirec = actDirec
            antForpa = RSalb!codforpa
            antDtoPP = RSalb!DtoPPago
            antDtoGn = RSalb!DtoGnral
        End If
        RSalb.MoveNext
    Wend
    RSalb.Close
    Set RSalb = Nothing
        
    'Facturar la ultima Factura generada del blucle
    If cadW <> "" Then
        cadW = cadW & ")"
        If PgbVisible Then LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
        
        If Not vFactu.PasarAlbaranesAFactura(TipoAlb, cadW, TextosCSB, ErroresAux, False) Then
            If b Then b = False
            AnyadirAvisos "Error Facturando el Cliente: " & Format(vFactu.Cliente, "000000") & " " & vFactu.NombreClien & vbCrLf & ErroresAux
        Else 'a?adirlo a la lista de facturas a imprimir
            ListFactu = ListFactu & "," & vFactu.NumFactu
        End If
        If PgbVisible Then
'            LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
            IncrementarProgresNew PBar1, Inc
        End If
        Espera 0.2
    End If
    
    TipoFac = vFactu.codtipom
    Set vFactu = Nothing
    
    
    If b Then
        TraspasoAlbaranesFacturasCliente = True
        LblBar.Caption = "Proceso finalizado correctamente."
        If MostrarMsgOK Then MsgBox "Las Facturas de los Albaranes seleccionados se generaron correctamente.", vbInformation
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        SQL = "ATENCI?N:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    Espera 0.2
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("VENFAC")
    TerminaBloquear
    
    
    If ListFactu <> "" Then ListFactu = Mid(ListFactu, 2)
            
    
    If ImprimeLasFacturasGeneradas Then
        If ListFactu <> "" Then
            HazPulsarAceptarEnFrmImprimir = False
            If TipoDeFactura = "ALM" And vParamAplic.EntradaRapidaFacturasMostrador Then HazPulsarAceptarEnFrmImprimir = True
            
            ImprimirFacturas ListFactu, FechaFact, , DevuelveTipoDocumentoFactura(TipoDeFactura), NumeroCopias, False, HazPulsarAceptarEnFrmImprimir

        End If
    End If
    
    'Voy a imprimir la hoja con las observaciones de la facturacion
    'Es decir si el cliente tiene observaciones de facturacion las mostrara ahora
    If ListFactu <> "" Then InformeObservacionFacturacion_ ListFactu, FechaFact
    
    
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Albaranes", Err.Description
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("VENFAC")
        TerminaBloquear
    End If
End Function





Public Function ComprobarFitosAlbaranesFacturasCliente(cadSQL As String, cadWhere As String) As Boolean
Dim RN As ADODB.Recordset
Dim SQL As String
Dim Col As Collection
Dim ErroL As String

    ComprobarFitosAlbaranesFacturasCliente = False
    
    'No deberia haber llegado
    If Not vParamAplic.ManipuladorFitosanitarios2 Then
        ComprobarFitosAlbaranesFacturasCliente = True
        Exit Function
    End If
    
    
    
    SQL = "DELETE FROM tmpnseries WHERE codusu = " & vUsu.codigo
    conn.Execute SQL
    Espera 0.5
    ErroL = ""
    
    Set RN = New ADODB.Recordset
    
    'Vamos a ver todos los albaranes que vamos a facturar
    SQL = "insert into tmpnseries(codusu,codartic,numlinealb,nummante) "
    SQL = SQL & " select " & vUsu.codigo & ",codtipom,numalbar,'' from  scaalb,sclien where scaalb.codclien=sclien.codclien AND " & cadWhere
    conn.Execute SQL

    'Quitamos los que no llevan articulos fitosnaitarios
    SQL = "delete from tmpnseries where codusu=" & vUsu.codigo & " and (codartic,numlinealb)"
    SQL = SQL & " in (select codtipom,numalbar from slialb inner join sartic on slialb.codartic=sartic.codartic group by 1,2 having sum(if(numserie<>'',1,0))=0)"
    conn.Execute SQL

    'Veremos cuales de los albaranes NO esta identificado el manipulador
    SQL = "select * from scaalb where (codtipom,numalbar) in (select codartic,numlinealb from tmpnseries where codusu = " & vUsu.codigo & ") and coalesce(manipuladornumcarnet,'')='' ORDER BY codtipom,numalbar"
    RN.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        SQL = RN!codtipom & " " & Format(RN!NumAlbar, "000000") & "  " & RN!Nomclien & vbCrLf
        ErroL = ErroL & SQL
        RN.MoveNext
    Wend
    RN.Close

    If ErroL <> "" Then
        SQL = "Falta identificar carnet manipulador fitosanitarios" & vbCrLf & String(60, "=") & vbCrLf
        ErroL = SQL & ErroL
    End If
    
    'Vamos a ver que todos los articulos con fitosanitarios tiene asignado los numeros de lote
    
    'Priemro veremos la cantidad en los albaranes
    SQL = "select codtipom,numalbar,sum(cantidad) lacanti from slialb inner join sartic on slialb.codartic=sartic.codartic where numserie<>''"
    SQL = SQL & " and (codtipom,numalbar) in (select codartic,numlinealb from tmpnseries where codusu = " & vUsu.codigo & " ) group by 1,2"
    RN.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RN.EOF
        'numserie
        SQL = Format(RN!lacanti * 100, "000000000")  '9 posiciones mas el signo
        SQL = "UPDATE tmpnseries SET numserie='" & SQL & "' WHERE codusu =" & vUsu.codigo
        SQL = SQL & " AND codartic= '" & RN!codtipom & "' AND numlinealb = " & RN!NumAlbar
        conn.Execute SQL
    
        RN.MoveNext
    Wend
    RN.Close
    
    'Ahora  veremos la cantidad en los lotes
    SQL = "select codtipom,numalbar,sum(cantidad) lacanti from slialblotes WHERE "
    SQL = SQL & " (codtipom,numalbar) in (select codartic,numlinealb from tmpnseries where codusu = " & vUsu.codigo & " ) group by 1,2"
    RN.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RN.EOF
        '  nummante
        SQL = Format(RN!lacanti * 100, "000000000") '9 posiciones mas el signo
        SQL = "UPDATE tmpnseries SET nummante='" & SQL & "' WHERE codusu =" & vUsu.codigo
        SQL = SQL & " AND codartic= '" & RN!codtipom & "' AND numlinealb = " & RN!NumAlbar
        conn.Execute SQL
    
        RN.MoveNext
    Wend
    RN.Close
    
    
    'Ahora veremos la cantidad de distintos que hay
    SQL = " select * from tmpnseries where codusu = " & vUsu.codigo & " AND nummante<>numserie order by codartic,numlinealb "
    RN.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    While Not RN.EOF
    
        SQL = SQL & RN!codArtic & Format(RN!numlinealb, "0000000") & " -> "
    
        If RN!nummante = "" Then
            SQL = SQL & " sin asignar lotes"
        Else
            SQL = SQL & " Lineas albaran: " & Format(Val(RN!numSerie / 100), FormatoCantidad)
            SQL = SQL & " //   Lotes  : " & Format(Val(RN!nummante / 100), FormatoCantidad)
            
        End If
        SQL = SQL & vbCrLf
        
        RN.MoveNext
    Wend
    RN.Close
    
    If SQL <> "" Then
        If ErroL <> "" Then ErroL = ErroL & vbCrLf & vbCrLf & vbCrLf
        ErroL = ErroL & "Lotes mal asignados: " & vbCrLf & String(60, "=") & vbCrLf
        ErroL = ErroL & SQL

    End If
    
    If ErroL <> "" Then
        Errores = ErroL
        LanzarErrorFitos
        Errores = ""
    Else
        'Todo bien
        ComprobarFitosAlbaranesFacturasCliente = True
    End If
    
    
    
    
    
    
    
End Function

Private Sub LanzarErrorFitos()
Dim NF As Integer
On Error GoTo eLanzarErrorFitos
    If Dir(App.Path & "\errfacFito.txt", vbArchive) <> "" Then Kill App.Path & "\errfacFito.txt"
    NF = FreeFile
    Open App.Path & "\errfacFito.txt" For Output As #NF
    Print #NF, Errores
    Close #NF
    
    Shell "notepad.exe " & App.Path & "\errfacFito.txt", vbNormalFocus
    
eLanzarErrorFitos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Sub

