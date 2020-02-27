Attribute VB_Name = "ModHistorico"
Option Explicit
'Modulo para el traspaso de registros de cabecera y lineas de las tablas
'de OFERTAS,PEDIDOS,ALBARANES
'A las tablas del HISTORICO de Ofertas,Pedidos,Albaranes
'OFERTAS:
' scapre --> schpre
' slipre --> slhpre
'PEDIDOS:
' scaped --> schped
' sliped --> slhped


Dim CodTipoMov As String
Dim NomTabla As String 'nombre de la tabla
Dim NomTablaH As String 'nombre de la tabla del historico al que movemos
Dim NomTablaLin As String 'nombre tabla de lineas
Dim NomTablaLinH As String 'nombre tabla de lineas del historico


Public Function ActualizarElTraspaso(ByRef ADonde As String, cadWhere As String, codMovim As String, Optional cadL As String) As Boolean
'codMovim: tipo de movimiento que estamos hacienco: OFE,PEV,ALV,PEC,ALC,....
    
    ActualizarElTraspaso = False
    CodTipoMov = codMovim
    
    'Insertamos en cabeceras Historico
    ADonde = "Insertando datos en histórico cabeceras "
    If Not InsertarCabeceraHistorico(cadWhere, cadL) Then Exit Function
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Histórico lineas "
    If Not InsertarLineasHistorico(cadWhere) Then Exit Function
    
    'Borramos cabeceras y lineas
    ADonde = "Borrar cabeceras y lineas"
    If Not BorrarTraspaso(False, cadWhere) Then Exit Function

    ActualizarElTraspaso = True
End Function

Public Function ActualizarElTraspasoSinBorrar(ByRef ADonde As String, cadWhere As String, codMovim As String, Optional cadL As String) As Boolean
    ActualizarElTraspasoSinBorrar = False
    CodTipoMov = codMovim
    
    'Insertamos en cabeceras Historico
    ADonde = "Insertando datos en histórico cabeceras "
    If Not InsertarCabeceraHistorico(cadWhere, cadL) Then Exit Function
'    IncrementarProgres 2
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Histórico lineas "
    If Not InsertarLineasHistorico(cadWhere) Then Exit Function


    ActualizarElTraspasoSinBorrar = True
End Function



Private Function InsertarCabeceraHistorico(cadWhere As String, Optional cadeN As String) As Boolean
Dim SQL As String
Dim Aux As String
Dim SegundaTablaCabeceras As Boolean
Dim CadenaInsercicon As String

On Error Resume Next

    

    NomTablaLinH = ""
    SegundaTablaCabeceras = False
    CadenaInsercicon = ""
    Select Case CodTipoMov
      Case "PEV" 'pedidos de venta a clientes
        NomTabla = "scaped"
        NomTablaH = "schped"
        NomTablaLinH = "slhped"
        SQL = " SELECT numpedcl,fecpedcl," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & ","
        SQL = SQL & "fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        SQL = SQL & "coddirec,nomdirec,referenc,codtraba,codagent,codforpa,dtoppago,dtognral,"
        SQL = SQL & "tipofact,observa01,observa02,observa03,observa04,observa05,servcomp,restoped,numofert,fecofert,observap1,observap2,recogecl,actuacion,coddiren,observacrm"
                   'Enero 2016        Nov 16
        SQL = SQL & ", PideCliente,observaciones,cerrado"
        
        
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ALI", "ALT", "ALO", "ALE" '[1.3.1] 'Albaran de venta a clientes
        NomTabla = "scaalb"
        NomTablaH = "schalb"
        NomTablaLinH = "slhalb"
        SQL = " SELECT codtipom,numalbar,fechaalb," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & ","
        SQL = SQL & "factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        SQL = SQL & "coddirec,nomdirec,referenc,facturkm,cantidkm,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,"
        SQL = SQL & "tipofact,observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa "
        SQL = SQL & ",aportacion, pesoalba, portes, fecenvio, docarchiv"
        SQL = SQL & ",tipliquid, actuacion"
        SQL = SQL & ",tipoimp,origdat"
        SQL = SQL & ",coddiren,tipAlbaran"
        SQL = SQL & ",albImpreso , codzonas,observacrm"
        'Ocvubre 2015
        SQL = SQL & ", ManipuladorNumCarnet , ManipuladorFecCaducidad , ManipuladorNombre,TipoCarnet"
        'Enero 2016               abri16      ago17
        SQL = SQL & ", PideCliente,numbultos,fechaAux,puntos"
        'NOV 2018
        SQL = SQL & ", codinter,codnatura,chofer,notasportes"
        'JUN 19-DIC19
        SQL = SQL & ",  FechaEnt , perrecep ,  latitud, Longitud ,dnirecep"
            
        CadenaInsercicon = " codtipom,numalbar,fechaalb,codigusu,fechelim,trabelim,codincid,"
        CadenaInsercicon = CadenaInsercicon & "factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        CadenaInsercicon = CadenaInsercicon & "coddirec,nomdirec,referenc,facturkm,cantidkm,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,"
        CadenaInsercicon = CadenaInsercicon & "tipofact,observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa "
        CadenaInsercicon = CadenaInsercicon & ",aportacion, pesoalba, portes, fecenvio, docarchiv"
        CadenaInsercicon = CadenaInsercicon & ",tipliquid, actuacion,tipoimp,origdat,coddiren,tipAlbaran,albImpreso , codzonas,observacrm"
        'Ocvubre 2015
        CadenaInsercicon = CadenaInsercicon & ", ManipuladorNumCarnet , ManipuladorFecCaducidad , ManipuladorNombre,TipoCarnet"
        'Enero 2016               abri16      ago17
        CadenaInsercicon = CadenaInsercicon & ", PideCliente,numbultos,fechaAux,puntos"
        'NOV 2018
        CadenaInsercicon = CadenaInsercicon & ", codinter,codnatura,chofer,notasportes"
        'JUN 19-DIC19
        CadenaInsercicon = CadenaInsercicon & ",  FechaEnt , perrecep ,  latitud, Longitud ,dnirecep"
            
        CadenaInsercicon = "(" & CadenaInsercicon & ")  "
            
            
        If InstalacionEsEulerTaxco Then SegundaTablaCabeceras = True
            
      Case "OFE" 'Ofertas a Clientes
        NomTabla = "scapre"
        NomTablaH = "schpre"
        NomTablaLinH = "slhpre"
        SQL = " SELECT numofert, fecofert," & "'" & Format(Now, FormatoFecha) & "' as fechamov, fecentre, aceptado, codclien, nomclien, domclien, codpobla, "
        SQL = SQL & "pobclien, proclien, nifclien, telclien, coddirec, nomdirec, referenc, codtraba, codagent, codforpa, dtoppago, dtognral, tipofact, "
        SQL = SQL & "plazos01, plazos02, plazos03, asunto01, asunto02, asunto03, asunto04, asunto05, observa01, observa02, observa03, observa04, observa05, "
        SQL = SQL & "concepto, seguiofe ,actuacion,coddiren,mailconfir,observacrm,obscompra," & cadeN & " as motivoTraspaso "
        
      Case "ALC" 'Albaranes a Proveedores (Compras)
        NomTabla = "scaalp"
        NomTablaH = "schalp"
        NomTablaLinH = "slhalp"
        SQL = " (numalbar,fechaalb,codprove,codigusu,fechelim,trabelim,codincid,nomprove,domprove,"
        SQL = SQL & "codpobla,pobprove,proprove,nifprove,telprove,codforpa,codtraba,codtrab1,dtoppago,dtognral,"
        SQL = SQL & "observa1,observa2,observa3,observa4,observa5,numpedpr,fecpedpr,fecenvio,docarchiv,codenvio,NReferencia,SReferencia,fecentrega) "
        SQL = SQL & " SELECT numalbar,fechaalb,codprove," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & ","
        SQL = SQL & "nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,"
        SQL = SQL & "codforpa,codtraba,codtrab1,dtoppago,dtognral,"
        SQL = SQL & "observa1,observa2,observa3,observa4,observa5,numpedpr,fecpedpr,fecenvio,docarchiv,codenvio,NReferencia,SReferencia,fecentrega"
      
      Case "PEC" 'Pedidos a Proveedores (Compras)
        NomTabla = "scappr"
        NomTablaH = "schppr"
        NomTablaLinH = "slhppr"
        SQL = " SELECT numpedpr,fecpedpr," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & ","
        SQL = SQL & "codprove,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,"
        SQL = SQL & "coddirea,coddiref,codforpa,codtraba,codtrab1,dtognral,dtoppago,"
        SQL = SQL & "restoped,codclien,observa1,observa2,observa3,observa4,observa5,tipoporte,obra,coddirre"
        SQL = SQL & ",NReferencia , SReferencia, CodEnvio, fecentrega"
    End Select
    
    'Borramos en hco primero
    'Primero las cabceras
    If CodTipoMov = "OFE" Then
        'Si es oferta existe el rigeso que la eliminacion es el trasapso a HCO con loc cual los aprametros son muchos
        'Por eso vamos a montar un select de eliminar
        Aux = "DELETE " & NomTablaLinH & ".* FROM " & NomTablaLinH & "  ," & NomTablaH & " WHERE "
        Aux = Aux & NomTablaLinH & ".numofert = " & NomTablaH & ".numofert AND " & Replace(cadWhere, NomTabla, NomTablaH)
        
    
    
    Else
        Aux = Replace(cadWhere, NomTabla, NomTablaLinH)
        Aux = "DELETE FROM " & NomTablaLinH & " WHERE " & Aux
    End If
    conn.Execute Aux
    
    
    Aux = Replace(cadWhere, NomTabla, NomTablaH)
    Aux = "DELETE FROM " & NomTablaH & " WHERE " & Aux
    conn.Execute Aux
           
        
    SQL = SQL & " FROM " & NomTabla & " WHERE " & cadWhere
    SQL = "INSERT INTO " & NomTablaH & CadenaInsercicon & SQL
    
    conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        Exit Function
    End If
    
    
    
    
    If SegundaTablaCabeceras Then
       
        
        NomTablaLinH = "bombamarca,bombaModelo,motormarca,motorModelo,observaciones,TrabajoExterior,TipoPortes,ReferPedido,FechaPed,numrepar"
        NomTablaLinH = NomTablaLinH & ",RecepAgenClien , RecepPortes, RecepAgenCliMat, RecpNumExp, FechaAlb, TipoBombResSuperHor,"
        NomTablaLinH = NomTablaLinH & "TipoBombResSuperVer, TipoBombLimSuperHor, TipoBombLimSuperVer, TipoBombResSumPoz, TipoBombLimSumPoz, TipoBombResSumVer,"
        NomTablaLinH = NomTablaLinH & "TipoBombLimSumVer, TipoBomAgitadorRes, TipoBomAgitadorLim, TipoBomResOtrosEqu, TipoBomLimOtrosEqu, DatosBommarca,"
        NomTablaLinH = NomTablaLinH & "DatosBomNumCurva, DatosBomModelo, DatosBomNumSerie, DatosBomAno, DatosBomH, DatosBomTipoRodete, DatosBomCaudal,"
        NomTablaLinH = NomTablaLinH & "DatosBomUdCaudal, DatosMotorMarca, DatosMotorModelo, DatosMotorNumSerie, DatosMotorV, DatosMotorI,DatosMotorCV,"
        NomTablaLinH = NomTablaLinH & "DatosMotorKw,DatosMotorrpm,NumParteTrabajo,NumTrabajExterno,NumReparacion,NumAlbaranVenta"
        
        
        
        
        SQL = DevuelveDesdeBD(conAri, "fechaalb", "scaalb", cadWhere & " AND 1", "1")
        If SQL = "" Then MsgBox "Error obeniendo fecha albaran. Avise soporte tecnico. El programa continua", vbExclamation
        SQL = " SELECT codtipom,numalbar," & DBSet(SQL, "F") & " fechaalb," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & "," & NomTablaLinH
        Aux = Replace(cadWhere, "scaalb", "scaalb_eu")
        SQL = SQL & " FROM scaalb_eu WHERE " & Aux
        SQL = "INSERT INTO schalb_eu(codtipom,numalbar,fechaalb1,codigusu,fechelim ,trabelim ,codincid," & NomTablaLinH & ") " & SQL
        conn.Execute SQL
    End If
    
    
    NomTablaLinH = ""
    
    
    
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabeceraHistorico = False
    Else
        InsertarCabeceraHistorico = True
    End If
End Function


Private Function InsertarLineasHistorico(cadWhere As String) As Boolean
Dim SQL As String
Dim Aux As String
Dim EsAlbaran As Boolean




On Error Resume Next


    EsAlbaran = False
    Select Case CodTipoMov
      Case "PEV" 'pedidos ventas a clientes
        NomTablaLin = "sliped"
        NomTablaLinH = "slhped"
        SQL = " SELECT scaped.numpedcl,scaped.fecpedcl,sliped.numlinea,sliped.codalmac,sliped.codartic,sliped.nomartic,sliped.ampliaci,sliped.cantidad,servidas,numbultos,precioar,dtoline1,dtoline2,importel,origpre,numlote,codccost,codtipor,codcapit,solicitadas,idL "
        SQL = SQL & " FROM scaped INNER JOIN sliped on scaped.numpedcl=sliped.numpedcl "
        SQL = SQL & " WHERE " & cadWhere
        '25-JUN: pvpInferior
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ALI", "ALT", "ALO", "ALE" '[1.3.1] 'Albaranes ventas a clientes, Mantenimientos y Reparaciones
        NomTablaLin = "slialb"
        NomTablaLinH = "slhalb"
        SQL = " SELECT scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,slialb.numlinea,slialb.codalmac,slialb.codartic,slialb.nomartic,slialb.ampliaci,slialb.cantidad,slialb.numbultos,precioar,dtoline1,dtoline2,importel,origpre ,codproveX,numlote,codccost"
        SQL = SQL & ",codtipor,codcapit ,precoste,slialb.codtraba,pvpInferior,comisionagente,idL,ordenlin,dtoCantidad "
        SQL = SQL & " FROM scaalb INNER JOIN slialb on scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
        SQL = SQL & " WHERE " & cadWhere
        EsAlbaran = True
      Case "OFE" 'Ofertas a clientes
        NomTablaLin = "slipre"
        NomTablaLinH = "slhpre"
        SQL = " SELECT scapre.numofert,scapre.fecofert,slipre.numlinea,slipre.codalmac,slipre.codartic,slipre.nomartic,slipre.ampliaci,slipre.cantidad,precioar,dtoline1,dtoline2,importel,origpre,codprovex,codcapit,esopcion "
        SQL = SQL & " FROM scapre INNER JOIN slipre on scapre.numofert=slipre.numofert"
        SQL = SQL & " WHERE " & cadWhere

      Case "ALC" 'Albaranes compras a proveedores
        NomTablaLin = "slialp"
        NomTablaLinH = "slhalp"
        SQL = "(numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,numlotes,codccost,codtipomV,numalbarV,fechaalbV) "
        SQL = SQL & " SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codprove,slialp.numlinea,slialp.codartic,slialp.codalmac,slialp.nomartic,slialp.ampliaci,slialp.cantidad,precioar,dtoline1,dtoline2,importel,numlotes,codccost,codtipomV,numalbarV,fechaalbV "
        SQL = SQL & " FROM scaalp INNER JOIN slialp on scaalp.numalbar=slialp.numalbar AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        SQL = SQL & " WHERE " & cadWhere
      
      Case "PEC" 'Pedidos compras a proveedores
        NomTablaLin = "slippr"
        NomTablaLinH = "slhppr"
        
        'FALTAn LOS DOS primeros numped y fecped    y falta codclien ,coddirec
        Aux = "numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,"
        Aux = Aux & "recibida,precioar,dtoline1,dtoline2,importel,codccost,actuacion "
        Aux = Aux & ", codtipomV , numalbarV, fechaalbV"
        
        'SQL = " SELECT scappr.numpedpr,scappr.fecpedpr,slippr.numlinea,slippr.codartic,slippr.codalmac,slippr.nomartic,slippr.ampliaci,slippr.cantidad,slippr.recibida,precioar,dtoline1,dtoline2,importel,slippr.codccost,slippr.codclien ,slippr.coddirec ,slippr.actuacion "
        SQL = " FROM scappr INNER JOIN slippr on scappr.numpedpr=slippr.numpedpr "
        SQL = SQL & " WHERE " & cadWhere
              
        SQL = "(numpedpr,fecpedpr,codclien ,coddirec," & Aux & ") SELECT scappr.numpedpr,scappr.fecpedpr,slippr.codclien ,slippr.coddirec," & Aux & SQL
    End Select
    
    SQL = "INSERT INTO " & NomTablaLinH & SQL
    
    conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        Exit Function
    End If
    'DAVID 03/NOV/2010
    'En ofertas, ademas de cbeceras lineas, hay lineas 2
    If CodTipoMov = "OFE" Then
        'NomTablaLin = "slipresail" 'mod by masl 28/10/2010
        'NomTablaLinH = "slhpresail"
        SQL = " SELECT scapre.numofert,nomarti1,caudal11,caudal12,caudal13,attm11,attm12,attm13,importe1,nomarti2,caudal21,caudal22,caudal23,"
        SQL = SQL & "attm21,attm22,attm23,importe2,nomarti3,caudal31,caudal32,caudal33,attm31,attm32,attm33,importe3"
        SQL = SQL & " FROM scapre INNER JOIN slipresail slipre on scapre.numofert=slipre.numofert"
        SQL = SQL & " WHERE " & cadWhere
        SQL = "INSERT INTO slhpresail " & SQL
        If Not ejecutar(SQL, True) Then MsgBox "Error insertando en tabla slipresail" & vbCrLf & "El programa continuara generando el pedido. " & vbCrLf & "Avise a soporte técnico", vbExclamation
            
            
            
        'ENE 2015
        If InstalacionEsEulerTaxco Then
            
            
            SQL = " SELECT scapre.numofert,numlinea,ficheroDesc,ficheronombre"
            SQL = SQL & " FROM scapre INNER JOIN sliprePdfs  on scapre.numofert=sliprePdfs.numofert"
            SQL = "INSERT INTO slhprePdfs " & SQL
            SQL = SQL & " WHERE " & cadWhere
            If Not ejecutar(SQL, True) Then MsgBox "Error insertando en tabla slhprePdfs " & vbCrLf & "El programa continuara generando el pedido. " & vbCrLf & "Avise a soporte técnico", vbExclamation
                
        End If
    End If
    
    
    If EsAlbaran Then
        If InstalacionEsEulerTaxco Then
            SQL = cadWhere
            SQL = Replace(SQL, "scaalb", "slialb_eu")
            SQL = "INSERT INTO slhalb_eu SELECT * from slialb_eu WHERE " & SQL
            If Not ejecutar(SQL, True) Then MsgBox "Error insertando en tabla hco costes " & vbCrLf & "El programa continuara generando el pedido. " & vbCrLf & "Avise a soporte técnico", vbExclamation
                
            SQL = cadWhere
            SQL = Replace(SQL, "scaalb", "slialb_eu2")
            SQL = "INSERT INTO slhalb_eu2 SELECT * from slialb_eu2 WHERE " & SQL
            If Not ejecutar(SQL, True) Then MsgBox "Error insertando en tabla hco lineas especiales " & vbCrLf & "El programa continuara generando el pedido. " & vbCrLf & "Avise a soporte técnico", vbExclamation
            
        End If
    End If

    
    
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function



Private Function BorrarTraspaso(EnHistorico As Boolean, cadWhere As String) As Boolean
'Si EnHistorico=true borra de las tablas de historico: "schtra" y "slhtra"
'Si EnHistorico=false borra de las tablas de traspaso: "scatra" y "slitra"
Dim SQL As String
Dim RS As ADODB.Recordset
Dim cad As String, cadAux As String
Dim EsAlbaran As Boolean
    BorrarTraspaso = False
    On Error GoTo EBorrar
    
    
    EsAlbaran = False
    'Eliminamos las lineas
    Select Case CodTipoMov
      Case "PEV" 'pedidos ventas  a clientes
        SQL = "Select numpedcl from scaped WHERE " & cadWhere
        cadAux = " numpedcl IN "
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ALI", "ALT", "ALO", "ALE" '[1.3.1] 'albaranes ventas a clientes,Mantenimientos y Reparaciones
        SQL = "Select numalbar from scaalb WHERE " & cadWhere
        cadAux = "codtipom=" & DBSet(CodTipoMov, "T") & " AND numalbar IN "
        EsAlbaran = True
      Case "OFE" 'Ofertas a clientes
        SQL = "Select numofert from scapre WHERE " & cadWhere
        cadAux = " numofert IN "
      Case "ALC" 'Albaranes compras a proveedores
'        SQL = "Select numalbar,fechaalb,codprove from scaalp WHERE " & cadWHERE
'        cadAux = " numalbar IN "
    End Select
    
    If CodTipoMov <> "ALC" And CodTipoMov <> "PEC" Then
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        While Not RS.EOF
            If CodTipoMov <> "ALC" Then
                cad = cad & RS.Fields(0).Value & ","
            Else
                cad = cad & "numalbar="
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        'Quitar la ultima coma de la cadena
        cad = Mid(cad, 1, Len(cad) - 1)
        
        cadAux = cadAux & "(" & cad & ")"
    Else
        cadAux = Replace(cadWhere, NomTabla, NomTablaLin)
    End If
    
    SQL = "DELETE FROM " & NomTablaLin & " WHERE " & cadAux
    
    conn.Execute SQL
    
    '03/11/2010 DAVID.  Por M.Angel
    'Si es una oferta
    If CodTipoMov = "OFE" Then
        SQL = "DELETE FROM slipresail WHERE " & cadAux
        ejecutar SQL, False  'Si da error me da lo mismo. Qu siga la fiesta
        
        
        If InstalacionEsEulerTaxco Then
            SQL = "DELETE FROM sliprePdfs   WHERE " & cadAux
            ejecutar SQL, False  'Si da error me da lo mismo. Qu siga la fiesta
        End If

        
        
    End If
    
    '10/12/2012  Moixent y Alzira llevan campos en los albaranes
    'Es decir, hay una tabla mas para borrar
    If EsAlbaran Then
        SQL = "DELETE FROM slialbcampos WHERE " & cadAux
        ejecutar SQL, False  'Si da error me da lo mismo. Qu siga la fiesta
        
        'Si tiene Manipulador fitosanitarios...
        If vParamAplic.ManipuladorFitosanitarios2 Then
            SQL = "DELETE FROM slialblotes WHERE " & cadAux
            ejecutar SQL, False  'Si da error me da lo mismo. Qu siga la fiesta
        End If
        
        
        If InstalacionEsEulerTaxco Then
                SQL = "DELETE from slialb_eu where "
                SQL = SQL & cadAux
                ejecutar SQL, False
                
                SQL = "DELETE from slialb_eu2 where "
                SQL = SQL & cadAux
                ejecutar SQL, False
                
                
        End If
    
        If vParamAplic.CartaPortes Then
            SQL = "DELETE from scaalb_portes where "
            SQL = SQL & cadAux
            ejecutar SQL, False
            
        End If
    End If
    
    
    
    'La cabecera
    SQL = "Delete from " & NomTabla
    SQL = SQL & " WHERE " & cadWhere
    conn.Execute SQL
    BorrarTraspaso = True
    
EBorrar:
    If Err.Number <> 0 Then
        BorrarTraspaso = False
    Else
        BorrarTraspaso = True
    End If
End Function



'========================================================

Public Sub CargarTagsHco(ByRef F As Form, vTabla As String, vTablaHco As String)
'Sustituye en los tags del formulario la tabla de Reparaciones (scarep)
'por la del historico de Reparaciones (schrep)
Dim Control As Object
Dim vtag As String

    For Each Control In F.Controls
        If Control.Tag <> "" Then
            vtag = Control.Tag
'            vtag = SustituirCadenas(vtag, vTabla, vTablaHco)
            vtag = Replace(vtag, vTabla, vTablaHco)
            Control.Tag = vtag
        End If
    Next Control
End Sub
