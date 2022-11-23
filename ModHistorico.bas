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
Dim Sql As String
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
        Sql = " SELECT numpedcl,fecpedcl," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & ","
        Sql = Sql & "fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        Sql = Sql & "coddirec,nomdirec,referenc,codtraba,codagent,codforpa,dtoppago,dtognral,"
        Sql = Sql & "tipofact,observa01,observa02,observa03,observa04,observa05,servcomp,restoped,numofert,fecofert,observap1,observap2,recogecl,actuacion,coddiren,observacrm"
        Sql = Sql & ", PideCliente,observaciones,cerrado,estado "
        
        
        
        
        CadenaInsercicon = " numpedcl,fecpedcl,codigusu,fechelim,trabelim,codincid,"
        CadenaInsercicon = CadenaInsercicon & "fecentre,sementre,visadore,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        CadenaInsercicon = CadenaInsercicon & "coddirec,nomdirec,referenc,codtraba,codagent,codforpa,dtoppago,dtognral,"
        CadenaInsercicon = CadenaInsercicon & "tipofact,observa01,observa02,observa03,observa04,observa05,servcomp,restoped,numofert,fecofert,observap1,observap2,recogecl,actuacion,coddiren,observacrm"
        CadenaInsercicon = CadenaInsercicon & ", PideCliente,observaciones,cerrado,estado "
        CadenaInsercicon = "(" & CadenaInsercicon & ")  "
        
        
        
        
        
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ALI", "ALT", "ALO", "ALE", "ALB", "ALD" '[1.3.1] 'Albaran de venta a clientes
        NomTabla = "scaalb"
        NomTablaH = "schalb"
        NomTablaLinH = "slhalb"
        Sql = " SELECT codtipom,numalbar,fechaalb," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & ","
        Sql = Sql & "factursn,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,"
        Sql = Sql & "coddirec,nomdirec,referenc,facturkm,cantidkm,codtraba,codtrab1,codtrab2,codagent,codforpa,codenvio,dtoppago,dtognral,"
        Sql = Sql & "tipofact,observa01,observa02,observa03,observa04,observa05,numofert,fecofert,numpedcl,fecpedcl,fecentre,sementre,esticket,numtermi,numventa "
        Sql = Sql & ",aportacion, pesoalba, portes, fecenvio, docarchiv"
        Sql = Sql & ",tipliquid, actuacion"
        Sql = Sql & ",tipoimp,origdat"
        Sql = Sql & ",coddiren,tipAlbaran"
        Sql = Sql & ",albImpreso , codzonas,observacrm"
        'Ocvubre 2015
        Sql = Sql & ", ManipuladorNumCarnet , ManipuladorFecCaducidad , ManipuladorNombre,TipoCarnet"
        'Enero 2016               abri16      ago17
        Sql = Sql & ", PideCliente,numbultos,fechaAux,puntos"
        'NOV 2018
        Sql = Sql & ", codinter,codnatura,chofer,notasportes"
        'JUN 19-DIC19
        Sql = Sql & ",  FechaEnt , perrecep ,  latitud, Longitud ,dnirecep"
            
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
        Sql = " SELECT numofert, fecofert," & "'" & Format(Now, FormatoFecha) & "' as fechamov, fecentre, aceptado, codclien, nomclien, domclien, codpobla, "
        Sql = Sql & "pobclien, proclien, nifclien, telclien, coddirec, nomdirec, referenc, codtraba, codagent, codforpa, dtoppago, dtognral, tipofact, "
        Sql = Sql & "plazos01, plazos02, plazos03, asunto01, asunto02, asunto03, asunto04, asunto05, observa01, observa02, observa03, observa04, observa05, "
        Sql = Sql & "concepto, seguiofe ,actuacion,coddiren,mailconfir,observacrm,obscompra," & cadeN & " as motivoTraspaso "
        
      Case "ALC" 'Albaranes a Proveedores (Compras)
        NomTabla = "scaalp"
        NomTablaH = "schalp"
        NomTablaLinH = "slhalp"
        Sql = " (numalbar,fechaalb,codprove,codigusu,fechelim,trabelim,codincid,nomprove,domprove,"
        Sql = Sql & "codpobla,pobprove,proprove,nifprove,telprove,codforpa,codtraba,codtrab1,dtoppago,dtognral,"
        Sql = Sql & "observa1,observa2,observa3,observa4,observa5,numpedpr,fecpedpr,fecenvio,docarchiv,codenvio,NReferencia,SReferencia,fecentrega,fentrada,emailenviado) "
        Sql = Sql & " SELECT numalbar,fechaalb,codprove," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & ","
        Sql = Sql & "nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,"
        Sql = Sql & "codforpa,codtraba,codtrab1,dtoppago,dtognral,"
        Sql = Sql & "observa1,observa2,observa3,observa4,observa5,numpedpr,fecpedpr,fecenvio,docarchiv,codenvio,NReferencia,SReferencia,fecentrega,fentrada,emailenviado"
      
      Case "PEC" 'Pedidos a Proveedores (Compras)
        NomTabla = "scappr"
        NomTablaH = "schppr"
        NomTablaLinH = "slhppr"
        Sql = " SELECT numpedpr,fecpedpr," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & ","
        Sql = Sql & "codprove,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,"
        Sql = Sql & "coddirea,coddiref,codforpa,codtraba,codtrab1,dtognral,dtoppago,"
        Sql = Sql & "restoped,codclien,observa1,observa2,observa3,observa4,observa5,tipoporte,obra,coddirre"
        Sql = Sql & ",NReferencia , SReferencia, CodEnvio, fecentrega"
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
           
        
    Sql = Sql & " FROM " & NomTabla & " WHERE " & cadWhere
    Sql = "INSERT INTO " & NomTablaH & CadenaInsercicon & Sql
    
    conn.Execute Sql
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
        
        
        
        
        Sql = DevuelveDesdeBD(conAri, "fechaalb", "scaalb", cadWhere & " AND 1", "1")
        If Sql = "" Then MsgBox "Error obeniendo fecha albaran. Avise soporte tecnico. El programa continua", vbExclamation
        Sql = " SELECT codtipom,numalbar," & DBSet(Sql, "F") & " fechaalb," & vUsu.Codigo Mod 1000 & " as codigusu," & cadeN & "," & NomTablaLinH
        Aux = Replace(cadWhere, "scaalb", "scaalb_eu")
        Sql = Sql & " FROM scaalb_eu WHERE " & Aux
        Sql = "INSERT INTO schalb_eu(codtipom,numalbar,fechaalb1,codigusu,fechelim ,trabelim ,codincid," & NomTablaLinH & ") " & Sql
        conn.Execute Sql
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
Dim Sql As String
Dim Aux As String
Dim EsAlbaran As Boolean




On Error Resume Next


    EsAlbaran = False
    Select Case CodTipoMov
      Case "PEV" 'pedidos ventas a clientes
        NomTablaLin = "sliped"
        NomTablaLinH = "slhped"
        Sql = " SELECT scaped.numpedcl,scaped.fecpedcl,sliped.numlinea,sliped.codalmac,sliped.codartic,sliped.nomartic,sliped.ampliaci,sliped.cantidad,servidas,numbultos,precioar,dtoline1,dtoline2,importel,origpre,numlote,codccost,codtipor,codcapit,solicitadas,idL,precoste "
        Sql = Sql & " FROM scaped INNER JOIN sliped on scaped.numpedcl=sliped.numpedcl "
        Sql = Sql & " WHERE " & cadWhere
        '25-JUN: pvpInferior
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ALI", "ALT", "ALO", "ALE", "ALD", "ALB" '[1.3.1] 'Albaranes ventas a clientes, Mantenimientos y Reparaciones
        NomTablaLin = "slialb"
        NomTablaLinH = "slhalb"
        Sql = " SELECT scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,slialb.numlinea,slialb.codalmac,slialb.codartic,slialb.nomartic,slialb.ampliaci,slialb.cantidad,slialb.numbultos,precioar,dtoline1,dtoline2,importel,origpre ,codproveX,numlote,codccost"
        Sql = Sql & ",codtipor,codcapit ,precoste,slialb.codtraba,pvpInferior,comisionagente,idL,ordenlin,dtoCantidad "
        Sql = Sql & " FROM scaalb INNER JOIN slialb on scaalb.codtipom=slialb.codtipom AND scaalb.numalbar=slialb.numalbar "
        Sql = Sql & " WHERE " & cadWhere
        EsAlbaran = True
      Case "OFE" 'Ofertas a clientes
        NomTablaLin = "slipre"
        NomTablaLinH = "slhpre"
        Sql = " SELECT scapre.numofert,scapre.fecofert,slipre.numlinea,slipre.codalmac,slipre.codartic,slipre.nomartic,slipre.ampliaci,slipre.cantidad,precioar,dtoline1,dtoline2,importel,origpre,codprovex,codcapit,esopcion "
        Sql = Sql & " FROM scapre INNER JOIN slipre on scapre.numofert=slipre.numofert"
        Sql = Sql & " WHERE " & cadWhere

      Case "ALC" 'Albaranes compras a proveedores
        NomTablaLin = "slialp"
        NomTablaLinH = "slhalp"
        Sql = "(numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,numlotes,codccost,codtipomV,numalbarV,fechaalbV) "
        Sql = Sql & " SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codprove,slialp.numlinea,slialp.codartic,slialp.codalmac,slialp.nomartic,slialp.ampliaci,slialp.cantidad,precioar,dtoline1,dtoline2,importel,numlotes,codccost,codtipomV,numalbarV,fechaalbV "
        Sql = Sql & " FROM scaalp INNER JOIN slialp on scaalp.numalbar=slialp.numalbar AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        Sql = Sql & " WHERE " & cadWhere
      
      Case "PEC" 'Pedidos compras a proveedores
        NomTablaLin = "slippr"
        NomTablaLinH = "slhppr"
        
        'FALTAn LOS DOS primeros numped y fecped    y falta codclien ,coddirec
        Aux = "numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,"
        Aux = Aux & "recibida,precioar,dtoline1,dtoline2,importel,codccost,actuacion "
        Aux = Aux & ", codtipomV , numalbarV, fechaalbV"
        
        'SQL = " SELECT scappr.numpedpr,scappr.fecpedpr,slippr.numlinea,slippr.codartic,slippr.codalmac,slippr.nomartic,slippr.ampliaci,slippr.cantidad,slippr.recibida,precioar,dtoline1,dtoline2,importel,slippr.codccost,slippr.codclien ,slippr.coddirec ,slippr.actuacion "
        Sql = " FROM scappr INNER JOIN slippr on scappr.numpedpr=slippr.numpedpr "
        Sql = Sql & " WHERE " & cadWhere
              
        Sql = "(numpedpr,fecpedpr,codclien ,coddirec," & Aux & ") SELECT scappr.numpedpr,scappr.fecpedpr,slippr.codclien ,slippr.coddirec," & Aux & Sql
    End Select
    
    Sql = "INSERT INTO " & NomTablaLinH & Sql
    
    conn.Execute Sql
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        Exit Function
    End If
    'DAVID 03/NOV/2010
    'En ofertas, ademas de cbeceras lineas, hay lineas 2
    If CodTipoMov = "OFE" Then
        'NomTablaLin = "slipresail" 'mod by masl 28/10/2010
        'NomTablaLinH = "slhpresail"
        Sql = " SELECT scapre.numofert,nomarti1,caudal11,caudal12,caudal13,attm11,attm12,attm13,importe1,nomarti2,caudal21,caudal22,caudal23,"
        Sql = Sql & "attm21,attm22,attm23,importe2,nomarti3,caudal31,caudal32,caudal33,attm31,attm32,attm33,importe3"
        Sql = Sql & " FROM scapre INNER JOIN slipresail slipre on scapre.numofert=slipre.numofert"
        Sql = Sql & " WHERE " & cadWhere
        Sql = "INSERT INTO slhpresail " & Sql
        If Not ejecutar(Sql, True) Then MsgBox "Error insertando en tabla slipresail" & vbCrLf & "El programa continuara generando el pedido. " & vbCrLf & "Avise a soporte técnico", vbExclamation
            
            
            
        'ENE 2015
        If InstalacionEsEulerTaxco Then
            
            
            Sql = " SELECT scapre.numofert,numlinea,ficheroDesc,ficheronombre"
            Sql = Sql & " FROM scapre INNER JOIN sliprePdfs  on scapre.numofert=sliprePdfs.numofert"
            Sql = "INSERT INTO slhprePdfs " & Sql
            Sql = Sql & " WHERE " & cadWhere
            If Not ejecutar(Sql, True) Then MsgBox "Error insertando en tabla slhprePdfs " & vbCrLf & "El programa continuara generando el pedido. " & vbCrLf & "Avise a soporte técnico", vbExclamation
                
        End If
    End If
    
    
    If EsAlbaran Then
        If InstalacionEsEulerTaxco Then
            Sql = cadWhere
            Sql = Replace(Sql, "scaalb", "slialb_eu")
            Sql = "INSERT INTO slhalb_eu SELECT * from slialb_eu WHERE " & Sql
            If Not ejecutar(Sql, True) Then MsgBox "Error insertando en tabla hco costes " & vbCrLf & "El programa continuara generando el pedido. " & vbCrLf & "Avise a soporte técnico", vbExclamation
                
            Sql = cadWhere
            Sql = Replace(Sql, "scaalb", "slialb_eu2")
            Sql = "INSERT INTO slhalb_eu2 SELECT * from slialb_eu2 WHERE " & Sql
            If Not ejecutar(Sql, True) Then MsgBox "Error insertando en tabla hco lineas especiales " & vbCrLf & "El programa continuara generando el pedido. " & vbCrLf & "Avise a soporte técnico", vbExclamation
            
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
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Cad As String, cadAux As String
Dim EsAlbaran As Boolean
    BorrarTraspaso = False
    On Error GoTo EBorrar
    
    
    EsAlbaran = False
    'Eliminamos las lineas
    Select Case CodTipoMov
      Case "PEV" 'pedidos ventas  a clientes
        Sql = "Select numpedcl from scaped WHERE " & cadWhere
        cadAux = " numpedcl IN "
      Case "ALV", "ALM", "ALR", "ALS", "ART", "ALI", "ALT", "ALO", "ALE", "ALD", "ALB" '[1.3.1] 'albaranes ventas a clientes,Mantenimientos y Reparaciones
        Sql = "Select numalbar from scaalb WHERE " & cadWhere
        cadAux = "codtipom=" & DBSet(CodTipoMov, "T") & " AND numalbar IN "
        EsAlbaran = True
      Case "OFE" 'Ofertas a clientes
        Sql = "Select numofert from scapre WHERE " & cadWhere
        cadAux = " numofert IN "
      Case "ALC" 'Albaranes compras a proveedores
'        SQL = "Select numalbar,fechaalb,codprove from scaalp WHERE " & cadWHERE
'        cadAux = " numalbar IN "
    End Select
    
    If CodTipoMov <> "ALC" And CodTipoMov <> "PEC" Then
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not RS.EOF
            If CodTipoMov <> "ALC" Then
                Cad = Cad & RS.Fields(0).Value & ","
            Else
                Cad = Cad & "numalbar="
            End If
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        'Quitar la ultima coma de la cadena
        Cad = Mid(Cad, 1, Len(Cad) - 1)
        
        cadAux = cadAux & "(" & Cad & ")"
    Else
        cadAux = Replace(cadWhere, NomTabla, NomTablaLin)
    End If
    
    Sql = "DELETE FROM " & NomTablaLin & " WHERE " & cadAux
    
    conn.Execute Sql
    
    '03/11/2010 DAVID.  Por M.Angel
    'Si es una oferta
    If CodTipoMov = "OFE" Then
        Sql = "DELETE FROM slipresail WHERE " & cadAux
        ejecutar Sql, False  'Si da error me da lo mismo. Qu siga la fiesta
        
        
        If InstalacionEsEulerTaxco Then
            Sql = "DELETE FROM sliprePdfs   WHERE " & cadAux
            ejecutar Sql, False  'Si da error me da lo mismo. Qu siga la fiesta
        End If

        
        
    End If
    
    '10/12/2012  Moixent y Alzira llevan campos en los albaranes
    'Es decir, hay una tabla mas para borrar
    If EsAlbaran Then
        Sql = "DELETE FROM slialbcampos WHERE " & cadAux
        ejecutar Sql, False  'Si da error me da lo mismo. Qu siga la fiesta
        
        'Si tiene Manipulador fitosanitarios...
        If vParamAplic.ManipuladorFitosanitarios2 Then
            Sql = "DELETE FROM slialblotes WHERE " & cadAux
            ejecutar Sql, False  'Si da error me da lo mismo. Qu siga la fiesta
        End If
        
        
        If InstalacionEsEulerTaxco Then
                Sql = "DELETE from slialb_eu where "
                Sql = Sql & cadAux
                ejecutar Sql, False
                
                Sql = "DELETE from slialb_eu2 where "
                Sql = Sql & cadAux
                ejecutar Sql, False
                
                Sql = "DELETE from scaalb_eu where "
                Sql = Sql & cadAux
                ejecutar Sql, False
                
        End If
    
        If vParamAplic.CartaPortes Then
            Sql = "DELETE from scaalb_portes where "
            Sql = Sql & cadAux
            ejecutar Sql, False
            
        End If
    End If
    
    
    
    'La cabecera
    Sql = "Delete from " & NomTabla
    Sql = Sql & " WHERE " & cadWhere
    conn.Execute Sql
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
