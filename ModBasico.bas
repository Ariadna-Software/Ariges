Attribute VB_Name = "ModBasico"
Option Explicit


Public Sub AyudaProveedores(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;S|txtAux(2)|T|Nombre comercial|4500|;S|txtAux(3)|T|NIF|1500|;S|txtAux(4)|T|F.Ult.Compra|1500|;"
    frmCom.CadenaConsulta = "SELECT sprove.codprove, sprove.nomprove, sprove.nomcomer, sprove.nifprove, sprove.fechamov "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999999|sprove|codprove|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sprove|nomprove|||"
    frmCom.Tag3 = "Nombre comercial|T|N|||sprove|nomcomer|||"
    frmCom.Tag4 = "NIF|T|N|||sprove|nifprove|||"
    frmCom.Tag5 = "Fecha Ult.Compra|F|N|||sprove|fechamov|dd/mm/yyyy||"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 40
    frmCom.Maxlen3 = 40
    frmCom.Maxlen4 = 15
    frmCom.Maxlen5 = 10

    frmCom.pConn = conAri

    frmCom.tabla = "sprove"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "codprove"
    frmCom.TipoCP = "N"
    frmCom.Formulario = "Proveedores"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Proveedores"
    
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 6000
    
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaProveedoresV(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|NIF|2005|;S|txtAux(1)|T|Nombre|5595|;"
    frmCom.CadenaConsulta = "SELECT sprvar.nifprove, sprvar.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sprvar "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "N.I.F.|T|N|||sprvar|nifprove||S|"
    frmCom.Tag2 = "Nombre Proveedor Varios|T|N|||sprvar|nomprove||N|"
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 30

    frmCom.pConn = conAri

    frmCom.tabla = "sprvar"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "nifprove"
    frmCom.TipoCP = "N"
    frmCom.Formulario = "ProveedoresV"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Proveedores Varios"
    
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 600
    
    
    frmCom.Show vbModal
End Sub



Public Sub AyudaCentroCoste(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;"
    
    If vParamAplic.ContabilidadNueva Then
        frmCom.CadenaConsulta = "SELECT ccoste.codccost, ccoste.nomccost "
        frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM ccoste "
        frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    
        If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
        
        frmCom.Tag1 = "Código|T|N|||ccoste|codccost||S|"
        frmCom.Tag2 = "Nombre|T|N|||ccoste|nomccost|||"
        frmCom.Tag3 = ""
    Else
        frmCom.CadenaConsulta = "SELECT cabccost.codccost, cabccost.nomccost "
        frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM cabccost "
        frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    
        If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
        
        frmCom.Tag1 = "Código|T|N|||cabccost|codccost||S|"
        frmCom.Tag2 = "Nombre|T|N|||cabccost|nomccost|||"
        frmCom.Tag3 = ""
    End If
    
    frmCom.Maxlen1 = 4
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 0

    frmCom.pConn = conConta

    If vParamAplic.ContabilidadNueva Then
        frmCom.tabla = "ccoste"
        frmCom.CampoCP = "codccost"
    Else
        frmCom.tabla = "cabccost"
        frmCom.CampoCP = "codccost"
    End If
    frmCom.TipoCP = "T"
    frmCom.Titulo = "Centros de Coste"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = ""
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 4900
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    frmCom.Titulo = "Centros de Coste"
    
    Redimensiona frmCom, -1500
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaCtasContables(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional NumDigit As Integer, Optional BD As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1405|;S|txtAux(1)|T|Nombre|5595|;"
    frmBas.CadenaConsulta = "SELECT cuentas.codmacta, cuentas.nommacta "
    If BD <> "" Then
        frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM " & BD & ".cuentas "
    Else
        frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM cuentas "
    End If
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    
    If NumDigit <> 0 Then
        frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and length(codmacta) = " & DBSet(NumDigit, "N")
    Else
        frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and apudirec = 'S'"
    End If

    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|T|N|||cuentas|codmacta||S|"
    frmBas.Tag2 = "Nombre|T|N|||cuentas|nommacta|||"
    frmBas.Tag3 = ""
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conConta
    
    If BD <> "" Then
        frmBas.tabla = BD & ".cuentas"
    Else
        frmBas.tabla = "cuentas"
    End If
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CampoCP = "codmacta"
    frmBas.Formulario = "Cuentas"
    frmBas.TipoCP = "T"
    frmBas.Titulo = "Cuentas Contables"
    frmBas.LenCta = NumDigit
    
    frmBas.CodigoActual = ""
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmBas.DataGrid1.Height = 8140 '7420
    frmBas.DataGrid1.Top = 870
    frmBas.FrameBotonGnral.visible = True
    frmBas.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmBas, 0
    
    
    frmBas.Show vbModal
    
    
End Sub

Public Sub AyudaFamilias(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|4095|;S|txtAux(2)|T|Codigo|1000|;S|txtAux(3)|T|Nombre Proveedor|4000|;"
    frmCom.CadenaConsulta = "SELECT sfamia.codfamia, sfamia.nomfamia, sfamia.codprove, sprove.nomprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sfamia left join sprove on sfamia.codprove = sprove.codprove "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|sfamia|codfamia|0000|S|"
    frmCom.Tag2 = "Descripción|T|N|||sfamia|nomfamia|||"
    frmCom.Tag3 = "Proveedor|N|S|0||sfamia|codprove|000000||"
    frmCom.Tag4 = "Nombre|T|S|||sprove|nomprove|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 25
    frmCom.Maxlen3 = 10
    frmCom.Maxlen4 = 25
    frmCom.pConn = conAri
    
    frmCom.tabla = "sfamia left join sprove on sfamia.codprove = sprove.codprove "
    frmCom.CampoCP = "sfamia.codfamia"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Familias"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = ""
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    
    
    Redimensiona frmCom, 3000
    
    frmCom.Show vbModal
End Sub






Public Sub AyudaTDiariosContables(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1005|;S|txtAux(1)|T|Nombre|5995|;"
    frmBas.CadenaConsulta = "SELECT tiposdiario.numdiari, tiposdiario.desdiari "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM tiposdiario "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código de Diario|N|N|||tiposdiario|numdiari|00|S|"
    frmBas.Tag2 = "Descripción del diario|T|N|||tiposdiario|desdiari|||"
    frmBas.Tag3 = ""
    frmBas.Maxlen1 = 2
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conConta
    
    frmBas.tabla = "tiposdiario"
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CampoCP = "numdiari"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Tipos de Diario"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 0
    
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaConceptosContables(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1005|;S|txtAux(1)|T|Nombre|5995|;"
    frmBas.CadenaConsulta = "SELECT conceptos.codconce, conceptos.nomconce "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM conceptos "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Código|N|N|||conceptos|codconce|##0|S|"
    frmBas.Tag2 = "Descripción|T|N|||conceptos|nomconce|||"
    frmBas.Tag3 = ""
    frmBas.Maxlen1 = 2
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 0
    
    frmBas.pConn = conConta
    
    frmBas.tabla = "conceptos"
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CampoCP = "codconce"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Conceptos Contables"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    Redimensiona frmBas, 0
    
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaTIvaContabilidad(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
' en total son 7000 = 905 + 4595 hay que quitarle al width 1500
    
    
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1005|;S|txtAux(1)|T|Nombre|3995|;S|txtAux(2)|T|% Iva|1000|;S|txtAux(3)|T|% Rec.|1000|;S|txtAux(4)|T|Tipo|2000|;"
    frmBas.CadenaConsulta = "SELECT tiposiva.codigiva, tiposiva.nombriva, tiposiva.porceiva, tiposiva.porcerec,  CASE tiposiva.tipodiva WHEN 0 THEN ""IVA"" WHEN 1 THEN ""IGIC"" WHEN 2 THEN ""BIEN DE INVERSION"" WHEN 3 THEN ""NO DEDUCIBLE"" WHEN 4 THEN ""SUPLIDOS"" END ntipo  "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM tiposiva "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    
    frmBas.Tag1 = "Código IVA|N|N|0|99|tiposiva|codigiva|00|S|"
    frmBas.Tag2 = "Descripcion Tipo IVA|T|N|||tiposiva|nombriva|||"
    frmBas.Tag3 = "% IVA|N|N|0|100|tiposiva|porceiva|#0.00||"
    frmBas.Tag4 = "% Recargo de IVA|N|N|0|100|tiposiva|porcerec|#0.00||"
    frmBas.Tag5 = "Tipo de IVA|T|N|||tiposiva|ntipo|||"
    frmBas.Maxlen1 = 2
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 4
    frmBas.Maxlen4 = 4
    frmBas.Maxlen5 = 20
    
    
    frmBas.pConn = conConta
    
    frmBas.tabla = "tiposiva"
    frmBas.DatosADevolverBusqueda = "0|1|2|"
    frmBas.CampoCP = "codigiva"
    frmBas.TipoCP = "N"
    frmBas.Caption = "Tipos de IVA Contabilidad"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmBas.DataGrid1.Height = 7420
    frmBas.DataGrid1.Top = 870
    frmBas.FrameBotonGnral.visible = True
    frmBas.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmBas, 2000
    
    frmBas.Show vbModal
    
    
End Sub


Public Sub AyudaAlmMovArtPrev(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Artículo|2205|;S|txtAux(1)|T|Nombre|4795|;"
    
    frmCom.CadenaConsulta = "SELECT distinct smoval.codartic, sartic.nomartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM smoval left join sartic on smoval.codartic = sartic.codartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Artículo|T|N|||smoval|codartic||S|"
    frmCom.Tag2 = "Denominacion|T|N|||sartic|nomartic||N|"
    
    frmCom.Maxlen1 = 16
    frmCom.Maxlen2 = 35
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "smoval"
    frmCom.CampoCP = "codartic"
    frmCom.TipoCP = "N"
    frmCom.Titulo = "Movimientos Artículos"
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = ""
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaArticulos(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional parAlmacen As String, Optional DesdeTPV As Boolean, Optional SinAvanzada As Boolean)
Dim incre As Long

    frmCom.CadenaTots = "S|txtAux(0)|T|Código|2100|;S|txtAux(1)|T|Descripción|4640|;S|txtAux(2)|T|Cod.Asociación|1600|;S|txtAux(3)|T|Stock|1500|;S|txtAux(4)|T|PVP|1300|;"
    If DesdeTPV Then
        frmCom.CadenaTots = frmCom.CadenaTots & "S|txtAux(5)|T|PVP IVA|1500|;"
        incre = 1500
    Else
        If InstalacionEsEulerTaxco Then
            frmCom.CadenaTots = frmCom.CadenaTots & "S|txtAux(5)|T|Ctr. stock|1000|;"
            incre = 1000
        Else
            frmCom.CadenaTots = frmCom.CadenaTots & "S|txtAux(5)|T|Referencia Provedor|2400|;"
            incre = 2400
        End If
    End If
    
    If parAlmacen = "" Then
        If vUsu.AlmacenPorDefecto2 <> "" Then
            parAlmacen = vUsu.AlmacenPorDefecto2
        Else
            parAlmacen = DevuelveDesdeBDNew(conAri, "salmpr", "min(codalmac)", "", "")
        End If
    End If
    
    frmCom.CadenaConsulta = "Select sartic.codartic,nomartic,codtelem,salmac.canstock,"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & "preciove,"
    If DesdeTPV Then
        frmCom.CadenaConsulta = frmCom.CadenaConsulta & "if(isnull(porcen1),preciove,preciove*(1+(porcen1/100)))"
    Else
        If InstalacionEsEulerTaxco Then
            frmCom.CadenaConsulta = frmCom.CadenaConsulta & "if(ctrstock=0,'','Si') "
        Else
            frmCom.CadenaConsulta = frmCom.CadenaConsulta & "referprov"
        End If
    End If
    'CadenaConsulta = CadenaConsulta & " FROM  (sartic INNER JOIN salmac ON sartic.codartic=salmac.codartic AND codalmac = " & parAlmacen & " ) "
    'If Me.DesdeTPV Then CadenaConsulta = CadenaConsulta & " LEFT OUTER JOIN tmpinformes ON sartic.codigiva=tmpinformes.codigo1 AND codusu = " & vUsu.codigo
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM  sartic ,salmac "
    If DesdeTPV Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & ", tmpinformes  "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE "
    If DesdeTPV Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " codusu = " & vUsu.Codigo & " AND "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " codalmac = " & parAlmacen & " AND sartic.codartic=salmac.codartic "
    If DesdeTPV Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " AND sartic.codigiva=tmpinformes.codigo1 "
    
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|T1|N|||sartic|codartic||S|"
    frmCom.Tag2 = "Descripcion|T|N|||sartic|nomartic||N|"
    frmCom.Tag3 = "Cod. Asociacionl|T|N|||sartic|codtelem||N|"
    frmCom.Tag4 = "Stock|N|N|||salmac|canstock|#,###,###,##0.00|N|"
    frmCom.Tag5 = "Stock|N|N|||sartic|preciove|#,##0.0000|N|"
    If DesdeTPV Then
        frmCom.Tag6 = "Stock|N|N|||salmac|canstock|#,###,###,##0.0000|N|"
    Else
        If InstalacionEsEulerTaxco Then
            frmCom.Tag6 = "Stock|N|N|||salmac|canstock|#,###,###,##0.0000|N|"
        Else
            frmCom.Tag6 = "refprove|T|N|||sartic|referprov|||"""
        End If
    End If
    
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 40
    frmCom.Maxlen3 = 30
    frmCom.Maxlen4 = 30
    frmCom.Maxlen5 = 30
    frmCom.Maxlen6 = 30
        
    
    frmCom.pConn = conAri
    frmCom.CampoDeOrdenacion = "sartic.nomartic"
    frmCom.tabla = "sartic ,salmac "
    If DesdeTPV Then frmCom.tabla = frmCom.tabla & ", tmpinformes "
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "sartic.codartic"
    frmCom.TipoCP = "T"
    frmCom.Formulario = "Articulos"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Articulos"
    
    
    frmCom.CodigoActual = ""
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
     '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 4140 + incre
    
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaTrabajadores(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|1305|;S|txtAux(1)|T|Nombre|4095|;S|txtAux(2)|T|Login|2000|;S|txtAux(3)|T|EMail|4600|;"
    frmCom.CadenaConsulta = "SELECT straba.codtraba, straba.nomtraba, straba.login, straba.maitraba "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM straba "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE true "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código Trabajador|N|N|0|9999|straba|codtraba|0000|S|"
    frmCom.Tag2 = "Nombre Trabajador|T|N|||straba|nomtraba||N|"
    frmCom.Tag3 = "Login Trabajador|T|S|||straba|login||N|"
    frmCom.Tag4 = "e-mail|T|S|||straba|maitraba||N|"
    frmCom.Maxlen1 = 4
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 20
    frmCom.Maxlen4 = 40

    frmCom.pConn = conAri

    frmCom.tabla = "straba"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "codtraba"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    
    frmCom.Caption = "Trabajadores"
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 5000
    
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaAlmMovimientosPrev(frmCom As frmBasico2, EsHco As Boolean, Optional CodActual As String, Optional cWhere As String)
Dim tabla As String

    tabla = "scamov"
    If EsHco Then tabla = "schmov"
    frmCom.CadenaTots = "S|txtAux(0)|T|NºMovimiento|1705|;S|txtAux(1)|T|Fecha|1450|;S|txtAux(2)|T|Almacén|950|;S|txtAux(3)|T|Descripción Almacén|4395|;"
    
    frmCom.CadenaConsulta = "SELECT distinct " & tabla & ".codmovim, " & tabla & ".fecmovim, " & tabla & ".codalmac, salmpr.nomalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM " & tabla & " left join salmpr on " & tabla & ".codalmac = salmpr.codalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Movimiento|N|S|0||" & tabla & "|codmovim|0000000|S|"
    frmCom.Tag2 = "Fecha|F|N|||" & tabla & "|fecmovim|dd/mm/yyyy|N|"
    frmCom.Tag3 = "Cod. Almacen|N|N|0|999|" & tabla & "|codalmac|000|N|"
    frmCom.Tag4 = "Descripcion Almacen|T|N|||salmpr|nomalmac||N|"
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 35
    
    frmCom.pConn = conAri
    
    frmCom.tabla = tabla & " left join salmpr on " & tabla & ".codalmac = salmpr.codalmac "
    frmCom.CampoCP = tabla & ".codmovim"
    frmCom.TipoCP = "N"
    frmCom.Titulo = "Movimientos Almacén"
    If EsHco Then frmCom.Titulo = "Histórico de " & frmCom.Titulo

    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = ""
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 1500
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaAlmMovTraspasoPrev(frmCom As frmBasico2, EsHco As Boolean, Optional CodActual As String, Optional cWhere As String)
Dim tabla As String

    tabla = "scatra"
    If EsHco Then tabla = "schtra"
    frmCom.CadenaTots = "S|txtAux(0)|T|NºTraspaso|1705|;S|txtAux(1)|T|Fecha|1450|;S|txtAux(2)|T|Código|950|;S|txtAux(3)|T|Almacén Origen|4395|;S|txtAux(4)|T|Código|950|;S|txtAux(5)|T|Almacén Destino|4395|;"
    
    frmCom.CadenaConsulta = "SELECT distinct " & tabla & ".codtrasp, " & tabla & ".fechatra, " & tabla & ".almaorig, oo.nomalmac, " & tabla & ".almadest, dd.nomalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (" & tabla & " left join salmpr oo on " & tabla & ".almaorig = oo.codalmac) "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " left join salmpr dd on " & tabla & ".almadest = dd.codalmac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Traspaso|N|S|0||scatra|codtrasp|0000000|S|"
    frmCom.Tag2 = "Fecha|F|N|||scatra|fechatra|dd/mm/yyyy|N|"
    frmCom.Tag3 = "Almacen Origen|N|N|0|999|scatra|almaorig|000|N|"
    frmCom.Tag4 = "Descripcion Almacen Origen|T|N|||salmpr|nomalmac||N|"
    frmCom.Tag5 = "Almacen Destino|N|N|0|999|scatra|almadest|000|N|"
    frmCom.Tag6 = "Descripcion Almacen Destino|T|N|||salmpr|nomalmac||N|"
    
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 35
    frmCom.Maxlen5 = 3
    frmCom.Maxlen6 = 35
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(" & tabla & " left join salmpr oo on " & tabla & ".codalmac = oo.codalmac) left join salmpr dd on " & tabla & ".almadest = dd.codalmac"
    
    frmCom.CampoCP = tabla & ".codtrasp"
    frmCom.TipoCP = "N"
    frmCom.Titulo = "Traspaso de Almacén"
    If EsHco Then frmCom.Titulo = "Histórico de " & frmCom.Titulo

    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = ""
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    Redimensiona frmCom, 6845
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaClientes(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|4595|;S|txtAux(2)|T|Nombre comercial|4500|;S|txtAux(3)|T|NIF|1500|;S|txtAux(4)|T|F.Ult.Movim|1500|;S|txtAux(5)|T|Fec.Alta|1500|;"
    frmCom.CadenaConsulta = "SELECT sclien.codclien, sclien.nomclien, sclien.nomcomer, sclien.nifclien, sclien.fechamov, sclien.fechaalt "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999999|sclien|codclien|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sclien|nomclien|||"
    frmCom.Tag3 = "Nombre comercial|T|N|||sclien|nomcomer|||"
    frmCom.Tag4 = "NIF|T|N|||sclien|nifclien|||"
    frmCom.Tag5 = "Fecha Ult.Movim.|F|S|||sclien|fechamov|dd/mm/yyyy||"
    frmCom.Tag6 = "Fecha Alta|F|N|||sclien|fechaalt|dd/mm/yyyy||"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 40
    frmCom.Maxlen3 = 40
    frmCom.Maxlen4 = 15
    frmCom.Maxlen5 = 10
    frmCom.Maxlen6 = 10

    frmCom.pConn = conAri
    frmCom.CampoDeOrdenacion = "sclien.nomclien"
    frmCom.tabla = "sclien"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "codclien"
    frmCom.TipoCP = "N"
    frmCom.Formulario = "Clientes"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Clientes"
    
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 7500
    
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaBancosPropios(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|6095|;"
    frmCom.CadenaConsulta = "SELECT sbanpr.codbanpr, sbanpr.nombanpr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sbanpr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|9999|sbanpr|codbanpr|0000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sbanpr|nombanpr|||"
    frmCom.Maxlen1 = 2
    frmCom.Maxlen2 = 40

    frmCom.pConn = conAri

    frmCom.tabla = "sbanpr"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "codbanpr"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Bancos Propios"
    
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 0
    
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaFormasPago(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional BD As String)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|6095|;S|txtAux(2)|T|Nro.Vtos|1200|;S|txtAux(3)|T|Primer Vto|1200|;S|txtAux(4)|T|Resto Vtos|1200|;"
    frmCom.CadenaConsulta = "SELECT sforpa.codforpa, sforpa.nomforpa, sforpa.numerove, sforpa.primerve, sforpa.restoven "
    
    If BD <> "" Then
        frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM " & BD & ".sforpa "
    Else
        frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sforpa "
    End If
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|sforpa|codforpa|000|S|"
    frmCom.Tag2 = "Descripción|T|N|||sforpa|nomforpa|||"
    frmCom.Tag3 = "Nº Vencimientos|N|N|1|99999|sforpa|numerove|0|N|"
    frmCom.Tag4 = "Primer Vencimiento|N|N|0||sforpa|primerve|0|N|"
    frmCom.Tag5 = "Resto Vencimientos|N|S|||sforpa|restoven|0|N|"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 5
    frmCom.Maxlen4 = 5
    frmCom.Maxlen5 = 5
    frmCom.pConn = conAri
    
    If BD <> "" Then
        frmCom.tabla = BD & ".sforpa"
    Else
        frmCom.tabla = "sforpa"
    End If
    
    frmCom.CampoCP = "codforpa"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Formas de Pago"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 3600
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaAgentesComerciales(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|6095|;"
    frmCom.CadenaConsulta = "SELECT sagent.codagent, sagent.nomagent "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sagent "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código Agente Comercial|N|N|0|9999|sagent|codagent|0000|S|"
    frmCom.Tag2 = "Nombre del Agente Comercial|T|N|||sagent|nomagent||N|"
    frmCom.Maxlen1 = 4
    frmCom.Maxlen2 = 30
    frmCom.pConn = conAri
    
    frmCom.tabla = "sagent"
    frmCom.CampoCP = "codagent"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Agentes Comerciales"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 0
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaClientesV(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|NIF|2005|;S|txtAux(1)|T|Nombre|5595|;S|txtAux(2)|T|Teléfono|2000|;"
    frmCom.CadenaConsulta = "SELECT sclvar.nifclien, sclvar.nomclien, sclvar.telclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sclvar "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "N.I.F.|T|N|||sclvar|nifclien||S|"
    frmCom.Tag2 = "Nombre Cliente Varios|T|N|||sclvar|nomclien||N|"
    frmCom.Tag3 = "Teléfono|T|S|||sclvar|telclien||N|"
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 15

    frmCom.pConn = conAri

    frmCom.tabla = "sclvar"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "nifclien"
    frmCom.TipoCP = "T"
    frmCom.Formulario = "ClientesV"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Clientes Varios"
    
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 2600
    
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaCartas(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|1005|;S|txtAux(1)|T|Descipción|5995|;"
    frmCom.CadenaConsulta = "SELECT scartas.codcarta, scartas.descarta "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scartas "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Codigo Carta|N|N|0|999|scartas|codcarta|000|S|"
    frmCom.Tag2 = "Descripción|T|S|||scartas|descarta||N|"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 50

    frmCom.pConn = conAri

    frmCom.tabla = "scartas"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "codcarta"
    frmCom.TipoCP = "N"
    frmCom.Formulario = "Cartas"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Cartas"
    
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 0
    
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaDireccionesCompra(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Cód|505|;N||||0|;S|Combo1(0)|C|Tipo|1150|;S|txtAux(1)|T|Descripción|5005|;S|txtAux(2)|T|Dirección|4000|;S|txtAux(3)|T|Población|2005|;"
    frmCom.CadenaConsulta = "SELECT sdirpr.coddirec, tipodire, if(tipodire=0,'Albarán','Factura') tipo, sdirpr.nomdirec, "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " sdirpr.domdirec, sdirpr.pobdirec "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sdirpr "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Cod. Dirección|N|N|0|999|sdirpr|coddirec|000|S|"
    frmCom.Tag10 = "Tipo Dirección|N|N|||sdirpr|tipodire||N|"
    frmCom.Tag2 = "Nombre Dirección|T|N|||sdirpr|nomdirec||N|"
    frmCom.Tag3 = "Domicilio|T|N|||sdirpr|domdirec||N|"
    frmCom.Tag4 = "Población|T|N|||sdirpr|pobdirec||N|"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 30
    frmCom.Maxlen4 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "sdirpr"
    frmCom.CampoCP = "coddirec"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Direcciones de Compra"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 5700
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaTarifaPrecios(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Artículo|2100|;S|txtAux(1)|T|Descripción|4700|;S|txtAux(2)|T|Tarifa|700|;S|txtAux(3)|T|Nombre|1300|;S|txtAux(4)|T|Precio|1500|;"
    frmCom.CadenaConsulta = "SELECT slista.codartic, sartic.nomartic, slista.codlista, starif.nomlista, slista.precioac"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (slista LEFT JOIN sartic ON slista.codartic=sartic.codartic) left join starif on slista.codlista = starif.codlista"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Cod.Artículo|T|N|||slista|codartic||S|"
    frmCom.Tag2 = "Nombre Artículo|T|N|||sartic|nomartic||N|"
    frmCom.Tag3 = "Cod.Tarifa|N|N|0|999|slista|codlista|000|S|"
    frmCom.Tag4 = "Nombre Tarifa|T|N|||starif|nomlista||N|"
    frmCom.Tag5 = "Precio Actual|N|N|0|999999.0000|slista|precioac|###,##0.0000|N|"
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 30
    frmCom.Maxlen5 = 15
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(slista LEFT JOIN sartic ON slista.codartic=sartic.codartic) left join starif on slista.codlista = starif.codlista"
    frmCom.CampoCP = "slista.codartic"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Tarifas de Artículos"
    
    frmCom.DatosADevolverBusqueda = "0|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 3300
    
    frmCom.Show vbModal

End Sub

Public Sub AyudaPreciosEspeciales(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Clientes|1000|;S|txtAux(1)|T|Nombre|4500|;S|txtAux(2)|T|Artículo|2100|;S|txtAux(3)|T|Descripción|4500|;"
    frmCom.CadenaConsulta = "SELECT sprees.codclien, sclien.nomclien, sprees.codartic, sartic.nomartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (sprees LEFT JOIN sclien on sprees.codclien = sclien.codclien) LEFT JOIN sartic ON sprees.codartic=sartic.codartic"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Cliente|N|N|0|999999|sprees|codclien|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sclien|nomclien||N|"
    frmCom.Tag3 = "Artículo|T|N|||sprees|codartic||S|"
    frmCom.Tag4 = "Nombre Artículo|T|N|||sartic|nomartic||N|"
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(sprees LEFT JOIN sclien on sprees.codclien = sclien.codclien) LEFT JOIN sartic ON sprees.codartic=sartic.codartic"
    frmCom.CampoCP = "sprees.codartic"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Precios Especiales"
    
    frmCom.DatosADevolverBusqueda = "0|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 5100
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaPromociones(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Artículo|2100|;S|txtAux(1)|T|Descripción|4700|;S|txtAux(2)|T|Tarifa|700|;S|txtAux(3)|T|Nombre|1300|;"
    frmCom.CadenaConsulta = "SELECT spromo.codartic, sartic.nomartic, spromo.codlista, starif.nomlista"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (spromo LEFT JOIN sartic ON spromo.codartic=sartic.codartic) left join starif on spromo.codlista = starif.codlista"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Cod. Artículo|T|N|||spromo|codartic||S|"
    frmCom.Tag2 = "Nombre Artículo|T|N|||sartic|nomartic||N|"
    frmCom.Tag3 = "Cod. Tarifa|N|N|0|999|spromo|codlista|000|S|"
    frmCom.Tag4 = "Nombre Tarifa|T|N|||starif|nomlista||N|"
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(spromo LEFT JOIN sartic ON spromo.codartic=sartic.codartic) left join starif on spromo.codlista = starif.codlista"
    frmCom.CampoCP = "spromo.codartic"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Promociones Tarifas"
    
    frmCom.DatosADevolverBusqueda = "0|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 1800
    
    frmCom.Show vbModal

End Sub



Public Sub AyudaOfertas(frmOfe As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
Dim Sql As String
Dim tabla As String

    Screen.MousePointer = vbHourglass


    frmOfe.CadenaTots = "S|txtAux(0)|T|Nº Ofe.|1105|;S|txtAux(1)|T|Fecha|1290|;S|txtAux(2)|T|Cod.Cli.|1050|;S|txtAux(3)|T|Nombre|4200|;"
    frmOfe.CadenaTots = frmOfe.CadenaTots & "S|txtAux(4)|T|Trab.|900|;S|txtAux(5)|T|Nom. trab.|2700|;S|txtAux(6)|T|Base|1400|;"
    
    
    
    tabla = " (select scapre.numofert,scapre.fecofert,codclien,nomclien,scapre.codtraba,nomtraba"
    tabla = tabla & " ,round(sum(importel) *  ( (100 - ( coalesce(dtoppago,0)   +   coalesce(dtognral,0) ))  /100   ),2) as base"
    tabla = tabla & " from scapre inner join straba on scapre.codtraba=straba.codtraba"
    tabla = tabla & " inner join slipre on scapre.numofert=slipre.numofert WHERE true "
    
    If cWhere <> "" Then tabla = tabla & " and " & cWhere
    
    tabla = tabla & " GROUP BY scapre.numofert ) as tt"
    
    Sql = "SELECT numofert,fecofert,codclien,nomclien,codtraba,nomtraba,base FROM " & tabla & " WHERE true"
    
    
    
    
    
    
    
    frmOfe.CadenaConsulta = Sql
    
    frmOfe.Tag1 = "Oferta|N|N|0|999999|tt|numofert|000000|S|"
    frmOfe.Tag2 = "Fecha|F|N|||tt|fecofert|dd/mm/yyyy||"
    frmOfe.Tag3 = "Cod.Cli|N|N|||tt|codclien|00000||"
    frmOfe.Tag4 = "Nombre|T|N|||tt|nomclien|||"
    frmOfe.Tag5 = "C.Trab.|N|N|||tt|codtraba|0000||"
    frmOfe.Tag6 = "Trabajador|T|N|||tt|nomtraba|||"
    frmOfe.Tag7 = "base|N|N|||tt|base|" & FormatoImporte & "||"
    
    frmOfe.Maxlen1 = 6
    frmOfe.Maxlen2 = 20
    frmOfe.Maxlen3 = 20
    frmOfe.Maxlen4 = 20
    frmOfe.Maxlen5 = 50
    frmOfe.Maxlen6 = 50
    frmOfe.Maxlen7 = 50
    

    frmOfe.pConn = conAri

    frmOfe.tabla = tabla
    frmOfe.DatosADevolverBusqueda = "0|1|"
    frmOfe.CampoCP = "numofert"
    frmOfe.TipoCP = "N"
    frmOfe.Formulario = "Ofertas"
    If SinAvanzada Then frmOfe.DeConsulta = True
    frmOfe.Caption = "Ofertas"
    
    
    frmOfe.CodigoActual = 0
    If CodActual <> "" Then frmOfe.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmOfe.DataGrid1.Height = 7420
    frmOfe.DataGrid1.Top = 870
    frmOfe.FrameBotonGnral.visible = True
    frmOfe.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmOfe, 6000
    

    frmOfe.Show vbModal
    
    
   
    
    
    
End Sub

Public Sub AyudaClientesPotenciales(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Nombre|5095|;S|txtAux(2)|T|Nombre comercial|5000|;"
    frmCom.CadenaConsulta = "SELECT sclipot.codclien, sclipot.nomclien, sclipot.nomcomer "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sclipot "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código Cliente|N|N|0|999999|sclipot|codclien|000000|S|"
    frmCom.Tag2 = "Nombre Cliente|T|N|||sclipot|nomclien||N|"
    frmCom.Tag3 = "Nombre Comercial|T|N|||sclipot|nomcomer||N|"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 40
    frmCom.Maxlen3 = 30

    frmCom.pConn = conAri

    frmCom.tabla = "sclipot"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "codclien"
    frmCom.TipoCP = "N"
    frmCom.Formulario = "Clientes Potenciales"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Clientes Potenciales"
    
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 4000
    
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaPreciosProveedor(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Proveedor|1200|;S|txtAux(1)|T|Nombre|4300|;S|txtAux(2)|T|Artículo|2100|;S|txtAux(3)|T|Descripción|4500|;"
    frmCom.CadenaConsulta = "SELECT slispr.codprove, sprove.nomprove, slispr.codartic, sartic.nomartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (slispr LEFT JOIN sprove on slispr.codprove = sprove.codprove) LEFT JOIN sartic ON slispr.codartic=sartic.codartic"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Cod. Proveedor|N|N|0|999999|slispr|codprove|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sprove|nomprove||N|"
    frmCom.Tag3 = "Cod. Artículo|T1|N|||slispr|codartic||S|"
    frmCom.Tag4 = "Nombre Artículo|T|N|||sartic|nomartic||N|"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 15
    frmCom.Maxlen4 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(slispr LEFT JOIN sprove on slispr.codprove = sprove.codprove) LEFT JOIN sartic ON slispr.codartic=sartic.codartic"
    frmCom.CampoCP = "slispr.codartic"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Precios Proveedor"
    
    frmCom.DatosADevolverBusqueda = "0|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 5100
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaDepartamentos(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
  
  '  frmCom.CadenaTots = "Código|sdirec|coddirec|N||30·Denominacion|sdirec|nomdirec|T||70·"
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|1305|;S|txtAux(1)|T|Denominacion|5095|;"
    frmCom.CadenaConsulta = "SELECT sdirec.coddirec, sdirec.nomdirec "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sdirec "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE true "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0||sdirec|coddirec|0000|S|"
    frmCom.Tag2 = "Descrip|T|N|||sdirec|nomdirec|||"
    frmCom.Maxlen1 = 2
    frmCom.Maxlen2 = 40

    frmCom.pConn = conAri

    frmCom.tabla = "sdirec"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "coddirec"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    
    If vParamAplic.HayDeparNuevo = 1 Then
        frmCom.Caption = "Departamentos"
    Else
        frmCom.Caption = "Direccion"
    End If
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 0
    
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaTarifasTaxi(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
  
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|1305|;S|txtAux(1)|T|Descripcion|5295|;"
    frmCom.CadenaConsulta = "SELECT slista_taxi.codtarifa, slista_taxi.descripcion "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM slista_taxi "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE true "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|slista_taxi|codtarifa|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||slista_taxi|descripcion|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 25

    frmCom.pConn = conAri

    frmCom.tabla = "slista_taxi"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "codtarifa"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    
    frmCom.Caption = "Tarifas Taxímetro"
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, -400
    
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaOrdenesReparacion(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
Dim tabla As String

    frmCom.CadenaTots = "S|txtAux(0)|T|Fecha|1605|;S|txtAux(1)|T|Albaran|1295|;S|txtAux(2)|T|Matricula|2300|;S|txtAux(3)|T|Facturado|1800|;"
    
    frmCom.CadenaConsulta = "select tt.fechaalb,tt.numalbar,tt.bombamarca , tt.facturado from "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & "(select scaalb.fechaalb,scaalb.numalbar,bombamarca , '' facturado from scaalb left join scaalb_eu "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " on scaalb.codtipom = scaalb_eu.codtipom and scaalb.numalbar = scaalb_eu.numalbar WHERE scaalb.codtipom='ALO'  and  "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " scaalb.codclien=" & DBSet(CodActual, "N")
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " Union"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " select s.fechaalb,s.numalbar,bombamarca, 'Si' facturado  from scafac f inner join scafac1 s  on f.codtipom=s.codtipom and f.numfactu=s.numfactu and f.fecfactu=s.fecfactu "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " left join scafac_eu l on  l.codtipom=s.codtipom and l.numfactu=s.numfactu and l.fecfactu=s.fecfactu and  s.codtipoa=l.codtipoa and s.numalbar=l.numalbar "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " where  s.codtipoa='ALO'  and codclien=" & DBSet(CodActual, "N")
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " order by 4,1,2) tt "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " where true "
    
    
    'If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Fecha|F|N|||tt|fechaalb|dd/mm/yyyy|S|"
    frmCom.Tag2 = "Albaran|N|N|||tt|numalbar|0000000||"
    frmCom.Tag3 = "Matrícula|T|N|||tt|bombamarca|||"
    frmCom.Tag4 = "Facturado|T|N|||tt|facturado|0000000||"
    frmCom.Maxlen1 = 10
    frmCom.Maxlen2 = 7
    frmCom.Maxlen3 = 20
    frmCom.Maxlen4 = 10
    

    frmCom.pConn = conAri

    tabla = "(select scaalb.fechaalb,scaalb.numalbar,bombamarca , '' facturado from scaalb left join scaalb_eu"
    tabla = tabla & " on scaalb.codtipom = scaalb_eu.codtipom and scaalb.numalbar = scaalb_eu.numalbar WHERE scaalb.codtipom='ALO'  and "
    tabla = tabla & " scaalb.codclien=" & DBSet(CodActual, "N")
    tabla = tabla & " Union"
    tabla = tabla & " select s.fechaalb,s.numalbar,bombamarca, 'Si' facturado  from scafac f inner join scafac1 s  on f.codtipom=s.codtipom and f.numfactu=s.numfactu and f.fecfactu=s.fecfactu"
    tabla = tabla & " left join scafac_eu l on  l.codtipom=s.codtipom and l.numfactu=s.numfactu and l.fecfactu=s.fecfactu and  s.codtipoa=l.codtipoa and s.numalbar=l.numalbar"
    tabla = tabla & " where  s.codtipoa='ALO'  and codclien=" & DBSet(CodActual, "N")
    tabla = tabla & " order by 4,1,2) tt "

    frmCom.tabla = tabla
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "numalbar"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    
    frmCom.Caption = "Ordenes de Reparación"
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    
    Redimensiona frmCom, 0
    
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaPedidosCompra(frmCom As frmBasico2, tabla As String, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Pedido|1100|;S|txtAux(1)|T|Fecha Ped|1500|;S|txtAux(2)|T|Proveedor|1200|;S|txtAux(3)|T|Nombre|5300|;"
    frmCom.CadenaConsulta = "SELECT " & tabla & ".numpedpr, " & tabla & ".fecpedpr, " & tabla & ".codprove, sprove.nomprove"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM " & tabla & " LEFT JOIN sprove ON " & tabla & ".codprove=sprove.codprove"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Pedido|N|S|0||" & tabla & "|numpedpr|0000000|S|"
    frmCom.Tag2 = "Fecha Pedido|F|N|||" & tabla & "|fecpedpr|dd/mm/yyyy|N|"
    frmCom.Tag3 = "Cod. Proveedor|N|N|0|999999|" & tabla & "|codprove|000000|N|"
    frmCom.Tag4 = "Nombre Proveedor|T|N|||sprove|nomprove||N|"
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 6
    frmCom.Maxlen4 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(" & tabla & " LEFT JOIN sprove ON " & tabla & ".codprove=sprove.codprove)"
    frmCom.CampoCP = tabla & ".numpedpr"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    If tabla = "scappr" Then
        frmCom.Caption = "Pedidos de Compra"
        frmCom.DatosADevolverBusqueda = "0|"
    Else
        frmCom.Caption = "Histórico Pedidos de Compra"
        frmCom.DatosADevolverBusqueda = "0|1|"
    End If
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 2100
    
    frmCom.Show vbModal

End Sub




Public Sub AyudaAlbaranesCompra(frmCom As frmBasico2, tabla As String, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Albaran|1500|;S|txtAux(1)|T|Fecha Alb|1500|;S|txtAux(2)|T|Proveedor|1200|;S|txtAux(3)|T|Nombre|5300|;"
    frmCom.CadenaConsulta = "SELECT " & tabla & ".numalbar, " & tabla & ".fechaalb, " & tabla & ".codprove, sprove.nomprove"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM " & tabla & " LEFT JOIN sprove ON " & tabla & ".codprove=sprove.codprove"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Albaran|T|N|0||" & tabla & "|numalbar||S|"
    frmCom.Tag2 = "Fecha Albaran|F|N|||" & tabla & "|fechaalb|dd/mm/yyyy|S|"
    frmCom.Tag3 = "Cod. Proveedor|N|N|0|999999|" & tabla & "|codprove|000000|S|"
    frmCom.Tag4 = "Nombre Proveedor|T|N|||sprove|nomprove||N|"
    frmCom.Maxlen1 = 10
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 6
    frmCom.Maxlen4 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(" & tabla & " LEFT JOIN sprove ON " & tabla & ".codprove=sprove.codprove)"
    frmCom.CampoCP = tabla & ".numalbar"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    If tabla = "scaalp" Then
        frmCom.Caption = "Albaranes de Compra"
    Else
        frmCom.Caption = "Histórico Albaranes de Compra"
    End If
    
    frmCom.DatosADevolverBusqueda = "0|1|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 2500
    
    frmCom.Show vbModal

End Sub

Public Sub AyudaFacturasCompra(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
Dim tabla As String
    
    tabla = "scafpc"


    frmCom.CadenaTots = "S|txtAux(0)|T|Factura|1500|;S|txtAux(1)|T|F.Factura|1500|;S|txtAux(2)|T|F.Recepción|1500|;S|txtAux(3)|T|Proveedor|1200|;S|txtAux(4)|T|Nombre|5300|;"
    frmCom.CadenaConsulta = "SELECT " & tabla & ".numfactu, " & tabla & ".fecfactu, " & tabla & ".fecrecep, " & tabla & ".codprove, sprove.nomprove"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM " & tabla & " LEFT JOIN sprove ON " & tabla & ".codprove=sprove.codprove"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Factura|T|N|0||" & tabla & "|numfactu||S|"
    frmCom.Tag2 = "Fecha Albaran|F|N|||" & tabla & "|fecfactu|dd/mm/yyyy|S|"
    frmCom.Tag3 = "Fecha Recepcion|F|N|||" & tabla & "|fecrecep|dd/mm/yyyy|S|"
    frmCom.Tag4 = "Cod. Proveedor|N|N|0|999999|" & tabla & "|codprove|000000|S|"
    frmCom.Tag5 = "Nombre Proveedor|T|N|||sprove|nomprove||N|"
    frmCom.Maxlen1 = 10
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 10
    frmCom.Maxlen4 = 6
    frmCom.Maxlen5 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(" & tabla & " LEFT JOIN sprove ON " & tabla & ".codprove=sprove.codprove)"
    frmCom.CampoCP = tabla & ".numfactu"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Histórico Facturas de Compra"
    
    frmCom.DatosADevolverBusqueda = "0|1|3|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 4000
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaCRMTipos(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
Dim tabla As String
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Denominación|6200|;"
    frmCom.CadenaConsulta = "SELECT scrmtipo.codigo, scrmtipo.denominacion "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scrmtipo "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0||scrmtipo|codigo|0000|S|"
    frmCom.Tag2 = "Denominacion|T|N|||scrmtipo|denominacion||N|"
    frmCom.Maxlen1 = 4
    frmCom.Maxlen2 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "scrmtipo"
    frmCom.CampoCP = "codigo"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Tipos Acciones Comerciales"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    
    frmCom.Show vbModal

End Sub

Public Sub AyudaMtoAccionesComerciales(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
Dim tabla As String
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Usuario|1000|;S|txtAux(1)|T|Fecha|2200|;S|txtAux(2)|T|Cliente|1000|;S|txtAux(3)|T|Nombre|4770|;"
    frmCom.CadenaConsulta = "SELECT scrmacciones.usuario, scrmacciones.fechora, scrmacciones.codclien, sclien.nomclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scrmacciones inner join sclien on scrmacciones.codclien = sclien.codclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Usuario|T|N|||scrmacciones|usuario||S|"
    frmCom.Tag2 = "Fecha/Hora|FH|N|||scrmacciones|fechora|dd/mm/yyyy hh:mm:ss|S|"
    frmCom.Tag3 = "Cliente|N|N|||scrmacciones|codclien|000000|S|"
    frmCom.Tag4 = "Nombre Cliente|T|N|||sclien|nomclien|||"
    
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 20
    frmCom.Maxlen3 = 6
    frmCom.Maxlen4 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "scrmacciones inner join sclien on scrmacciones.codclien = sclien.codclien"
    frmCom.CampoCP = "usuario"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Mantenimiento Acciones Comerciales"
    
    frmCom.DatosADevolverBusqueda = "0|1|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 2000
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaMantenimientos(frmCom As frmBasico2, Direc As String, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
Dim tabla As String
    
    tabla = "scaman"
    If EsAnulados Then tabla = "scamana"
    
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Cliente|1000|;S|txtAux(1)|T|Nombre|4700|;S|txtAux(2)|T|" & Direc & "|1000|;S|txtAux(3)|T|Nombre|4270|;S|txtAux(4)|T|NºMantenim.|1500|;"
    frmCom.CadenaConsulta = "SELECT " & tabla & ".codclien, sclien.nomclien, " & tabla & ".coddirec, sdirec.nomdirec, " & tabla & ".nummante "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (" & tabla & " inner join sclien on " & tabla & ".codclien = sclien.codclien) left join sdirec on " & tabla & ".coddirec = sdirec.coddirec "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código Cliente|N|N|0|999999|" & tabla & "|codclien|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sclien|nomclien|||"
    frmCom.Tag3 = "Cód. Dirección|N|S|0|999|" & tabla & "|coddirec|000|S|"
    frmCom.Tag4 = "Nombre|T|N|||sdirec|nomdirec|||"
    frmCom.Tag5 = "Nº Mantenimiento|T|N|||" & tabla & "|nummante||S|"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 30
    frmCom.Maxlen5 = 10
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(" & tabla & " inner join sclien on " & tabla & ".codclien = sclien.codclien) left join sdirec on " & tabla & ".coddirec = sdirec.coddirec "
    frmCom.CampoCP = "codclien"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Mantenimientos"
    
    frmCom.DatosADevolverBusqueda = "0|4|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 5500
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaMantenimientosAux(frmCom As frmBasico2, Titulo As String, Desc As String, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
Dim tabla As String
    
    tabla = "sdirec"
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Codigo|1000|;S|txtAux(1)|T|" & Desc & "|4700|;"
    frmCom.CadenaConsulta = "SELECT " & tabla & ".coddirec, sdirec.nomdirec "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM " & tabla
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|" & tabla & "|coddirec|000|S|"
    frmCom.Tag2 = Desc & "|T|N|||" & tabla & "|nomdirec|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = tabla
    frmCom.CampoCP = "coddirec"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = Titulo
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, -1300
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaNrosSerie(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
Dim tabla As String
    
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Nro.Serie|2000|;S|txtAux(1)|T|Artículo|2100|;S|txtAux(2)|T|Descripción|5000|;S|txtAux(3)|T|Código|800|;S|txtAux(4)|T|Tipo Artículo|3500|;"
    frmCom.CadenaConsulta = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic, sserie.codtipar, stipar.nomtipar "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (sserie inner join sartic on sserie.codartic = sartic.codartic) left join stipar on sserie.codtipar = stipar.codtipar "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Serie|T|N|||sserie|numserie||S|"
    frmCom.Tag2 = "Artículo|T|N|||sserie|codartic||S|"
    frmCom.Tag3 = "Descripcion|T|N|||sartic|nomartic|||"
    frmCom.Tag4 = "Tipo Artículo|T|N|||sserie|codtipar||N|"
    frmCom.Tag5 = "Descripcion|T|N|||stipar|nomtipar|||"
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 16
    frmCom.Maxlen3 = 30
    frmCom.Maxlen4 = 2
    frmCom.Maxlen5 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(sserie inner join sclien on sserie.codartic = sartic.codartic) left join stipar on sserie.codtipar = stipar.codtipar "
    frmCom.CampoCP = "sserie.numserie"
    frmCom.TipoCP = "T"
    frmCom.Formulario = "Series"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Números de Serie"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 6400
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaMantenimientoReports(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Codigo|800|;S|txtAux(1)|T|Descripción|5100|;S|txtAux(2)|T|Fichero RPT|3200|;"
    frmCom.CadenaConsulta = "SELECT scryst.codcryst, scryst.nomcryst, scryst.documrpt "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scryst "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código Documento|N|N|||scryst|codcryst|0000|S|"
    frmCom.Tag2 = "Descripción|T|N|||scryst|nomcryst|||"
    frmCom.Tag3 = "Fichero rpt|T|N|||scryst|documrpt|||"
    frmCom.Maxlen1 = 4
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "scryst"
    frmCom.CampoCP = "codcryst"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Tipos de Documentos"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 2100
    
    frmCom.Show vbModal

End Sub

Public Sub AyudaReparaciones(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
Dim tabla As String
    
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Nro.Serie|2000|;S|txtAux(1)|T|Artículo|2100|;S|txtAux(2)|T|Descripción|5000|;S|txtAux(3)|T|NºReparación|1800|;"
    frmCom.CadenaConsulta = "SELECT scarep.numserie, scarep.codartic, sartic.nomartic, scarep.numrepar "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scarep inner join sartic on scarep.codartic = sartic.codartic  "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Serie|T|S|||scarep|numserie||N|"
    frmCom.Tag2 = "Cod. Artículo|T|N|||scarep|codartic||N|"
    frmCom.Tag3 = "Descripcion|T|N|||sartic|nomartic|||"
    frmCom.Tag4 = "Nº Reparación|N|S|0|9999999|scarep|numrepar|0000000|S|"
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 16
    frmCom.Maxlen3 = 30
    frmCom.Maxlen4 = 7
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "scarep inner join sartic on scarep.codartic = sartic.codartic "
    frmCom.CampoCP = "scarep.numserie"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Reparaciones"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 3900
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaAvisos(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
Dim tabla As String
    
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Nº Aviso|1000|;S|txtAux(1)|T|Fecha Aviso|1500|;S|txtAux(2)|T|Cliente|1000|;S|txtAux(3)|T|Nombre|5000|;"
    frmCom.CadenaConsulta = "SELECT scaavi.numaviso, scaavi.fechaavi, scaavi.codclien, sclien.nomclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scaavi inner join sclien on scaavi.codclien = sclien.codclien  "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Nº Aviso|N|S|0||scaavi|numaviso|0000000|S|"
    frmCom.Tag2 = "Fecha Aviso|F|N|||scaavi|fechaavi|dd/mm/yyyy|N|"
    frmCom.Tag3 = "Cliente|N|N|0|999999|scaavi|codclien|000000|N|"
    frmCom.Tag4 = "Nombre cliente|T|N|||sclien|nomclien|||"
    frmCom.Maxlen1 = 7
    frmCom.Maxlen2 = 10
    frmCom.Maxlen3 = 6
    frmCom.Maxlen4 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "scaavi inner join sclien on scaavi.codclien = sclien.codclien "
    frmCom.CampoCP = "scaavi.numaviso"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Avisos de Clientes"
    
    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 1500
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaFrecuencias(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
Dim tabla As String
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Codigo|900|;S|txtAux(1)|T|Cliente|4500|;S|txtAux(2)|T|Dpto|600|;S|txtAux(3)|T|Nombre|3500|;S|txtAux(4)|T|Expediente|2000|;S|txtAux(5)|T|F.Inicio|1400|;"
    frmCom.CadenaConsulta = "SELECT scafre.codclien,sclien.nomclien,scafre.coddirec,sdirec.nomdirec,scafre.numexped,scafre.fechaini "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (scafre inner join sclien on scafre.codclien = sclien.codclien) left join sdirec on scafre.codclien = sdirec.codclien and scafre.coddirec = sdirec.coddirec "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Cod.Clien|N|N|||scafre|codclien|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sclien|nomclien||N|"
    frmCom.Tag3 = "Dpto|N|N|||scafre|coddirec|000|S|"
    frmCom.Tag4 = "Nombre dpto|T|N|||sdirec|nomdirec|||"
    frmCom.Tag5 = "Numexp|T|N|||scafre|numexped||S|"
    frmCom.Tag6 = "Fecha inicio|F|N|||scafre|fechaini|dd/mm/yyyy|S|"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 30
    frmCom.Maxlen3 = 3
    frmCom.Maxlen4 = 30
    frmCom.Maxlen5 = 15
    frmCom.Maxlen6 = 10
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "(scafre inner join sclien on scafre.codclien = sclien.codclien) left join sdirec on scafre.codclien = sdirec.codclien and scafre.coddirec = sdirec.coddirec"
    frmCom.CampoCP = "scafre.codclien"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Frecuencias"
    
    frmCom.DatosADevolverBusqueda = "0|2|4|5|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 5900
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaTelNombreGrupo(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
Dim tabla As String
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Codigo|900|;S|txtAux(1)|T|Nombre Grupo|4500|;"
    frmCom.CadenaConsulta = "SELECT tel_desc_nombres_grupo.grupo,tel_desc_nombres_grupo.nombre "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM tel_desc_nombres_grupo "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Codigo|N|N|||tel_desc_nombres_grupo|grupo|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||tel_desc_nombres_grupo|nombre||N|"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 35
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "tel_desc_nombres_grupo"
    frmCom.CampoCP = "tel_desc_nombres_grupo.grupo"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Nombre de Grupos"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, -1600
    
    frmCom.Show vbModal

End Sub



Public Sub AyudaTelefonos(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
Dim tabla As String
    frmCom.CadenaTots = "S|txtAux(0)|T|Codigo|1100|;S|txtAux(1)|T|Nombre|5300|;S|txtAux(2)|T|Teléfono|1500|;S|txtAux(3)|T|Operador|1500|;"
    frmCom.CadenaConsulta = "SELECT sclien.codclien, sclien.nomclien, sclientfno.idtelefono, stfnooperador.nombre "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (sclien inner join sclientfno on sclien.codclien=sclientfno.codclien) inner join stfnooperador on stfnooperador.codoperador=sclientfno.operador "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Codigo|N|N|||sclien|codclien|000000|S|"
    frmCom.Tag2 = "Nombre|T|N|||sclien|nomclien||N|"
    frmCom.Tag3 = "Teléfono|T|N|||sclientfno|IdTelefono||N|"
    frmCom.Tag4 = "Operador|T|N|||stfnooperador|nombre||N|"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 35
    frmCom.Maxlen3 = 15
    frmCom.Maxlen4 = 35
    
    frmCom.pConn = conAri
    
    frmCom.tabla = "sclien inner join sclientfno on sclien.codclien=sclientfno.codclien) inner join stfnooperador on stfnooperador.codoperador=sclientfno.operador "
    frmCom.CampoCP = "sclientfno.idtelefono"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Teléfonos"
    
    frmCom.DatosADevolverBusqueda = "0|1|2|3|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 2400
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaContadoresAgua(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Contador|1700|;S|txtAux(1)|T|Cliente|1000|;S|txtAux(2)|T|Nombre|5500|;"
    frmCom.CadenaConsulta = "SELECT aguacontadores.contador, aguacontadores.codclien, sclien.nomclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM aguacontadores inner join sclien on aguacontadores.codclien=sclien.codclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere

    frmCom.Tag1 = "Numero contador|T|N|||aguacontadores|contador||S|"
    frmCom.Tag2 = "Cliente|N|N|||aguacontadores|codclien|000000||"
    frmCom.Tag3 = "Nombre|T|N|||sclien|nomclien||N|"
    frmCom.Maxlen1 = 15
    frmCom.Maxlen2 = 6
    frmCom.Maxlen3 = 35

    frmCom.pConn = conAri

    frmCom.tabla = "aguacontadores inner join sclien on aguacontadores.codclien=sclien.codclien "
    frmCom.CampoCP = "aguacontadores.contador"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Contadores"

    frmCom.DatosADevolverBusqueda = "0|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual

    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui

    Redimensiona frmCom, 1200

    frmCom.Show vbModal

End Sub


Public Sub AyudaContadoresAguaMod(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean, Optional EsAnulados As Boolean)
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Cliente|1000|;S|txtAux(1)|T|Nombre|5500|;S|txtAux(2)|T|Contador|1700|;S|txtAux(3)|T|Facturar|1000|;"
    frmCom.CadenaConsulta = "SELECT aguacontadores.codclien, sclien.nomclien, aguacontadores.contador, if(coalesce(aguacontadoresconce.Facturar,0)=1,""Si"","""")  "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM (aguacontadores inner join sclien on aguacontadores.codclien=sclien.codclien) left join aguacontadoresconce ON  aguacontadores.Contador = aguacontadoresconce.Contador  "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere

    frmCom.Tag1 = "Cliente|N|N|||aguacontadores|codclien|000000||"
    frmCom.Tag2 = "Nombre|T|N|||sclien|nomclien||N|"
    frmCom.Tag3 = "Numero contador|T|N|||aguacontadores|contador||S|"
    frmCom.Tag4 = "Facturar|T|N|||aguacontadoresconce|facturar|||"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 35
    frmCom.Maxlen3 = 15
    frmCom.Maxlen4 = 10

    frmCom.pConn = conAri

    frmCom.tabla = "(aguacontadores inner join sclien on aguacontadores.codclien=sclien.codclien) left join aguacontadoresconce ON  aguacontadores.Contador = aguacontadoresconce.Contador  "
    frmCom.CampoCP = "aguacontadores.contador"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Contadores"

    frmCom.DatosADevolverBusqueda = "2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual

    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui

    Redimensiona frmCom, 2200

    frmCom.Show vbModal

End Sub



Public Sub AyudaCalibres(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
  
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|1305|;S|txtAux(1)|T|Descripcion|5495|;"
    frmCom.CadenaConsulta = "SELECT aguacalibre.codcalibre, aguacalibre.nomcalibre "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM aguacalibre "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE true "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|aguacalibre|codcalibre|000|S|"
    frmCom.Tag2 = "Nombre|T|N|||aguacalibre|nomcalibre|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 25

    frmCom.pConn = conAri

    frmCom.tabla = "aguacalibre"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "codcalibre"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    
    frmCom.Caption = "Calibres"
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, -190
    
    
    frmCom.Show vbModal
End Sub


Public Sub AyudaHistoricoInventario(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
  
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|2085|;S|txtAux(1)|T|Descripcion|5495|;"
    frmCom.CadenaConsulta = "SELECT distinct shinve.codartic, sartic.nomartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM shinve left join sartic on shinve.codartic = sartic.codartic "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE true "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|T|N|||shinve|codartic||S|"
    frmCom.Tag2 = "Nombre|T|N|||sartic|nomartic|||"
    frmCom.Maxlen1 = 16
    frmCom.Maxlen2 = 40

    frmCom.pConn = conAri

    frmCom.tabla = "shinve left join sartic on shinve.codartic = sartic.codartic"
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CampoCP = "shinve.codartic"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    
    frmCom.Caption = "Historico Inventario"
    
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 590
    
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaFacturasCliente(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
Dim tabla As String
    
    tabla = "scafac"

    frmCom.CadenaTots = "S|txtAux(0)|T|Tipo|800|;S|txtAux(1)|T|Factura|1000|;S|txtAux(2)|T|F.Factura|1500|;S|txtAux(3)|T|Cliente|900|;S|txtAux(4)|T|Nombre|5300|;"
    frmCom.CadenaConsulta = "SELECT codtipom, numfactu, fecfactu, codclien, nomclien "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM scafac "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Tipo Movimiento|T|N|||" & tabla & "|codtipom||S|"
    frmCom.Tag2 = "Nº Factura|T|N|0||" & tabla & "|numfactu||S|"
    frmCom.Tag3 = "Fecha Factura|F|N|||" & tabla & "|fecfactu|dd/mm/yyyy|S|"
    frmCom.Tag4 = "Cliente|N|N|||" & tabla & "|codclien|000000||"
    frmCom.Tag5 = "Nombre Cliente|T|N|||" & tabla & "|nomclien|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 7
    frmCom.Maxlen3 = 10
    frmCom.Maxlen4 = 6
    frmCom.Maxlen5 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = tabla
    frmCom.CampoCP = tabla & ".numfactu"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Histórico Facturas de Cliente"
    
    frmCom.DatosADevolverBusqueda = "0|1|2|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 2500
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaDireccionesEnvio(frmCom As frmBasico2, Titulo As String, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
Dim tabla As String
    
    
    frmCom.CadenaTots = "S|txtAux(0)|T|Codigo|1000|;S|txtAux(1)|T|Descripción|4700|;"
    frmCom.CadenaConsulta = "SELECT sdirenvio.coddiren, sdirenvio.nomdiren "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM sdirenvio "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|999|sdirenvio|coddiren|000|S|"
    frmCom.Tag2 = "Descripción|T|N|||sdirenvio|nomdiren|||"
    frmCom.Maxlen1 = 3
    frmCom.Maxlen2 = 30
    
    frmCom.pConn = conAri
    
    frmCom.tabla = tabla
    frmCom.CampoCP = "coddiren"
    frmCom.TipoCP = "T"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = Titulo
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = 0
    If CodActual <> "" Then frmCom.CodigoActual = CodActual
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, -1300
    
    frmCom.Show vbModal

End Sub


Public Sub AyudaUnidadesNegocio(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)
    frmCom.CadenaTots = "S|txtAux(0)|T|Código|905|;S|txtAux(1)|T|Descripción|5095|;"
    frmCom.CadenaConsulta = "SELECT IdUnidad, Nombre "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM unidadesnegocio"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código de Banco Propio|N|N|0|9999|unidadesnegocio|IdUnidad|0000|S|"
    frmCom.Tag2 = "Denominación|T|N|||unidadesnegocio|Nombre||N|"
    frmCom.Maxlen1 = 4
    frmCom.Maxlen2 = 30
    frmCom.pConn = conAri
    
    frmCom.tabla = "unidadesnegocio"
    frmCom.CampoCP = "IdUnidad"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Unidades de Negocio"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = ""
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, -1000
    
    frmCom.Show vbModal
End Sub

Public Sub AyudaAsociadosGesSoc(frmCom As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional SinAvanzada As Boolean)

    frmCom.CadenaTots = "S|txtAux(0)|T|Código|1005|;S|txtAux(1)|T|Descripción|5595|;S|txtAux(2)|T|NIF|1800|;"
    frmCom.CadenaConsulta = "SELECT Idasoc, nomlargo, nif "
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " FROM asociados"
    frmCom.CadenaConsulta = frmCom.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmCom.CadenaConsulta = frmCom.CadenaConsulta & " and " & cWhere
    
    frmCom.Tag1 = "Código|N|N|0|99999|asociados|idasoc|000000|S|"
    frmCom.Tag2 = "Nombre Trabajador|T|N|||asociados|NomLargo||N|"
    frmCom.Tag3 = "NIF|T|N|||asociados|nif|||"
    frmCom.Maxlen1 = 6
    frmCom.Maxlen2 = 40
    frmCom.Maxlen3 = 9
    frmCom.pConn = conAri
    
    frmCom.tabla = "asociados"
    frmCom.CampoCP = "idasoc"
    frmCom.TipoCP = "N"
    If SinAvanzada Then frmCom.DeConsulta = True
    frmCom.Caption = "Asociados"
    
    frmCom.DatosADevolverBusqueda = "0|1|"
    frmCom.CodigoActual = ""
    
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmCom.DataGrid1.Height = 7420
    frmCom.DataGrid1.Top = 870
    frmCom.FrameBotonGnral.visible = True
    frmCom.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmCom, 1400
    
    frmCom.Show vbModal
End Sub



Public Sub AyudaFacturasAnyadeAlbaranCostes(frmAlb As frmBasico2, codClien As String, SinAvanzada As Boolean)
    frmAlb.CadenaTots = "S|txtAux(0)|T|Tipo|1105|;S|txtAux(1)|T|Albaran.|1495|;S|txtAux(2)|T|Fecha|1350|;S|txtAux(3)|T|Referenc|5500|;S|txtAux(4)|T|Observa1|4500|;"
    frmAlb.CadenaConsulta = "SELECT scaalb.codtipom, scaalb.numalbar, scaalb.fechaalb, scaalb.referenc,scaalb.observa01"
    frmAlb.CadenaConsulta = frmAlb.CadenaConsulta & " FROM scaalb"
    frmAlb.CadenaConsulta = frmAlb.CadenaConsulta & " WHERE codclien=" & RecuperaValor(codClien, 1) & " AND scaalb.codtipom <> 'ALV' "
    frmAlb.CadenaConsulta = frmAlb.CadenaConsulta & " AND NOT (codtipom,numalbar) IN (select codtipom,numalbar FROM slialb)"
    
    
    frmAlb.Tag1 = "Tipo|T|N|||scaalb|codtipom|||"
    frmAlb.Tag2 = "Albaran|N|N|0|9999999|scaalb|numalbar|000000|S|"
    frmAlb.Tag3 = "Fecha|F|N|||scaalb|fechaalb|dd/mm/yyyy||"
    frmAlb.Tag4 = "Referencia|T|N|||scaalb|referenc|||"
    frmAlb.Tag5 = "Observa.|T|N|||scaalb|observa01|||"
    
    frmAlb.Maxlen1 = 6
    frmAlb.Maxlen2 = 10
    frmAlb.Maxlen3 = 10
    frmAlb.Maxlen4 = 35
    frmAlb.Maxlen5 = 35
    
    frmAlb.pConn = conAri

    frmAlb.tabla = "scaalb"
    frmAlb.DatosADevolverBusqueda = "0|1|"
    frmAlb.CampoCP = "codclien"
    
    frmAlb.TipoCP = "N"
    'frmAlb.Formulario = "Albaranes cliente " & RecuperaValor(Codclien, 2)
    frmAlb.DeConsulta = False
    
    frmAlb.Caption = "Albaranes cliente " & RecuperaValor(codClien, 2)
    
    frmAlb.CodigoActual = 0
'    If CodActual <> "" Then
'    frmAlb.CodigoActual = CodActual
    
    '[Monica]17/04/2018: añadimos en este caso los botones de busqueda
    frmAlb.DataGrid1.Height = 7420
    frmAlb.DataGrid1.Top = 870
    frmAlb.FrameBotonGnral.visible = True
    frmAlb.FrameBotonGnral.Enabled = True
    ' hasta aqui
    
    Redimensiona frmAlb, 7500
    
    
    frmAlb.Show vbModal
End Sub








Private Sub Redimensiona(frmBas As frmBasico2, Cant As Integer)
    frmBas.Width = frmBas.Width + Cant
    frmBas.DataGrid1.Width = frmBas.DataGrid1.Width + Cant
    frmBas.cmdAceptar.Left = frmBas.cmdAceptar.Left + Cant
    frmBas.cmdCancelar.Left = frmBas.cmdCancelar.Left + Cant
    frmBas.cmdRegresar.Left = frmBas.cmdRegresar.Left + Cant
End Sub


