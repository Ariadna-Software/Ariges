VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPreciosFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Modulo para Calcula el Precio a Aplicar a las lineas de una Oferta, Pedido o Albaran
' Y los descuentos correspondientes si se aplica descuento
'TABLAS: slista, sprees, sdtofm

'------------------------------------------------
'Variables del modulo CPreciosFact
'------------------------------------------------

Private mvarDtoPermitido As Boolean
Private mvarCodArtic As String
Private mvarCodLista As String
Private mvarCodClien As String
Private mvarCodFamia As String
Private mvarCodMarca As String
Private mvarCodActividad As String
Private mvarAplicaPromo As Boolean  'interno. Si no tiene la marca de promociones NO se le aplicara ninguna


Private mvarDescuento1 As String
Private mvardescuento2 As String

Dim FechaMenor As Boolean
'------------------------------------------------
'Propiedades del modulo CPreciosFact
'------------------------------------------------

Public Property Let DtoPermitido(ByVal vData As Boolean)
     mvarDtoPermitido = vData
End Property

Public Property Get DtoPermitido() As Boolean
     DtoPermitido = mvarDtoPermitido
End Property

Public Property Let CodigoArtic(ByVal vData As String)
     mvarCodArtic = vData
End Property

Public Property Get CodigoArtic() As String
     CodigoArtic = mvarCodArtic
End Property

Public Property Let CodigoLista2(ByVal vData As String)
     mvarCodLista = vData
End Property

Public Property Get CodigoLista2() As String
     CodigoLista2 = mvarCodLista
End Property



Public Property Let CodActividad(ByVal vData As String)
     mvarCodActividad = vData
End Property

Public Property Get CodActividad() As String
     CodActividad = mvarCodActividad
End Property


Public Property Let CodigoClien(ByVal vData As String)
     mvarCodClien = vData
End Property

Public Property Get CodigoClien() As String
     CodigoClien = mvarCodClien
End Property

Public Property Let CodigoFamia(ByVal vData As String)
     mvarCodFamia = vData
End Property

Public Property Get CodigoFamia() As String
     CodigoFamia = mvarCodFamia
End Property

Public Property Let CodigoMarca(ByVal vData As String)
     mvarCodMarca = vData
End Property

Public Property Get CodigoMarca() As String
     CodigoMarca = mvarCodMarca
End Property

Public Property Let Descuento1(ByVal vData As String)
     mvarDescuento1 = vData
End Property

Public Property Get Descuento1() As String
     Descuento1 = mvarDescuento1
End Property


Public Property Let Descuento2(ByVal vData As String)
     mvardescuento2 = vData
End Property

Public Property Get Descuento2() As String
     Descuento2 = mvardescuento2
End Property

'------------------------------------------------
'Metodos del modulo CPreciosFact
'------------------------------------------------

'Hemos desdoblado la function
'Para no molestar a nadie si no lleva lo de precio minimo

Public Function ObtenerPrecio(PCaja As Boolean, FechaOfe As String, OrigP As String, Comision As String) As String
    If vParamAplic.PrecioMinimo Then
        'ObtenerPrecio
        ObtenerPrecio = ObtenerPrecioMinimo(PCaja, FechaOfe, OrigP, Comision)
    Else
        ObtenerPrecio = ObtenerPrecioNormal(PCaja, FechaOfe, OrigP, Comision)
    End If
End Function

Private Function ObtenerPrecioNormal(PCaja As Boolean, FechaOfe As String, OrigP As String, Comision As String) As String
'PCaja -> IN: si vale true se obiene el precio de unidad si se vende por caja sino se obtiene precio
'         OUT: si sale con precio de pvp del articulo se pone a False, independientemente de que se
'              pueda vender por cajas o no
'FechaOfe -> IN: Fecha de la Oferta para comprobar si la oferta esta en los periodos de promocion, etc...
'OrigP -> OUT: Aqui se devolvera con que valor precio se sale de la función, e.d.:
'               P=sale con precio de Promoción
'               E=sale con Precio Especial
'               etc....
'RETURN: valor del precio
Dim PrecioFact As String
Dim Dto1Fact As String
Dim Dto2Fact As String
Dim Aux As String
    On Error GoTo EObtenerPre
    Comision = ""
    'COMPROBAR PROMOCIONES
    '==============================================================
    'Buscar Precio en la tabla spromo si encuentro valor salgo con ese precio
    PrecioFact = ObtenerPromocion(FechaOfe, PCaja)
    If PrecioFact <> "" Then 'Ha obtenido el Precio
        If DtoPermitido Then
        'Buscar en la tabla sdtofm los descuentos
            Dto1Fact = ObtenerDescuentosN(FechaOfe, Dto2Fact, PCaja, Comision)
        End If
        OrigP = "P" 'Indica que sale con precio de Promocion
        
    Else 'No encuentra Precio en spromo y Buscar en Sprees(Tabla Precios Especiales)
        'COMPROBAR PRECIO ESPECIAL
        '===========================================================================
        PrecioFact = ObtenerPrecioEsp2(FechaOfe, PCaja, Dto1Fact, Comision)
        Dto2Fact = "0"
        If PrecioFact <> "" Then 'Ha obtenido el precio
            If PrecioFact = "0" Then
                If vParamAplic.NumeroInstalacion = vbFenollar Then
                    Aux = ObtenerTarifaArt(FechaOfe, PCaja)
                    If Aux <> "" Then PrecioFact = Aux
                End If
            End If
            If DtoPermitido And Dto1Fact = "" Then
            'Buscar en la tabla sdtofm los descuentos
                Dto1Fact = ObtenerDescuentosN(FechaOfe, Dto2Fact, PCaja, Comision)
            End If
            OrigP = "E" 'Indica que sale con Precio Especial
        Else 'No encuentra precio en sprees y Buscar en slista(Tabla Tarifas Articulos)
            'COMPROBAR TARIFAS ARTICULOS
            '=======================================================================
            PrecioFact = ObtenerTarifaArt(FechaOfe, PCaja)
            If PrecioFact <> "" Then 'Ha obtenido el precio
                If DtoPermitido Then
                    'Buscar en la tabla sdtofm los descuentos
                    Dto1Fact = ObtenerDescuentosN(FechaOfe, Dto2Fact, PCaja, Comision)
                End If
                OrigP = "T"
            Else 'Buscar en sartic el precio del articulo (preciove)
                PrecioFact = ObtenerPrecioVenta
                PCaja = False
                OrigP = "A"
                'Nuevo###. David 14 Mayo 2008
                'Permitimos descuento
                DtoPermitido = True
                                  
                Dto1Fact = ObtenerDescuentosN(FechaOfe, Dto2Fact, PCaja, Comision)
                
            End If
            
            If vParamAplic.NumeroInstalacion = vbFenollar Then
                If Dto1Fact = "" Then
                    
                    Dto1Fact = DescuentosPorCatalogo
                End If
            End If
            
        End If
    End If
    ObtenerPrecioNormal = ComprobarCero(PrecioFact)
    Descuento1 = ComprobarCero(Dto1Fact)
    Descuento2 = ComprobarCero(Dto2Fact)
    Exit Function
    
EObtenerPre:
    MuestraError Err.Number, "Obtener Precio.", Err.Description
End Function

'Aqui, aunque obtengamos un valor, seguiremos calculando para que obtenga el minimo
Private Function ObtenerPrecioMinimo(PCaja As Boolean, FechaOfe As String, OrigP_ As String, Comision2 As String) As String
Dim ElPrecio2 As String
Dim PrecioMinConDto As String
Dim Pre As String
'Dim OrigP As String
Dim Dto1Ft As String
Dim Dto2Ft As String
Dim DtoPermit As Boolean
Dim ComisionMinima As String


Dim PrecioFact As String
Dim Dto1Fact_ As String
Dim Dto2Fact_ As String
Dim Aux As String
Dim auxCom As String

    On Error GoTo EObtenerPre
    ElPrecio2 = "99999999999" 'Empezamos a este nivel
    PrecioMinConDto = ElPrecio2
    ComisionMinima = "0"
    
    
    ' 4 Junio 2015
    'Tania.
    'En promociones, la marca de dtopermitido significa que NO va a buscar los dtos
    ' , pero SI que pdoria modificarlos en las lineas de albra-ped-ofert
    
    
    
    
    
    
    'COMPROBAR PROMOCIONES
    '==============================================================
    'Buscar Precio en la tabla spromo si encuentro valor salgo con ese precio
    PrecioFact = ObtenerPromocion(FechaOfe, PCaja)
    If PrecioFact <> "" Then 'Ha obtenido el Precio
        Dto1Fact_ = ""
        Dto2Fact_ = ""
        If DtoPermitido Then
            'Buscar en la tabla sdtofm los descuentos
            'Esta comentado. Ver arriba la explicacion
            'Dto1Fact_ = ObtenerDescuentosN(FechaOfe, Dto2Fact_, PCaja, auxCom)
                        
        End If
        OrigP_ = "P" 'Indica que sale con precio de Promocion
        ElPrecio2 = PrecioFact
        PrecioMinConDto = CalcularImporte("1", ElPrecio2, Dto1Fact_, Dto2Fact_, vParamAplic.TipoDtos)
        
        DtoPermit = DtoPermitido
        Dto1Ft = Dto1Fact_
        Dto2Ft = Dto2Fact_
        PrecioFact = ""  'Para que entre en el siguiente
        If auxCom <> "" Then ComisionMinima = auxCom
    End If
    
    
    'No encuentra Precio en spromo y Buscar en Sprees(Tabla Precios Especiales)
    'COMPROBAR PRECIO ESPECIAL
    '===========================================================================
    Dto1Fact_ = ""
    Dto2Fact_ = "0"
    auxCom = ""
    PrecioFact = ObtenerPrecioEsp2(FechaOfe, PCaja, Dto1Fact_, auxCom)
    
    
    If PrecioFact <> "" Then 'Ha obtenido el precio
        If auxCom = "" Then auxCom = "0"
        
        If ComisionMinima = "0" Then
            ComisionMinima = auxCom
        Else
            If CCur(auxCom) < CCur(ComisionMinima) Then
                ComisionMinima = auxCom
            End If
        End If
        
        
        If DtoPermitido And Dto1Fact_ = "" Then
        'Buscar en la tabla sdtofm los descuentos
            Dto1Fact_ = ObtenerDescuentosN(FechaOfe, Dto2Fact_, PCaja, auxCom)
        Else
            'Dto1Fact_ = ""
            Dto2Fact_ = ""
        End If
        'OrigP = "E" 'Indica que sale con Precio Especial
        Pre = CalcularImporte("1", PrecioFact, Dto1Fact_, Dto2Fact_, vParamAplic.TipoDtos)
        
        'Veremos el minimo
        If CCur(PrecioMinConDto) > CCur(Pre) Then
            'Si, es menor
            DtoPermit = DtoPermitido
            OrigP_ = "E"
            ElPrecio2 = PrecioFact
            Dto1Ft = Dto1Fact_
            Dto2Ft = Dto2Fact_
            PrecioMinConDto = Pre
        End If
        
        
        PrecioFact = ""  'Para que entre en el siguiente
        
    End If
    
    
        
    'COMPROBAR TARIFAS ARTICULOS
    '=======================================================================
    auxCom = ""
    
    PrecioFact = ObtenerTarifaArt(FechaOfe, PCaja)
    If PrecioFact <> "" Then 'Ha obtenido el precio
        If DtoPermitido Then
            'Buscar en la tabla sdtofm los descuentos
            Dto1Fact_ = ObtenerDescuentosN(FechaOfe, Dto2Fact_, PCaja, auxCom)
        Else
            Dto1Fact_ = ""
            Dto2Fact_ = ""
        End If
        'OrigP = "T"
        Pre = CalcularImporte("1", PrecioFact, Dto1Fact_, Dto2Fact_, vParamAplic.TipoDtos)
        
        'Veremos el minimo
        If CCur(PrecioMinConDto) > CCur(Pre) Then
            'Si, es menor
            DtoPermit = DtoPermitido
            OrigP_ = "T"
            ElPrecio2 = PrecioFact
            Dto1Ft = Dto1Fact_
            Dto2Ft = Dto2Fact_
            PrecioMinConDto = Pre
        End If
        
        'Comision
        If auxCom <> "" Then
            If ComisionMinima = "0" Then
                ComisionMinima = auxCom
            Else
                If CCur(auxCom) < CCur(ComisionMinima) Then
                    ComisionMinima = auxCom
                End If
            End If
        End If
        
        
        
     End If
        
            
     'El precio venta
     
     PrecioFact = ObtenerPrecioVenta
     PCaja = False
     'Nuevo###. David 14 Mayo 2008
     Dto1Fact_ = 0
     Dto2Fact_ = 0
     Pre = CalcularImporte("1", PrecioFact, Dto1Fact_, Dto2Fact_, vParamAplic.TipoDtos)
     If CCur(PrecioMinConDto) > CCur(Pre) Then
        'DtoPermitido = True  'pq de tarifa sii permite dot
        DtoPermit = True
         'Si, es menor
         OrigP_ = "A"
         ElPrecio2 = PrecioFact
         Dto1Ft = Dto1Fact_
         Dto2Ft = Dto2Fact_
     End If
                
    DtoPermitido = DtoPermit
    ObtenerPrecioMinimo = ComprobarCero(ElPrecio2)
    Descuento1 = ComprobarCero(Dto1Ft)
    Descuento2 = ComprobarCero(Dto2Ft)
    If ComisionMinima = "0" Then ComisionMinima = ""
    Comision2 = ComisionMinima
    Exit Function
    
EObtenerPre:
    MuestraError Err.Number, "Obtener Precio.", Err.Description
End Function





Private Function ObtenerPromocion(FechaOfe As String, PorCaja As Boolean) As String
'Comprueba si existe una promocion en la tabla: spromo (Promociones Tarifas)
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Precio As String

    On Error GoTo EObtenerPromo

    Precio = ""
    
    If Not mvarAplicaPromo Then Exit Function
    
    
    'Obtener Precio Actual
    SQL = "SELECT * FROM spromo "
    SQL = SQL & " WHERE codartic=" & DBSet(CodigoArtic, "T") & " AND codlista=" & CodigoLista2
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
    'Existe una promocion para ese articulo y tarifa, comprobar fechas
        If EntreFechas(DBLet(RS!FechaIni), FechaOfe, DBLet(RS!FechaFin)) Then
        'Dentro Periodo Promocion Actual
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precioa1))
                If Precio = "0" Then Precio = DBLet(RS!precioac)
            Else
                Precio = DBLet(RS!precioac)
            End If
            If Precio = "0" Then Precio = ""
        ElseIf DBLet(RS!fechain1) = "" And DBLet(RS!fechafi1) = "" Then
            Precio = ""
        ElseIf EntreFechas(DBLet(RS!fechain1), FechaOfe, DBLet(RS!fechafi1)) Then
        'DEntro Periodo Promocion Nueva
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precion1))
                If Precio = "0" Then Precio = DBLet(RS!precionu)
            Else
                Precio = ComprobarCero(DBLet(RS!precionu))
            End If
        End If
        DtoPermitido = CBool(RS!dtopermi)
    Else 'No existe Promocion salir y pasar al siguiente paso
        Precio = ""
        DtoPermitido = False
    End If
    RS.Close
    Set RS = Nothing
    
    ObtenerPromocion = Precio
    Exit Function
    
EObtenerPromo:
   MuestraError Err.Number, "Precio Promoción.", Err.Description
End Function


Private Function ObtenerDescuentosOLD(FechaOfe As String, Dto2Fact, PorCaja As Boolean) As String
Dim RS As ADODB.Recordset
Dim SQL As String
Dim CampoDto1 As String, campoDto2 As String
Dim vCodMarca As String
    
    '- obtener la familia y la marca del articulo
    vCodMarca = "Codmarca"
    CodigoFamia = DevuelveDesdeBDNew(conAri, "sartic", "Codfamia", "codartic", CodigoArtic, "T", vCodMarca)
    CodigoMarca = vCodMarca
    
    'ANTE NOVIEMBRE 2013
'    If PorCaja Then 'Obtener Descuentos Por Caja
'        CampoDto1 = "dtocaja1"
'        campoDto2 = "dtocaja2"
'    Else 'Obtener Descuentos por Unidad
'        CampoDto1 = "dtoline1"
'        campoDto2 = "dtoline2"
'    End If
    
    CampoDto1 = "dtoline1"
    campoDto2 = "dtoline2"
    If PorCaja Then 'Obtener Descuentos Por Caja
        If vParamAplic.NumeroInstalacion <> vbHerbelca Then  'PARA TODOS menos para herbelca
            CampoDto1 = "dtocaja1"
            campoDto2 = "dtocaja2"
        End If
    End If
    
    
    
    
    'Obtener Descuento:
    '1- buscar descuento para la familia y marca del articulo
    SQL = "SELECT " & CampoDto1 & ", " & campoDto2 & ", codactiv FROM sdtofm "
    SQL = SQL & " WHERE codclien=" & CodigoClien
    SQL = SQL & " AND (fechadto<= '" & Format(FechaOfe, FormatoFecha) & "')"
    SQL = SQL & " AND ((codfamia=" & CodigoFamia & " AND codmarca=" & CodigoMarca & ")"
    SQL = SQL & " OR (codfamia=" & CodigoFamia & " AND codmarca=0)"
    SQL = SQL & " OR (codfamia=0 AND codmarca=" & CodigoMarca & ")"
    SQL = SQL & " OR (codfamia=0 AND codmarca=0))"
    If vParamAplic.OrdenDtos = 0 Then
        SQL = SQL & " ORDER BY codfamia DESC, codmarca DESC"
    Else
        SQL = SQL & " ORDER BY codmarca DESC,codfamia DESC"
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
    'Existe el descuento, y salimos con los descuentos
        ObtenerDescuentosOLD = DBLet(RS.Fields(0).Value)
        Dto2Fact = DBLet(RS.Fields(1).Value)
    Else
        'Buscar si existe en Promociones Nueva
        ObtenerDescuentosOLD = ""
        Dto2Fact = ""
    End If
    
    RS.Close
    Set RS = Nothing
End Function



Private Function ObtenerPrecioEsp2(FechaOfe As String, PorCaja As Boolean, Dto As String, ComisionEspecial As String) As String
'Comprueba si existe un Precio Especial en la tabla: sprees (Precios Especiales)
'Dim devuelve As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Precio As String
On Error GoTo EObtenerPreEsp

    Precio = ""
    
    'Obtener Precio Actual
    SQL = "SELECT * FROM sprees "
    SQL = SQL & " WHERE codclien=" & CodigoClien & " AND codartic=" & DBSet(CodigoArtic, "T")

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        'Existe una promocion, y salimos con el precio actual o nuevo
        DtoPermitido = CBool(RS!dtopermi)
        
        If Not EsFechaIgualPosteriorDavid(FechaOfe, DBLet(RS!fechanue)) Then
        'Precio Actual
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precioa1))
                'Agos2019
                If Precio = "0" Then Precio = DBLet(RS!precioac)
                
            Else
                Precio = DBLet(RS!precioac)
            End If
            If DtoPermitido Then Dto = DBLet(RS!dtoespec)
        'ElseIf EsFechaIgualPosterior(RS!fechanue, FechaOfe, False) Then
        Else
            'Precio Nuevo
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precion1))
                'Agos2019
                If Precio = "0" Then Precio = DBLet(RS!precionu)
            Else
                Precio = ComprobarCero(DBLet(RS!precionu))
            End If
            If DtoPermitido Then Dto = DBLet(RS!dtoespe1)
        End If
        
        'Comision especial
        ComisionEspecial = DBLet(RS!Comision)
        
    Else  'No existe Precio Especial salir y pasar al siguiente paso
        Precio = ""
        DtoPermitido = False
        Dto = ""
        ComisionEspecial = ""
    End If
    RS.Close
    Set RS = Nothing
    ObtenerPrecioEsp2 = Precio
EObtenerPreEsp:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Especia", Err.Description
End Function


Private Function ObtenerTarifaArt(FechaOfe As String, PorCaja As Boolean) As String
'Comprueba si existe una Tarifa de Articulos en la tabla: slista (Tarifas Articulos)
'Dim devuelve As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Precio As String

    On Error GoTo EObtenerTarifa

    Precio = ""
    
    'Obtener Precio Actual
    SQL = "SELECT * FROM slista "
    SQL = SQL & " WHERE codartic=" & DBSet(CodigoArtic, "T") & " AND codlista=" & CodigoLista2
    'SQL = SQL & " and ( isnull(fechanue) or (fechanue<= '" & Format(FechaOfe, FormatoFecha) & "'))"

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
    'Existe una promocion, y salimos con el precio actual o nuevo
        DtoPermitido = CBool(RS!dtopermi)
        
        'MARZo 2015
        'If IsNull(RS!fechanue) Or (EsFechaIgualPosterior(FechaOfe, DBLet(RS!fechanue), False)) Then
        If Not EsFechaIgualPosteriorDavid(FechaOfe, DBLet(RS!fechanue)) Then
            'Precio Actual
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precioa1))
                'Cambio ENerp 2011.  Si va por caja, el precio por caja es 0 pongo el precio normal
                If Precio = "0" Then Precio = DBLet(RS!precioac)
            
            Else
                Precio = DBLet(RS!precioac)
            End If
        'ElseIf EsFechaIgualPosterior(RS!fechanue, FechaOfe, False) Then
        Else
            'Precio Nuevo
            If PorCaja Then
                Precio = ComprobarCero(DBLet(RS!precion1))
                'Cambio ENerp 2011
                If Precio = "0" Then Precio = DBLet(RS!precionu)
            Else
                Precio = ComprobarCero(DBLet(RS!precionu))
            End If
        End If
    Else  'No existe Precio Especial salir y pasar al siguiente paso
        Precio = ""
        DtoPermitido = False
    End If
    RS.Close
    Set RS = Nothing
    ObtenerTarifaArt = Precio
EObtenerTarifa:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Tarifa Artículo", Err.Description
End Function


Private Function ObtenerPrecioVenta() As String
'Obtiene el precio de Venta al Publico (preciove) de la tabla sartic
    ObtenerPrecioVenta = DevuelveDesdeBDNew(conAri, "sartic", "preciove", "codartic", CodigoArtic, "T")
End Function


Public Function ObtenerNumCajas(TUnidades As String, UniCaja As String) As Integer
Dim NumCajas As Integer
Dim cantidad As Integer, UniPorCaja As Integer
On Error Resume Next

    cantidad = CInt(TUnidades)
    UniPorCaja = CInt(UniCaja)
    If UniPorCaja > 1 Then 'Se vende en cajas
        NumCajas = Int(cantidad / UniPorCaja)
    Else 'No se vende por cajas
        NumCajas = 0
    End If
    ObtenerNumCajas = NumCajas
End Function



Public Function FijarTarifaActividad()
Dim RS As ADODB.Recordset
Dim Aux As String

    Aux = "Select codactiv,codtarif,promocio from sclien where codclien = " & CodigoClien
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        mvarCodActividad = 0
        mvarCodLista = 0
        mvarAplicaPromo = False
    Else
        mvarAplicaPromo = DBLet(RS!promocio, "N") = 1
        mvarCodActividad = DBLet(RS!codactiv, "N")
        mvarCodLista = RS!codTarif
    End If
    RS.Close
    Set RS = Nothing
    
End Function






'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'
'       16 Marzo 2010
'       Nueva funcion para obtener los descuentos
'
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Private Function ObtenerDescuentosN(FechaOfe As String, Dto2Fact, PorCaja As Boolean, Comision As String) As String
Dim RS As ADODB.Recordset
Dim SQL As String
Dim CampoDto1 As String, campoDto2 As String
Dim vCodMarca As String

    '- obtener la familia y la marca del articulo
    vCodMarca = "Codmarca"
    CodigoFamia = DevuelveDesdeBDNew(conAri, "sartic", "Codfamia", "codartic", CodigoArtic, "T", vCodMarca)
    CodigoMarca = vCodMarca
    
    'Antes Nov 2013
'    If PorCaja Then 'Obtener Descuentos Por Caja
'        CampoDto1 = "dtocaja1"
'        campoDto2 = "dtocaja2"
'    Else 'Obtener Descuentos por Unidad
'        CampoDto1 = "dtoline1"
'        campoDto2 = "dtoline2"
'    End If
    CampoDto1 = "dtoline1"
    campoDto2 = "dtoline2"
    If PorCaja Then 'Obtener Descuentos Por Caja
        If vParamAplic.NumeroInstalacion <> vbHerbelca Then       'Para todos menos para herbelca
            CampoDto1 = "dtocaja1"
            campoDto2 = "dtocaja2"
        End If
    End If
    
    Comision = ""
    
    
    'Obtener Descuento:
    '1- buscar descuento para la cliente,familia y marca
    SQL = "SELECT " & CampoDto1 & ", " & campoDto2 & ", codactiv,comision FROM sdtofm "
    SQL = SQL & " WHERE codclien=" & CodigoClien
    SQL = SQL & " AND (fechadto<= '" & Format(FechaOfe, FormatoFecha) & "')"
    SQL = SQL & " AND ((codfamia=" & CodigoFamia & " AND codmarca=" & CodigoMarca & ")"
    SQL = SQL & " OR (codfamia=" & CodigoFamia & " AND codmarca is null)"
    SQL = SQL & " OR (codfamia is null AND codmarca=" & CodigoMarca & ")"
    SQL = SQL & " OR (codfamia is null AND codmarca is null))"
    If vParamAplic.OrdenDtos = 0 Then
        SQL = SQL & " ORDER BY codfamia DESC, codmarca DESC"
    Else
        SQL = SQL & " ORDER BY codmarca DESC,codfamia DESC"
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
    'Existe el descuento, y salimos con los descuentos
        ObtenerDescuentosN = DBLet(RS.Fields(0).Value)
        Dto2Fact = DBLet(RS.Fields(1).Value)
        Comision = DBLet(RS!Comision)
    Else
        'Buscar si existe en Promociones Nueva
        ObtenerDescuentosN = ""
        Dto2Fact = ""
    End If
    
    RS.Close
    
    'YA lo ha encentrado
    If ObtenerDescuentosN <> "" Then Exit Function
    
    
    SQL = "SELECT " & CampoDto1 & ", " & campoDto2 & ",comision FROM sdtofm "
    SQL = SQL & " WHERE codactiv = " & CodActividad
    SQL = SQL & " AND (fechadto<= '" & Format(FechaOfe, FormatoFecha) & "')"
    SQL = SQL & " AND ((codfamia=" & CodigoFamia & " AND codmarca=" & CodigoMarca & ")"
    SQL = SQL & " OR (codfamia=" & CodigoFamia & " AND codmarca is null)"
    SQL = SQL & " OR (codfamia is null AND codmarca=" & CodigoMarca & ")"
    SQL = SQL & " OR (codfamia is null AND codmarca is null))"
    If vParamAplic.OrdenDtos = 0 Then
        SQL = SQL & " ORDER BY codfamia DESC, codmarca DESC"
    Else
        SQL = SQL & " ORDER BY codmarca DESC,codfamia DESC"
    End If
    

    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
    'Existe el descuento, y salimos con los descuentos
        ObtenerDescuentosN = DBLet(RS.Fields(0).Value)
        Dto2Fact = DBLet(RS.Fields(1).Value)
        Comision = DBLet(RS!Comision)
    Else
        'Buscar si existe en Promociones Nueva
        ObtenerDescuentosN = ""
        Dto2Fact = ""
    End If
    RS.Close
    
    
    'YA lo ha encentrado
    If ObtenerDescuentosN <> "" Then Exit Function
    
    
    'fANALMENTE VEO SI HAY DESCUENTOS A LA FAMILIA O A LA MARCA
    SQL = "SELECT dtoline1, dtoline2 FROM sdtofm  WHERE (fechadto<= '" & Format(FechaOfe, FormatoFecha) & "')"
    SQL = SQL & " AND codactiv IS NULL AND codclien IS NULL AND "
    SQL = SQL & " ((codmarca = " & CodigoMarca & " AND codfamia is NULL ) OR (codfamia= " & CodigoFamia & " and codmarca is null))"
    If vParamAplic.OrdenDtos = 0 Then
        SQL = SQL & " ORDER BY codfamia DESC, codmarca DESC"
    Else
        SQL = SQL & " ORDER BY codmarca DESC,codfamia DESC"
    End If
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
    'Existe el descuento, y salimos con los descuentos
        ObtenerDescuentosN = DBLet(RS.Fields(0).Value)
        Dto2Fact = DBLet(RS.Fields(1).Value)
    Else
        'Buscar si existe en Promociones Nueva
        ObtenerDescuentosN = ""
        Dto2Fact = ""
    End If
    RS.Close
    
    
    
    
    
    'EN fenollar buscamos por catalogo
    
    
    
    
    
    
    Set RS = Nothing
End Function



'***************************************************************************************
'***************************************************************************************
'***************************************************************************************
'***************************************************************************************
'***************************************************************************************
'***************************************************************************************
'***************************************************************************************
'***************************************************************************************
'
'
'  HERBELCA.  Para saber si esta vendiendo por debajo de lo que le toca
'               para las ventas ECO
'
'
'***************************************************************************************
Public Function ObtenerPrecioDtoFamilia(PCaja As Boolean, FechaOfe As String, OrigP As String) As String
    
'    If vParamAplic.PrecioMinimo Then
'        'ObtenerPrecio
'        ObtenerPrecio = ObtenerPrecioMinimo(PCaja, FechaOfe, OrigP)
'    Else
'        ObtenerPrecio = ObtenerPrecioNormal(PCaja, FechaOfe, OrigP)

    
'    End If
    ObtenerPrecioDtoFamilia = ObtenerPrecioConDtoFamilia(PCaja, FechaOfe, OrigP)
End Function



Private Function ObtenerPrecioConDtoFamilia(PCaja As Boolean, FechaOfe As String, OrigP_ As String) As String
Dim ElPrecio2 As String
Dim PrecioMinConDto As String
Dim Pre As String

Dim Dt1F_BD As String
Dim Dt2F_BD As String



Dim Dto1Fam As String
Dim Dto2Fam As String
Dim DtoPermit As Boolean

Dim PrecioFact As String
Dim Dto1Fact_ As String
Dim Dto2Fact_ As String

Dim Comi As String

    On Error GoTo EObtenerPre
    ElPrecio2 = "99999999999" 'Empezamos a este nivel
    PrecioMinConDto = ElPrecio2
    
    
    
    'Modificacion1, busco los descuentos
    Dt1F_BD = ObtenerDescuentoDeLaFamilia(Dt2F_BD, PCaja)
    
    
    'COMPROBAR PROMOCIONES
    '==============================================================
    'Buscar Precio en la tabla spromo si encuentro valor salgo con ese precio
    PrecioFact = ObtenerPromocion(FechaOfe, PCaja)
    Dto1Fact_ = ""
    Dto2Fact_ = ""
    If PrecioFact <> "" Then 'Ha obtenido el Precio
        If DtoPermitido Then
       
            Dto1Fact_ = Dt1F_BD
            Dto2Fact_ = Dt2F_BD
       
        End If
        OrigP_ = "P" 'Indica que sale con precio de Promocion
        ElPrecio2 = PrecioFact
        PrecioMinConDto = CalcularImporte("1", ElPrecio2, Dto1Fact_, Dto2Fact_, vParamAplic.TipoDtos)
        DtoPermit = DtoPermitido
        Dto1Fam = Dto1Fact_
        Dto2Fam = Dto2Fact_
        PrecioFact = ""  'Para que entre en el siguiente
    End If
    
    
    'No encuentra Precio en spromo y Buscar en Sprees(Tabla Precios Especiales)
    'COMPROBAR PRECIO ESPECIAL
    '===========================================================================
    Dto1Fact_ = ""
    Dto2Fact_ = "0"
    PrecioFact = ObtenerPrecioEsp2(FechaOfe, PCaja, Dto1Fact_, Comi)  'comi no va a ningun lado
    
    If PrecioFact <> "" Then 'Ha obtenido el precio
        If DtoPermitido Then
            'Buscar en la tabla sdtofm los descuentos
            Dto1Fact_ = Dt1F_BD
            Dto2Fact_ = Dt2F_BD
        End If
        'OrigP = "E" 'Indica que sale con Precio Especial
        Pre = CalcularImporte("1", PrecioFact, Dto1Fact_, Dto2Fact_, vParamAplic.TipoDtos)
        
        'Veremos el minimo
        If CCur(PrecioMinConDto) > CCur(Pre) Then
            'Si, es menor
            DtoPermit = DtoPermitido
            OrigP_ = "E"
            ElPrecio2 = PrecioFact
            Dto1Fam = Dto1Fact_
            Dto2Fam = Dto2Fact_
            PrecioMinConDto = Pre
        End If
        
        PrecioFact = ""  'Para que entre en el siguiente
        
    End If
    
    
        
    'COMPROBAR TARIFAS ARTICULOS
    '=======================================================================
    PrecioFact = ObtenerTarifaArt(FechaOfe, PCaja)
    If PrecioFact <> "" Then 'Ha obtenido el precio
        If DtoPermitido Then
        'Buscar en la tabla sdtofm los descuentos
            Dto1Fact_ = Dt1F_BD
            Dto2Fact_ = Dt2F_BD
        Else
            Dto1Fact_ = ""
            Dto2Fact_ = ""
        End If
        'OrigP = "T"
        Pre = CalcularImporte("1", PrecioFact, Dto1Fact_, Dto2Fact_, vParamAplic.TipoDtos)
        
        'Veremos el minimo
        If CCur(PrecioMinConDto) > CCur(Pre) Then
            'Si, es menor
            DtoPermit = DtoPermitido
            OrigP_ = "T"
            ElPrecio2 = PrecioFact
            Dto1Fam = Dto1Fact_
            Dto2Fam = Dto2Fact_
            PrecioMinConDto = Pre
        End If
        
     End If
        
            
     'El precio venta
     
     PrecioFact = ObtenerPrecioVenta
     PCaja = False
     Dto1Fact_ = 0
     Dto2Fact_ = 0
     Pre = CalcularImporte("1", PrecioFact, Dto1Fact_, Dto2Fact_, vParamAplic.TipoDtos)
     If CCur(PrecioMinConDto) > CCur(Pre) Then
        'DtoPermitido = True  'pq de tarifa sii permite dot
        DtoPermit = True
         'Si, es menor
         OrigP_ = "A"
         ElPrecio2 = PrecioFact
         
         
         Dto1Fam = Dto1Fact_
         Dto2Fam = Dto2Fact_
     End If
                
    DtoPermitido = DtoPermit
    ObtenerPrecioConDtoFamilia = ComprobarCero(ElPrecio2)
    Descuento1 = ComprobarCero(Dto1Fam)
    Descuento2 = ComprobarCero(Dto2Fam)
    Exit Function
    
EObtenerPre:
    MuestraError Err.Number, "Obtener Precio.", Err.Description
End Function



Private Function ObtenerDescuentoDeLaFamilia(Dto2Fact, PorCaja As Boolean) As String
Dim RS As ADODB.Recordset
Dim SQL As String
Dim CampoDto1 As String, campoDto2 As String


    '- obtener la familia y la marca del articulo
    CodigoFamia = DevuelveDesdeBDNew(conAri, "sartic", "Codfamia", "codartic", CodigoArtic, "T")

    

    CampoDto1 = "dtoline1"
    campoDto2 = "dtoline2"
    If PorCaja Then 'Obtener Descuentos Por Caja
        If vParamAplic.NumeroInstalacion <> vbHerbelca Then       'Para todos menos para herbelca
            CampoDto1 = "dtocaja1"
            campoDto2 = "dtocaja2"
        End If
    End If
    
    
    
    
    'Obtener Descuento:
    '1- buscar descuento para la cliente,familia y marca
    'SQL = "SELECT dtoline1,dtoline2 FROM sfamiadtos "
    SQL = "SELECT max(dtoline1),max(dtoline2) FROM sfamiadtos "
    SQL = SQL & " WHERE  codfamia=" & CodigoFamia
        
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
    'Existe el descuento, y salimos con los descuentos
        ObtenerDescuentoDeLaFamilia = DBLet(RS.Fields(0).Value)
        Dto2Fact = DBLet(RS.Fields(1).Value)
    Else
        ObtenerDescuentoDeLaFamilia = ""
        Dto2Fact = ""
    End If
    
    RS.Close
    
    Set RS = Nothing
End Function


Private Function DescuentosPorCatalogo() As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Catalogos As String
Dim CatPpal As String
    Set RS = New ADODB.Recordset
    DescuentosPorCatalogo = ""
    
    SQL = "Select * from sagrupaart where codartic=" & DBSet(mvarCodArtic, "T")
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    CatPpal = ""
    If Not RS.EOF Then
        Catalogos = ""
        While Not RS.EOF
            'ESTO DEBERIA SACARLO ANTES POR SI CREAN NUEVAS AGRUPACIONES. De momento para que funcione
            '
            SQL = ""
            Select Case RS!codagrupa
            Case "00"
                
                CatPpal = CatPpal & ", 0"
            Case "01"
                 CatPpal = CatPpal & ", 1"
                 
            Case "02"
                 CatPpal = CatPpal & ", 2"
            Case "03"
                 CatPpal = CatPpal & ", 3"
            Case "04"
                 CatPpal = CatPpal & ", 4"
            Case Else
                
                     SQL = DBSet(RS!codagrupa, "T")
            End Select
            If SQL <> "" Then Catalogos = Catalogos & ", " & SQL
            RS.MoveNext
        Wend
        RS.Close
        SQL = ""
        If CatPpal <> "" Then SQL = "( agruppal in (" & Mid(CatPpal, 2) & ") )"
            
        If Catalogos <> "" Then
            Catalogos = " ( sagrupacli.codagrupa in (" & Mid(Catalogos, 2) & ") )"
            If SQL <> "" Then SQL = SQL & " OR "
            SQL = SQL & Catalogos
        End If
        
        If SQL <> "" Then
            SQL = " sagrupacli.codclien =" & mvarCodClien & " AND sagrupa.codagrupa =sagrupacli.codagrupa AND tipo='C' AND " & SQL
            SQL = "Select * from    sagrupacli,sagrupa WHERE  " & SQL & " order by AgruPpal ,tipo"
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                
                DescuentosPorCatalogo = DBLet(RS!Dto1)
            
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing
End Function
