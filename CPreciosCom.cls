VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPreciosCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Modulo para Calcular el Precio a Aplicar a un PEDIDO de Compras (a PROVEEDOR)
' Y los descuentos correspondientes si se aplica descuento
'TABLAS: slispr, sdtomp

'------------------------------------------------
'Variables del modulo CPreciosCom
'------------------------------------------------

'TABLA: slispr
'---------------
Private mvarCodArtic As String
Private mvarCodProve As String

Private mvarCantFija As Integer 'Cantidad fija
Private mvarCantMin As Integer  'Cantidad minima

Private mvarDtoPermitido As Boolean 'si permite descuento, buscará posibles descuentos en la tabla "sdtomp"

'Datos Promocion
Private mvarFecIniPromo As String
Private mvarFecFinPromo As String
Private mvarPrecioPromo As Currency
Private mvarDtoPermiPromo As Boolean

'Valores actuales
Private mvarPrecioAct As Currency
Private mvarPrecioNue As Currency
Private mvarFechaNue As String

'**********************************
'Laura: campos añadidos el 11/09/06
Private mvarDtoAct1 As Currency
Private mvarDtoAct2 As Currency
'**********************************

'TABLA: sdtomp
'---------------
Private mvarCodFamia As String
Private mvarCodMarca As String



'valor del descuento que devuelve para aplicarlo
Private mvarDescuento1 As String
Private mvardescuento2 As String


'------------------------------------------------
'Propiedades del modulo CPreciosCom
'------------------------------------------------


Public Property Let CantFija(ByVal vData As Integer)
     mvarCantFija = vData
End Property

Public Property Get CantFija() As Integer
     CantFija = mvarCantFija
End Property


Public Property Let CantMinima(ByVal vData As Integer)
     mvarCantMin = vData
End Property

Public Property Get CantMinima() As Integer
     CantMinima = mvarCantMin
End Property


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


Public Property Let CodigoProve(ByVal vData As String)
     mvarCodProve = vData
End Property

Public Property Get CodigoProve() As String
     CodigoProve = mvarCodProve
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


Public Property Let FechaIniPromo(ByVal vData As String)
     mvarFecIniPromo = vData
End Property

Public Property Get FechaIniPromo() As String
     FechaIniPromo = mvarFecIniPromo
End Property


Public Property Let FechaFinPromo(ByVal vData As String)
     mvarFecFinPromo = vData
End Property

Public Property Get FechaFinPromo() As String
     FechaFinPromo = mvarFecFinPromo
End Property


Public Property Let PrecioPromo(ByVal vData As Currency)
     mvarPrecioPromo = vData
End Property

Public Property Get PrecioPromo() As Currency
     PrecioPromo = mvarPrecioPromo
End Property


Public Property Let DtoPermiPromo(ByVal vData As Boolean)
     mvarDtoPermiPromo = vData
End Property

Public Property Get DtoPermiPromo() As Boolean
     DtoPermiPromo = mvarDtoPermiPromo
End Property


Public Property Let PrecioActual(ByVal vData As Currency)
     mvarPrecioAct = vData
End Property

Public Property Get PrecioActual() As Currency
     PrecioActual = mvarPrecioAct
End Property


Public Property Let PrecioNuevo(ByVal vData As Currency)
     mvarPrecioNue = vData
End Property

Public Property Get PrecioNuevo() As Currency
     PrecioNuevo = mvarPrecioNue
End Property

Public Property Let FechaNuevo(ByVal vData As String)
     mvarFechaNue = vData
End Property

Public Property Get FechaNuevo() As String
     FechaNuevo = mvarFechaNue
End Property


'Descuento actual 1: de la tabla slispr
Public Property Let DescuentoActual1(ByVal vData As Currency)
     mvarDtoAct1 = vData
End Property

Public Property Get DescuentoActual1() As Currency
     DescuentoActual1 = mvarDtoAct1
End Property



'Descuento actual 2: de la tabla slispr
Public Property Let DescuentoActual2(ByVal vData As Currency)
     mvarDtoAct2 = vData
End Property

Public Property Get DescuentoActual2() As Currency
     DescuentoActual2 = mvarDtoAct2
End Property






'Descuentos de la tabla sdtomp
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
'Metodos del modulo CPreciosCom
'------------------------------------------------

Public Function Leer(codArtic As String, codProve As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo ELeer

    SQL = "SELECT * FROM slispr "
    SQL = SQL & " WHERE codartic=" & DBSet(codArtic, "T") & " AND codprove=" & codProve
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        mvarCantFija = DBLet(RS!CantFija, "N")
        mvarCantMin = DBLet(RS!cantmini, "N")

        mvarDtoPermitido = RS!dtopermi
        mvarCodArtic = DBLet(RS!codArtic, "T")
        mvarCodProve = RS!codProve
        
        'datos de promocion
        mvarFecIniPromo = DBLet(RS!FechaIni, "F")
        mvarFecFinPromo = DBLet(RS!FechaFin, "F")
        mvarPrecioPromo = DBLet(RS!preciopr, "N")
        mvarDtoPermiPromo = DBLet(RS!dtoperm1, "B")
        
        'valores actuales
        mvarPrecioAct = DBLet(RS!precioac, "N")
        mvarPrecioNue = DBLet(RS!precionu, "N")
        mvarFechaNue = DBLet(RS!fechanue, "F")
        mvarDtoAct1 = DBLet(RS!dtoline1, "N")
        mvarDtoAct2 = DBLet(RS!dtoline2, "N")
        
        Leer = True
    End If
    RS.Close
    Set RS = Nothing
    
ELeer:
    If Err.Number <> 0 Then Leer = False
End Function



Public Function ComprobarCantidad(Cantidad As Integer) As Boolean
'Comprueba si la cantidad a insertar en la linea de pedido o Albaran de Compras
'cumple la cantidad fija o minima si alguna de las 2 tiene valor para el artic-prove

    ComprobarCantidad = False
    
    'Si hay cantidad Fija solo se pueden comprar multiplos de esa cantidad
    If CantFija > 0 Then
        'comprobar q la cantidad es multiplo de CantFija
        If Not (Cantidad >= CantFija) Or Not (Cantidad Mod CantFija = 0) Then
            MsgBox "Cantidad Fija a comprar del artículo " & CodigoArtic & " es:  " & CantFija & "  (o múltiplos)", vbExclamation
            Exit Function
        End If
        
    'Si hay cantidad minima no se puede comprar menos de esa cantidad
    ElseIf CantMinima > 0 Then
        If Not (Cantidad >= CantMinima) Then
            If vParamAplic.NumeroInstalacion <> 2 Then
                MsgBox "Cantidad Mínima a comprar del artículo " & CodigoArtic & " es:  " & CantMinima, vbExclamation
                Exit Function
            End If
        End If
    End If
    
    ComprobarCantidad = True
End Function



Public Function ObtenerPrecio(FechaPed As String) As String
'FechaPed -> IN: Fecha del Pedido para comprobar si esta en los periodos de promocion, etc...
'RETURN: valor del precio
Dim PrecioCom As String
Dim Dto1Com As String
Dim Dto2Com As String

    On Error GoTo EObtenerPre

   
    'COMPROBAR PROMOCIONES
    '==============================================================
    'Buscar Precio en la tabla "slispr" si encuentro valor salgo con ese precio
    PrecioCom = ObtenerPromocion(FechaPed)
    If PrecioCom <> "" Then 'Ha obtenido el Precio
        If DtoPermiPromo Then
            'mirar si los campos de descuento de la tabla slispr tienen valor
            If (DescuentoActual1 + DescuentoActual2) > 0 Then
                Dto1Com = DescuentoActual1
                Dto2Com = DescuentoActual2
            Else
                'Buscar en la tabla "sdtomp" los descuentos
                Dto1Com = ObtenerDescuentos2(FechaPed, Dto2Com)
            End If
        End If
        
    Else 'No encuentra Precio de Promocion en slispr y Buscar en la misma tabla: slispr (Tabla Cab. Precios Prov.)
        'COMPROBAR PRECIO Prov.
        '===========================================================================
        PrecioCom = ObtenerPrecioProv(FechaPed)
        If PrecioCom <> "" Then 'Ha obtenido el precio
            'mirar si los campos de descuento de la tabla slispr tienen valor
            If (DescuentoActual1 + DescuentoActual2) > 0 Then
                Dto1Com = DescuentoActual1
                Dto2Com = DescuentoActual2
            ElseIf DtoPermitido Then
            'Buscar en la tabla sdtomp los descuentos
                Dto1Com = ObtenerDescuentos2(FechaPed, Dto2Com)
            End If
        Else 'Buscar en sartic el precio ultima compra del articulo (preciouc)
            PrecioCom = ObtenerPrecioUltCompra
        End If
    End If
    
    ObtenerPrecio = ComprobarCero(PrecioCom)
    Descuento1 = ComprobarCero(Dto1Com)
    Descuento2 = ComprobarCero(Dto2Com)
    
EObtenerPre:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener Precio", Err.Description
End Function


Private Function ObtenerPromocion(FechaPed As String) As String
'Comprueba si existe una promocion en la tabla: slispr (Cab. Precios Prov.)
Dim Precio As String
On Error GoTo EObtenerPromo

    Precio = ""
    
    'Existe una promocion para ese articulo y proveedor, comprobar fechas
    If FechaIniPromo <> "" And FechaFinPromo <> "" Then
        If EntreFechas(FechaIniPromo, FechaPed, FechaFinPromo) Then
            'Dentro Periodo Promocion Actual
            Precio = PrecioPromo
        End If
    Else 'No existe Promocion salir y pasar al siguiente paso
        Precio = ""
        DtoPermiPromo = False
    End If
    ObtenerPromocion = Precio
    
EObtenerPromo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Promoción", Err.Description
End Function

'27 Sept 2010
'pAso a public la fucnion.
'Asi, puede ir a ver a la tabla sdtomp sin tener que tenere tarifas ni nada
Public Function ObtenerDescuentos2(FechaPed As String, Dto2Com) As String
'IN -> FechaPed: fecha del pedido para comprobar si hay algun descuento valido en esa fecha
'                para el Proveedor, y para la familia y marca del Articulo
'OUT -> ObtenerDescuentos: Aqui se devuelve el valor del Dto1 si lo encuentra
'        Dto2com: Devuelve el valor del Dto2 si lo encuentra
Dim RS As ADODB.Recordset
Dim SQL As String
Dim CampoDto1 As String, campoDto2 As String
Dim vCodMarca As String
    
    'Obtenemos la Familia y la Marca del Articulo
    vCodMarca = "Codmarca"
    CodigoFamia = DevuelveDesdeBDNew(conAri, "sartic", "Codfamia", "codartic", CodigoArtic, "T", vCodMarca)
    CodigoMarca = vCodMarca
    
    CampoDto1 = "dtoline1"
    campoDto2 = "dtoline2"
    
    'Obtener Descuento
    SQL = "SELECT " & CampoDto1 & ", " & campoDto2 & " FROM sdtomp "
    SQL = SQL & " WHERE codprove=" & CodigoProve & " AND codfamia=" & CodigoFamia
    SQL = SQL & " AND (codmarca is null or codmarca=" & CodigoMarca & ")"
    SQL = SQL & " and (fechadto<= '" & Format(FechaPed, FormatoFecha) & "') ORDER BY codmarca desc"

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
    'Existe el descuento, y salimos con los descuentos
        ObtenerDescuentos2 = DBLet(RS.Fields(0).Value, "N")
        Dto2Com = DBLet(RS.Fields(1).Value, "N")
    Else
        ObtenerDescuentos2 = ""
        Dto2Com = ""
    End If
    RS.Close
    Set RS = Nothing
End Function


Private Function ObtenerPrecioProv(FechaPed As String) As String
'Comprueba si existe un Precio de Proveedor en la tabla: slispr (Cab. Precios Prov.)
Dim Precio As String

    On Error GoTo EObtenerPreProv

    Precio = ""
    
    If FechaNuevo <> "" Then
    'Existe un precio para el Proveedor, y salimos con el precio actual o nuevo
        
        'Laura: modificado el 11/09/06
        '************************************************************
'        If (EsFechaIgualPosterior(FechaPed, FechaNuevo, False)) Then
'            'Precio Actual
'            Precio = CStr(PrecioActual)
'        ElseIf EsFechaIgualPosterior(FechaNuevo, FechaPed, False) Then
'            'Precio Nuevo
'            Precio = CStr(PrecioNuevo)
'        End If
        If EsFechaIgualPosterior(FechaNuevo, FechaPed, False) Then
            'Precio Nuevo
            Precio = CStr(PrecioNuevo)
        Else
            'Precio Actual
            Precio = CStr(PrecioActual)
        End If
        '*************************************************************
        
    'Laura: añadido el 11/09/06
    '*****************************
    ElseIf PrecioActual > 0 Then
        Precio = PrecioActual
    '*****************************
    
    Else  'No existe Precio Especial salir y pasar al siguiente paso
        Precio = ""
        DtoPermitido = False
    End If
    ObtenerPrecioProv = Precio
    
EObtenerPreProv:
    If Err.Number <> 0 Then MuestraError Err.Number, "Precio Proveedor", Err.Description
End Function


Private Function ObtenerPrecioUltCompra() As String
'Obtiene el precio de ultima compra (preciouc) de la tabla sartic
    ObtenerPrecioUltCompra = DevuelveDesdeBDNew(conAri, "sartic", "preciouc", "codartic", CodigoArtic, "T")
End Function

