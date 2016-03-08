VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ARTICULOS
'TABLA: sartic

'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodArtic As String 'Cod. articulo
Private mNomArtic As String 'Nombre articulo

Private mCodFamia As String 'Familia del Articulo
Private mCodCateg As String 'categoria del Articulo

Private mCodStatu As Byte 'cod. estatus: 1=Bloqueado, 2=Caducado
Private mCtrStock As Byte 'hay control de stock (si/No)
Private mArtVario As Byte 'Si es articulo de varios o no
Private mConjunto As Boolean  'Si es articulo de conjunto o no

Private mCodigIVA As Byte

Private mPrecioVe As String 'Precio de venta (P.V.P.)
Private mUniCajas As Integer 'unidades por caja
Private mPreciomp As Currency 'Precio medio ponderado
Private mPreciost As Currency 'Precio standard
Private mPreciouc As Currency 'Precio ultima compra
Private mMargeCom As Currency 'margen comercial

Private mTextoVen As String 'texto a mostrar para ventas
Private mTextoCom As String 'texto a mostrar para compras

Private mCodProve As Long   'codigo proveedor



'------------------------------------------------
'Propiedades del modulo CArticulo
'------------------------------------------------

'**** Codigo del articulo
Public Property Let codigo(ByVal vData As String)
     mCodArtic = vData
End Property

Public Property Get codigo() As String
     codigo = mCodArtic
End Property


'**** Nombre del Articulo
Public Property Let Nombre(ByVal vData As String)
     mNomArtic = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomArtic
End Property



'**** Familia del Articulo
Public Property Let Familia(ByVal vData As String)
     mCodFamia = vData
End Property

Public Property Get Familia() As String
     Familia = mCodFamia
End Property


'**** Categoria del Articulo
Public Property Let categoria(ByVal vData As String)
     mCodCateg = vData
End Property

Public Property Get categoria() As String
     categoria = mCodCateg
End Property



'**** cod. status del Articulo
Public Property Let Status(ByVal vData As Byte)
     mCodStatu = vData
End Property

Public Property Get Status() As Byte
     Status = mCodStatu
End Property



'**** Articulo tiene control de stock (si/no)
Public Property Let CtrStock(ByVal vData As Byte)
     mCtrStock = vData
End Property

Public Property Get CtrStock() As Byte
     CtrStock = mCtrStock
End Property



'**** Articulo es de Varios (Si/No)
Public Property Let EsDeVarios(ByVal vData As Byte)
     mArtVario = vData
End Property

Public Property Get EsDeVarios() As Byte
     EsDeVarios = mArtVario
End Property



'**** Articulo es de Conjunto (Si/No)
Public Property Let EsConjunto(ByVal vData As Boolean)
     mConjunto = vData
End Property

Public Property Get EsConjunto() As Boolean
     EsConjunto = mConjunto
End Property



'**** Precio de Venta
Public Property Let PrecioVenta(ByVal vData As String)
     mPrecioVe = vData
End Property

Public Property Get PrecioVenta() As String
     PrecioVenta = mPrecioVe
End Property


'**** Unidades por caja
Public Property Let UnidCaja(ByVal vData As Integer)
     mUniCajas = vData
End Property

Public Property Get UnidCaja() As Integer
     UnidCaja = mUniCajas
End Property



'**** Precio medio ponderado
Public Property Let PrecioMedPon(ByVal vData As Currency)
     mPreciomp = vData
End Property

Public Property Get PrecioMedPon() As Currency
     PrecioMedPon = mPreciomp
End Property


'**** Precio standard
Public Property Let PrecioStan(ByVal vData As Currency)
     mPreciost = vData
End Property

Public Property Get PrecioStan() As Currency
     PrecioStan = mPreciost
End Property



'**** Precio ultima compra
Public Property Let PrecioUltCom(ByVal vData As Currency)
     mPreciouc = vData
End Property

Public Property Get PrecioUltCom() As Currency
     PrecioUltCom = mPreciouc
End Property


'**** Margen comercial
Public Property Let MargenComercial(ByVal vData As Currency)
     mMargeCom = vData
End Property

Public Property Get MargenComercial() As Currency
     MargenComercial = mMargeCom
End Property



'**** Tipo de IVA del articulo
Public Property Let TipoIVA(ByVal vData As Byte)
     mCodigIVA = vData
End Property

Public Property Get TipoIVA() As Byte
     TipoIVA = mCodigIVA
End Property




'**** Texto para Ventas
Public Property Let TextoVentas(ByVal vData As String)
     mTextoVen = vData
End Property

Public Property Get TextoVentas() As String
     TextoVentas = mTextoVen
End Property


'**** Texto para Compras
Public Property Let TextoCompras(ByVal vData As String)
     mTextoCom = vData
End Property

Public Property Get TextoCompras() As String
     TextoCompras = mTextoCom
End Property


'**** Codigo proveedor
Public Property Let CodProve(ByVal vData As Long)
     mCodProve = vData
End Property

Public Property Get CodProve() As Long
     CodProve = mCodProve
End Property





'------------------------------------------------
'Procedimientos del modulo CArticulo
'------------------------------------------------

Public Function Existe(vCodartic As String) As Boolean
'Comprueba si existe el articulo en la BD
Dim devuelve As String
Dim Nom As String

    On Error GoTo EExiste
    
    Nom = "nomartic"
    devuelve = DevuelveDesdeBD(conAri, "codartic", "sartic", "codartic", vCodartic, "T", Nom)
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el Artículo " & vCodartic & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
        codigo = vCodartic
        Nombre = Nom
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function



Public Function LeerDatos(vCodartic As String) As Boolean
'Leer los datos de un Articulo dado
'Lee de la BD: Ariges, Tabla: sartic
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    LeerDatos = False
    
    SQL = "SELECT codartic,nomartic,codfamia,codstatu,ctrstock,preciove,unicajas,codigiva,artvario,conjunto,textoven,textocom,codcateg,preciomp,preciost,preciouc,margecom,codprove "
    SQL = SQL & " FROM sartic "
    SQL = SQL & " WHERE codartic=" & DBSet(vCodartic, "T")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        codigo = CStr(RS!codArtic)
        
        Nombre = RS!NomArtic
        Familia = RS!Codfamia
        categoria = DBLet(RS!codCateg, "T")
        
        Status = RS!codstatu
        CtrStock = RS!CtrStock
        EsDeVarios = RS!artvario
        Me.EsConjunto = RS!Conjunto
        
        PrecioVenta = RS!preciove
        TipoIVA = RS!CodigIVA
        UnidCaja = RS!unicajas
        
        PrecioMedPon = DBLet(RS!PrecioMP, "N") 'precio medio ponderado
        PrecioStan = DBLet(RS!preciost, "N") 'precio standard
        PrecioUltCom = DBLet(RS!precioUC, "N") 'precio ultima compra
        MargenComercial = DBLet(RS!margecom, "N") 'margen comercial
        
        TextoVentas = DBLet(RS!textoven, "T")
        TextoCompras = DBLet(RS!textocom, "T")
        
        CodProve = DBLet(RS!CodProve, "N")
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Artículo" & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function



Public Sub MostrarStatusArtic(Bloquea As Boolean)
'Recuperamos el estatus (codstatu) del articulo, y si tiene
'valor 1 o 2 lo mostramos en un mensaje
'codstatu=0 normal

'ABRIL 2014
'añadimos el estado OBSOLETO Bloquea = True
'No bloquea nunca, pero da msgbox
'ANTES
'codstatu=1 bloqueado, muestra un mensaje
'codstatu=2 caducado, muestra un mensaje y bloquea linea de articulo

'AHORA
'codstatu=1 Obsoleto
'codstatu=2 bloqueado, muestra un mensaje
'codstatu=3 caducado, muestra un mensaje y bloquea linea de articulo


    Bloquea = False

    Select Case Status
        ' 0 = Normal
        Case 1 'Obsoleto
            MsgBox "El articulo " & codigo & " esta OBSOLETO.", vbExclamation
        Case 2 'BLOQUEADO
            MsgBox "El articulo " & codigo & " esta BLOQUEADO." & vbCrLf & "No se puede insertar una linea con este artículo.", vbExclamation, "Situación artículo"
            Bloquea = True
        Case 3 'CADUCADO
            MsgBox "El articulo " & codigo & " esta CADUCADO.", vbInformation, "Situación artículo"
    End Select
End Sub


Public Function ExisteEnAlmacen(codAlm As String, Optional CanStock As String) As Boolean
Dim devuelve As String
Dim Cantidad As String

    On Error GoTo EExisteAlm

    'Comprobar que existe de ese articulo en el almacen seleccionado
    Cantidad = "canstock"
    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", codigo, "T", Cantidad, "codalmac", codAlm, "N")
    
    If devuelve = "" Then
        MsgBox "No existe unidades del Artículo: " & codigo & "  en el Almacen: " & codAlm, vbExclamation
        ExisteEnAlmacen = False
    Else
        ExisteEnAlmacen = True
        CanStock = Cantidad
    End If
    
EExisteAlm:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Artículo", Err.Description
End Function




Public Function ExistenciaTotalAlmacenes() As Currency
'-- devuelve la cantidad de stock total en todos los almacenes del articulo
Dim rst As ADODB.Recordset
Dim SQL As String
    
    On Error GoTo EExisTotal
    
    If codigo <> "" Then
        SQL = "SELECT SUM(canstock) FROM salmac where codartic=" & DBSet(codigo, "T")
        Set rst = New ADODB.Recordset
        rst.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not rst.EOF Then
            ExistenciaTotalAlmacenes = rst.Fields(0).Value
        End If
        rst.Close
        Set rst = Nothing
    End If
    
    Exit Function
    
EExisTotal:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener existencia total en Almacenes del Artículo", Err.Description
End Function


Public Function EnInventario(codAlm As String) As Boolean
'SI  se esta haciendo inventario del articulo o no
Dim devuelve As String

    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "statusin", "codartic", codigo, "T", , "codalmac", codAlm, "N")
    
    If devuelve = "1" Then
        EnInventario = True
        MsgBox "El artículo " & codigo & " esta inventariandose.", vbInformation
    Else
        EnInventario = False
    End If
End Function


Public Function EsInstalacion() As Boolean
'Coprobar si el articulo es instalacion
Dim devuelve As String
    
    If Familia = "" Then _
        Familia = DevuelveDesdeBDNew(conAri, "sartic", "codfamia", "codartic", codigo, "T")
        
    
    devuelve = DevuelveDesdeBDNew(conAri, "sfamia", "instalac", "codfamia", Familia, "N")
    If devuelve = "0" Then 'No es instalacion
        EsInstalacion = False
    ElseIf devuelve = "1" Then
        EsInstalacion = True
    End If

End Function


Public Sub MostrarTextoCom()
'Recuperamos el texto para compras y ventas del articulo, y si tiene
'valor lo mostramos en un mensaje
     MsgBox TextoCompras, vbInformation, "Texto para compras"
End Sub


Public Sub MostrarTextoVen()
'Recuperamos el texto para compras y ventas del articulo, y si tiene
'valor lo mostramos en un mensaje
     MsgBox TextoVentas, vbInformation, "Texto para ventas"
End Sub



Public Function ActualizarUltFechaCompra(vFecha As String, vPrecio As String) As Boolean
'Modificar fecha ult.compra y precio ult.compra en la tabla sartic
Dim SQL As String
Dim actualiza As Boolean

    On Error GoTo EActFecha

    'Obtenemos la ult. fecha de compra que tiene el articulo
    SQL = DevuelveDesdeBDNew(conAri, "sartic", "ultfecco", "codartic", codigo, "T")
    
    actualiza = False
    
    If SQL = "" Then
        'No tiene fecha ult. compra
        actualiza = True
    ElseIf CDate(vFecha) >= CDate(SQL) Then
        'fecha de compra posterior
        actualiza = True
    End If
    
    If actualiza Then
        SQL = "UPDATE sartic SET ultfecco=" & DBSet(vFecha, "F")
        'actualizar el ult. precio de compra si el nuevo valor no es 0
        If CCur(vPrecio) > 0 Then SQL = SQL & ",preciouc=" & DBSet(vPrecio, "N")
        SQL = SQL & " WHERE codartic=" & DBSet(codigo, "T", "N")
        
        conn.Execute SQL
    End If
    
EActFecha:
    If Err.Number <> 0 Then
        ActualizarUltFechaCompra = False
    Else
        ActualizarUltFechaCompra = True
    End If
End Function



Public Function ActualizarPrecioMedPond(vCantidad As Currency, vPrecio As Currency, Optional vCantAnt As Currency) As Boolean
'-- Modificar precio medio ponderado en la tabla sartic, se utiliza al crear
'-- albaranes de compra
'(IN) vCantidad=cantidad comprada o modificada
'(IN) vPrecio= precio de compra
'(IN) vCantAnt= cantidad anterior a la modificacion
Dim SQL As String
Dim actualiza As Boolean
Dim existencia As Currency
Dim Importe As Currency

    On Error GoTo EActPMP

    
    If PrecioMedPon = 0 Then
        'Obtenemos el precio medio ponderado que tiene el articulo
        'leemos de la tabla por si no se leyeron todos los campos de la sartic
        SQL = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", codigo, "T")
        If SQL <> "" Then PrecioMedPon = CCur(SQL)
    End If
    
    actualiza = False
    If PrecioMedPon = 0 Then
        'se actualiza con el precio de compra
        PrecioMedPon = vPrecio
        actualiza = True
    Else
        'se actualiza con la formula:
        '((existencia_total_almacenes_antes_compra * precio_medio_ponderado_actual) + importe_compra) / (existencia_total_almacenes_antes_compra + cantidad_compra)
        
        existencia = ExistenciaTotalAlmacenes
        'como aqui ya le habiamos sumado la cantidad q acabamos de comprar
        'se la quitamos para obtener la existencia antes de la compra
'        existencia = existencia - vCantidad

        'si hemos modificado la linea del albaran, a la existencia habrá q
        'quitarle la cantidad q habia antes
        If vCantAnt <> 0 Then existencia = existencia - vCantAnt
        
        'obtenemos el nuevo precio_medio_ponderado a partir del valor actual
        If (existencia + vCantidad) <> 0 Then
            Importe = Round2((vCantidad * vPrecio), 2)
            PrecioMedPon = Round2(((existencia * PrecioMedPon) + Importe) / (existencia + vCantidad), 4)
'        PrecioMedPon = Round(PrecioMedPon, 4)
            actualiza = True
        End If
        
        
        'ABril 2011
        'Stcks negativos
        'Si la cantidad anterior es menor que cero, entonces el pmp es importe / cantidad, es decir el precio actual
        If existencia < 0 Then
            Importe = Round2((vCantidad * vPrecio), 2)
            PrecioMedPon = Round2(Importe / vCantidad, 4)
            actualiza = True
        End If
    End If

    
    If actualiza Then
        SQL = "UPDATE sartic SET preciomp=" & DBSet(PrecioMedPon, "N")
        SQL = SQL & " WHERE codartic=" & DBSet(codigo, "T", "N")
        
        conn.Execute SQL
    End If
    
    ActualizarPrecioMedPond = True
    Exit Function
    
EActPMP:
'    If Err.Number <> 0 Then
        ActualizarPrecioMedPond = False
'    Else
'        ActualizarPrecioMedPond = True
'    End If
End Function






Public Function ReestablecerUltFechaCompra(vAlmac As String) As Boolean
'Modificar fecha ult.compra y precio ult.compra en la tabla sartic
'con los valores del ultimo movimiento que encontremos en la smoval
Dim SQL As String
Dim cadSel As String
Dim RS As ADODB.Recordset

    On Error GoTo ERestFecha

    'Obtenemos el ultimo movimiento de albaran de compra que tiene el articulo
    cadSel = " detamovi='ALC' and codartic=" & DBSet(codigo, "T") & " AND codalmac=" & vAlmac & " AND and tipomovi=1 "
    
    SQL = "select * from smoval WHERE " & cadSel
    SQL = SQL & " and horamovi =(select max(horamovi) from smoval where " & cadSel & ")"

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'actualizamos la sartic con esos valores
        SQL = "UPDATE sartic SET ultfecco=" & DBSet(RS!FechaMov, "F")
        SQL = SQL & ", preciouc=" & DBSet(RS!impormov, "N")
        SQL = SQL & " WHERE codartic=" & DBSet(codigo, "T")
        conn.Execute SQL
    End If
    RS.Close
    Set RS = Nothing

    
ERestFecha:
    If Err.Number <> 0 Then
        ReestablecerUltFechaCompra = False
    Else
        ReestablecerUltFechaCompra = True
    End If
End Function




Public Function ReestablecerPrecioMedPon(vCantidad As Currency, vPrecio As Currency) As Boolean
'reestablecer el valor del precio medio ponderado
'       (pmp_actual * stock_actual) - (vCantidad * vPrecio)
'pmp=   ---------------------------------------------------
'       (stock_actual - cantidad)

Dim existencia As Currency 'existencia de stock en almacenes
Dim SQL As String
Dim Importe As Currency

    On Error GoTo EResPMP
    
    ReestablecerPrecioMedPon = False
    
    If PrecioMedPon = 0 Then
        'Obtenemos el precio medio ponderado que tiene el articulo
        'leemos de la tabla por si no se leyeron todos los campos de la sartic
        SQL = DevuelveDesdeBDNew(conAri, "sartic", "preciomp", "codartic", codigo, "T")
        If SQL <> "" Then PrecioMedPon = CCur(SQL)
    End If
    
    
    existencia = ExistenciaTotalAlmacenes
    
    If (existencia - vCantidad) <> 0 Then
        Importe = Round2((vCantidad * vPrecio), 2)
        PrecioMedPon = Round2(((PrecioMedPon * existencia) - Importe) / (existencia - vCantidad), 4)
'        PrecioMedPon = Round(PrecioMedPon, 4)
        
        'Abril 2011
        If PrecioMedPon < 0 And existencia < 0 Then PrecioMedPon = mPreciouc
            
    Else
        PrecioMedPon = 0
    End If
    
    SQL = "UPDATE sartic SET preciomp=" & DBSet(PrecioMedPon, "N")
    SQL = SQL & " WHERE codartic=" & DBSet(codigo, "T", "N")
        
    conn.Execute SQL
    
    ReestablecerPrecioMedPon = True
    Exit Function
    
EResPMP:
    ReestablecerPrecioMedPon = False
    MsgBox Err.Number, "Reestablecer precio medio ponderado.", Err.Description
End Function




Public Function ObtenerPorceIVA() As Currency
'Devuelve el % de IVA del articulo correspondiente al codigiva
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo EPorce
    
    If codigo <> "" Then 'codartic
        'Obtenemos el % de IVA a aplicar
        SQL = "SELECT porceiva FROM tiposiva WHERE codigiva=" & TipoIVA
        Set RS = New ADODB.Recordset
        RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SQL = ""
        If Not RS.EOF Then SQL = RS.Fields(0).Value
        
        RS.Close
        Set RS = Nothing
        ObtenerPorceIVA = CCur(SQL)
    End If
    
EPorce:
    If Err.Number <> 0 Then
        SQL = "Se ha producido un error. " & "Obteniendo porcentaje de IVA." & vbCrLf
        SQL = SQL & "Número: " & Err.Number & vbCrLf
        SQL = SQL & "Descripción: " & Err.Description
        MsgBox SQL, vbExclamation
        Set RS = Nothing
    End If
End Function


''David
''Modifica:  Para que guarde en una variable el valor del % del iva
'Public Function ObtenerPrecioConIVA2(ByRef CadenaIVA As String, Optional newPrecio As String) As Currency
''Devuelve el precio del articulo tras aplicarle el IVA correspondiente
''si newPrecio tiene valor le aplicamos el IVA a ese precio ya que habremos
''introducido el precio del articulo manualmente
''si no le pasamos newPrecio se aplicará al precio_venta del articulo
'Dim SQL As String
'Dim Porce As Currency ' % de IVA
'
'    On Error GoTo EPrecio
'
'    If Codigo <> "" Then 'codartic
'        'Obtener el % de IvA q se le aplica al articulo
'        Porce = ObtenerPorceIVA
'        CadenaIVA = CStr(Porce)
'        If Porce <> CCur(0) Then
'            If newPrecio = "" Then
'                'al precio de venta del articulo
'                SQL = CalcularDto(PrecioVenta, CStr(Porce))
'                ObtenerPrecioConIVA2 = Round(CCur(ComprobarCero(PrecioVenta)) + CCur(ComprobarCero(SQL)), 4)
'            Else
'                'al precio especial del articulo (para un cliente tn promociones)
'                SQL = CalcularDto(newPrecio, CStr(Porce))
'                ObtenerPrecioConIVA2 = Round(CCur(ComprobarCero(newPrecio)) + CCur(ComprobarCero(SQL)), 4)
'            End If
'        Else
'            If newPrecio = "" Then
'                ObtenerPrecioConIVA2 = CCur(ComprobarCero(PrecioVenta))
'            Else
'                ObtenerPrecioConIVA2 = CCur(ComprobarCero(newPrecio))
'            End If
'        End If
'    End If
'
'EPrecio:
'    If Err.Number <> 0 Then
'        SQL = "Se ha producido un error. " & "Calculando precio con IVA." & vbCrLf
'        SQL = SQL & "Número: " & Err.Number & vbCrLf
'        SQL = SQL & "Descripción: " & Err.Description
'        MsgBox SQL, vbExclamation
'    End If
'End Function









Public Function ObtenerPrecioSinIVA(Optional newPrecio As String) As Currency
'Devuelve el precio del articulo tras quitarle el IVA correspondiente
'IN -> newPrecio: precio con IVA
'si newPrecio tiene valor le quitamos el IVA a ese precio ya que habremos
'introducido el precio del articulo manualmente
'si no le pasamos newPrecio será al precio_venta del articulo
Dim SQL As String
Dim Porce As Currency ' % de IVA
Dim PreuSinIVA As Currency

    On Error GoTo EPrecio
    
    If codigo <> "" Then 'codartic
        'Obtener el % de IvA q se le aplica al articulo
        Porce = ObtenerPorceIVA
        Porce = Porce / 100
        
        'precio al que le vamos a quitar el IVA
        If newPrecio = "" Then
            'al precio de venta del articulo
            PreuSinIVA = CCur(ComprobarCero(PrecioVenta))
        Else
            'al precio especial del articulo (para un cliente tn promociones)
            'o art. de varios y precio manual
            PreuSinIVA = CCur(ComprobarCero(newPrecio))
            If Porce <> CCur(0) Then
                'quitarle el IVA al precio
                PreuSinIVA = PreuSinIVA / (1 + Porce)
                PreuSinIVA = Round(PreuSinIVA, 4)
            End If
        End If
        ObtenerPrecioSinIVA = PreuSinIVA
    End If

'    curImporte = curImporte / (1 + curPorce) 'importe sin iva
'    curCuota = Round((curPorce * curImporte), 2)
'    curImporte = Round(curImporte, 2)
'
'    'valores que devuelve: Importe sin iva, cuota de iva
'    ImporteSinIVA = Format(curImporte, FormatoImporte)
'    sCuota = Format(curCuota, FormatoImporte)
'
    Exit Function
    
EPrecio:
    If Err.Number <> 0 Then
        SQL = "Se ha producido un error. " & "Calculando precio sin IVA." & vbCrLf
        SQL = SQL & "Número: " & Err.Number & vbCrLf
        SQL = SQL & "Descripción: " & Err.Description
        MsgBox SQL, vbExclamation
    End If
End Function


'Nuevo###
'Cambiado por David 14 Mayo 2008
' Si tiene descuentos linea, debe calcular el importe con respecto a esos descuentos
Public Function ObtPrecioParaCliente2(codCli As String, canti As String, FecVen As String, des1 As String, des2 As String) As String
'Obtiene el precio que se le va a aplicar a un cliente
'en la venta de un determinado articulo
'Habrá clientes con promociones, precios especiales, etc.

'IN -> codclien: cliente para el q comprobamos si hay precio especial
       'canti: cantidad a vender para saber si se aplica precio por caja
'     fecven: fecha de la venta

'OUT -> des, des2: descuento 1 y descuento 2

Dim cPrecioF As CPreciosFact
Dim NumCajas As Integer
Dim RestoUnid As Integer
Dim PorCaja As Boolean
Dim OrigP As String
Dim Precio As String
Dim cadMen As String
Dim Impo As Currency
Dim SiNoCancel


    On Error GoTo ErrPrecioCli

    'Obtener el precio correspondiente y los descuentos
        
        
    'Dejo los dtos a cero
    des1 = "0"
    des2 = "0"
    'Comprobar si el articulo se vende por cajas antes de entrar a la función
    If UnidCaja >= 0 Then '+ de 1 unidad x caja, significa q se vende por cajas
        Set cPrecioF = New CPreciosFact
        
        'Si se puede vender por cajas(unidcaja>1) poner numero de cajas en una linea con el
        'precio de caja, y otra linea con el resto unidades con precio unidad
'            Cantidad = txtAux(Index).Text
        NumCajas = cPrecioF.ObtenerNumCajas(canti, UnidCaja)
        RestoUnid = CInt(ComprobarCero(canti)) - NumCajas * CInt(UnidCaja)
            
        'Obtenemos la Tarifa del Cliente
        cPrecioF.CodigoArtic = codigo
        cPrecioF.CodigoClien = codCli
        
        cPrecioF.FijarTarifaActividad
        'cPrecioF.CodActividad = DevuelveDesdeBDNew(conAri, "sclien", "codactiv", "codclien", codCli, "N")
        
                
        PorCaja = (NumCajas > 0)
        Precio = cPrecioF.ObtenerPrecio(PorCaja, FecVen, OrigP, "")
        'Si PorCaja vuelve de ObtenerPrecio a false se calcula con precio unidad aunque NumCajas>0
        'Ya que a regresado con pvp del Articulo
        If PorCaja And NumCajas > 0 And RestoUnid > 0 Then
            cadMen = "El Artículo puede venderse por Cajas (" & UnidCaja & "uds. por Caja)." & vbCrLf
            cadMen = cadMen & vbCrLf & "¿Desea insertar dos Lineas?:   " & vbCrLf
            cadMen = cadMen & vbCrLf & "   Linea 1:  " & NumCajas * CInt(UnidCaja) & " uds a Precio Caja"
            cadMen = cadMen & vbCrLf & "   Linea 2:  " & CInt(canti) - NumCajas * CInt(UnidCaja) & " uds a Precio Unidad"
            
            Do
                SiNoCancel = MsgBox(cadMen, vbQuestion + vbYesNoCancel + vbDefaultButton3)
            Loop Until SiNoCancel = vbYes Or SiNoCancel = vbNo
            If SiNoCancel = vbYes Then
                canti = NumCajas * CInt(UnidCaja)
            Else
                'si no vendemos por cajas traer el precio de la funcion poniendo
                'venta por cajas a false.
                Precio = cPrecioF.ObtenerPrecio(False, FecVen, OrigP, "")
            End If
         
        Else
            If CCur(cPrecioF.Descuento1) > 0 Or CCur(cPrecioF.Descuento2) > 0 Then
                des1 = cPrecioF.Descuento1
                des2 = cPrecioF.Descuento2
            End If
        End If
        
        ObtPrecioParaCliente2 = Precio
        Set cPrecioF = Nothing
    End If
    Exit Function
    
ErrPrecioCli:
    MuestraError Err.Number, "Obtener precio para el cliente", Err.Description
End Function



Public Function TieneNumLote() As Boolean
'si el articulo pertenece a una categoria que tiene control de numero de lote
Dim Cad As String

    If categoria <> "" Then
        'ver si la categoria a la q pertenece lleva control de lotes
        Cad = DevuelveDesdeBDNew(conAri, "scateg", "ctrlotes", "codcateg", categoria, "T")
        TieneNumLote = (Cad = "1")
    Else
        'el articulo no pertenece a una categoria
        TieneNumLote = False
    End If
End Function



Public Function AplicarMargenComercial() As Currency
'Aplica el margen comercial que queremos obtener al precio de ultima compra
'y obtiene el nuevo precio de venta al publico.
Dim margen As Currency

    If Me.MargenComercial <> 0 Then 'el campo margen comercial tiene valor
'        margen = CalcularDto(Me.PrecioUltCom, Me.MargenComercial)
'        margen = Round(margen, 4)
        margen = Round2(((Me.PrecioUltCom * Me.MargenComercial) / 100), 4)
        AplicarMargenComercial = Me.PrecioUltCom + margen 'nuevo pvp
        
'    ElseIf Me.PrecioUltCom > Me.PrecioVenta Then
'        AplicarMargenComercial = Me.PrecioUltCom
        
    Else
        AplicarMargenComercial = Me.PrecioVenta
    End If
    
End Function


Public Function CambiaPrecioVenta(newPrecio As Currency, newMargen As Currency, fecCambio As String, cadError) As Boolean
'actualiza el precio de venta del articulo (PVP) y la fecha en que se realiza la modificacion
'actualiza el campo sartic.preciove
'actualiza el campo sartic.ultfecpvp
Dim SQL As String

    On Error GoTo ErrPrecio

    CambiaPrecioVenta = False
    If codigo <> "" Then
        SQL = "UPDATE sartic SET preciove=" & DBSet(newPrecio, "N")
        SQL = SQL & ", margecom=" & DBSet(newMargen, "N")
        
        '## LAURA 23/06/2008
        '   actualizar la fecha de ultima modificacion PVP
        SQL = SQL & ", ultfecpvp=" & DBSet(fecCambio, "F")
        '##
        
        SQL = SQL & " WHERE codartic=" & DBSet(codigo, "T")
        conn.Execute SQL
        CambiaPrecioVenta = True
    Else
        cadError = cadError & "Falta código del artículo para actualizar precio."
    End If
    
    Exit Function

ErrPrecio:
    cadError = cadError & "Cambiar precio artículo." & vbCrLf
    cadError = cadError & Err.Number & ": " & Err.Description
'    MuestraError Err.Number, "Cambiar precio artículo.", Err.Description
End Function



Private Function TarifasTieneMargen() As Boolean

End Function



'## Laura 22/11/2006
'Public Function ActualizarStockLinConjunto(ByRef cSto As CStock, cadError As String) As Boolean
''para el artículo q estamos actualizando su stock
'' comprobar si es de conjunto, y si lo es Actualizar stock para cada linea del conjunto
'Dim b As Boolean
'Dim SQL As String
'Dim RS As ADODB.Recordset
'Dim cStock1 As CStock 'stock de las sublineas de conjunto
'
'    On Error GoTo ErrActConj
'
'    If Me.EsConjunto Then
'        'para cada linea de conjunto del articulo actualizar stock e insertar movimiento
'        SQL = "SELECT * FROM sarti1 WHERE codartic=" & DBSet(Me.Codigo, "T")
'        Set RS = New ADODB.Recordset
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        While Not RS.EOF And b
''            Set cStock1 = New CStock
'            Set cStock1 = cSto
'            cStock1.Cantidad = cSto.Cantidad * DBLet(RS!codarti1, "N")
'            b = cStock1.ActualizarStock
'            'precios!!!!
'
'
'            Set cStock1 = Nothing
'            RS.MoveNext
'        Wend
'        Set RS = Nothing
'    Else
'        'si no es cjto no hacer nada salimo con ok
'        b = True
'    End If
'
'
'    ActualizarStockLinConjunto = b
'    Exit Function
'
'ErrActConj:
'    ActualizarStockLinConjunto = False
'    cadError = Err.Description
'End Function

