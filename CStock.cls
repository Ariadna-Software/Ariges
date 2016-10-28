VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=======================================================================
'MODULO para el control de Stock en Almacenes
'Comprueba si hay stock de un Articulo en un Almacen
'Comprueba si hay control de stock para permitir movimiento o no
'Inserta el movimiento de Entrada/salida de stock en la tabla "smoval"
'Actualiza el valor del stock tras el movimiento en la tabla "salmac"
'=======================================================================


'------------------------------------------------
'Variables del modulo CStock
'------------------------------------------------
Private mvarcodartic As String
Private mvarCodAlmac As Integer

Private mvarFechaMov As String
Private mvarHoraMov As String
Private mvarTipoMov As String '0=Salida, 1=Entrada
Private mvarDetaMov As String 'referencia a codtipom de la tabla stipom
Private mvarCantidad As Single 'Cantidad de Stock a mover
Private mvarImporte As Currency 'Importe de la linea de Movimiento de stock
Private mvarCodOpe As Long   'Codigo Trabajador/Cliente/Proveedor segun el tipo de Movimiento. ENERO 08: LONG
Private mvarDocum As String 'Codigo de la tabla de Movimientos, Traspaso, Albaran,... al que va asociado el movimiento
Private mvarLinea As Integer 'Numero de linea del documento




'MARZO 2010
'Punto de PEDIDO.




Private cArticulos As Collection
Private nI As Integer

'variable local
Dim MuestraMenCStock  As Boolean 'Indica si se muestra el mensaje en el control de Stock
                             'Cuando se pasa de Pedido a Albaran si alguna linea no tiene
                             'suficiente stock no se muestra mensaje, al final muestra detalle de articulos
                             'sin stock

Dim artPuntoPedido As String
Dim artsotckminimo As String

        'Para cuando ponga el punto de pedido(o minimo) lo muestre la ventana del msgbox

'------------------------------------------------
'Propiedades del modulo CStock
'------------------------------------------------

Public Property Let codArtic(ByVal vData As String)
     mvarcodartic = vData
End Property

Public Property Get codArtic() As String
     codArtic = mvarcodartic
End Property


Public Property Let codAlmac(ByVal vData As Integer)
     mvarCodAlmac = vData
End Property

Public Property Get codAlmac() As Integer
     codAlmac = mvarCodAlmac
End Property


Public Property Let FechaMov(ByVal vData As String)
     mvarFechaMov = vData
End Property

Public Property Get FechaMov() As String
     FechaMov = mvarFechaMov
End Property



Public Property Let HoraMov(ByVal vData As String)
     mvarHoraMov = vData
End Property

Public Property Get HoraMov() As String
     HoraMov = mvarHoraMov
End Property


Public Property Let tipoMov(ByVal vData As String)
     mvarTipoMov = vData
End Property

Public Property Get tipoMov() As String
     tipoMov = mvarTipoMov
End Property


Public Property Let DetaMov(ByVal vData As String)
     mvarDetaMov = vData
End Property

Public Property Get DetaMov() As String
     DetaMov = mvarDetaMov
End Property


Public Property Let cantidad(ByVal vData As Single)
     mvarCantidad = vData
End Property

Public Property Get cantidad() As Single
     cantidad = mvarCantidad
End Property


Public Property Let Importe(ByVal vData As Currency)
     mvarImporte = vData
End Property

Public Property Get Importe() As Currency
     Importe = mvarImporte
End Property


Public Property Let Trabajador(ByVal vData As Long)
     mvarCodOpe = vData
End Property

Public Property Get Trabajador() As Long
     Trabajador = mvarCodOpe
End Property


Public Property Let Documento(ByVal vData As String)
     mvarDocum = vData
End Property

Public Property Get Documento() As String
     Documento = mvarDocum
End Property


Public Property Let LineaDocu(ByVal vData As Integer)
     mvarLinea = vData
End Property

Public Property Get LineaDocu() As Integer
     LineaDocu = mvarLinea
End Property



'------------------------------------------------
'Metodos del modulo CStock
'------------------------------------------------


Private Function HayStock(ByRef ColArt As Collection, vCantidad As Single, ByRef PuntoPedidooMinimo As Byte) As Boolean
'IN: stockOrig: stock existente en almacen Origen
'    stockTras: stock a traspasar del origen a otro almacen
'OUT: True=Hay stock disponible en almacen para realizar el movimiento
'     False=No hay suficiento stock en el almacen
    Dim stockOrig As Single
    Dim stockTras As Single
    Dim stockPed As Single
    'Marzo 2010. Punto de pedido
        
    
    Dim B As Boolean

    PuntoPedidooMinimo = 0
    HayStock = False
    For nI = 1 To ColArt.Count
        stockOrig = ColArt(nI).cantidadStock
        stockTras = ColArt(nI).CantidadNecesaria * cantidad
    
        If stockTras > 0 Then
            If stockOrig >= stockTras Then
            'Si cantidad en stock > cantidad a traspasar entonces
                B = True
            Else
                B = False
            End If
            
            
            'Punto de pedido
            stockPed = ColArt(nI).cantidadPuntoPedido
            If stockPed > 0 Then
                
                'TIENE FIJADO EL PUNTO DE PEDIDO
                'Vemos cuanto se va a quedar el stock
                stockOrig = stockOrig - stockTras
                
                If stockOrig < CSng(ColArt(nI).cantidadPuntoPedido) Then
                    PuntoPedidooMinimo = 2
                    artPuntoPedido = artPuntoPedido & ColArt(nI).codArtic & "|"
                End If
            End If
            
            'No es menor que el punto de pedido
            If PuntoPedidooMinimo = 0 Then
                'Compruebo con el minimo
                'stockOrig = ColArt(nI).cantidadPuntoPedido
                stockOrig = ColArt(nI).cantidadStock
                If stockOrig > 0 Then
                    stockOrig = stockOrig - stockTras
                    'TIENE FIJADO EL PUNTO MINIMO
                    If stockOrig < CSng(ColArt(nI).cantidadMinimo) Then
                        PuntoPedidooMinimo = 1
                        artsotckminimo = artsotckminimo & ColArt(nI).codArtic & "|"
                    End If
                End If
            End If
        Else
            B = True
        End If
        If Not B Then Exit For
    Next nI
    
    HayStock = B
End Function


Private Function ControlStock2(CL As Collection) As Boolean
'Si hay control de Stock no se permite realizar el movimiento
'Sino hay control se pregunta si se permite realizar el movimiento
Dim devuelve As String
Dim B As Boolean
Dim NomArtic As String
Dim Rotacion As String

        
        devuelve = "Control de Stock  "
        If CL.Count > 1 Then devuelve = devuelve & "(Conjuntos)"
        devuelve = devuelve & ": " & vbCrLf & "----------------------" & vbCrLf & vbCrLf & vbCrLf
        devuelve = devuelve & " No hay suficiente Stock en el almacen  para: "
        For nI = 1 To CL.Count

                Rotacion = "concat(rotacion,""|"",codprove)"
                NomArtic = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", CL(nI).codArtic, "T", Rotacion)
                NomArtic = CL(nI).codArtic & " - " & NomArtic & " ("
                
                NomArtic = NomArtic & CL(nI).cantidadStock & ")"
                If nI > 1 Then NomArtic = "       L: " & NomArtic
                If nI = 2 Then NomArtic = "----------------------------" & vbCrLf & NomArtic
                devuelve = devuelve & vbCrLf & NomArtic
                
                
                'SOLO EL PRIMERO y en oferta
                If (DetaMov = "PEV" Or DetaMov = "OFE") And nI = 1 Then
                    If Mid(Rotacion, 1, 1) = "0" Then
                    
                        'NO es de rotacion
                        'Rotacion = "codprove = " & Val(Mid(Rotacion, 3)) & " AND codartic"
                        'Rotacion = DevuelveDesdeBDNew(conAri, "slispr", "cantmini", Rotacion, CL(nI).Codartic, "T")
                        
                        'Marzo 2012.
                        'Va a por UNICAJAS, no a por cantmini
                        Rotacion = DevuelveDesdeBDNew(conAri, "sartic", "unicajas", "codartic", CL(nI).codArtic, "T")
                        If Rotacion = "" Then Rotacion = "0"
                        
                        If Val(Rotacion) > 1 Then
                            'Va por cajas, tiene que pedir multiplos de unicajas
                                If (mvarCantidad Mod Val(Rotacion)) <> 0 Then
                                    'Rotacion = vbCrLf & "            Cantidad minima: " & Right("     " & Rotacion, 6) & vbCrLf & vbCrLf
                                    'Rotacion = String(43, "*") & vbCrLf & String(43, "*") & vbCrLf & Rotacion & String(43, "*") & vbCrLf & String(43, "*")
                                    'devuelve = devuelve & vbCrLf & vbCrLf & Rotacion
                        
                                    'Marzo 2012
                                    Rotacion = "            Unidades caja: " & Right("     " & Rotacion, 6) & vbCrLf
                                    Rotacion = vbCrLf & Rotacion & Space(25) & "Stock: " & Right("     " & CL(nI).cantidadStock, 6) & vbCrLf & vbCrLf
                                    Rotacion = String(43, "*") & vbCrLf & String(43, "*") & vbCrLf & Rotacion & String(43, "*") & vbCrLf & String(43, "*")
                                    devuelve = "Material de NO rotacion. " & vbCrLf & vbCrLf & vbCrLf & devuelve & vbCrLf & vbCrLf & Rotacion
                                End If
                        Else
                            'Mayo 2012
                            'HERBELCA
                            'Aunque  se cajas =1 hay que indicar que el articulo no es de rotacion
                            If vParamAplic.NumeroInstalacion = 2 Then devuelve = "Material de NO rotacion. " & vbCrLf & vbCrLf & vbCrLf & devuelve & vbCrLf & vbCrLf
                        End If
                    End If
                End If
        Next nI

        
                
        
        
        
        If DetaMov = "OFE" Or DetaMov = "PEV" Or DetaMov = "PEW" Or DetaMov = "PEZ" Then 'No hay control de stock para bloquear solo informa
            If HayArticulosEquivalenciaConStock Then
                LanzaEquivalencias CL(1).codArtic, devuelve
            Else
                MsgBox devuelve, vbInformation
            End If
            B = False
        Else
            If vParamAplic.ControlStock Then
                'Si hay control Stock no permitir traspaso
                B = True   '¿True?
                If MuestraMenCStock Then
                    Select Case DetaMov
                        Case "REG"
                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Movimiento de Almacen. "
                        Case "TRA"
                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Traspaso de Almacen. "
                        Case "ALV", "ALR", "ALC", "ALI", "ALS"
                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Albaran. "
                    End Select
                    MsgBox devuelve, vbExclamation
                End If
            Else
                Select Case DetaMov
                Case "REG"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Movimiento de Almacen? "
                Case "TRA"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Traspaso de Almacen? "
                Case "ALV", "ALR", "ALC", "ALI", "ALS"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Albaran? "
                Case "PRO"
                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar continuar con la producción? "
                End Select
                
                
                'If Not MuestraMenCStock Then Stop
                
                
                If MsgBox(devuelve, vbQuestion + vbYesNo) = vbYes Then
                    B = False
                Else
                    B = True
                End If
            End If
        End If
        ControlStock2 = B
End Function


Private Function ObtenerStockNuevo3(ByRef ColArticulos As Collection) As Boolean
'OUT: cantidad de Stock que hay en el almacen vColAlm del articulo vCodArtic
'Modificacion  19 Nov.  NO msgbox. Ya que si es conjunto lo mostrare al final
Dim Cant As String
Dim devuelve As String
Dim RT As ADODB.Recordset


    devuelve = ""
    Set RT = New ADODB.Recordset
    
    For nI = 1 To ColArticulos.Count
'        Debug.Print ColArticulos.item(nI).codArtic
        'Marzo 2010
        'cant = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", ColArticulos.item(nI).codartic, "T", , "codalmac", CStr(codAlmac), "N")
        Cant = "Select canstock,stockmin,puntoped,fechainv from salmac WHERE codartic = " & DBSet(ColArticulos.item(nI).codArtic, "T") & " AND codalmac = " & CStr(codAlmac)
        RT.Open Cant, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RT.EOF Then
            
            Cant = vbCrLf & "     .-"
            'Significa que el articulo tiene conjuntos y es del conjunto
            If nI > 1 Then devuelve = devuelve & "(Lin) "
            Cant = Cant & ColArticulos.item(nI).codArtic
            devuelve = devuelve & Cant
        Else
            'SI que tiene estock
            ColArticulos.item(nI).cantidadStock = DBLet(RT!CanStock, "N")
            ColArticulos.item(nI).cantidadPuntoPedido = DBLet(RT!puntoped, "N")
            ColArticulos.item(nI).cantidadMinimo = DBLet(RT!stockmin, "N")
            If IsNull(RT!FechaINV) Then
                ColArticulos.item(nI).FecInve = "01/01/1900"
            Else
                ColArticulos.item(nI).FecInve = RT!FechaINV
            End If
                
        End If
        RT.Close
        
    Next
    Set RT = Nothing
    
    'ANtes solo comprobaba el primero
    'cant = DevuelveDesdeBDNew(conAri, "salmac", "canstock", "codartic", codartic, "T", , "codalmac", CStr(codAlmac), "N")
    If devuelve <> "" Then 'No hay de ese artículo(s) en Almacen
        devuelve = "No existe en el Almacen: " & codAlmac & vbCrLf
        devuelve = devuelve & "Artículo " & vbCrLf
        MsgBox devuelve, vbExclamation
        ObtenerStockNuevo3 = False
    Else
        ObtenerStockNuevo3 = True
    End If
    
End Function





Public Function MueveStock() As Boolean
Dim vArtic As CArticulo
Dim B As Boolean

    Set vArtic = New CArticulo
    If vArtic.LeerDatos(codArtic) Then
        If vArtic.EsInstalacion Then
            'Si el articulo es una instalacion no mueve stock
            B = False
        Else
            ' ver si hay control de stock para el articulo
            If vArtic.CtrStock Then
'                    vStock = ObtenerStock
                    B = True
            Else
                B = False
            End If
        End If
   End If
   Set vArtic = Nothing
   MueveStock = B
End Function



Public Function MoverStock(YaTieneLeidoLosDatosObjeto As Boolean, ActualizandoBD As Boolean, Optional NoMenStock As Boolean) As Boolean
Dim vStock As String
Dim B As Boolean
Dim devuelve As String
Dim vArtic As CArticulo
Dim vD As cLArticulo
Dim MinimoPedido As Byte
    '0.- Nada   1.- Minimo   2.- Punto pedido
    MuestraMenCStock = Not NoMenStock
       
    MoverStock = False
    
    'Comprobar si hay suficiente stock para realizar el Movimiento de stock si
    'el articulo NO es INSTALACION
    Set vArtic = New CArticulo
    If vArtic.LeerDatos(codArtic) Then
        If vArtic.EsInstalacion Then
            B = True
        Else
                        
     
            'meto el primer articulo ppal.
            
            'si hay control de stock para el articulo
            If vArtic.CtrStock Then
                If Not YaTieneLeidoLosDatosObjeto Then
                    Set cArticulos = New Collection
                    Set vD = New cLArticulo
                    vD.codArtic = vArtic.codigo
                    vD.CantidadNecesaria = 1
                    cArticulos.Add vD
                
                    'No se si deberia ir en la clase articulos o en esta
                    If vArtic.EsConjunto And Not vParamAplic.Produccion Then AñadirArticulosConjunto vD
                End If
                
                'Aqui leemos del almacen en articulo PPAL
                'vStock = ObtenerStock2
                'If (vStock = "Error") Then Exit Function
                
                'Ademas de comprobar la existencia en BD
                If Not ObtenerStockNuevo3(cArticulos) Then
                    If Not YaTieneLeidoLosDatosObjeto Then
                        Set cArticulos = Nothing
                        Set vD = Nothing
                        Exit Function
                    End If
                End If
                
                
                artPuntoPedido = ""
                artsotckminimo = ""
                'If HayStock(cArticulos, Cantidad, MinimoPedido) Then
                B = HayStock(cArticulos, cantidad, MinimoPedido)
                    
                If Not B Then
                    'No hay stock para el movimiento. Ver control de Stock
                    B = Not (ControlStock2(cArticulos))
                Else
                    'Ha ido bien , pero debe realizar pedidos (o esta en minimo)
                    If MinimoPedido > 0 Then
                        'Mostrar mensaje de punto pedido o minimo
                        'Ademas, si esta insertando en BD
                        'insertaremos en la tmp
                        MostrarMensajePuntopedidoMinimo cArticulos, ActualizandoBD, Not NoMenStock
                    End If

                End If
                
            Else
                B = False
            End If

            
            'ComprobarFechaInventario True, ""


            If Not YaTieneLeidoLosDatosObjeto Then
                Set vD = Nothing
                Set cArticulos = Nothing
            End If
        End If
    End If
    Set vArtic = Nothing
    MoverStock = B
End Function


Private Sub MostrarMensajePuntopedidoMinimo(CL As Collection, InsertaEnBD As Boolean, MuestraMSG As Boolean)
Dim devuelve As String
Dim Cad As String
Dim Esta As Boolean
Dim Minimo As Boolean
Dim SQL As String  'Para insertar en la tabla de pedidos
Dim Can As Single
        
        If artPuntoPedido <> "" Then artPuntoPedido = "|" & artPuntoPedido
        If artsotckminimo <> "" Then artsotckminimo = "|" & artsotckminimo
        devuelve = ""
        For nI = 1 To CL.Count
                Cad = "|" & CL(nI).codArtic & "|"
                If InStr(1, artsotckminimo, Cad) > 0 Then
                    Esta = True
                    Minimo = True
                Else
                    If InStr(1, artPuntoPedido, Cad) > 0 Then
                        Esta = True
                        Minimo = False
                    Else
                        Esta = False
                    End If
                End If
             
             
                If Esta Then
                    
                    Cad = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", CL(nI).codArtic, "T")
                    Cad = CL(nI).codArtic & " - " & Cad '& " ("
                    If Minimo Then
                        Cad = "MINIMO STOCK   " & Cad
                        'Cad = Cad & CL(nI).cantidadStock & ")"
                    Else
                        Cad = "Ha alcanzado el PUNTO de PEDIDO    " & Cad
                        'Cad = Cad & CL(nI).cantidadStock & ")"
                    End If
                    If nI > 1 Then Cad = "       L: " & Cad
                    If nI = 2 Then Cad = String(40, "-") & vbCrLf & Cad
                    devuelve = devuelve & vbCrLf & Cad
                    
                    
                    'Insertamos en tm
                    If InsertaEnBD Then
                            SQL = "insert into `spedidos` (`fechahora`,`codusu`,`codalmac`,`codartic`,`minimo`,`puntope`,"
                            SQL = SQL & "`cantidad`,`proceso`) values (concat(curdate(),' ' ,curtime())," & vUsu.codigo
                            SQL = SQL & "," & mvarCodAlmac
                            SQL = SQL & "," & DBSet(CL(nI).codArtic, "T")
                            SQL = SQL & "," & DBSet(CL(nI).cantidadMinimo, "N")
                            SQL = SQL & "," & DBSet(CL(nI).cantidadPuntoPedido, "N")
                            
                            'Necesaria
                            Can = CL(nI).cantidadStock
                            Can = CL(nI).CantidadNecesaria * Can
                            SQL = SQL & "," & DBSet(Can, "N")

                            
                            Cad = mvarTipoMov & " "
                            Cad = Cad & mvarDetaMov & ": " & mvarDocum & "  lin" & mvarLinea
                            SQL = SQL & ",'" & Cad & "')"
                            ejecutar SQL, False
                       
                    End If
                    
                End If
        Next nI

        If devuelve = "" Then Exit Sub
        
        If Not MuestraMSG Then Exit Sub
                
                
        If vParamAplic.NumeroInstalacion = 2 Then Exit Sub
                
        'Mostramos el msgh
        devuelve = vbCrLf & String(40, "-") & vbCrLf & devuelve & vbCrLf & vbCrLf
        If CL.Count > 1 Then devuelve = devuelve & "(Conjuntos)" & vbCrLf
        devuelve = "STOCKs.  Punto pedido / minimo  " & devuelve
                
        MsgBox devuelve, vbInformation
        
End Sub


Public Function ActualizarStock(Optional EsTicket As Boolean, Optional OcultarElMsgbox As Boolean) As Boolean
'Actualiza el stock en la tabla smoval y en la tabla salmac
'EsTicket: si estamos generando un ticket no se comprueba el stock auque se quede negativo
Dim B As Boolean
Dim ADonde As String, cadMen As String
Dim CadValues As String 'cadena para la SQL de insertar en la tabla salmac
Dim vTipoMov As String
Dim cart As Collection
Dim vD As cLArticulo
Dim miCantidad As Single
Dim miImporte As Single
Dim F_inve As String
    ActualizarStock = False
    B = False
    On Error GoTo EActualizarStock
    
    MuestraMenCStock = Not OcultarElMsgbox
    
    'Veremos si el articulo es conjunto
    '-------------------------------------------------------------------------------
    Set vD = New cLArticulo
    Set cArticulos = New Collection
    vD.codArtic = mvarcodartic
    vD.CantidadNecesaria = 1
    
    


    
    If mvarDetaMov <> "ALC" Then
        cArticulos.Add vD  'lo añado aqui
        cadMen = DevuelveDesdeBDNew(conAri, "sartic", "conjunto", "codartic", mvarcodartic, "T")
        If cadMen = "" Then cadMen = "0"
        If Val(cadMen) = 1 And Not vParamAplic.Produccion Then AñadirArticulosConjunto vD
        
    Else
        F_inve = DevuelveDesdeBD(conAri, "fechainv", "salmac", "codartic=" & DBSet(mvarcodartic, "T") & " AND codalmac", CStr(mvarCodAlmac), "N")
        If F_inve = "" Then F_inve = "01/01/1900"
        vD.FecInve = F_inve
        cArticulos.Add vD 'aqui lo añado
    End If

    If tipoMov = "S" Then
        vTipoMov = "0" 'Movimiento de Salida de Almacen
    Else
        vTipoMov = "1" 'Movimiento de Entrada de Almacen
    End If
        
    '---- Insertamos en Movimientos Artículos
    ADonde = "Insertando datos en Movimientos de Articulos (tabla: smoval)."
    If FechaMov = "" Then FechaMov = Now
    'Construir la cadena VALUES para insertar en la tabla "smoval"
    For nI = 1 To cArticulos.Count
        
        miCantidad = cArticulos(nI).CantidadNecesaria * cantidad
        If nI = 1 Then
            'La primera es el articulo ppal. El importe es el que es
            miImporte = Importe
        Else
            miImporte = cArticulos(nI).CantidadNecesaria * cantidad * cArticulos(nI).precioUC
        
        End If
        CadValues = "VALUES (" & DBSet(cArticulos(nI).codArtic, "T") & ", " & codAlmac & ", " & DBSet(FechaMov, "F") & ", '"
        If HoraMov = "" Then HoraMov = Format(FechaMov, "dd/mm/yyyy") & " " & Format(Now, "hh:mm:ss")
        CadValues = CadValues & Format(HoraMov, "yyyy-mm-dd hh:mm:ss") & "', " & vTipoMov & " , '" & DetaMov & "', " & DBSet(miCantidad, "N") & ", "
        CadValues = CadValues & DBSet(miImporte, "N") & ", " & Trabajador & ", " & ValorNulo & ", " & DBSet(Documento, "T") & ", " & LineaDocu & ")"
        
        B = InsertarMovimArticulos(CadValues)
        
        If Not B Then
            ADonde = ADonde & "Linea: " & nI
            Exit For
        End If
    Next nI
    
    '---- Modificar stock del articulo
    If B Then
        ADonde = "Actualizando Stocks Almacenes (tabla: salmac)."
        If MueveStock Then
            If vTipoMov = "1" Then 'Entrada: no se comprueba si hay stock
'                If ActualizarStockAlmacen() Then b = True
                B = ActualizarStockAlmacen()
            Else
                'If DetaMov = "FTI" Then
                If EsTicket Then
                    'En los tickets se actualiza el stock aunque no haya
                    'no se paraliza la venta
'                    If ActualizarStockAlmacen() Then b = True
                    B = ActualizarStockAlmacen()
                Else
                    'comprobar si hay stock, y si no hay si permito que
                    'se mueva o no.
                    If MoverStock(True, True, OcultarElMsgbox) Then
'                        If ActualizarStockAlmacen() Then b = True
                        B = ActualizarStockAlmacen()
                    Else
                        B = False
                    End If
                End If
            End If
        End If
    End If
    
    
    '---- si es articulo de conjunto actualizar stock e insertar movimientos
    'para cada una de las líneas de conjunto
    'if b then b=ActualizarStockLinConjunto
    
    
    ActualizarStock = B
    Exit Function
    
EActualizarStock:
    If Err.Number <> 0 Then
        cadMen = "Actualizar Stock." & vbCrLf & "----------------------------" & vbCrLf
        cadMen = cadMen & ADonde
        ActualizarStock = False
        MuestraError Err.Number, cadMen, Err.Description
    End If
End Function


Private Function InsertarMovimArticulos(CadValues As String) As Boolean
Dim SQL As String
On Error Resume Next

    SQL = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
    SQL = SQL & CadValues
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarMovimArticulos = False
    Else
        InsertarMovimArticulos = True
    End If
End Function


Private Function ActualizarStockAlmacen() As Boolean
Dim SQL As String
Dim ElArticulo As String
Dim miCantidad As Single
On Error Resume Next

    ActualizarStockAlmacen = False
    For nI = 1 To cArticulos.Count
        ElArticulo = DBSet(cArticulos(nI).codArtic, "T")
        miCantidad = cArticulos(nI).CantidadNecesaria * cantidad
        If cantidad <> 0 Then
            If tipoMov = "E" Then 'Mov. de Entrada
                '==== Aumentar el stock en el Almacen
                SQL = "UPDATE salmac Set canstock = canstock + " & DBSet(miCantidad, "N")
                SQL = SQL & " WHERE codartic =" & ElArticulo & " AND "
                SQL = SQL & " codalmac =" & codAlmac
            Else 'Mov. de Salida
                '==== Disminuir Stock en Almacen Origen
                SQL = "UPDATE salmac Set canstock = canstock - " & DBSet(miCantidad, "N")
                SQL = SQL & " WHERE codartic =" & ElArticulo & " AND "
                SQL = SQL & " codalmac =" & codAlmac
            End If
            conn.Execute SQL
        End If
        If Err.Number <> 0 Then
            'Hay error , almacenamos y salimos
            ActualizarStockAlmacen = False
        Else
            ActualizarStockAlmacen = True
        End If
    Next nI
    
End Function


Public Function DevolverStock2() As Boolean
'Cuando se elimina una linea de movimiento de stock:
'Actualiza el stock en la tabla "salmac" aumentando de nuevo la cantidad de la linea
'Eliminar la linea que se introdujo en la "smoval"
Dim B As Boolean
Dim ADonde As String, cadMen As String
Dim vD As cLArticulo

    DevolverStock2 = False
    B = False
    
    On Error GoTo EDevolverStock
    
    'Veremos si el articulo es conjunto
    '-------------------------------------------------------------------------------
    Set vD = New cLArticulo
    Set cArticulos = New Collection
    vD.codArtic = mvarcodartic
    vD.CantidadNecesaria = 1  'el principal siempre es uno
    cArticulos.Add vD

    cadMen = DevuelveDesdeBDNew(conAri, "sartic", "conjunto", "codartic", mvarcodartic, "T")
    If cadMen = "" Then cadMen = "0"
    'ANTES NOVIEMBRE 2013
    'If Val(cadMen) = 1 And Not vParamAplic.Produccion Then AñadirArticulosConjunto vD
    If Val(cadMen) = 1 And Not vParamAplic.Produccion Then
        If mvarDetaMov <> "ALC" Then AñadirArticulosConjunto vD
    End If
    
    
    
    'Modificar stock
    ADonde = "Actualizando Stocks Almacenes (tabla: salmac)."
    If MueveStock Then
        If ActualizarStockAlmacen() Then B = True
    Else
        B = True
    End If
    
    If B Then
        'Eliminar en Movimientos Artículos
        ADonde = "Eliminando datos en Movimientos de Articulos (tabla: smoval)."
        If EliminarMovimArticulos Then B = True
    End If
    
    DevolverStock2 = B
    
EDevolverStock:
    If Err.Number <> 0 Then
        cadMen = "Actualizar Stock." & vbCrLf & "----------------------------" & vbCrLf
        cadMen = cadMen & ADonde
        MuestraError Err.Number, cadMen, Err.Description
        DevolverStock2 = False
    End If
    Set vD = Nothing
    Set cArticulos = Nothing
End Function


Private Function EliminarMovimArticulos() As Boolean
Dim SQL As String
On Error GoTo EElim

    For nI = 1 To cArticulos.Count
        SQL = DBSet(cArticulos(nI).codArtic, "T")
        SQL = "DELETE FROM smoval WHERE codartic=" & SQL & " AND codalmac=" & codAlmac
        SQL = SQL & " AND detamovi='" & DetaMov & "' AND document= " & DBSet(Documento, "T") & " AND numlinea= " & LineaDocu
        SQL = SQL & " AND fechamov=" & DBSet(FechaMov, "F")
        'en albaranes de compra añadir el proveedor
        If DetaMov = "ALC" Then
            SQL = SQL & " AND codigope=" & Me.Trabajador 'en compras aqui guardamos el proveedor
        End If
        
        
        conn.Execute SQL
    
    
    Next nI
    EliminarMovimArticulos = True
    Exit Function
    
EElim:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        EliminarMovimArticulos = False
    
    End If
    Err.Clear
End Function


Public Function ModificarStock(oldCant As Single) As Boolean
'Cuando se modifica una linea de movimiento de stock:
'Modificar la linea que se introdujo en la "smoval"
'Actualiza el stock en la tabla "salmac" aumentando/disminuyendo la diferencia de cantidad de la linea
Dim B As Boolean
Dim ADonde As String, cadMen As String
Dim difCant As Single, difTipoM As String
Dim vD  As cLArticulo

    ModificarStock = False
    B = False
    
    On Error GoTo EModificarStock
    
    'Modificar en Movimientos Artículos (tabla: smoval)
    ADonde = "Modificando datos en Movimientos de Articulos (tabla: smoval)."
    If ModificarMovimArticulos Then B = True
    
    If B Then
        
        Set vD = New cLArticulo
        Set cArticulos = New Collection
        vD.codArtic = mvarcodartic
        vD.CantidadNecesaria = 1  'el principal siempre es uno
        cArticulos.Add vD
        
        
        'Modificar stock
        ADonde = "Modificando Stocks Almacenes (tabla: salmac)."
        'Obtener si se suma o resta la diferencia de cantidad
        difCant = ObtenerDiferenciaStock(oldCant, difTipoM)
        mvarCantidad = difCant
        mvarTipoMov = difTipoM
        B = ActualizarStockAlmacen()
    End If
    ModificarStock = B
    
EModificarStock:
    If Err.Number <> 0 Then
        cadMen = "Actualizar Stock." & vbCrLf & "----------------------------" & vbCrLf
        cadMen = cadMen & ADonde
        MuestraError Err.Number, cadMen, Err.Description
        ModificarStock = False
    End If
End Function


Private Function ModificarMovimArticulos() As Boolean
'Modifica la cantidad y el importe en la tabla "smoval"
Dim SQL As String
On Error Resume Next
    
    SQL = "UPDATE smoval SET cantidad=" & DBSet(cantidad, "N") & ", impormov=" & DBSet(Importe, "N") 'Importe
    SQL = SQL & " WHERE codartic=" & DBSet(codArtic, "T") & " AND codalmac=" & codAlmac
    SQL = SQL & " AND detamovi='" & DetaMov & "' AND document= " & DBSet(Documento, "T") & " AND numlinea= " & LineaDocu
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        ModificarMovimArticulos = False
    Else
        ModificarMovimArticulos = True
    End If
End Function


Private Function ObtenerDiferenciaStock(oldCant As Single, difTipoM As String) As Single
'Cuando se modifica la cantidad de una linea hay que reestablecer el stock en la
'tabla "salmac" aumentando o restando la diferencia
'IN -> oldCant: cantidad que habia antes de modificar la cantidad
'OUT -> difTipoMov: Indicara si se devera aumentar o disminuir el stock (Entrada o Salida)
'       RETURN: cantidad a aumentar o disminuir en almacen: salmac
Dim TipoM As Byte
Dim difCant As Single
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo EDifStock

    'Obtener el tipo de movimiento en la smoval
    SQL = "SELECT tipomovi FROM smoval "
    SQL = SQL & " WHERE codartic=" & DBSet(codArtic, "T") & " AND codalmac=" & codAlmac
    SQL = SQL & " AND detamovi='" & DetaMov & "' AND document= " & DBSet(Documento, "T") & " AND numlinea= " & LineaDocu
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        TipoM = RS.Fields(0).Value
    End If
    RS.Close
    Set RS = Nothing
    
    'En funcion del tipo de Movimiento y la diferencia de stock se sumara o se resta la diferencia de stock
    difCant = oldCant - cantidad
    If TipoM = 1 Then 'Entrada
        If difCant >= 0 Then 'restar la diferencia
            difTipoM = "S"
        Else 'sumar la diferencia
            difTipoM = "E"
        End If
    ElseIf TipoM = 0 Then 'Salida
        If difCant >= 0 Then
            'sumar la diferencia
            difTipoM = "E"
        Else 'restar la diferencia
            difTipoM = "S"
        End If
    End If
    ObtenerDiferenciaStock = Abs(difCant)
    
EDifStock:
    If Err.Number <> 0 Then
        ObtenerDiferenciaStock = -1 'error
    End If
End Function



'## Laura 22/11/2006
'Private Function ActualizarStockArticConjunto() As Boolean
'Dim cArt As CArticulo
'Dim b As Boolean
'Dim menError As String
'
'    On Error GoTo ErrActConj
'
'    b = False
'    Set cArt = New CArticulo
'    b = cArt.LeerDatos(Me.codArtic) 'articulo q estamos actualizando stock
'    If b Then
'        If cArt.EsConjunto Then
'            'actualizar stock de las lineas de conjunto
'            b = cArt.ActualizarStockLinConjunto(Me, menError)
'        Else
'            b = True 'no es conjunto y salimos con ok
'        End If
'    End If
'    Set cArt = Nothing
'
'    ActualizarStockArticConjunto = b
'    Exit Function
'
'ErrActConj:
'    ActualizarStockArticConjunto = False
'    MuestraError Err.Number, "Actualizar stock artículo de conjunto", Err.Description
'End Function



Private Sub AñadirArticulosConjunto(ByRef vDa As cLArticulo)
Dim RS As ADODB.Recordset
Dim C As String


    'No controlo el error. Si da error que pete en la sub que lo llama
    'On Error GoTo EAñadirArticulosConjunto
    'Antes marzo 2011
    'C = "Select codarti1,cantidad  from sarti1 where codartic = '" & mvarcodartic & "' order by numlinea"
    
    'ahora
    C = "Select codarti1,cantidad,preciouc  from sarti1,sartic where  sarti1.codarti1=sartic.codartic and sarti1.codartic= '" & mvarcodartic & "' order by numlinea"
    Set RS = New ADODB.Recordset
    RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Articulo marcado como conjunto pero no tiene lineas", vbExclamation
    Else
        While Not RS.EOF
            Set vDa = New cLArticulo
            vDa.CantidadNecesaria = RS!cantidad
            vDa.codArtic = RS!codarti1
            vDa.precioUC = DBLet(RS!precioUC, "N")
            cArticulos.Add vDa
            RS.MoveNext
            Set vDa = Nothing
        Wend
    End If
    RS.Close
    Set RS = Nothing

End Sub





Private Sub Class_Terminate()
    Set cArticulos = Nothing
End Sub



Public Function ComprobarFechaInventario(MostrarMsg As Boolean, LineaErrorFecha As String) As Boolean
Dim NF As Integer
Dim Cad As String
        
        If cArticulos Is Nothing Then
            'Solo comprobaremos del articulo este
            Cad = "codartic=" & DBSet(Me.codArtic, "T") & " AND codalmac "
            Cad = DevuelveDesdeBD(conAri, "fechainv", "salmac", Cad, CStr(codAlmac))
            If Cad <> "" Then
                If Cad <> "0:00:00" Then
                    'OK
                    If CDate(Cad) > CDate(FechaMov) Then
                        Cad = codArtic & "  inventariado en " & Cad
                    Else
                        Cad = ""
                    End If
                Else
                    Cad = "" 'ok
                End If
            End If
        
        
        Else
            Cad = ""
            For NF = 1 To cArticulos.Count
                    If CDate(Me.FechaMov) < CDate(cArticulos.item(NF).FecInve) Then
                        Cad = Cad & cArticulos.item(NF).codArtic & "  inventariado en " & cArticulos.item(NF).FecInve & vbCrLf
                    End If
            Next
        End If
        
        
        If Cad <> "" Then
            If MostrarMsg Then
                Cad = "Fecha inventario posterior." & vbCrLf & Cad
                Cad = String(40, "*") & vbCrLf & Cad & vbCrLf & vbCrLf & String(40, "*")
                MsgBox Cad, vbExclamation
            Else
                LineaErrorFecha = LineaErrorFecha & vbCrLf & Cad
            End If
        End If
End Function




'-----------------------------------------------------------------------
Private Function HayArticulosEquivalenciaConStock() As Boolean
Dim R As ADODB.Recordset
Dim Aux As String


    
    
    HayArticulosEquivalenciaConStock = False
    Aux = "select * from salmac where codalmac=" & mvarCodAlmac & "  and codartic in (select codarti1 from sarti6 where codartic=" & DBSet(mvarcodartic, "T") & ")   "
    Set R = New ADODB.Recordset
    R.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not R.EOF Then
            HayArticulosEquivalenciaConStock = True
    End If
    R.Close
    Set R = Nothing
End Function


Private Sub LanzaEquivalencias(arti As String, mensa As String)
    frmFacEquivalencias.vCodartic = arti & "|" & mvarCodAlmac & "|"
    frmFacEquivalencias.Text1 = mensa
    frmFacEquivalencias.Show vbModal
End Sub
