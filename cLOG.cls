VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'En el frmlog, grabamos la tabla tmpnlotes con las acciones para linkar con slog


'ACCIONES DISPONIBLES
'
'       1.- Eliminar facturas
'       2.- Facturacion
'       3.- Contabilizar facturas
'       4.- Correccion precios
'       5.- Facturacion desde hasta COMPRAS
'       6.- Facturacion TICKETS agrupados
'
'       7.- Eliminar articulo
'       8.- Modificar FACTURA
'       9.- Recepcion fra pro
'       10.-    Eliminar de hco cuando ya existe . para que no de entrada duplicada en BD
'       11.-    Traspaso OFERTAS a historico
'       12.-    DTO LINEA VENTA(ofer/ped/alb)  supera el maximo permitido en dfamia_dtos
'       13.-    Copia precios
'       14.-    cambio precio o descuento
'       15.-    OFERTA. Cuando pasa a pedido el precio del articulo ha cambiado desde que se hizo
'       16.-    Actualziacion de los dtos "a piñon" desde frmfamilias
'       17.-    Alb-Pedido con riesgo superado
'       18.-    Actualizar preciomp o precioUC
'       19.-    Recalculo PMP...
'       20.-    Cambio trabajador albaran
'       21.-    Cambio Nombre (proveedor/cliente)
'       22.-    Actualizar CCC en otras secciones

'       23.-    Cambio titularidad telefono
'       24.-    Cambio masivo cuotas/coperadora
'
'       25.- Modifi fecha factura
'       26.- Eliminar reestableciendo albaranes
'       27.- Cambiar familia marca
'       28.- OFE a FAZ (Herbelca)


'       29.-    VARIOS
'               Por ejemplo, en Bolbaite AGUA sera acepta la importacion del fichero
'

'       30 .- Linea pedido de cliente
'       31 .- Bloqueo cliente
'       32 .- Cambiar proveedor (herbelca)
'       33 .- Ajuste canstock desde movimientos articulos desde inventario
'       34 .- Eliminar albaran
'       35 .- Tabla numeros de lote. Modificar
'       36 .- Cambio agente pedido
'       37 .- Eliminar linea albaran
'       38. Cambiar Tipo alb.
Public Function Insertar(Accion As Byte, ByRef ElUsuario As Usuario, Descripcion As String) As Boolean
Dim SQL As String

    On Error GoTo EI
    Insertar = False
    
    SQL = "insert into `slog` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`) values ( "
    SQL = SQL & " now()," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    SQL = SQL & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        SQL = SQL & "NULL)"
    Else
        SQL = SQL & "'" & DevNombreSQL(Descripcion) & "')"
    End If
    conn.Execute SQL
    
    
    Insertar = True
    Exit Function
EI:
    MuestraError Err.Number, Err.Description

End Function




Public Function DevuelveAcciones(ByRef ListaAcciones As Collection) As Boolean
 On Error GoTo EDevuelveAcciones
    ListaAcciones.Add "1|Eliminar facturas|"
    ListaAcciones.Add "2|Facturación|"
    ListaAcciones.Add "3|Contabilizar facturas|"
    ListaAcciones.Add "4|Corrección precios|"
    ListaAcciones.Add "5|Facturacion COMPRAS|"
    ListaAcciones.Add "6|Facturacion TICKETS agrupados|"
    ListaAcciones.Add "7|Eliminar articulo|"
    ListaAcciones.Add "8|Modificar Factura|"
    ListaAcciones.Add "9|Recepcion FRA PROV.|"
    ListaAcciones.Add "10|Eli. automatica de HCO|"   'Cuando vamos a pasar a hco y ya existe
    ListaAcciones.Add "11|Traspaso ofertas a HCO|"   '
    ListaAcciones.Add "12|Dto superior|"   'Estamos dandole un dto superior al que tiene en la tabla de dtosxfamila
    ListaAcciones.Add "13|Copia precios|"   'copia precios de compra a vtas o viceversa
    ListaAcciones.Add "14|Cambio precio dto|"   'cambio precio o descuento
    ListaAcciones.Add "15|OFE. precios cambiados|"   ' OFERTA. Cuando pasa a pedido el precio del articulo ha cambiado desde que se hizo
    ListaAcciones.Add "16|Actualizar dto-fam-marca|"
    ListaAcciones.Add "17|Alb-Pedido con riesgo superado|"
    ListaAcciones.Add "18|Modificar cuenta contabilidad|"
    ListaAcciones.Add "19|Recalculo PMP-UPC|"
    ListaAcciones.Add "20|Cambio trabajador albaran|"
    ListaAcciones.Add "21|Nombre prov-clien|"
    ListaAcciones.Add "22|Actualizar CCC|"
    ListaAcciones.Add "23|Cambio titular telefono|"
    ListaAcciones.Add "24|Cambio masivo cuota/telefono|"
    ListaAcciones.Add "25|Cambiar fecha factura|"
    ListaAcciones.Add "26|Eliminar reestab. albaran|"
    ListaAcciones.Add "27|Cambiar familia marca|"
    ListaAcciones.Add "28|Oferta a presupuesto|"
    ListaAcciones.Add "29|Varios|"
    ListaAcciones.Add "30|Elim. linea pedido cliente|"
    ListaAcciones.Add "31|Bloqueo cliente|"
    ListaAcciones.Add "32|Cambio proveedor|"
    ListaAcciones.Add "33|Ajuste stock desde mov. inventario|"
    ListaAcciones.Add "34|Eliminar albaran|"
    ListaAcciones.Add "35|Modificar numero lote|"
    ListaAcciones.Add "36|Cambio agente pedido|"
    ListaAcciones.Add "37|Eliminar linea albaran|"
    ListaAcciones.Add "38|Cambiar tipo albaran|"
    DevuelveAcciones = True
    Exit Function
EDevuelveAcciones:
    MuestraError Err.Number
End Function

