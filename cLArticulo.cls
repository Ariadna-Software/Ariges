VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarcodartic As String 'nombre del cliente
Private mvarcantidadnecesaria As Single
Private mvarcantidadStock As Single


'Marzo 2010
'Punto de pedido y minimo
Private mvarcantidadMinimo As Single
Private mvarcantidadPuntoPedido As Single

'Marzo 2011
'Precio UC para cuando metemos en la smoval, que no meta el importe
Private mvarPrecioUC As Single


'Sept 2011
'Fecha inventario
Private mvarFecInve As Date

Public Property Let codartic(ByVal vData As String)
     mvarcodartic = vData
End Property

Public Property Get codartic() As String
     codartic = mvarcodartic
End Property


Public Property Let CantidadNecesaria(ByVal vData As Single)
     mvarcantidadnecesaria = vData
End Property

Public Property Get CantidadNecesaria() As Single
     CantidadNecesaria = mvarcantidadnecesaria
End Property


Public Property Let cantidadStock(ByVal vData As Single)
     mvarcantidadStock = vData
End Property

Public Property Get cantidadStock() As Single
     cantidadStock = mvarcantidadStock
End Property


Public Property Let cantidadMinimo(ByVal vData As Single)
     mvarcantidadMinimo = vData
End Property

Public Property Get cantidadMinimo() As Single
     cantidadMinimo = mvarcantidadMinimo
End Property

Public Property Let cantidadPuntoPedido(ByVal vData As Single)
     mvarcantidadPuntoPedido = vData
End Property

Public Property Get cantidadPuntoPedido() As Single
     cantidadPuntoPedido = mvarcantidadPuntoPedido
End Property

Public Property Let PrecioUC(ByVal vData As Single)
     mvarPrecioUC = vData
End Property

Public Property Get PrecioUC() As Single
     PrecioUC = mvarPrecioUC
End Property



Public Property Let FecInve(ByVal vData As Date)
     mvarFecInve = vData
End Property

Public Property Get FecInve() As Date
     FecInve = mvarFecInve
End Property



Public Sub Reestablecer()
    mvarcantidadStock = 0
    mvarcantidadnecesaria = 0
    mvarcodartic = ""
    mvarcantidadMinimo = 0
    mvarcantidadPuntoPedido = 0
    mvarPrecioUC = 0
    mvarFecInve = "01/01/1900"
End Sub
