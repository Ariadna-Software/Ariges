VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAridoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Private mvarcodigo As Long 'copia local
Private mvarcodext As Integer 'copia local
Private mvarcodcarpeta As Integer 'copia local
Private mvarcampo1 As String 'copia local
Private mvarcampo2 As String 'copia local
Private mvarcampo3 As String 'copia local
Private mvarcampo4 As String 'copia local
Private mvarfecha1 As Date 'copia local
Private mvarfecha2 As Date 'copia local
'Private mvarfecha3 As Date 'copia local
Private mvarimporte1 As Currency 'copia local
'Private mvarimporte2 As Currency 'copia local
'Private mvarobserva As String 'copia local
Private mvartamnyo As Currency 'copia local
Private mvaruserprop As Long 'copia local
Private mvargroupprop As Long 'copia local

'Private mvarlecturau As Long 'copia local
'Private mvarlecturag As Long 'copia local
'Private mvarescriturau As Long 'copia local
'Private mvarescriturag As Long 'copia local

Private mvarcadUpdate As String






Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset 'Para la base de datos
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo
'------------------------------------------------


Public Property Let Codigo(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codigo = 5
     mvarcodigo = vData
End Property

Public Property Get Codigo() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codigo
     Codigo = mvarcodigo
End Property


Public Property Let codext(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codext = 5
     mvarcodext = vData
End Property

Public Property Get codext() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codext
     codext = mvarcodext
End Property



Public Property Let codcarpeta(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codcarpeta = 5
     mvarcodcarpeta = vData
End Property

Public Property Get codcarpeta() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codcarpeta
     codcarpeta = mvarcodcarpeta
End Property


Public Property Let campo1(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.campo1 = 5
     mvarcampo1 = vData
End Property

Public Property Get campo1() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.campo1
     campo1 = mvarcampo1
End Property


Public Property Let campo2(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.campo2 = 5
     mvarcampo2 = vData
End Property

Public Property Get campo2() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.campo2
     campo2 = mvarcampo2
End Property


Public Property Let campo3(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.campo3 = 5
     mvarcampo3 = vData
End Property

Public Property Get campo3() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.campo3
     campo3 = mvarcampo3
End Property


Public Property Let campo4(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.campo4 = 5
     mvarcampo4 = vData
End Property

Public Property Get campo4() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.campo4
     campo4 = mvarcampo4
End Property


Public Property Let fecha1(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fecha1 = 5
     mvarfecha1 = vData
End Property

Public Property Get fecha1() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fecha1
     fecha1 = mvarfecha1
End Property


Public Property Let fecha2(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fecha2 = 5
     mvarfecha2 = vData
End Property

Public Property Get fecha2() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fecha2
     fecha2 = mvarfecha2
End Property

Public Property Let importe1(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.importe1 = 5
     mvarimporte1 = vData
End Property

Public Property Get importe1() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.importe1
     importe1 = mvarimporte1
End Property


Public Property Let cadUpdate(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.campo4 = 5
     mvarcadUpdate = vData
End Property

Public Property Get cadUpdate() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.campo4
     cadUpdate = mvarcadUpdate
End Property
