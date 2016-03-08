VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNumLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TABLA: slotes
'Nº LOTES DE LOS ARTICULOS

'Variables locales que contienen valores de propiedad
Private mNumLotes As String 'num. lote
Private mCodartic As String 'cod. articulo
Private mFecEntra As String 'fecha de entrada
Private mCanEntra As Single 'cantidad de entrada




'------------------------------------------------
'Propiedades del modulo CNumLote
'------------------------------------------------

'**** num. Lote

Public Property Let numLote(ByVal vData As String)
     mNumLotes = vData
End Property

Public Property Get numLote() As String
     numLote = mNumLotes
End Property


'**** cod. Artic

Public Property Let Articulo(ByVal vData As String)
     mCodartic = vData
End Property

Public Property Get Articulo() As String
     Articulo = mCodartic
End Property


'**** Fecha de entrada

Public Property Let FechaEnt(ByVal vData As String)
     mFecEntra = vData
End Property

Public Property Get FechaEnt() As String
     FechaEnt = mFecEntra
End Property



'**** Cantidad de entrada

Public Property Let CantidadEnt(ByVal vData As Single)
     mCanEntra = vData
End Property

Public Property Get CantidadEnt() As Single
     CantidadEnt = mCanEntra
End Property



'------------------------------------------------
'Procedimientos del modulo CNumLote
'------------------------------------------------


Public Function LeerDatos(vCodArtic As String, vNumLote As String, vFecEntra As String) As Boolean
'Leer los datos de un numero de lote
'Lee de la BD: Ariges, Tabla: slotes
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    LeerDatos = False
    
    SQL = "SELECT codartic,numlotes,fecentra,canentra "
    SQL = SQL & " FROM slotes "
    SQL = SQL & " WHERE codartic=" & DBSet(vCodArtic, "T") & " AND numlotes=" & DBSet(vNumLote, "T")
    SQL = SQL & " AND fecentra=" & DBSet(vFecEntra, "F")
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        mCodartic = CStr(RS!codArtic)
        mNumLotes = RS!numlotes
        mFecEntra = DBLet(RS!fecentra, "F") 'fecha entrada
        
        mCanEntra = DBLet(RS!canentra, "N") 'cantidad de entrada

        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    Set RS = Nothing
    LeerDatos = False
    
    SQL = "Se ha producido un error. " & "Datos Nº de Lote." & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
End Function




Public Function Eliminar(Optional vCanti As Single) As Boolean
'eliminamos una linea de nº de lote de un artículo
'varías líneas del mismo día que tendrían la misma clave primaria podrían
'sumar cantidades a la misma línea, en ese caso al eliminar pasamos la cantidad
'de la linea a eliminar
Dim SQL As String

    On Error GoTo ErrElim
    
    vCanti = CantidadEnt - vCanti
    
    If vCanti = 0 Then
        'borrar
        'borramos toda la línea
        SQL = "DELETE FROM slotes  "
    Else
        'actualizar la cantidad
        SQL = "UPDATE slotes SET canentra=canentra + " & DBSet(vCanti, "N")
    End If
    SQL = SQL & " WHERE codartic=" & DBSet(Articulo, "T") & " AND numlotes=" & DBSet(numLote, "T") & " AND fecentra=" & DBSet(FechaEnt, "F")
    Conn.Execute SQL
    Eliminar = True
    Exit Function

ErrElim:
    Eliminar = False
    MuestraError Err.Number, "Eliminar nº de lote.", Err.Description
End Function
