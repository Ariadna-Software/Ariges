VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAgente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CLIENTES
'TABLA: sclien


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodAgent As Integer
Private mNomAgent As String
Private mComsionNormal As Currency
Private mComsionEco As Currency
Private mComsionPVPMin As Currency


Public Property Get CodAgent() As Integer
     CodAgent = mCodAgent
End Property
Public Property Get NomAgent() As String
     NomAgent = mNomAgent
End Property

Public Property Get ComsionNormal() As Currency
     ComsionNormal = mComsionNormal
End Property
Public Property Get ComsionEco() As Currency
     ComsionEco = mComsionEco
End Property
Public Property Get ComsionPVPMin() As Currency
     ComsionPVPMin = mComsionPVPMin
End Property



Public Function LeerDatos(vAgent As String) As Boolean

'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo eLeer
    
    LeerDatos = False
    
    SQL = "SELECT codagent,nomagent,comisioc,comsio1n,comsiopvpmin FROM sagent "
    SQL = SQL & " WHERE codagent=" & vAgent
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        
    Else
        'mNomAgent  mComsionNormal mComsionEco mComsionPVPMin
        mCodAgent = CStr(RS!CodAgent)
        
        mNomAgent = RS!NomAgent
        mComsionNormal = DBLet(RS!comisioc, "N")
        mComsionEco = DBLet(RS!comsio1n, "N")
        mComsionPVPMin = DBLet(RS!comsiopvpmin, "N")
        LeerDatos = True
    End If
    RS.Close
    
     
eLeer:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function
