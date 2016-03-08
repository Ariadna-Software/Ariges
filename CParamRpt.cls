VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros: SCRYST
Private mvarCodcryst As Integer  'Codigo del tipo de Documento
Private mvarNomcryst As String   'Descripción del Documento
Private mvarDocumRpt As String   'Documento .rpt asociado
Private mvarCodigiso As String   'Código ISO
Private mvarCodigrev As Integer  'Código Revisión (puede ser nulo)
Private mvarLineapi1 As String   'Linea pie 1
Private mvarLineapi2 As String   'Linea pie 2
Private mvarLineapi3 As String   'Linea pie 3
Private mvarLineapi4 As String   'Linea pie 4
Private mvarLineapi5 As String   'Linea pie 5
Private mvarImprimeDirecto As Boolean
Private mvarPDFrpt As String

Private mvarMultiInforme As Integer   'si lleva mutiinforme pondra el codigo de scryst



Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes


'------------------------------------------------
'Propiedades del modulo CParamRpt
'------------------------------------------------
Public Property Let Codigo(ByVal vData As Integer)
     mvarCodcryst = vData
End Property

Public Property Get Codigo() As Integer
     Codigo = mvarCodcryst
End Property


Public Property Let Descripcion(ByVal vData As String)
     mvarNomcryst = vData
End Property

Public Property Get Descripcion() As String
     Descripcion = mvarNomcryst
End Property

Public Property Let Documento(ByVal vData As String)
     mvarDocumRpt = vData
End Property

Public Property Get Documento() As String
     Documento = mvarDocumRpt
End Property

Public Property Let CodigoISO(ByVal vData As String)
     mvarCodigiso = vData
End Property

Public Property Get CodigoISO() As String
     CodigoISO = mvarCodigiso
End Property


Public Property Let CodigoRevision(ByVal vData As Integer)
     mvarCodigrev = vData
End Property

Public Property Get CodigoRevision() As Integer
     CodigoRevision = mvarCodigrev
End Property

Public Property Let LineaPie1(ByVal vData As String)
     mvarLineapi1 = vData
End Property

Public Property Get LineaPie1() As String
     LineaPie1 = mvarLineapi1
End Property

Public Property Let LineaPie2(ByVal vData As String)
     mvarLineapi2 = vData
End Property

Public Property Get LineaPie2() As String
     LineaPie2 = mvarLineapi2
End Property

Public Property Let LineaPie3(ByVal vData As String)
     mvarLineapi3 = vData
End Property

Public Property Get LineaPie3() As String
     LineaPie3 = mvarLineapi3
End Property

Public Property Let LineaPie4(ByVal vData As String)
     mvarLineapi4 = vData
End Property

Public Property Get LineaPie4() As String
     LineaPie4 = mvarLineapi4
End Property

Public Property Let LineaPie5(ByVal vData As String)
     mvarLineapi5 = vData
End Property

Public Property Get LineaPie5() As String
     LineaPie5 = mvarLineapi5
End Property


Public Property Let ImprimeDirecto(ByVal vData As Boolean)
     mvarImprimeDirecto = vData
End Property

Public Property Get ImprimeDirecto() As Boolean
     ImprimeDirecto = mvarImprimeDirecto
End Property



Public Property Let PDFrpt(ByVal vData As String)
     mvarPDFrpt = vData
End Property

Public Property Get PDFrpt() As String
     PDFrpt = mvarPDFrpt
End Property



'Solo GET
Public Property Get MultiInforme() As Integer
    MultiInforme = mvarMultiInforme
End Property



'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function Leer(vCodigo As Byte) As Byte
Dim RS As ADODB.Recordset
On Error GoTo Err3
    
    SQL = "SELECT * FROM scryst WHERE codcryst=" & vCodigo
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        Leer = 1
    Else
        mvarCodcryst = RS!codcryst
        mvarNomcryst = RS!nomcryst
        mvarDocumRpt = RS!documrpt
        mvarCodigiso = DBLet(RS!codigiso, "T")
        If IsNull(RS!codigrev) Then
            mvarCodigrev = -1
        Else
            mvarCodigrev = DBLet(RS!codigrev, "N")
        End If
        mvarLineapi1 = DBLet(RS!lineapi1, "T")
        mvarLineapi2 = DBLet(RS!lineapi2, "T")
        mvarLineapi3 = DBLet(RS!lineapi3, "T")
        mvarLineapi4 = DBLet(RS!lineapi4, "T")
        mvarLineapi5 = DBLet(RS!lineapi5, "T")
        mvarImprimeDirecto = DBLet(RS!ImprimeDirecto, "N") = 1
        mvarPDFrpt = RS!PDFrpt  'no pue ser nulo
        
        
        mvarMultiInforme = 0
        If DBLet(RS!LlevaMulitInformes, "N") = 1 Then mvarMultiInforme = mvarCodcryst
        
        
        Leer = 0
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Leer = 1
End Function


Public Function Modificar(vCodigo As Integer) As Byte
'Modifica los parametros de una empresa
'Modifica la Tabla: sparam, BD: Ariges
On Error GoTo Err4

    SQL = "UPDATE scryst SET "
    SQL = SQL & "nomcryst = '" & mvarNomcryst
    SQL = SQL & "', documrpt = '" & mvarDocumRpt
    SQL = SQL & "', codigiso = '" & mvarCodigiso
    If mvarCodigrev < 0 Then
        SQL = SQL & "', codigrev = null"
    Else
        SQL = SQL & "', codigrev = " & mvarCodigrev
    End If
    SQL = SQL & ", lineapi1 = '" & mvarLineapi1
    SQL = SQL & "', lineapi2 = '" & mvarLineapi2
    SQL = SQL & "', lineapi3 = '" & mvarLineapi3
    SQL = SQL & "', lineapi4 = '" & mvarLineapi4
    SQL = SQL & "', lineapi5 = '" & mvarLineapi5
    SQL = SQL & "', ImprimeDirecto = " & DBSet(mvarImprimeDirecto, "B")
    SQL = SQL & ", pdfrpt = " & DBSet(mvarPDFrpt, "T")
    SQL = SQL & " WHERE codcryst =" & vCodigo & ";"

    conn.Execute SQL

    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 1
End Function


Public Function Insertar() As Byte
'Modifica los parametros de una empresa
'Modifica la Tabla: sparam, BD: Ariges
On Error GoTo Err4

    SQL = "INSERT INTO scryst (codcryst, nomcryst, documrpt,codigiso, codigrev, lineapi1, lineapi2, lineapi3, lineapi4, lineapi5,ImprimeDirecto,pdfrpt) VALUES( "
    SQL = SQL & mvarCodcryst
    SQL = SQL & " ," & DBSet(mvarNomcryst, "T")
    SQL = SQL & "," & DBSet(mvarDocumRpt, "T")
    SQL = SQL & "," & DBSet(mvarCodigiso, "T")
    SQL = SQL & ", " & DBSet(mvarCodigrev, "N", "S")
'    If mvarCodigrev < 0 Then
'        SQL = SQL & ", null"
'    Else
'        SQL = SQL & "," & mvarCodigrev
'    End If
    SQL = SQL & " ," & DBSet(mvarLineapi1, "T")
    SQL = SQL & " ," & DBSet(mvarLineapi2, "T")
    SQL = SQL & " ," & DBSet(mvarLineapi3, "T")
    SQL = SQL & " ," & DBSet(mvarLineapi4, "T")
    SQL = SQL & " ," & DBSet(mvarLineapi5, "T")
    SQL = SQL & " ," & DBSet(ImprimeDirecto, "B")
    SQL = SQL & " ," & DBSet(PDFrpt, "T") & ")"
    conn.Execute SQL
    Insertar = 0
    Exit Function

Err4:
    Mens = "Insertar. Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Insertar = 1
End Function
