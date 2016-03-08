VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TelBaseDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mtipo As String
Private mrs As ADODB.Recordset

Public Property Let Tipo(ByVal vData As String)
    mtipo = vData
End Property

Public Property Get Tipo() As String
    Tipo = mtipo
End Property





Public Function ejecutar(vSQL As String) As Boolean
    If Not (conn Is Nothing) Then
        conn.Execute vSQL
    Else
        '-- Falta la conexión y no podemos continuar
    End If
End Function

Public Function cursor(vSQL As String) As ADODB.Recordset
    '-- Devuelve un recordset con los valores obtenidos por SQL
    Set mrs = New ADODB.Recordset
    mrs.Open vSQL, conn, adOpenForwardOnly
    Set cursor = mrs
End Function

Public Function Fecha(vData) As String
    '-- Devuelve un string con la fecha correspondiente
    If Not IsNull(vData) And IsDate(vData) Then
        Select Case mtipo
            Case "ACCESS"
                Fecha = "#" & Format(vData, "yyyy/mm/dd") & "#"
            Case "MYSQL"
                Fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
            Case Default
        End Select
    Else
        Fecha = "NULL"
    End If
End Function

Public Function numero(vData, Optional decimales As Integer) As String
    If Not IsNull(vData) Then
        numero = CStr(vData)
        If decimales Then
            numero = TransformaComasPuntos(CStr(vData / (10 ^ decimales)))
        Else
            numero = TransformaComasPuntos(CStr(vData))
        End If
    Else
        numero = "0"
    End If
    If numero = "" Then numero = 0
End Function

Public Function texto(vData) As String
    If Not IsNull(vData) Then
        texto = "'" & CApos(CStr(vData)) & "'"
    Else
        texto = "''"
    End If
End Function
