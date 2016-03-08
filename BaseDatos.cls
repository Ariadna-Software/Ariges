VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mtipo As String
Private conn As ADODB.Connection
Private mrs As ADODB.Recordset

Public Property Let Tipo(ByVal vData As String)
    mtipo = vData
End Property

Public Property Get Tipo() As String
    Tipo = mtipo
End Property


Public Function abrir(vODBC As String, vUsu As String, vPass As String) As Boolean
    '-- Abrimos la conexión en función del ODBC pasado
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;" & _
                            "Data Source=" & vODBC & _
                            ";UID=" & vUsu & _
                            ";PWD=" & vPass
    conn.Open
    abrir = True
End Function

Public Function asignar(ByRef cn As ADODB.Connection) As Boolean
    Set conn = New ADODB.Connection
    Set conn = cn
End Function


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

Public Function fecha(vData) As String
    '-- Devuelve un string con la fecha correspondiente
    If vData = "0:00:00" Then
        fecha = "NULL"
    Else
        If Not IsNull(vData) And IsDate(vData) Then
            Select Case mtipo
                Case "ACCESS"
                    fecha = "#" & Format(vData, "yyyy/mm/dd") & "#"
                Case "MYSQL"
                    fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
                Case "MULTIBASE"
                    fecha = "'" & Format(vData, "dd/mm/yyyy") & "'"
                Case Default
                    fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
            End Select
        Else
            fecha = "NULL"
        End If
    End If
End Function

Public Function hora(vData) As String
    '-- Devuelve un string con la fecha correspondiente
    If vData = "0:00:00" Then
        hora = "NULL"
    Else
        If Not IsNull(vData) And IsDate(vData) Then
            Select Case mtipo
                Case "ACCESS"
                    hora = "#" & Format(vData, "hh:mm:ss") & "#"
                Case "MYSQL"
                    hora = "'" & Format(vData, "hh:mm:ss") & "'"
                Case "MULTIBASE"
                    hora = "'" & Format(vData, "hh:mm") & "'"
                Case Default
            End Select
        Else
            hora = "NULL"
        End If
    End If
End Function

Public Function numero(vData, Optional decimales As Integer) As String
    If Not IsNull(vData) Then
        numero = CStr(vData)
        If decimales Then
            'numero = TransformaComasPuntos(CStr(vData / (10 ^ decimales)))
            numero = TransformaComasPuntos(Format((vData / (10 ^ decimales)), "0." & String(decimales, "0")))
        Else
            numero = TransformaComasPuntos(CStr(vData))
        End If
    Else
        numero = "0"
    End If
    If numero = "" Then numero = 0
End Function

Public Function Texto(vData) As String
Dim cad As String

    If Not IsNull(vData) Then
'        vData = CApos(CStr(vData))
        cad = (CStr(vData))
        NombreSQL cad
        Texto = "'" & cad & "'"
    Else
        Texto = "''"
    End If
End Function
