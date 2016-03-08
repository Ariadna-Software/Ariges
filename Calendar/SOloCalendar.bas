Attribute VB_Name = "SOloCalendar"
Public COnn As Connection
Public ConnConta As Connection
Public vEmpresa As Cempresa
Public vConfig As Configuracion
Public vUsu As Usuario
Public Const conAri = 1

Public Function Dblet(Campo As Variant, Optional Tipo As String) As Variant

End Function


Public Function ComputerName() As String

End Function

Public Function DevuelveDesdeBD(Kcon As Byte, vPC As String, vtab As String, Nom As String, mvarPC As String, Tipo As String) As String

End Function


Public Sub Main()
    If AbrirConexion Then
        Set vEmpresa = New Cempresa
        Set vConfig = New Configuracion
        If vConfig.Leer Then
            Set vUsu = New Usuario
            
        
            vUsu.Leer ("root")
            vUsu.CadenaConexion = "ariges2"
            If vEmpresa.LeerDatos = 0 Then frmMainCalendar.Show vbModal
    
        End If
    End If
End Sub


Public Function AbrirConexion() As Boolean
Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set COnn = Nothing
    Set COnn = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    COnn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

'        cad = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=accUPVMED"
'        cad = cad & ";UID=" & Usuario
'        cad = cad & ";PWD=" & Pass
'        Conn.ConnectionString = cad
    
    'cad = "DSN=plannertours;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=plannertours;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
    '---- Laura: 17/10/2006
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE= vAriges;DATABASE=ariges2;SERVER=PCDAVIDG"
'    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DSN= vAriges;DATABASE=" & vUsu.CadenaConexion
    
    cad = cad & ";UID=root"
    cad = cad & ";PWD=aritel"
    '---- Laura: 29/09/2006
    cad = cad & ";PORT=3306;OPTION=3;STMT=;"
    '----
   
'    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
'    cad = cad & ";UID=" & vConfig.User
'    cad = cad & ";PWD=" & vConfig.password
    
    COnn.ConnectionString = cad
    COnn.Open
    COnn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MsgBox Err.Description
End Function
