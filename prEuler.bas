Attribute VB_Name = "prEuler"
Option Explicit

Public NumRegElim As Long
'Conexión a la BD Ariges de la empresa
Public conn As ADODB.Connection
Public ConnConta As ADODB.Connection
Public Const ValorNulo = "Null"
Public vUsu As Usuario

Public Const vbTaxco = 8

Public FormatoFecha As String
Public Const conAri As Byte = 1 'Si conAri entonces trabajaremos con conexion conn a la BD ARIGES
Public Const conConta As Byte = 2 'Si conConta entonces trabajaremos con conexion connConta a la BD CONTA


Public miRsAux As ADODB.Recordset
Public NombreCheck As String

Public vbMyMonday




Public Type vParamAplicDef
   HayDeparNuevo As Integer
   DireccionesEnvio As Boolean
   NumeroInstalacion As Integer
End Type

Public vParamAplic As vParamAplicDef


Public Type vParamDef
   CifEmpresa As String
End Type

Public vParam As vParamDef

Public EulerParam As String

Public Function AbrirConexion() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set conn = Nothing
    Set conn = New Connection
'    Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente

'        cad = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=accUPVMED"
'        cad = cad & ";UID=" & Usuario
'        cad = cad & ";PWD=" & Pass
'        Conn.ConnectionString = cad
    
    'cad = "DSN=plannertours;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=plannertours;UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    
    '---- Laura: 17/10/2006
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE=vAriges;DATABASE=Ariges1;"
    If False Then
        Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE=vAriges;DATABASE=Ariges7;"
    End If
    Cad = Cad & ";"   'UID=" & vConfig.User
    Cad = Cad & ";"   'PWD=" & vConfig.password
    Cad = Cad & ";Persist Security Info=true"
    
    conn.ConnectionString = Cad
    conn.Open
    conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión BD:Ariges.", Err.Description
End Function





'Public Function AbrirConexionUsuarios() As Boolean
'Dim Cad As String
'On Error GoTo EAbrirConexion
'
'
'    AbrirConexionUsuarios = False
'    Set conn = Nothing
'    Set conn = New Connection
'    'Conn.CursorLocation = adUseClient
'    conn.CursorLocation = adUseServer
'    'Cad = "DSN=vUsuarios;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=usuarios;"
'    'Cad = Cad & "SERVER=" & vConfig.SERVER & ";UID=" & vConfig.User & ";PASSWORD=" & vConfig.password & ";PORT=3306;OPTION=3;STMT=;"
'
'    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=usuarios;SERVER=" & vConfig.SERVER
'
'    Cad = Cad & ";UID=" & vConfig.User
'    Cad = Cad & ";PWD=" & vConfig.password
'    Cad = Cad & ";OPTION=3;STMT=;Persist Security Info=true"
'
'    conn.ConnectionString = Cad
'    conn.Open
'    AbrirConexionUsuarios = True
'    Exit Function
'EAbrirConexion:
'    MuestraError Err.Number, "Abrir conexión usuarios.", Err.Description
'End Function
'
'
'


Public Sub Main()

        FormatoFecha = "yyyy-mm-dd"
        
        Set vUsu = New Usuario
        
        vParam.CifEmpresa = "B20899563"
        vParamAplic.DireccionesEnvio = False
        vParamAplic.HayDeparNuevo = 0
        vParamAplic.NumeroInstalacion = 1
         
       
        
        
        
        If AbrirConexion() = False Then
            MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
            End
        Else
            'Carga Parametros Generales y Contables de la empresa
            'LeerParametros
        End If
         EulerParam = DevuelveDesdeBD(conAri, "pathDocs", "eulerparam", "1", "1")
        
        vUsu.Leer "root"
        frmEulerReloj.Show vbModal
        End
End Sub



Public Sub MuestraError(numero As Long, Optional CADENA As String, Optional Desc As String)
    Dim Cad As String
    
    
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    Cad = "Se ha producido un error: " & vbCrLf
    If CADENA <> "" Then
        Cad = Cad & vbCrLf & CADENA & vbCrLf & vbCrLf
    End If
    
    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
    
    If numero <> 513 Then Cad = Cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
    
    MsgBox Cad, vbExclamation
End Sub


Public Function Espera(Segundos As Single)
Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function



Public Function DBSet(vData As Variant, Tipo As String, Optional esNULO As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
'Tipos
'       T
'       N
'       F
'       H
'       FH
'       B
'       S   single O DOUBLE. sINGLE DE MOMENTO.    MAYO 2009
Dim Cad As String
Dim ValorNumericoCero As Boolean

    On Error GoTo Error1

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If
    
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If esNULO = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        Cad = (CStr(vData))
                        NombreSQL Cad
                        DBSet = "'" & Cad & "'"
                    End If
                    
                Case "N", "S"   'Numero  y  SINGLE
                    
                    If CStr(vData) = "" Then
                        ValorNumericoCero = True
                    
                    Else
                        If Tipo = "S" Then
                            ValorNumericoCero = CSng(vData) = 0
                        Else
                            ValorNumericoCero = CCur(vData) = 0
                        End If
                    End If
                    
                    If ValorNumericoCero Then
                        If esNULO <> "" Then
                            If esNULO = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        If Tipo = "N" Then
                            Cad = CStr(ImporteFormateado(CStr(vData)))
                        Else
                            'Sngle
                            Cad = CStr(ImporteFormateadoSingle(CStr(vData)))
                        End If
                        DBSet = TransformaComasPuntos(Cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If esNULO = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If

                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If esNULO = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If

                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
Error1:
    If Err.Number <> 0 Then MuestraError Err.Number, "Formato para la BD.", Err.Description
End Function



Public Sub NombreSQL(ByRef CADENA As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String

    J = 1
    '-- (RAFA/ALZIRA) 07052006
    Do
        I = InStr(J, CADENA, "\")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    

    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    
End Sub

Public Function DevNombreSQL(CADENA As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = CADENA
End Function



'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir numérico
Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
    End If
End Function
Public Function ImporteFormateadoSingle(Importe As String) As Single
Dim I As Integer

    If Importe = "" Then
        ImporteFormateadoSingle = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateadoSingle = Importe
    End If
End Function




'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
Dim I As Integer
    Do
        I = InStr(1, CADENA, ",")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "." & Mid(CADENA, I + 1)
        End If
    Loop Until I = 0
    TransformaComasPuntos = CADENA
End Function



Public Function ejecutar(ByRef SQL As String, OcultarMsg As Boolean) As Boolean
    On Error Resume Next
    conn.Execute SQL
    If Err.Number <> 0 Then
        If Not OcultarMsg Then MuestraError Err.Number, Err.Description, SQL
        ejecutar = False
    Else
        ejecutar = True
    End If
End Function



Public Function DevuelveDesdeBD(vBD As Byte, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef otroCampo As String) As String
    Dim Rs As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    Cad = "Select " & kCampo
    If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
    Cad = Cad & " FROM " & Ktabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
'    Debug.Print cad
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    If vBD = 1 Then 'BD 1: Ariges
        Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        Stop
        'RS.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function




Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = "0"
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function




Public Sub CargarCombo_Tabla(ByRef Cbo As ComboBox, NomTabla As String, NomCodigo As String, nomDescrip As String, Optional strWhere As String, Optional ItemNulo As Boolean, Optional Ordenacion As String)
'Carga un objeto ComboBox con los registros de una Tabla
'(IN) cbo: ComboBox en el q se van a cargar los datos
'(IN) nomTabla: nombre de la tabla de la q leeremos los datos a cargar
'(IN) nomCodigo: nombre del campo codigo de la tabla q queremos cargar
'(IN) nomDescrip: nombre del campo descripcion de la tabla a cargar
'(IN) strWhere: para filtrar los registros de la tabla q queremos cargar
'(IN) ItemNulo: si es true se añade el primer item con linea en blanco
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim I As Integer

    On Error GoTo ErrCombo
    
    Cbo.Clear
    
    SQL = "SELECT " & NomCodigo & "," & nomDescrip & " FROM " & NomTabla
    If strWhere <> "" Then SQL = SQL & " WHERE " & strWhere
    SQL = SQL & " ORDER BY "
    
    If Ordenacion <> "" Then
        SQL = SQL & Ordenacion
    Else
        SQL = SQL & nomDescrip
    End If
    
    
'    If AbrirRecordset(SQL, RS) Then
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    '- si valor del parametro ItemNulo=true hay que añadir linea en blanco
    If Not Rs.EOF And ItemNulo Then
        Cbo.AddItem "  "
        Cbo.ItemData(Cbo.NewIndex) = 0
    End If
    
    If Not Rs.EOF Then
        If IsNumeric(Rs.Fields(0).Value) Then
            '- si el codigo NomCodigo es numerico en el ItemData se carga el campo clave primaria
            '- y en List la descripcion NomDescrip
            While Not Rs.EOF
              Cbo.AddItem Rs.Fields(1).Value 'descrip
              Cbo.ItemData(Cbo.NewIndex) = Rs.Fields(0).Value 'codigo
              Rs.MoveNext
            Wend
        Else
            '- si el codigo NomCodigo en alfanumerico no se puede cargar
            '- el codigo en ItemData y cargamos un indice ficticio
            '- y en el List el campo codigo NomCodigo
            I = 1
            While Not Rs.EOF
              Cbo.AddItem Rs.Fields(0).Value 'campo del codigo
              Cbo.ItemData(Cbo.NewIndex) = I
              I = I + 1
              Rs.MoveNext
            Wend
        End If
    End If
'    End If
    
'    CerrarRecordset RS
    Rs.Close
    Set Rs = Nothing
    Exit Sub
    
ErrCombo:
    MuestraError Err.Number, "Cargar combo." & NomTabla, Err.Description
End Sub


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef CADENA As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim Cad As String

    I = 0
    cont = 1
    Cad = ""
    Do
        J = I + 1
        I = InStr(J, CADENA, "|")
        If I > 0 Then
            If cont = Orden Then
                Cad = Mid(CADENA, J, I - J)
                I = Len(CADENA) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = Cad
End Function


Public Sub PonerFoco(ByRef Text As TextBox)
On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoBtn(ByRef btn As CommandButton)
On Error Resume Next
    If btn.Visible Then btn.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub KEYpressGnral(KeyAscii As Integer, Modo As Byte, cerrar As Boolean)
'IN: codigo keyascii tecleado, y modo en que esta el formulario
'OUT: si se tiene que cerrar el formulario o no
    cerrar = False
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then cerrar = True
    End If
End Sub

Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte, Optional cadkey As Integer)
'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1
'en los formularios de Mantenimiento
On Error Resume Next

    If Modo = 5 Then Exit Sub
    
    If (Modo <> 0 And Modo <> 2) Then
        If Modo = 1 Then
            Text.BackColor = vbYellow  'Modo 1: Busqueda
        Else
            If Text.Locked Then 'si el control esta bloqueado pasamos el foco al sig. campo
                Text.BackColor = &H80000018 'amarillo claro
                 If cadkey = 0 Then cadkey = 40
                 KEYdown cadkey
                 Exit Sub
            Else
                Text.BackColor = vbWhite
            End If
        End If
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub KEYdown(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            SendKeys "+{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Function PerderFocoGnral(ByRef Text As TextBox, Modo As Byte) As Boolean
Dim Comprobar As Boolean
On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnral = False
        Exit Function
    End If
    
    With Text
        'Quitamos blancos por los lados
        .Text = Trim(.Text)
        
        If .BackColor = vbYellow Then .BackColor = vbWhite
        
        'Si no estamos en modo: 3=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
        If (Modo <> 3 And Modo <> 4 And Modo <> 1) Then
            PerderFocoGnral = False
            Exit Function
        End If
        
        If Modo = 1 Then
            'Si estamos en modo busqueda y contiene un caracter especial no realizar
            'las comprobaciones
            Comprobar = ContieneCaracterBusqueda(.Text)
            If Comprobar Then
                PerderFocoGnral = False
                Exit Function
            End If
        End If
        PerderFocoGnral = True
    End With
    If Err.Number <> 0 Then Err.Clear
End Function




Public Function ContieneCaracterBusqueda(CADENA As String) As Boolean
'Comprueba si la cadena contiene algun caracter especial de busqueda
' >,>,>=,: , ....
'si encuentra algun caracter de busqueda devuelve TRUE y sale
Dim B As Boolean
Dim I As Integer
Dim CH As String


    'Febrero 2012, el 29
    'NULL
    If UCase(CADENA) = "NULL" Then
        ContieneCaracterBusqueda = True
        Exit Function
    End If

    'For i = 1 To Len(cadena)
    I = 1
    B = False
    Do
        CH = Mid(CADENA, I, 1)
        Select Case CH
            Case "<", ">", ":", "="
                B = True
            Case "*", "%", "?", "_", "\", ":" ', "."
                B = True
            Case Else
                B = False
        End Select
    'Next i
        I = I + 1
    Loop Until (B = True) Or (I > Len(CADENA))
    ContieneCaracterBusqueda = B
End Function





Public Function PonerFormatoEntero(ByRef T As TextBox) As Boolean
'Comprueba que el valor del textbox es un entero y le pone el formato
Dim mTag As CTag
Dim Cad As String
Dim Formato As String
On Error GoTo EPonerFormato

    If T.Text = "" Then Exit Function
    PonerFormatoEntero = True
    
    Set mTag = New CTag
    mTag.Cargar T
    If mTag.Cargado Then
       Cad = mTag.Nombre 'descripcion del campo
       Formato = mTag.Formato
    End If
    Set mTag = Nothing

    If Not EsEnteroNew(T.Text) Then
        PonerFormatoEntero = False
        MsgBox "El campo " & Cad & " tiene que ser un número entero.", vbExclamation
        PonerFoco T
    Else
         T.Text = Format(T.Text, Formato)
    End If
    
EPonerFormato:
    If Err.Number <> 0 Then Err.Clear
End Function


'*********** LAURA : 13/09/2005
Public Function EsEnteroNew(texto As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEnteroNew = False

    If Not IsNumeric(texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, texto, ".")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 0 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, texto, ",")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 0 Then res = False
        End If
    End If
    EsEnteroNew = res
End Function





Public Function EsFechaOK(T As String) As Boolean
Dim Cad As String
Dim mes As String, dia As String
    
    Cad = T
    If InStr(1, Cad, "/") = 0 Then
       'debe ser una cadena tipo:020105 y la convertimos a 02/01/05
       If Not IsNumeric(Cad) Then
            EsFechaOK = False
            Exit Function
       End If
        
      '==== Anade: Laura 04/02/2005 =============
        If Len(Cad) < 6 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el dia es correcto, valores entre 1-31
        dia = Mid(Cad, 1, 2)
        If dia < 1 Or dia > 31 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el mes es correcto, valores entre 1-12
        mes = Mid(Cad, 3, 2)
        If mes < 1 Or mes > 12 Then
            EsFechaOK = False
            Exit Function
        End If
      '============================================
        
        If Len(T) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T) = 6 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        End If
    Else
        dia = Mid(Cad, 1, 2)
        mes = Mid(Cad, 4, 2)
    End If
    
    If IsDate(Cad) Then
        EsFechaOK = True
        T = Format(Cad, "dd/mm/yyyy")
      '==== Añade: Laura 08/02/2005
        If Month(T) <> Val(mes) Then EsFechaOK = False
        If Day(T) <> Val(dia) Then EsFechaOK = False
      '====
    Else
        EsFechaOK = False
    End If
End Function



Public Sub limpiar(ByRef formulario As Form)
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub











Public Function ObtenerBusqueda(ByRef formulario As Form, paraRPT As Boolean, Optional CHECK As String) As String
Dim Control As Object
Dim Carga As Boolean
Dim mTag As CTag
Dim Aux As String
Dim Cad As String
Dim SQL As String
Dim tabla As String, columna As String
Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.Visible Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        If Not paraRPT Then
                            Cad = " MAX(" & mTag.columna & ")"
                        Else
                            Cad = " MAX({" & mTag.tabla & "." & mTag.columna & "})"
                        End If
                    Else
                        If Not paraRPT Then
                            Cad = " MIN(" & mTag.columna & ")"
                        Else
                            Cad = " MIN({" & mTag.tabla & "." & mTag.columna & "})"
                        End If
                    End If
                    If Not paraRPT Then
                        SQL = "Select " & Cad & " from " & mTag.tabla
                    Else
                        SQL = "Select " & Cad & " from {" & mTag.tabla & "}"
                    End If
                    SQL = ObtenerMaximoMinimo(SQL)
                    
                    Select Case mTag.TipoDato
                    Case "N"
                        If Not paraRPT Then
                            SQL = mTag.tabla & "." & mTag.columna & " = " & TransformaComasPuntos(SQL)
                        Else
                            SQL = "{" & mTag.tabla & "." & mTag.columna & "} = " & TransformaComasPuntos(SQL)
                        End If
                    Case "F"
                        If Not paraRPT Then
                            SQL = mTag.tabla & "." & mTag.columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        Else
                            SQL = "{" & mTag.tabla & "." & mTag.columna & "} = '" & Format(SQL, "yyyy-mm-dd") & "'"
                        End If
                    Case Else
                        If Not paraRPT Then
                            SQL = mTag.tabla & "." & mTag.columna & " = '" & SQL & "'"
                        Else
                            SQL = "{" & mTag.tabla & "." & mTag.columna & "} = '" & SQL & "'"
                        End If
                    End Select
                    SQL = "(" & SQL & ")"
                End If
            End If
        End If
    Next

    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.Visible Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Not paraRPT Then
                        SQL = mTag.tabla & "." & mTag.columna & " is NULL"
                    Else
                        SQL = "{" & mTag.tabla & "." & mTag.columna & "} is NULL"
                    End If
                    SQL = "(" & SQL & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If (TypeOf Control Is TextBox) And Control.Visible Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If mTag.Cargado Then
                    Aux = Trim(Control.Text)
                    Aux = QuitarCaracterEnter(Aux) 'Si es multilinea quitar ENTER
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            If Not paraRPT Then
                                tabla = mTag.tabla & "."
                            Else
                                tabla = "{" & mTag.tabla & "."
                            End If
                        Else
                            tabla = ""
                        End If
                        If Not paraRPT Then
                            columna = mTag.columna
                        Else
                            columna = mTag.columna & "}"
                        End If
                    Rc = SeparaCampoBusqueda(mTag.TipoDato, tabla & columna, Aux, Cad, paraRPT)
                    If Rc = 0 Then
                        If SQL <> "" Then SQL = SQL & " AND "
                        If Not paraRPT Then
                            SQL = SQL & "(" & Cad & ")"
                        Else
                            SQL = SQL & "(" & Cad & ")"
                        End If
                    End If
                End If
            End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If

        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    If mTag.TipoDato <> "T" Then
                        Cad = Control.ItemData(Control.ListIndex)
                        If Not paraRPT Then
                            Cad = mTag.tabla & "." & mTag.columna & " = " & Cad
                        Else
                            Cad = "{" & mTag.tabla & "." & mTag.columna & "} = " & Cad
                        End If
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    Else
                        Cad = Control.List(Control.ListIndex)
                        If Not paraRPT Then
                            Cad = mTag.tabla & "." & mTag.columna & " = '" & Cad & "'"
                        Else
                            Cad = "{" & mTag.tabla & "." & mTag.columna & "} = '" & Cad & "'"
                        End If
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            End If


        'CHECK
                'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    
                    Aux = ""
                    If Control.Value = 1 Then
                        Aux = "1"
                    Else
                        If CHECK <> "" Then
                            CheckBusqueda Control
                            tabla = NombreCheck & "|"
                            If InStr(1, CHECK, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                        End If
                    End If
                    If Aux <> "" Then
                        If Not paraRPT Then
                            Cad = mTag.tabla & "." & mTag.columna
                        Else
                            Cad = "{" & mTag.tabla & "." & mTag.columna & "} "
                        End If
                        
                        Cad = Cad & " = " & Aux
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If 'cargado
                End If '<>""
            End If
        End If
    
    Next Control
    ObtenerBusqueda = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function



Private Function ObtenerMaximoMinimo(ByRef vSQL As String) As String
Dim Rs As Recordset
    ObtenerMaximoMinimo = ""
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.EOF) Then
            ObtenerMaximoMinimo = CStr(Rs.Fields(0))
        End If
    End If
    Rs.Close
    Set Rs = Nothing
End Function



Public Function QuitarCaracterEnter(vcad As String) As String
Dim I As Integer

    Do
        I = InStr(1, vcad, Chr(13))
        If I > 0 Then 'Hay ENTER
            vcad = Mid(vcad, 1, I - 1) & Mid(vcad, I + 2)
        End If
    Loop Until I = 0
    QuitarCaracterEnter = vcad
End Function



Public Function SeparaCampoBusqueda(Tipo As String, campo As String, CADENA As String, ByRef DevSQL As String, Optional paraRPT) As Byte
Dim Cad As String
Dim Aux As String
Dim CH As String
Dim fin As Boolean
Dim I, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
Cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    '==== Laura: 11/07/05
    If IsNumeric(CADENA) Then
        CADENA = CStr(ImporteFormateado(CADENA))
        CADENA = TransformaComasPuntos(CADENA)
    End If
    '====================
    I = CararacteresCorrectos(CADENA, "N")
    If I > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo numerico
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not IsNumeric(Cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = campo & " >= " & Cad & " AND " & campo & " <= " & Aux
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                DevSQL = "1=1"
             Else
                    fin = False
                    I = 1
                    Cad = ""
                    Aux = "NO ES NUMERO"
                    While Not fin
                        CH = Mid(CADENA, I, 1)
                        If CH = ">" Or CH = "<" Or CH = "=" Then
                            Cad = Cad & CH
                            Else
                                Aux = Mid(CADENA, I)
                                fin = True
                        End If
                        I = I + 1
                        If I > Len(CADENA) Then fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If Cad = "" Then Cad = " = "
                    DevSQL = campo & " " & Cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(CADENA, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo de fechas
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not EsFechaOK(Cad) Or Not EsFechaOK(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        
'        If Not Left(campo, 1) = "{" Then
'                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
'                Else
'                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
'                End If
        
        If paraRPT Then
            Cad = "Date(" & Year(Cad) & "," & Month(Cad) & "," & Day(Cad) & ")"
            Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
            DevSQL = campo & " >=" & Cad & " AND " & campo & " <= " & Aux
        Else
            Cad = Format(Cad, FormatoFecha)
            Aux = Format(Aux, FormatoFecha)
            'En my sql es la ' no el #
            'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
            DevSQL = campo & " >='" & Cad & "' AND " & campo & " <= '" & Aux & "'"
        End If
        '----
        'ELSE
    Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                fin = False
                I = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not fin
                    CH = Mid(CADENA, I, 1)
                    If CH = ">" Or CH = "<" Or CH = "=" Then
                        Cad = Cad & CH
                        Else
                            Aux = Mid(CADENA, I)
                            fin = True
                    End If
                    I = I + 1
                    If I > Len(CADENA) Then fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOK(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                If Not Left(campo, 1) = "{" Then
                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
                Else
                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
                End If
                If Cad = "" Then Cad = " = "
                DevSQL = campo & " " & Cad & " " & Aux
            End If
    End If
    
  Case "H"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(CADENA, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo de fechas
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not EsFechaOK(Cad) Or Not EsFechaOK(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        
'        If Not Left(campo, 1) = "{" Then
'                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
'                Else
'                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
'                End If
        
        If paraRPT Then
            Cad = "Date(" & Year(Cad) & "," & Month(Cad) & "," & Day(Cad) & ")"
            Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
            DevSQL = campo & " >=" & Cad & " AND " & campo & " <= " & Aux
        Else
            Cad = Format(Cad, FormatoFecha)
            Aux = Format(Aux, FormatoFecha)
            'En my sql es la ' no el #
            'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
            DevSQL = campo & " >='" & Cad & "' AND " & campo & " <= '" & Aux & "'"
        End If
        '----
        'ELSE
    Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                fin = False
                I = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not fin
                    CH = Mid(CADENA, I, 1)
                    If CH = ">" Or CH = "<" Or CH = "=" Then
                        Cad = Cad & CH
                        Else
                            Aux = Mid(CADENA, I)
                            fin = True
                    End If
                    I = I + 1
                    If I > Len(CADENA) Then fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOK(Aux) Then
                    'Veo si es una hora
                    If EsHoraOK(Aux) Then MsgBox "Debe especificar fecha/hora", vbExclamation
                    
                    Exit Function
                    
                    
                Else
                
                
                
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If Not Left(campo, 1) = "{" Then
                        Aux = Format(Aux, FormatoFecha)
                        DevSQL = campo & " >= '" & Aux & " 00:00:00' AND " & campo & " <= '" & Aux & " 23:59:59'"
                    Else
                        Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
                    
                    
                        DevSQL = campo & " " & Cad & " " & Aux
                    End If
                End If
            End If
    End If
    
Case "T"
    '---------------- TEXTO ------------------
    I = CararacteresCorrectos(CADENA, "T")
    If I = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If CADENA = ">>" Or CADENA = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    
    'Comprobamos si es LIKE o NOT LIKE
    Cad = Mid(CADENA, 1, 2)
    If Cad = "<>" Then
        CADENA = Mid(CADENA, 3)
        If Left(campo, 1) <> "{" Then
            'No es consulta seleccion para Report.
            DevSQL = campo & " NOT LIKE '"
        Else
            'Consulta de seleccion para Crystal Report
            DevSQL = "NOT (" & campo & " LIKE """ & CADENA & """)"
        End If
    Else
        If Left(campo, 1) <> "{" Then
        'NO es para report
            DevSQL = campo & " LIKE '"
        Else  'Es para report
            I = InStr(1, CADENA, "*")
            'Poner Consulta de seleccion para Crystal Report
            If I > 0 Then
                DevSQL = campo & " LIKE """ & CADENA & """"
            Else
                DevSQL = campo & " = """ & CADENA & """"
            End If
        End If
    End If
    
    
    'Cambiamos el * por % puesto que en ADO es el caraacter para like
    I = 1
    Aux = CADENA
    If Not Left(campo, 1) = "{" Then
      'No es para report
       While I <> 0
           I = InStr(1, Aux, "*")
           If I > 0 Then
                Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
            End If
        Wend
    End If
    
    'Cambiamos el ? por la _ pue es su omonimo
    I = 1
    While I <> 0
        I = InStr(1, Aux, "?")
        If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
    Wend
    
    
    'Poner el valor de la expresion
    If Left(campo, 1) <> "{" Then
        'No es consulta seleccion para Report.
        DevSQL = DevSQL & Aux & "'"
    'Else
        'Consulta de seleccion para Crystal Report
        'DevSQL = DevSQL & CADENA & """)"
    End If
    
    '=========
    'ANTES
'    If cad = "<>" Then
'        '====David
'        'Aux = Mid(CADENA, 3)
'        'LAura
'        Aux = Mid(Aux, 3)
'        '====
'        If Left(Campo, 1) <> "{" Then
'            'Mo es consulta seleccion para Report.
'            DevSQL = Campo & " NOT LIKE '" & Aux & "'"
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " <> " & Aux & ""
'        End If
'    Else
'        If Left(Campo, 1) <> "{" Then
'            DevSQL = Campo & " LIKE '" & Aux & "'"
'        ElseIf Left(Aux, 4) = "like" Then
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " " & Aux
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " = """ & Aux & """"
'        End If
'    End If
    
    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    I = InStr(1, CADENA, "<>")
    If I = 0 Then
        'IGUAL A valor
        Cad = " = "
        Else
            'Distinto a valor
        Cad = " <> "
    End If
    'Verdadero o falso
    I = InStr(1, CADENA, "V")
    If I > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = campo & " " & Cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Private Function CararacteresCorrectos(vcad As String, Tipo As String) As Byte
Dim I As Integer
Dim CH As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " ", "-"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "a" To "z"
            Case "A" To "Z"
            Case "0" To "9"
            Case "*", "%", "?", "_", "\", "/", ":", ".", " " ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "<", ">"
            Case "Ñ", "ñ"
            Case "-", "+", ",", """" 'Añade Laura
            'Abril 2014
            Case "[", "]"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
    
Case "F"
    'Tipo Fecha. Aceptamos Numeros , "/" ,":"
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next I

Case "B"
    'Numeros , "/" ,":"
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next I
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function






Public Function QuitarCaracterNULL(vcad As String) As String
Dim I As Integer

    Do
        I = InStr(1, vcad, vbNullChar)
        If I > 0 Then 'Hay null
            vcad = Mid(vcad, 1, I - 1) & Mid(vcad, I + 2)
        End If
    Loop Until I = 0
    QuitarCaracterNULL = vcad
End Function


Public Sub CheckBusqueda(ByRef CH As CheckBox)
    NombreCheck = ""
    NombreCheck = CH.Name & "("
    On Error Resume Next
    NombreCheck = NombreCheck & CH.Index
    If Err.Number <> 0 Then Err.Clear
    NombreCheck = NombreCheck & ")"
End Sub


Public Function EsHoraOK(T As String) As Boolean
Dim Cad As String
    
    Cad = T
    If InStr(1, Cad, ":") = 0 Then
        Select Case Len(T)
            Case 8
                Cad = Mid(Cad, 1, 2) & ":" & Mid(Cad, 3, 2) & ":" & Mid(Cad, 5)
            Case 6
                Cad = Mid(Cad, 1, 2) & ":" & Mid(Cad, 3, 2) & ":" & Mid(Cad, 5)
            Case 4
                Cad = Mid(Cad, 1, 2) & ":" & Mid(Cad, 3, 2) & ":00"
        End Select
    End If
    
    If IsDate(Cad) Then
        EsHoraOK = True
        T = Format(Cad, "hh:mm:ss")
    Else
        EsHoraOK = False
    End If
End Function




Public Function InsertarDesdeForm(ByRef formulario As Form, Optional opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        
        ElseIf TypeOf Control Is TextBox And Control.Visible = True Then
            If (opcion = 1 And Control.Name = "Text1") Or opcion = 0 Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.columna & ""
                        
                            'Parte VALUES
                            Cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & Cad
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                    Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    conn.Execute Cad, , adCmdText
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function






Public Function ValorParaSQL(Valor, ByRef vtag As CTag) As String
Dim Dev As String
Dim D As Single
Dim I As Integer
Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vtag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004

                    'Ademas de la coma lleva puntos
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                Else

                    V = CSng(Valor)
                    Valor = V
                End If
            Else

            End If
            Dev = TransformaComasPuntos(CStr(Valor))

        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
        Case "H"
            Dev = "'" & Format(Valor, FormatoFecha & " hh:mm:ss") & "'"
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select

    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vtag.Vacio = "S" Then
            Dev = ValorNulo
        Else
            'Modifica Laura: 04/10/05
            If vtag.TipoDato = "N" Then
                Dev = "0"
            Else
                Dev = "''"
            End If
        End If
    End If
    ValorParaSQL = Dev
End Function






Public Function ModificaDesdeFormulario(ByRef formulario As Form, opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        ElseIf TypeOf Control Is TextBox And Control.Visible = True Then
            If (opcion = 1 And Control.Name = "Text1") Or (opcion = 3 And Control.Name = "txtAux") Then
            If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                                 cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox And Control.Visible Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If

        ElseIf TypeOf Control Is ComboBox And Control.Visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        ElseIf TypeOf Control Is OptionButton And Control.Visible Then
            If Control.Enabled Then
                If Control.Value = True And Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Aux = Control.Index
                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWhere
    conn.Execute Aux, , adCmdText

    ModificaDesdeFormulario = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function




Public Sub TerminaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=1"
End Sub



Public Function CheckValueLeer(NombreForm As String) As Byte
Dim NombreFichero As String

On Error GoTo ECheckValueLeer
CheckValueLeer = 0  'UNCHECKED
'Se podria hacer un select para que no lie mucho los nombres en las carpetas
NombreFichero = DevNombreFichero(NombreForm)
If NombreFichero <> "" Then
    If Dir(NombreFichero) <> "" Then CheckValueLeer = 1
End If


Exit Function
ECheckValueLeer:
    Err.Clear
End Function

Private Function DevNombreFichero(Nombre As String) As String

DevNombreFichero = App.Path & "\" & Nombre & ".xdf"
End Function

Public Function BLOQUEADesdeFormulario(ByRef formulario As Form, Optional opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWhere As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is CommonDialog Then
        
        ElseIf TypeOf Control Is TextBox And Control.Visible = True Then
            If (opcion = 1 And Control.Name = "Text1") Or opcion <> 1 Then
                If Control.Tag <> "" Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWhere <> "" Then cadWhere = cadWhere & " AND "
                             cadWhere = cadWhere & "(" & mTag.columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & cadWhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
    
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function



Public Sub PreparaBloquear()
    conn.Execute "commit"
    conn.Execute "set autocommit=0"
End Sub




Public Sub PonerFormatoFecha(ByRef T As TextBox)
Dim Cad As String

    Cad = T.Text
    If Cad <> "" Then
        If Not EsFechaOK(Cad) Then
            MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
            Cad = "mal"
        End If
        If Cad <> "" And Cad <> "mal" Then
            T.Text = Cad
        Else
            T.Text = ""
            PonerFoco T
        End If
    End If
End Sub


Public Sub PonerIndicador(ByRef lblIndicador As Label, Modo As Byte)
'Pone el titulo del label lblIndicador
    lblIndicador.FontBold = True
    Select Case Modo
        Case 0    'Modo Inicial
            lblIndicador.Caption = ""
        Case 1 'Modo Buscar
            lblIndicador.Caption = "BUSQUEDA"
        Case 2    'Preparamos para que pueda Modificar
        
        Case 3 'Modo Insertar
            lblIndicador.Caption = "INSERTAR"
        Case 4 'MODIFICAR
            lblIndicador.Caption = "MODIFICAR"
        Case Else
            lblIndicador.Caption = ""
    End Select
End Sub



Public Sub DesplazamientoVisible(ByRef toolb As Toolbar, iniBoton As Byte, bol As Boolean, nreg As Byte)
'Oculta o Muestra las botones de  flechas de desplazamiento de la toolbar
Dim I As Byte

    Select Case nreg
        Case 0, 1 '0 o 1 registro no mostrar los botones despl.
            For I = iniBoton To iniBoton + 3
                toolb.Buttons(I).Visible = False
            Next I
        Case Else '>1 reg, mostrar si bol
            For I = iniBoton To iniBoton + 3
                toolb.Buttons(I).Visible = bol
            Next I
    End Select
End Sub





Public Sub BloquearText1(ByRef formulario As Form, Modo As Byte)
'Bloquea controles q se llamen TEXT1 si no estamos en Modo: 3.-Insertar, 4.-Modificar
'si estamos en modo modificar bloquea solo los campos que son clave primaria
'IN ->  formulario: formulario en el que se van a poner los controles textbox en modo visualización
'       Modo: modo del mantenimiento (Insertar, Modificar,Buscar...)
Dim I As Byte
Dim B As Boolean
Dim vtag As CTag
On Error Resume Next

    With formulario
        B = (Modo = 3 Or Modo = 4 Or Modo = 1) 'And ModoLineas = 1))
        
        For I = 0 To .text1.Count - 1 'En principio todos los TExt1 tiene TAG
            Set vtag = New CTag
            vtag.Cargar .text1(I)
            If vtag.Cargado Then
                If vtag.EsClave And (Modo = 2 Or Modo = 4 Or Modo = 5) Then
                    .text1(I).Locked = True
                    .text1(I).BackColor = &H80000018 'amarillo claro
                Else
                    .text1(I).Locked = Not B  '((Not b) And (Modo <> 1))
                    If B Then
                        .text1(I).BackColor = vbWhite
                    Else
                        .text1(I).BackColor = &H80000018 'amarillo claro
                    End If
                    If Modo = 3 Then .text1(I).Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
                End If
            Else
                .text1(I).Locked = Not B  '((Not b) And (Modo <> 1))
                If B Then
                    .text1(I).BackColor = vbWhite
                Else
                    .text1(I).BackColor = &H80000018 'amarillo claro
                End If
            End If
        Set vtag = Nothing
        Next I
        
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub PonerLongCamposGnral(ByRef formulario As Form, Modo As Byte, opcion As Byte)
    Dim I As Integer
    
    On Error Resume Next

    With formulario
        If Modo = 1 Then 'BUSQUEDA
            Select Case opcion
                Case 1 'Para los TEXT1
                    For I = 0 To .text1.Count - 1
                        With .text1(I)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = (.HelpContextID * 2) + 1 'el doble + 1
                            End If
                        End With
                    Next I
                
                Case 3 'para los TXTAUX
                    For I = 0 To .txtAux.Count - 1
                        With .txtAux(I)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = (.HelpContextID * 2) + 1 'el doble + 1
                            End If
                        End With
                    Next I
            End Select
            
        Else 'resto de modos
            Select Case opcion
                Case 1
                    For I = 0 To .text1.Count - 1
                        With .text1(I)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next I
                Case 3
                    For I = 0 To .txtAux.Count - 1
                        With .txtAux(I)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next I
            End Select
        End If
    End With
    If Err.Number <> 0 Then Err.Clear
End Sub
 


Public Sub DesplazamientoData(ByRef vData As Adodc, Index As Integer)
'Para desplazarse por los registros de control Data
    If vData.Recordset.EOF Then Exit Sub
    Select Case Index
        Case 0 'Primer Registro
            If Not vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 1 'Anterior
            vData.Recordset.MovePrevious
            If vData.Recordset.BOF Then vData.Recordset.MoveFirst
        Case 2 'Siguiente
            vData.Recordset.MoveNext
            If vData.Recordset.EOF Then vData.Recordset.MoveLast
        Case 3 'Ultimo
            vData.Recordset.MoveLast
    End Select
End Sub


Public Function CompForm(ByRef formulario As Form, opcion As Byte) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is CommonDialog Then
        ElseIf False Then  'TypeOf Control Is MSComm
        ElseIf TypeOf Control Is TextBox And Control.Visible = True Then
            If (opcion = 1 And Control.Name = "Text1") Or (opcion = 2 And Control.Name = "Text3") Or (opcion = 3 And Control.Name = "txtAux") Then
                Carga = mTag.Cargar(Control)
                If Carga = True Then
                    Correcto = mTag.Comprobar(Control)
                    If Not Correcto Then Exit Function
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox And Control.Visible = True Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function

                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function




Public Function ParaGrid(ByRef Control As Control, AnchoPorcentaje As Integer, Optional Desc As String) As String
Dim mTag As CTag
Dim Cad As String
'====Modificado por Laura Junio 2004:
'====Se añade el formato empipado
'Montamos al final: "Cod Diag.|tabla|columna|tipo|formato|10·"

    ParaGrid = ""
    Cad = ""
    Set mTag = New CTag
    mTag.Cargar Control
    If mTag.Cargado Then
        If Control.Tag <> "" Then
            'Si es texto monta esta parte de sql
            If TypeOf Control Is TextBox Then
                If Desc <> "" Then
                    Cad = Desc
                Else
                    Cad = mTag.Nombre
                End If
                Cad = Cad & "|"
                
                '----------------------
                'Añade Laura - 1/9/04
                Cad = Cad & mTag.tabla & "|"
                '----------------------
                
                Cad = Cad & mTag.columna & "|"
                Cad = Cad & mTag.TipoDato & "|"
                
                '----------------------
                'Añade Laura - Junio/04
                Cad = Cad & mTag.Formato & "|"
                '----------------------
                
                Cad = Cad & AnchoPorcentaje & "·"
    
            'CheckBOX
            ElseIf TypeOf Control Is CheckBox Then
    
            ElseIf TypeOf Control Is ComboBox Then
                If Desc <> "" Then
                    Cad = Desc
                Else
                    Cad = mTag.Nombre
                End If
                Cad = Cad & "|"
                '----------------------
                'Añade Laura - 1/9/04
                Cad = Cad & mTag.tabla & "|"
                '----------------------
                Cad = Cad & mTag.columna & "|"
                Cad = Cad & mTag.TipoDato & "|"
                Cad = Cad & mTag.Formato & "|"
                Cad = Cad & AnchoPorcentaje & "·"
            
    
            End If 'De los elseif
        End If
        Set mTag = Nothing
        ParaGrid = Cad
    End If
End Function



Public Function ValorDevueltoFormGrid(ByRef Control As Control, ByRef CadenaDevuelta As String, Orden As Integer) As String
Dim mTag As CTag
Dim Cad As String
Dim Aux As String
'Montamos al final: " columnatabla = valordevuelto "

    ValorDevueltoFormGrid = ""
    Cad = ""
    Set mTag = New CTag
    mTag.Cargar Control
    If mTag.Cargado Then
        If Control.Tag <> "" Then
            'Si es texto monta esta parte de sql
            If TypeOf Control Is TextBox Then
                Aux = RecuperaValor(CadenaDevuelta, Orden)
                If Aux <> "" Then Cad = mTag.columna & " = " & ValorParaSQL(Aux, mTag)
            'CheckBOX
           ' ElseIf TypeOf Control Is CheckBox Then
           '
           ' ElseIf TypeOf Control Is ComboBox Then
           '
           '
            End If 'De los elseif
        End If
    End If
    Set mTag = Nothing
    ValorDevueltoFormGrid = Cad
End Function



Public Function SituarDataTrasEliminar(ByRef vData As Adodc, NumReg, Optional NoActualiza As Boolean) As Boolean
'NumReg: numero de registro que acabo de eliminar
'NoActualiza: si se hace el refresh o no, por defecto siempre se hace el refresh
'             pero si hemos eliminado de un Grid ya se hizo en el cargaGrid y
'             no lo volvemos a hacer para mantener las columnas bien.

    On Error GoTo ESituarDataElim
    
        If NoActualiza = False Then vData.Refresh
        
        If Not vData.Recordset.EOF Then    'Solo habia un registro
            If NumReg > vData.Recordset.RecordCount Then
                vData.Recordset.MoveLast
            Else
                vData.Recordset.MoveFirst
                vData.Recordset.Move NumReg - 1
            End If
            SituarDataTrasEliminar = True
        Else
            SituarDataTrasEliminar = False
        End If
        
ESituarDataElim:
    If Err.Number <> 0 Then
        Err.Clear
        SituarDataTrasEliminar = False
    End If
End Function



Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Cad As String
Dim Valor As Variant
Dim campo As String  'Campo en la base de datos
Dim I As Integer


    On Error GoTo EPonerCamposForma


    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'TEXTO
        If TypeOf Control Is CommonDialog Then
        
        ElseIf (TypeOf Control Is TextBox) And (Control.Visible = True) And (Control.Name = "Text1") Then
'                If TypeOf control Is TextBox Then

            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    
                    If mTag.columna <> "" Then
                        'Debug.Print mTag.columna
                        'If mTag.columna = "porciva3re" Then Stop
                        
                        campo = mTag.columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(campo))
                        Else
                            Valor = vData.Recordset.Fields(campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            If mTag.TipoDato = "N" Then
                                If Val(Valor) = 0 Then
                                    Control.Text = ""
                                Else
                                    Control.Text = Valor
                                End If
                            Else
                                Control.Text = Valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) And (Control.Visible = True) Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    campo = mTag.columna
                    Valor = vData.Recordset.Fields(campo)
                    Else
                        Valor = 0
                End If
                If IsNull(Valor) Then Valor = 0
                Control.Value = Valor
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) And Control.Visible Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    campo = mTag.columna
                    Valor = DBLet(vData.Recordset.Fields(campo))
                    I = 0
                    For I = 0 To Control.ListCount - 1
                        If Control.ItemData(I) = Val(Valor) Then
                            Control.ListIndex = I
                            Exit For
                        End If
                    Next I
                    If I = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    Cad = Err.Description
    Cad = "Poner campos formulario. " & vbCrLf & campo & vbCrLf & Cad & vbCrLf
    MsgBox Cad, vbExclamation
End Function




Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim I As Integer
Dim J As Integer

On Error GoTo EPonerOpcionesMenuGeneral

'Añadir, modificar y borrar deshabilitados si no nivel
With formulario

    'LA TOOLBAR  .--> Requisito, k se llame toolbar1
    For I = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(I).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(I).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(I).Enabled = False
            End If
        End If
    Next I

    'Esto es un poco salvaje. Por si acaso , no existe en este trozo pondremos los errores on resume next

    On Error Resume Next

    'Los MENUS
    'K sean mnAlgo
    J = Val(.mnNuevo.HelpContextID)
    If J < vUsu.Nivel Then .mnNuevo.Enabled = False

    J = Val(.mnModificar.HelpContextID)
    If J < vUsu.Nivel Then .mnModificar.Enabled = False

    J = Val(.mnEliminar.HelpContextID)
    If J < vUsu.Nivel Then .mnEliminar.Enabled = False
    
    J = Val(.mnLineas.HelpContextID)
    If J < vUsu.Nivel Then .mnLineas.Enabled = False
    
    On Error GoTo 0
End With

Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub




Public Function SituarDataMULTI(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String) As Boolean
'Situa un DataControl en el registo que cumple vwhere
On Error GoTo ESituarData
        'Actualizamos el recordset
        vData.Refresh
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find vData.Recordset, vWhere
        'vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarDataMULTI = False
End Function


Public Sub Multi_Find(ByRef oRs As ADODB.Recordset, sCriteria As String)
'para el situarDataMULTI
On Error Resume Next
    Dim clone_rs As ADODB.Recordset
    Set clone_rs = oRs.Clone
    
    clone_rs.Filter = sCriteria
    
    If clone_rs.EOF Or clone_rs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    clone_rs.Close
    Set clone_rs = Nothing
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function DBLetMemo(vData As Variant) As String
    On Error Resume Next
    
    DBLetMemo = vData
    
'    If IsNull(DBLetMemo) Then DBLetMemo = ""
    
    If Err.Number <> 0 Then
        Err.Clear
        DBLetMemo = ""
    End If
End Function




'Funciona para claves primarias formadas por 2 campos
Public Function DevuelveDesdeBDNew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef otroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim Rs As Recordset
Dim Cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    Cad = "Select " & kCampo
    If otroCampo <> "" Then Cad = Cad & ", " & otroCampo
    Cad = Cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        Cad = Cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            Cad = Cad & Val(valorCodigo1)
        Case "T"
            Cad = Cad & DBSet(valorCodigo1, "T")
        Case "F"
            Cad = Cad & "'" & valorCodigo1 & "'"
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        Cad = Cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            Cad = Cad & DBSet(ValorCodigo2, "T")
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        Cad = Cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                Cad = Cad & "-1"
            Else
                Cad = Cad & Val(ValorCodigo3)
            End If
        Case "T"
            Cad = Cad & "'" & ValorCodigo3 & "'"
        Case "F"
            Cad = Cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    If vBD = conAri Then 'BD 1: Ariges
        Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        Rs.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
    If Not Rs.EOF Then
        DevuelveDesdeBDNew = DBLet(Rs.Fields(0))
        If otroCampo <> "" Then otroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function






Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function






Public Function Comprueba_CuentaBan2(CC As String, Optional OcultarMsgbox As Boolean) As Boolean
Dim Cad As String

    'Validar que la cuenta bancaria es correcta
    Comprueba_CuentaBan2 = False
    If Trim(CC) <> "" Then
        If Not Comprueba_CC(CC, Cad) Then
            If Not OcultarMsgbox Then MsgBox "La cuenta bancaria no es correcta." & vbCrLf & Cad, vbInformation
        Else
            Comprueba_CuentaBan2 = True
        End If
    End If
End Function
'------------------------------


Private Function Comprueba_CC(CC As String, ByRef MensajeError As String) As Boolean
    Dim Calculado As String
    Dim I, i2, i3, i4 As Integer

    
    MensajeError = "Longitud <>20"
    
    '-- Esta función comprueba la corrección de un número de cuenta pasado en CC
    If Len(CC) <> 20 Then Exit Function '-- Las cuentas deben contener 20 dígitos en total
    MensajeError = ""
    
    '-- Calculamos el primer dígito de control
    I = Val(Mid(CC, 1, 1)) * 4
    I = I + Val(Mid(CC, 2, 1)) * 8
    I = I + Val(Mid(CC, 3, 1)) * 5
    I = I + Val(Mid(CC, 4, 1)) * 10
    I = I + Val(Mid(CC, 5, 1)) * 9
    I = I + Val(Mid(CC, 6, 1)) * 7
    I = I + Val(Mid(CC, 7, 1)) * 3
    I = I + Val(Mid(CC, 8, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    
    Calculado = i4
    If i4 <> Val(Mid(CC, 9, 1)) Then MensajeError = "N"
    
    '-- Calculamos el segundo dígito de control
    I = Val(Mid(CC, 11, 1)) * 1
    I = I + Val(Mid(CC, 12, 1)) * 2
    I = I + Val(Mid(CC, 13, 1)) * 4
    I = I + Val(Mid(CC, 14, 1)) * 8
    I = I + Val(Mid(CC, 15, 1)) * 5
    I = I + Val(Mid(CC, 16, 1)) * 10
    I = I + Val(Mid(CC, 17, 1)) * 9
    I = I + Val(Mid(CC, 18, 1)) * 7
    I = I + Val(Mid(CC, 19, 1)) * 3
    I = I + Val(Mid(CC, 20, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    
    
    Calculado = Calculado & i4
    
    If i4 <> Val(Mid(CC, 10, 1)) Then MensajeError = "N"
    
    If MensajeError <> "" Then
        MensajeError = "CC calculado: " & Calculado & "  -  " & Mid(CC, 9, 2)
    Else
        Comprueba_CC = True
    End If

End Function



Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim Ent As Integer
Dim Cad As String
  
  ' Comprobaciones
  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un número."
    Exit Function
  End If
  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If
  
  ' Redondeo.
  Cad = "0"
  If NumDigitsAfterDecimals <> 0 Then Cad = Cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Format(Number, Cad)
  
End Function


Public Function InstalacionEsEulerTaxco() As Boolean
    InstalacionEsEulerTaxco = True
End Function




'Devuelve en otros campos, 2 valores
Public Function DevuelveDesdeBD2(vBD As Byte, kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Tipo As String, ByRef otroCampo1 As String, ByRef otroCampo2 As String) As String
    Dim Rs As Recordset
    Dim Cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD2 = ""
    Cad = "Select " & kCampo
    Cad = Cad & ", " & otroCampo1 & ", " & otroCampo2
    Cad = Cad & " FROM " & Ktabla
    Cad = Cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        Cad = Cad & ValorCodigo
    Case "T", "F", "T1"
        Cad = Cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
'    Debug.Print cad
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    If vBD = 1 Then 'BD 1: Ariges
        Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Else    'BD 2: Conta
        Rs.Open Cad, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    End If
    
    If Not Rs.EOF Then
        DevuelveDesdeBD2 = DBLet(Rs.Fields(0))
        otroCampo1 = DBLet(Rs.Fields(1))
        otroCampo2 = DBLet(Rs.Fields(2))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function

