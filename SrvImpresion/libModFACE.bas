Attribute VB_Name = "libModFACE"




'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef CADENA As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim cad As String

    I = 0
    cont = 1
    cad = ""
    Do
        J = I + 1
        I = InStr(J, CADENA, "|")
        If I > 0 Then
            If cont = Orden Then
                cad = Mid(CADENA, J, I - J)
                I = Len(CADENA) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = cad
End Function




Public Function BloqueoManual(cadTabla As String, cadWhere As String, Optional OcultarMsg As Boolean) As Boolean
Dim Aux As String

On Error GoTo EBLOQ
    BloqueoManual = False
    If cadWhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "INSERT INTO zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & cadTabla
        Aux = Aux & "',""" & cadWhere & """)"
        conn.Execute Aux
        BloqueoManual = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If conn.Errors.Count > 0 Then
            If conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            If Not OcultarMsg Then MsgBox "Registro bloqueado por otro usuario", vbExclamation
        End If
    End If
'    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueoManual(cadTabla As String) As Boolean
Dim SQL As String

'Solo me interesa la tabla
On Error Resume Next

        SQL = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.Codigo & " and tabla='" & cadTabla & "'"
        conn.Execute SQL
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function


'======== Añade: Laura
Public Function ContieneCaracterBusqueda(CADENA As String) As Boolean
'Comprueba si la cadena contiene algun caracter especial de busqueda
' >,>,>=,: , ....
'si encuentra algun caracter de busqueda devuelve TRUE y sale
Dim b As Boolean
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
    b = False
    Do
        CH = Mid(CADENA, I, 1)
        Select Case CH
            Case "<", ">", ":", "="
                b = True
            Case "*", "%", "?", "_", "\", ":" ', "."
                b = True
            Case Else
                b = False
        End Select
    'Next i
        I = I + 1
    Loop Until (b = True) Or (I > Len(CADENA))
    ContieneCaracterBusqueda = b
End Function




Public Function SugerirCodigoSiguienteStr(NomTabla As String, NomCodigo As String, Optional CondLineas As String) As String
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo ESugerirCodigo

    'SQL = "Select Max(codtipar) from stipar"
    SQL = "Select Max(" & NomCodigo & ") from " & NomTabla
    If CondLineas <> "" Then
        SQL = SQL & " WHERE " & CondLineas
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, , , adCmdText
    SQL = "1"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If IsNumeric(RS.Fields(0)) Then
                SQL = CStr(RS.Fields(0) + 1)
            Else
                If Asc(Left(RS.Fields(0), 1)) <> 122 Then 'Z
                SQL = Left(RS.Fields(0), 1) & CStr(Asc(Right(RS.Fields(0), 1)) + 1)
                End If
            End If
        End If
    End If
    RS.Close
    Set RS = Nothing
    SugerirCodigoSiguienteStr = SQL
ESugerirCodigo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



'-----------------------------------
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
        Case "T", "T1"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
            
        Case "FH"
        
            Dev = "'" & Format(Valor, FormatoFecha & " hh:mm:ss") & "'"
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




Public Sub PonerFoco(ByRef Text As TextBox)
On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
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






Public Sub KEYpressGnral(KeyAscii As Integer, Modo As Byte, Cerrar As Boolean)
'IN: codigo keyascii tecleado, y modo en que esta el formulario
'OUT: si se tiene que cerrar el formulario o no
    Cerrar = False
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Cerrar = True
    End If
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



Public Sub KEYdownLineas(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 37 'Desplazamiento Flecha Izquierda
            SendKeys "+{tab}"
        Case 38 'Desplazamieto Flecha Hacia Arriba
            SendKeys "+{tab}"
        Case 39 'Desplaz. Flecha Derecha
            SendKeys "{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub





Public Function PonerFormatoEntero(ByRef T As TextBox) As Boolean
'Comprueba que el valor del textbox es un entero y le pone el formato
Dim mTag As CTag
Dim cad As String
Dim Formato As String
On Error GoTo EPonerFormato

    If T.Text = "" Then Exit Function
    PonerFormatoEntero = True
    
    Set mTag = New CTag
    mTag.Cargar T
    If mTag.Cargado Then
       cad = mTag.Nombre 'descripcion del campo
       Formato = mTag.Formato
    End If
    Set mTag = Nothing

    If Not EsEnteroNew(T.Text) Then
        PonerFormatoEntero = False
        MsgBox "El campo " & cad & " tiene que ser un número entero.", vbExclamation
        PonerFoco T
    Else
         T.Text = Format(T.Text, Formato)
    End If
    
EPonerFormato:
    If Err.Number <> 0 Then Err.Clear
End Function



'*********** LAURA : 13/09/2005
Public Function EsEnteroNew(Texto As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEnteroNew = False

    If Not IsNumeric(Texto) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, Texto, ".")
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
                I = InStr(L, Texto, ",")
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



Private Function AbrirConexionServicio(BBDD As String) As Boolean
Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionServicio = False
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
    
    If Cambia_ODBC Then
        cad = "DRIVER={MySQL ODBC 3.51 Driver};;DESC=;DATA SOURCE=vAriges2;DATABASE=" & BBDD
    Else
        cad = "DRIVER={MySQL ODBC 3.51 Driver};;DESC=;DATA SOURCE=vAriges;DATABASE=" & BBDD
    End If
    cad = cad & ";;;Persist Security Info=true"
    
    conn.ConnectionString = cad
    conn.Open
    conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexionServicio = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión BD:Ariges.", Err.Description
End Function



Public Sub Main()
Dim CADENA As String
Dim Tipo As String
Dim RN As ADODB.Recordset
Dim DatoImpresion As String
Dim BD As String
Dim ElDestino As String



    'LLLEVARA EL ID de la tabla de intercambio  info_intercambio   infoIntercambioId
    CADENA = Command
    
    If CADENA = "" Then CADENA = "1"
    
    
    
    
    
    
        Set vConfig = New Configuracion
        If vConfig.Leer = 1 Then
            Error = "error leyendo Config.cfg: "
        Else
            Cambia_ODBC = False
            If AbrirConexionServicio("usuarios") Then
            
                    
                    
                    If AbrirConexionServicio(BD) Then
                        
                        FormatoFecha = "yyyy-mm-dd"
                        
                        
                        Set vUsu = New Usuario
                        vUsu.CadenaConexion = Replace(BD, "ariges", "")
                        vUsu.CadenaConexion = BD
                        
                        Set vEmpresa = New Cempresa
                        vEmpresa.LeerDatos
                        
                        Set vParamAplic = New CParamAplic
                        If vParamAplic.Leer(True) = 0 Then
                    
                            codClien = -1
                            
                            If Tipo = "OFE" Then
                                ImprimeOferta DatoImpresion
                            ElseIf Tipo = "PED" Then
                                ImprimePEdido DatoImpresion
                            ElseIf Tipo = "ALB" Then
                                ImprimeAlbaran 45, DatoImpresion
                            ElseIf Tipo = "FAC" Then
                                ImprimeFactura DatoImpresion
                            Else
                                Error = "Tpo ¿incorrecto? " & Tipo
                            End If
                        Else
                            Error = "Error abriendo conexion Ariges ODBC; " & BD
                        End If
                    Else
                        Error = "Error abriendo conexion Ariges ODBC; " & BD
                    End If
                End If
                
            Else
                Error = "Error abriendo conexion ODBC "
            End If
        End If  'de config
        
        
        'Si llega aqui, y la cadena error no esta vacia UPDATEAMOS a dos
        'Si es vacia, updateamos a uno
        If Destino <> "" Then
            'PathDestino = "Z:\aa bb"
            'ElDestino = """" & PathDestino & "\" & Destino & """"
            ElDestino = PathDestino & "\" & Destino
            If CopiarFichero(App.Path & "\docum.pdf", ElDestino) Then
                Error = ""
            Else
                
                
            End If
        End If
        
        If Cambia_ODBC Then
        
            Cambia_ODBC = False
            AbrirConexionServicio "usuarios"
        End If
        
        DatoImpresion = "UPDATE usuarios.info_intercambio SET estado = "
        If Error <> "" Then
           DatoImpresion = DatoImpresion & " 3"
           DatoImpresion = DatoImpresion & " , obs = " & DBSet(Error, "T")
        Else
            'todo OK
            DatoImpresion = DatoImpresion & " 1"
            DatoImpresion = DatoImpresion & " , obs = null "
            DatoImpresion = DatoImpresion & " , fichero = " & DBSet(Destino, "T")
            
        End If
        DatoImpresion = DatoImpresion & " WHERE infoIntercambioId = " & CADENA
        ejecutar DatoImpresion, True
        
        
        DatoImpresion = "UPDATE usuarios.info_parametros SET ExportacionFinalizada=0;"
        ejecutar DatoImpresion, True
        
    Else
        'cadena=""
        'Error = "Mal lanzado el programa"
    End If
    
    
    End
End Sub




