Attribute VB_Name = "libSrvImpresion"
Option Explicit

Dim Error As String
Dim Destino As String
Dim PathDestino As String
Dim codClien As Long





Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomRPT As String
Dim vImprimedirecto As Boolean
Dim cadPDFrpt As String
Dim pRptvMultiInforme As Integer
Dim devuelve As String
Dim ImprimeDirecto As Boolean
Dim NumCopias As Integer

Dim Cambia_ODBC As Boolean



'Funciones que estan en otras mods de ariges, operon necesitamos

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
    'CADENA = "6"
    'EnDesarrolloServicioImpresion = True
    If CADENA <> "" Then
    
    
    
        Set vConfig = New Configuracion
        If vConfig.Leer = 1 Then
            Error = "error leyendo Config.cfg: "
        Else
            Cambia_ODBC = False
            If AbrirConexionServicio("usuarios") Then
            
                
            
                Tipo = ""
                Set RN = New ADODB.Recordset
                RN.Open "Select * from info_intercambio where  infoIntercambioId= " & CADENA, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RN.EOF Then
                    Error = "No existe registro en la tabla info_intercambio (" & CADENA & ")"
                Else
                    Select Case RN!Tipo
                    Case "OFE", "PED", "ALB", "FAC"
                        Tipo = RN!Tipo
                        DatoImpresion = RN!clave
                        BD = RN!sistema
                        
                    Case Else
                        Error = "Tipo impresion no desarrollada. ID: " & CADENA & " TIPO: " & RN!Tipo
                    End Select
                End If
                RN.Close
            
                PathDestino = ""
                RN.Open "Select PathDestino,quearigescambiaODBC from info_parametros where  true ORDER BY infoparametrosId ", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RN.EOF Then
                    PathDestino = DBLet(RN!PathDestino, "T")
                    If Not IsNull(RN!quearigescambiaODBC) Then Cambia_ODBC = (LCase(BD) = LCase(RN!quearigescambiaODBC))
                End If
                RN.Close
                
                If PathDestino = "" Then
                    Error = "Falta configurar infoparametros"
                    Tipo = ""
                End If
        
                If Tipo <> "" Then
                    
                    
                    
                    
                    If AbrirConexionServicio(BD) Then
                        
                        FormatoFecha = "yyyy-mm-dd"
                        
                        
                        Set vUsu = New usuario
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




Private Sub ImprimeOferta(QueOferta As String)
     
    'comprobamos que existe
    devuelve = DevuelveDesdeBD(conAri, "numofert", "scapre", "numofert", QueOferta, "N")
    If devuelve = "" Then
        Error = "No existe OFERTA " & QueOferta
        Exit Sub
    End If
    
    indRPT = 5
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomRPT, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then
        Error = "Leyendo scryst 5"
        Exit Sub
    End If
         
    
    cadParam = "|pCodigoISO=""""|pCodigoRev=""""|pCodigoISO=""""|pCodigoRev=""""|pCodCarta=0|pCodUsu=0|"
    devuelve = DevuelveDesdeBDNew(conAri, "scapre", "codclien", "numofert", QueOferta, "N")
    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", devuelve, "N")
    If devuelve = "" Then devuelve = "0"
    cadParam = cadParam & "pTipoIVA=" & devuelve
    devuelve = DevuelveDesdeBD(conAri, "artSeparador", "spara1", "1", "1")
    cadParam = cadParam & "|Separador=""" & devuelve & """|"
     
     
    cadFormula = "{scapre.numofert} IN [" & QueOferta & "] "
     
    LanzaImprimir 31
    
    If Dir(App.Path & "\docum.pdf") = "" Then
        Error = "Error exportando PDF"
    Else
        Destino = "OFE" & Format(Val(QueOferta), "000000") & ".pdf"
    End If
End Sub


Private Function CopiarFichero(Origen As String, ElDestino As String) As Boolean
    On Error Resume Next
    FileCopy Origen, ElDestino
    If Err.Number <> 0 Then
        Error = Err.Description
        Err.Clear
    End If
End Function


'QueFactura:  CODTIPOM_numfactu_fecfactu
Private Sub ImprimeFactura(QueFactura As String)
Dim EsFraTelefono As Boolean
Dim TipoM As String
Dim Numfac As Long
Dim Fefac As Date
Dim J As Integer
Dim Aux As String
Dim NumeroTerminal As Integer
        NumeroTerminal = 0
        EsFraTelefono = False
        'FAS_1_2019-10-02
        Aux = ""
        J = InStr(1, QueFactura, "_")
        If J = 0 Then
            Aux = "separador 1"
        Else
            TipoM = Mid(QueFactura, 1, J - 1)
            QueFactura = Mid(QueFactura, J + 1)
            J = InStr(1, QueFactura, "_")
            If J = 0 Then
                 Aux = "separador 2"
             Else
                cadSelect = Mid(QueFactura, J + 1)
                If Len(cadSelect) = 10 Then cadSelect = Mid(cadSelect, 9, 2) & "/" & Mid(cadSelect, 6, 2) & "/" & Mid(cadSelect, 1, 4)
                QueFactura = Mid(QueFactura, 1, J - 1)
             End If
       End If
       If Aux <> "" Then
            Error = "localizando factura: " & Aux
            Exit Sub
       End If
        
       If Not EsFechaOK(cadSelect) Then Aux = "No es fecha correcta: " & cadSelect
       If Not IsNumeric(QueFactura) Then Aux = "No es campo numerico: " & QueFactura
        If Aux <> "" Then
            Error = "Campos factura: " & Aux
            Exit Sub
       End If
       Fefac = CDate(cadSelect)
       Numfac = Val(QueFactura)
       'VEmpos si existe la factura
       cadSelect = ""
       Aux = "codtipom =" & DBSet(TipoM, "T") & " AND numfactu =" & Numfac & " AND fecfactu =" & DBSet(Fefac, "F") & " AND 1"
       Aux = DevuelveDesdeBD(conAri, "codclien", "scafac", Aux, "1")
       If Aux = "" Then
            Error = "No existe factura: " & TipoM & Numfac & Fefac
            Exit Sub
       End If
       codClien = Val(Aux)
        
        
       If EsFraTelefono Then
            'ImprimirFraTelefonia
        
        Else
            'If CInt(DBLet(Data3.Recordset!NumTermi, "N")) > 0 Then
            If NumeroTerminal > 0 Then
                'Es factura del TPV
                BotonImprimeFactura 63, TipoM, Numfac, Fefac
            Else
                
                'Impresion normal
                'Indice = 53  '53: Informe de Facturas
                BotonImprimeFactura (53), TipoM, Numfac, Fefac
            End If
        End If
        
        If Error = "" Then
            If Dir(App.Path & "\docum.pdf") = "" Then
                Error = "Error exportando PDF"
            Else
                Destino = TipoM & Format(Numfac, "000000") & "_" & Format(Fefac, "yyyymmdd") & ".pdf"
            End If
                
        End If
        
        
        
End Sub

Private Sub BotonImprimeFactura(OpcionListado As Integer, HcoMov As String, NumFactu As Long, Fefactu As Date)
Dim devuelve As String
Dim ImprimeDirecto As Boolean

    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0

    
    '===================================================
    '============ PARAMETROS ===========================
    If (OpcionListado = 53) Then
        If HcoMov = "FAZ" Then
            'Factura B
            indRPT = 30
        
        'EULER
        ElseIf HcoMov = "FAO" Then
            indRPT = 78 'Orden trabajo
        ElseIf HcoMov = "FAE" Then
            indRPT = 79 'trabajo exterior
        'TELEFONIA
        ElseIf HcoMov = "FAT" Then
            indRPT = 63 'Facturas telefonia
        Else
            indRPT = 12 'Facturas Clientes
            
        End If
        
        'En taxco
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            'Facturas alvic
            cadParam = "|" & HcoMov & "|"
            If InStr(1, "|FA1|FA2|FA3|FAB|FAD|", cadParam) > 0 Then indRPT = 93
            cadParam = ""
        End If
        
    Else
        If (OpcionListado = 89) Then
            indRPT = OpcionListado
    
        ElseIf (OpcionListado = 94) Then
            indRPT = OpcionListado
        Else
            'OpcionListado = 53
            '-----------------------------------------------
            indRPT = 18 'Facturas Clientes TPV
        End If
    End If
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomRPT, ImprimeDirecto, pPdfRpt, pRptvMultiInforme) Then Exit Sub
      
      
      
    'PUNTO VERDE
    '--------------------------------------------------------------------------
    If vParamAplic.ArtReciclado <> "" Then
        cadParam = cadParam & "PuntoVerde= """ & vParamAplic.ArtReciclado & """|"
        numParam = numParam + 1
    End If
      
    

  

        'Cod Tipo Movimiento
        devuelve = "{scafac.codtipom}='" & HcoMov & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        'Nº Factura
        devuelve = "{scafac.numfactu}=" & NumFactu
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
        
        'Fecha Factura
        devuelve = "{scafac.fecfactu}= Date(" & Year(Fefactu) & "," & Month(Fefactu) & "," & Day(Fefactu) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        
    
  
  
        LanzaImprimir OpcionListado
         
  
         
         If Dir(App.Path & "\docum.pdf") = "" Then
            Error = "Error exportando PDF"
        Else
            '' HcoMov As String, NumFactu As Long, Fefactu As Date
            Destino = HcoMov & Format(NumFactu, "000000") & "_" & Format(Fefactu, "yyyymmdd") & ".pdf"
        End If
  
     
     
     
End Sub



Private Sub ImprimePEdido(QuePedido As String)

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim devuelve As String, campo As String
Dim CodPed As String
Dim campo1 As String, campo2 As String, campo3 As String
    
    
    devuelve = DevuelveDesdeBD(conAri, "codclien", "scaped", "numpedcl", QuePedido, "N")
    If devuelve = "" Then
        Error = "No existe pedido " & QuePedido
        Exit Sub
    End If
    codClien = Val(devuelve)

   
    indRPT = 7 '7: Pedidos de Clientes
         
  
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomRPT, vImprimedirecto, cadPDFrpt, pRptvMultiInforme) Then Exit Sub
     
    
        campo1 = "numpedcl"
        campo2 = "fecpedcl"
        campo3 = "codclien"
  
  
        cadFormula = "{scaped.numpedcl} = " & QuePedido
  
  

        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", CStr(codClien), "N")
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
        
        'PORTES
        cadParam = cadParam & "vPortes=""" & vParamAplic.ArtPortesN & """|"
        numParam = numParam + 1
  

   
        cadParam = cadParam & "valorado= " & 1 & "|"
        numParam = numParam + 1

    
    LanzaImprimir 12
    
    
    If Dir(App.Path & "\docum.pdf") = "" Then
        Error = "Error exportando PDF"
    Else
        Destino = "PED" & Format(Val(QuePedido), "000000") & ".pdf"
    End If
    
    
    
    
End Sub


'opcion listad: 45 normal
Private Sub ImprimeAlbaran(OpcionListado As Byte, QueAlbaran As String)
Dim devuelve As String
Dim hcoCodTipoM As String
Dim EsHistorico As Boolean


    EsHistorico = False
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0

    hcoCodTipoM = Mid(QueAlbaran, 1, 3)
    QueAlbaran = Mid(QueAlbaran, 4)
    
    
    devuelve = "codtipom =" & DBSet(hcoCodTipoM, "T") & " AND numalbar =" & QueAlbaran & " AND 1"
    devuelve = DevuelveDesdeBD(conAri, "codclien", "scaalb", devuelve, "1", "N")
    If devuelve = "" Then
        Error = "No existe albaran " & hcoCodTipoM & QueAlbaran
        Exit Sub
    End If
    codClien = Val(devuelve)
    
    
    
    
    '===================================================
    '============ PARAMETROS ===========================
    'ALBARANES
    If hcoCodTipoM = "ALZ" Then
        indRPT = 29   'Albaranes B
    ElseIf hcoCodTipoM = "ALR" Then
        indRPT = 36
    ElseIf hcoCodTipoM = "ALS" Then
        indRPT = 39
    ElseIf hcoCodTipoM = "ALI" Then
        indRPT = 56
    Else
        If EsHistorico Then
            indRPT = 11 'Hist. Albaranes clientes
        Else
            indRPT = 10 'Albaran Clientes
        End If
    End If
    
    If Not PonerParamRPT2(indRPT, cadParam, numParam, nomRPT, False, pPdfRpt, pRptvMultiInforme) Then Exit Sub
   
    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
    'tabla temporal para el calculo del bruto total para cada tipo de IVA
    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
   
    'PORTES
    cadParam = cadParam & "vPortes=""" & vParamAplic.ArtPortesN & """|"
    numParam = numParam + 1
    
    'PUNTO VERDE
    cadParam = cadParam & "PuntoVerde=""" & vParamAplic.ArtReciclado & """|"
    numParam = numParam + 1
    
    'Si se imprimen importes y/o
    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", CStr(codClien), "N")
    If devuelve = "" Then devuelve = "0"
    ' 0 "Todo"
    ' 1 "Cantidad y Precio"
    ' 2 "Cantidad"
    cadParam = cadParam & "Albarcon=" & devuelve & "|"
    numParam = numParam + 1
    
    
'    'Nombre fichero .rpt a Imprimir
'    frmImprimir.SeleccionaRPTCodigo = 0 'pRptvMultiInforme
'    If Not ImpresionDirecta Then
'        frmImprimir.NombreRPT = nomDocu
'        frmImprimir.NombrePDF = pPdfRpt
'    End If
        
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de Albaran
   
        
       
       '     devuelve = "{scaalb.codtipom}=" & DBSet(hcoCodTipoM, "T")
       ' Else
            devuelve = "{scaalb.codtipom}='" & hcoCodTipoM & "'"  'lo que habia
       ' End If
        
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        'Nº Albaran
        devuelve = "{scaalb.numalbar}=" & Val(QueAlbaran)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
        
        If EsHistorico Then
'            'El campo fecha tambien es clave primaria
'            devuelve = Text1(1).Text
'            devuelve = "{" & NombreTabla & ".fechaalb}=Date(" & Year(devuelve) & "," & Month(devuelve) & "," & Day(devuelve) & ")"
'            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'
'            devuelve = "{" & NombreTabla & ".fechaalb}='" & Format(Text1(1).Text, FormatoFecha) & "'"
'            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        End If
        
    
   
    '=========================================================================
    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
    'que se aplica a ese cliente
    If hcoCodTipoM = "ALI" Then
        'facturas internas VAN sin IVA         Si los ALZ no
        cadParam = cadParam & "pTipoIVA=2|"
        numParam = numParam + 1
    Else
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", CStr(codClien), "N")
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
    End If
    
    
    'If ImpresionDirecta Then
    If False Then
        
        'Imrpimie directamente. Tipo 4tonda.  -----------
       ' If MsgBox("¿Imprimir el albarán?", vbQuestion + vbYesNo) = vbYes Then ImprimirDirectoAlb cadSelect
    Else
    
        'En visreport hay un sub para imprmir
        davidNumalbar = 0
        'If Not EsInformePortes Then
        If Not False Then
            davidCodtipom = hcoCodTipoM
            davidNumalbar = Val(QueAlbaran)
        End If
    
'        With frmImprimir
'            'Febrero 2010
'            'If indRPT = 34 Then
'                .outTipoDocumento = 0
'            'Else
'            '    .outTipoDocumento = 4
'            '    .outClaveNombreArchiv = Text1(30).Text & Text1(0).Text
'            '    .outCodigoCliProv = CLng(Text1(4).Text)
'            '    .NumeroCopias = vParamAplic.NumCop_AlbaranNormal
'            'End If
'
'            .FormulaSeleccion = cadFormula
'            .OtrosParametros = cadParam
'            .NumeroParametros = numParam
'            .SoloImprimir = False
'            .EnvioEMail = False
'            .opcion = OpcionListado
'            If indRPT = 34 Then
'                .Titulo = "Portes albaran "
'            Else
'                .Titulo = "Albaran de Cliente"
'            End If
'            .ConSubInforme = True
'            .Show vbModal
'
'
'
'
'            If Not EsHistorico Then
'                If Not EsInformePortes Then
'                    If HaPulsadoElBotonDeImprimir Then
'                        'UPDATEAMOS scaalb para que no reimpimrpima los albaranes
'                        'Cod Tipo Movimiento
'                        devuelve = "scaalb.codtipom = '" & CodTipoMov & "' AND scaalb.numalbar = " & Val(Text1(0).Text)
'                        devuelve = "UPDATE scaalb SET albImpreso = 1 WHERE " & devuelve
'                        Me.chkImpreso.Value = 1
'                        ejecutar devuelve, False
'                    End If
'                End If
'            End If
'        End With

         LanzaImprimir 45
         
         If Dir(App.Path & "\docum.pdf") = "" Then
            Error = "Error exportando PDF"
        Else
            Destino = hcoCodTipoM & Format(Val(QueAlbaran), "000000") & ".pdf"
        End If
         
         
    End If
End Sub
























'Respeta ARIGES
' OpcionListado:  31 oferta
Private Sub LanzaImprimir(OpcionListado As Integer)


 With frmImprimirServ
        
        
        If Dir(App.Path & "\docum.pdf") <> "" Then
            'App.LogEvent "Espera docum"
            Espera 0.5
            Kill App.Path & "\docum.pdf"
        End If
        .CambiaODBC = Cambia_ODBC
        .outTipoDocumento = 0
       ' If DatosEnvioMail <> "" Then
            .outTipoDocumento = 0 'RecuperaValor(DatosEnvioMail, 1)
            .outCodigoCliProv = 0 'RecuperaValor(DatosEnvioMail, 2)
            .outClaveNombreArchiv = "" 'RecuperaValor(DatosEnvioMail, 3)
       ' End If
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .opcion = OpcionListado
        .SoloImprimir = False
        .MostrarTreeDesdeFuera = False 'MostrarTreeEnPrint
        .EnvioEMail = True
        
        .Titulo = ""  'Titulo
        .NombreRPT = nomRPT
        .NumeroCopias = 1  'NumeroDeCopias
        .SeleccionaRPTCodigo = 0 'pRptvMultiInforme
        If cadPDFrpt <> "" Then .NombrePDF = cadPDFrpt
        .ConSubInforme = True 'conSubRPT
        .Show vbModal
    End With
End Sub



Public Sub ActualizarTablasMenusArigesNuevo()

End Sub
