Attribute VB_Name = "LibImagen"
Option Explicit


''  -- Modos de Trabajo
'Public Const vbNorm = 0  ' modo normal
'Public Const vbHistNue = 1  ' modo de recuperar historico
'Public Const vbHistAnt = 2  ' modo de recuperar historico de los antiguos
'
'
'Public Const vbMaxGrupos = 31
'
'Public ModoTrabajo As Byte  '---------------------
'
'Public FormatoFecha As String
'
'Public Conn As Connection
'Public vUsu As Cusuarios
'Public vConfig As CConfiguracion

Public miRsAux As ADODB.Recordset


Public listacod As Collection
Public listaimpresion As Collection  'Esta lista servira para cuando queramos imprimir

'Cuiado con esta varibale
Public DatosModificados As Boolean


'Saber si ha coipado el archivo al server
Public DatosCopiados As String
'
'Public SeHaEjecutadoFTP As Boolean


Public Type RegistroTipoMensaje   ' Crea un tipo definido por el usuario.
   Descripcion As String
   Color As Long
   Icono As Integer
End Type


Public ArrayTipoMen() As RegistroTipoMensaje
Public TotalTipos As Integer   'Menos 1. Es decir, si hay tres tipos la var vale 2


''Usuario As String, Pass As String --> Directamente el usuario
'Public Function AbrirConexion() As Boolean
'Dim Cad As String
'On Error GoTo EAbrirConexion
'
'
'    AbrirConexion = False
'    Set Conn = Nothing
'    Set Conn = New Connection
'    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
'    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
'
'
'
'
'    'cadenaconexion
'    Cad = "DSN=Aridoc;DESC= DSN;DATABASE=aridoc;SERVER=localhost;;PORT=;OPTION=;STMT=;"
'    Cad = "DSN=Aridoc;"" "
'    'cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
'    'cad = cad & ";UID=" & vConfig.User
'    'cad = cad & ";PWD=" & vConfig.password
'
'
'    Conn.ConnectionString = Cad
'    Conn.Open
'    Conn.Execute "Set AUTOCOMMIT = 1"
'    AbrirConexion = True
'    Exit Function
'EAbrirConexion:
'    MsgBox "Abrir conexión." & Err.Description, vbExclamation
'End Function
'
'
'Public Sub Main()
'Dim CadenaComandos As String
'
'
''    If App.PrevInstance Then
''        MsgBox "Ya se esta ejecutando ARIDOC. Tenga paciencia", vbExclamation, "ARIDOC"
''        End
''        Exit Sub
''    End If
'    FormatoFecha = "yyyy-mm-dd"
'
'    CadenaComandos = Command
'    ModoTrabajo = 0
'    SeHaEjecutadoFTP = False
'    'Opcion /s /b     Subir ,bajar un fichero ya estando en aridoc
'
'    'CadenaComandos = "/s erdo 5 S c:\m.jpg"
'
'    'CadenaComandos = "/f root ""Nueva2"" "
'    'CadenaComandos = "/n erdo ""C:\Archivos de programa\Microsoft Visual Studio\Common\Graphics\Bitmaps\Gauge\DOME.bmp"" ""raiz\rama2\obra 10"" "
'    'CadenaComandos = "/a"
'    'CadenaComandos = Trim(CadenaComandos)
'
'    'CadenaComandos = "/u erdo 24 /f1:01/01/2001 /c1:""hola caracola"" /f3:05/06/2005"
'
'    'CadenaComandos = "/f root ""2001-2002"" ""raiz"""
'    'CadenaComandos = "/N root ""C:\Datos\Aridoc\Raiz\2001-2002\SOCIOS\INFORMES\575.iux"" ""raiz\Raiz\2001-2002\SOCIOS\INFORMES"""
'    'CadenaComandos = CadenaComandos & "/F root ""Raiz"" ""raiz"""
'
'
'    If CadenaComandos = "" Then
'        frmInicio.Show
'    Else
'        OpcionesFlag True
'        LanzarShellPedido CadenaComandos
'        '   Rc = vUsu.Leer(Val(Cad))
'
'        If SeHaEjecutadoFTP Then SubCerrarFTP
'        Set Conn = Nothing
'        OpcionesFlag False
'        End   'FUERZO EL FINAL
'    End If
'End Sub
'Private Sub SubCerrarFTP()
'    On Error Resume Next
'     frmMovimientoArchivo.Inet1.Cancel
'     Err.Clear
'End Sub
'Private Sub OpcionesFlag(Poner As Boolean)
'Dim N As String
'Dim NF As Integer
'
'    N = App.Path & "\flag.txt"
'    If Poner Then
'        If Dir(N, vbArchive) = "" Then
'            NF = FreeFile
'            Open N For Output As #NF
'            Print #NF, Now
'            Close #NF
'        End If
'    Else
'        'Eliminar el flag
'        If Dir(N, vbArchive) <> "" Then Kill N
'    End If
'End Sub
'
'
''Realmente este es para insertar
'Public Sub GestionarEquipo()
'Dim Cad As String
'
'
'
'        Set miRSAux = New ADODB.Recordset
'        Cad = "Select max(codequipo) from equipos"
'        miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        vUsu.PC = 1
'        If Not miRSAux.EOF Then vUsu.PC = DBLet(miRSAux.Fields(0), "N") + 1
'        miRSAux.Close
'        Set miRSAux = Nothing
'
'        'De momento un insert normal
'        Cad = "INSERT INTO equipos (codequipo, descripcion,  cargaIconsExt) VALUES (" & vUsu.PC
'        Cad = Cad & ",'" & vUsu.NomPC & "',1)"
'        Conn.Execute Cad
'
'
'
'End Sub
'
'Public Function ComprobacionesPrevias() As Boolean
'Dim MiNombre As String
'Dim MiRuta As String
'
'    On Error GoTo EComprobacionesPrevias
'
'    'Tiene k existir la carpeta Imagenes
'    'imagenes
'    If Dir(App.Path & "\imagenes", vbDirectory) = "" Then MkDir App.Path & "\imagenes"
'    'Tambien la temporal
'    If Dir(App.Path & "\temp", vbDirectory) = "" Then MkDir App.Path & "\temp"
'
'
'    'Tiene k estar vacia
'    'Como algunos les habremos cambiados la extension
'    'La volvemos a poner a lecturaescritura
'    MiRuta = App.Path & "\temp\"
'    MiNombre = Dir(MiRuta, vbDirectory)    ' Recupera la primera entrada.
'    Do While MiNombre <> ""   ' Inicia el bucle.
'       ' Ignora el directorio actual y el que lo abarca.
'       If MiNombre <> "." And MiNombre <> ".." Then
'          ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'          If (GetAttr(MiRuta & MiNombre) And vbDirectory) = vbDirectory Then
'
'          Else
'              SetAttr MiRuta & MiNombre, vbNormal
'              Kill MiRuta & MiNombre
'          End If   ' solamente si representa un directorio.
'       End If
'       MiNombre = Dir   ' Obtiene siguiente entrada.
'    Loop
'
'
'
'
'    ComprobacionesPrevias = True
'
'    Exit Function
'EComprobacionesPrevias:
'    MiRuta = "Comprobaciones Previas" & vbCrLf & Err.Description & vbCrLf & MiRuta & MiNombre
'    MiRuta = MiRuta & vbCrLf & vbCrLf & "No se han podido borrar todos los archivos. ¿Desea continuar de igual modo ?"
'    If MsgBox(MiRuta, vbCritical + vbYesNoCancel) <> vbYes Then
'        ComprobacionesPrevias = False
'    Else
'         ComprobacionesPrevias = True
'    End If
'End Function
'
'

'
'Public Sub Mensajes1(Valor As Integer)
'Dim m As String
'Select Case Valor
'
'Case 0  'Desde hacer copiar
'        m = "No se puede realizar copias desde Histórico."
'Case 1 'Crear diretorio
'        m = "No se puede eliminar la carpeta."
'Case 2 'Eliminar archivos
'        m = "No se pueden eliminar archivos."
'Case 3, 4
'        m = "No se puede Corta&Pegar."
'Case 5 'insertar
'        m = "No se puede insertar nuevos archivos."
'Case 6
'        m = "No se puede realizar MOVER."
'Case 7
'        m = "No se puede importar nuevos archivos."
'Case 8
'        m = "No se pueden crear carpetas nuevas."
'Case 9
'        m = " No hay ningún archivo seleccionado para imprimir."
'Case 10
'        m = " No se puede verificar los documentos de la gestión documental desde el hco"
'Case 11
'        m = " No se puede crear CARPETAS desde historico"
'Case 12
'        m = " No se puede verificar CARPETAS desde historico"
'Case 13
'        m = " No se puede modificar archivos desde historico"
'Case 14
'        m = " No se puede verificar desde historico"
'Case 15
'        m = "Imposible realizar estos cambios en el historico"
'Case 16
'        m = "Opcion no disponible en modo historico"
'End Select
'
'MsgBox m, vbInformation
'End Sub
'
''Para obtener el nompath de un archivo a partir del
''treeview1.fullpath
'Public Function DamePath(CADENA As String)
''Dim aux
''Dim l As Integer
''Dim I As Integer
''
''l = Len(mConfig.Carpeta) + 2
''I = InStr(1, cadena, mConfig.Carpeta)
''If I > 0 Then
''    aux = Mid(cadena, l)
''    Else
''        MsgBox "Error calculando PATH relativo", vbExclamation
''        aux = cadena
''End If
''DamePath = aux
'End Function
'
'
'Public Function devuelvePATH(NomPath As String) As String
'Dim I
'Dim CADENA
'Dim Cad2 As String
'Cad2 = NomPath
'Do
'    I = InStr(1, Cad2, "/")
'    If I > 0 Then
'        CADENA = Mid(Cad2, 1, I - 1)
'        Cad2 = CADENA & "\" & Mid(Cad2, I + 1)
'    End If
'Loop Until I = 0
'devuelvePATH = Cad2
'End Function
'
'
'Public Sub MostrarError(NumeroError As Long, Optional Texto As String)
'Dim Cad
'Cad = "Se ha producido un error." & vbCrLf
'If Texto <> "" Then Cad = Cad & Texto & vbCrLf
'Cad = Cad & "Número: " & NumeroError & vbCrLf
'Cad = Cad & "Descripción: " & Error(NumeroError) & vbCrLf
'MsgBox Cad, vbExclamation
'End Sub
'
'
'
'
'
'''-----------------------------------------------
'''-----------------------------------------------
'''-----------------------------------------------
'''-----------------------------------------------
'''Estas funciones estaban antes en admin, y ahora las hemos sacado
''Public Function CompruebaCarpeta(ByVal kCarpeta As String, men As String) As Byte
''' 1 ---> No tiene archivos
''' 2 ---> Si tiene archivos
''' 3 ---> Tiene SUB directorios
''Dim Aux As String
''Dim miNombre As String
''Dim TieneDir As Boolean
''Dim TieneArch As Boolean
''
''
''TieneDir = False
''TieneArch = False
''Aux = kCarpeta
''miNombre = Dir(Aux, vbDirectory)
''Do While miNombre <> ""
''   If miNombre <> "." And miNombre <> ".." Then
''        If (GetAttr(Aux & miNombre) And vbDirectory) = vbDirectory Then
''            TieneDir = True
''            Exit Do
''            Else
''                TieneArch = True
''                Exit Do
''            End If
''     End If
''   miNombre = Dir ' Obtiene siguiente entrada.
''   Loop
''
''
''If TieneDir Then
''   men = "La carpeta contiene Subcarpetas"
''   CompruebaCarpeta = 3
''   Else
''        If TieneArch Then
''            men = "La carpeta contiene archivos"
''            CompruebaCarpeta = 2
''        Else
''            men = "NO tiene"
''            CompruebaCarpeta = 1
''        End If
''    End If
''End Function
'
'
'
'
'Public Function TratarCarpeta(vCarpeta As String) As Byte
''Dim valor
''Dim directorio As String
''Dim subcarpeta As String
''Dim Fin As Boolean
''Dim ruta As String
''Dim Camino   ' aqui tendremos la ruta de la carpeta que la contiene
''Dim st1 As String ' Para poder llamar a la funcion carpeta
''
''
''TratarCarpeta = 0 ' 0 correcto    1 .- error
''directorio = vCarpeta
''Fin = False
''subcarpeta = ""
''ruta = inicial & Carpeta & "\"
''Camino = ruta
''While Not Fin
''    valor = InStr(1, directorio, "\")
''    If valor = 0 Then
''        Fin = True
''        subcarpeta = directorio
''        Else
''            subcarpeta = Mid(directorio, 1, valor - 1)
''        End If
''
''    ruta = ruta & subcarpeta & "\"
''    If Dir(ruta, vbDirectory) = "" Then ' la carpeta no existe
''        If CompruebaCarpeta(Camino, st1) <> 2 Then
''            MkDir (ruta)
''            SeHanCreadoCarpetas = True
''            Else
''                TratarCarpeta = 1
''                Exit Function
''            End If
''    End If
''    Camino = Camino & subcarpeta & "\"
''    directorio = Mid(directorio, Len(subcarpeta) + 2, Len(directorio))
''Wend
'End Function
'
'
'Private Function CopiaArchivo(NA As String, Car_Des As String) As Byte
''On Error GoTo ErrorCopiaArchivo
''CopiaArchivo = 1
''    FileCopy mConfig.carpetaInt & "\" & NA, Car_Des & "\" & NA
''    Kill mConfig.carpetaInt & "\" & NA
''CopiaArchivo = 0
''Exit Function
''ErrorCopiaArchivo:
''    MsgBox "Se ha producido un error copiando archivo: " & vbCrLf & _
''        "    .-" & mConfig.carpetaInt & "\" & NA & vbCrLf & _
''        "Número: " & Err.Number & vbCrLf & _
''        "Descripción: " & Err.Description, vbExclamation
'End Function
'
''-----------------------------------------------
''-----------------------------------------------
''-----------------------------------------------
''-----------------------------------------------
'
'
'
'
'Public Function ProcesaLinea2(ByRef L As String) As String
'Dim I, C, l2
'Dim J As Byte
'l2 = ""
''Para que no tenga que hacer cada vez el select, y sabiendo que casi todo son letras y numero
''Para saber si lo tenemos que modificar
''comprobaremos que el ASC es mayor 165 para saber si hay que hacer cambios, o no
''If InStr(1, l, "CAMPA") Then Stop
'For I = 1 To Len(L)
'    C = Mid(L, I, 1)
'    J = Asc(C)
'    If J > 125 Then
'        'Caracteres especiales
'        Select Case J
'        Case 165
'            C = "Ñ"
'        Case 166
'            C = "ª"
'        Case 167
'            C = "/"
'        Case 179
'            C = "|"
'        Case 191, 192, 193, 194, 196
'            C = "-"
'        Case 217, 218 ' Estas son las esquinas
'            C = "-"
'        End Select
'    End If
'    l2 = l2 & C
'Next I
'ProcesaLinea2 = l2
'End Function
'
'Public Function espera(Segundos As Single)
'    Dim T1
'    T1 = Timer
'    Do
'    Loop Until Timer - T1 > Segundos
'End Function
'
'
'
'
'
''DIAS:  0.-  Dentro del mes
''       1.-  Hace mas de un mes
''       2.-  hace mas de dos meses
''       3.- NUNCA se ha hecho o fichero no existe
'Public Sub FicheroVerificacion(Grabar As Boolean, ByRef Dias As Byte, Optional ByRef LosErrores As String)
'Dim NF As Integer
'Dim d As Long
'
'On Error GoTo EF
'
'    NF = FreeFile
'    If Grabar Then
'        Open App.Path & "\carpeta.dat" For Output As #NF
'        Print #NF, Format(Now, "dd/mm/yyyy")
'        Print #NF, "Hora: " & Format(Now, "hh:mm")
'        Print #NF, "Revisión carpetas. "
'        If LosErrores = "" Then
'            Print #NF, "Todo bien"
'        Else
'            Print #NF, LosErrores
'        End If
'        Close NF
'    Else
'        If Dir(App.Path & "\carpeta.dat") = "" Then
'            Dias = 3   'Fichero no existe
'        Else
'            Open App.Path & "\carpeta.dat" For Input As #NF
'            Line Input #NF, LosErrores
'            If IsDate(LosErrores) Then
'                'FECHA
'                d = DateDiff("d", Now, CDate(LosErrores))
'
'                If Abs(d) > 62 Then
'                    Dias = 2
'                Else
'                    Dias = 1
'                End If
'            Else
'                Dias = 3 'No podremos precisar
'            End If
'            Close NF
'
'        End If
'    End If
'    Exit Sub
'EF:
'    MsgBox Err.Description
'End Sub
'
'
'
'
'
'
'Public Sub MuestraError(Numero As Long, Optional CADENA As String, Optional Desc As String)
'    Dim Cad As String
'
'    'Con este sub pretendemos unificar el msgbox para todos los errores
'    'que se produzcan
'    On Error Resume Next
'    Cad = "Se ha producido un error: " & vbCrLf
'    If CADENA <> "" Then
'        Cad = Cad & vbCrLf & CADENA & vbCrLf & vbCrLf
'    End If
'
'    If Desc <> "" Then Cad = Cad & vbCrLf & Desc & vbCrLf & vbCrLf
'    Cad = Cad & "Número: " & Numero & vbCrLf & "Descripción: " & Error(Numero)
'    MsgBox Cad, vbExclamation
'End Sub
'
'
'
'
'
'
'
'Public Function DevuelveNombreFichero(campo1 As String, Extension As String, ByRef NombreFinalFichero As String, ParaEmail As Boolean) As Integer
'Dim I As Integer
'Dim Cad As String
'
'    On Error GoTo ED
'
'    Do
'        I = InStr(1, campo1, " ")
'        If I > 0 Then campo1 = Mid(campo1, 1, I - 1) & Mid(campo1, I + 1)
'    Loop Until I = 0
'
'    'QUito comas tambien
'    Do
'        I = InStr(1, campo1, ",")
'        If I > 0 Then campo1 = Mid(campo1, 1, I - 1) & "_" & Mid(campo1, I + 1)
'    Loop Until I = 0
'
'    'quito los dospuntos : por _
'
'    Do
'        I = InStr(1, campo1, ":")
'        If I > 0 Then campo1 = Mid(campo1, 1, I - 1) & "_" & Mid(campo1, I + 1)
'    Loop Until I = 0
'
'
'    I = 0
'    Do
'        If ParaEmail Then
'            Cad = App.Path & "\mail\" & campo1
'        Else
'            Cad = App.Path & "\temp\" & campo1
'        End If
'        If I > 0 Then Cad = Cad & "(" & I & ")"
'        Cad = Cad & "." & Extension
'        I = I + 1
'    Loop Until Dir(Cad, vbArchive) = "" Or I > 100
'    NombreFinalFichero = Cad
'    Exit Function
'ED:
'    MuestraError Err.Number, "Devuelve nombre fichero: " & Err.Description
'    DevuelveNombreFichero = 101
'End Function
'
''Public Function TraerFicheroFisico(ByRef Carpeta As Ccarpetas, Destino As String, codigo As Long) As Boolean
'Public Function TraerFicheroFisico(ByRef Carpeta As Ccarpetas, Destino As String, codigo) As Boolean
'
'
'        TraerFicheroFisico = False
'        'Llevamos el fichero
'        DatosCopiados = "NO"
'        Set frmMovimientoArchivo.vOrigen = Carpeta
'        frmMovimientoArchivo.Opcion = 2
'        frmMovimientoArchivo.Origen = codigo
'        frmMovimientoArchivo.Destino = Destino
'        frmMovimientoArchivo.Show vbModal
'
'        'Y si se producen errores No abrimos
'        If DatosCopiados = "" Then TraerFicheroFisico = True
'
'End Function
'
'Public Function TextoParaComonDialog2(SoloNuevo As Boolean) As String
'Dim SQL As String
'
'    TextoParaComonDialog2 = ""
'    SQL = " SELECT extensionpc.*,descripcion,exten from extensionpc,extension where "
'    SQL = SQL & " extensionpc.codext = extension.codext AND codequipo=" & vUsu.PC
'    If SoloNuevo Then SQL = SQL & " AND extension.nuevo=1"
'    'Que este habilitada
'    SQL = SQL & " AND extension.Deshabilitada =0"
'    SQL = SQL & " order by descripcion" '     'Visor<>""Predeterminado"""
'    Set miRSAux = New ADODB.Recordset
'    miRSAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    SQL = ""
'    If Not miRSAux.EOF Then
'
'
'        While Not miRSAux.EOF
'            SQL = SQL & "|" & miRSAux!Descripcion & "   (*." & miRSAux!Exten & ")|*." & miRSAux!Exten
'            miRSAux.MoveNext
'        Wend
'        SQL = Mid(SQL, 2) 'Quito el primer |
'    End If
'    miRSAux.Close
'    Set miRSAux = Nothing
'    TextoParaComonDialog2 = SQL
'End Function
'
'


Public Sub PonerArrayTiposMensaje()
Dim L As Long
Dim Fin As Integer
Dim i As Integer
Dim J As Integer
Dim Cortar11 As String
'Public Type RegistroTipoMensaje   ' Crea un tipo definido por el usuario.
'   Descripcion As String * 30
'   Color As Long
'End Type
'
'Public ArrayTipoMen() As RegistroTipoMensaje
    TotalTipos = 0
    Cortar11 = "Select count(*) from mailtipo"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cortar11, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Fin = 0
    If Not miRsAux.EOF Then Fin = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If Fin = 0 Then Exit Sub
    
    
    ReDim ArrayTipoMen(Fin)
    TotalTipos = Fin
    
    Cortar11 = "Select * from mailtipo order by tipo "
    miRsAux.Open Cortar11, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 0
    i = 0
    
    
    While Not miRsAux.EOF
        
        If miRsAux!Tipo - J > 1 Then
            J = J + 1
            For Fin = J To miRsAux!Tipo - 1
                ArrayTipoMen(Fin).Color = 0
                ArrayTipoMen(Fin).Descripcion = ""
                ArrayTipoMen(Fin).Icono = 0
            Next Fin
            i = miRsAux!Tipo
        End If
        
        ArrayTipoMen(i).Color = DBLet(miRsAux!Color, "N")
        ArrayTipoMen(i).Descripcion = miRsAux!Descripcion
        ArrayTipoMen(i).Icono = miRsAux!numico
        J = miRsAux!Tipo
        
        miRsAux.MoveNext
        i = i + 1
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

End Sub

'
'Public Sub CodificacionLinea(Leer As Boolean, ByRef Linea As String)
'Dim I As Integer
'Dim C As String
'Dim C2 As String
'    C = Linea
'    Linea = ""
'
'
'        'Escribir
'        For I = 1 To Len(C)
'            C2 = Mid(C, I, 1)
'            If Leer Then
'                C2 = Chr(Asc(C2) - 3)
'            Else
'                C2 = Chr(Asc(C2) + 3)
'            End If
'            Linea = Linea & C2
'        Next I
'End Sub
'
'
'Public Sub AsignarCampoMemo(ByRef Campo As String, ByRef nombrecampo As String, ByRef ADO As ADODB.Recordset)
'    On Error Resume Next
'    Campo = ADO.Fields(nombrecampo).Value
'    If Err.Number <> 0 Then
'        Err.Clear
'        Campo = ""
'    End If
'End Sub
'
