VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'#Generador DAVID(Ariadna)



'-------------------------------------------------------
'-------------------------------------------------------
'Declaracion de las variables
'---- Servidor donde esta MYSQL
Private mvarSERVER As String 'Variable local

'---- Usuario general BD
Private mvarUser As String 'Variable local

'---- Password usuario general BD
Private mvarPassword As String 'Variable local

'---- Explorador a Utilizar para abrir las Webs de Clientes/Proveedores o de la Ayuda
Private mvarExplorador As String
'---- Programa de Correo para abrir en email de Clientes/Proveedores
Private mvarCorreo As String


'--------
'Private mvarPedirPasswd As Boolean

'---- Carpeta Integraciones
'Private mvarIntegraciones As String


'Variables de este modulo
Dim NF As Integer
Dim Mens As String
'variables locales para almacenar los valores de las propiedades
Private mvarAplicacion As String 'copia local

Public Property Let Aplicacion(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Aplicacion = 5
    mvarAplicacion = vData
End Property


Public Property Get Aplicacion() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Aplicacion
    Aplicacion = mvarAplicacion
End Property






'------------------------------------------------------
'------------------------------------------------------
'      Propiedades del modulo de configuracion
'------------------------------------------------------
'------------------------------------------------------
'----
'---- Servidor donde esta MYSQL
'----
Public Property Let SERVER(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarSERVER = vData
End Property

Public Property Get SERVER() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     SERVER = mvarSERVER
End Property

'----
'---- Usuario general BD
'----
Public Property Let User(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.User = 5
     mvarUser = vData
End Property

Public Property Get User() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.User
     User = mvarUser
End Property

'---- Password usuario general BD
Public Property Let password(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarPassword = vData
End Property

Public Property Get password() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     password = mvarPassword
End Property

'---- Explorador
Public Property Let Explorador(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarExplorador = vData
End Property

Public Property Get Explorador() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     Explorador = mvarExplorador
End Property

'---- Programa de Correo
Public Property Let Correo(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarCorreo = vData
End Property

Public Property Get Correo() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     Correo = mvarCorreo
End Property


'---------------------------
' Pedir password
'---------------------------
'Public Property Let PedirPasswd(ByVal vData As Boolean)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.SERVER = 5
'     mvarPedirPasswd = vData
'End Property

'Public Property Get PedirPasswd() As Boolean
''se usa cuando se asigna un valor a la propiedad, en el
''lado derecho de la asignacion
''Syntax: Debug.Print X.SERVER
'     PedirPasswd = mvarPedirPasswd
'End Property


'----
'---- Integraciones
'----
'Public Property Let Integraciones(ByVal vData As String)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.SERVER = 5
'     mvarIntegraciones = vData
'End Property

'Public Property Get Integraciones() As String
''se usa cuando se asigna un valor a la propiedad, en el
''lado derecho de la asignacion
''Syntax: Debug.Print X.SERVER
'     Integraciones = mvarIntegraciones
'End Property


'//////////////////////////////////////////////////////
'             LEER  Y  GUARDAR
'//////////////////////////////////////////////////////
'
'
'
'------------------------------------------------------
'     LEER
'
Public Function Leer() As Byte
    '--------------------------------------------------
    '   Leer:
    '   Busca un fichero de nombre Config.ini y lee de él los valores
    '   de parametrización para la aplicación. El fichero se busca
    '   en el directorio desde el que se ejecuta la aplicación.
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim LinAnt As String
    Dim Lin As String
    Dim NomParam As String ' Para leer el nombre del parámetro
    Dim ValParam As String ' Para leer el valor del parámetro
    Dim NFile As Integer
    
    On Error GoTo Err1
    Fichero = App.Path & "\ConfigAriges.ini"
    NFile = FreeFile
    Open Fichero For Input As #NFile   ' Abre el archivo.
    Do While Not EOF(NFile) ' Repite el bucle hasta el final del archivo.
        Line Input #NFile, Lin
        Select Case Lin
           Case "[SERVER]", "[USER]", "[PASSWORD]", "[PEDIR]", "[INTEG]", "[EXPLORADOR]", "[CORREO]", "[APLICACION]"
               
               ValParam = ""
               NomParam = Lin

           Case "<-------->"
               Select Case NomParam
                    Case "[SERVER]": mvarSERVER = ValParam
                    Case "[USER]": mvarUser = ValParam
                    Case "[PASSWORD]": mvarPassword = ValParam
                    Case "[EXPLORADOR]": mvarExplorador = ValParam
                    Case "[CORREO]": mvarCorreo = ValParam
                    Case "[APLICACION]": mvarAplicacion = ValParam
'                    Case "[PEDIR]": mvarPedirPasswd = CBool(ValParam)
'                    Case "[INTEG]": mvarIntegraciones = ValParam
                End Select
            Case Else
                If ValParam = "" Then
                    ValParam = Lin
                Else
                    ValParam = ValParam & vbCrLf & Lin
                End If
        End Select
    Loop
    Close #NFile    ' Cierra el archivo.
    
    'Desencriptamos password
    Desencripta (True)
    Desencripta (False) 'Desencripta el usuario
    Leer = 0
    Exit Function
Err1:
   Leer = 1
End Function




'------------------------------------------------------
'     GUARDAR

Public Function Grabar() As Byte
    '--------------------------------------------------
    '   GuardaConfig:
    '   Guardamos en el fichero Config.ini los valores de los
    '   parámetros de la aplicación
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim NFile As Integer

    On Error GoTo Err1
    Fichero = App.Path & "\ConfigAriges.ini"
    NFile = FreeFile
    Open Fichero For Output As #NFile   ' Abre el archivo.
    Print #NFile, "[SERVER]"
    Print #NFile, mvarSERVER
    Print #NFile, "<-------->"
    Print #NFile, "[USER]"
    Print #NFile, Encripta(mvarUser)
    Print #NFile, "<-------->"
    Print #NFile, "[PASSWORD]"
    Print #NFile, Encripta(mvarPassword)
    Print #NFile, "<-------->"
    Print #NFile, "[EXPLORADOR]"
    Print #NFile, "C:\Archivos de programa\Internet Explorer\IEXPLORE.EXE"
    Print #NFile, "<-------->"
    Print #NFile, "[CORREO]"
    Print #NFile, "C:\Archivos de programa\Microsoft Office\OFFICE11\OUTLOOK.EXE"
    Print #NFile, "<-------->"
    
    
'    Print #NFile, "[PEDIR]"
'    Print #NFile, mvarPedirPasswd
'    Print #NFile, "<-------->"
'    Print #NFile, "[INTEG]"
'    Print #NFile, mvarIntegraciones
'    Print #NFile, "<-------->"
       
    Close #NFile
    Grabar = 0
    Exit Function
Err1:
    Mens = "Error " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbCritical
    Grabar = 1
End Function


Private Sub Desencripta(ElPassword As Boolean)
Dim C As String
Dim i As Integer

If ElPassword Then
    C = mvarPassword
    Else
    C = mvarUser
End If
Mens = ""
For i = 1 To Len(C)
    If (i Mod 2) <> 0 Then
        'Solo valen los caracteres impares
        NF = Asc(Mid(C, i, 1))
        NF = NF + 1
        Mens = Mens & Chr(NF)
    End If
Next i

If ElPassword Then
    mvarPassword = Mens
    Else
    mvarUser = Mens
End If

End Sub



Private Function Encripta(ByRef C As String) As String
Dim i As Byte

Mens = ""
Randomize
For i = 1 To Len(C)
        'Solo valen los caracteres impares
        NF = Asc(Mid(C, i, 1))
        NF = NF - 1
        Mens = Mens & Chr(NF)
        'Y el aleatorio , entre 48 y 122  <- 74 ->
        NF = Int((74 * Rnd)) ' Genera valores aleatorios entre 0y74.
        NF = NF + 48
        Mens = Mens & Chr(NF)
Next i
Encripta = Mens

End Function
