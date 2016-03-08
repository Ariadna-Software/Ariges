VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'EMPRESA
'TABLA:


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mvarCodempre As String 'copia local
Private mvarNomempre As String 'copia local
Private mvarNomresum As String 'copia local
Private mvarAriges As String

Private mvarnumnivel As Integer 'copia local
Private mvarnumdigi1 As Integer 'copia local
Private mvarnumdigi2 As Integer 'copia local
Private mvarnumdigi3 As Integer 'copia local
Private mvarnumdigi4 As Integer 'copia local
Private mvarnumdigi5 As Integer 'copia local
Private mvarnumdigi6 As Integer 'copia local
Private mvarnumdigi7 As Integer 'copia local
Private mvarnumdigi8 As Integer 'copia local
Private mvarnumdigi9 As Integer 'copia local
Private mvarnumdigi10 As Integer 'copia local



'VARIABLES
'Campo calculado de numero de digitos en la ULTIMA NIVEL CONTABLE
Private mvarDigitosUltimoNivel As Byte

'David
'-----------------------------------------------
'Para  no tener que leer "cienes y cienes" de veces
'las fechas de inicio y fin en la contabilidad
'las leo en esta primera lectura de datos
'Ya que si tiene esta tabla(empresa), tiene la de parametros
Private mvarFechaIni As Date
Private mvarFechaFin As Date

'Tambien añadire si lleva analitica
Private mvarTieneAnalitica As Boolean

'Llevara empipado
'   GrupoGasto|GruopVta|OtroGrupoAnalitica|   o lo k es lo mismo: grupogto grupovta Subgrupo1
Private mvarRaizAnalitica As String



Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private RS As ADODB.Recordset
Private RC As Byte




'------------------------------------------------
'Propiedades del modulo Cempresa
'------------------------------------------------

'*** codigo de la empresa

Public Property Let codempre(ByVal vData As String)
     mvarCodempre = vData
End Property

Public Property Get codempre() As String
     codempre = mvarCodempre
End Property


'*** nombre de la empresa

Public Property Let nomempre(ByVal vData As String)
     mvarNomempre = vData
End Property

Public Property Get nomempre() As String
     nomempre = mvarNomempre
End Property


'*** nombre resumen de la empresa
Public Property Let nomresum(ByVal vData As String)
     mvarNomresum = vData
End Property

Public Property Get nomresum() As String
     nomresum = mvarNomresum
End Property


'*** Base de Datos de la empresa

Public Property Let BDAriges(ByVal vData As String)
     mvarAriges = vData
End Property

Public Property Get BDAriges() As String
     BDAriges = mvarAriges
End Property


'*** nivel de las cuentas de contabilidad

Public Property Let numnivel(ByVal vData As Integer)
     mvarnumnivel = vData
End Property

Public Property Get numnivel() As Integer
     numnivel = mvarnumnivel
End Property


'*** digito de 1er nivel de las cuentas de contabilidad
Public Property Let numdigi1(ByVal vData As Integer)
     mvarnumdigi1 = vData
End Property

Public Property Get numdigi1() As Integer
     numdigi1 = mvarnumdigi1
End Property

'*** digito de nivel 2

Public Property Let numdigi2(ByVal vData As Integer)
     mvarnumdigi2 = vData
End Property

Public Property Get numdigi2() As Integer
     numdigi2 = mvarnumdigi2
End Property

'*** digito de nivel 3

Public Property Let numdigi3(ByVal vData As Integer)
     mvarnumdigi3 = vData
End Property

Public Property Get numdigi3() As Integer
     numdigi3 = mvarnumdigi3
End Property

'*** digito de nivel 4
Public Property Let numdigi4(ByVal vData As Integer)
     mvarnumdigi4 = vData
End Property

Public Property Get numdigi4() As Integer
     numdigi4 = mvarnumdigi4
End Property


'*** digito de nivel 8
Public Property Let numdigi8(ByVal vData As Integer)
     mvarnumdigi8 = vData
End Property

Public Property Get numdigi8() As Integer
     numdigi8 = mvarnumdigi8
End Property


'*** digito de nivel *

Public Property Let numdigi9(ByVal vData As Integer)
     mvarnumdigi9 = vData
End Property

Public Property Get numdigi9() As Integer
     numdigi9 = mvarnumdigi9
End Property

'*** digito de nivel 10

Public Property Let numdigi10(ByVal vData As Integer)
     mvarnumdigi10 = vData
End Property

Public Property Get numdigi10() As Integer
     numdigi10 = mvarnumdigi10
End Property


'*** digitos de la cuenta de ultimo nivel

Public Property Get DigitosUltimoNivel() As Byte
     DigitosUltimoNivel = mvarDigitosUltimoNivel
End Property


Public Property Let FechaIni(ByVal vData As Date)
     mvarFechaIni = vData
End Property

Public Property Get FechaIni() As Date
     FechaIni = mvarFechaIni
End Property

Public Property Let FechaFin(ByVal vData As Date)
     mvarFechaFin = vData
End Property

Public Property Get FechaFin() As Date
     FechaFin = mvarFechaFin
End Property




Public Property Let TieneAnalitica(ByVal vData As Boolean)
     mvarTieneAnalitica = vData
End Property

Public Property Get TieneAnalitica() As Boolean
     TieneAnalitica = mvarTieneAnalitica
End Property


'   GrupoGasto|GruopVta|OtroGrupoAnalitica|
'o lo k es lo mismo: grupogto grupovta Subgrupo1
Public Property Get RaizAnalitica() As String
     RaizAnalitica = mvarRaizAnalitica
End Property





'------------------------------------------------
'FUNCIONES del modulo Cempresa
'------------------------------------------------

''****   AGREGAR  ***
'Public Function Agregar() As Byte
''Añadir una nueva empresa en la BD de usuarios tabla:empresasariges
'
'    On Error GoTo Err1
'    'La Base de datos "Usuarios" esta en el mismo servidor que BD "Ariges"
'    SQL = "INSERT INTO usuarios.empresasariges (codempre, nomempre, nomresum)"
'    SQL = SQL & " VALUES ("
'    'Datos
'    SQL = SQL & mvarCodempre
'    SQL = SQL & "," & DBSet(mvarNomempre, "T")
'    SQL = SQL & "," & DBSet(mvarNomresum, "T")
'    SQL = SQL & ")"
'
'    Conn.Execute SQL
'    Agregar = 0
'    Exit Function
'
'Err1:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Agregar = 1
'End Function



'****   LEER  ***
Public Function LeerDatos() As Byte
'Leer los parametros de una empresa
'Lee de la BD: Usuarios, Tabla: EmpresasAriges

    On Error GoTo Err3
    SQL = "SELECT * FROM usuarios.empresasariges WHERE ariges='" & vUsu.CadenaConexion & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If RS.EOF Then
        LeerDatos = 1
    Else
        Mens = "Datos basicos"
        mvarCodempre = RS!codempre
        mvarNomempre = RS!nomempre
        mvarNomresum = DBLet(RS!nomresum)
        mvarAriges = DBLet(RS!AriGes)
        Mens = "Datos de Gestión"
        LeerDatos = 0
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set RS = Nothing
    LeerDatos = 1
End Function


'****   LEER  ***
Public Function LeerNiveles() As Byte
'Leer los parametros de la empresa en BDatos "Conta" correspondiente
'a la empresa de "Ariges" que estamos conectados

    On Error GoTo Err3
    
    SQL = "SELECT * FROM empresa "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If RS.EOF Then
        LeerNiveles = 1
    Else
        Mens = "Datos basicos empresa en contabilidad."
        mvarnumnivel = RS!numnivel
        mvarnumdigi1 = RS!numdigi1
        mvarnumdigi2 = DBLet(RS!numdigi2, "N")
        mvarnumdigi3 = DBLet(RS!numdigi3, "N")
        mvarnumdigi4 = DBLet(RS!numdigi4, "N")
        mvarnumdigi5 = DBLet(RS!numdigi5, "N")
        mvarnumdigi6 = DBLet(RS!numdigi6, "N")
        mvarnumdigi7 = DBLet(RS!numdigi7, "N")
        mvarnumdigi8 = DBLet(RS!numdigi8, "N")
        mvarnumdigi9 = DBLet(RS!numdigi9, "N")
        mvarnumdigi10 = DBLet(RS!numdigi10, "N")
        'Fijamos la longitud maxima de digitos en ctas de ultimo nivel
        FijaLongitudMaxima
        LeerNiveles = 0
    End If
    RS.Close
    
    
    'Asigno los valores por defecto a las variables de fechaini y fechafin
    If LeerNiveles = 0 Then
        LeerNiveles = 2
        SQL = "Select * from parametros"
        RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            MsgBox "Fecha inicio y fin de ejercicio", vbExclamation
        Else
            mvarFechaIni = RS!FechaIni
            mvarFechaFin = RS!FechaFin
            mvarTieneAnalitica = RS!Autocoste = 1
            'grupogto grupovta Subgrupo1
            mvarRaizAnalitica = DBLet(RS!grupogto, "T") & "|" & DBLet(RS!grupovta, "T") & "|" & DBLet(RS!Subgrupo1, "T") & "|"
        End If
        RS.Close
    End If
    
    
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set RS = Nothing
    LeerNiveles = 1
End Function



'****   MODIFICAR  ***
Public Function ModificarDatos() As Byte
'Modifica los parametros de una empresa
'Modifica la Tabla:empresasariges en la BD: Usuarios
    
    On Error GoTo Err4

    'Modificamos en la tabla Usuarios.empresasariges
    SQL = "UPDATE usuarios.empresasariges SET nomempre= '" & mvarNomempre
    SQL = SQL & "', nomresum= '" & mvarNomresum
    SQL = SQL & "' WHERE codempre=" & mvarCodempre
    conn.Execute SQL
    
    ModificarDatos = 0
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    ModificarDatos = 1
End Function



'---------------------------------------------------------
'FUNCIONES del modulo Cempresa referentes a Contabilidad
'---------------------------------------------------------

Private Sub FijaLongitudMaxima()
Dim I As Byte

    Select Case mvarnumnivel
        Case 1
            I = mvarnumdigi1
        Case 2
            I = mvarnumdigi2
        Case 3
            I = mvarnumdigi3
        Case 4
            I = mvarnumdigi4
        Case 5
            I = mvarnumdigi5
        Case 6
            I = mvarnumdigi6
        Case 7
            I = mvarnumdigi7
        Case 8
            I = mvarnumdigi8
        Case 9
            I = mvarnumdigi9
        Case 10
            I = mvarnumdigi10
    End Select
    mvarDigitosUltimoNivel = I
End Sub

