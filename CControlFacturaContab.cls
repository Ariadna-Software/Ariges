VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CControlFacturaContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------
'
'   Modulo comprobacion facturas a contabilizar
'
'   Metodos:
'   .-FechaCorrectaContabilizazion(ByRef CnnCta As Connection, FechaContab As Date) as string
'           parametros: Conexion abierta a la contabilidad
'                       Fecha contabilizacion
'           devuelve: "" =todook
'                     en otro caso el error de pk la fecha no es correcta para integrar
'
'
'
'   .-FechaCorrectaIVA(ByRef CnnCta As Connection, FechaContab As Date)
'           parametros: Conexion abierta a la contabilidad
'                       Fecha contabilizacion
'           devuelve una cadena diciendo si esta liquidado
'
'
'
'
'
'   .-CuentasBloqueadas(ByRef CnnCta As Connection, ByVal CadenaCuentas As String, FechaContab As Date, ByRef CtaBloq As Collection)
'           parametros: Conexion abierta a la contabilidad
'                       CadenaCuentas: Cuentas separadas entre si por comas
'                       Fecha contabilizacion
'           se le pasa un objeto CtaBloq collection
'           si hay cuetas bloquedas insertara en CtaBloq una linea por cada cta con codmacta descripcion
'
'
'
'
'   .-FechaRecepMenorQueProveedor(ByRef CnnCta As Connection, FechaRecpcion As Date)
'       True si existe una fecharecepcion mayor que la que queremos contabilizar
'
'
'
'
'   .-FechaMenorUltimaFacturaCliente(ByRef CnnCta As Connection, FechaFactura As Date, Serie As String)
'       True si existe una fecha para esa serie mayor que la que quiero contabilizar




'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'
' SI SE MODIFICA HAY QUE CAMONICARLO A OTROS PROYECTOS
'  arioli, ariagro....
'
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************











Private DatosCargados As Boolean
Private mvarFechaIni As Date
Private mvarFechaFin As Date
Private mvarFechaActiva As Date
Private mvarUltMEsIVA As Integer
Private mvarUltAnyoIVA As Integer
Private mvarEsIvaMensual As Boolean





'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'
'
'   Publico
'
'
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Devolvera el ERROR
' Comprobara si esta en el ejercicio(actual y siguiente) y con la fecha activa
Public Function FechaCorrectaContabilizazion(ByRef CnnCta As Connection, FechaContab As Date) As String
Dim Cad As String
    
    Cad = LeerDatos(CnnCta)
    If Cad <> "" Then
        'Error leyendo parametros
        FechaCorrectaContabilizazion = Cad
        Exit Function
    End If
    
    
    
    If FechaContab < mvarFechaIni Then
        FechaCorrectaContabilizazion = "Fecha anterior a inicio de ejercicio"
        Exit Function
    End If
    
    If FechaContab < mvarFechaActiva Then
        FechaCorrectaContabilizazion = "Fecha anterior a fecha activa"
        Exit Function
    End If
    
    If FechaContab > mvarFechaFin Then
        FechaCorrectaContabilizazion = "Fecha fuera de ejercicios"
        Exit Function
    End If
    
    
    FechaCorrectaContabilizazion = ""  'Correcto
    
    
End Function

'Lo he separado pk podria darse el caso de que QUIERA continuar, a sabiendas que esta liquidado
Public Function FechaCorrectaIVA(ByRef CnnCta As Connection, FechaContab As Date) As String
Dim FechaUltIVA As Date
Dim J As Integer
Dim Cad As String

    Cad = LeerDatos(CnnCta)
    If Cad <> "" Then
        'Error leyendo parametros
        FechaCorrectaIVA = Cad
        Exit Function
    End If
    
    
    'Cojo el dia uno del mes utlimo
    'le sunmo un mes y le quito un dia
    
    If mvarEsIvaMensual Then
       J = mvarUltMEsIVA
    Else
        'Trimestral
        J = mvarUltMEsIVA * 3
    End If
    FechaUltIVA = "01/" & J & "/" & mvarUltAnyoIVA
    FechaUltIVA = DateAdd("m", 1, FechaUltIVA)  'mas un mes
    FechaUltIVA = DateAdd("d", -1, FechaUltIVA)  'menos un dia

    If FechaContab <= FechaUltIVA Then
        FechaCorrectaIVA = "Fecha pertence a un periodo liquidado"
    Else
        'OK
        FechaCorrectaIVA = ""
    End If
    
End Function


'Cnn. La conexion a
'Cadenacuentas  llevara (separada por comas) las cuentas que quiero comprobar si es
'
Public Function CuentasBloqueadas(ByRef CnnCta As Connection, ByVal CadenaCuentas As String, FechaContab As Date, ByRef CtaBloq As Collection)
Dim SQL As String
Dim Aux As String
Dim i As Integer
Dim RT As ADODB.Recordset


    Set CtaBloq = New Collection
    Set RT = New ADODB.Recordset
    Do
        'Para reestringir la select un poco
        If Len(CadenaCuentas) > 450 Then
            i = InStr(450, CadenaCuentas, ",")
            If i > 0 Then
                Aux = Mid(CadenaCuentas, 1, i - 1)
                CadenaCuentas = Mid(CadenaCuentas, i + 1)
            Else
                Aux = CadenaCuentas
                CadenaCuentas = ""
            End If
        Else
            Aux = CadenaCuentas
            CadenaCuentas = ""
        End If
        
        SQL = "Select * from cuentas WHERE apudirec='S' and fecbloq<='" & Format(FechaContab, "yyyy-mm-dd") & "'"
        SQL = SQL & " AND codmacta IN (" & Aux & ")"
        RT.Open SQL, CnnCta, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RT.EOF
            'Devolver empipado ciertos campos, para que pueden mostrar los que quieran
            SQL = RT!Codmacta & "|" & RT!nommacta & "|" & RT!fecbloq & "|"
            CtaBloq.Add SQL
            RT.MoveNext
        Wend
        RT.Close
    Loop Until CadenaCuentas = ""
    Set RT = Nothing
End Function

Public Function FechaMenorUltimaFacturaCliente(ByRef CnnCta As Connection, FechaFactura As Date, Serie As String) As Boolean
    FechaMenorUltimaFacturaCliente = FechaUltFacturaMayor(CnnCta, FechaFactura, Serie)
End Function

Public Function FechaRecepMenorQueProveedor(ByRef CnnCta As Connection, FechaRecpcion As Date) As Boolean
    FechaRecepMenorQueProveedor = FechaUltFacturaMayor(CnnCta, FechaRecpcion, "")
End Function


Private Function FechaUltFacturaMayor(ByRef CnnCta As Connection, fecFac As Date, Serie As String) As Boolean
Dim RT As ADODB.Recordset
Dim SQL As String
Dim LaFecha As Date

    Set RT = New ADODB.Recordset
    If Serie = "" Then
        SQL = "select max(fecrecpr) as fecha from cabfactprov "
            
    Else
        'Estamos en proveedores
        SQL = "Select max(fecfaccl)as fecha from cabfact WHERE numserie = '" & Serie & "'"
    End If
    LaFecha = CDate("01/01/1900")
    RT.Open SQL, CnnCta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT!Fecha) Then LaFecha = RT!Fecha
    End If
    RT.Close
    Set RT = Nothing
    
    
    If LaFecha > fecFac Then
        'ERROR. La fecha factura menor k una contabilizada
        'If Serie = "" Then
        '   ' FechaUltFac = "Fecha recepcion menor que la fecha recepcion en contabilidad"
        'Else
        '   ' FechaUltFac = "Fecha factura menor que la fecha factura en contabilidad"
        'End If
        FechaUltFacturaMayor = True
    Else
        FechaUltFacturaMayor = False
    End If
End Function



Private Function LeerDatos(ByRef CN As Connection) As String
Dim RS As ADODB.Recordset
Dim SQL As String

    If DatosCargados Then
        LeerDatos = "" 'Ya estaban cargados
        Exit Function
    End If
    
    SQL = "Select * from parametros "
    Set RS = New ADODB.Recordset
    RS.Open SQL, CN, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        'MSGBOX
        LeerDatos = "No existe parametros para la empresa"
    Else
        mvarFechaIni = RS!FechaIni
        mvarFechaFin = RS!FechaFin
        mvarFechaFin = DateAdd("yyyy", 1, mvarFechaFin)
        If IsNull(RS!FechaActiva) Then
            mvarFechaActiva = mvarFechaIni
        Else
            mvarFechaActiva = RS!FechaActiva
        End If
        'Utlimo peridodo liquidado
        mvarUltMEsIVA = DBLet(RS!perfactu, "N")
        If mvarUltMEsIVA = 0 Then mvarUltMEsIVA = 1
        mvarUltAnyoIVA = DBLet(RS!anofactu, "N")
        If mvarUltAnyoIVA = 0 Then mvarUltAnyoIVA = 1900
        
        mvarEsIvaMensual = DBLet(RS!periodos, "N") = 1

    
        
        LeerDatos = ""
        DatosCargados = True
    End If
    RS.Close
    
    
End Function

Private Sub Class_Initialize()
    DatosCargados = False
End Sub
