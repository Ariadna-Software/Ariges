VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CLIENTES
'TABLA: sclien


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodClien As String 'Cod. cliente
Private mNomClien As String 'Nombre cliente
Private mNomComer As String 'Nombre comercial cliente
Private mDomClien As String 'Domicilio cliente
Private mCodpobla As String 'cod. postal cliente
Private mPobClien As String 'poblacion cliente
Private mProClien As String 'Provincia cliente
Private mNIFClien As String 'NIF cliente
Private mTfnClien As String 'Telefono administracion
Private mMaiClien As String 'Mail administracion

Private mCodEnvio As String 'cod. envio
Private mCodAgent As String 'cod. agente
Private mCodForpa As String 'cod. forma de pago
Private mDtoPpago As String 'descuento pronto pago
Private mDtoGnral As String 'descuento general
Private mTipoFact As Byte   'tipo facturacion (0:factura colectiva, 1:facturaxAlbaran

Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria
'Noviembre 2013
Private mIBAN As String


Private mObservac As String 'Observaciones cliente
Private mCliVario As Boolean 'si es cliente de varios

Private mCodTarif As Integer 'cod. tarifa de precios
Private mCodSitua As Byte 'codigo de situacion
Private mKilometr As Integer 'kimlometros

Private mFecUltMovim As String 'Fecha del ult. movimiento del cliente

Private mWwwClien As String 'web del cliente

Private mZona As String 'web del cliente

Private mPersonaContacto As String


Private mDirEnvioHabitual As Integer   ' Si no trabajamos con direcciones de envio, o no la tiene configurada tendra un -1

'------------------------------------------------
'Propiedades del modulo CCliente
'------------------------------------------------

'**** Codigo del cliente

Public Property Let codigo(ByVal vData As String)
     mCodClien = vData
End Property

Public Property Get codigo() As String
     codigo = mCodClien
End Property


'**** Nombre del cliente

Public Property Let Nombre(ByVal vData As String)
     mNomClien = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomClien
End Property


'**** Domicilio del cliente

Public Property Let Domicilio(ByVal vData As String)
     mDomClien = vData
End Property

Public Property Get Domicilio() As String
     Domicilio = mDomClien
End Property

'**** CPostal del cliente

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property

'**** Poblacion del cliente

Public Property Let Poblacion(ByVal vData As String)
     mPobClien = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobClien
End Property

'**** Provincia del cliente

Public Property Let Provincia(ByVal vData As String)
     mProClien = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProClien
End Property

'**** NIF del cliente

Public Property Let NIF(ByVal vData As String)
     mNIFClien = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFClien
End Property

'**** Telefono administracion del cliente

Public Property Let TfnoClien(ByVal vData As String)
     mTfnClien = vData
End Property

Public Property Get TfnoClien() As String
     TfnoClien = mTfnClien
End Property

'**** e-MAil administracion del cliente

Public Property Let EMailAdm(ByVal vData As String)
     mMaiClien = vData
End Property

Public Property Get EMailAdm() As String
     EMailAdm = mMaiClien
End Property


'**** Web del cliente

Public Property Let WebClien(ByVal vData As String)
     mWwwClien = vData
End Property

Public Property Get WebClien() As String
     WebClien = mWwwClien
End Property



'**** Tipo Facturacion del cliente

Public Property Let TipoFactu(ByVal vData As Byte)
     mTipoFact = vData
End Property

Public Property Get TipoFactu() As Byte
     TipoFactu = mTipoFact
End Property

'**** Cod. Forma Envio del Cliente

Public Property Let FEnvio(ByVal vData As String)
     mCodEnvio = vData
End Property

Public Property Get FEnvio() As String
     FEnvio = mCodEnvio
End Property


'**** Agente del Cliente

Public Property Let Agente(ByVal vData As String)
     mCodAgent = vData
End Property

Public Property Get Agente() As String
     Agente = mCodAgent
End Property


'**** Forma de pago del Cliente

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Descuento pronto pago del Cliente

Public Property Let DtoPPago(ByVal vData As String)
     mDtoPpago = vData
End Property

Public Property Get DtoPPago() As String
     DtoPPago = mDtoPpago
End Property

'**** Descuento general del Cliente

Public Property Let DtoGnral(ByVal vData As String)
     mDtoGnral = vData
End Property

Public Property Get DtoGnral() As String
     DtoGnral = mDtoGnral
End Property


'**** Banco del Cliente

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal del Cliente

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la cuenta del cliente

Public Property Let DigControl(ByVal vData As String)
     mDigContr = vData
End Property

Public Property Get DigControl() As String
     DigControl = mDigContr
End Property


'**** Cuenta Bancaria de la Factura

Public Property Let CuentaBan(ByVal vData As String)
     mCuentaba = vData
End Property

Public Property Get CuentaBan() As String
     CuentaBan = mCuentaba
End Property

Public Property Let IBAN(ByVal vData As String)
     mIBAN = vData
End Property

Public Property Get IBAN() As String
     IBAN = mIBAN
End Property


'**** Observaciones del cliente

Public Property Let Observaciones(ByVal vData As String)
     mObservac = vData
End Property

Public Property Get Observaciones() As String
     Observaciones = mObservac
End Property


'**** cliente de varios (SI/NO)

Public Property Let DeVarios(ByVal vData As Boolean)
     mCliVario = vData
End Property

Public Property Get DeVarios() As Boolean
     DeVarios = mCliVario
End Property



'**** Tarifa precios del cliente

Public Property Let Tarifa(ByVal vData As Integer)
     mCodTarif = vData
End Property

Public Property Get Tarifa() As Integer
     Tarifa = mCodTarif
End Property



'**** situacion cliente

Public Property Let Situacion(ByVal vData As Byte)
     mCodSitua = vData
End Property

Public Property Get Situacion() As Byte
     Situacion = mCodSitua
End Property


'**** Kilometros cliente

Public Property Let Kilometros(ByVal vData As Integer)
     mKilometr = vData
End Property

Public Property Get Kilometros() As Integer
     Kilometros = mKilometr
End Property



'**** Fecha ult. movimiento del cliente

Public Property Let FechaUltMov(ByVal vData As String)
     mFecUltMovim = vData
End Property

Public Property Get FechaUltMov() As String
     FechaUltMov = mFecUltMovim
End Property




Public Property Let PersonaContacto(ByVal vData As String)
     mPersonaContacto = vData
End Property

Public Property Get PersonaContacto() As String
     PersonaContacto = mPersonaContacto
End Property

'Zona por defecto
Public Property Let Zona(ByVal vData As String)
     mZona = vData
End Property

Public Property Get Zona() As String
     Zona = mZona
End Property



'**** Direccion envio habitual
Public Property Let DirEnvioHabitual(ByVal vData As Integer)
     mDirEnvioHabitual = vData
End Property

Public Property Get DirEnvioHabitual() As Integer
     DirEnvioHabitual = mDirEnvioHabitual
End Property

'------------------------------------------------
'Procedimientos del modulo CCliente
'------------------------------------------------

Public Function Existe(vCodClien As String) As Boolean
'Comprueba si existe el cliente en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBD(conAri, "codclien", "sclien", "codclien", vCodClien, "N")
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el Cliente " & vCodClien & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function



Public Function LeerDatos(vCodClien As String) As Boolean
'Leer los datos de un cliente dado
'Lee de la BD: Ariges, Tabla: sclien
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    
    LeerDatos = False
    
    SQL = "SELECT codclien,nomclien,nomcomer,codpobla,domclien,pobclien,proclien,nifclien,wwwclien,codenvio,codagent,codforpa,codbanco,codsucur,digcontr,cuentaba,telclie1,maiclie1,clivario,dtoppago,dtognral,codtarif,codsitua,tipofact,kilometr,fechamov,observac,perclie1,codzonas,coddirenhab,iban "
    SQL = SQL & " FROM sclien "
    SQL = SQL & " WHERE codclien=" & vCodClien
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Rs.EOF Then
        LeerDatos = False
    Else
        mCodClien = CStr(Rs!codclien)
        
        mNomClien = Rs!Nomclien
        mNomComer = Rs!nomcomer
        mDomClien = DBLet(Rs!domclien, "T")
        mCodpobla = DBLet(Rs!codpobla, "T")
        mPobClien = DBLet(Rs!pobclien, "T")
        mProClien = DBLet(Rs!proclien, "T")
        mNIFClien = DBLet(Rs!nifClien, "T")
        mWwwClien = DBLet(Rs!wwwclien, "T")
        mTfnClien = DBLet(Rs!telclie1, "T")
        mMaiClien = DBLet(Rs!maiclie1, "T")
       
        mCodEnvio = CStr(DBLet(Rs!CodEnvio, "N"))
        mCodAgent = CStr(Rs!codagent)
        mCodForpa = CStr(Rs!codforpa)
        mDtoPpago = CStr(Rs!DtoPPago)
        mDtoGnral = CStr(Rs!DtoGnral)
        mTipoFact = Rs!tipofact
        
        mCodbanco = CStr(DBLet(Rs!codbanco, "N"))
        mCodsucur = CStr(DBLet(Rs!codsucur, "N"))
        mDigContr = DBLet(Rs!digcontr, "T")
        mCuentaba = DBLet(Rs!cuentaba, "T")
        mIBAN = DBLet(Rs!IBAN, "T")
        
        mObservac = DBLetMemo(Rs!observac)
        mCliVario = Rs!CliVario
        
        Tarifa = Rs!codTarif
        Situacion = Rs!codsitua
        mKilometr = DBLet(Rs!kilometr, "N")

        mFecUltMovim = DBLet(Rs!FechaMov, "T")
        
        mPersonaContacto = DBLet(Rs!perclie1, "T")
        
        mZona = CStr(DBLet(Rs!codzonas, "N"))
        
        
        'Direccion envio habitual
        mDirEnvioHabitual = -1
        
        
        If vParamAplic.DireccionesEnvio Then
            If Not IsNull(Rs!coddirenhab) Then mDirEnvioHabitual = Rs!coddirenhab
        End If
        
        
        
        LeerDatos = True
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Cliente" & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set Rs = Nothing
    LeerDatos = False
End Function



Public Function LeerDatosCliVario(vNIF As String) As Boolean
'si es cliente de varios recuperamos los datos nombre,direc,etc.
'de la tabla de varios: sclvar
Dim Rs As ADODB.Recordset
Dim SQL As String
    
    On Error GoTo ELeer
    LeerDatosCliVario = False
    
    SQL = "SELECT nifclien,nomclien,domclien,codpobla,pobclien,proclien,telclien "
    SQL = SQL & " FROM sclvar "
    SQL = SQL & " WHERE nifclien=" & DBSet(vNIF, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    mPersonaContacto = ""
    If Rs.EOF Then
        LeerDatosCliVario = False
        
        mNomClien = ""
        mNomComer = ""
        mDomClien = ""
        mCodpobla = ""
        mPobClien = ""
        mProClien = ""
        mTfnClien = ""
        MsgBox "No existe el cliente vario: " & vNIF, vbInformation
    Else
        mNIFClien = Rs!nifClien
        mNomClien = Rs!Nomclien
        mNomComer = ""
        mDomClien = Rs!domclien
        mCodpobla = DBLet(Rs!codpobla, "T")
        mPobClien = DBLet(Rs!pobclien, "T")
        mProClien = DBLet(Rs!proclien, "T")
        mTfnClien = DBLet(Rs!telclien, "T")
        
        LeerDatosCliVario = True
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Cliente de Varios." & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set Rs = Nothing
    LeerDatosCliVario = False
End Function



Public Function ActualizarClienteV(vNIF As String) As Boolean
'Actualiza los datos del cliente de varios en la tabla: sclvar
Dim SQL As String
Dim devuelve As String

    On Error GoTo EActualizarCliente
    ActualizarClienteV = False
    
    devuelve = DevuelveDesdeBDNew(conAri, "sclvar", "nifclien", "nifclien", vNIF, "T")
    
    If devuelve <> "" Then 'Existe el cliente de varios y lo actualizamos
        SQL = "UPDATE sclvar SET nomclien=" & DBSet(Nombre, "T") & ", domclien=" & DBSet(Domicilio, "T") & ", "
        SQL = SQL & " codpobla=" & DBSet(CPostal, "T") & ", pobclien=" & DBSet(Poblacion, "T") & ", "
        SQL = SQL & " proclien=" & DBSet(Provincia, "T") & ", telclien=" & DBSet(TfnoClien, "T")
        SQL = SQL & " WHERE nifclien=" & DBSet(vNIF, "T")
    Else 'No existe el cliente y lo añadimos
        SQL = "INSERT INTO sclvar (nifclien, nomclien, domclien, codpobla, pobclien, proclien, telclien) "
        SQL = SQL & " VALUES (" & DBSet(vNIF, "T") & ", " & DBSet(Nombre, "T") & ", " & DBSet(Domicilio, "T") & ", "
        SQL = SQL & DBSet(CPostal, "T") & ", " & DBSet(Poblacion, "T") & ", " & DBSet(Provincia, "T") & ", " & DBSet(TfnoClien, "T") & ")"
    End If
    conn.Execute SQL, , adCmdText
    
    ActualizarClienteV = True
    
EActualizarCliente:
    If Err.Number <> 0 Then
        ActualizarClienteV = False
        MuestraError Err.Number, "Actualizando Cliente de Varios.", Err.Description
    End If
End Function


Public Function ClienteBloqueado(Optional DesdeOFertasPedidos As Boolean) As Boolean
'NUEVO
'Diciembre 2010
'Si esta en ofertas / pedidos CLIENTE buscaremos en la tabla el campo: clioferped
' Si es 1 tb bloquea la creacion de ofertas / pedidos


'devuelve true si el cliente esta bloqueado
'si la situación del cliente es distinta de NORMAL(codsitua=0) entonces
'mostrar un mensaje con la situación especial del cliente
Dim Tipo As String
Dim devuelve As String

    On Error GoTo EBloqueado
    ClienteBloqueado = False
    
    If Situacion <> 0 Then '0: situacion normal
        If DesdeOFertasPedidos Then
            Tipo = "clioferped"
        Else
            Tipo = "tipositu"
        End If
        devuelve = DevuelveDesdeBDNew(conAri, "ssitua", "nomsitua", "codsitua", Situacion, "N", Tipo)
        
        If Tipo = "1" Then 'Cliente Bloqueado por Situación Especial.
            MsgBox UCase("Cliente Bloqueado por: ") & vbCrLf & devuelve, vbInformation, "Situación Especial del Cliente."
            ClienteBloqueado = True
        Else
            MsgBox devuelve, vbInformation, "Situación Especial del Cliente."
        End If
    End If
    
EBloqueado:
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function DptoCliente(vCodDpto As String, NomDpto As String) As Boolean
'comprueba si existe el departamento para el cliente y devuelve su nombre
'IN: vCodDpto= codigo del departamento o direccion del cliente
'OUT: NomDppto= devuelve el nombre del dpto o direccion si lo encuentra
Dim devuelve As String
Dim cadDpto As String
    
    If codigo = "" Then
        devuelve = "Debe introducir valor al campo Cliente, para cargar sus "
        If vParamAplic.HayDeparNuevo = 1 Then 'Se trabaja con dpto
            devuelve = devuelve & "Departamentos. "
        ElseIf vParamAplic.HayDeparNuevo = 0 Then
            'la aplicacion trabaja con direcciones
            devuelve = devuelve & "Direcciones. "
        Else
            'la aplicacion trabaja con obras
            devuelve = devuelve & "Obras. "
        End If
        MsgBox devuelve, vbInformation
        Exit Function
    End If

    'Comprobar que el cliente seleccionado tiene esa direccion
    devuelve = DevuelveDesdeBDNew(conAri, "sdirec", "nomdirec", "codclien", codigo, "N", , "coddirec", vCodDpto, "N")
    NomDpto = devuelve
    If devuelve = "" Then
        DptoCliente = False
        
        If vParamAplic.HayDeparNuevo = 1 Then 'Se trabaja con dpto
            cadDpto = " el Departamento "
        ElseIf vParamAplic.HayDeparNuevo = 0 Then
            'la aplicacion trabaja con direcciones
            cadDpto = " la Dirección "
        Else
            cadDpto = " la obra "
        End If
        
        devuelve = "No existe" & cadDpto & vCodDpto & " para el cliente: "
        devuelve = devuelve & codigo
        MsgBox devuelve, vbInformation
    Else
        DptoCliente = True
    End If
End Function


Public Function TieneRefObligatoria(vReferen As String) As Boolean
'Comprueba si el campo de referencia obligatoria es necesario que se rellene
'al introducir datos de ese cliente
Dim devuelve As String

    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "referobl", "codclien", codigo, "N")
    If devuelve = "1" And vReferen = "" Then 'Referencia Obligatoria
        TieneRefObligatoria = True
        devuelve = "La Referencia del Cliente es Obligatoria."
        MsgBox devuelve, vbInformation
    Else
        TieneRefObligatoria = False
    End If
End Function


Public Function exentoIVA() As Byte
'Comprueba si el cliente esta exento de IVA para calcular el %IVA de las ofertas,pedidos
'albaranes, facturas,...
'Dim devuelve As String

    exentoIVA = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", codigo, "N")
    
    '0.- Normal    1.- R.equival     2 Exento     3.- Exento Intracom   4.- Reducido
    
    
    
    
    
End Function


Public Function Bonifica() As Boolean
'Indica si al cliente se le aplican bonificaciones o no
Dim Cad As String

    On Error GoTo EBonifica
    Bonifica = False
    
    Cad = DevuelveDesdeBDNew(conAri, "sclien", "codtarif", "codclien", codigo, "N")
    If Cad <> "" Then
        Cad = DevuelveDesdeBDNew(conAri, "starif", "bonifica", "codlista", Cad, "N")
        If Cad <> "" Then Bonifica = CBool(Cad)
    End If
    
EBonifica:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar bonificación del cliente.", Err.Description
    End If
End Function


Public Function ActualizaUltFecMovim(vFecha As String) As Boolean
'Actualiza la fecha del ultimo movimiento del cliente si la fecha del movimiento
'que se pasa por parametro es posterior a la que tiene el cliente
Dim SQL As String
Dim b As Boolean
Dim actualiza As Boolean

    On Error GoTo EActFecha
    
    If FechaUltMov = "" Then _
        FechaUltMov = DevuelveDesdeBDNew(conAri, "sclien", "fechamov", "codclien", codigo, "N")
    
    
    If FechaUltMov = "" Then
        actualiza = True
    ElseIf CDate(vFecha) > CDate(FechaUltMov) Then
        actualiza = True
    Else
        actualiza = False
    End If
    
    If actualiza Then
        SQL = "UPDATE sclien SET fechamov=" & DBSet(vFecha, "F")
        SQL = SQL & " WHERE codclien=" & codigo
        conn.Execute SQL, , adCmdText
    End If
    b = True
        
EActFecha:
    If Err.Number <> 0 Then b = False
    ActualizaUltFecMovim = b
End Function




Public Sub MostrarObservaciones()
'Recuperamos el texto de las observaciones del cliente, y si tiene
'valor lo mostramos en un mensaje
    If Me.Observaciones <> "" Then _
         MsgBox Me.Observaciones, vbInformation, "Observaciones del cliente"
End Sub




Public Function ComprobarCtaBancaria(cadMen As String) As Boolean
'comprueba q el cliente tenga cuenta bancaria y que sea correcta
Dim Cta As String

    If Me.Banco = 0 And Me.Sucursal = 0 And Me.DigControl = "" And Me.CuentaBan = "" Then
        ComprobarCtaBancaria = False
        cadMen = "El cliente no tiene asignada cuenta bancaria."
        Exit Function
    End If
    
    
    
    Cta = Format(Me.Banco, "0000") & Format(Me.Sucursal, "0000") & Format(Me.DigControl, "00") & Format(Me.CuentaBan, "0000000000")
    
    If Val(ComprobarCero(Cta)) = 0 Then
        ComprobarCtaBancaria = False
        cadMen = "El cliente no tiene asignada cuenta bancaria."
        Exit Function
    End If
    
    
    If Not Comprueba_CuentaBan2(Cta, True) Then
        cadMen = "La cuenta bancaria del cliente no es correcta."
        ComprobarCtaBancaria = False
    Else
        ComprobarCtaBancaria = True
    End If
End Function



Public Function Obtener_EMailConfirmacion(Optional codDpto As String) As String
'Obtener el email confirmacion para el cliente-departamento
Dim dirMail As String

    dirMail = ""
    
    '1º) email para el cliente-departamento
    If codDpto <> "" Then
        dirMail = DevuelveDesdeBDNew(conAri, "sdirec", "maidirec", "codclien", Me.codigo, "N", , "coddirec", codDpto, "N")
    End If
    
    '2º) si no email comercial cliente
    If dirMail = "" Then
        dirMail = DevuelveDesdeBDNew(conAri, "sclien", "maiclie2", "codclien", Me.codigo, "N")
    End If
    
    '3º) si no email administracion cliente
    If dirMail = "" Then
        dirMail = DevuelveDesdeBDNew(conAri, "sclien", "maiclie1", "codclien", Me.codigo, "N")
    End If
    
    Obtener_EMailConfirmacion = dirMail
End Function


Public Sub PonDatosDireccionEnvio(ByRef txtCod As TextBox, txtName As TextBox)
    If vParamAplic.DireccionesEnvio Then
        
        If Me.DirEnvioHabitual >= 0 Then
            txtCod.Text = Me.DirEnvioHabitual
            txtName.Text = DevuelveDesdeBDNew(conAri, "sdirenvio", "nomdiren", "codclien=" & Me.codigo & " AND coddiren", CStr(Me.DirEnvioHabitual), "N")
        Else
            txtCod.Text = ""
            txtName.Text = ""
        End If
        
    End If
End Sub
