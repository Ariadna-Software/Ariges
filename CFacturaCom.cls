VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFacturaCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'COMPRAS
'FACTURA: se corresponde con la tabla ariges.scafpc (cabecera de facturas de Proveedores)
'-------------------------------------------------------------------------

'Variables locales que contienen valores de propiedad
Private mCodProve As String 'Cod. proveedor
Private mNumFactu As String 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mFecRecep As String 'FEcha Recepcion

Private mCodmacta_ As String 'cuenta del proveedor para insertar in spagop
Private mKTipoIVA As Integer  '0, Normal  1, Intra  2. Extranjero  3.  REA  4 Modulos

Private mNomProve As String 'nombre del proveedor
Private mDomProve As String 'domicilio
Private mCodpobla As String 'cpostal
Private mPobprove As String 'poblacion
Private mProprove As String 'provincia
Private mNifProve As String 'NIF proveedor
Private mTelProve As String 'telefono

Private mCodForpa As Integer 'Cod. Forma de pago
Private mCodTraba As String 'Cod. trabajador factura

Private mBrutofac As Currency 'total importe bruto de las lineas de la factura(sin dtos ni IVA)
Private mDtoPpago As Currency 'descuento pronto pago
Private mDtoGnral As Currency 'descuento general
Private mImpPPago As Currency 'Importe al aplicar al bruto el dto pronto pago
Private mImpGnral As Currency  'Importe al aplicar al bruto el dto general
Private mBaseImp As Currency  'Base imponible de la factura (bruto - dtopp - dtogn)

Private mBaseiva1 As Currency  'base imponible tipo IVA 1
Private mBaseiva2 As Currency  'base imponible tipo IVA 2
Private mBaseiva3 As Currency  'base imponible tipo IVA 3
Private mTipoiva1 As Byte   'Cod. tipo de iva 1
Private mTipoiva2 As Byte   'Cod. tipo de iva 2
Private mTipoiva3 As Byte   'Cod. tipo de iva 3
Private mPorciva1 As Currency    '% de iva 1
Private mPorciva2 As Currency    '% de iva 2
Private mPorciva3 As Currency    '% de iva 3

Private mPorcrec1 As Currency   '% de recargo 1
Private mPorcrec2 As Currency   '% de recargo 2
Private mPorcrec3 As Currency   '% de recargo 3

Private mImpoiva1 As Currency   'Importe de iva 1
Private mImpoiva2 As Currency   'Importe de iva 2
Private mImpoiva3 As Currency   'Importe de iva 3
Private mImporec1 As Currency   'Importe de recargo 1
Private mImporec2 As Currency   'Importe de recargo 2
Private mImporec3 As Currency   'Importe de recargo 3

Private mTotalfac As Currency  'total factura (con dtos y con IVA)

Private mBanPr As String 'Banco propio que ponemos como cuenta prevista de pago
Private mCtaPrev As String 'Cuenta prevista de pago


'DAVID###
Private mCCC_Entidad As String
Private mCCC_Oficina As String
Private mCCC_CC As String
Private mCCC_CTa As String
Private mIBAN As String
Private mPais As String


'retencion ----------------------------------------------------
Private mTipoRet As Byte   '0.- Retencion NORMAL
                           '1.- Reteccion coooperativas agricolas
Private mPorRet As Currency
Private mImpRet As Currency


Private mISP As Boolean   'Inversion de sujeto pasivo


'------------------------------------------------
'Propiedades del modulo CFacturaCom
'------------------------------------------------

'**** Proveedor de la Factura

Public Property Let Proveedor(ByVal vData As String)
     mCodProve = vData
End Property

Public Property Get Proveedor() As String
     Proveedor = mCodProve
End Property


'**** Nº de la Factura

Public Property Let Numfactu(ByVal vData As String)
     mNumFactu = vData
End Property

Public Property Get Numfactu() As String
     Numfactu = mNumFactu
End Property


'**** Fecha de la Factura

Public Property Let FecFactu(ByVal vData As String)
     mFecFactu = vData
End Property

Public Property Get FecFactu() As String
     FecFactu = mFecFactu
End Property


'**** Fecha Recepcion de la Factura

Public Property Let FecRecep(ByVal vData As String)
     mFecRecep = vData
End Property

Public Property Get FecRecep() As String
     FecRecep = mFecRecep
End Property


'**** cta contable (codmacta) del Proveedor de la Factura

Public Property Let CtaProve(ByVal vData As String)
     mCodmacta_ = vData
End Property

Public Property Get CtaProve() As String
     CtaProve = mCodmacta_
End Property


'Sera:
Public Property Let TipoIVA(ByVal vData As Integer)
     mKTipoIVA = vData
End Property

Public Property Get TipoIVA() As Integer
     TipoIVA = mKTipoIVA
End Property


'**** Nombre del Proveedor de la Factura

Public Property Let NombreProv(ByVal vData As String)
     mNomProve = vData
End Property

Public Property Get NombreProv() As String
     NombreProv = mNomProve
End Property

'**** Domicilio del Proveedor de la Factura

Public Property Let DomicilioProv(ByVal vData As String)
     mDomProve = vData
End Property

Public Property Get DomicilioProv() As String
     DomicilioProv = mDomProve
End Property


'**** CPostal del Proveedor de la Factura

Public Property Let CPostalProv(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostalProv() As String
     CPostalProv = mCodpobla
End Property


'**** Poblacion del Proveedor de la Factura

Public Property Let PoblacionProv(ByVal vData As String)
     mPobprove = vData
End Property

Public Property Get PoblacionProv() As String
     PoblacionProv = mPobprove
End Property

'**** Provincia del Proveedor de la Factura

Public Property Let ProvinciaProv(ByVal vData As String)
     mProprove = vData
End Property

Public Property Get ProvinciaProv() As String
     ProvinciaProv = mProprove
End Property


'**** NIF del Proveedor de la Factura

Public Property Let NIFProv(ByVal vData As String)
     mNifProve = vData
End Property

Public Property Get NIFProv() As String
     NIFProv = mNifProve
End Property


'**** Telefono del Proveedor de la Factura

Public Property Let TelefProv(ByVal vData As String)
     mTelProve = vData
End Property

Public Property Get TelefProv() As String
     TelefProv = mTelProve
End Property


'**** Trabajador de la Factura

Public Property Let Trabajador(ByVal vData As String)
     mCodTraba = vData
End Property

Public Property Get Trabajador() As String
     Trabajador = mCodTraba
End Property


'**** Forma Pago de la Factura

Public Property Let ForPago(ByVal vData As Integer)
     mCodForpa = vData
End Property

Public Property Get ForPago() As Integer
     ForPago = mCodForpa
End Property


'**** Importe Bruto Factura

Public Property Let BrutoFac(ByVal vData As Currency)
     mBrutofac = vData
End Property

Public Property Get BrutoFac() As Currency
     BrutoFac = mBrutofac
End Property


'**** Descuento pronto pago de la Factura

Public Property Let DtoPPago(ByVal vData As Currency)
     mDtoPpago = vData
End Property

Public Property Get DtoPPago() As Currency
     DtoPPago = mDtoPpago
End Property

'**** Descuento general de la Factura

Public Property Let DtoGnral(ByVal vData As Currency)
     mDtoGnral = vData
End Property

Public Property Get DtoGnral() As Currency
     DtoGnral = mDtoGnral
End Property

'**** Importe pronto pago de la Factura

Public Property Let ImpPPago(ByVal vData As Currency)
     mImpPPago = vData
End Property

Public Property Get ImpPPago() As Currency
     ImpPPago = mImpPPago
End Property


'**** Importe general de la Factura

Public Property Let ImpGnral(ByVal vData As Currency)
     mImpGnral = vData
End Property

Public Property Get ImpGnral() As Currency
     ImpGnral = mImpGnral
End Property

'**** Importe Base Imponible Factura

Public Property Let BaseImp(ByVal vData As Currency)
     mBaseImp = vData
End Property

Public Property Get BaseImp() As Currency
     BaseImp = mBaseImp
End Property

'**** Base IVA1 de la Factura

Public Property Let BaseIVA1(ByVal vData As Currency)
     mBaseiva1 = vData
End Property

Public Property Get BaseIVA1() As Currency
     BaseIVA1 = mBaseiva1
End Property

'**** Base IVA2 de la Factura

Public Property Let BaseIVA2(ByVal vData As Currency)
     mBaseiva2 = vData
End Property

Public Property Get BaseIVA2() As Currency
     BaseIVA2 = mBaseiva2
End Property

'**** Base IVA3 de la Factura

Public Property Let BaseIVA3(ByVal vData As Currency)
     mBaseiva3 = vData
End Property

Public Property Get BaseIVA3() As Currency
     BaseIVA3 = mBaseiva3
End Property

'**** Tipo iva 1 de la Factura

Public Property Let TipoIVA1(ByVal vData As Byte)
     mTipoiva1 = vData
End Property

Public Property Get TipoIVA1() As Byte
     TipoIVA1 = mTipoiva1
End Property

'**** Tipo iva 2 de la Factura

Public Property Let TipoIVA2(ByVal vData As Byte)
     mTipoiva2 = vData
End Property

Public Property Get TipoIVA2() As Byte
     TipoIVA2 = mTipoiva2
End Property

'**** Tipo iva 3 de la Factura

Public Property Let TipoIVA3(ByVal vData As Byte)
     mTipoiva3 = vData
End Property

Public Property Get TipoIVA3() As Byte
     TipoIVA3 = mTipoiva3
End Property

'**** % de iva 1 de la Factura

Public Property Let PorceIVA1(ByVal vData As Currency)
     mPorciva1 = vData
End Property

Public Property Get PorceIVA1() As Currency
     PorceIVA1 = mPorciva1
End Property

'**** % de iva 2 de la Factura

Public Property Let PorceIVA2(ByVal vData As Currency)
     mPorciva2 = vData
End Property

Public Property Get PorceIVA2() As Currency
     PorceIVA2 = mPorciva2
End Property

'**** % de iva 3 de la Factura

Public Property Let PorceIVA3(ByVal vData As Currency)
     mPorciva3 = vData
End Property

Public Property Get PorceIVA3() As Currency
     PorceIVA3 = mPorciva3
End Property

'**** Importe de IVA 1 de la Factura

Public Property Let ImpIVA1(ByVal vData As Currency)
     mImpoiva1 = vData
End Property

Public Property Get ImpIVA1() As Currency
     ImpIVA1 = mImpoiva1
End Property

'**** Importe de IVA 2 de la Factura

Public Property Let ImpIVA2(ByVal vData As Currency)
     mImpoiva2 = vData
End Property

Public Property Get ImpIVA2() As Currency
     ImpIVA2 = mImpoiva2
End Property

'**** Importe de IVA 3 de la Factura

Public Property Let ImpIVA3(ByVal vData As Currency)
     mImpoiva3 = vData
End Property

Public Property Get ImpIVA3() As Currency
     ImpIVA3 = mImpoiva3
End Property

'**** Total Factura

Public Property Let TotalFac(ByVal vData As Currency)
     mTotalfac = vData
End Property

Public Property Get TotalFac() As Currency
     TotalFac = mTotalfac
End Property


'**** Banco propio de pago de la Factura

Public Property Let BancoPr(ByVal vData As String)
     mBanPr = vData
End Property

Public Property Get BancoPr() As String
     BancoPr = mBanPr
End Property

'**** Cuenta Prevista de pago de la Factura
Public Property Let CuentaPrev(ByVal vData As String)
     mCtaPrev = vData
End Property
Public Property Get CuentaPrev() As String
     CuentaPrev = mCtaPrev
End Property


'**** Cuenta banacaria: ENTIDAD
Public Property Let CCC_Entidad(ByVal vData As String)
     mCCC_Entidad = vData
End Property
Public Property Get CCC_Entidad() As String
     CCC_Entidad = mCCC_Entidad
End Property

'**** Cuenta banacaria: OFICINA
Public Property Let CCC_Oficina(ByVal vData As String)
     mCCC_Oficina = vData
End Property
Public Property Get CCC_Oficina() As String
     CCC_Oficina = mCCC_Oficina
End Property

'**** Cuenta banacaria: Codigo control
Public Property Let CCC_CC(ByVal vData As String)
     mCCC_CC = vData
End Property
Public Property Get CCC_CC() As String
     CCC_CC = mCCC_CC
End Property

'Cuenta banacaria.  Cuenta
Public Property Let CCC_CTa(ByVal vData As String)
     mCCC_CTa = vData
End Property
Public Property Get CCC_CTa() As String
     CCC_CTa = mCCC_CTa
End Property

Public Property Let Iban(ByVal vData As String)
     mIBAN = vData
End Property
Public Property Get Iban() As String
     Iban = mIBAN
End Property


'Porcentaje y total retencion PARA liquidacion proveedores
'es decir, socios tipo cooperativas agrarias
Public Property Let PorRet(ByVal vData As Currency)
     mPorRet = vData
End Property

Public Property Get PorRet() As Currency
     PorRet = mPorRet
End Property

Public Property Let ImpRet2(ByVal vData As Currency)
     mImpRet = vData
End Property

Public Property Get ImpRet2() As Currency
     ImpRet2 = mImpRet
End Property


Public Property Let TipoRet(ByVal vData As Byte)
     mTipoRet = vData
End Property

Public Property Get TipoRet() As Byte
     TipoRet = mTipoRet
End Property

Public Property Let ISP(ByVal vData As Boolean)
     mISP = vData
End Property

Public Property Get ISP() As Boolean
     ISP = mISP
End Property



'------------------------------------------------
'Procedimientos del modulo CFacturaCom
'------------------------------------------------

Public Function LeerDatos3(vCodProve As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla scafpc
'Lee de la BD: Ariges, Tabla: scafpc
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim Mens As String
Dim SQL As String

    On Error GoTo ELeer
    
    LeerDatos3 = False
    SQL = "SELECT codprove,numfactu,fecfactu,codforpa,brutofac, totalfac FROM scafpc "
    SQL = SQL & " WHERE codprove=" & vCodProve
    SQL = SQL & " AND numfactu='" & vNumfactu & "'"
    SQL = SQL & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If RS.EOF Then
        
    Else
        Mens = "Datos Factura"
        mCodProve = RS!Codprove
        mNumFactu = CStr(RS!Numfactu)
        mFecFactu = CStr(RS!FecFactu)
        mCodForpa = CStr(RS!codforpa)
        mBrutofac = CStr(RS!BrutoFac)
        mTotalfac = CStr(RS!TotalFac)
        'obterner la ctaCble del proveedor
        SQL = "tipprove"
        CtaProve = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", vCodProve, "N", SQL)
        Me.TipoIVA = CInt(SQL)
        SQL = ""
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de pago que
        'se introdujo en el formulario de pasar albaranes a factura
        'si no venimos de moficicar y la leemos de spagos antes de actualizar
        If CuentaPrev = "" Then
            'leer la cuenta prevista de pagos de la tabla de contabilidad: conta.spagos
            'cuando vamos a borrar una factura porque modificarmos
            If vParamAplic.ContabilidadNueva Then
                SQL = DevuelveDesdeBDNew(conConta, "pagos", "ctabanc1", "codmacta", CtaProve, "T", , "numfactu", vNumfactu, "T", "fecfactu", vFecfactu, "F")
            Else
                SQL = DevuelveDesdeBDNew(conConta, "spagop", "ctabanc1", "ctaprove", CtaProve, "T", , "numfactu", vNumfactu, "T", "fecfactu", vFecfactu, "F")
            End If
            If SQL <> "" Then
                CuentaPrev = SQL
                LeerDatos3 = True
            Else
                Mens = "La cuenta prevista de cobro para la factura no puede ser nula. "
                If vUsu.Nivel <= 1 Then
                    Mens = Mens & vbCrLf & "¿Continuar?"
                    If MsgBox(Mens, vbQuestion + vbYesNoCancel) = vbYes Then
                        CuentaPrev = "''"
                        LeerDatos3 = True
                    End If
                Else
                    MsgBox Mens, vbExclamation
                End If
            End If
        Else
            LeerDatos3 = True
        End If
        

    End If
    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    If Not RS Is Nothing Then Set RS = Nothing
    LeerDatos3 = False
End Function

Public Sub FijarTipoIvaProveedor(ByRef vCodProve As Long)
        '0, Normal  1, Intra  2. Extranjero  3.  REA  4 Modulos
        mKTipoIVA = CInt(DevuelveDesdeBDNew(conAri, "sprove", "tipprove", "codprove", CStr(vCodProve), "N"))
        
End Sub

'======================================================================
'GRABAR EN TESORERIA
'======================================================================

Public Function InsertarEnTesoreria(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim SQL As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim I As Byte

    On Error GoTo EInsertarTesoreria

'    b = False
    InsertarEnTesoreria = False
    CadValues2 = ""

    'Obtener el Nº de Vencimientos de la forma de pago
    SQL = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago de la tabla de parametros: spara1
            SQL = " SELECT  diapago1, diapago2, diapago3,mesnogir "
            SQL = SQL & " FROM spara1 "
            'SQL = SQL & " WHERE codigo=1"
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RS.EOF Then
               'vamos creando la cadena para insertar en spagosp de la CONTA
               CadValuesAux2 = "('" & CtaProve & "', " & DBSet(Numfactu, "T") & ", '" & Format(FecFactu, FormatoFecha) & "', "
              
              'Primer Vencimiento
              '------------------------------------------------------------
              I = 1
              'FECHA VTO
              FecVenci = CDate(FecFactu)
              '=== Modificado: Laura 23/01/2007
'              FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
              FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
              '==================================
              'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
              FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))

              'Comprobar si  tiene mes a no girar
              FecVenci1 = FecVenci
              If DBSet(RS!mesnogir, "N") <> 0 Then
                  FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(RS!mesnogir, "N"), DBSet(0, "N"), RS!DiaPago1, RS!DiaPago2, RS!DiaPago3)
              End If
             
              CadValues2 = CadValuesAux2 & I
              CadValues2 = CadValues2 & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
              'IMPORTE del Vencimiento
              If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
              Else
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
                        ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
                    End If
              End If
              CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CuentaPrev, "T") & ","
              
              If Val(mCCC_Entidad) = 0 Then
                'NO TIENE ENTIDAD
                If vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & "null,"   'Solo llevamos el IBAN
                Else
                    CadValues2 = CadValues2 & "null,null,null,null,null,"
                End If
              Else
                
                
               
                If vParamAplic.ContabilidadNueva Then
                    'En contabiliadd nueva, el IBAN no solo es ES??   es toooooda la cuenta
                    SQL = mIBAN & Right("0000" & mCCC_Entidad, 4) & Right("0000" & mCCC_Oficina, 4)
                    SQL = SQL & mCCC_CC & Right(String(10, "0") & mCCC_CTa, 10)
                    CadValues2 = CadValues2 & DBSet(SQL, "T") & ","
                Else
                     CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                    CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
                    CadValues2 = CadValues2 & DBSet(mIBAN, "T", "S") & ","
                End If
              End If

              'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
              SQL = "Factura num.: " & Numfactu & "-" & Format(FecFactu, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "',"
              SQL = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "'"
 

              'Si hay contabilidad nueva añadira los siguiens valores
              '  nomprove , domprove, pobprove, cpprove, proprove, nifProve, codpais, numSerie)"
            
              If vParamAplic.ContabilidadNueva Then
                    
                    If mNomProve = "" Then
                        mPais = "ES"
                        'No tiene cargado los datos del prove
                        SQL = "concat(nomprove ,'|', domprove,'|', pobprove,'|', codpobla,'|', proprove,'|', nifProve,'|',codpais,'|')"
                        SQL = DevuelveDesdeBD(conAri, SQL, "sprove", "codprove", CStr(mCodProve))
                        If Len(SQL) > 6 Then
                                mNomProve = RecuperaValor(SQL, 1)
                                mDomProve = RecuperaValor(SQL, 2)
                                mPobprove = RecuperaValor(SQL, 3)
                                mCodpobla = RecuperaValor(SQL, 4)
                                mProprove = RecuperaValor(SQL, 5)
                                mNifProve = RecuperaValor(SQL, 6)
                                mPais = RecuperaValor(SQL, 7)
                        End If
                    Else
                        SQL = DevuelveDesdeBD(conAri, "codpais", "sprove", "codprove", CStr(mCodProve))
                        If SQL = "" Then SQL = "ES"
                        mPais = SQL
                    End If
                    '  nomprove , domprove, pobprove, cpprove, proprove, nifProve, codpais, numSerie)"
                    CadValues2 = CadValues2 & "," & DBSet(mNomProve, "T") & "," & DBSet(mDomProve, "T") & "," & DBSet(mPobprove, "T")
                    CadValues2 = CadValues2 & "," & DBSet(mCodpobla, "T") & "," & DBSet(mProprove, "T") & "," & DBSet(mNifProve, "T")
                    CadValues2 = CadValues2 & "," & DBSet(mPais, "T") & ",1"
                    
                    
                    
                    
              End If
              CadValues2 = CadValues2 & ")"
              
              
              'Resto Vencimientos
              '--------------------------------------------------------------------
              For I = 2 To rsVenci!numerove
                 'FECHA Resto Vencimientos
                  '==== Modificado: Laura 23/01/2007
                  'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                  FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                  '==================================================
                  'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                  FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))

                  'Comprobar si tiene mes a no girar
                  FecVenci1 = FecVenci
                  If DBSet(RS!mesnogir, "N") <> 0 Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(RS!mesnogir, "N"), DBSet(0, "N"), RS!DiaPago1, RS!DiaPago2, RS!DiaPago3)
                  End If

                  CadValues2 = CadValues2 & ", " & CadValuesAux2 & I & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "

                  'IMPORTE Resto de Vendimientos
                  ImpVenci = Round(TotalFac / rsVenci!numerove, 2)

                  CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "',"
                  
                  
                  'David. Para que ponga la cuenta bancaria (SI LA tiene)
                  If vParamAplic.ContabilidadNueva Then
                        If mCCC_Entidad = "" Then
                            SQL = ""
                        Else
                            SQL = mIBAN & Right("0000" & mCCC_Entidad, 4) & Right("0000" & mCCC_Oficina, 4)
                            SQL = SQL & mCCC_CC & Right(String(10, "0") & mCCC_CTa, 10)
                        End If
                        CadValues2 = CadValues2 & DBSet(SQL, "T", "S") & ","
                  Else
                        CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                        CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & "," & DBSet(mIBAN, "T", "S") & ","
                  End If
                  SQL = "Factura num.: " & Numfactu & "-" & Format(FecFactu, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "',"
                  SQL = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(SQL) & "'"
                  
                  
                  
                  If vParamAplic.ContabilidadNueva Then
                  
                        '  nomprove , domprove, pobprove, cpprove, proprove, nifProve, codpais, numSerie)"
                        CadValues2 = CadValues2 & "," & DBSet(mNomProve, "T") & "," & DBSet(mDomProve, "T") & "," & DBSet(mPobprove, "T")
                        CadValues2 = CadValues2 & "," & DBSet(mCodpobla, "T") & "," & DBSet(mProprove, "T") & "," & DBSet(mNifProve, "T")
                        CadValues2 = CadValues2 & "," & DBSet(mPais, "T") & ",1"
                    
                  End If
                  CadValues2 = CadValues2 & ")"
                  
              Next I
            End If
        End If
        RS.Close
        Set RS = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    'Grabar tabla spagop de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        'antes de grabar en la spago comprobar que existe en conta.sforpa la
        'forma de pago de la factura. Sino existe insertarla

        'vemos si existe en la conta
        If vParamAplic.ContabilidadNueva Then
            CadValuesAux2 = DevuelveDesdeBDNew(conConta, "formapago", "codforpa", "codforpa", ForPago, "N")
        Else
            CadValuesAux2 = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
        End If
        'si no existe la forma de pago en conta, insertamos la de ariges
        If CadValuesAux2 = "" Then
            If vParamAplic.ContabilidadNueva Then
                
            
                cadValuesAux = "codforpa,nomforpa,tipforpa,numerove,primerve,restoven,"
                CadValuesAux2 = "INSERT INTO formapago(" & cadValuesAux & "iban) SELECT "
                CadValuesAux2 = CadValuesAux2 & cadValuesAux & "obsiban FROM " & vUsu.CadenaConexion
                SQL = CadValuesAux2 & ".sforpa WHERE codforpa = " & ForPago
            Else
                cadValuesAux = "tipforpa"
                CadValuesAux2 = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", ForPago, "N", cadValuesAux)
                'insertamos e sforpa de la CONTA
                
                SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                    
                SQL = SQL & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
            End If
            ConnConta.Execute SQL
        End If

        'Insertamos en la tabla spagop de la CONTA
        'SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1) "
        'David. Cuenta bancaria y descripcion textos
        If vParamAplic.ContabilidadNueva Then
            SQL = "INSERT INTO pagos (codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1," 'entidad,oficina,control,cuentaba,
            SQL = SQL & "iban,text1csb,text2csb "
            SQL = SQL & ", nomprove , domprove, pobprove, cpprove, proprove, nifProve, codpais, numSerie)"
        Else
            SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,iban,text1csb,text2csb )"
        End If
        
        

        SQL = SQL & " VALUES " & CadValues2
        ConnConta.Execute SQL
    End If

    b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreria = b
End Function

'Si es LiquidacionProveedores entonces
' es modificaion Terrasana. para cuando generamos las facturas de proveedores a nuestros socios
' con tipo iva REA
' La diferencia PPAL es en el Arimoney y que en la lineas de la factura , en la conta
' es el lugar donde van las retenciones
Public Function TraspasoAlbaranesAFactura(cadSQL As String, InsertarEnLaTesoreria As Boolean, MarcarAlbaranesComoContabilizados As Boolean, LiquidacionProveedores As Boolean) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String
Dim ImporteAuxiliarTesoreria As Currency

    On Error GoTo ETraspasoAlbFac
    TraspasoAlbaranesAFactura = False
    
    conn.BeginTrans
    ConnConta.BeginTrans
    

    
    cadSQL = Replace(cadSQL, "slialp", "scaalp")
    
    'Cuenta Prevista de Cobro de las Facturas
    CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", BancoPr, "N")
    
    
    'Insertar la Factura
    b = InsertarFactura(cadSQL, MenError)
    
    If b Then
        'Insertar en Tesoreria
        
        If InsertarEnLaTesoreria Then
            
            b = InsertarEnTesoreria(MenError)
            
        End If
        
        
        If MarcarAlbaranesComoContabilizados Then
            If b Then b = Me.MarcarContabilizada(MenError)
        End If
        
        'Eliminar los Albaranes y sus lineas
        If b Then
            b = EliminarAlbaranes(cadSQL, MenError)
        End If
    End If

     If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        TraspasoAlbaranesAFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        TraspasoAlbaranesAFactura = False
        MsgBox "ATENCIÓN:" & vbCrLf & "La Factura NO se generó correctamente!!!." & vbCrLf & MenError, vbInformation
    End If

    Exit Function


ETraspasoAlbFac:
    If Err.Number <> 0 Then
'        b = False
        conn.RollbackTrans
        ConnConta.RollbackTrans
        TraspasoAlbaranesAFactura = False
        MuestraError Err.Number, "Pasando Albaranes a Factura.(Recepción)", Err.Description
    End If

End Function


Private Function InsertarFactura(cadSQL As String, cadError As String) As Boolean
'Insertamos en las tablas de factura de proveedor: scafpc, scafpa, slifpc
Dim b As Boolean
    
    InsertarFactura = False
    
    
    'Cabecera Factura
    cadError = ""
    b = InsertarCabeceraFactu(cadError)
    If Not b Then
'        cadError = "Error al insertar cabecera de factura"
        Exit Function
    End If
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL, cadError)
    If Not b Then
'        cadError = "Error al insertar cabecera albaranes de factura"
        Exit Function
    End If
    
    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL, cadError)
    If Not b Then
'        cadError = "Error al insertar líneas de factura"
        Exit Function
    End If
    
    InsertarFactura = True
End Function



Private Function InsertarCabeceraFactu(Optional cadErr As String) As Boolean
'Inserta la cabecera de la factura en la tabla: scafpc
Dim SQL As String
Dim vProve As CProveedor
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo EInsertar

    'Obtener los datos del proveedor: nombre, domicilio, NIF,....
    Set vProve = New CProveedor
    If vProve.Existe(Proveedor) Then
        vProve.LeerDatos (Proveedor)
        '#### Leer estos datos de la tabla scaalpr y no de sprove
        NombreProv = vProve.Nombre
        DomicilioProv = vProve.Domicilio
        CPostalProv = vProve.CPostal
        PoblacionProv = vProve.Poblacion
        ProvinciaProv = vProve.Provincia
        NIFProv = vProve.NIF
        TelefProv = vProve.TfnoAdmon
        CtaProve = vProve.CuentaCble
    End If

    'Insertar en la tabla cabecera de la factura de compras
    SQL = "INSERT INTO scafpc (codprove,numfactu,fecfactu,fecrecep,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,codforpa,codtraba,brutofac,dtoppago,dtognral,impppago,impgnral,baseiva1,baseiva2,baseiva3, "
    SQL = SQL & "tipoiva1,tipoiva2,tipoiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,totalfac,intconta"
    'RETENCION SOCIOS, y no socios: segun tiporet
    SQL = SQL & ",PorRet,ImpRet,tiporet,InvSujPas"
    SQL = SQL & ") VALUES (" & Proveedor & "," & DBSet(Numfactu, "T") & "," & DBSet(FecFactu, "F") & "," & DBSet(FecRecep, "F") & ","
    SQL = SQL & DBSet(NombreProv, "T", "N") & "," & DBSet(DomicilioProv, "T", "N") & "," & DBSet(CPostalProv, "T", "N") & "," & DBSet(PoblacionProv, "T", "N") & "," & DBSet(ProvinciaProv, "T", "N") & ","
    SQL = SQL & DBSet(NIFProv, "T", "N") & "," & DBSet(TelefProv, "T", "N") & "," & ForPago & ","
    SQL = SQL & Trabajador & "," & DBSet(BrutoFac, "N", "N") & ","
    SQL = SQL & DBSet(DtoPPago, "N", "N") & "," & DBSet(DtoGnral, "N", "N") & ","
    SQL = SQL & DBSet(ImpPPago, "N", "N") & "," & DBSet(ImpGnral, "N", "N") & ","
    SQL = SQL & DBSet(BaseIVA1, "N", "N") & "," & DBSet(BaseIVA2, "N", "S") & "," & DBSet(BaseIVA3, "N", "S") & ","
    Nulo2 = "N"
    Nulo3 = "N"
    If BaseIVA2 = 0 Then Nulo2 = "S"
    If BaseIVA3 = 0 Then Nulo3 = "S"
    SQL = SQL & DBSet(TipoIVA1, "N", "N") & "," & DBSet(TipoIVA2, "N", Nulo2) & "," & DBSet(TipoIVA3, "N", Nulo3) & ","
    SQL = SQL & DBSet(PorceIVA1, "N", "N") & "," & DBSet(PorceIVA2, "N", Nulo2) & "," & DBSet(PorceIVA3, "N", Nulo3) & ","
    SQL = SQL & DBSet(ImpIVA1, "N", "N") & "," & DBSet(ImpIVA2, "N", Nulo2) & "," & DBSet(ImpIVA3, "N", Nulo3) & ","
    SQL = SQL & DBSet(TotalFac, "N", "N") & ",0,"
    'Retencion socios
    SQL = SQL & DBSet(PorRet, "N", "S") & "," & DBSet(ImpRet2, "N", "S")
    SQL = SQL & "," & TipoRet & "," & DBSet(mISP, "B") & ")"
    
    conn.Execute SQL
    Set vProve = Nothing
    InsertarCabeceraFactu = True

EInsertar:
    If Err.Number <> 0 Then
        Set vProve = Nothing
        InsertarCabeceraFactu = False
        cadErr = "Error al insertar cabecera factura: " & Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function



Private Function InsertarCabAlbaranesFactu(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: scafpa
Dim SQL As String

    On Error GoTo EInsertarAlb

    SQL = "INSERT INTO scafpa (codprove,numfactu,fecfactu,numalbar,fechaalb,numpedpr,fecpedpr,codtrab1,codtrab2,observa1,observa2,observa3,observa4,observa5,fecenvio,docarchiv,codenvio,NReferencia,SReferencia,fecentrega) "
    SQL = SQL & " SELECT " & Proveedor & " as codprove," & DBSet(Numfactu, "T") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & " numalbar,fechaalb,numpedpr,fecpedpr,codtrab1,codtraba,observa1,observa2,observa3,observa4,observa5,fecenvio,docarchiv,codenvio,NReferencia,SReferencia,fecentrega "
    SQL = SQL & " FROM scaalp WHERE " & cadSQL
    conn.Execute SQL

    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        cadErr = "Error al insertar cabecera albaranes Factura: " & Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
End Function


Private Function InsertarLineasFactu(cadSQL As String, cadErr As String) As Boolean
'Insertar las lineas de los Albaranes de la factura en la tabla: slifpc
Dim SQL As String

    On Error GoTo EInsertarLin
    'Julio 2015 EULER
    SQL = "INSERT INTO slifpc (codprove,numfactu,fecfactu,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,"
    SQL = SQL & "dtoline1,dtoline2,importel,numlotes,codccost,codclien,coddirec,actuacion ,codtipomV,numalbarV,fechaalbV,numpedV)"
    SQL = SQL & " SELECT " & Proveedor & " as codprove," & DBSet(Numfactu, "T") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu,"
    SQL = SQL & " numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,numlotes,codccost,"
    SQL = SQL & " codclien,coddirec,actuacion,codtipomV,numalbarV,fechaalbV,numpedV "
    SQL = SQL & " FROM slialp WHERE " & Replace(cadSQL, "scaalp", "slialp")
    conn.Execute SQL

    InsertarLineasFactu = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactu = False
        cadErr = "Error al insertar líneas de factura: " & Err.Description
    Else
        InsertarLineasFactu = True
    End If
End Function


Private Function EliminarAlbaranes(cadSQL As String, cadError As String) As Boolean
'Eliminamos de las tablas de Albaranes: scaalp, slialp
Dim SQL As String

    On Error GoTo EEliminar

    EliminarAlbaranes = False
    
    'ELiminar lineas
    SQL = "DELETE FROM slialp "
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaalp", "slialp")
    conn.Execute SQL
    
    'Eliminar Cabeceras Albaranes
    SQL = "DELETE FROM scaalp "
    SQL = SQL & " WHERE " & cadSQL
    conn.Execute SQL
        
    EliminarAlbaranes = True
    Exit Function

EEliminar:
    EliminarAlbaranes = False
    cadError = "Error al Eliminar albaranes: " & Err.Description
End Function



Public Function CalcularDatosFactura2(cadWhere As String, NomTabla As String, NomTablaLin As String, FechaDoc As Date, InvSujetoPasivo As Boolean) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim RS As ADODB.Recordset
Dim I As Integer

Dim SQL As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim Impaux As Currency
Dim ImpIva As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA
Dim CodigoDeIva As Integer

    CalcularDatosFactura2 = False
    On Error GoTo ECalcular

   
    cadWhere = Replace(cadWhere, NomTabla, NomTablaLin)
    
    If mKTipoIVA = 1 Then
        
        SQL = "SELECT " & vParamAplic.IVA_Intracomunitario & " CodigIVA ,sum(importel) as bruto"
    Else
        SQL = "SELECT " & "sartic.codigiva,sum(importel) as bruto"
    End If
    
    SQL = SQL & " FROM " & NomTablaLin & " LEFT OUTER JOIN sartic ON " & NomTablaLin & ".codartic=sartic.codartic"
    SQL = SQL & " WHERE " & cadWhere
    SQL = SQL & " GROUP BY codigiva "
    SQL = SQL & " ORDER BY codigiva "

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    TotBruto = 0
    I = 1
    While Not RS.EOF
        'Aqui vamos acumulando la suma del importe bruto de las lineas
        TotBruto = TotBruto + RS!bruto
        ImpBImIVA = RS!bruto
        
        'Aplicarle el dto ppago
        '---- Laura: 24/10/2006
        Impaux = CalcularPorcentaje(RS!bruto, DtoPPago, 2)
        ImpBImIVA = ImpBImIVA - Impaux '(bruto - DtoPP)
        
        'Aplicarle el dto grnal
        '---- Laura: 24/10/2006
        Impaux = CalcularPorcentaje(RS!bruto, DtoGnral, 2)
        ImpBImIVA = ImpBImIVA - Impaux '(bruto - Dtogn)
        

        
        
        If FechaDoc < vParamAplic.FechaCambioIva Then
            cadAux = CStr(vParamAplic.DevuelveTipoIVA_OLD(False, RS!Codigiva, True))
            If cadAux = "" Then cadAux = RS!Codigiva
        Else
             cadAux = RS!Codigiva
        End If
        CodigoDeIva = CInt(cadAux)
        cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(CodigoDeIva), "N")
        If cadAux = "" Then cadAux = "0"
        ImpIva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + ImpIva
        
        
        Select Case I
            Case 1  'IVA 1
                TipoIVA1 = CodigoDeIva
                
                BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA1 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA1 = ImpIva
                
            Case 2  'IVA 2
                TipoIVA2 = CodigoDeIva
                
                BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA2 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA2 = ImpIva

            Case 3  'IVA 3
                TipoIVA3 = CodigoDeIva
                
                BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA3 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA3 = ImpIva
        End Select
        I = I + 1
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing

    'TOTALES
    BrutoFac = TotBruto
    
    'Aplicarle el dto ppago
'    ImpPPago = CCur(CalcularDto(CStr(TotBruto), CStr(DtoPPago)))
'    '---- Laura: 27/09/2006
'    ImpPPago = Round(ImpPPago, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
    '----
    
    'Aplicarle el dto general
'    ImpGnral = CCur(CalcularDto(CStr(TotBruto), CStr(DtoGnral)))
'    '---- Laura: 27/09/2006
'    ImpGnral = Round(ImpGnral, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
    '----
    
    'Base Imponible
    BaseImp = TotBruto - ImpPPago - ImpGnral
    
    'TOTAL de la factura
    'Si es inversion de sujeto pasivo el IVA no lo suma en el total factura
    TotalFac = BaseImp
    If Not mISP Then
        'No es inversion SUJ pasivo, si ademas, no es intracom NO sumo eliva
        If mKTipoIVA <> 1 Then TotalFac = BaseImp + TotImpIVA
    End If
    If PorRet > 0 And ImpRet2 <> 0 Then TotalFac = TotalFac - ImpRet2
    
    
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFactura2 = False
    Else
        CalcularDatosFactura2 = True
    End If
End Function




'#### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function MarcarContabilizada(Optional cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    MarcarContabilizada = False
    
    SQL = "UPDATE scafpc SET intconta=1 "
    SQL = SQL & " WHERE codprove=" & Me.Proveedor & " AND numfactu=" & DBSet(Me.Numfactu, "T") & " AND fecfactu=" & DBSet(Me.FecFactu, "F")
    conn.Execute SQL
    
    MarcarContabilizada = True
    Exit Function
    
EActualizar:
    MarcarContabilizada = False
    cadErr = "Error al marcar como contabilizada:" & Err.Description
End Function

