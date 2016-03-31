VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'VENTAS
'FACTURA: se corresponde con la tabla ariges.scafac (cabecera de facturas)
'-------------------------------------------------------------------------

'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodTipoM As String 'Tipo de Movimiento="FAV", "FAR" (Factura venta o reparacion)
Private mNumFactu As Long 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mCodClien As String 'Cod. cliente

Private mNomClien As String 'nombre del cliente
Private mDomClien As String 'domicilio
Private mCodpobla As String 'cpostal
Private mPobClien As String 'poblacion
Private mProClien As String 'provincia
Private mNIFClien As String 'NIF proveedor
Private mTelclien As String 'telefono

Private mCoddirec As String 'codigo direc./dpto
Private mNomdirec As String  'nombre direc./dpto
Private mCodAgent As Integer 'codigo de agente


Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria
Private mIBAN As String 'IBAN

Private mCodForpa As String 'Cod. Forma de pago
Private mTipForpa As Integer 'tipo de forma de pago: efectivo,...

Private mDtoPpago As Currency 'descuento pronto pago
Private mDtoGnral As Currency 'descuento general
Private mBrutofac As Currency 'importe bruto de la factura
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

Private mImpoiva1 As Currency   'Importe de iva 1
Private mImpoiva2 As Currency   'Importe de iva 2
Private mImpoiva3 As Currency   'Importe de iva 3


'DAVID.   ENERO 2008
Private mPorciva1RE As Currency    '% de iva 1
Private mPorciva2RE As Currency    '% de iva 2
Private mPorciva3RE As Currency    '% de iva 3

Private mImpoiva1RE As Currency   'Importe de iva 1
Private mImpoiva2RE As Currency   'Importe de iva 2
Private mImpoiva3RE As Currency   'Importe de iva 3





Private mTotalfac2 As String 'total factura

Private mLetraSer As String 'Letra de serie para los cobros    reutlizada en renting
Private mBanPr As String 'Banco propio que ponemos como cuenta prevista de cobro
Private mCtaPrev As String 'Cuenta prevista de cobro

Private mNumtermi As Integer 'terminal de venta del TPV
Private mNumventa As Long 'Nº de venta del TPV



'David.
'Para las facturas de mantenimeinto, hay un campo que es "concepto factura" varchar(60)
'Si tiene valor entonces ese valor es el que grabaremos en la linea y como obervacion de la factura
'meteremos lo de antes ("texto" de: mes de año.  Ejemplo: manteimiento de Junio de 2007
'Si no tiene valor seguira como estaba
Private mObservacion As String

' VBLES para los Mantenimientos -------------------------------------------
Dim TipCoMan As String 'tipo de contrato para Facturas de Manteniminetos
Dim OpeFactu As String 'Operador conectado o el que pasamos como parametro
Dim AlmFactu As String 'Cod Almacen del trabajador conectado, para insertar lines mantenim, desplaz, bonificaciones,....
Dim MesFactu As String
Dim nummante As String
Dim ArticFac As String 'cod articulo que se va a utilizar para facturar mantenimientos
Dim manCCost As String 'Centro de coste  para las lineas
'--------------------------------------------------------------------------

' VBLES para los Desplazamientos ----------------------------------
Dim TotalKm As Integer
Dim PrecioKm As Currency
Dim ImporteL As Currency
Dim Ampliacion As String  'ampliacion de la linea: "Dia,Dia,Dia..." de los albaranes con desplazamiento
Dim codArtic As String
Dim despIVA As Integer 'tipo de IVA del articulo para desplazamientos
'------------------------------------------------------------------

' VBLES para las Bonificaciones -----------------------------------
Dim TotBonif1 As Currency 'Total de bonificacion a factura con articulo de tipo de IVA1
Dim IVABonif1 As Byte 'tipo de IVA
Dim TotBonif2 As Currency
Dim IVABonif2 As Byte
Dim TotBonif3 As Currency
Dim IVABonif3 As Byte
Dim SQLBonif As String
'------------------------------------------------------------------

' VBLES para los Cheques Regalo -----------------------------------
Dim ImpCheque As Currency
'------------------------------------------------------------------



'vbles para las facturas que vienen de una venta de ticket ---------
Private NumTicket As String
Private NumAlbTicket As String '01/09/06 laura   'reutlizada en renting
'-------------------------------------------------------------------

Private mAportacion As Currency   'Viene de Radival, pero se podria utilizar en mas
                                 ' Es la aportacion que hace Telefonica, en los moviles
                                 
                                 
                                 
Private NuevasNormasBancarias As Boolean
                                 
                                 
                                 
Dim FechaFacturarRenting As Date
                                 


'------------------------------------------------
'Propiedades del modulo CFactura
'------------------------------------------------

'**** Tipo de Movimiento de la Factura

Public Property Let codtipom(ByVal vData As String)
     mCodTipoM = vData
End Property

Public Property Get codtipom() As String
     codtipom = mCodTipoM
End Property


'**** Nº de la Factura

Public Property Let NumFactu(ByVal vData As Long)
     mNumFactu = vData
End Property

Public Property Get NumFactu() As Long
     NumFactu = mNumFactu
End Property


'**** Fecha de la Factura

Public Property Let FecFactu(ByVal vData As String)
     mFecFactu = vData
End Property

Public Property Get FecFactu() As String
     FecFactu = mFecFactu
End Property


'**** Cliente de la Factura

Public Property Let Cliente(ByVal vData As String)
     mCodClien = vData
End Property

Public Property Get Cliente() As String
     Cliente = mCodClien
End Property


'**** Nombre del Cliente de la Factura

Public Property Let NombreClien(ByVal vData As String)
     mNomClien = vData
End Property

Public Property Get NombreClien() As String
     NombreClien = mNomClien
End Property


'**** Domicilio del Cliente de la Factura

Public Property Let DomicilioClien(ByVal vData As String)
     mDomClien = vData
End Property

Public Property Get DomicilioClien() As String
     DomicilioClien = mDomClien
End Property


'**** CPostal del Cliente de la Factura

Public Property Let CPostal(ByVal vData As String)
     mCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodpobla
End Property


'**** Poblacion del Cliente de la Factura

Public Property Let Poblacion(ByVal vData As String)
     mPobClien = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobClien
End Property


'**** Provincia del Cliente de la Factura

Public Property Let Provincia(ByVal vData As String)
     mProClien = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProClien
End Property


'**** NIF del Cliente de la Factura

Public Property Let NIF(ByVal vData As String)
     mNIFClien = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFClien
End Property


'**** Telefono del Cliente de la Factura

Public Property Let Telefono(ByVal vData As String)
     mTelclien = vData
End Property

Public Property Get Telefono() As String
     Telefono = mTelclien
End Property


'**** Direc./Dpto de Cliente de la Factura

Public Property Let DirDpto(ByVal vData As String)
     mCoddirec = vData
End Property

Public Property Get DirDpto() As String
     DirDpto = mCoddirec
End Property


'**** Nombre Direc./Dpto de Cliente de la Factura

Public Property Let NombreDirDpto(ByVal vData As String)
     mNomdirec = vData
End Property

Public Property Get NombreDirDpto() As String
     NombreDirDpto = mNomdirec
End Property


'**** Agente de Cliente de la Factura

Public Property Let Agente(ByVal vData As Integer)
     mCodAgent = vData
End Property

Public Property Get Agente() As Integer
     Agente = mCodAgent
End Property


'**** Forma Pago de la Factura

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Tipo de Forma Pago de la Factura
'(para tickets saber si insertar en tesoreria o no: si es efectivo no se inserta en scobro)

Public Property Let TipForPago(ByVal vData As Integer)
     mTipForpa = vData
End Property

Public Property Get TipForPago() As Integer
     TipForPago = mTipForpa
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let LetraSerie(ByVal vData As String)
     mLetraSer = vData
End Property

Public Property Get LetraSerie() As String
     LetraSerie = mLetraSer
End Property


'**** Banco propio de pago de la Factura

Public Property Let BancoPr(ByVal vData As String)
     mBanPr = vData
End Property

Public Property Get BancoPr() As String
     BancoPr = mBanPr
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let CuentaPrev(ByVal vData As String)
     mCtaPrev = vData
End Property

Public Property Get CuentaPrev() As String
     CuentaPrev = mCtaPrev
End Property


'**** Banco de la Factura

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal de la Factura

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la Factura

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


'**** Importe Bruto Factura

Public Property Let BrutoFac(ByVal vData As Currency)
     mBrutofac = vData
End Property

Public Property Get BrutoFac() As Currency
     BrutoFac = mBrutofac
End Property


'**** Importe pronto pago de la Factura

Public Property Let ImpPPago(ByVal vData As Currency)
     vData = vData
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




'*****************************************************++
'   Recargo de equivalencia

'**** % de iva 1 de la Factura

Public Property Let PorceIVA1RE(ByVal vData As Currency)
     mPorciva1RE = vData
End Property

Public Property Get PorceIVA1RE() As Currency
     PorceIVA1RE = mPorciva1RE
End Property


'**** % de iva 2 de la Factura

Public Property Let PorceIVA2RE(ByVal vData As Currency)
     mPorciva2RE = vData
End Property

Public Property Get PorceIVA2RE() As Currency
     PorceIVA2RE = mPorciva2RE
End Property


'**** % de iva 3 de la Factura

Public Property Let PorceIVA3RE(ByVal vData As Currency)
     mPorciva3RE = vData
End Property

Public Property Get PorceIVA3RE() As Currency
     PorceIVA3RE = mPorciva3RE
End Property


'**** Importe de IVA 1 de la Factura

Public Property Let ImpIVA1RE(ByVal vData As Currency)
     mImpoiva1RE = vData
End Property

Public Property Get ImpIVA1RE() As Currency
     ImpIVA1RE = mImpoiva1RE
End Property


'**** Importe de IVA 2 de la Factura

Public Property Let ImpIVA2RE(ByVal vData As Currency)
     mImpoiva2RE = vData
End Property

Public Property Get ImpIVA2RE() As Currency
     ImpIVA2RE = mImpoiva2RE
End Property


'**** Importe de IVA 3 de la Factura

Public Property Let ImpIVA3RE(ByVal vData As Currency)
     mImpoiva3RE = vData
End Property

Public Property Get ImpIVA3RE() As Currency
     ImpIVA3RE = mImpoiva3RE
End Property












'**** Total Factura

Public Property Let TotalFac(ByVal vData As String)
     mTotalfac2 = vData
End Property

Public Property Get TotalFac() As String
     TotalFac = mTotalfac2
End Property


'**** Nº terminal de la venta del TPV

Public Property Let NumTerminal(ByVal vData As Integer)
     mNumtermi = vData
End Property

Public Property Get NumTerminal() As Integer
     NumTerminal = mNumtermi
End Property



'**** Nº venta del TPV

Public Property Let NumVenta(ByVal vData As Long)
     mNumventa = vData
End Property

Public Property Get NumVenta() As Long
     NumVenta = mNumventa
End Property


'**** Para los MANTENIMIENTOS.  sera 1 la linea. Si esta vacio lo cargaremos con la amplicaicon normal

Public Property Let Observacion(ByVal vData As String)
     mObservacion = vData
End Property

Public Property Get Observacion() As String
     Observacion = mObservacion
End Property


'-----   Aportacion
Public Property Let Aportacion(ByVal vData As Currency)
     mAportacion = vData
End Property

Public Property Get Aportacion() As Currency
     Aportacion = mAportacion
End Property



'------------------------------------------------
'Procedimientos del modulo CFactura
'------------------------------------------------

Public Function LeerDatos(vCodtipom As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla scafac
'Lee de la BD: Ariges, Tabla: scafac
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Mens As String
Dim SQL As String

    On Error GoTo ELeer
    
    
    LeerDatos = False
    SQL = "SELECT codtipom,numfactu,fecfactu,codclien,codforpa,codbanco,codsucur,digcontr,cuentaba,brutofac, totalfac,iban,Aportacion FROM scafac "
    SQL = SQL & " WHERE codtipom='" & vCodtipom & "'"
    SQL = SQL & " AND numfactu=" & vNumfactu
    SQL = SQL & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If Rs.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Factura"
        mCodTipoM = Rs!codtipom
        mNumFactu = CStr(Rs!NumFactu)
        mFecFactu = CStr(Rs!FecFactu)
        mCodClien = CStr(Rs!codClien)
        ForPago = CStr(Rs!codforpa)
        TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N")
        mCodbanco = CStr(DBLet(Rs!codbanco, "N"))
        mCodsucur = CStr(DBLet(Rs!codsucur, "N"))
        mDigContr = DBLet(Rs!digcontr, "T")
        mCuentaba = DBLet(Rs!cuentaba, "T")
        mIBAN = DBLet(Rs!IBAN, "T")
        
        mAportacion = DBLet(Rs!Aportacion, "N")
        
        mBrutofac = CStr(Rs!BrutoFac)
        mTotalfac2 = CStr(Rs!TotalFac)

        mLetraSer = ObtenerLetraSerie(mCodTipoM)
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de cobro que
        'se introdujo en el formulario de pasar albaranes a factura
        If CuentaPrev = "" Then
            'leer la cuenta prevista de cobros de la tabla de contabilidad: conta.scobro
            'cuando vamos a borrar una factura porque modificarmos
            'ya que este dato se pidio al facturar (SI EXISTE LA FACTURA EN TESORERIA)
            If vParamAplic.ContabilidadNueva Then
                SQL = "SELECT COUNT(*) FROM cobros WHERE numserie='" & LetraSerie & "' and numfactu=" & NumFactu
                SQL = SQL & " AND fecfactu=" & DBSet(FecFactu, "F")
            Else
                SQL = "SELECT COUNT(*) FROM scobro WHERE numserie='" & LetraSerie & "' and codfaccl=" & NumFactu
                SQL = SQL & " AND fecfaccl=" & DBSet(FecFactu, "F")
            End If
            If RegistrosAListar(SQL, conConta) > 0 Then
                If vParamAplic.ContabilidadNueva Then
                    SQL = DevuelveDesdeBDNew(conConta, "cobros", "ctabanc1", "numserie", mLetraSer, "T", , "numfactu", vNumfactu, "N", "fecfactu", vFecfactu, "F")
                Else
                    SQL = DevuelveDesdeBDNew(conConta, "scobro", "ctabanc1", "numserie", mLetraSer, "T", , "codfaccl", vNumfactu, "N", "fecfaccl", vFecfactu, "F")
                End If
                If SQL <> "" Then
                    mCtaPrev = SQL
                    LeerDatos = True
                Else
                    LeerDatos = False
                    Mens = "La cuenta prevista de cobro para la factura no puede ser nula."
                End If
            Else
                LeerDatos = True
            End If
        Else
            LeerDatos = True
        End If
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    LeerDatos = False
End Function


'======================================================================
'GRABAR EN TESORERIA
'======================================================================

Public Function InsertarEnTesoreria(vTextosCSB As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim SQL As String, Codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAuxConta As String 'para insertar en conta.scobro
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Currency 'importe para insertar en la svenci
Dim ImpVenci2 As Currency 'importe para insertar en conta.scobro
Dim Knumerovenci As Byte
Dim TotalFactura3 As Currency   'Por si acaso lleva aportacion al terminal
Dim ImporteDeLaFactura As Currency  'por si lleva pago por adelantado
Dim FormaPagoAdelantado As Integer  '-1=nada
Dim ImporteAdelantado As Currency
Dim UtilizaFormaPagoAlternativa As Boolean
'1 Julio 2009. Los graba en scobro
Dim CadenaDatosFiscales As String
Dim J As Integer
Dim FraTelefonia As Boolean
Dim NumeroDeVencimientos As Byte
Dim NuevaNorma19 As Boolean

Dim FormapagoAportacion  As Integer   'De momento NO la lee de parametros

    On Error GoTo EInsertarTesoreria

    Set rsVenci = New ADODB.Recordset

    'Enero 2013
    'Facturas de telefonia, en los textos de los vencimientos
    'Tenemos que poner otras cosas
    FraTelefonia = False
    If mCodTipoM = "FAT" Then
            FraTelefonia = True
    Else
        
        'Si es interna y algo mas
        
        If mCodTipoM = "FAI" Then
'            If mNumventa = 1 Then
'                FraTelefonia = True
'                mNumventa = 0
'            End If
        End If
    End If
        
        
        'Facturas de telefonia.
        '-------------------------------------------------------
        
  
    If FraTelefonia Then
    
        
       
    
        If vParamAplic.TieneTelefonia2 <> 2 Then
        
            If NuevasNormasBancarias Then
                SQL = "SELECT * from tel_cab_factura WHERE serie= '" & LetraSerie & "' AND ano = " & Year(FecFactu)
                SQL = SQL & " AND NumFact = " & NumFactu
                rsVenci.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                'NO puede ser eof
                textcsb33 = "'NIF " & Right(Space(9) & rsVenci!NIF, 9) & " TFNO :" & Mid(rsVenci!Telefono, 1, 9)
                                        
                textcsb33 = textcsb33 & "  FACT: " & rsVenci!Serie & Format(rsVenci!NumFact, "000000") & "-" & Format(rsVenci!Ano, "0000") & _
                                        " DE " & Format(rsVenci!Fecha, "dd/mm/yyyy")
        
                textcsb33 = textcsb33 & "  TOT. " & Format(rsVenci!total, "#,###,##0.00") & "'"
                
                'la ultima linea del CSB, la der y la izda
                CadValues2 = "'BASE: " & Format(rsVenci!BaseImponible, "##0.00")
                                                            
                CadValues2 = CadValues2 & "  CUOTA: " & Format(rsVenci!Cuota, "##0.00") & "  (" & Format(rsVenci!TipoIVA, "#0.00") & "%)'"
        
            
            
                CadValues = ""
                vTextosCSB = CadValues2
                
                If Not vParamAplic.ContabilidadNueva Then
                    'Lleva el formato antiuo con todos los txtcsb...
                    For Knumerovenci = 1 To 8
                        vTextosCSB = vTextosCSB & ",NULL"
                    Next
                    
                End If
                rsVenci.Close
                
            Else
                
                'Cuotas
                'Linea(2) = "CUOTAS              " & completaTextoIz(Format(totCuotas, "#,###,##0.00"), 12)
                If Not vParamAplic.ContabilidadNueva Then
                    ImporteDeLaFactura = TelSumlineas(0, LetraSerie, Year(FecFactu), NumFactu)
                    CadValues = Format(ImporteDeLaFactura, "#,###,##0.0000")
                    CadValues = "CUOTAS              " & Right(Space(12) & CadValues, 12)
                    vTextosCSB = vTextosCSB & "," & DBSet(CadValues, "T")
                    
                    'Consumos
                    'Linea(3) = "CONSUMOS            " & completaTextoIz(Format(totConsumos, "#,###,##0.00"), 12)
                    ImporteDeLaFactura = TelSumlineas(1, LetraSerie, Year(FecFactu), NumFactu)
                    CadValues = Format(ImporteDeLaFactura, "#,###,##0.0000")
                    CadValues = "CONSUMOS            " & Right(Space(12) & CadValues, 12)
                    vTextosCSB = vTextosCSB & "," & DBSet(CadValues, "T")
                    
                    'Cargos
                    ' Linea(4) = "CARGOS              " & completaTextoIz(Format(totCargos, "#,###,##0.00"), 12)
                    ImporteDeLaFactura = TelSumlineas(3, LetraSerie, Year(FecFactu), NumFactu)
                    CadValues = Format(ImporteDeLaFactura, "#,###,##0.0000")
                    CadValues = "CARGOS              " & Right(Space(12) & CadValues, 12)
                    vTextosCSB = vTextosCSB & "," & DBSet(CadValues, "T")
            
            
                    'Descuentos
                    ' Linea(5) = "DESCUENTOS         " & completaTextoIz(Format(totDescuentos, "#,###,##0.00"), 12)
                    ImporteDeLaFactura = TelSumlineas(2, LetraSerie, Year(FecFactu), NumFactu)
                    CadValues = Format(ImporteDeLaFactura, "#,###,##0.0000")
                    CadValues = "DESCUENTOS          " & Right(Space(12) & CadValues, 12)
                    vTextosCSB = vTextosCSB & "," & DBSet(CadValues, "T")
                    
                    'Separador
                    'Linea(6) = "--------------------" & String(12, "-")
                    CadValues = String(40, "-")
                    vTextosCSB = vTextosCSB & "," & DBSet(CadValues, "T") & "," & DBSet(CadValues, "T")
            
            
                    'TOTALES
                    'Linea(7) = "BASE                " & completaTextoIz(Format(mrs!BaseImponible, "#,###,##0.00"), 12) & _
                                                                " IVA: " & Format(mrs!TipoIVA, "#0") & "%" & _
                                                                " CUOTA: " & Format(mrs!Cuota, "#,###,##0.00") & _
                                                                " TOTAL: " & Format(mrs!total, "#,###,##0.00")
                
                
                End If
                    
                
                
                
                'Y el priemro
                SQL = "SELECT * from tel_cab_factura WHERE serie= '" & LetraSerie & "' AND ano = " & Year(FecFactu)
                SQL = SQL & " AND NumFact = " & NumFactu
                rsVenci.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                'NO puede ser eof
                textcsb33 = "'TELEFONO:" & rsVenci!Telefono & "    NIF:" & rsVenci!NIF & "'"
                                        
                CadValues = "FACTURA:" & rsVenci!Serie & " " & Format(rsVenci!Ano, "0000") & "-" & Format(rsVenci!NumFact, "000000") & _
                                        " FECHA: " & Format(rsVenci!Fecha, "dd/mm/yyyy")
        
                'En la norma nueva, lo metemos todo en el primer campo
                If vParamAplic.ContabilidadNueva Then
                    textcsb33 = Replace(textcsb33, "'", "")   'quito las comillas
                    textcsb33 = textcsb33 & " " & CadValues
                    textcsb33 = Replace(textcsb33, "  ", " ")   'quito dos espacios en blanco segidos
                    textcsb33 = "'" & Mid(textcsb33, 1, 80) & "'"
                    CadValues = ""
                End If
                
        
                'la ultima linea del CSB, la der y la izda
                CadValues2 = "BASE                " & Right(Space(12) & Format(rsVenci!BaseImponible, "#,###,##0.00"), 12) & _
                                                            " IVA:" & Format(rsVenci!TipoIVA, "#0") & ""
                CadValues3 = "%      CUOTA: " & Format(rsVenci!Cuota, "#,###,##0.00") & "     TOTAL: " & Format(rsVenci!total, "#,###,##0.00")
        
        
                If vParamAplic.ContabilidadNueva Then
                    CadValues = CadValues2 & CadValues3
                    Do
                        Knumerovenci = InStr(1, CadValues, "  ")
                        If Knumerovenci > 0 Then CadValues = Replace(CadValues, "  ", " ")
                    Loop Until Knumerovenci = 0
                    CadValues3 = ""
                    CadValues2 = ""
                    
                    vTextosCSB = "" 'El resto de campos los ponemos quitamos ya que solo estan 2 CSB text
                End If
                
                'El Cadvalues sera le txtcsb41, es decir, el primero del INSERT
                vTextosCSB = DBSet(CadValues, "T") & vTextosCSB
                
                If Not vParamAplic.ContabilidadNueva Then vTextosCSB = vTextosCSB & "," & DBSet(CadValues2, "T") & "," & DBSet(CadValues3, "T")
                rsVenci.Close
            End If
       Else
            'CATADAU
            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd mmm yyyy") & "'"
            
            vTextosCSB = ""
            For Knumerovenci = 1 To 9
                vTextosCSB = vTextosCSB & ",NULL"
            Next
            If vParamAplic.ContabilidadNueva Then vTextosCSB = ", NULL"   'Solo dos campos txtcsb
            
            vTextosCSB = Mid(vTextosCSB, 2) 'QUITO LA primera coma
            
       End If



    Else
        'Resto de Fras normales
        If vTextosCSB = "" Then
            vTextosCSB = "NULL"
            If Not vParamAplic.ContabilidadNueva Then vTextosCSB = vTextosCSB & ",NULL,NULL"
            
            
        Else
            
            CadValues = DBSet(RecuperaValor(vTextosCSB, 1), "T", "S")
            CadValues3 = ""
            If Not vParamAplic.ContabilidadNueva Then
                CadValues2 = ", " & DBSet(RecuperaValor(vTextosCSB, 2), "T", "S")
                CadValues3 = ", " & DBSet(RecuperaValor(vTextosCSB, 3), "T", "S")
            End If
            vTextosCSB = CadValues & CadValues2 & CadValues3
            
        End If
    
        'Enero 2013
        'Se han añadido mas txtcsb para los recibos bancarios de telefono
        If Not vParamAplic.ContabilidadNueva Then vTextosCSB = vTextosCSB & ",NULL,NULL,NULL,NULL,NULL,NULL"
            
    End If
    
    b = False
    InsertarEnTesoreria = False
    CadValues3 = ""
    CadValues = ""
    CadValues2 = ""

    'campo para insertar en conta.scobro de Tesoreria. pAra las de telefonia ya lo ha creado arriba
    If mCodTipoM <> "FAT" Then textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd mmm yyyy") & "'"


    'Datos fiscales en scobro     Julio 2009
    'nomclien,domclien,pobclien, cpclien,proclien
    CadenaDatosFiscales = DBSet(mNomClien, "T") & "," & DBSet(mDomClien, "T") & "," & DBSet(mPobClien, "T")
    CadenaDatosFiscales = CadenaDatosFiscales & "," & DBSet(mCodpobla, "T") & "," & DBSet(mProClien, "T")
    
    
    ImporteDeLaFactura = TotalFac
    Knumerovenci = 1
    FormaPagoAdelantado = -1
    ImporteAdelantado = 0
    UtilizaFormaPagoAlternativa = False
    
    'Obtener el Nº de Vencimientos de la forma de pago
    
    J = -1   'por si acaso no llega al minimo para utilizar la forma de pago alternativa
    
    'Solo lo haremos para importes > que cero y sin cheque ni aportacion
    If ImporteDeLaFactura > 0 And ImpCheque = 0 And Aportacion = 0 Then
        'ImoporteAdelantado  FormaPagoAdelantado

        SQL = "SELECT * FROM sforpa WHERE codforpa=" & ForPago
        rsVenci.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If rsVenci.EOF Then Err.Raise 513, "Leyendo forma de pago: " & ForPago
        
        'Si pago algo por adelantado
        If Not IsNull(rsVenci!forpapor) Then
            ImpVenci = DBLet(rsVenci!poradela, "N")
            If ImpVenci > 0 Then
                ImporteAdelantado = Round2((ImporteDeLaFactura * ImpVenci) / 100, 2)
                ImporteDeLaFactura = ImporteDeLaFactura - ImporteAdelantado
                FormaPagoAdelantado = rsVenci!forpapor
                Knumerovenci = Knumerovenci + 1  'Empezaran en el dos. El primero lo reservo para el adelantado
            End If
        End If
        
        'Si utiliza la forma de pago alternativa
        If Not IsNull(rsVenci!forpaalt) Then
            If Not IsNull(rsVenci!ImporMin) Then
                'Veamos si no lleva al mimino
                'OK. No llega al minimo, utlizo la forma de pago alternativa
                If rsVenci!ImporMin > ImporteDeLaFactura Then J = rsVenci!forpaalt
            End If
        End If
        
        
        If J >= 0 Then
            'FORMA DE PAGO ALTERNATIVA.....
            'SI QUE TIENE
            UtilizaFormaPagoAlternativa = True
            ForPago = J
            rsVenci.Close
            J = -1
        Else
            J = 0
        End If
    
    End If
    'Lo volvemos a abrir, o abrimos por primera vez
    If J = -1 Then
        SQL = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
        rsVenci.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        

    End If
    
    If Not rsVenci.EOF Then
    
        If TotalFac < 0 Then
            'Fras NEGATIVAS solo hay un vencimiento
            NumeroDeVencimientos = 1
        Else
            NumeroDeVencimientos = CByte(rsVenci!numerove)
        End If
    
    
        
        If rsVenci!numerove > 0 And CCur(ImporteDeLaFactura) <> 0 Then
        
            'Comporbamos si el importe es <>0
            'Obtener los dias de pago del cliente
            SQL = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            SQL = SQL & " FROM sclien "
            SQL = SQL & " WHERE codclien=" & Cliente
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            Codmacta = DBSet(Rs!Codmacta, "T")
'            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not Rs.EOF Then
                cadValuesAux = "('" & codtipom & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                CadValuesAuxConta = "('" & LetraSerie & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "', "
                '                    Añadire a la cadena fija esta los valores de textcsb41,txcs
                CadValuesAuxConta = CadValuesAuxConta & vTextosCSB & ","
                '-------- Primer Vencimiento
               
                'FECHA VTO
                FecVenci = CDate(FecFactu)
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(Rs!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & Knumerovenci & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAuxConta & Knumerovenci & ", "
                CadValues2 = CadValues2 & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                TotalFactura3 = ImporteDeLaFactura - Aportacion
                If NumeroDeVencimientos = 1 Then
                    ImpVenci = TotalFactura3
                    ImpVenci2 = TotalFactura3 - ImpCheque
                Else
                    
                    ImpVenci = Round2(TotalFactura3 / NumeroDeVencimientos, 2)
                    ImpVenci2 = Round2((TotalFactura3 - ImpCheque) / NumeroDeVencimientos, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * NumeroDeVencimientos <> TotalFactura3 Then
                        ImpVenci = Round(ImpVenci + (TotalFactura3 - ImpVenci * NumeroDeVencimientos), 2)
                    End If
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If (ImpVenci2 * NumeroDeVencimientos) + ImpCheque <> TotalFactura3 Then
                        ImpVenci2 = Round(ImpVenci2 + (TotalFactura3 - ImpCheque - (ImpVenci2 * NumeroDeVencimientos)), 2)
                    End If
                End If
                
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", "
                CadValues2 = CadValues2 & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & DBSet(mIBAN, "T", "S") & ", "
                CadValues2 = CadValues2 & textcsb33 & ", " & DBSet(Agente, "N")
                
                'departamento y transfer
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                '1 Julio 2009
                ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                 CadValues2 = CadValues2 & "," & CadenaDatosFiscales & ")"
                
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For J = 2 To NumeroDeVencimientos
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(Rs!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    End If
                    Knumerovenci = Knumerovenci + 1
                    CadValues = CadValues & ", " & cadValuesAux & Knumerovenci & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & Knumerovenci & ", " & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFactura3 / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFactura3 - ImpCheque) / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & DBSet(IBAN, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                    '1 Julio 2009
                    ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                    CadValues2 = CadValues2 & "," & CadenaDatosFiscales & ")"
                    
                Next J
                
                '--- Cheque regalo: laura 1/12/2006   y/o    Aportacion terminal
                'si hay cheque regalo insertar una linea más para la forma de pago correspondiente y el importe del cheque
                If ImpCheque > 0 Then
                
                    Knumerovenci = J
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & Knumerovenci & "," & Codmacta & ", " & vParamAplic.ForPagoChequeRegalo & ", "
                    
                    'FECHA VTO
                    CadValuesAuxConta = "primerve"
                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", vParamAplic.ForPagoChequeRegalo, "N", CadValuesAuxConta)
                    FecVenci = CDate(FecFactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                        MsgBox "FALTA cheque regalo con forma de pago no en EFECTIVO", vbInformation
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & DBSet(IBAN, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL)"
                End If
                
                
                'Aportacion al terminal
                'Si tiene cuenta aportacion entonces añadiremos un cobro
                If Me.Aportacion > 0 Then
                    If vParamAplic.ctaAportacion = "" Then
                        MsgBox "Error cta aportacion NULL", vbExclamation
                        Exit Function
                    End If
                        
                   'Montamos EL SQL para el cobro de la aportacion al termina
                    'ForPago
                    Knumerovenci = Knumerovenci + 1
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & Knumerovenci & ",'" & vParamAplic.ctaAportacion & "', "
                    
                    FormapagoAportacion = -1
                    
                    'Vemos primero Efectivos con texto efectivo o contador
                    CadValuesAuxConta = "(nomforpa like '%efec%' or nomforpa like '%conta%') and tipforpa"
                    CadValuesAuxConta = DevuelveDesdeBDNew(conAri, "sforpa", "codforpa", CadValuesAuxConta, "0", "N", CadValuesAuxConta)
                    If CadValuesAuxConta <> "" Then
                        'OK ya tenemos la forma de pago
                        TipForPago = 0
                        FormapagoAportacion = CadValuesAuxConta
                        CadValuesAuxConta = 0
                    Else
                        'Provamos de otro modo
                        CadValuesAuxConta = "primerve"
                        'TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N", CadValuesAuxConta)
                        TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", ForPago, "N", CadValuesAuxConta)
                        FormapagoAportacion = ForPago
                    End If
                    FecVenci = CDate(FecFactu)
                    FecVenci = FecVenci + CByte(CadValuesAuxConta)
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
                       ' MsgBox "FALTA config con forma de pago aportacion no en EFECTIVO. ", vbInformation
                       ' FecVenci = MecVenci
                        
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    
                    
                    CadValues2 = CadValues2 & FormapagoAportacion & ", "
                    
                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
                    CadValues2 = CadValues2 & DBSet(Aportacion, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & DBSet(IBAN, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                    'departamento
                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                    'Modificacion que metia los datos fiscales
                    CadValues2 = CadValues2 & ",NULL,NULL,NULL,NULL,NULL)"
                End If
                
            End If
            Rs.Close
            
            'Si habia un primer pago como aportacion entonces lo metemos aqui, con el numero venci=1
            If ImporteAdelantado <> 0 Then
                    FecVenci1 = Now
                    Knumerovenci = 1
                    CadValues = CadValues & ", " & cadValuesAux & Knumerovenci & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAuxConta & Knumerovenci & ", " & Codmacta & ", " & FormaPagoAdelantado & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = ImporteAdelantado
                    ImpVenci2 = ImporteAdelantado
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & DBSet(IBAN, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") & ",NULL"
                    '1 Julio 2009
                    ' Datos fiscales en scobro nomclien , domclien, pobclien, cpclien, proclien
                    CadValues2 = CadValues2 & "," & CadenaDatosFiscales & ")"
                    
            End If
            
            
            
            
            
            
        Else
            'totalfac =0 and numerovtos >=1
            b = True
        End If
        
        Set Rs = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        SQL = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        SQL = SQL & " VALUES " & CadValues
        conn.Execute SQL
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        '01/09/06
'        If (NumTicket = "") Or (NumTicket <> "" And TipForPago <> 0) Then
            If CuentaPrev <> "" Then
                'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                'forma de pago de la factura. Sino existe insertarla
                'vemos si existe en la conta
                b = InsertarFormaPagoEnConta(ForPago, MenError)
                
                'si hay cheque regalo comprobar q existe su forma de pago en conta
                'sino insertarla
                If b And ImpCheque > 0 Then b = InsertarFormaPagoEnConta(vParamAplic.ForPagoChequeRegalo, MenError)


                
                
                If b Then
                    'Insertamos en la tabla scobro de la CONTA
                    If vParamAplic.ContabilidadNueva Then
                        SQL = "INSERT INTO cobros (numserie, numfactu, fecfactu,text41csb, "
                        SQL = SQL & "numorden , Codmacta, codforpa, FecVenci, ImpVenci, ctabanc1, entidad, oficina, control"
                        SQL = SQL & ", cuentaba,iban,text33csb,agente,departamento,transfer "
                        
                    Else
                        SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl,text41csb ,text42csb, text43csb,"
                        SQL = SQL & "text51csb,text52csb,text53csb,text61csb,text62csb,text63csb,"
                        SQL = SQL & "numorden , Codmacta, codforpa, FecVenci, ImpVenci, ctabanc1, codbanco, codsucur,digcontr "
                        SQL = SQL & ", cuentaba,iban,text33csb,agente,departamento,transfer "
                        
                    End If
                    SQL = SQL & ",nomclien,domclien,pobclien, cpclien,proclien)"
                    SQL = SQL & " VALUES " & CadValues2
                    ConnConta.Execute SQL
                End If
                
            Else
                'DAVID ####
                'ENERO 2008
                'Si no inserto en tesoreria por que ctaprvsta="" entonces dejo continuar
                b = True
        
        End If





    End If
    
    
    If UtilizaFormaPagoAlternativa Then
        SQL = "UPDATE scafac set codforpa = " & ForPago
        SQL = SQL & " WHERE codtipom='" & Me.codtipom & "' AND numfactu = " & Me.NumFactu & " and fecfactu=" & DBSet(Me.FecFactu, "F")
        ejecutar SQL, False
    End If
'    If b Then b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Insertar en Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreria = b
End Function





Private Sub ReiniciarImportesFactura()
    Me.BaseImp = 0
    Me.BaseIVA1 = 0
    Me.BaseIVA2 = 0
    Me.BaseIVA3 = 0
    
    Me.BrutoFac = 0
    
    Me.ImpIVA1 = 0
    Me.ImpIVA2 = 0
    Me.ImpIVA3 = 0
    
    Me.PorceIVA1 = 0
    Me.PorceIVA2 = 0
    Me.PorceIVA3 = 0
    
    
    Me.ImpIVA1RE = 0
    Me.ImpIVA2RE = 0
    Me.ImpIVA3RE = 0
    
    Me.PorceIVA1RE = 0
    Me.PorceIVA2RE = 0
    Me.PorceIVA3RE = 0
    
    
    Me.TipoIVA1 = 0
    Me.TipoIVA2 = 0
    Me.TipoIVA3 = 0
    
    Me.ImpGnral = 0
    Me.ImpPPago = 0
    Me.TotalFac = 0
    

End Sub


Public Function CalcularDatosFactura(cadWhere As String, NomTabla As String, NomTablaLin As String, IVA_Antiguo As Boolean) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim I As Integer

Dim SQL As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim Impaux As Currency
Dim ImpIva As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vBruto As Currency

 ' 0.-IVA normal      1.- Recargo equivalencia    2.- Exento  3.-Exento INTRACOM
 ' Junio 2012:        4.- Reducido
Dim exento_IVA As Byte
Dim conDesplaz As Boolean
    
    
Dim TipoIVAAplicable As Integer
Dim RE1 As String
Dim impRE As Currency
Dim TotRe As Currency


Dim ImporteParaComparar As Currency
Dim NueroIvas As Integer


    CalcularDatosFactura = False
    On Error GoTo ECalcular

    If codtipom <> "FAM" Then ReiniciarImportesFactura

   '------------------------------------------------------------------------
   'En estos dos procedimientos solo entra cuando va a generar la factura y
   'codTipoMov="FAV", cuando llamamos desde pedidos, albaranes,etc, para mostrar
   'el resumen de la factura codTipoMov="ALV","PEV",... y no entra
   '-------------------------------------------------------------------------
   
   'comprobar si hay que facturar DESPLAZAMIENTOS
   CalcularDatosFactura = CalcularDesplazamiento(cadWhere, NomTabla)
   If Not CalcularDatosFactura Then Exit Function
   
   'comprobar si hay que facturar BONIFICACIONES
    CalcularDatosFactura = CalcularBonificacion(cadWhere, NomTablaLin)
   If Not CalcularDatosFactura Then Exit Function
   '-------------------------------------------------------------------------
   
   'comprobar si el cliente de la factura esta EXENTO de IVA
   If codtipom = "ALZ" Or codtipom = "FAZ" Or codtipom = "PEZ" Or codtipom = "FAI" Or codtipom = "ALI" Then
        'Albaranes de B. No llevan IVA     Feb 2011.    Los pedidos de B tampoco llevaran IVA
        '                                  Mar 2011.    La facturacion interna (ALI,FAI) tampoco llevan IVA
        exento_IVA = 2   'Exento IVA
   Else
        exento_IVA = ClienteExentoIVA
   End If
   
  
   If exento_IVA = 2 Or exento_IVA = 3 Then
        cadWhere = Replace(cadWhere, NomTabla, NomTablaLin)
        SQL = "SELECT sum(importel) as bruto"
        SQL = SQL & " FROM " & NomTablaLin
        SQL = SQL & " WHERE " & cadWhere
   Else
    'Agrupar el importe bruto por tipos de iva
       If codtipom = "FAM" Then
            'para los mantenimientos obtenemos de la tabla de tipos de contrato(stipco.codartic)
            'el articulo que se va a utilizar para facturar los mantenimientos
    '        CadAux = DevuelveDesdeBDNew(conAri, "stipco", "codartic", "codtipco", TipCoMan, "T")
            If ArticFac <> "" Then 'tenemos el articulo
                SQL = "SELECT codartic,nomartic,sartic.codigiva, tiposiva.porceiva "
                SQL = SQL & " FROM sartic INNER JOIN conta" & vParamAplic.NumeroConta & ".tiposiva ON sartic.codigiva=conta" & vParamAplic.NumeroConta & ".tiposiva.codigiva"
                SQL = SQL & " WHERE codartic=" & DBSet(ArticFac, "T")
            End If
       Else
            cadWhere = Replace(cadWhere, NomTabla, NomTablaLin)
            
            If exento_IVA = 4 Then
                'Se le aplica el ivareducido. Reemplzao los articulos que lleven IVA normal por iva reducido
                SQL = "SELECT if(sartic.codigiva=" & vParamAplic.TipoIVA1 & "," & vParamAplic.TipoIVA2 & ","
                SQL = SQL & "sartic.codigiva) as codigiva,sum(importel) as bruto"
            Else
            
                'NORMAL, todo normal
                SQL = "SELECT sartic.codigiva,sum(importel) as bruto"
            End If
            SQL = SQL & " FROM " & NomTablaLin & " LEFT OUTER JOIN sartic ON " & NomTablaLin & ".codartic=sartic.codartic"
            SQL = SQL & " WHERE " & cadWhere
            SQL = SQL & " GROUP BY codigiva "
            SQL = SQL & " ORDER BY codigiva "
        End If
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenKeyset, adLockOptimistic, adCmdText
    NueroIvas = 0 'Numero IVAS
    ImporteParaComparar = 0
    
    'VEre cuantos IVAS tiene y el bruto. Para cuadrar los decimales
    If codtipom = "FAM" Then
        'No hago la comrobacion para cuadre decimales
        NueroIvas = 1
    Else
            
        While Not Rs.EOF
            TotBruto = DBLet(Rs!bruto, "N")
            ImporteParaComparar = ImporteParaComparar + TotBruto
            NueroIvas = NueroIvas + 1
            Rs.MoveNext
        Wend
    
    End If
    If NueroIvas > 0 Then
        TotImpIVA = CalcularPorcentaje(ImporteParaComparar, DtoPPago, 2)
        TotBruto = CalcularPorcentaje(ImporteParaComparar, DtoGnral, 2)
        ImporteParaComparar = ImporteParaComparar - TotImpIVA - TotBruto '(bruto - DtoPP -dtogral

        Rs.MoveFirst
    End If
    TotBruto = 0
    TotImpIVA = 0
    TotRe = 0
    I = 1
    conDesplaz = False

    While Not Rs.EOF
        'Aqui vamos acumulando la suma del importe bruto de las lineas
        If codtipom = "FAM" Then
            vBruto = BrutoFac
        Else
            vBruto = Rs!bruto
        End If
        
        TotBruto = TotBruto + vBruto
        ImpBImIVA = vBruto
        
        If exento_IVA < 2 Or exento_IVA = 4 Then    'Con IVA normal o RE   o Reducido
            'comprobar que si hay km de DESPLAZAMIENTO se añadan al tipo de IVA correspondiente
            If TotalKm > 0 Then
                If despIVA = Rs!CodigIVA Then
                    vBruto = vBruto + ImporteL
                    TotBruto = TotBruto + ImporteL
                    conDesplaz = True
                End If
            End If
        
            'comprobar si hay BONIFICACIONES que aplicar a la Factura
            If IVABonif1 = Rs!CodigIVA Then
                vBruto = vBruto - TotBonif1
                TotBruto = TotBruto - TotBonif1
            ElseIf IVABonif2 = Rs!CodigIVA Then
                vBruto = vBruto - TotBonif2
                TotBruto = TotBruto - TotBonif2
                
            ElseIf IVABonif3 = Rs!CodigIVA Then
                vBruto = vBruto - TotBonif3
                TotBruto = TotBruto - TotBonif3
            End If
            
        
            'Aplicarle el dto ppago
            '---- Laura: 05/10/2006
            Impaux = CalcularPorcentaje(vBruto, DtoPPago, 2)
'            ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoPPago)))
            '---- Laura: 27/09/2006
'            ImpAux = Round(ImpAux, 2)
            '----
            ImpBImIVA = vBruto - Impaux '(bruto - DtoPP)
            
            'Aplicarle el dto grnal
            '---- Laura: 05/10/2006
            Impaux = CalcularPorcentaje(vBruto, DtoGnral, 2)
            ImpBImIVA = ImpBImIVA - Impaux '(bruto - Dtogn)
            
            
            'NUEVO. 22 Mayo 2008
            'Si tiene mas de un tipo de iva, compruebo en la ultima de las bases
            'Si el total sera igual al total a comparar
            If I > 1 And I = NueroIvas Then
                'ImporteParaComparar
                If ImporteParaComparar - ImpBImIVA <> 0 Then
                    ImpBImIVA = ImporteParaComparar
                End If
            Else
                ImporteParaComparar = ImporteParaComparar - ImpBImIVA  'Para comparar en la ultima (si ha lugar)
            End If
            
            'Obtener el % de IVA
            cadAux = Rs!CodigIVA    'Por si no tiene  RE o si teniendo RE no tiene el correspondiente
            If exento_IVA = 1 Then cadAux = CStr(vParamAplic.DevuleveTipoIVA_RE(Rs!CodigIVA))
            TipoIVAAplicable = CInt(cadAux)
            
            If IVA_Antiguo And exento_IVA < 2 Then
               
                    'JUNIO 2010
                    'Debido al cambio de IVA, los albaranes rectificativos podran ir en el IVA que llevaran en su momento
                    'Es decir, que aunque ahora sea al 18(a partir de Julio), para un albr rectifi podria ser
                    'con iva 16
                    cadAux = CStr(vParamAplic.DevuelveTipoIVA_OLD(exento_IVA = 1, Rs!CodigIVA))
                    TipoIVAAplicable = CInt(cadAux)
            End If
            
            
            
            'ANTES
            'cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(RS!Codigiva), "N")
            RE1 = "porcerec"
            cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", cadAux, "N", RE1)
            If exento_IVA <> 1 Then
                RE1 = "0"
                impRE = 0
            Else
                impRE = CalcularPorcentaje(ImpBImIVA, CCur(RE1), 2)
                TotRe = TotRe + impRE
            End If
            

            
            
            'aplicar el IVA a la base imponible de ese tipo
            '---- Laura: 05/10/2006
            ImpIva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + ImpIva
            
            
            Select Case I
                Case 1  'IVA 1
                    TipoIVA1 = TipoIVAAplicable    'Antes Rs!codigiva
                    
                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA1 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA1 = ImpIva
                    
                    'Recargo equivalencia
                    PorceIVA1RE = RE1
                    ImpIVA1RE = impRE
                     
                Case 2  'IVA 2
                    TipoIVA2 = TipoIVAAplicable     'Antes Rs!codigiva
                    
                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA2 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA2 = ImpIva
    
                    'Recargo equivalencia
                    PorceIVA2RE = RE1
                    ImpIVA2RE = impRE
    
    
                Case 3  'IVA 3
                    TipoIVA3 = TipoIVAAplicable     'Antes Rs!codigiva
                    
                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                    
                    PorceIVA3 = cadAux '% de IVA
                    
                    'Importe total con IVA
                    ImpIVA3 = ImpIva
                    
                    'Recargo equivalencia
                    PorceIVA3RE = RE1
                    ImpIVA3RE = impRE
                    
            End Select
            I = I + 1
        Else 'EXENTO DE IVA
            'si hay que factura KM añadirlos al bruto
            If TotalKm > 0 Then vBruto = vBruto + ImporteL
            
            If exento_IVA = 2 Then
                'Exento normal
                If vParamAplic.IVA_Exento2 = 0 Then MsgBox "Falta configurar el IVA exento", vbExclamation
                TipoIVA1 = vParamAplic.IVA_Exento2
            Else
                'exento intracom
                If vParamAplic.IVA_Intracomunitario = 0 Then MsgBox "Falta configurar el IVA intracomunitario", vbExclamation
                TipoIVA1 = vParamAplic.IVA_Intracomunitario
                
            End If
        
                
                
            'DAVID.  20 Julio 2009
            Impaux = CalcularPorcentaje(vBruto, DtoPPago, 2)
            ImpIVA1 = Impaux
            Impaux = CalcularPorcentaje(vBruto, DtoGnral, 2)
            ImpIVA1 = ImpIVA1 + Impaux
                
            vBruto = vBruto - ImpIVA1
            ImpIVA1 = 0
                
            
                
                BaseIVA1 = vBruto 'BASE IMPONIBLE
                
                PorceIVA1 = 0 '% de IVA
                
                'Importe total con IVA
                ImpIVA1 = 0
                
                'Recargo equivalencia
                PorceIVA1RE = 0
                ImpIVA1RE = 0

        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'si tiene desplazamiento comprobar que se ha añadido en el IVA
    'si no calcularlo
    If TotalKm > 0 And conDesplaz = False Then
    
        'FALTA##### con la conversion del Tipo de iva a RE
    
        TotBruto = TotBruto + ImporteL
        vBruto = ImporteL
        ImpBImIVA = ImporteL
        
        'Aplicarle el dto ppago
        '---- Laura: 05/10/2006
        Impaux = CalcularPorcentaje(vBruto, DtoPPago, 2)
'        ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoPPago)))
        ImpBImIVA = ImpBImIVA - Impaux '(bruto - DtoPP)
            
        'Aplicarle el dto grnal
        '---- Laura: 05/10/2006
        Impaux = CalcularPorcentaje(vBruto, DtoGnral, 2)
'        ImpAux = CCur(CalcularDto(CStr(vBruto), CStr(DtoGnral)))
        ImpBImIVA = ImpBImIVA - Impaux '(bruto - Dtogn)
            
        'Obtener el % de IVA
        cadAux = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(despIVA), "N")
            
        'cuota IVA: aplicar el IVA a la base imponible de ese tipo
        'Laura: 05/10/2006
        Impaux = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
'        ImpAux = CalcularDto(CStr(ImpBImIVA), cadAux)
        ImpIva = Impaux
            
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + ImpIva
        
        Select Case I
            Case 2 'IVA 2
                TipoIVA2 = despIVA
                
                BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA2 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA2 = ImpIva
                
            Case 3 'IVA 3
                TipoIVA3 = Rs!CodigIVA
                
                BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA3 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA3 = ImpIva
        End Select
    End If
    
    BrutoFac = TotBruto
    
    'Aplicarle el dto ppago
    '---- Laura: 05/10/2006
    ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
'    ImpPPago = CCur(CalcularDto(CStr(TotBruto), CStr(DtoPPago)))
    '---- Laura: 27/09/2006
'    ImpPPago = Round(ImpPPago, 2)
    '----
    
    'Aplicarle el dto general
    '---- Laura: 05/10/2006
    ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
'    ImpGnral = CCur(CalcularDto(CStr(TotBruto), CStr(DtoGnral)))
    '---- Laura: 27/09/2006
'    ImpGnral = Round(ImpGnral, 2)
    '----
    
    
    'Base Imponible
    BaseImp = TotBruto - ImpPPago - ImpGnral
    
    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA + TotRe
    
    CalcularDatosFactura = True
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFactura = False
    Else
        CalcularDatosFactura = True
    End If
End Function



Public Function PasarAlbaranesAFactura(TipoAlb As String, cadSQL As String, textoCSB As String, Optional ErroresAux As String, Optional EsTraspasoOfeFAZ As Boolean) As Boolean
'IN ->  cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'       TextosCSB:  Para la tesoreria
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String
Dim HacerCambioIVA As Boolean
Dim Aux As String
Dim ComprobarRecargoFinanciero As Boolean
Dim PorceRecargo As Currency
'Dim PortesTipoHerbelca As Currency


    On Error GoTo ETraspasoAlbFac

    PasarAlbaranesAFactura = False
    b = True
    ErroresAux = ""
    ComprobarRecargoFinanciero = False
    
    Select Case TipoAlb
        Case "ALV", "ALM": 'ALV: Albaranes venta a clientes
                    'ALM: Albaran mostrador
            codtipom = "FAV"
            
            '#David 08/11/2010 recuperar facturas de Alzira
            If TipoAlb = "ALM" Then
                If vParamAplic.FrasMostradorSerieDistinta Then codtipom = "FMO"
            End If
            
            ComprobarRecargoFinanciero = True
        Case "ALR": 'Albaranes de reparacion en clientes
            codtipom = "FAR"
            ComprobarRecargoFinanciero = True
        Case "ART": 'Albaranes de factura rectificativa
            codtipom = "FRT"
        Case "ALS": 'Albaranes de servicios [SERVICIOS]
            codtipom = "FAS"
            ComprobarRecargoFinanciero = True
        Case "ALZ"  'Albaranes "presupuestos". Es decir, el "B"
            codtipom = "FAZ"
        Case "ALI"
            codtipom = "FAI"
        Case "ALT"
            'Telfonia
            codtipom = "FAT"
        Case "ALG"
            'Agua Bolbaite
            codtipom = "FAG"
            
        'EULER
        Case "ALO"
            'Orden de trabajo
            codtipom = "FAO"
        Case "ALE"
            'Trabajo externo
            codtipom = "FAE"
            
            
    End Select
      
     
    OpeFactu = PonerTrabajadorConectado("")
    If OpeFactu <> "" Then
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N")
    Else
        AlmFactu = "1"
    End If
      
      
    'Veo si el cliente lleva portes y/o recargo financiero
    
    
    'Previo. Veremos si la forma de pago, y el cliente llevan RECARGO FINANCIERO
    If vParamAplic.ArticuloRecargoFinanciero = "" Then ComprobarRecargoFinanciero = False
      
    
        
  
    
    If ComprobarRecargoFinanciero Then
        Aux = DevuelveDesdeBD(conAri, "porgasfi", "sforpa", "codforpa", CStr(Me.ForPago))
        If Aux = "" Then Aux = "0"
        If CCur(Aux) = 0 Then
            ComprobarRecargoFinanciero = False 'NO METO Recargo financiero
        Else
            PorceRecargo = CCur(Aux)
        End If
    End If
      
    'Si lleva recargo financiero, pero el cliente no se le aplica..
    If ComprobarRecargoFinanciero Then
        Aux = DevuelveDesdeBD(conAri, "Recargofinanciero", "sclien", "codclien", CStr(mCodClien))
        If Aux = "" Then Aux = "0"
        If Aux = "0" Then ComprobarRecargoFinanciero = False
    End If
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
    
    
    If ComprobarRecargoFinanciero Then
        Aux = Replace(cadSQL, "scaalb", "slialb")
        
        Aux = DevuelveDesdeBD(conAri, "count(*)", "slialb", Aux & " AND codartic ", vParamAplic.ArticuloRecargoFinanciero, "T")
        If Aux = "" Then Aux = "0"
        If Val(Aux) = 0 Then
            MenError = "Precálculo total factura"
            Aux = cadSQL  'Como lo va a cambiar en la funcion, la guardo antes
            b = CalcularDatosFactura(cadSQL, "scaalb", "slialb", False)
            cadSQL = Aux 'reestablezco cadsql
            If b Then
                'AQUI es donde, al ultimo albaran, le meto la linea de recargo financiero
                Aux = cadSQL & " AND 1"
                Aux = DevuelveDesdeBD(conAri, "max(numalbar)", "scaalb", Aux, "1", "N")
                If Aux = "" Then Err.Raise 513, "Error SQL: " & cadSQL
                MenError = Replace(cadSQL, "scaalb", "slialb")
                MenError = MenError & " AND 1"
                MenError = DevuelveDesdeBD(conAri, "max(numlinea)", "slialb", MenError, "1", "N")
                If MenError = "" Then MenError = "0"
                MenError = Val(MenError) + 1
                Aux = "'" & TipoAlb & "'," & Aux & "," & MenError & "," & AlmFactu & "," & DBSet(vParamAplic.ArticuloRecargoFinanciero, "T") & ","
                
                'nomartic,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre
                PorceRecargo = Round((PorceRecargo * CCur(mTotalfac2)) / 100, 2)
                MenError = DevuelveDesdeBD(conAri, "nomartic", "sartic", "codartic", vParamAplic.ArticuloRecargoFinanciero, "T")
                If MenError = "" Then MenError = "** SIN REFERENCIA **"
                Aux = Aux & DBSet(MenError, "T") & ",1,0," & TransformaComasPuntos(CStr(PorceRecargo)) & ",0,0," & TransformaComasPuntos(CStr(PorceRecargo)) & ",'M')"
                Aux = " VALUES (" & Aux
                'codtipom,numalbar,numlinea,codalmac,codartic,nomartic,cantidad,numbultos,precioar,dtoline1,dtoline2,importel
                MenError = "Insertando recargo financiero"
                Aux = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre)" & Aux
                conn.Execute Aux
            End If
            
        End If
    End If
    
    
    HacerCambioIVA = False
    If TipoAlb = "ART" Then
        'Obtengo la fecha de la factura a la que rectifica
        Aux = DevuelveDesdeBD(conAri, "fecfactu", "scaalb", cadSQL & " AND 1", "1")
        If Aux <> "" Then
            If CDate(Aux) < vParamAplic.FechaCambioIva Then HacerCambioIVA = True
        End If
    End If
    
    'Insertar la Factura (scafac,scafac1,slifac)
    If b Then
        MenError = "Error al Insertar en tablas de factura."
        b = InsertarFactura(cadSQL, , HacerCambioIVA, EsTraspasoOfeFAZ)
    End If
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    '#Laura: 14/11/2006 Recuperar facturas Alzira
    If b Then
        MenError = "Error al actualizar en Tesoreria."
        b = InsertarEnTesoreria(textoCSB, MenError)
    End If
    
    
    
    'Eliminar los Albaranes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar los albaranes."
        b = EliminarAlbaranes(cadSQL)
    End If
    
    'Si alguna linea de albaran tiene asociados num. serie actualizar
    'el campo numfactu de dichos num serie
    If b Then
        MenError = "Error al actualizar los Nº de serie."
        b = ActualizarNumSeries(cadSQL)
    End If
    
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarAlbaranesAFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarAlbaranesAFactura = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
    Espera 0.25
End Function



Public Function PasarMtosAFactura(TipCo As String, Operador As String, mes As String, nMante As String, CenCoste As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los Mantenimientos que vamos a Facturar
'Desde Mantenimientos Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String
Dim CC As String

    On Error GoTo ETraspasoFac

    PasarMtosAFactura = False
    b = False
      
    codtipom = "FAM"
    
    TipCoMan = TipCo
    OpeFactu = Operador
    MesFactu = mes
    nummante = nMante
    'CEntro de coste
    manCCost = ""
    
    
    ArticFac = DevuelveDesdeBDNew(conAri, "stipco", "codartic", "codtipco", TipCoMan, "T")
    If ArticFac = "" Then
        MenError = "El tipo de contrato no tiene articulo para facturar."
        Exit Function
    End If
      
    
    If OpeFactu <> "" Then
        'Cojo el del trabajador para no hacer el select otra vez
        manCCost = "codccost"
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N", manCCost)
    Else
        AlmFactu = "1"
    End If
    
    'CEntro de coste
    
    If vEmpresa.TieneAnalitica Then
        'Tienen analitica
        Select Case vParamAplic.ModoAnalitica
        Case 0
            'Trabajador
            'YA lo he hecho arriba
            ' manCCost = devuelve......
        
        Case 1
            'Familia
            manCCost = DevuelveDesdeBD(conAri, "codccost", "sartic,sfamia", "sartic.codfamia=sfamia.codfamia AND sartic.codartic", ArticFac, "T")
        Case Else
            'proyecto
            manCCost = CenCoste
        
        End Select
        If manCCost = "" Then
            MsgBox "Centro de coste no especificado", vbExclamation
            Exit Function
        End If
    Else
        manCCost = ""
    End If
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
    
       
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura"
    b = InsertarFactura("")
    
    
    If b Then
        'Insertar en la Tabla para detalle equipos, para reimprimir
       MenError = "Error al insertar detalle de equipos."
       InsertarDetEquip
    End If
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    If b Then
        MenError = "Error al actualizar en Tesoreria"
        b = InsertarEnTesoreria("", MenError)
    End If
    
   
   
    'Actualizo en la tabla mantenimientos la fecha ultimio mantenimiento
    If b Then
        MenError = "Actualizar fecha ultimo mantenimiento"
        ArticFac = "UPDATE scaman SET ulmesfac = " & mes & " WHERE codclien = " & Cliente & " AND nummante= '" & nummante & "'"
        conn.Execute ArticFac
    End If
   
ETraspasoFac:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Pasando Mantenimientos a Factura.", Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarMtosAFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarMtosAFactura = False
    End If
    Espera 0.2
End Function



Public Function PasarTicketAFactura(cadSQL As String, Optional ErroresAux As String, Optional NTicket As String, Optional NAlbTicket As String, Optional impRegalo As String) As Boolean
'IN -> cadSQL: cadena para seleccion de las ventas del ticket que vamos a Facturar
'Desde ventas Genera las Facturas de ticket correspondiente
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETrasTicFac

    OpeFactu = ErroresAux 'Aqui hemos pasado el trabajador
    NumTicket = NTicket
    NumAlbTicket = NAlbTicket '01/09/06 laura
    ImpCheque = CCur(ComprobarCero(impRegalo)) 'importe del cheque regalo (si hay)
    
    PasarTicketAFactura = False
    b = False
    ErroresAux = ""
      
      
    'por defecto se coge el almacen con el q trabaja el empleado
    If OpeFactu <> "" Then
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N")
    Else
        AlmFactu = "1"
    End If
    
    
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFacturaTicket(cadSQL)
    
    
    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    If b Then
        MenError = "Error al actualizar en Tesoreria."
        b = InsertarEnTesoreria("", MenError)
    End If
    
    
    'FALTA
    If UCase(App.Title) <> "ARIGES" Then
        'De momento solo para AritPV
        'Antes de eliminar hay que reducir la cantidad de los lotes
        MenError = "Ajustando cantidad restante lotes."
        AjustarCantidadLotes cadSQL
        
    End If
    
    
    'Eliminar los Albaranes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar la venta."
        b = EliminarVenta(cadSQL)
    End If
    
    
    
ETrasTicFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If Not b Then ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    PasarTicketAFactura = b
End Function

Private Sub AjustarCantidadLotes(ElTiket As String)
Dim RL As ADODB.Recordset
Dim Cadena As String

    'AQUI AQUI
    Set RL = New ADODB.Recordset
    
    Cadena = " SELECT  codartic,numlote,fecentra,sum(cantidad) unidades FROM slivenlotes "
    Cadena = Cadena & " WHERE " & Replace(ElTiket, "scaven", "slivenlotes")
    Cadena = Cadena & "  group by 1,2,2"
    RL.Open Cadena, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        If DBLet(RL!unidades, "N") <> 0 Then
            Cadena = "-"
            If RL!unidades < 0 Then Cadena = "+"
                
            Cadena = "UPDATE slotes SET vendida=vendida  " & Cadena & " " & DBSet(Abs(RL!unidades), "N")
            Cadena = Cadena & " WHERE codartic=" & DBSet(RL!codArtic, "T")
            Cadena = Cadena & " AND numlotes=" & DBSet(RL!numLote, "T")
            Cadena = Cadena & " AND fecentra=" & DBSet(RL!fecentra, "F")
            conn.Execute Cadena
        End If
        RL.MoveNext
    Wend
    RL.Close
    Set RL = Nothing
End Sub

Private Function InsertarFactura(cadSQL As String, Optional cadNumFac As String, Optional HacerCambiosIVA As Boolean, Optional EsTraspasoOfeFAZ As Boolean) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim devuelve As String

    InsertarFactura = False
    
    
    'Obtener el Contador de FACTURA
    '--------------------------------------------------------------
    Set vTipoMov = New CTiposMov
    
    
    If vTipoMov.Leer(codtipom) Then
    
        If codtipom = "FAT" Then
            'Las facturas de telefonia se corresponden con el mismo NUMero de albaran
            
            'El primer numero es el DE la factura final
            'Si tuviera mas de un albaran, los siguientes a partir de la coma serian los que
            'se introdujeron a mano en albaranes de telefonia
            
            If InStr(1, cadSQL, ",") > 0 Then
                NumFactu = Mid(cadSQL, 1, InStr(1, cadSQL, ",") - 1)
            Else
                NumFactu = cadSQL
            End If
            'aqui aqui aqui
            
            cadSQL = " scaalb.codtipom='ALT' AND scaalb.numalbar IN (" & cadSQL & ")"
        Else
                
                If EsTraspasoOfeFAZ Then
                    'Es un traspaso de una oferta a FAZ
                    
                    Dim I As Integer
                    I = InStr(1, cadSQL, " IN (")
                    If I = 0 Then Err.Raise 513, "Imposible obtener numero albaran"
                    devuelve = Mid(cadSQL, I + 5)
                    I = InStr(1, devuelve, ")")
                    If I = 0 Then Err.Raise 513, "Imposible obtener numero albaran (II)"
                    devuelve = Mid(devuelve, 1, I - 1)
                    NumFactu = CLng(devuelve)
                Else
                    'Comprobar si mientras tanto se incremento el contador de Facturas
                    'para ello vemos si existe una Factura con ese contador y si existe la incrementamos
                    Do
                        NumFactu = vTipoMov.ConseguirContador(codtipom)
                        devuelve = DevuelveDesdeBDNew(conAri, "scafac", "numfactu", "codtipom", codtipom, "T", , "numfactu", CStr(NumFactu), "N", "fecfactu", FecFactu, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (codtipom)
                            NumFactu = vTipoMov.ConseguirContador(codtipom)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                End If
    
        End If
    Else
        Exit Function
    End If
        
    LetraSerie = Trim(vTipoMov.LetraSerie) 'la guardamos para INSERT en scobro de la Conta

    'calcular el total de la factura antes de insertar
    b = CalcularDatosFactura(cadSQL, "scaalb", "slialb", HacerCambiosIVA)
    If Not b Then Exit Function
    
    '### Laura: 14/11/2006 Recuperar facturas Alzira
    devuelve = ""
    
    'Cabecera Factura
    b = InsertarCabeceraFactu
    If Not b Then Exit Function
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL)
    If Not b Then Exit Function
    
    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL)
    If Not b Then Exit Function
    
    'Comprobar si los albaranes tienen km de desplazamiento e insertar
    'lineas de factura del desplazamiento e incrementar el bruto de la factura
    'Insertar lineas de desplazamiento
    b = InsertarDesplaz
    If Not b Then Exit Function
    
    'si hay bonificaciones en la factura
    If SQLBonif <> "" Then
        b = InsertarBonificaciones
        If Not b Then Exit Function
    End If
    
   
 
    
    'Incrementar contador del tipo de movimiento
    '-------------------------------------------------------------
    'Se incrementa el contador para TODOS menos para las Fras de telefonia. El contador lo maneja la importacion del fichero
    If codtipom <> "FAT" Then
        If Not EsTraspasoOfeFAZ Then vTipoMov.IncrementarContador (codtipom)
    End If

    Set vTipoMov = Nothing
      
    InsertarFactura = True
End Function




Private Function InsertarFacturaTicket(cadSQL As String) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
'Es un ticket de venta
Dim b As Boolean
Dim devuelve As String

    InsertarFacturaTicket = False
    
    
    'Calcular los campos de la factura y cuadrarla al
    'pasar las lineas sin IVA
    b = CalcularDatosFacturaTicket(cadSQL)
    If Not b Then Exit Function

    
    'Cabecera Factura
    b = InsertarCabeceraFactu
    If Not b Then Exit Function
    
    'Cabeceras Albaranes de la Factura
    b = InsertarCabAlbaranesFactu(cadSQL)
    If Not b Then Exit Function

    'Insertar lineas de la factura
    b = InsertarLineasFactu(cadSQL)
    If Not b Then Exit Function
    
       
'    'Incrementar contador del tipo de movimiento
'    '-------------------------------------------------------------
'    vTipoMov.IncrementarContador (codTipoM)
'    Set vTipoMov = Nothing
      
    InsertarFacturaTicket = True
End Function




Private Function InsertarCabeceraFactu() As Boolean
'Inserta la cabecera de la factura en la tabla: scafac
Dim SQL As String
Dim vClien As CCliente
Dim Nulo2 As String
Dim Nulo3 As String
Dim ctaDpto As Boolean
Dim Cad As String, cad2 As String

    On Error GoTo EInsertar

    
    'Insertar en la tabla cabecera de la factura de compras
    SQL = "INSERT INTO scafac (codtipom,numfactu,fecfactu,codclien,nomclien,domclien,codpobla,pobclien,proclien,nifclien,telclien,coddirec,nomdirec,"
    SQL = SQL & "codagent,codforpa,dtoppago,dtognral,codbanco,codsucur,digcontr,cuentaba,iban,brutofac,impdtopp,impdtogr,baseimp1,baseimp2,baseimp3, "
    SQL = SQL & "codigiv1,codigiv2,codigiv3,porciva1,porciva2,porciva3,imporiv1,imporiv2,imporiv3,totalfac,intconta "
            ' Añadimos IVA recargo de equivalencia                               02 Junio 08
    SQL = SQL & ",porciva1re,imporiv1re, porciva2re,imporiv2re, porciva3re,imporiv3re,aportacion,aridoc)"
            'JUNIO 2011. NO hace falta añadir el campo FACTURAe ya que siempre que entre, entra a CERO
            
    
    
    SQL = SQL & " VALUES (" & DBSet(codtipom, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & "," & DBSet(Cliente, "N") & ","
    SQL = SQL & DBSet(NombreClien, "T") & "," & DBSet(DomicilioClien, "T", "N") & "," & DBSet(CPostal, "T", "N") & "," & DBSet(Poblacion, "T", "N") & "," & DBSet(Provincia, "T", "N") & ","
    SQL = SQL & DBSet(NIF, "T", "N") & "," & DBSet(Telefono, "T") & "," & DBSet(DirDpto, "N", "S") & "," & DBSet(NombreDirDpto, "T") & "," & DBSet(Agente, "N") & ","
    SQL = SQL & ForPago & "," & DBSet(DtoPPago, "N") & "," & DBSet(DtoGnral, "N") & ","
    
    If codtipom = "FAV" Or codtipom = "FRT" Or codtipom = "FAR" Or codtipom = "FAS" Or codtipom = "FAI" Or codtipom = "FAT" Or codtipom = "FMO" Then
        'si son facturas de VENTA o Factura Rectificativa
        'Cargar los datos del banco del cliente
        'si se trabaja con departamentos y la factura va a un departamento
        'comprobar si tiene banco y asignar el del departamento
        ctaDpto = False
        'DAVID 05/07/2010    Direccion Departamento Obra.  Agrupa <>direccion
        If vParamAplic.HayDeparNuevo > 0 Then
            If Me.DirDpto <> "" Then
                'comprobar que el dpto del cliente tiene cta bancaria
                cad2 = "digcontr"
                Cad = DevuelveDesdeBDNew(conAri, "sdirec", "cuentaba", "codclien", Me.Cliente, "N", cad2, "coddirec", Me.DirDpto, "N")
                If Cad <> "" And cad2 <> "" Then ctaDpto = True
            End If
        End If
        
        If ctaDpto Then 'cogemos la cta bancaria del departamento
            Me.CuentaBan = Cad
'            cad = DevuelveDesdeBDNew(conAri, "sdirec", "digcontr", "codclien", Me.Cliente, "N", , "coddirec", Me.DirDpto, "N")
            Me.DigControl = cad2
            cad2 = "concat(codsucur,'|',iban,'|')"
            Cad = DevuelveDesdeBDNew(conAri, "sdirec", "codbanco", "codclien", Me.Cliente, "N", cad2, "coddirec", Me.DirDpto, "N")
            
            If InStr(1, cad2, "concat(") > 0 Then
                IBAN = "ERR"
                cad2 = "0000"
                Cad = "0"
            Else
                IBAN = RecuperaValor(cad2, 2)
                cad2 = RecuperaValor(cad2, 1)
            End If
            Me.Banco = Cad
            Me.Sucursal = cad2
        Else 'cuenta bancaria del cliente
            Set vClien = New CCliente
            If vClien.LeerDatos(Cliente) Then
                Banco = vClien.Banco
                Sucursal = vClien.Sucursal
                DigControl = vClien.DigControl
                CuentaBan = vClien.CuentaBan
                IBAN = vClien.IBAN
            Else
                Banco = ValorNulo
                Sucursal = ValorNulo
                DigControl = ValorNulo
                CuentaBan = ValorNulo
                IBAN = ValorNulo
            End If
            Set vClien = Nothing
        End If
    ElseIf codtipom = "FAG" Then
        'Si la forma de pago del contador es Recibo bancario entonces pondre la cuenta
        'Si en el contador tiene la cuenta puesta, pues esa. Si no, la del cliente
        'La hemos asignado en la generacion de facturas de agua
        'Stop
    
    End If
    
    SQL = SQL & DBSet(Banco, "N") & "," & DBSet(Sucursal, "N") & "," & DBSet(DigControl, "T") & "," & DBSet(CuentaBan, "T") & "," & DBSet(IBAN, "T") & ","
    
    SQL = SQL & DBSet(BrutoFac, "N") & ","
    SQL = SQL & DBSet(ImpPPago, "N") & "," & DBSet(ImpGnral, "N") & ","
    SQL = SQL & DBSet(BaseIVA1, "N") & "," & DBSet(BaseIVA2, "N", "S") & "," & DBSet(BaseIVA3, "N", "S") & ","
    Nulo2 = "N"
    Nulo3 = "N"
    If BaseIVA2 = 0 Then Nulo2 = "S"
    If BaseIVA3 = 0 Then Nulo3 = "S"
    SQL = SQL & DBSet(TipoIVA1, "N") & "," & DBSet(TipoIVA2, "N", Nulo2) & "," & DBSet(TipoIVA3, "N", Nulo3) & ","
    SQL = SQL & DBSet(PorceIVA1, "N") & "," & DBSet(PorceIVA2, "N", Nulo2) & "," & DBSet(PorceIVA3, "N", Nulo3) & ","
     
    SQL = SQL & DBSet(ImpIVA1, "N") & "," & DBSet(ImpIVA2, "N", Nulo2) & "," & DBSet(ImpIVA3, "N", Nulo3) & ","
     
    SQL = SQL & DBSet(TotalFac, "N") & ",0"
    
    'AÑADIMOS EL recargo de equivalencia en el INSERT  porciva1re,imporiv1re .......
    SQL = SQL & "," & DBSet(PorceIVA1RE, "N", "S") & "," & DBSet(ImpIVA1RE, "N", "S")
    SQL = SQL & "," & DBSet(PorceIVA2RE, "N", "S") & "," & DBSet(ImpIVA2RE, "N", "S")
    SQL = SQL & "," & DBSet(PorceIVA3RE, "N", "S") & "," & DBSet(ImpIVA3RE, "N", "S") & ","
    
    'aportacion,aridoc
    SQL = SQL & DBSet(mAportacion, "N", "N") & ",NULL)"
    
    conn.Execute SQL
    InsertarCabeceraFactu = True
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabeceraFactu = False
        MuestraError Err.Number, "Insertar cabecera factura.", Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function


Private Function InsertarCabAlbaranesFactu(cadSQL As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: scafac1
'NumTicket: para cuando de venta generamos Factura en numalbar guardamos el ticket que la genero
Dim SQL As String
Dim devuelve As String
Dim MantenimientoObservacion As String
Dim nt As Byte
Dim RTiket As ADODB.Recordset

    On Error GoTo EInsertarAlb

    devuelve = ""
    If codtipom = "FTI" Then
'        cadSQL = Replace(cadSQL, "sliven", "scaven")
    Else
        If InStr(cadSQL, "scaven") > 0 Then 'viene de generar factura en ticket
'            codTipoM = "FTI"
        Else
            cadSQL = Replace(cadSQL, "slialb", "scaalb")
        End If
    End If
    '
    SQL = "INSERT INTO scafac1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,"
    SQL = SQL & "referenc,codenvio,codtraba,codtrab1,codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa,actuacion,tipoimp ,"
    SQL = SQL & "origdat,fecenvio,docarchiv,coddiren,SerieTfno, ManipuladorNumCarnet , ManipuladorFecCaducidad , ManipuladorNombre,TipoCarnet,PideCliente)"
    
    
    'Octubre 2013
    'FAT
    'sementre ,numofert,SerieTfno  --> Luego  SerieTfno hay que añadirlo a todos
    
    
    If (codtipom = "FAV" And NumTicket <> "") Then
        'Facturas generadas desde una venta de tickets: en numalbar guardamos numticket
'        SQL = SQL & " VALUES('" & codTipoM & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "','FTI'," & NumTicket & ","
        '01/09/06 laura
        SQL = SQL & " VALUES('" & codtipom & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "','ATI'," & NumTicket & ","
        
        SQL = SQL & "'" & Format(FecFactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codenvio", "codclien", Cliente, "N")
        SQL = SQL & DBSet(devuelve, "N") & ","
        'trabajadores
        SQL = SQL & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        
        'observaciones
        '---- Laura 10/10/06: poner en la observacion 1 el valor de scaven.observa1
        Set RTiket = New ADODB.Recordset
        
        'devuelve = DevuelveDesdeBDNew(conAri, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.FecFactu, "F")
        
        devuelve = " FROM scaven WHERE numtermi= " & Me.NumTerminal & " AND numventa = " & Me.NumVenta & " AND fecventa = " & DBSet(Me.FecFactu, "F")
        devuelve = "Select observa1, ManipuladorNumCarnet , ManipuladorFecCaducidad, ManipuladorNombre, TipoCarnet " & devuelve
        RTiket.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        devuelve = ""
        If Not RTiket.EOF Then devuelve = DBLet(RTiket!observa1, "T")
        
        SQL = SQL & DBSet(devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        

        
        SQL = SQL & NumTerminal & "," & NumVenta
        
        'SAIL
        SQL = SQL & ",NULL,NULL,NULL"  'actuacion,tipoimp ,origdat
        'fecenvio,docarchiv,coddiren  SerieTfno
        SQL = SQL & ",NULL,0,NULL,NULL,"
        
        ' ManipuladorNumCarnet , ManipuladorFecCaducidad, ManipuladorNombre, TipoCarnet
        devuelve = "NULL,NULL,NULL,NULL"
        If Not RTiket.EOF Then
            If DBLet(RTiket!ManipuladorNumCarnet, "T") <> "" Then
                devuelve = DBSet(RTiket!ManipuladorNumCarnet, "T", "S") & "," & DBSet(RTiket!ManipuladorFecCaducidad, "F", "S")
                devuelve = devuelve & "," & DBSet(RTiket!ManipuladorNombre, "F", "S") & "," & DBSet(RTiket!TipoCarnet, "N")
            End If
        End If
        SQL = SQL & devuelve & ",0)"
        RTiket.Close
    ElseIf (codtipom <> "FAM" And codtipom <> "FTI") And cadSQL <> "" Then 'Facturas de VEnta
        SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu,"
        SQL = SQL & " codtipom as codtipoa, numalbar,fechaalb,numpedcl,fecpedcl,sementre,NumOfert , fecofert, referenc, CodEnvio, "
        SQL = SQL & " codtraba,codtrab1,codtrab2,observa01,observa02,observa03,observa04,observa05,numtermi,numventa,actuacion,tipoimp"
        SQL = SQL & " ,origdat,fecenvio,docarchiv,coddiren, "
        'OCTUBRE 2013
        'Letraser . Variable mLetraSer. Albara de telefonoia 'ALT'
        SQL = SQL & " if(codtipom='ALT','" & mLetraSer & "','') Letraser "
        
        
        'Manipulador
        SQL = SQL & " , ManipuladorNumCarnet , ManipuladorFecCaducidad , ManipuladorNombre,TipoCarnet,PideCliente "
        
        SQL = SQL & " FROM scaalb WHERE " & cadSQL
        
        If codtipom = "FAT" Then SQL = SQL & " ORDER BY sementre desc" 'para que el primer albaran sea el de la factura de telefono
        
    Else 'Bonificaciones/Desplazamientos
        'Facturas de Mantenimiento
        'Facturas de Ticket del TPV
        SQL = SQL & " VALUES('" & codtipom & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "',"
'        If (codTipoM = "FAM" Or codTipoM = "FTI") Then
        '01/09/06 laura
        If (codtipom = "FAM") Then
            SQL = SQL & "'',0,"
        ElseIf codtipom = "FTI" Then
            SQL = SQL & "'ATI'," & NumAlbTicket & ","
        Else
            SQL = SQL & "'',9999999,"
        End If
        SQL = SQL & "'" & Format(FecFactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "codenvio", "codclien", Cliente, "N")
        SQL = SQL & DBSet(devuelve, "N") & ","
        'trabajadores
        SQL = SQL & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        'observaciones 1
        If codtipom = "FTI" Then
            '---- Laura 19/12/06: poner en la observacion 1 el valor de scaven.observa1
            devuelve = DevuelveDesdeBDNew(conAri, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.FecFactu, "F")
        Else
            'La linea viene en el campo concefac que hemos guardado en observaciones
            If codtipom = "FAM" Then
                If nummante = "DBZ" Then
                    'ESTAMOS EN RENTING
                    
                     devuelve = ""
                     
                     
                     
                     
                     
                     
                     
                Else
                    MantenimientoObservacion = DevuelveDesdeBDNew(conAri, "scaman", "tipopago", "codclien", Me.Cliente, "N", , "nummante", nummante, "T")
                    If MantenimientoObservacion = "1" Then 'Trimestral
                        'Manolo dice que es el trimestre del mes que facturo
                        nt = ((CInt(MesFactu) - 1) \ 3) + 1
                        If nt = 1 Or nt = 3 Then
                            MantenimientoObservacion = nt & "er"
                        Else
                            MantenimientoObservacion = nt & "º"
                        End If
                        MantenimientoObservacion = MantenimientoObservacion & " TRIMESTRE DE " & Year(FecFactu)
                        
                        
                    ElseIf MantenimientoObservacion = "2" Then 'Semestral
                        nt = ((CInt(MesFactu) - 1) \ 6) + 1
                        If nt = 1 Then
                            MantenimientoObservacion = "1er"
                        Else
                            MantenimientoObservacion = "2º"
                        End If
                        MantenimientoObservacion = MantenimientoObservacion & " SEMESTRE DE " & Year(FecFactu)
                        
                        
                    'Noviembre 2013
                    ElseIf MantenimientoObservacion = "4" Then 'Bimensual
                        MantenimientoObservacion = "BIMENSUAL "
                        
                        
                    ElseIf MantenimientoObservacion = "3" Then 'Anual
                        MantenimientoObservacion = "ANUAL " & Year(FecFactu)
                    Else
                        'si es mensual: ej: "JUNIO de 2005" (ponemos el mes seleccionado para facturar)
                        MantenimientoObservacion = MonthName(CLng(MesFactu)) & " de " & Year(FecFactu)
                        MantenimientoObservacion = UCase(MantenimientoObservacion)
                    End If
                                                                
                    If mObservacion <> "" Then
                        devuelve = MantenimientoObservacion
                        
                    Else
                        devuelve = ""
                        mObservacion = MantenimientoObservacion
                    End If
                End If
            Else
                devuelve = ""
            End If
           
        End If
        SQL = SQL & DBSet(devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
        '                                                                                   sail + fecenvio  + docarchiv + coddiren + letratf
        SQL = SQL & "," & DBSet(NumTerminal, "N", "S") & "," & DBSet(NumVenta, "N", "S") & ",NULL,NULL,NULL,NULL,0,NULL,NULL"
        ' ManipuladorNumCarnet , ManipuladorFecCaducidad , ManipuladorNombre,TipoCarnet
        SQL = SQL & ",NULL,NULL,NULL,NULL,0)"
    End If
    
    conn.Execute SQL




    'SOLO PARA EULER
    '.----------------------------------
    If vParamAplic.NumeroInstalacion = 4 Then
        'METERA EN UNA TABLA que es scafac_EU
        If codtipom = "FAO" Or codtipom = "FAE" Or codtipom = "FAR" Then
            devuelve = "bombamarca,bombaModelo,motormarca,motorModelo,observaciones,TrabajoExterior,TipoPortes,ReferPedido,FechaPed"
            
            'Febrero 2016
            devuelve = devuelve & ",Rep_OrdenTrabajo,Rep_TrabajoExterior,RecepAgenClien,RecepPortes,RecepAgenCliMat,RecpNumExp,FechaAlb,TipoBombResSuperHor"
            devuelve = devuelve & ",TipoBombResSuperVer,TipoBombLimSuperHor,TipoBombLimSuperVer,TipoBombResSumPoz,TipoBombLimSumPoz,TipoBombResSumVer"
            devuelve = devuelve & ",TipoBombLimSumVer,TipoBomAgitadorRes,TipoBomAgitadorLim,TipoBomResOtrosEqu,TipoBomLimOtrosEqu,DatosBommarca,DatosBomNumCurva"
            devuelve = devuelve & ",DatosBomModelo,DatosBomNumSerie,DatosBomAno,DatosBomH,DatosBomTipoRodete,DatosBomCaudal,DatosBomUdCaudal,DatosMotorMarca"
            devuelve = devuelve & ",DatosMotorModelo , DatosMotorNumSerie, DatosMotorV, DatosMotorI, DatosMotorCV, DatosMotorKw, DatosMotorrpm, NumParteTrabajo, NumTrabajExterno"
            
            
            SQL = "INSERT INTO scafac_eu(codtipom,numfactu,fecfactu,codtipoa,numalbar," & devuelve & ") "
            SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu,"
            SQL = SQL & " codtipom as codtipoa, numalbar," & devuelve
            SQL = SQL & " FROM scaalb_eu scaalb WHERE " & cadSQL
            If Not ejecutar(SQL, True) Then
                devuelve = SQL & vbCrLf & vbCrLf & vbCrLf & "El programa continuará. Avise soporte técnico"
                MsgBox devuelve, vbCritical
            End If
    
        End If
    End If


    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        MuestraError Err.Number, "Insertar cabecera albaranes factura.", Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
    Set RTiket = Nothing
End Function


Private Function InsertarLineasFactu(cadSQL As String) As Boolean
'Insertar las lineas de los Albaranes de la factura en la tabla: slifac
Dim SQL As String
Dim Cad As String
Dim cad2 As String
Dim cart As CArticulo
Dim Observas As Boolean   'DE SAIL. Tabla slialt
Dim RL As ADODB.Recordset
Dim Aux2 As String

    On Error GoTo EInsertarLin

    '##########################################
    'Cambios para llevar la TRAZA con codprove
    '   si cojo slialb:  tengo la columna



    'Otra modificacion:   16 Mayo 08
    ' Se permiten dtos, con lo cual la donde ponia 0 as dto ya no vale
    
    'Modifo.  30/08/2010,  Entra SAIL
    SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre"
    If NumTicket <> "" Then SQL = SQL & ",precioiv"
    SQL = SQL & ",preciomp,preciost,preciouc,codprovex,numlote,codccost,"     'ENERO 2008.  Traza sobre codprove.
                                                                            '[19/10/2009]: añadir centro coste
    
    'SAIL
    '----- codtipor codcapit precoste codtraba       ''pvpInferior --> Herbelca
    SQL = SQL & "codtipor, codcapit, precoste, codtraba,pvpInferior,comisionagente)"
    Observas = False
    If codtipom = "FAM" Then 'facturas de mantenimientos
    
        If nummante = "DBZ" Then
            'FACTURACION RENTING
           
            NumAlbTicket = "Select sclienrenting.*,TipoFraRenting FROM sclienrenting,sclien " & NumAlbTicket
            If InStr(1, NumAlbTicket, " codclien") = 0 Then
                'Facturacion cliente por cliente
                NumAlbTicket = NumAlbTicket & " AND sclienrenting.codclien = " & Cliente
            End If
            Set RL = New ADODB.Recordset
    
            RL.Open NumAlbTicket, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumAlbTicket = ""
            mNumventa = 0
            While Not RL.EOF
            
                
                    mLetraSer = ""
                    ArticFac = DevuelveDesdeBDNew(conAri, "stipco", "codartic", "codtipco", RL!codtipco, "T")
                    mNumventa = mNumventa + 1
                        
                    'nombre del articulo para facturar,preciomp,preciost, preciouc
                    If ArticFac <> "" Then
                        Set cart = New CArticulo
                        If cart.LeerDatos(ArticFac) Then
                            Cad = cart.Nombre   'Aunque luego machaco cad con la nueva amplicacion: contrato + numerocontrato
                        End If
            
                    End If
                   
                    'Como "NOMARTIC", vamos a poner Referencia: referenciarenting
                    
                    If RL!TipoFraRenting = 3 Then
                        Cad = " Trimestre: " & ((Month(FechaFacturarRenting) - 1) \ 3) + 1 & "º"
                      
                    ElseIf RL!TipoFraRenting = 6 Then
                        Cad = " Semestre: " & ((Month(FechaFacturarRenting) - 1) \ 6) + 1 & "º"
                    ElseIf RL!TipoFraRenting = 12 Then
                        Cad = " Año: " & Year(FechaFacturarRenting)
                    Else
                        Cad = " MES: " & Format(FechaFacturarRenting, "mm/yyyy")
                    End If
                    
                    
                    'Teinsa lo dejo como esta
                    
                    If vParamAplic.NumeroInstalacion = 3 Then
                        Cad = "Referencia: " & RL!Referencia & Cad
                        Cad = DBSet(Cad, "T")
                    Else
                        'aPara el resto.  De momento nosotros
                        Aux2 = Trim(DBLet(RL!obser, "T"))
                        
                        
                        If Aux2 <> "" Then
                            NombreSQL Aux2
                    
                            Aux2 = Aux2 & "\r\n"
                        End If
                        
                        Cad = Aux2 & "Periodo " & Cad
                        Cad = "'" & Cad & "'"
                        
                    
                    End If
                  
                  
                    'Ampliacion=aux2
                    Aux2 = Cad
                    Cad = DBSet(cart.Nombre, "T")
                    

                    Cad = Cad & "," & Aux2
                    
                    NumAlbTicket = NumAlbTicket & ", ('" & codtipom & "'," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','',0," & mNumventa & ","
                
                    NumAlbTicket = NumAlbTicket & DBSet(AlmFactu, "N") & "," & DBSet(cart.codigo, "T") & "," & Cad & ","
                        
                    'cantidad,numbultos, precio, dtos
                    If InStr(1, RL.Source, " AND id =") Then
                        'Va una unica linea, por lo tanto, el importe es el de la factura
                        NumAlbTicket = NumAlbTicket & "1,1," & DBSet(BrutoFac, "N") & ", 0,0, "
                        'importe, origen precio
                        NumAlbTicket = NumAlbTicket & DBSet(BrutoFac, "N") & ", 'M',"
                    Else
                        'Es de cada renting de los que facturamos
        
                        NumAlbTicket = NumAlbTicket & RL!TipoFraRenting & ",1," & DBSet(RL!Importe, "N") & ", 0,0, "
                        'importe, origen precio
                        NumAlbTicket = NumAlbTicket & DBSet(RL!TipoFraRenting * RL!Importe, "N") & ", 'M',"
                    End If
                    'guardar los precios del articulo
                    NumAlbTicket = NumAlbTicket & DBSet(cart.PrecioMedPon, "N") & "," & DBSet(cart.PrecioStan, "N") & "," & DBSet(cart.PrecioUltCom, "N") & ","
                    'traza ENERO 2008                    'NUmero lote Junio2009
                    NumAlbTicket = NumAlbTicket & DBSet(cart.Codprove, "N") & ",NULL,"
                    'CEntro de coste
                    NumAlbTicket = NumAlbTicket & DBSet(manCCost, "T", "S")
                    'SAIL codtipor codcapit precoste codtraba ,pvpInferior,comisionagente
                    NumAlbTicket = NumAlbTicket & ",NULL,NULL,NULL,NULL,0,0"
                    NumAlbTicket = NumAlbTicket & ")"
                    Set cart = Nothing
    
    
                    RL.MoveNext
                Wend
                RL.Close
                NumAlbTicket = Mid(NumAlbTicket, 2)
                SQL = SQL & " VALUES " & NumAlbTicket
                mLetraSer = ""
                NumAlbTicket = ""
                mNumventa = 0
        Else
            'FRAS MANTENIMIENTO
            SQL = SQL & " VALUES('" & codtipom & "'," & NumFactu & ",'" & Format(FecFactu, FormatoFecha) & "','',0,1,"
                
            SQL = SQL & DBSet(AlmFactu, "N") & ","
                
            'Articulo para facturar mantenimientos de la tabla de tipos contrato: stipco.codartic
            SQL = SQL & DBSet(ArticFac, "T") & ", "
                
            'nombre del articulo para facturar,preciomp,preciost, preciouc
            If ArticFac <> "" Then
                Set cart = New CArticulo
                If cart.LeerDatos(ArticFac) Then
                    Cad = cart.Nombre   'Aunque luego machaco cad con la nueva amplicacion: contrato + numerocontrato
                End If
    '            cad2 = "preciomp"
    '            cad = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", ArticFac, "T", cad2)
            End If
            
            Cad = DevuelveDesdeBDNew(conAri, "stipco", "nomtipco", "codtipco", TipCoMan, "T")
            Cad = Cad & "  Nº: " & nummante
            
            SQL = SQL & DBSet(Cad, "T") & ","
                
    
            Cad = mObservacion
            
            SQL = SQL & DBSet(Cad, "T") & ","
                
            'cantidad,numbultos, precio, dtos
            SQL = SQL & "1,1," & DBSet(BrutoFac, "N") & ", 0,0, "
            'importe, origen precio
            SQL = SQL & DBSet(BrutoFac, "N") & ", 'M',"
            
            'guardar los precios del articulo
            SQL = SQL & DBSet(cart.PrecioMedPon, "N") & "," & DBSet(cart.PrecioStan, "N") & "," & DBSet(cart.PrecioUltCom, "N") & ","
            'traza ENERO 2008                    'NUmero lote Junio2009
            SQL = SQL & DBSet(cart.Codprove, "N") & ",NULL,"
            'CEntro de coste
            SQL = SQL & DBSet(manCCost, "T", "S")
            'SAIL codtipor codcapit precoste codtraba   pvpinf comisionagente
            SQL = SQL & ",NULL,NULL,NULL,NULL,0,0"
            SQL = SQL & ")"
            Set cart = Nothing
        End If
    ElseIf (codtipom = "FTI") Then 'factura de ticket
        'se entra en esta opcion cuando desde el TPV generamos un ticket
        SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu, 'ATI' as codtipoa," & NumAlbTicket & " as numalbar," & "numlinea,"
        SQL = SQL & AlmFactu & " as codalmac," & "sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,1 as numbultos,precioar,"
        'David 16 Mayo 08
        '0 as dtoline1,0 as dtoline2, cantidad*precioar as importel,'' as origpre,precioiv "
        SQL = SQL & " dto1,dto2, implineareal,'' as origpre,precioiv "
        
         '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & ", sartic.preciomp,sartic.preciost,sartic.preciouc,sartic.codprove,NULL as numlote"
        '-- [19/10/2009] [LAURA] : añadir centro de coste
        SQL = SQL & ",codccost"
        'SAIL codtipor codcapit precoste codtraba ,ppvpinferior,comisionagente
        SQL = SQL & ",NULL,NULL,NULL,NULL,0,0"
        '--
        SQL = SQL & " FROM sliven "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & " INNER JOIN sartic ON sliven.codartic=sartic.codartic "
        '--
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven")
        SQL = SQL & " ORDER BY numlinea"
        
    ElseIf NumTicket <> "" Then 'Ticket de venta genera factura
        SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu, 'ATI' as codtipoa," & NumTicket & " as numalbar," & "numlinea,"
        SQL = SQL & AlmFactu & " as codalmac," & "sliven.codartic,sliven.nomartic," & ValorNulo & " as ampliaci,cantidad,1 as numbultos,precioar,"
        
        'David 16 Mayo 08
        '0 as dtoline1,0 as dtoline2, cantidad*precioar as importel,'' as origpre,precioiv "
        SQL = SQL & " dto1,dto2, implineareal,'' as origpre,precioiv "
        
        
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & ", sartic.preciomp,sartic.preciost,sartic.preciouc,sartic.codprove,"
        
        'Enero 2013. Meto el numeseri en la venta del tpv
        'SQL = SQL & "NULL as numlote"
        SQL = SQL & "numserie as numlote"
        
        '-- [19/10/2009] [LAURA] : añadir centro de coste
        SQL = SQL & ",codccost"
        'SAIL codtipor codcapit precoste codtraba pvp inf comisionagente
        SQL = SQL & ",NULL,NULL,NULL,NULL,0,0"
        '--
        SQL = SQL & " FROM sliven "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & " INNER JOIN sartic ON sliven.codartic=sartic.codartic "
        '--
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    
    Else 'pasamos lineas de albaranes a lineas factura
        SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,numlinea,codalmac,slialb.codartic,slialb.nomartic,ampliaci,cantidad,numbultos,precioar,dtoline1,dtoline2,importel,origpre "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & ", sartic.preciomp,sartic.preciost,sartic.preciouc,codprovex,numlote"   'ENERO 2008. Traza cond codprove
        '--
        ' -- [19/10/2009] [LAURA]: se añade el centro de coste
        SQL = SQL & ",codccost"
        
        'SAIL
        SQL = SQL & ",codtipor, codcapit, precoste, codtraba,pvpInferior,comisionagente"
            
        SQL = SQL & " FROM slialb "
        '-- Laura 15/09/2006: añadir precios del articulo a las lineas de factura
        SQL = SQL & " INNER JOIN sartic ON slialb.codartic=sartic.codartic "
        '--
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
        Observas = True
    End If
    
    conn.Execute SQL




    'Para SAIL
    If vParamAplic.TipoFormularioClientes = 1 And Observas Then
        'Para SAIL
        SQL = "INSERT INTO slifaT (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,observa)"
        SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa, "
        SQL = SQL & "numalbar,numlinea,observa "
        SQL = SQL & " FROM slialt "
        '--
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialt")
        conn.Execute SQL
    End If


    If vParamAplic.Ariagro <> "" Then
        'Veremos si tiene datos
        Aux2 = "T"
        If codtipom = "FAV" Then
            If NumTicket = "" Then Aux2 = ""
        Else
            'Aux2 = ""   'Ene 2013. Para todos los demas """
            'If codtipom = "FAI" Or codtipom = "FAS" Or codtipom = "FRT" Or codtipom = "FAM" Then Aux2 = ""
            If codtipom <> "FTI" Then Aux2 = ""
        End If
        
        
        
        If Aux2 = "T" Then
            'Viene  de tiket, o bien por factura o bien por tiket
            SQL = "INSERT INTO slifaccampos(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codcampo,fechahora,nomvarie , nompartida)"
            SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu,'ATI' as codtipoa,"
            If NumTicket <> "" Then
                SQL = SQL & NumTicket
            Else
                SQL = SQL & NumAlbTicket
            End If
            SQL = SQL & " as numalbar,numlinea,codcampo,Now(),nomvarie , nompartida "
            SQL = SQL & " FROM sliven2 "
            '--
            SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven2")
            conn.Execute SQL
        
        Else
            SQL = "INSERT INTO slifaccampos(codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codcampo,fechahora,nomvarie , nompartida)"
            SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa, "
            SQL = SQL & "numalbar,numlinea,codcampo,Now(),nomvarie , nompartida "
            SQL = SQL & " FROM slialbcampos "
            '--
            SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialbcampos")
            
            'Los mantenimientos NO llevan campos... de momento
            If codtipom <> "FAM" Then conn.Execute SQL
        End If
        
        
        If vParamAplic.ManipuladorFitosanitarios2 Then
            If codtipom <> "FAM" Then
                'Los mantenimientos NO llevan fitos
                SQL = "INSERT INTO slifaclotes( codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic)"
                
                If NumTicket <> "" Then
                    'Viene de ticket
            
                    'LOTES fitosanitarios
                    '
                    SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu, 'ATI' as codtipoa," & NumTicket & " as numalbar," & "numlinea,"
                    SQL = SQL & "sublinea,cantidad,numlote,fecentra,codartic"
                    SQL = SQL & " FROM slivenlotes   WHERE "
                    SQL = SQL & Replace(cadSQL, "scaven", "slivenlotes")
      
                
                
                Else
                    SQL = SQL & " SELECT '" & codtipom & "' as codtipom," & DBSet(NumFactu, "N") & " as numfactu," & DBSet(FecFactu, "F") & " as fecfactu," & "codtipom as codtipoa, "
                    SQL = SQL & "numalbar,numlinea,sublinea,cantidad,numlote,fecentra,codartic"
                    SQL = SQL & " FROM slialblotes "
                    '--
                    SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialblotes")
                End If
                
                conn.Execute SQL
            End If
        End If
    End If

    InsertarLineasFactu = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactu = False
        MuestraError Err.Number, "Insertar lineas factura.", Err.Description
    Else
        InsertarLineasFactu = True
    End If
End Function


Private Function EliminarAlbaranes(cadSQL As String) As Boolean
'Eliminamos de las tablas de Albaranes: scaalb, slialb
Dim SQL As String

    On Error GoTo EEliminar

    EliminarAlbaranes = False
    
    'Para SAIL borramos tb las observaciones de la linea
    If vParamAplic.TipoFormularioClientes = 1 Then
        'Para SAIL
        SQL = "DELETE FROM slialt "
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialt")
        conn.Execute SQL
        
        If vParamAplic.NumeroInstalacion = 4 Then
            SQL = "DELETE FROM scaalb_eu "
            SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "scaalb_eu")
            conn.Execute SQL
        
        End If
    End If
    
    'Para empresas con ariagro
    If vParamAplic.Ariagro <> "" Then
        'Para SAIL
        SQL = "DELETE FROM slialbcampos "
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialbcampos")
        conn.Execute SQL
    End If
    
    'Para empresas con ariagro
    If vParamAplic.ManipuladorFitosanitarios2 Then
        'Para SAIL
        SQL = "DELETE FROM slialblotes "
        SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialblotes")
        conn.Execute SQL
    End If
    
    
    
    'ELiminar lineas albaranes
    SQL = "DELETE FROM slialb "
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "slialb")
    conn.Execute SQL
    
    Espera 0.1
    
    'Eliminar Cabeceras Albaranes
    SQL = "DELETE FROM scaalb "
    SQL = SQL & " WHERE " & Replace(cadSQL, "slialb", "scaalb")
    conn.Execute SQL
        
    EliminarAlbaranes = True

EEliminar:
    If Err.Number <> 0 Then
        EliminarAlbaranes = False
    Else
        EliminarAlbaranes = True
    End If
End Function



Private Function InsertarDetEquip() As Boolean
Dim SQL As String
Dim Cad As String

    'si son facturas de MAntenimiento y parametro de stipco: detallar equipos en la factura=true
    'guardar en una tabla el resumen por tipo de articulo de los Nº Series de equipos
    'asociados al MAntenimiento
    InsertarDetEquip = False
    
    On Error GoTo EDetalle
       
    SQL = DevuelveDesdeBDNew(conAri, "stipco", "detequip", "codtipco", TipCoMan, "T")
    
    If SQL = "1" Then 'Detalla Equipo
        Cad = ObtenerSQLcomponentes(" WHERE codclien=" & Cliente & " AND nummante=" & DBSet(nummante, "T"))
        SQL = "Select distinct '" & codtipom & "' as codtipom," & NumFactu & " as numfactu,'" & Format(FecFactu, FormatoFecha) & "' as fecfactu,"
        Cad = Replace(Cad, "Select distinct", SQL)
        SQL = "INSERT INTO srecom (codtipom,numfactu,fecfactu,codtipar,nomtipar,cantidad) "
        SQL = SQL & Cad

        conn.Execute SQL
    End If
    
EDetalle:
    If Err.Number <> 0 Then
        InsertarDetEquip = False
    Else
        InsertarDetEquip = True
    End If
End Function


'ANTES:  ClienteExentoIVA  as boolean
'Ahora:  ClienteExentoIVA2 as byte
'           0.- Exento
'           1.- RE
'           2.- EXNTO
'           3.-  exento intracom
'           4.- Reducido
Private Function ClienteExentoIVA() As Byte
Dim vClien As CCliente

    Set vClien = New CCliente
    vClien.codigo = Cliente
    ClienteExentoIVA = vClien.exentoIVA
    Set vClien = Nothing
    
    
End Function



Private Function InsertarDesplaz() As Boolean
'Si los albaranes a facturar tienen KM de desplazamiento añadirlos a la factura
'Añadir cantkm * preciokm al bruto de la factura
'Insertar linea en tabla slifac con el importe del desplazamiento
Dim SQL As String
Dim Cad As String
Dim b As Boolean

    On Error GoTo EDesplaz
    InsertarDesplaz = False
    
    If TotalKm > 0 Then
        'Insertar en scafac1 un albaran a 9999999, para el desplazamiento (y servira tb para bonificaciones)
        '-----------------------------------------------------------
        b = InsertarCabAlbaranesFactu("")
        If Not b Then Exit Function
            
               
        'Insertar linea de Desplazamiento en tabla: slifac (lineas factura)
        '------------------------------------------------------------------
        SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        SQL = SQL & "cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
        'codtipom,numfactu,fecfactu,codtipoa,numalbar
        SQL = SQL & " VALUES (" & DBSet(codtipom, "T") & "," & NumFactu & "," & DBSet(FecFactu, "F") & ",'',9999999,"
        'numlinea
        SQL = SQL & "1,"
        SQL = SQL & DBSet(AlmFactu, "N") & ","
            
        'codartic,nomartic,ampliaci
        Cad = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", codArtic, "T")
        SQL = SQL & DBSet(codArtic, "T") & "," & DBSet(Cad, "T") & "," & DBSet(Ampliacion, "T") & ","
        'cantidad,precioar,dtoline1,dtoline2,importel,origpre
        SQL = SQL & DBSet(TotalKm, "N") & "," & DBSet(PrecioKm, "N") & ",0,0," & DBSet(ImporteL, "N") & ",'M')"
        
        conn.Execute SQL
        
    End If
    InsertarDesplaz = True
     
EDesplaz:
    If Err.Number <> 0 Then
        InsertarDesplaz = False
    Else
        InsertarDesplaz = True
    End If
End Function


Private Function CalcularDesplazamiento(cadWhere As String, NomTabla As String) As Boolean
'comprobar si los albaranes a facturar tienen DESPLAZAMIENTOS
'calcula en vbles: Totalkm, Ampliacion los valores para insertar en linea factura
Dim Rs As ADODB.Recordset
Dim SQL As String
    
    CalcularDesplazamiento = False
    On Error GoTo EDesp
    
    If codtipom = "FAV" Or codtipom = "FAR" Then
        'Obtenemos el total de km a facturar en los albaranes seleccionados
        SQL = "SELECT fechaalb,sum(cantidkm)" ',sum(sclien.kilometr) "
        SQL = SQL & " FROM " & NomTabla '& " INNER JOIN sclien ON " & NomTabla & ".codclien=sclien.codclien "
        SQL = SQL & " WHERE " & cadWhere '& " AND facturkm=1 "
        SQL = SQL & " GROUP BY fechaalb"
        SQL = SQL & " ORDER BY fechaalb"
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        TotalKm = 0
        Ampliacion = ""
        While Not Rs.EOF
            If Rs.Fields(1).Value <> 0 Then
                TotalKm = TotalKm + Rs.Fields(1).Value
                Ampliacion = Ampliacion & Day(Rs.Fields(0).Value) & ", "
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        If Ampliacion <> "" Then 'quitar la ultima coma
            Ampliacion = Mid(Ampliacion, 1, Len(Ampliacion) - 2)
            Ampliacion = "DIAS: " & Ampliacion
        End If
        
        If TotalKm > 0 Then
            'obtenemos el precio del km de los parametros, y el cod Articulo
            codArtic = "codartid"
            SQL = DevuelveDesdeBDNew(conAri, "spara1", "preukmcl", "codigo", "1", "N", codArtic)
            PrecioKm = CCur(ComprobarCero(SQL))
            ImporteL = Round(TotalKm * PrecioKm, 2)
            
            'obtenemos el tipo IVA del articulo de desplazamiento
            SQL = DevuelveDesdeBDNew(conAri, "sartic", "codigiva", "codartic", codArtic, "T")
            despIVA = CInt(ComprobarCero(SQL))
        End If
    End If
    CalcularDesplazamiento = True
    
EDesp:
    If Err.Number <> 0 Then
        CalcularDesplazamiento = False
    Else
        CalcularDesplazamiento = True
    End If
End Function



Private Function CalcularBonificacion(cadWhere As String, NomTablaLin As String) As Boolean
'comprobar los articulos de los albaranes a facturar tienen BONIFICACION
Dim Rs As ADODB.Recordset
Dim RSaux As ADODB.Recordset
Dim SQL As String, cadIVA As String
Dim vClien As CCliente
Dim EsExcluyente As Boolean
Dim Bonif1 As Currency, Bonif2 As Currency
Dim Canti1 As Single, Canti2 As Single
Dim numlinea As Integer

    CalcularBonificacion = False
    On Error GoTo EBonif
    
    If codtipom = "FAV" Then 'Factura de Venta
        'Obtener el cod. tarifa del cliente al que se va a facturar y
        'comprobar si para este tipo de tarifa se aplica Bonificacion
        Set vClien = New CCliente
        vClien.codigo = Cliente
        SQLBonif = ""
        TotBonif1 = 0
        TotBonif2 = 0
        TotBonif3 = 0
        IVABonif1 = 0
        IVABonif2 = 0
        IVABonif3 = 0
            
        If vClien.Bonifica Then
            'Comprobar para los codartic de las lineas de albaran que vamos
            'a factura, para cuales hay bonificacion y la cantidad de la linea
            'cumple la cantidad para que se le aplica la bonificacion
            SQL = "SELECT slialb.codartic, sum(slialb.cantidad) as cantidad "
            SQL = SQL & "FROM " & NomTablaLin
            SQL = SQL & " INNER JOIN sbonif ON " & NomTablaLin & ".codartic=sbonif.codartic "
            SQL = SQL & " WHERE " & Replace(cadWhere, "scaalb", NomTablaLin)
            SQL = SQL & " GROUP BY " & NomTablaLin & ".codartic "
            SQL = SQL & " order by " & NomTablaLin & ".codartic "
            
           
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            numlinea = 1
            
            While Not Rs.EOF 'para cada codartic
                'Para cada linea de slialb a la que se le aplica bonificacion
                'obtener la bonif. que se le va a aplicar

                'mirar primero si hay algun hastaca2<=cantidad
                '--------------------------------------------------
                SQL = "SELECT * from sbonif "
                SQL = SQL & " WHERE codartic=" & DBSet(Rs!codArtic, "T") '& " AND (" & RS!Cantidad & "<=sbonif.hastaca1)"
                SQL = SQL & " order by hastaca1 desc "
                Set RSaux = New ADODB.Recordset
                RSaux.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                Bonif1 = 0
                Bonif2 = 0
                If Not RSaux.EOF Then
'                    If (1 <= Rs!Cantidad) And (Rs!Cantidad <= RSaux!hastaca1) Then
                    If (0 < Rs!cantidad) And (Rs!cantidad <= RSaux!hastaca1) Then
                        Bonif1 = RSaux!impboni1
                        Canti1 = Rs!cantidad
                    End If
                    If (RSaux!hastaca1 < Rs!cantidad) And (Rs!cantidad <= RSaux!hastaca2) Then
                        Bonif1 = RSaux!impboni1
                        Canti1 = Rs!cantidad
                        Bonif2 = RSaux!impboni2
                        Canti2 = Rs!cantidad
                    End If
                    
                    
                    'comprobar si la bonificacion es excluyente
                    EsExcluyente = CBool(RSaux!excluyen)
                    
                    If EsExcluyente Then
                    'se aplicara una bonificacion u otra
                       If Bonif2 > 0 Then Bonif1 = 0
                    End If
                    
                    If Bonif1 > 0 Then
                        numlinea = numlinea + 1
                        SQLBonif = SQLBonif & "('" & codtipom & "'," & NumFactu & "," & DBSet(FecFactu, "F") & ",'',9999999," & numlinea & "," & AlmFactu & "," & DBSet(RSaux!codarti1, "T") & ","
                        cadIVA = "codigiva"
                        SQL = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", RSaux!codarti1, "T", cadIVA)
                        SQLBonif = SQLBonif & DBSet(SQL, "T") & ",''," & DBSet(Canti1, "N") & "," & DBSet(-Bonif1, "N") & ",0,0," & DBSet(0 - Round(Canti1 * Bonif1, 2), "N") & ",'M'),"
                        If IVABonif1 = 0 Or IVABonif1 = CByte(cadIVA) Then
                            IVABonif1 = CByte(cadIVA)
                            TotBonif1 = TotBonif1 + Round(Canti1 * Bonif1, 2)
                        ElseIf IVABonif2 = 0 Or IVABonif2 = CByte(cadIVA) Then
                            IVABonif2 = CByte(cadIVA)
                            TotBonif2 = TotBonif2 + Round((Canti1 * Bonif1), 2)
                        ElseIf IVABonif3 = 0 Or IVABonif3 = CByte(cadIVA) Then
                            IVABonif3 = CByte(cadIVA)
                            TotBonif3 = TotBonif3 + Round(Canti1 * Bonif1, 2)
                        End If
                    End If
                    
                    If Bonif2 > 0 Then
'                        SQL = "(codtipom='" & codTipoM & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(FecFactu, "F") & " and codtipoa='' and numalbar=0)"
                        numlinea = numlinea + 1
                        SQLBonif = SQLBonif & "('" & codtipom & "'," & NumFactu & "," & DBSet(FecFactu, "F") & ",'',9999999," & numlinea & "," & AlmFactu & "," & DBSet(RSaux!codarti2, "T") & ","
                        SQL = DevuelveDesdeBDNew(conAri, "sartic", "nomartic", "codartic", RSaux!codarti2, "T")
                        SQLBonif = SQLBonif & DBSet(SQL, "T") & ",''," & DBSet(Canti2, "N") & "," & DBSet(-Bonif2, "N") & ",0,0," & DBSet(0 - Round(Canti2 * Bonif2, 2), "N") & ",'M'),"
                        If IVABonif1 = 0 Or IVABonif1 = CByte(cadIVA) Then
                            IVABonif1 = CByte(cadIVA)
                            TotBonif1 = TotBonif1 + Round((Canti2 * Bonif2), 2)
                        ElseIf IVABonif2 = 0 Or IVABonif2 = CByte(cadIVA) Then
                            IVABonif2 = CByte(cadIVA)
                            TotBonif2 = TotBonif2 + Round((Canti2 * Bonif2), 2)
                        ElseIf IVABonif3 = 0 Or IVABonif3 = CByte(cadIVA) Then
                            IVABonif3 = CByte(cadIVA)
                            TotBonif3 = TotBonif3 + Round((Canti2 * Bonif2), 2)
                        End If
                    End If
'                    RSaux.MoveNext
                End If
                RSaux.Close
                Set RSaux = Nothing
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
            If SQLBonif <> "" Then 'quitamos ultima coma
                SQLBonif = Mid(SQLBonif, 1, Len(SQLBonif) - 1)
            End If
        End If
    End If
    Set vClien = Nothing
    CalcularBonificacion = True
    
    
EBonif:
    If Err.Number <> 0 Then
        CalcularBonificacion = False
    Else
        CalcularBonificacion = True
    End If
End Function



Private Function InsertarBonificaciones() As Boolean
'Si los albaranes a facturar tienen lineas de articulos con bonificacion
'añadirlos a la factura
'Insertar lineas en tabla slifac con el importe de las bonificaciones
Dim SQL As String
Dim b As Boolean

    On Error GoTo EBonifica

    InsertarBonificaciones = False
    
    If SQLBonif > "" Then
        'Insertar en scafac1 un albaran a 0, para la bonificacion
        'si no se inserto antes para los desplazamientos
        '-----------------------------------------------------------
        SQL = "(codtipom='" & codtipom & "' and numfactu=" & NumFactu & " and fecfactu=" & DBSet(FecFactu, "F") & " and codtipoa='' and numalbar=9999999)"
        SQL = "Select count(*) from scafac1 where " & SQL
        If RegistrosAListar(SQL) = 0 Then
            b = InsertarCabAlbaranesFactu("")
            If Not b Then Exit Function
        End If
               
        'Insertar lineas de Desplazamiento en tabla: slifac (lineas factura)
        '------------------------------------------------------------------
        SQL = "INSERT INTO slifac (codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,"
        SQL = SQL & "cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
        SQL = SQL & " VALUES " & SQLBonif
        conn.Execute SQL
        
    End If
    InsertarBonificaciones = True
     
EBonifica:
    If Err.Number <> 0 Then
        InsertarBonificaciones = False
    Else
        InsertarBonificaciones = True
    End If
End Function




Private Function ActualizarNumSeries(cadSQL As String) As Boolean
'Actualiza el campo "numfactu" de la tabla de num. serie: sserie
'Para ponerle la factura asociada
Dim SQL As String

    On Error GoTo EActualizar

    SQL = "UPDATE sserie SET numfactu= " & NumFactu
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaalb", "sserie")
    conn.Execute SQL

    ActualizarNumSeries = True
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarNumSeries = False
    Else
        ActualizarNumSeries = True
    End If
End Function



Private Function CalcularDatosFacturaTicket(cadSQL As String) As Boolean
Dim b As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim I As Byte
Dim ImpTotal As Currency
Dim ImpDif As Currency

    On Error GoTo ETicket
    


    
    
    '23 Nov 2010
    'Antes
    SQL = "SELECT codigiva,sum(importel) as totalimp, from sliven "
    'Ahora
    SQL = "SELECT codigiva,sum(implineareal) as totalimp from sliven "   'Ahi tenemos el importe de la linea SIN IVA
    SQL = SQL & " WHERE " & Replace(cadSQL, "scaven", "sliven")
    SQL = SQL & " group by codigiva"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not Rs.EOF And I <= 3
        SQL = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", CStr(Rs!CodigIVA), "N")
        Select Case I
            Case 1
                TipoIVA1 = Rs.Fields(0).Value
                PorceIVA1 = CCur(SQL)
                
                '23 NOv 2010      'Para la dos y al 3 tipo de iva la funcion es la misma
                'El importe viene sin IVA. Se calcula al reves
                'BaseIVA1 = Round2(RS.Fields(1).Value / (1 + (PorceIVA1 / 100)), 2)
                'ImpIVA1 = RS.Fields(1).Value - BaseIVA1
                
                BaseIVA1 = Rs.Fields(1).Value
                ImpIVA1 = Round2(((BaseIVA1 * PorceIVA1) / 100), 2)
            Case 2
                TipoIVA2 = Rs.Fields(0).Value
                PorceIVA2 = CCur(SQL)
                BaseIVA2 = Rs.Fields(1).Value
                ImpIVA2 = Round2(((BaseIVA2 * PorceIVA2) / 100), 2)



            Case 3
                TipoIVA3 = Rs.Fields(0).Value
                PorceIVA3 = CCur(SQL)
                BaseIVA3 = Rs.Fields(1).Value
                ImpIVA3 = Round2(((BaseIVA3 * PorceIVA3) / 100), 2)
    
        End Select
        Rs.MoveNext
        I = I + 1
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    
    
    'comprobar que suma bases imp + imporiva = total factura
    ImpTotal = (BaseIVA1 + ImpIVA1) + (BaseIVA2 + ImpIVA2) + (BaseIVA3 + ImpIVA3)
    BrutoFac = BaseIVA1 + BaseIVA2 + BaseIVA3
    BaseImp = BrutoFac
    
    If Not (ImpTotal = TotalFac) Then
        ImpDif = TotalFac - ImpTotal
        
        If ImpIVA1 = 0 Then
            'Sobre el DOS
            ImpIVA2 = ImpIVA2 + ImpDif
        Else
            ImpIVA1 = ImpIVA1 + ImpDif
        End If
        
        
        
'        If BaseIVA2 = 0 And BaseIVA3 = 0 Then ImpIVA1 = ImpIVA1 + ImpDif
'        ImpTotal = (BaseIVA1 + ImpIVA1) + (BaseIVA2 + ImpIVA2) + (BaseIVA3 + ImpIVA3)
    End If
    
    b = True
    
    
ETicket:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Calculando datos de ticket a factura.", Err.Description
        b = False
    End If
    CalcularDatosFacturaTicket = b
End Function


Private Function InsertarFormaPagoEnConta(nForPa As String, cadErr As String) As Boolean
Dim cadAux As String
Dim cadAux2 As String
Dim SQL As String

    On Error GoTo ErrInsForpa
    InsertarFormaPagoEnConta = False
    
    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
    'forma de pago de la factura. Sino existe insertarla
    
    'vemos si existe en la conta
    If vParamAplic.ContabilidadNueva Then
        cadAux = DevuelveDesdeBDNew(conConta, "formapago", "codforpa", "codforpa", nForPa, "N")
        If cadAux = "" Then Err.Raise 513, , "No existe forma de pago: " & nForPa   'para que no
    Else
        cadAux = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", nForPa, "N")
    End If
    
    
    'si no existe la forma de pago en conta, insertamos la de ariges
    If cadAux = "" Then
        cadAux2 = "tipforpa"
        cadAux = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", nForPa, "N", cadAux2)
        If cadAux <> "" Then
            'insertamos e sforpa de la CONTA
            SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
            SQL = SQL & " VALUES(" & nForPa & ", " & DBSet(cadAux, "T") & ", " & cadAux2 & ")"
            ConnConta.Execute SQL
            InsertarFormaPagoEnConta = True
        Else
            InsertarFormaPagoEnConta = False
        End If
    Else
        InsertarFormaPagoEnConta = True
    End If
    
    
    Exit Function
    
ErrInsForpa:
    InsertarFormaPagoEnConta = False
    cadErr = "Insertar forma de pago en Contablilidad: " & vbCrLf & Err.Description
End Function



'#### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function MarcarContabilizada(Optional cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    MarcarContabilizada = False
    
    SQL = "UPDATE scafac SET intconta=1 "
    SQL = SQL & " WHERE codtipom=" & DBSet(Me.codtipom, "T") & " AND numfactu=" & Me.NumFactu & " AND fecfactu=" & DBSet(Me.FecFactu, "F")
    conn.Execute SQL
    
    MarcarContabilizada = True
    Exit Function
    
EActualizar:
    MarcarContabilizada = False
    cadErr = Err.Description
End Function






'####  Renting
Public Function PasarRentingAFactura(TipCo As String, Operador As String, CenCoste As String, CadenaSQL As String, PeridoFacturar As Date) As Boolean
'IN -> cadSQL: cadena para seleccion de los Mantenimientos que vamos a Facturar
'Desde Mantenimientos Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String
Dim CC As String

    On Error GoTo ETraspasoFac

    PasarRentingAFactura = False
    b = False
      
    codtipom = "FAM"
    
    TipCoMan = TipCo
    OpeFactu = Operador
    MesFactu = ""
    nummante = "DBZ"  'indicara que no es un manteimiento, sino un renting
    'CEntro de coste
    manCCost = ""
    
    
    FechaFacturarRenting = PeridoFacturar  'llevara la fecha del peridodo de facturacion
    
    
    
    'Cojo el primero solo para fijar los IVAS's
    ArticFac = DevuelveDesdeBDNew(conAri, "stipco", "codartic", "codtipco", TipCoMan, "T")
    If ArticFac = "" Then
        MenError = "El tipo de contrato no tiene articulo para facturar."
        Exit Function
    End If
      
    
    If OpeFactu <> "" Then
        'Cojo el del trabajador para no hacer el select otra vez
        manCCost = "codccost"
        AlmFactu = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", OpeFactu, "N", manCCost)
    Else
        AlmFactu = "1"
    End If
    
    'CEntro de coste
    
    If vEmpresa.TieneAnalitica Then
        'Tienen analitica
        Select Case vParamAplic.ModoAnalitica
        Case 0
            'Trabajador
            'YA lo he hecho arriba
            ' manCCost = devuelve......
        
        Case 1
            'Familia
            manCCost = DevuelveDesdeBD(conAri, "codccost", "sartic,sfamia", "sartic.codfamia=sfamia.codfamia AND sartic.codartic", ArticFac, "T")
        Case Else
            'proyecto
            manCCost = CenCoste
        
        End Select
        If manCCost = "" Then
            MsgBox "Centro de coste no especificado", vbExclamation
            Exit Function
        End If
    Else
        manCCost = ""
    End If
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
    
       
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura"
    mLetraSer = "UNOSOLO"
    NumAlbTicket = CadenaSQL
    b = InsertarFactura(NumAlbTicket)

    mLetraSer = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", codtipom, "T")

    'Insertar los Vencimientos de la Factura. Tabla: svenci
    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
    If b Then
        MenError = "Error al actualizar en Tesoreria"
        b = InsertarEnTesoreria("", MenError)
    End If
    
    NumAlbTicket = ""
    mLetraSer = "" 'reutilizada
   
    'Actualizo en la tabla mantenimientos la fecha ultimio mantenimiento
    If b Then
        MenError = "Actualizar fecha ultimo renting"
        mCodAgent = DiasMes(Month(PeridoFacturar), Year(PeridoFacturar))
        ArticFac = "UPDATE sclienrenting,sclien set ultfec= '" & Format(PeridoFacturar, "yyyy-mm-") & Format(mCodAgent, "00") & "'  " & CadenaSQL
        'Y EL CLIENTE
        ArticFac = ArticFac & " AND sclienrenting.codclien = " & mCodClien
        conn.Execute ArticFac
    End If
   
ETraspasoFac:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Pasando Mantenimientos a Factura.", Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarRentingAFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarRentingAFactura = False
    End If
    Espera 0.2
End Function



'Para scobro de telefonia
Public Function TelSumlineas(Cual As Byte, Serie As String, Ano As Integer, NumFact As Long) As Currency
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    
    Select Case Cual
    Case 0
        mSQL = "select sum(Importe) from tel_lin_factura_cuotas"
    Case 1
        mSQL = "select sum(Importe) from tel_lin_factura_consumos"
    Case 2
        mSQL = "select sum(Importe) from tel_lin_factura_descuentos"
    Case Else
        mSQL = "select sum(Importe) from tel_lin_factura_especial"
    End Select
    
    mSQL = mSQL & " where serie = " & DBSet(Serie, "T")
    mSQL = mSQL & " and ano = " & Ano
    mSQL = mSQL & " and numfact = " & NumFact
    Set mrs = New ADODB.Recordset
    mrs.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not mrs.EOF Then
        If IsNull(mrs.Fields(0)) Then
            TelSumlineas = 0
        Else
            TelSumlineas = mrs.Fields(0)
        End If
    End If
    mrs.Close
    Set mrs = Nothing
End Function

Public Sub PonerValorNuevasNormasBancarias(Valor As Boolean)
    NuevasNormasBancarias = Valor
End Sub


Private Sub Class_Initialize()
    NuevasNormasBancarias = False
End Sub
