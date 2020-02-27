VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPARA1
Private mvarTipoDtos As Byte      'Tipo de Descuentos: 0= Aditivo 1=Sobre Resto
Private mvarOrdenDtos As Byte   'Orden de Dtos para obtener los dtos a aplicar al precio segun la tabla: sdtofm
                                'Si F=0004 y M=0004 y existe dtos para (F,M)=(0000,0004) y para (F,M)=(0004,0000) siendo 0000=todas las marcas/todas las familias
                                'si ordendto=0(Fam/Marca) aplicaremos el dto (F,M)=(0004,0000)
                                'si ordendto=1 (Marca/Fam) aplicaremos el dto (F,M)=(0000,0004)
                                
Private mvarCodTarif As Integer   'Cod. Tarifa de PVP
Private mvarCajaComp As Boolean      'Cajas completas precios (S/N)
Private mvarPreuKmCli As Single 'Precio Km Desplazamiento Clientes
Private mvarPreuKmTec As Single 'Precio Km Desplazamiento Técnicos
Private mvarnomGerente As String  'Nombre del Director Gerente
Private mvarnomAdmin As String  'Nombre del Responsable Administración
Private mvarHayMante As Boolean   'Si hay Mantenimiento (S/N)
Private mvarHayRepar As Boolean   'Si hay Reparaciones (S/N)
Private mvarHayFrecu As Boolean   'Si hay Frecuencias (S/N)
Private mvarHayServicio As Boolean   'Si hay Servicios (S/N)

'Nuevo Julio 2010
Private mvarHayDeparNuevo As Byte    'Si hay 0.- Direccion  / 1.- Departamento /  2.- Obra  (SAIL pej.)

Private mvarCtrStock As Boolean   'Si hay Control de Stock (S/N)
Private mvarInventar As Boolean  'Si se realiza Inventario ordenando lista dos por Proveedor o no
Private mvarHaynserie As Boolean 'Si hay Nº de Serie en Compras (S/N)
Private mvarDiasSiMan As Integer 'Nº de dias si hay mantenimiento
Private mvarDiasNoMan As Integer 'Nº de dias si NO hay mantenimiento

'Datos Facturacion
Private mvarCodArticDes As String 'Articulo para facturar desplazamientos
Private mvarDiaPago1 As Byte 'Dia de pago 1 para compras
Private mvarDiaPago2 As Byte 'Dia de pago 2 para compras
Private mvarDiaPago3 As Byte 'Dia de pago 3 para compras
Private mvarMesnogir As Byte 'Mes a no girar para compras
Private mvarCodForPa As String 'forma de pago para cheque regalo

'Datos de Internet
Private mvarDireMail As String 'Direccion email
Private mvarSMTPhost As String 'Host SMTP para envio de email
Private mvarSMTPuser As String 'Usuario SMTP para envio de email
Private mvarSMTPpass As String 'Password SMTP para envio de email
Private mvarWebSoporte As String 'Web de soporte
Private mvarMailSoporte As String 'Mail de soporte
Private mvarWebVersion As String 'Version de la web

'Datos de la Contabilidad
Private mvarSerConta As String 'Servidor BD de Contabilidad
Private mvarUsuConta As String 'Usuario BD de Contabilidad
Private mvarPasConta As String 'Password BD de Contabilidad
Private mvarNumConta As Byte 'Numero BD de Contabilidad


'Valores por defecto.   Toca meter un variable pq en el matenimiento (en el form) no deja modificar si no es con el objeto
'-----------------
'
'      Clientes
Private mvarDefActiv As Integer
Private mvarDefEnvio As Integer
Private mvarDefZona As Integer
Private mvarDefRuta As Integer
Private mvarDefSitu As Integer
Private mvarDefTarifa As Integer
Private mvarDefAgente As Integer

'Facturacion recargas telefonia
Private mvarCodarticTfnia As String



    'aVISOS  al entrar
    Private mvaravipedcli  As Integer
    Private mvaravipedpro As Integer
    Private mvaravialbcli As Integer
    Private mvaravialbpro As Integer
    Private mvaravimanteni As Integer
    Private mvaravirepara As Integer
    Private mvaraviavisos As Integer
    Private mvarTieneAvisos As Boolean 'De este no hay valor en BD. Simplememnte si alguno de ellos tiene valor, el tendra valor


'Si LLEVA REcargo de equivalencia
Private mvarRecargoEquivalencia As Boolean
Private mvarTipoIVAre1 As Integer
Private mvarTipoIVAre2 As Integer
Private mvarTipoIVAre3 As Integer
Private mvarIVA1 As Integer
Private mvarIVA2 As Integer
Private mvarIVA3 As Integer


'Retencion y REA
Private mvarIVAREA As Integer
Private mvarCtaReten As String
Private mvarPorReten As Currency

Private mvarIVA_Exento As Integer
Private mvarIVA_Intracomunitario As Integer

'Tickets TPV agrupados a contabilidad
Private mvarContabilizarTicketAgrupados As Boolean

'Albaranes y facturas en B
Private mvarContabilidadB As Byte

'Cueenta para la aportacion al terminal en RADIVAL
Private mvarctaaportacion As String

'Produccion
Private mvarProduccion As Boolean

'Descriptores
Private mvarDescriptores As Boolean

'Que llevara al campo observaciones de la factura cli/pro
Private mvarObsFactura As Byte

'Reciclado / PUNTO VERDE
Private mvarArtReciclado As String


'FONTENAS
Private mvarArtPortes As String
Private mvarImporteMinimo As Currency
Private mvarAbonoKilos As Currency


'Produccion
Private mvarComponentePorcentaje As Boolean   'Si el componente entra como porcentaje o como cantidad


' ---- [14/09/2009] (LAURA)
Private mvarDtoxCanti As Boolean 'Lleva dtos por cantidad para facturacion: Pedidos, Ofertas, Albaranes
Private mvarCreaTarifArt2 As Byte 'crea articulo en todas las tarifas
'                               David cambia a byte
'                               0.- NO crea
'                               1.- Solo en tarifa ppal
'                               2.- En todas tarifas



' ---- [19/10/2009] [LAURA]: añadir campo modo analitica
Private mvarModoAnalit As Byte 'modo analitica: 0=trabajador, 1=Familia, 2=Proyecto

Private mvarEnvioDesdeOutlook As Boolean  'Si envia por outlook



'Exe nuestro k llama a el Outlook
Private mvarExeEnvioMail As String

'Dias de garantia despues de efectuar un reparacion
Private mvarDiasGarantia As Integer


'De momento solo tiene el DE leer
'Si lleva el trozo del CRM
Private mvarTieneCRM As Boolean




'Para el cambio del IVA
'Guardaremos los valores anteriores
Private mvarOLDIVAre1 As Integer
Private mvarOLDIVAre2 As Integer
'Private mvarOLDIVAre3 As Integer
Private mvarOLDIVA1 As Integer
Private mvarOLDIVA2 As Integer
'Private mvarOLDIVA3 As Integer

Private mvarNumCopiasFacturacion As Integer


'4 Octubre 2010
Private mvarArticuloRecargoFinanciero As String


'Julio2010
'SAIL.  La tbala de albaranes, pedidos, oferas es la misma, pero en funcion
'       de este parametro abrirar un frm u otro.
'       asi unos campos estarn ocultos para uo y no para otros
Private mvarTipoFormularioClientes As Byte
        '0  .- El estandar, el que habia hasta ahora
        '1  .- frm's tipo SAIL
        

'[13/10/2010] [DAVID]: Portes                     'Mirar LOG impresion albaranes
Private mvarTipoPortes As Integer
        'Si lleva articulo de portes, la apliacion de los portes puede ser como fontenas(tipo 0) que es
        'por albaran factura, y tipo 1 que es como herbelca
        '-1:  NO LLEVA PORTES
        '0- Por albaran/Factura    Para cada albaran , pedido o factura tendremos un tipo porte
        '1- Por factura "HERBELCA
        
'[13/10/2010] [DAVID]: Direcciones de envio, ademas de dpto/direccion. No entra en la facturacion
Private mvarDireccionesEnvio As Boolean


'[08/11/2010] [DAVID]: Fras de mostrador generaran facturas mostrador
Private mvarFrasMostradorSerieDistinta As Boolean


'[26/11/2010] [DAVID]: Al crear el albarn marcarlo con el check de ParaFacturar
Private mvarMarcarAlbaranFacturar As Boolean


' `OperAseguradas`    `situbloq`  'Situacion boque oper aseguradas' after `OperAseguradas`;
'[17/12/2010] [DAVID]
Private mvarOperacionesAseguradas As Boolean
Private mvarSituacionBloqueoOpAseg As Integer


' Rotacion.  Valores para el listado de rotacion
'[22/12/2010] [DAVID]
Private mvarRot_ConsumMes1 As Currency
Private mvarRot_ConsumMes2 As Currency
Private mvarRot_ConsumMesMax As Currency
Private mvarRot_ConsumMesMin As Currency

'LOG de cambio de precios o descuentos
'[22/12/2010] [DAVID]
Private mvarLogCambioPrecDto As Boolean


Private mvarAlmacenB As Integer 'Nos dira cual es el almacen de B
                                'No esta en el form. Solo por YOG.   Ademas me sirve para identificar a la empresa herbelca""


'[15/02/2011] [DAVID]
' Cuando una fra proveedor sea contabilizada se crearaan dos fras mas, una en proveedor y otra en cliente
Private mvarIvaIntracomAdicional As Integer   '
Private mvarCtaContabIntracom   As String   'cuenta a la que iran las dos fras intracom adicionales
Private mvarIntracomAdicionalContab As Boolean 'si pasan para que no sean contabilizadas(numasien=0)



'[15/02/2011] [DAVID]
Private mvarPathFacturaE As String

'[03/05/2011]  David
'           Numero de copias. Sera un varchar numerico que tendra
'           por posiciones
Private mvarNumeroCopias As String
' a partir de el string, CON GET, pondremos las variables.
' Es decir habran n GETs que devolveran un integer en funcion


'[19/05/2011] [DAVID]
Private mvarPrecioMinimo As Boolean    'En la rutina de precios cogera siempre el MINIMO (herbelca solo)
'   -------> El siguiente parametro grabara si el precio se ha modificado A LA BAJA.
'               es para herbelca solamente, con lo cual de momento, no creo valor en BD, slo mvarPreciominimo es HERBELCA
Private mvarGrabaModificarPrecioAlaBaja As Boolean

'[08/08/2011] [DAVID]
Private mvarArtSeparador As String  'Para las ofertas crea el articulo separador

'[16/12/2011] [DAVID]
Private mvarAriagro As String  'Si esta vinculado con Ariagro(ADV y/o Tikets campos)
Private TieneADV As Byte  '0- No establecida  1.- NO tiene ADV   2- Si



'[05/02/2012] [DAVID]
Private mvarRenting As Boolean

'[27/02/2012] [DAVID]
Private mvarGestionFlotas As Boolean


'ALMAGRUPO.  En teoria solo Herbelca
'Parametro SIN parametro
'Si existe una tabla y esta tiene valor es que COMUNICA
Private mvarComunicaAlmagrupo As Boolean


'[09/08/2012]
'Porcentaje de dias a partir del cual un talon / pagare se considerara devuelto
Private mvarPorcenPagoAgentTalPag As Currency

Private mvarFechaCambioIva As Date
Private mvarClientesPotenciales As Boolean  '17/oct/2012


'17/12/2012
Private mvarSituaEnCodigoArticulo As Boolean  'Este NO esta en BD. Cuando entre en el programa se establecera leyendo Dir(..\SituCod.dat


'08/01/2013
Private mvarTieneTelefonia2 As Byte '0:NO   1:Alzira   2:Catadau(coarval)
Private mvarArtiTelefonia As String 'Solo tiene LET, para guardar. No lo guardo en memoria

'05/03/2013
Private mvarRecMercanciaSoloPpal As Boolean  'Recepcion de mercancia solo busque en los pedidos del almacen ppal. En herbelca pondres

'09/05/2013
Private mvarArtTfniaIvaExento As String  'articulo telefonia IVA exento


'2014/01/17
'  0. Resto     1.- WHOSE
Private mvarQueEmpresaEs  As Byte



'30/08/2013
'cuando cambia NIF cliente tb se va a cambiarlo a la contbildad asociada
'Pero puede ser que tenga arigasol-aragros bodega......
'Con este parametro le diremos si tiene que ir a comparar con las otras BD
Private mvarComprobarBancoRestoAplicaciones As Boolean


'03/Sept/2013
'------------
Private mvarIncializarStockEnInventario As Boolean





'20/02/2014
'--------------
'No lleva campo en BD, hasta que no decida si creamos un objeto...
'Con lo cual, si es WHOSE, nada mas entrar, cargamos este valor que es MUY utilizado
Private mvarPathDocsWHOSE As String





'LLeva campo en la BD pero no esta el formulario
Private mvarAguasPotables As Boolean





'NO lleva Campo en BD. Cuando lea la empresa, que vea cuantos bancos tiene.
'Si tiene solo uno, ese sera el banco por defecto, sino=0
Private mvarCodigoUnicoBancoPropio_ As Integer


'[22/04/2014] [DAVID]
Private mvarActualizaPrecioEspecial As Boolean    'Cuando actualiza el precio de un articulo, si actualiza tb el precio especial

'[14/05/2014] [DAVID]
Private mvarEntradaRapidaFacturasMostrador As Boolean    'Lo pide NAVARRES. Cuando esta introduciendo las facturas de mostrador menos clicks




'[27/05/2014]
Private mvarNumeroInstalacion As Byte
    ' 0 Sin especificar
    ' 1.- ALZIRA
    ' 2.-  Herbelca (en muchos sitios pone almacenb>90
    ' 3.- Teinsa   buscar frecuencias
    ' 4.- EULER
    ' 5.- Fontenas
    ' 6.- Fenollar
'[14/10/2014]
Private mvarManipuladorFitosanitarios As Boolean

'Abril 2015
'Es oculto
Private mvarInvSujetoPasivo As Boolean



'Sept 2015
Private mvarHuertos As Boolean  'NO lleva en BD spara1. Si existe la tabla sclienhuertos pondremos a true


'Nov 2015
Private mvarRecalculoMargen As Boolean


'Marzo 2016. Nueva contabilidad ARIADNA. LA BD cambia
Private mvarContabilidadNueva As Boolean


'OCt2016
Private mvarLotesGeneralitat As Boolean  'NO lleva en BD spara1. Si tiene fitos y existe la tabla slotesgeneralitat pondremos a true


'MAY 2017
Private mvarInventarioCodigoArticulo As Boolean



'JUNIO217
Private mvarSIITiene As Boolean   'No tiene en spara1. Lo leera (cuando hay abierto ConnConta) de contabilidad asociada
Private mvarSIIDias As Integer
Private mvarSIIFechaIni As Date


'Agosto.
Private mvarPtosAsignar As Currency  'Si lleva >0 tendra sistema de puntos
Private mvarPtosImporteCalculo As Currency
Private mvarPtosEquivalencia As Currency    'ptos a euro
Private mvarPtosArticuloCanje As String     'Aticulo de canje
Private mvarPtosFechaIncio As String     'Aticulo de canje


'SEPTIE17
Private mvarPrecioHoraCosteEUL As Currency

'FEB 2018
' Si tiene telefonia, ver si tiene activado el modulo de venta a plazos de telefonos
' if mvarTieneTelefonia2>0 irá a la spara2 a leer el valor
Private mvarTelefoniaVtaPlazos As Boolean



'MAYO 2018
'Reimpresion facturas firmadas socio,
Private mvarPathFirmasFacturas As String
Private mvarPathFirmasAlbaran As String




'NOVIEMBRE 2018
Private mvarCartaPortes As Boolean

Private mvarHaciendoFrmulariosGrandes As Boolean



'ENE 2019
Private mvarImpresoraFenollarA As String
Private mvarImpresoraFenollarB As String


Private mvarMaximoImporteEfectivo As String

Private mvarPuedeModificarAriconta As Boolean
Private mvarTieneComponentes_y_Produccion As Boolean
Private mvarImportacionesCoarval As Boolean



'Nov2019.   Gasolinera
Private mvarLlevaGasolinera As Boolean



'hace como la conta
' rellena a ceros el codartic
' De momento TAXCO, con lo cual NO hay parametros
Private mvarDigitosCodartic As Byte



Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes







'------------------------------------------------
'Propiedades del modulo CParamAplic
'------------------------------------------------

Public Property Let TipoDtos(ByVal vData As Byte)
     mvarTipoDtos = vData
End Property

Public Property Get TipoDtos() As Byte
     TipoDtos = mvarTipoDtos
End Property


Public Property Let OrdenDtos(ByVal vData As Byte)
     mvarOrdenDtos = vData
End Property

Public Property Get OrdenDtos() As Byte
     OrdenDtos = mvarOrdenDtos
End Property

Public Property Let CodTarifa(ByVal vData As Integer)
     mvarCodTarif = vData
End Property

Public Property Get CodTarifa() As Integer
     CodTarifa = mvarCodTarif
End Property


Public Property Let CajasCompletas(ByVal vData As Boolean)
     mvarCajaComp = vData
End Property

Public Property Get CajasCompletas() As Boolean
     CajasCompletas = mvarCajaComp
End Property


Public Property Let PrecioKmClientes(ByVal vData As Single)
     mvarPreuKmCli = vData
End Property

Public Property Get PrecioKmClientes() As Single
     PrecioKmClientes = mvarPreuKmCli
End Property


Public Property Let PrecioKmTecnicos(ByVal vData As Single)
     mvarPreuKmTec = vData
End Property

Public Property Get PrecioKmTecnicos() As Single
     PrecioKmClientes = mvarPreuKmTec
End Property


Public Property Let NomGerente(ByVal vData As String)
     mvarnomGerente = vData
End Property


Public Property Get NomGerente() As String
     NomGerente = mvarnomGerente
End Property


Public Property Let NomRespAdmin(ByVal vData As String)
     mvarnomAdmin = vData
End Property

Public Property Get NomRespAdmin() As String
     NomRespAdmin = mvarnomAdmin
End Property


Public Property Let Mantenimientos(ByVal vData As Boolean)
     mvarHayMante = vData
End Property

Public Property Get Mantenimientos() As Boolean
     Mantenimientos = mvarHayMante
End Property


Public Property Let Reparaciones(ByVal vData As Boolean)
     mvarHayRepar = vData
End Property

Public Property Get Reparaciones() As Boolean
     Reparaciones = mvarHayRepar
End Property

Public Property Let Frecuencias(ByVal vData As Boolean)
'    If vData Then
'        MsgBox "Lo sentimos todavía no puede activar las frecuencias. En próximas versiones de la aplicación será posible", vbExclamation
'        vData = False
'    End If
     mvarHayFrecu = vData
End Property

Public Property Get Servicios() As Boolean
     Servicios = mvarHayServicio
End Property

Public Property Let Servicios(ByVal vData As Boolean)
    Dim Mens As String
    If vData Then
        If MovimientosParaServicios() Then
            vData = True
        Else
            Mens = "Para poder activar los servicios es necesario que los movimientos ALS y FAS estén dados de alta." & vbCrLf & _
                    "Délos de alta ahora y luego vuelva a intentar activar los servicios. Gracias."
            MsgBox Mens, vbExclamation
            vData = False
        End If
    End If
     mvarHayServicio = vData
End Property

Public Property Get Frecuencias() As Boolean
     Frecuencias = mvarHayFrecu
End Property


'JULIO 2010
'Si hay Departamento / Direccion  / Obra  (SAIL pej.)
'    0.- Direccion  / 1.- Departamento /  2.- Obra  (SAIL pej.)
Public Property Let HayDeparNuevo(ByVal vData As Byte)
     mvarHayDeparNuevo = vData
End Property
Public Property Get HayDeparNuevo() As Byte
     HayDeparNuevo = mvarHayDeparNuevo
End Property


Public Property Let ControlStock(ByVal vData As Boolean)
     mvarCtrStock = vData
End Property

Public Property Get ControlStock() As Boolean
     ControlStock = mvarCtrStock
End Property

Public Property Let InventarioxProv(ByVal vData As Boolean)
     mvarInventar = vData
End Property

Public Property Get InventarioxProv() As Boolean
     InventarioxProv = mvarInventar
End Property

Public Property Let NumSeries(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarHaynserie = vData
End Property

Public Property Get NumSeries() As Boolean
     NumSeries = mvarHaynserie
End Property


Public Property Let DiasSiMante(ByVal vData As Integer)
     mvarDiasSiMan = vData
End Property

Public Property Get DiasSiMante() As Integer
     DiasSiMante = mvarDiasSiMan
End Property


Public Property Let DiasNoMante(ByVal vData As Integer)
     mvarDiasNoMan = vData
End Property

Public Property Get DiasNoMante() As Integer
     DiasNoMante = mvarDiasNoMan
End Property


Public Property Let ArticDesplaz(ByVal vData As String)
     mvarCodArticDes = vData
End Property

Public Property Get ArticDesplaz() As String
     ArticDesplaz = mvarCodArticDes
End Property


Public Property Let DiaPago1(ByVal vData As Byte)
     mvarDiaPago1 = vData
End Property

Public Property Get DiaPago1() As Byte
     DiaPago1 = mvarDiaPago1
End Property

Public Property Let DiaPago2(ByVal vData As Byte)
     mvarDiaPago2 = vData
End Property

Public Property Get DiaPago2() As Byte
     DiaPago2 = mvarDiaPago2
End Property

Public Property Let DiaPago3(ByVal vData As Byte)
     mvarDiaPago3 = vData
End Property

Public Property Get DiaPago3() As Byte
     DiaPago3 = mvarDiaPago3
End Property


Public Property Let MesNoGirar(ByVal vData As Byte)
     mvarMesnogir = vData
End Property

Public Property Get MesNoGirar() As Byte
     MesNoGirar = mvarMesnogir
End Property

'forma de pago de los cheques regalo
Public Property Let ForPagoChequeRegalo(ByVal vData As String)
     mvarCodForPa = vData
End Property

Public Property Get ForPagoChequeRegalo() As String
     ForPagoChequeRegalo = mvarCodForPa
End Property




Public Property Let DireMail(ByVal vData As String)
     mvarDireMail = vData
End Property

Public Property Get DireMail() As String
     DireMail = mvarDireMail
End Property

Public Property Let SMTPhost(ByVal vData As String)
     mvarSMTPhost = vData
End Property

Public Property Get SMTPhost() As String
     SMTPhost = mvarSMTPhost
End Property

Public Property Let SMTPuser(ByVal vData As String)
     mvarSMTPuser = vData
End Property

Public Property Get SMTPuser() As String
     SMTPuser = mvarSMTPuser
End Property

Public Property Let SMTPpass(ByVal vData As String)
     mvarSMTPpass = vData
End Property

Public Property Get SMTPpass() As String
     SMTPpass = mvarSMTPpass
End Property

Public Property Let WebSoporte(ByVal vData As String)
     mvarWebSoporte = vData
End Property

Public Property Get WebSoporte() As String
     WebSoporte = mvarWebSoporte
End Property

Public Property Let MailSoporte(ByVal vData As String)
     mvarMailSoporte = vData
End Property

Public Property Get MailSoporte() As String
     MailSoporte = mvarMailSoporte
End Property

Public Property Let WebVersion(ByVal vData As String)
     mvarWebVersion = vData
End Property

Public Property Get WebVersion() As String
     WebVersion = mvarWebVersion
End Property



'servidor conta
Public Property Let ServidorConta(ByVal vData As String)
     mvarSerConta = vData
End Property

Public Property Get ServidorConta() As String
     ServidorConta = mvarSerConta
End Property



Public Property Let UsuarioConta(ByVal vData As String)
     mvarUsuConta = vData
End Property

Public Property Get UsuarioConta() As String
     UsuarioConta = mvarUsuConta
End Property


Public Property Let PasswordConta(ByVal vData As String)
     mvarPasConta = vData
End Property

Public Property Get PasswordConta() As String
     PasswordConta = mvarPasConta
End Property


Public Property Let NumeroConta(ByVal vData As Byte)
     mvarNumConta = vData
End Property

Public Property Get NumeroConta() As Byte
     NumeroConta = mvarNumConta
End Property




Public Property Let PorDefecto_Activ(ByVal vData As Integer)
     mvarDefActiv = vData
End Property

Public Property Get PorDefecto_Activ() As Integer
     PorDefecto_Activ = mvarDefActiv
End Property
Public Property Let PorDefecto_Envio(ByVal vData As Integer)
     mvarDefEnvio = vData
End Property

Public Property Get PorDefecto_Envio() As Integer
     PorDefecto_Envio = mvarDefEnvio
End Property
Public Property Let PorDefecto_Zona(ByVal vData As Integer)
     mvarDefZona = vData
End Property

Public Property Get PorDefecto_Zona() As Integer
     PorDefecto_Zona = mvarDefZona
End Property
Public Property Let PorDefecto_Ruta(ByVal vData As Integer)
     mvarDefRuta = vData
End Property

Public Property Get PorDefecto_Ruta() As Integer
     PorDefecto_Ruta = mvarDefRuta
End Property
Public Property Let PorDefecto_Situ(ByVal vData As Integer)
     mvarDefSitu = vData
End Property

Public Property Get PorDefecto_Situ() As Integer
     PorDefecto_Situ = mvarDefSitu
End Property

Public Property Let PorDefecto_Tarifa(ByVal vData As Integer)
     mvarDefTarifa = vData
End Property

Public Property Get PorDefecto_Tarifa() As Integer
     PorDefecto_Tarifa = mvarDefTarifa
End Property
Public Property Let PorDefecto_Agente(ByVal vData As Integer)
     mvarDefAgente = vData
End Property

Public Property Get PorDefecto_Agente() As Integer
     PorDefecto_Agente = mvarDefAgente
End Property


'0:NO   1:Alzira   2:Catadau(coarval)    3 ANna  / bolbaite / Navarres
Public Property Get TieneTelefonia2() As Byte
     TieneTelefonia2 = mvarTieneTelefonia2
End Property


Public Property Let avipedcli(ByVal vData As Integer)
     mvaravipedcli = vData
End Property

Public Property Get avipedcli() As Integer
     avipedcli = mvaravipedcli
End Property
    
Public Property Let avipedpro(ByVal vData As Integer)
     mvaravipedpro = vData
End Property

Public Property Get avipedpro() As Integer
     avipedpro = mvaravipedpro
End Property

Public Property Let avialbcli(ByVal vData As Integer)
     mvaravialbcli = vData
End Property

Public Property Get avialbcli() As Integer
     avialbcli = mvaravialbcli
End Property
Public Property Let avialbpro(ByVal vData As Integer)
     mvaravialbpro = vData
End Property

Public Property Get avialbpro() As Integer
     avialbpro = mvaravialbpro
End Property
    

Public Property Let avimanteni(ByVal vData As Integer)
     mvaravimanteni = vData
End Property

Public Property Get avimanteni() As Integer
     avimanteni = mvaravimanteni
End Property
    
Public Property Let avirepara(ByVal vData As Integer)
     mvaravirepara = vData
End Property

Public Property Get avirepara() As Integer
     avirepara = mvaravirepara
End Property

Public Property Let aviavisos(ByVal vData As Integer)
     mvaraviavisos = vData
End Property

Public Property Get aviavisos() As Integer
     aviavisos = mvaraviavisos
End Property

'Solo tiene GET
Public Property Get TieneAvisos() As Integer
     TieneAvisos = mvarTieneAvisos
End Property



'---------------------------------------
'RECARGO EQUIVALENCIA
Public Property Get TipoIVAre1() As Integer
     TipoIVAre1 = mvarTipoIVAre1
End Property

Public Property Let TipoIVAre1(ByVal vData As Integer)
     mvarTipoIVAre1 = vData
End Property
Public Property Get TipoIVAre2() As Integer
     TipoIVAre2 = mvarTipoIVAre2
End Property

Public Property Let TipoIVAre2(ByVal vData As Integer)
     mvarTipoIVAre2 = vData
End Property

Public Property Get TipoIVAre3() As Integer
     TipoIVAre3 = mvarTipoIVAre3
End Property

Public Property Let TipoIVAre3(ByVal vData As Integer)
     mvarTipoIVAre3 = vData
End Property

Public Property Get TipoIVA1() As Integer
     TipoIVA1 = mvarIVA1
End Property

Public Property Let TipoIVA1(ByVal vData As Integer)
     mvarIVA1 = vData
End Property
Public Property Get TipoIVA2() As Integer
     TipoIVA2 = mvarIVA2
End Property

Public Property Let TipoIVA2(ByVal vData As Integer)
     mvarIVA2 = vData
End Property

Public Property Get TipoIVA3() As Integer
     TipoIVA3 = mvarIVA3
End Property

Public Property Let TipoIVA3(ByVal vData As Integer)
     mvarIVA3 = vData
End Property

Public Property Get RecargoEquivalencia() As Integer
     RecargoEquivalencia = mvarRecargoEquivalencia
End Property





'CAMBIO IVA 2010
'-----------------------------------------------------
Public Property Get OLDIVAre1() As Integer
     mvarOLDIVAre1 = mvarOLDIVAre1
End Property

Public Property Let OLDIVAre1(ByVal vData As Integer)
     mvarOLDIVAre1 = vData
End Property
Public Property Get OLDIVAre2() As Integer
     OLDIVAre2 = mvarOLDIVAre2
End Property

Public Property Let OLDIVAre2(ByVal vData As Integer)
     mvarOLDIVAre2 = vData
End Property

'Public Property Get OLDIVAre3() As Integer
'     OLDIVAre3 = mvarOLDIVAre3
'End Property
'
'Public Property Let OLDIVAre3(ByVal vData As Integer)
'     mvarOLDIVAre3 = vData
'End Property

Public Property Get OLDIVA1() As Integer
     OLDIVA1 = mvarOLDIVA1
End Property

Public Property Let OLDIVA1(ByVal vData As Integer)
     mvarOLDIVA1 = vData
End Property

Public Property Get OLDIVA2() As Integer
     OLDIVA2 = mvarOLDIVA2
End Property

Public Property Let OLDIVA2(ByVal vData As Integer)
     mvarOLDIVA2 = vData
End Property

'Public Property Get OLDIVA3() As Integer
'     OLDIVA3 = mvarOLDIVA3
'End Property
'
'Public Property Let OLDIVA3(ByVal vData As Integer)
'     mvarOLDIVA3 = vData
'End Property


'RETENCION Y REGIMEN ESPECIAL AGRARIO
'---------------------------------------
Public Property Get IVA_REA() As Integer
     IVA_REA = mvarIVAREA
End Property

Public Property Let IVA_REA(ByVal vData As Integer)
     mvarIVAREA = vData
End Property


Public Property Get CtaReten() As String
     CtaReten = mvarCtaReten
End Property

Public Property Let CtaReten(ByVal vData As String)
     mvarCtaReten = vData
End Property

Public Property Get PorReten() As Currency
     PorReten = mvarPorReten
End Property

Public Property Let PorReten(ByVal vData As Currency)
     mvarPorReten = vData
End Property

'IVA exento
Public Property Get IVA_Exento2() As Integer
     IVA_Exento2 = mvarIVA_Exento
End Property

Public Property Let IVA_Exento2(ByVal vData As Integer)
     mvarIVA_Exento = vData
End Property



'IVA intracomunitario
Public Property Get IVA_Intracomunitario() As Integer
     IVA_Intracomunitario = mvarIVA_Intracomunitario
End Property

Public Property Let IVA_Intracomunitario(ByVal vData As Integer)
     mvarIVA_Intracomunitario = vData
End Property



'Contabilizacioon del Ticket del TPV a grupado
Public Property Get ContabilizarTicketAgrupados() As Boolean
     ContabilizarTicketAgrupados = mvarContabilizarTicketAgrupados
End Property

Public Property Let ContabilizarTicketAgrupados(ByVal vData As Boolean)
     mvarContabilizarTicketAgrupados = vData
End Property


'Numero contabilidad para las albaranes /facturas en B
Public Property Get ContabilidadB() As Byte
     ContabilidadB = mvarContabilidadB
End Property

Public Property Let ContabilidadB(ByVal vData As Byte)
     mvarContabilidadB = vData
End Property


'APortacion al terminal. Ej. Facturas de mostrador en RADIVAL
Public Property Get ctaAportacion() As String
     ctaAportacion = mvarctaaportacion
End Property

Public Property Let ctaAportacion(ByVal vData As String)
     mvarctaaportacion = vData
End Property


'Tiene produccion
Public Property Get Produccion() As Boolean
     Produccion = mvarProduccion
End Property

Public Property Let Produccion(ByVal vData As Boolean)
     mvarProduccion = vData
End Property

'Tiene Descriptores
Public Property Get Descriptores() As Boolean
     Descriptores = mvarDescriptores
End Property

Public Property Let Descriptores(ByVal vData As Boolean)
     mvarDescriptores = vData
End Property


'Campo observaciones en factura
Public Property Get ObsFactura() As Byte
     ObsFactura = mvarObsFactura
End Property

Public Property Let ObsFactura(ByVal vData As Byte)
     mvarObsFactura = vData
End Property


'Articulo de Recilado o punto Verde
Public Property Get ArtReciclado() As String
     ArtReciclado = mvarArtReciclado
End Property

Public Property Let ArtReciclado(ByVal vData As String)
     mvarArtReciclado = vData
End Property


'Octubre 2010
'Para saber si aplica los portes ver tipoportes
Public Property Get ArtPortesN() As String
     ArtPortesN = mvarArtPortes
End Property

Public Property Let ArtPortesN(ByVal vData As String)
     mvarArtPortes = vData
End Property

Public Property Get ImporteMinimo() As Currency
     ImporteMinimo = mvarImporteMinimo
End Property

Public Property Let ImporteMinimo(ByVal vData As Currency)
     mvarImporteMinimo = vData
End Property
Public Property Get AbonoKilos() As Currency
     AbonoKilos = mvarAbonoKilos
End Property

Public Property Let AbonoKilos(ByVal vData As Currency)
     mvarAbonoKilos = vData
End Property


Public Property Get ComponentePorcentaje() As Boolean
     ComponentePorcentaje = mvarComponentePorcentaje
End Property

Public Property Let ComponentePorcentaje(ByVal vData As Boolean)
     mvarComponentePorcentaje = vData
End Property


' ---- [14/09/2009] (LAURA)
Public Property Get DtoxCantidad() As Boolean
     DtoxCantidad = mvarDtoxCanti
End Property

Public Property Let DtoxCantidad(ByVal vData As Boolean)
     mvarDtoxCanti = vData
End Property

' ---- [14/09/2009] (LAURA)
'David cambia ENERO 2010
'                               David cambia a byte
'                               0.- NO crea
'                               1.- Solo en tarifa ppal
'                               2.- En todas tarifas

Public Property Get CreaTarifasArticulo() As Byte
     CreaTarifasArticulo = mvarCreaTarifArt2
End Property

Public Property Let CreaTarifasArticulo(ByVal vData As Byte)
     mvarCreaTarifArt2 = vData
End Property


' ---- [19/10/2009] [LAURA]: añadir campo modo analitica  0=trabajador, 1=Familia, 2=Proyecto
Public Property Get ModoAnalitica() As Byte
     ModoAnalitica = mvarModoAnalit
End Property

Public Property Let ModoAnalitica(ByVal vData As Byte)
     mvarModoAnalit = vData
End Property



' ---- [31/12/2009] (DAVID)
Public Property Get EnvioDesdeOutlook() As Boolean
     EnvioDesdeOutlook = mvarEnvioDesdeOutlook
End Property

Public Property Let EnvioDesdeOutlook(ByVal vData As Boolean)
     mvarEnvioDesdeOutlook = vData
End Property



'Febrero 2010
Public Property Get ExeEnvioMail() As String
     ExeEnvioMail = mvarExeEnvioMail
End Property

Public Property Let ExeEnvioMail(ByVal vData As String)
     mvarExeEnvioMail = vData
End Property

Public Property Get DiasGarantia() As Integer
     DiasGarantia = mvarDiasGarantia
End Property

Public Property Let DiasGarantia(ByVal vData As Integer)
     mvarDiasGarantia = vData
End Property




' ---- [13/04/2010] (DAVID) De momento NO tiene LET
Public Property Get TieneCRM() As Boolean
     TieneCRM = mvarTieneCRM
End Property

'Public Property Let TieneCRM(ByVal vData As Boolean)
'     mvarTieneCRM = vData
'End Property


Public Property Get NumCopiasFacturacion() As Integer
     NumCopiasFacturacion = mvarNumCopiasFacturacion
End Property

Public Property Let NumCopiasFacturacion(ByVal vData As Integer)
     mvarNumCopiasFacturacion = vData
End Property

' ---- [02/07/2010] (DAVID) De momento NO tiene LET
        '0  .- El estandar, el que habia hasta ahora
        '1  .- frm's tipo SAIL
Public Property Get TipoFormularioClientes() As Byte
    TipoFormularioClientes = mvarTipoFormularioClientes
End Property
'Public Property Let TipoFormularioClientes(ByVal vData As byte)
'     mvarTipoFormularioClientes = vData
'End Property


' ---- [04/10/2010] (DAVID)
'Articulo de Recilado o punto Verde
Public Property Get ArticuloRecargoFinanciero() As String
     ArticuloRecargoFinanciero = mvarArticuloRecargoFinanciero
End Property

Public Property Let ArticuloRecargoFinanciero(ByVal vData As String)
     mvarArticuloRecargoFinanciero = vData
End Property


' ---- [13/10/2010] (DAVID)
'Tipo Portes:       Por programa, si artiuloportes="" entonces meto un -1 al tipoportes(aunque en la BD tenga lo que tenga
'       0- FONTENAS
'       1- HERBELCA
'CUIDADO A CAMBIAR ESTA VARIABLE
' Para el LOG de impresion de albaranes  UTILIZO PrecioMinimo
' junto con mvarTipoPortes
Public Property Get TipoPortes() As Integer
     TipoPortes = mvarTipoPortes
End Property

Public Property Let TipoPortes(ByVal vData As Integer)
     mvarTipoPortes = vData
End Property




' ---- [14/10/2010] (DAVID)
Public Property Get DireccionesEnvio() As Boolean
     DireccionesEnvio = mvarDireccionesEnvio
End Property

Public Property Let DireccionesEnvio(ByVal vData As Boolean)
     mvarDireccionesEnvio = vData
End Property


' ---- [08/11/2010] (DAVID)
Public Property Get FrasMostradorSerieDistinta() As Boolean
     FrasMostradorSerieDistinta = mvarFrasMostradorSerieDistinta
End Property

Public Property Let FrasMostradorSerieDistinta(ByVal vData As Boolean)
     mvarFrasMostradorSerieDistinta = vData
End Property


'[26/11/2010] [DAVID]: Al crear el albarn marcarlo con el check de ParaFacturar
Public Property Get MarcarAlbaranFacturar() As Boolean
     MarcarAlbaranFacturar = mvarMarcarAlbaranFacturar
End Property

Public Property Let MarcarAlbaranFacturar(ByVal vData As Boolean)
     mvarMarcarAlbaranFacturar = vData
End Property







'[17/12/2010] [DAVID]: lleva operaciones aseguradas
Public Property Get OperacionesAseguradas() As Boolean
     OperacionesAseguradas = mvarOperacionesAseguradas
End Property

Public Property Let OperacionesAseguradas(ByVal vData As Boolean)
     mvarOperacionesAseguradas = vData
End Property

'[17/12/2010] [DAVID]: Si cuabndo supera el riesgo y bloqueamos... cual es esa incide
Public Property Get SituacionBloqueoOpAseg() As Integer
     SituacionBloqueoOpAseg = mvarSituacionBloqueoOpAseg
End Property

Public Property Let SituacionBloqueoOpAseg(ByVal vData As Integer)
     mvarSituacionBloqueoOpAseg = vData
End Property



'------------------------------------------------
'------------------------------------------------
'   Rotacion
Public Property Get Rot_ConsumMes1() As Currency
     Rot_ConsumMes1 = mvarRot_ConsumMes1
End Property

Public Property Let Rot_ConsumMes1(ByVal vData As Currency)
     mvarRot_ConsumMes1 = vData
End Property
Public Property Get Rot_ConsumMes2() As Currency
     Rot_ConsumMes2 = mvarRot_ConsumMes2
End Property
Public Property Let Rot_ConsumMes2(ByVal vData As Currency)
     mvarRot_ConsumMes2 = vData
End Property
Public Property Get Rot_ConsumMesMax() As Currency
     Rot_ConsumMesMax = mvarRot_ConsumMesMax
End Property
Public Property Let Rot_ConsumMesMax(ByVal vData As Currency)
     mvarRot_ConsumMesMax = vData
End Property
Public Property Get Rot_ConsumMesMin() As Currency
     Rot_ConsumMesMin = mvarRot_ConsumMesMin
End Property

Public Property Let Rot_ConsumMesMin(ByVal vData As Currency)
     mvarRot_ConsumMesMin = vData
End Property

'[11/01/2011] [DAVID]: Log de cambio de precios o dtos
Public Property Get LogCambioPrecDto() As Boolean
     LogCambioPrecDto = mvarLogCambioPrecDto
End Property

Public Property Let LogCambioPrecDto(ByVal vData As Boolean)
     mvarLogCambioPrecDto = vData
End Property


Public Property Get AlmacenB() As Integer
     AlmacenB = mvarAlmacenB
End Property



'***********************************************
'FRAS COMPRAS INTRCOMUNITARIAS
'
' Cunado se contabiliza una frapro intracounitaria, la del proveedor entrara con el IVA intracom 0
' y luego se crearan en la contb DOS fras mas (1cli+1pro) con un IVA <>0 (IvaIntracomAdicional)
' la cuenta base para esas fras sera o bien una especidicada en CtaContabIntracom o la del propio proveedor
' la serie de factura de cliente en conta lo traera de lo que ponga en stipom  para 'CFI'
' y en la contabilzacion se marcara en umasien un cero o null segun el parametro:IntracomAdicionalContab
Public Property Get IvaIntracomAdicional() As Integer
     IvaIntracomAdicional = mvarIvaIntracomAdicional
End Property
Public Property Let IvaIntracomAdicional(ByVal vData As Integer)
     mvarIvaIntracomAdicional = vData
End Property
'-------------------------------------------------------------
Public Property Get CtaContabIntracom() As String
     CtaContabIntracom = mvarCtaContabIntracom
End Property
Public Property Let CtaContabIntracom(ByVal vData As String)
     mvarCtaContabIntracom = vData
End Property
'---------------------------------------------------------
Public Property Get IntracomAdicionalContab() As Boolean
     IntracomAdicionalContab = mvarIntracomAdicionalContab
End Property
Public Property Let IntracomAdicionalContab(ByVal vData As Boolean)
     mvarIntracomAdicionalContab = vData
End Property


'-------------------------------------------------------------
Public Property Get PathFacturaE() As String
     PathFacturaE = mvarPathFacturaE
End Property
Public Property Let PathFacturaE(ByVal vData As String)
     mvarPathFacturaE = vData
End Property


'----------------------------------------------------------------------
' GETS para los numeros de copia
'      1ofer  2pedido  3ped_ruta  4alb  5alb_ruta  6FraRect
Public Property Get NumeroCopias() As String
     NumeroCopias = mvarNumeroCopias
End Property
Public Property Let NumeroCopias(ByVal vData As String)
     mvarNumeroCopias = vData
End Property
' ofer ped ped_ruta alb alb_ruta frrect
Public Property Get NumCop_AlbaranNormal() As Integer
    NumCop_AlbaranNormal = DevuelveNumeroCopias(4)
End Property
Public Property Get NumCop_AlbaranRuta() As Integer
    NumCop_AlbaranRuta = DevuelveNumeroCopias(5)
End Property
Public Property Get NumCop_Oferta() As Integer
    NumCop_Oferta = DevuelveNumeroCopias(1)
End Property
Public Property Get NumCop_Pedido() As Integer
    NumCop_Pedido = DevuelveNumeroCopias(2)
End Property
Public Property Get NumCop_PedZona() As Integer
    NumCop_PedZona = DevuelveNumeroCopias(3)
End Property
Public Property Get NumCop_FraRectifica() As Integer
    NumCop_FraRectifica = DevuelveNumeroCopias(6)
End Property


'CUIDADO A CAMBIAR ESTA VARIABLE
' Para el LOG de impresion de albaranes  UTILIZO PrecioMinimo
' junto con mvarTipoPortes
Public Property Get PrecioMinimo() As Boolean
     PrecioMinimo = mvarPrecioMinimo
End Property
Public Property Let PrecioMinimo(ByVal vData As Boolean)
     mvarPrecioMinimo = vData
End Property

'Articulo separacion ofertas
Public Property Get ArtSeparador() As String
     ArtSeparador = mvarArtSeparador
End Property

Public Property Let ArtSeparador(ByVal vData As String)
     mvarArtSeparador = vData
End Property

'Conexion Ariagro. SOLO GET
Public Property Get Ariagro() As String
     Ariagro = mvarAriagro
End Property



Public Property Get RecalculoMargen() As Boolean
     RecalculoMargen = mvarRecalculoMargen
End Property
Public Property Let RecalculoMargen(ByVal vData As Boolean)
     mvarRecalculoMargen = vData
End Property

'#Solo GET
'
' Si tiene telefonia, puede tener, o no, el modulo de venta telefonos a plazos
Public Property Get TelefoniaVtaPlazos() As Boolean
     TelefoniaVtaPlazos = mvarTelefoniaVtaPlazos
End Property



'Dic 2012     EN TPV tb vamos a poder añadir un campo a un tiket.
'Para ello en ariagro tendremos el nombre de la bD, como hasta ahora, pero
' para saber si tienen ADV veremos si tiene algun valor en la tabla de  advparametros
 'para eso .
' 0.- Variable sin establecer
' 1.- Variable establecida NO tiene ADVparametros
' 2.-   "                  SI
Public Function LlevaADV() As Boolean
    
    LlevaADV = False
    If mvarAriagro = "" Then Exit Function
    

    If TieneADV = 0 Then
        SQL = DevuelveDesdeBD(conAri, "count(*)", "advparametros", "1", "1")
        If SQL = "" Then SQL = "0"
        If Val(SQL) > 0 Then
            TieneADV = 2
        Else
            TieneADV = 1
        End If
    End If
    
    LlevaADV = TieneADV > 1
    
End Function


Public Property Get Renting() As Boolean
     Renting = mvarRenting
End Property


Public Property Get GestionFlotas() As Boolean
     GestionFlotas = mvarGestionFlotas
End Property

'No tienen coluimna en La BD
Public Property Get ComunicaAlmagrupo() As Boolean
     ComunicaAlmagrupo = mvarComunicaAlmagrupo
End Property



'Porcentaje de dias a partir del cual un talon / pagare se considerara devuelto
Public Property Get PorcenPagoAgentTalPag() As Currency
     PorcenPagoAgentTalPag = mvarPorcenPagoAgentTalPag
End Property

Public Property Let PorcenPagoAgentTalPag(ByVal vData As Currency)
     mvarPorcenPagoAgentTalPag = vData
End Property


'Solo get
Public Property Get FechaCambioIva() As Date
     FechaCambioIva = mvarFechaCambioIva
End Property
Public Property Get ClientesPotenciales() As Boolean
     ClientesPotenciales = mvarClientesPotenciales
End Property
Public Property Get SituaEnCodigoArticulo() As Boolean  'Este NO esta en BD. Cuando entre en el programa se establecera leyendo Dir(..\SituCod.dat
     SituaEnCodigoArticulo = mvarSituaEnCodigoArticulo
End Property


Public Property Let ArtiTelefonia(ByVal vData As String)
     mvarArtiTelefonia = vData
End Property
Public Property Get ArtiTelefonia() As String
     ArtiTelefonia = mvarArtiTelefonia
End Property



Public Property Let RecMercanciaSoloPpal(ByVal vData As Boolean)
     mvarRecMercanciaSoloPpal = vData
End Property
Public Property Get RecMercanciaSoloPpal() As Boolean
     RecMercanciaSoloPpal = mvarRecMercanciaSoloPpal
End Property



Public Property Let ArtTfniaIvaExento(ByVal vData As String)
     mvarArtTfniaIvaExento = vData
End Property
Public Property Get ArtTfniaIvaExento() As String
     ArtTfniaIvaExento = mvarArtTfniaIvaExento
End Property


Public Property Get GrabaModificarPrecioAlaBaja() As Boolean
     GrabaModificarPrecioAlaBaja = mvarGrabaModificarPrecioAlaBaja
End Property


Public Property Get ComprobarBancoRestoAplicaciones() As Boolean
    '0.- ARIGES y contariges
    '1.- Todas las aplicaciones
    ComprobarBancoRestoAplicaciones = mvarComprobarBancoRestoAplicaciones
End Property

'---------------------------------------------------------
'  0. Resto     1.- WHOSE      2.- GESSOCIAL
Public Property Get QueEmpresaEs() As Byte
    '0.- ARIGES y contariges
    '1.- Todas las aplicaciones
    QueEmpresaEs = mvarQueEmpresaEs
End Property


Public Property Let IncializarStockEnInventario(ByVal vData As Boolean)
     mvarIncializarStockEnInventario = vData
End Property
Public Property Get IncializarStockEnInventario() As Boolean
     IncializarStockEnInventario = mvarIncializarStockEnInventario
End Property





'NO lleva campo EN BD, ver descripcion above
Public Property Let PathDocsWHOSE(ByVal vData As String)
     mvarPathDocsWHOSE = vData
End Property
Public Property Get PathDocsWHOSE() As String
     PathDocsWHOSE = mvarPathDocsWHOSE
End Property



'Solo GET
Public Property Get AguasPotables() As String
     AguasPotables = mvarAguasPotables
End Property


'Si solo tiene un banco propio, este valor sera >0
'Si es cero=tiene varios bancos propios
Public Property Get CodigoUnicoBancoPropio() As String
    CodigoUnicoBancoPropio = mvarCodigoUnicoBancoPropio_
End Property


'Cuando actualiza el precio de un articulo, si actualiza tb el precio especial
Public Property Let ActualizaPrecioEspecial(ByVal vData As Boolean)
     mvarActualizaPrecioEspecial = vData
End Property
Public Property Get ActualizaPrecioEspecial() As Boolean
     ActualizaPrecioEspecial = mvarActualizaPrecioEspecial
End Property


Public Property Let EntradaRapidaFacturasMostrador(ByVal vData As Boolean)
     mvarEntradaRapidaFacturasMostrador = vData
End Property
Public Property Get EntradaRapidaFacturasMostrador() As Boolean
     EntradaRapidaFacturasMostrador = mvarEntradaRapidaFacturasMostrador
End Property


Public Property Let InventarioCodigoArticulo(ByVal vData As Boolean)
     mvarInventarioCodigoArticulo = vData
End Property
Public Property Get InventarioCodigoArticulo() As Boolean
     InventarioCodigoArticulo = mvarInventarioCodigoArticulo
End Property


'Path facturas firmadas signotech
Public Property Let PathFirmasAlbaran(ByVal vData As String)
     mvarPathFirmasAlbaran = vData
End Property
Public Property Get PathFirmasAlbaran() As String
     PathFirmasAlbaran = mvarPathFirmasAlbaran
End Property

Public Property Let PathFirmasFacturas(ByVal vData As String)
     mvarPathFirmasFacturas = vData
End Property
Public Property Get PathFirmasFacturas() As String
     PathFirmasFacturas = mvarPathFirmasFacturas
End Property




















Public Property Get ManipuladorFitosanitarios2() As Boolean
     ManipuladorFitosanitarios2 = mvarManipuladorFitosanitarios
End Property

Public Property Get InvSujetoPasivo() As Boolean
     InvSujetoPasivo = mvarInvSujetoPasivo
End Property


Public Property Get Huertos() As Boolean
     Huertos = mvarHuertos
End Property

Public Property Get LotesGeneralitat() As Boolean
     LotesGeneralitat = mvarLotesGeneralitat
End Property




'Indicara la empresa en la que estamos
' 0 Sin especificar     1.- ALZIRA    2.- HERBELCA   3.-TEinsa      4.- EULER
' 5 FONTENAS            6.-FENOLLAR   7.- AMESA
Public Property Get NumeroInstalacion() As Byte
     NumeroInstalacion = mvarNumeroInstalacion
End Property


Public Property Get ContabilidadNueva() As Boolean
     ContabilidadNueva = mvarContabilidadNueva
End Property


Public Property Get SII_Tiene() As Boolean
     SII_Tiene = mvarSIITiene
End Property

Public Property Get Sii_Finicio() As Date
     Sii_Finicio = mvarSIIFechaIni
End Property
Public Property Get Sii_Dias() As Integer
     Sii_Dias = mvarSIIDias
End Property





'Private mvarPtosAsignar As Currency  'Si lleva >0 tendra sistema de puntos
'Private mvarPtosImporteCalculo As Currency
Public Property Get PtosAsignar() As Currency
     PtosAsignar = mvarPtosAsignar
End Property
Public Property Get PtosImporteCalculo() As Currency
     PtosImporteCalculo = mvarPtosImporteCalculo
End Property
Public Property Get PtosEquivalencia() As Currency
     PtosEquivalencia = mvarPtosEquivalencia
End Property
Public Property Get PtosArticuloCanje() As String
     PtosArticuloCanje = mvarPtosArticuloCanje
End Property
Public Property Get PtosFechaIncio() As Date
     PtosFechaIncio = mvarPtosFechaIncio
End Property



Public Property Get PrecioHoraCosteEUL() As Currency
     PrecioHoraCosteEUL = mvarPrecioHoraCosteEUL
End Property
Public Property Let PrecioHoraCosteEUL(ByVal CosteHOra As Currency)
     mvarPrecioHoraCosteEUL = CosteHOra
End Property


Public Property Get CartaPortes() As Boolean
     CartaPortes = mvarCartaPortes
End Property


Public Property Get HaciendoFrmulariosGrandes() As Boolean
     HaciendoFrmulariosGrandes = mvarHaciendoFrmulariosGrandes
End Property


'Impresoaras ALBARANES Fenollar
Public Property Get ImpresoraFenollarA() As String
     ImpresoraFenollarA = mvarImpresoraFenollarA
End Property
Public Property Get ImpresoraFenollarB() As String
     ImpresoraFenollarB = mvarImpresoraFenollarB
End Property



Public Property Get MaximoImporteEfectivo() As String
     MaximoImporteEfectivo = mvarMaximoImporteEfectivo
End Property



Public Property Get PuedeModificarAriconta() As Boolean
     PuedeModificarAriconta = mvarPuedeModificarAriconta
End Property


'Lleva produccion y compon
Public Property Get TieneComponentes_y_Produccion() As Boolean
     TieneComponentes_y_Produccion = mvarTieneComponentes_y_Produccion
End Property


Public Property Get ImportacionesCoarval() As Boolean
    ImportacionesCoarval = mvarImportacionesCoarval
End Property
     
     
'llevan GET y let
Public Property Get TieneGasolinera() As Boolean
     TieneGasolinera = mvarLlevaGasolinera
End Property


Public Property Get DigitosCodartic() As Byte
     DigitosCodartic = mvarDigitosCodartic
End Property




'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------


' Metodos.



'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------


'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
Dim RS As ADODB.Recordset
Dim V As Variant
On Error GoTo Err3
    
    SQL = "SELECT * FROM spara1 "
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        Leer = 1
    Else
        mvarTipoDtos = RS!TipoDtos
        mvarOrdenDtos = RS!ordendto
        mvarCodTarif = RS!codTarif
        mvarCajaComp = RS!cajacomp  'Cajas completas precios (S/N)
        mvarPreuKmCli = DBLet(RS!preukmcl, "N") 'Precio Km Desplazamiento Clientes
        mvarPreuKmTec = DBLet(RS!preukmtc, "N") 'Precio Km Desplazamiento Técnicos
        mvarnomGerente = DBLet(RS!nomgeren, "T") 'Nombre del Director Gerente
        mvarnomAdmin = DBLet(RS!nomadmin, "T") 'Nombre del Responsable Administración
        mvarHayMante = RS!haymante  'Si hay Mantenimiento (S/N)
        mvarHayRepar = RS!hayrepar  'Si hay Reparaciones (S/N)
        mvarHayFrecu = RS!hayfrecu  'Si hay Frecuencias (S/N)
        mvarHayServicio = RS!hayservicio  'Si hay Servicios (S/N)
        mvarHayDeparNuevo = RS!haydepar  'Si hay Departamento o Direccion
        mvarCtrStock = RS!CtrStock  'Si hay Control de Stock (S/N)
        mvarInventar = RS!inventar
        mvarHaynserie = RS!haynserie 'Si hay Nº Serie en Compras (S/N)
        mvarDiasSiMan = RS!diassiman 'Dias reparacion si hay Mantenimiento
        mvarDiasNoMan = RS!diasnoman 'Dias reparacion si NO hay Mantenimiento
        Me.ForPagoChequeRegalo = DBLet(RS!codforpa, "T")
        
        'Datos Facturacion
        mvarCodArticDes = DBLet(RS!codartid, "T") 'articulo para facturar mantenimientos
        mvarDiaPago1 = DBLet(RS!DiaPago1, "N") 'dia pago 1 para compras
        mvarDiaPago2 = DBLet(RS!DiaPago2, "N") 'dia pago 2 para compras
        mvarDiaPago3 = DBLet(RS!DiaPago3, "N") 'dia pago 3 para compras
        mvarMesnogir = DBLet(RS!mesnogir, "N") 'mes a no girar para compras
        
        'Datos de Internet
        mvarDireMail = DBLet(RS!DireMail, "T") 'Direccion de email
        mvarSMTPhost = DBLet(RS!SMTPhost, "T") 'SMTP host para el envio de email
        mvarSMTPuser = DBLet(RS!SMTPuser, "T") 'SMTP usuario para el envio de email
        mvarSMTPpass = DBLet(RS!SMTPpass, "T") 'SMTP password para el envio de email
        mvarWebSoporte = DBLet(RS!WebSoporte, "T") 'Web de soporte
        mvarMailSoporte = DBLet(RS!MailSoporte, "T") 'email de soporte
        mvarWebVersion = DBLet(RS!WebVersion, "T") 'Version de la Web
        
        
        'Datos Contabilidad
        mvarSerConta = DBLet(RS!serconta, "T") 'servidor conta
        mvarUsuConta = DBLet(RS!UsuConta, "T") 'usuario conta
        mvarPasConta = DBLet(RS!pasconta, "T") 'password
        mvarNumConta = DBLet(RS!numconta, "N") 'numero conta
        
        
        'Valores por defecto
        mvarDefActiv = DBLet(RS!defactividad, "N") 'numero conta
        mvarDefEnvio = DBLet(RS!defenvio, "N") 'numero conta
        mvarDefZona = DBLet(RS!defzona, "N") 'numero conta
        mvarDefRuta = DBLet(RS!defruta, "N") 'numero conta
        mvarDefSitu = DBLet(RS!defstituacion, "N") 'numero conta
        mvarDefTarifa = DBLet(RS!deftarifa, "N") 'numero conta
        mvarDefAgente = DBLet(RS!defagente, "N") 'numero conta


        mvaravipedcli = DBLet(RS!avipedcli, "N")
        mvaravipedpro = DBLet(RS!avipedpro, "N")
        mvaravialbcli = DBLet(RS!avialbcli, "N")
        mvaravialbpro = DBLet(RS!avialbpro, "N")
        mvaravimanteni = DBLet(RS!avimanteni, "N")
        mvaravirepara = DBLet(RS!avirepara, "N")
        mvaraviavisos = DBLet(RS!aviavios, "N")
        PonVariableTienAvisos

        
        

        mvarTipoIVAre1 = DBLet(RS!IVAre1eq, "N")
        mvarTipoIVAre2 = DBLet(RS!IVAre2eq, "N")
        mvarTipoIVAre3 = DBLet(RS!IVAre3eq, "N")
        mvarRecargoEquivalencia = (mvarTipoIVAre1 > 0 Or mvarTipoIVAre2 > 0 Or mvarTipoIVAre3 > 0)
        
        mvarIVA1 = DBLet(RS!IVAre1, "N")
        mvarIVA2 = DBLet(RS!IVAre2, "N")
        mvarIVA3 = DBLet(RS!IVAre3, "N")

        'El articulo de telefonia no lo pongo como parametros
        'codartictel
        mvarCodarticTfnia = DBLet(RS!codartictel, "T") 'numero conta
                
                
        'Retencion
        mvarIVAREA = DBLet(RS!IVA_REA, "N")
        mvarCtaReten = DBLet(RS!CtaReten, "T")
        mvarPorReten = DBLet(RS!PorReten, "N")
                
                
        'IVA exento
        mvarIVA_Exento = DBLet(RS!ivaexento, "N")
        
        
        
        'Contabilidad
        mvarContabilidadB = DBLet(RS!conta_B, "N")
        
        'Contabiliza los tickets resumidos
        mvarContabilizarTicketAgrupados = (DBLet(RS!conttickagrupado, "N") = 1)
        
        'Cta aportacion
        mvarctaaportacion = DBLet(RS!ctaAportacion, "T")
        
        'produccion
        mvarProduccion = DBLet(RS!Produccion, "N") = 1
        'descriptores
        mvarDescriptores = DBLet(RS!Descriptores, "N") = 1
        'obs factura
        mvarObsFactura = DBLet(RS!ObsFactura, "N")
        '
        mvarIVA_Intracomunitario = DBLet(RS!IvaIntracom, "N")
        'Recliclado Punto verde
        mvarArtReciclado = DBLet(RS!ArtReciclado, "T")

        mvarArtPortes = DBLet(RS!ArticuloPortes, "T")
        mvarImporteMinimo = DBLet(RS!impminped, "N")
        mvarAbonoKilos = DBLet(RS!AbonoKilos, "N")
    
        mvarComponentePorcentaje = DBLet(RS!compoporcen, "N") = 1

        ' ---- [14/09/2009] (LAURA)
        Me.DtoxCantidad = IIf(RS!dtoxcanti = 1, True, False)
        Me.CreaTarifasArticulo = DBLet(RS!creatarifart, "N")
        ' ----
        
        ' ---- [19/10/2009] [LAURA]: añadir campo modo analitica
        Me.ModoAnalitica = RS!modanalitica
        
        
        'David 31/12/09
        Me.EnvioDesdeOutlook = DBLet(RS!EnvioDesdeOutlook, "N") = 1
        
        'David Febrero 2010
        ExeEnvioMail = DBLet(RS!arigesmail, "T")
        DiasGarantia = DBLet(RS!diasgaranrepa, "N")
        
        mvarTieneCRM = DBLet(RS!TieneCRM, "N") = 1
        
        
        'David. Junio 2010.  Cambio IVA   EL IVA superreducido NO cambia
        mvarOLDIVAre1 = DBLet(RS!iva_oldre1, "N")
        mvarOLDIVAre2 = DBLet(RS!iva_oldre2, "N")
        'mvarOLDIVAre3 = DBLet(RS!iva_oldre3, "N")
        
        mvarOLDIVA1 = DBLet(RS!iva_old1, "N")
        mvarOLDIVA2 = DBLet(RS!iva_old2, "N")
        'mvarOLDIVA3 = DBLet(RS!iva_old3, "N")

        mvarNumCopiasFacturacion = DBLet(RS!NumCopias, "N")
        If mvarNumCopiasFacturacion = 0 Then mvarNumCopiasFacturacion = 1
        
        'De momento NO lo tengo en la BD
        mvarTipoFormularioClientes = 0
        
        mvarArticuloRecargoFinanciero = DBLet(RS!artRecargoFina, "T")
        
        'tIPO PORTES
        If mvarArtPortes = "" Then
            mvarTipoPortes = -1
        Else
            SQL = DBLet(RS!TipoPortes, "T")
            If SQL = "" Then SQL = "-1"
            mvarTipoPortes = CInt(SQL)
        End If
        
        mvarDireccionesEnvio = (DBLet(RS!DirecEnvio, "N") = 1)
        
        
        mvarFrasMostradorSerieDistinta = (DBLet(RS!FraMostra, "N") = 1)
        
        mvarMarcarAlbaranFacturar = (DBLet(RS!AlbParaFcturar, "N") = 1)
        
        '`OperAseguradas`    `situbloq`
        mvarOperacionesAseguradas = (DBLet(RS!OperAseguradas, "N") = 1)
        mvarSituacionBloqueoOpAseg = DBLet(RS!situbloq, "N")

        'Rotacion
        'alter table `spara1` `mesconant1`  `mesconant2` `consummin`  `consummax`
        mvarRot_ConsumMes1 = DBLet(RS!mesconant1, "N")
        mvarRot_ConsumMes2 = DBLet(RS!mesconant2, "N")
        mvarRot_ConsumMesMin = DBLet(RS!consummin, "N")
        mvarRot_ConsumMesMax = DBLet(RS!consummax, "N")
        
        mvarLogCambioPrecDto = DBLet(RS!LogCambioPrecDto, "N") = 1
        
        'No esta en el form
        mvarAlmacenB = -1
        If Not IsNull(RS!AlmacenB) Then mvarAlmacenB = RS!AlmacenB
        
        'IVA intracomuntario para fras proveedor
        mvarIvaIntracomAdicional = DBLet(RS!IvaIntracomAdicional, "N")
        mvarCtaContabIntracom = DBLet(RS!CtaContabIntracom, "T")
        mvarIntracomAdicionalContab = DBLet(RS!IntracomAdicionContab, "N") = 1

        
        mvarPathFacturaE = DBLet(RS!PathFacturaE, "T")
        
        mvarNumeroCopias = DBLet(RS!NumeroCopias, "T")
        
        mvarPrecioMinimo = DBLet(RS!PrecioMinimo, "N") = 1
        mvarArtSeparador = DBLet(RS!ArtSeparador, "T")
        mvarAriagro = DBLet(RS!Ariagro, "T")
        
        
        mvarRenting = DBLet(RS!Renting, "N") = 1
        mvarGestionFlotas = DBLet(RS!GestionFlotas, "N") = 1
        
        mvarPorcenPagoAgentTalPag = DBLet(RS!PorcenpagAgenPag, "N")
        
        mvarFechaCambioIva = RS!FechaCambioIva 'NO ES NULL
        mvarClientesPotenciales = DBLet(RS!ClientesPotenciales, "N") = 1
        
        'Solo tiene GET
        mvarTieneTelefonia2 = DBLet(RS!Telefonia, "N")
        mvarArtiTelefonia = DBLet(RS!codartictel, "T")
        

        mvarRecMercanciaSoloPpal = DBLet(RS!RecMercanciaSoloPpal, "N") = 1
        
        mvarArtTfniaIvaExento = DBLet(RS!artTfoniaIvaExento, "T") 'articulo telefonia IVA exento
        
        
        'Solo tiene GET
        mvarComprobarBancoRestoAplicaciones = DBLet(RS!ComprobarCCC, "N")
        mvarAguasPotables = DBLet(RS!Agua, "N") = 1
        
        
        mvarIncializarStockEnInventario = DBLet(RS!IncializarStokInv, "N")
        
        
        mvarActualizaPrecioEspecial = DBLet(RS!ActualizaPrecioEspecial, "N") = 1
        
        
        mvarEntradaRapidaFacturasMostrador = DBLet(RS!FrasMostradorRapida, "N") = 1
        
        'Solo tiene GET
        mvarNumeroInstalacion = DBLet(RS!NumInstalacion, "N")
        
        
        mvarManipuladorFitosanitarios = DBLet(RS!ManipuladorFitosanitarios, "N") = 1
        
        mvarInvSujetoPasivo = DBLet(RS!InvSujetoPasivo, "N") = 1
    
        mvarRecalculoMargen = DBLet(RS!ActualizaMargen, "N") = 1
    
        
        mvarInventarioCodigoArticulo = DBLet(RS!InventarioPorCodigo, "N") = 1
        
        
        If Me.HayDeparNuevo = 2 Then mvarTipoFormularioClientes = 1  'Obras. SAIL-Y euler
        

        mvarPathFirmasFacturas = ""
        If CampoNuevoBD(RS, "PathFirmasFacturas", "T", V) Then mvarPathFirmasFacturas = V
        
        mvarPathFirmasAlbaran = ""
        If CampoNuevoBD(RS, "PathFirmasAlbaran", "T", V) Then mvarPathFirmasAlbaran = V

        
        'Solo GET
        mvarContabilidadNueva = DBLet(RS!NuevaContabilidad, "N") = 1
        
        
        mvarSIITiene = False
        
        
        
        mvarPtosAsignar = 0
        mvarPtosImporteCalculo = 0
        If Not IsNull(RS!PtosAsignar) Then
            mvarPtosAsignar = RS!PtosAsignar
            mvarPtosImporteCalculo = DBLet(RS!PtosImporteCalculo, "N")
            mvarPtosEquivalencia = RS!PtosEquivalencia
            mvarPtosArticuloCanje = RS!PtosArticuloCanje
            mvarPtosFechaIncio = RS!PtosFechaIncio
        End If
        
        mvarPrecioHoraCosteEUL = DBLet(RS!preciohoracoste, "N")
        
        
        'Si tiene telefonia entoces vere si tiene el modulo de venta a plazos
        mvarTelefoniaVtaPlazos = False
        If mvarTieneTelefonia2 > 0 Then ModuloTelefoniaVtaPlazos
        
        
        ComprobarProgramaEnvioMail
        
        
        
        'Como pone en la definicion, este parametro no va a BD, lleva el de precio minimo(HERBEL)
        mvarGrabaModificarPrecioAlaBaja = mvarPrecioMinimo
        
        
        '
        mvarCartaPortes = False
        If CampoNuevoBD(RS, "cartaPortes", "B", V) Then mvarCartaPortes = CBool(V)
        
        mvarHaciendoFrmulariosGrandes = False
        If CampoNuevoBD(RS, "HaciendoFrmulariosGrandes", "B", V) Then mvarHaciendoFrmulariosGrandes = CBool(V)
        
        mvarPuedeModificarAriconta = False   'Modificara o eliminara en contabilidad   SOLO conta nueva   De momento Fenollar
        mvarTieneComponentes_y_Produccion = False
        
        If mvarNumeroInstalacion = vbFenollar Then
            mvarImpresoraFenollarA = "\\PCPRUEBAS\CARLA"
            mvarImpresoraFenollarB = "\\PCMAJOSE\MAJO"
            mvarPuedeModificarAriconta = True
            mvarTieneComponentes_y_Produccion = True
        End If
        
        mvarMaximoImporteEfectivo = 2500
        If CampoNuevoBD(RS, "ImportMaxEfectivo", "N", V) Then mvarMaximoImporteEfectivo = DBLet(V)
        
        mvarLlevaGasolinera = False
        mvarDigitosCodartic = 0
        If MovimientosParaGasolinera Then
            mvarLlevaGasolinera = True
            mvarDigitosCodartic = 13
        End If
        
        
        If mvarNumeroInstalacion = vbTaxco Then mvarArtReciclado = "SI"
        
        
        
        Leer = 0
        
    End If

    RS.Close
    
    If Leer = 0 Then
        FijaParametroAlmagrupo RS  'hacer rs=nothing
        
        mvarQueEmpresaEs = 0
        If ComprobarEmpresaWHOSE(RS) Then
            mvarQueEmpresaEs = 1
        ElseIf ComprobarGessocial(RS) Then mvarQueEmpresaEs = 2
            
        End If
        
        
        'Navarres(aunque sirve para todos)
        'Si solo tiene un banco por defecto, sera ese el cargara
        FijarSiTieneUnSoloBancoPropio RS
        
        
        
        'Marjal CHIPOS, para saber si es marjal sin ponerle un numero instalacion
        FijarSiTieneHuertos RS
        
        
        'Lotes generalitat. Para meter los datos en
        FijarSiTieneLotesGeneralitat RS
        
        'Importaciones COARVAL
        FijarSiTieneImporCoarval RS
        
    End If
    
    Set RS = Nothing
    
    
    Exit Function

Err3:
    Mens = "Se ha producido un error(Leer Parametros)" & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Set RS = Nothing
    Leer = 1
End Function


Public Function Modificar(Codigo As Byte) As Byte
'Modifica los parametros de la Aplicación
'Modifica la Tabla: spara1, BD: Ariges
On Error GoTo EModificar

    SQL = "UPDATE spara1 SET "
    SQL = SQL & " tipodtos= " & mvarTipoDtos
    SQL = SQL & ", ordendto= " & mvarOrdenDtos
    SQL = SQL & ", codtarif= " & mvarCodTarif
    SQL = SQL & ", cajacomp= " & DBSet(mvarCajaComp, "B")
    SQL = SQL & ", preukmcl= " & DBSet(mvarPreuKmCli, "N")
    SQL = SQL & ", preukmtc= " & DBSet(mvarPreuKmTec, "N")
    SQL = SQL & ", nomgeren= " & DBSet(mvarnomGerente, "T")
    SQL = SQL & ", nomadmin= " & DBSet(mvarnomAdmin, "T")
    SQL = SQL & ", haymante = " & DBSet(mvarHayMante, "B")
    SQL = SQL & ", hayrepar = " & DBSet(mvarHayRepar, "B")
    SQL = SQL & ", hayfrecu = " & DBSet(mvarHayFrecu, "B")
    SQL = SQL & ", hayservicio = " & DBSet(mvarHayServicio, "B")
    SQL = SQL & ", haydepar = " & DBSet(mvarHayDeparNuevo, "N")
    SQL = SQL & ", ctrstock = " & DBSet(mvarCtrStock, "B")
    SQL = SQL & ", inventar= " & DBSet(mvarInventar, "B")
    SQL = SQL & ", haynserie= " & DBSet(mvarHaynserie, "B")
    SQL = SQL & ", diassiman= " & DBSet(mvarDiasSiMan, "N")
    SQL = SQL & ", diasnoman= " & DBSet(mvarDiasNoMan, "N")
    SQL = SQL & ", codartid= " & DBSet(mvarCodArticDes, "T")
    SQL = SQL & ", diapago1= " & DBSet(mvarDiaPago1, "N")
    SQL = SQL & ", diapago2= " & DBSet(mvarDiaPago2, "N")
    SQL = SQL & ", diapago3= " & DBSet(mvarDiaPago3, "N")
    SQL = SQL & ", mesnogir= " & DBSet(mvarMesnogir, "N")
    SQL = SQL & ", codforpa= " & DBSet(mvarCodForPa, "N", "S")
    SQL = SQL & ", diremail= " & DBSet(mvarDireMail, "T")
    SQL = SQL & ", smtphost= " & DBSet(mvarSMTPhost, "T")
    SQL = SQL & ", smtpuser= " & DBSet(mvarSMTPuser, "T")
    SQL = SQL & ", smtppass= " & DBSet(mvarSMTPpass, "T")
    SQL = SQL & ", websoporte= " & DBSet(mvarWebSoporte, "T")
    SQL = SQL & ", mailsoporte= " & DBSet(mvarMailSoporte, "T")
    SQL = SQL & ", webversion= " & DBSet(mvarWebVersion, "T")
    SQL = SQL & ", serconta= " & DBSet(mvarSerConta, "T")
    SQL = SQL & ", usuconta= " & DBSet(mvarUsuConta, "T")
    SQL = SQL & ", pasconta= " & DBSet(mvarPasConta, "T")
    SQL = SQL & ", numconta= " & DBSet(mvarNumConta, "N", "S")
    
    'Valore por defecto
    SQL = SQL & ", defactividad= " & DBSet(mvarDefActiv, "N", "N")
    SQL = SQL & ", defenvio= " & DBSet(mvarDefEnvio, "N", "N")
    SQL = SQL & ", defzona= " & DBSet(mvarDefZona, "N", "N")
    SQL = SQL & ", defruta= " & DBSet(mvarDefRuta, "N", "N")
    SQL = SQL & ", defstituacion= " & DBSet(mvarDefSitu, "N", "N")
    SQL = SQL & ", deftarifa= " & DBSet(mvarDefTarifa, "N", "N")
    SQL = SQL & ", defagente= " & DBSet(mvarDefAgente, "N", "N")
      
    'Telefonia
    SQL = SQL & ", codartictel= " & DBSet(mvarCodarticTfnia, "T", "S")
      
      
      
      
    'Avisos
    SQL = SQL & ", avipedcli= " & DBSet(mvaravipedcli, "N", "S")
    SQL = SQL & ", avipedpro= " & DBSet(mvaravipedpro, "N", "S")
    SQL = SQL & ", avialbcli= " & DBSet(mvaravialbcli, "N", "S")
    SQL = SQL & ", avialbpro= " & DBSet(mvaravialbpro, "N", "S")
    SQL = SQL & ", avimanteni= " & DBSet(mvaravimanteni, "N", "S")
    SQL = SQL & ", avirepara= " & DBSet(mvaravirepara, "N", "S")
    SQL = SQL & ", aviavios= " & DBSet(mvaraviavisos, "N", "S")
      
    PonVariableTienAvisos

    SQL = SQL & ", IVAre1eq= " & DBSet(mvarTipoIVAre1, "N", "S")
    SQL = SQL & ", IVAre2eq= " & DBSet(mvarTipoIVAre2, "N", "S")
    SQL = SQL & ", IVAre3eq= " & DBSet(mvarTipoIVAre3, "N", "S")
    mvarRecargoEquivalencia = (mvarTipoIVAre1 > 0 Or mvarTipoIVAre2 > 0 Or mvarTipoIVAre3 > 0)
    
    SQL = SQL & ", IVAre1= " & DBSet(mvarIVA1, "N", "S")
    SQL = SQL & ", IVAre2= " & DBSet(mvarIVA2, "N", "S")
    SQL = SQL & ", IVAre3= " & DBSet(mvarIVA3, "N", "S")

    'Retencion y REA
    SQL = SQL & ", IVA_REA = " & DBSet(mvarIVAREA, "N", "S")
    SQL = SQL & ", CtaReten = " & DBSet(mvarCtaReten, "T", "S")
    SQL = SQL & ", PorReten = " & DBSet(mvarPorReten, "N", "S")

    'IVA exento
    SQL = SQL & ", ivaexento = " & DBSet(mvarIVA_Exento, "N", "S")


    'Tickets agrupados
    SQL = SQL & ", conttickagrupado = " & DBSet(mvarContabilizarTicketAgrupados, "B")

    'Contabilidad "B". Radival
    SQL = SQL & ", conta_B  = " & DBSet(mvarContabilidadB, "N")

    'Ctaa aportacion termial
    SQL = SQL & ", ctaaportacion  = " & DBSet(mvarctaaportacion, "T")
    'Produccion
    SQL = SQL & ", Produccion  = " & DBSet(mvarProduccion, "B")
    'Descriptores
    SQL = SQL & ", descriptores = " & DBSet(mvarDescriptores, "B")
    'Observaciones factura proveedor/cliente
    SQL = SQL & ", ObsFactura  = " & DBSet(mvarObsFactura, "N")
    SQL = SQL & ", IvaIntracom  = " & DBSet(mvarIVA_Intracomunitario, "N")

    'Tasa recliclado PUNTO VERDE
    SQL = SQL & ", ArtReciclado  = " & DBSet(mvarArtReciclado, "T", "S")
    

    'FONTENAS. Portes
    SQL = SQL & ", ArticuloPortes  = " & DBSet(mvarArtPortes, "T", "S")
    SQL = SQL & ", impminped  = " & DBSet(mvarImporteMinimo, "N", "S")
    SQL = SQL & ", abonokilos  = " & DBSet(mvarAbonoKilos, "N", "S")
    SQL = SQL & ", compoporcen = " & DBSet(mvarComponentePorcentaje, "B")
    
    ' ---- [14/09/2009] (LAURA)
    SQL = SQL & ", dtoxcanti = " & DBSet(Me.DtoxCantidad, "B")
    SQL = SQL & ", creatarifart = " & DBSet(Me.CreaTarifasArticulo, "N")
    '-----------------
    
    ' ---- [19/10/2009] [LAURA]: añadir campo modo analitica
    SQL = SQL & ", modanalitica = " & DBSet(Me.ModoAnalitica, "N")
    
    
    ' ---- [31/12/2009] [DAVID]
    SQL = SQL & ", EnvioDesdeOutlook = " & DBSet(Me.EnvioDesdeOutlook, "B")
    
    
    ' ---- [02/02/2010] [DAVID]
    SQL = SQL & ", arigesmail = " & DBSet(ExeEnvioMail, "T", "S")
    ' ---- [10/02/2010] [DAVID]
    SQL = SQL & ", diasgaranrepa = " & DBSet(Me.DiasGarantia, "N", "S")
    
    
    ' ---- [09/06/2010] [DAVID] Cambio IVA
    SQL = SQL & ", iva_oldre1 = " & DBSet(mvarOLDIVAre1, "N", "S")
    SQL = SQL & ", iva_oldre2 = " & DBSet(mvarOLDIVAre2, "N", "S")
    'SQL = SQL & ", iva_oldre3 = " & DBSet(mvarOLDIVAre3, "N", "S")
    
    SQL = SQL & ", iva_old1 = " & DBSet(mvarOLDIVA1, "N", "S")
    SQL = SQL & ", iva_old2 = " & DBSet(mvarOLDIVA2, "N", "S")
    'SQL = SQL & ", iva_old3 = " & DBSet(mvarOLDIVA3, "N", "S")
    
    SQL = SQL & ",numcopias = " & DBSet(mvarNumCopiasFacturacion, "N")
    
    SQL = SQL & ",artRecargoFina = " & DBSet(mvarArticuloRecargoFinanciero, "T")
    
    
    'Tipo portes
    SQL = SQL & ",TipoPortes  = " & mvarTipoPortes
    SQL = SQL & ",DirecEnvio  = " & DBSet(mvarDireccionesEnvio, "B")
    
    'Fras mostrador serie distinta a facturas albaran
    SQL = SQL & ",FraMostra  = " & DBSet(mvarFrasMostradorSerieDistinta, "B")
    'Alb para facturar
    SQL = SQL & ",AlbParaFcturar = " & DBSet(mvarMarcarAlbaranFacturar, "B")
    
    '`OperAseguradas`    `situbloq`
    SQL = SQL & ",OperAseguradas = " & DBSet(mvarOperacionesAseguradas, "B")
    SQL = SQL & ",situbloq = " & DBSet(mvarSituacionBloqueoOpAseg, "N")

    'Rotacion
    SQL = SQL & ",mesconant1 = " & DBSet(mvarRot_ConsumMes1, "N", "S")
    SQL = SQL & ",mesconant2 = " & DBSet(mvarRot_ConsumMes2, "N", "S")
    SQL = SQL & ",consummin = " & DBSet(mvarRot_ConsumMesMin, "N", "S")
    SQL = SQL & ",consummax = " & DBSet(mvarRot_ConsumMesMax, "N", "S")
    
    SQL = SQL & ",LogCambioPrecDto = " & DBSet(mvarLogCambioPrecDto, "B")
    
    SQL = SQL & ",CtaContabIntracom = " & DBSet(mvarCtaContabIntracom, "T", "S")
    SQL = SQL & ",IvaIntracomAdicional = " & DBSet(mvarIvaIntracomAdicional, "N", "S")
    SQL = SQL & ",IntracomAdicionContab = " & DBSet(mvarIntracomAdicionalContab, "B")
   
    SQL = SQL & ",PathFacturaE = " & DBSet(mvarPathFacturaE, "T")
    
    'Nº Copias. Es para ofertas, facturas.....
    SQL = SQL & ",NumeroCopias = " & DBSet(mvarNumeroCopias, "T")
   
   
    SQL = SQL & ",PrecioMinimo = " & DBSet(mvarPrecioMinimo, "B")
    SQL = SQL & ",artSeparador = " & DBSet(mvarArtSeparador, "T")
    SQL = SQL & ",PorcenpagAgenPag = " & DBSet(mvarPorcenPagoAgentTalPag, "N", "N")
     
    SQL = SQL & ",codartictel = " & DBSet(mvarArtiTelefonia, "T")
    
    
    
    SQL = SQL & ",RecMercanciaSoloPpal = " & DBSet(mvarRecMercanciaSoloPpal, "B")
    
    SQL = SQL & ",artTfoniaIvaExento = " & DBSet(mvarArtTfniaIvaExento, "T")
            
    SQL = SQL & ", IncializarStokInv  = " & DBSet(IncializarStockEnInventario, "B")
    
    SQL = SQL & ", ActualizaPrecioEspecial  = " & DBSet(mvarActualizaPrecioEspecial, "B")
    
    SQL = SQL & ", FrasMostradorRapida  = " & DBSet(mvarEntradaRapidaFacturasMostrador, "B")
    SQL = SQL & ", ManipuladorFitosanitarios  = " & DBSet(mvarManipuladorFitosanitarios, "B")
    
    
    SQL = SQL & ", ActualizaMargen = " & DBSet(RecalculoMargen, "B")
    
    SQL = SQL & ", InventarioPorCodigo = " & DBSet(mvarInventarioCodigoArticulo, "B")

    'Coste hora euler
    SQL = SQL & ", preciohoracoste= " & DBSet(mvarPrecioHoraCosteEUL, "N", "S")
    'SIGNOTEC
    '    PathFirmasFacturas
    SQL = SQL & ", PathFirmasAlbaran= " & DBSet(mvarPathFirmasAlbaran, "T", "S")
    SQL = SQL & ", PathFirmasFacturas= " & DBSet(mvarPathFirmasFacturas, "T", "S")

    
    'Solo desde OG
    'SQL = SQL & ",Renting = " & DBSet(mvarRenting, "B")
    'GestionFlotas
    
    'where
    SQL = SQL & " WHERE codigo =" & Codigo & ";"

    conn.Execute SQL



    'Desde los ivas de retencion, hay una tabla(sivatra) que tiene la relacion entre
    ' iva sin retencion, iva con retencion
    'Entocnes cuando modifico, borro las entradas y las vuevlo a generer
    CargarTablaTraduccionIVAS
    
    
    'Todo bien
    Modificar = 1

    Exit Function

EModificar:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 0
End Function

Private Sub PonVariableTienAvisos()
    mvarTieneAvisos = (mvaravipedcli + mvaravipedpro + mvaravialbcli + mvaravialbpro + mvaravimanteni _
        + mvaravirepara + mvaraviavisos > 0)
End Sub


Private Function MovimientosParaServicios() As Boolean
    '-- Esta función comprueba si los movimientos ALS y FAS están creados
    '   este es un requisito previo si se desa activar la opción de servicios
    Dim SQL As String
    Dim RS As ADODB.Recordset
    SQL = "select count(*) from stipom where codtipom = 'ALS' or codtipom = 'FAS'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly
    If RS.Fields(0) = 2 Then
        MovimientosParaServicios = True
    Else
        MovimientosParaServicios = False
    End If
    RS.Close
     Set RS = New ADODB.Recordset
End Function


Private Function MovimientosParaGasolinera() As Boolean
    '-- Esta función comprueba si los movimientos ALS y FAS están creados
    '   este es un requisito previo si se desa activar la opción de servicios
    
    On Error GoTo eMovimientosParaGasolinera
    
    Dim SQL As String
    Dim RS As ADODB.Recordset
    SQL = "select count(*) from stipom where codtipom = 'ALD' or codtipom = 'FAD'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly
    If RS.Fields(0) = 2 Then
        MovimientosParaGasolinera = True
    Else
        MovimientosParaGasolinera = False
    End If
    RS.Close
    
    
eMovimientosParaGasolinera:
    If Err.Number <> 0 Then
        Err.Clear
        conn.Errors.Clear
    End If
    Set RS = New ADODB.Recordset
End Function



Public Function DevuleveTipoIVA_RE(TipoIVAOrigen As Integer) As Integer
Dim devuelve As Integer
    
    'IVA con recargo de equivalencia
    devuelve = TipoIVAOrigen
    If TipoIVAOrigen = TipoIVA1 Then
        devuelve = TipoIVAre1
    ElseIf TipoIVAOrigen = vParamAplic.TipoIVA2 Then
        devuelve = TipoIVAre2
    ElseIf TipoIVAOrigen = TipoIVA3 Then
        devuelve = TipoIVAre3
    Else
        MsgBox "NO se ha encontrado la correspondencia para el IVA RE: " & TipoIVAOrigen, vbExclamation
    End If
    DevuleveTipoIVA_RE = devuelve
End Function



Public Function DevuelveTipoIVA_OLD(ConRecargo As Boolean, TipoIVAOrigen As Integer, Optional OcultarMensaje As Boolean) As Integer
Dim devuelve As Integer
    
    devuelve = TipoIVAOrigen
    
    If ConRecargo Then
        'IVA con recargo de equivalencia
        If TipoIVAOrigen = TipoIVA1 Then
            devuelve = Me.OLDIVAre1
        ElseIf TipoIVAOrigen = vParamAplic.TipoIVA2 Then
            devuelve = Me.OLDIVAre2
'        ElseIf TipoIVAOrigen = TipoIVA3 Then
'            devuelve = OLDIVAre3
        Else
           If Not OcultarMensaje Then MsgBox "NO se ha encontrado la correspondencia para el ANTIGUO IVA RE: " & TipoIVAOrigen, vbExclamation
        End If
    
    
    Else
        If TipoIVAOrigen = TipoIVA1 Then
            devuelve = OLDIVA1
        ElseIf TipoIVAOrigen = vParamAplic.TipoIVA2 Then
            devuelve = OLDIVA2
'        ElseIf TipoIVAOrigen = TipoIVA3 Then
'            devuelve = OLDIVA3
        Else
            If Not OcultarMensaje Then MsgBox "NO se ha encontrado la correspondencia para el ANTIGUO IVA: " & TipoIVAOrigen, vbExclamation
        End If
    
    
    
    End If
    DevuelveTipoIVA_OLD = devuelve
End Function


Private Sub CargarTablaTraduccionIVAS()
Dim RS As ADODB.Recordset

    On Error GoTo ECargarTablaTraduccionIVAS
    If ConnConta Is Nothing Then
        MsgBox "Reinicie ARIGES y vuelva a modificar los parametros para cargar los tipos de IVA", vbExclamation
        Exit Sub
    End If
    
    SQL = "DELETE FROM sivatra"
    conn.Execute SQL
    
    
    Set RS = New ADODB.Recordset
    SQL = "select codigiva from tiposiva"
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = "INSERT INTO sivatra VALUES (" & RS!Codigiva & "," & RS!Codigiva & ")"
        
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    'Ahora acutalizo la relacion
    If mvarTipoIVAre1 > 0 Then
        SQL = "UPDATE sivatra set codigiva2=" & mvarTipoIVAre1 & " WHERE codigiva = " & mvarIVA1
        conn.Execute SQL
    End If
    
    If mvarTipoIVAre2 > 0 Then
        SQL = "UPDATE sivatra set codigiva2=" & mvarTipoIVAre2 & " WHERE codigiva = " & mvarIVA2
        conn.Execute SQL
    End If
        
    If mvarTipoIVAre3 > 0 Then
        SQL = "UPDATE sivatra set codigiva2=" & mvarTipoIVAre3 & " WHERE codigiva = " & mvarIVA3
        conn.Execute SQL
    End If
        
    Exit Sub
ECargarTablaTraduccionIVAS:
    MuestraError Err.Number, Err.Description, "Cargar Tabla Traduccion IVAS"
End Sub



Public Sub ComprobarProgramaEnvioMail()
On Error GoTo EC

    If Me.ExeEnvioMail <> "" Then
        If Dir(App.Path & "\" & Me.ExeEnvioMail, vbArchive) = "" Then
            MsgBox "No existe en este PC el archivo: " & Me.ExeEnvioMail, vbInformation
            ExeEnvioMail = ""
        End If
    End If
    
    If Dir(App.Path & "\SituCod.dat", vbArchive) <> "" Then mvarSituaEnCodigoArticulo = True
    
    Exit Sub
EC:
    ExeEnvioMail = ""
    mvarSituaEnCodigoArticulo = False
End Sub


' Las facturas VAN por otro proceso
'Posiciones
' CINCO primeras
'      ofer  pedido  ped_ruta  alb  alb_ruta
' SIG
'       FraRect
Private Function DevuelveNumeroCopias(posicion As Byte) As Integer
Dim C As String

    

    If posicion > Len(mvarNumeroCopias) Then
        C = "1"
    Else
        C = Mid(mvarNumeroCopias, posicion, 1)
        If C = "" Then C = "1"
        If Not IsNumeric(C) Then C = "1"
    End If
            
    DevuelveNumeroCopias = CInt(Val(C))
    
End Function






'Si existe la tabla de parametros de almagrupo Y NO es EOF
Private Sub FijaParametroAlmagrupo(ByRef R As ADODB.Recordset)
    On Error GoTo eMvarComunicaAlmagrupo
    Set R = New ADODB.Recordset
    mvarComunicaAlmagrupo = False
    SQL = "Select * from salmagrupoparam "
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    SQL = R.Fields(0)
    If SQL <> "" Then mvarComunicaAlmagrupo = True
    R.Close
eMvarComunicaAlmagrupo:
    Err.Clear
    
End Sub


'Si existe la tabla de parametros de whose
Private Function ComprobarEmpresaWHOSE(ByRef R As ADODB.Recordset) As Boolean
    On Error GoTo eComprobarEmpresaWHOSE
    Set R = New ADODB.Recordset
    ComprobarEmpresaWHOSE = False
    SQL = "Select * from sparamwhose "
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    SQL = R.Fields(0)
    If SQL <> "" Then ComprobarEmpresaWHOSE = True
    R.Close
    Exit Function
eComprobarEmpresaWHOSE:
    Err.Clear
    conn.Errors.Clear
    
End Function


Private Function ComprobarGessocial(ByRef R As ADODB.Recordset) As Boolean
    On Error GoTo eComprobarEmpresaWHOSE
    Set R = New ADODB.Recordset
    ComprobarGessocial = False
    SQL = "Select * from asociados "
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    SQL = R.Fields(0)
    If SQL <> "" Then ComprobarGessocial = True
    R.Close
    Exit Function
eComprobarEmpresaWHOSE:
    Err.Clear
    conn.Errors.Clear
   
End Function


Private Sub FijarSiTieneUnSoloBancoPropio(ByRef R As ADODB.Recordset)
    On Error GoTo eFijaSiTieneUnSoloBancoPropio
    Set R = New ADODB.Recordset
    mvarCodigoUnicoBancoPropio_ = 0
    SQL = "Select codbanpr from sbanpr "
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not R.EOF
        mvarCodigoUnicoBancoPropio_ = R!codbanpr
        SQL = SQL & "X"
            R.MoveNext
    Wend
    R.Close
    If Len(SQL) > 1 Then mvarCodigoUnicoBancoPropio_ = 0
    
     Exit Sub
eFijaSiTieneUnSoloBancoPropio:
    Err.Clear
    
End Sub



'MARJAL CHIPOS
Private Sub FijarSiTieneHuertos(ByRef R As ADODB.Recordset)
    'Para marja-chipos
    'Marjal CHIPOS
    
    On Error GoTo eFijarSiTieneHuertos
    mvarHuertos = False
    
    If mvarNumeroInstalacion > 0 Then Exit Sub
    
    SQL = "Select * from sclienhuertos where codclien=-1"
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Si llega aqui siginifca que la tabla existe
    mvarHuertos = True
    R.Close
eFijarSiTieneHuertos:
    Err.Clear
    conn.Errors.Clear
    Set R = Nothing
End Sub




'Fijar lotes generalitat
Private Sub FijarSiTieneLotesGeneralitat(ByRef R As ADODB.Recordset)
    
    On Error GoTo eFijarSiTieneHuertos
    Set R = Nothing
    Set R = New ADODB.Recordset
    mvarLotesGeneralitat = False
    
    If Not mvarManipuladorFitosanitarios Then Exit Sub
    
    SQL = "Select * from slotesgeneralitat where codartic='@#@@'"
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Si llega aqui siginifca que la tabla existe
    
    mvarLotesGeneralitat = True
    R.Close
eFijarSiTieneHuertos:
    Err.Clear
    conn.Errors.Clear
    Set R = Nothing
    
End Sub


'Fijar lotes generalitat
Private Sub FijarSiTieneImporCoarval(ByRef R As ADODB.Recordset)
    
    On Error GoTo eFijarSiTieneHuertos
    Set R = Nothing
    Set R = New ADODB.Recordset
    mvarImportacionesCoarval = False
    
    
    
    SQL = "Select * from sparamcoarval "
    R.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Si llega aqui siginifca que la tabla existe
    If Not IsNull(R.Fields(0)) Then mvarImportacionesCoarval = True

    R.Close
eFijarSiTieneHuertos:
    Err.Clear
    conn.Errors.Clear
    Set R = Nothing
    
End Sub








Public Sub SII_FijarValores()
Dim R
    On Error GoTo eFijarTieneSII
    mvarSIITiene = False
    If Not mvarContabilidadNueva Then Exit Sub
    
    If ConnConta Is Nothing Then
        MsgBox "Error conexion CONTA", vbExclamation
        Exit Sub
    End If
    Set R = New ADODB.Recordset
        
    SQL = "select SIIFechaInicio,SIIDiasAviso FROM parametros"
    R.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic
    If Not R.EOF Then
        If Not IsNull(R!SIIFechaInicio) Then
            'Tiene fecha inicio
            mvarSIITiene = True
            mvarSIIDias = DBLet(R!SIIDiasAviso, "N")
            mvarSIIFechaIni = R!SIIFechaInicio
        End If
    End If
    R.Close
    
eFijarTieneSII:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set R = Nothing
End Sub



Public Sub ModuloTelefoniaVtaPlazos()
Dim R
    On Error GoTo eFijarTieneSII
    
    Set R = New ADODB.Recordset
        
    SQL = "select VtaPlazosTelefono FROM spara2"
    R.Open SQL, conn, adOpenForwardOnly, adLockOptimistic
    If Not R.EOF Then
        If DBLet(R!VtaPlazosTelefono, "N") = 1 Then mvarTelefoniaVtaPlazos = True
    End If
    R.Close
    
eFijarTieneSII:
    Err.Clear
    conn.Errors.Clear
    Set R = Nothing
End Sub





Private Function CampoNuevoBD(ByRef R As ADODB.Recordset, campo As String, Tipo As String, ByRef ValorDevuelto) As Boolean
    On Error Resume Next
        
    ValorDevuelto = R(campo)
    If Err.Number <> 0 Then
        Err.Clear
        CampoNuevoBD = False
        
    Else
        'OK
        If Tipo = "B" Then
            ValorDevuelto = CBool(ValorDevuelto = 1)
        Else
            If IsNull(R(campo)) Then
                If Tipo = "T" Then
                    ValorDevuelto = ""
                Else
                    If Tipo = "F" Then
                        ValorDevuelto = "'1900-01-01'"
                    Else
                        ValorDevuelto = 0
                    End If
                End If
            End If
        End If
        CampoNuevoBD = True
    End If
End Function


Public Sub QuitarFormulariosGrandes()
     mvarHaciendoFrmulariosGrandes = False
End Sub


