VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamTPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros SPATPVG
Private mvarCodClien As String
Private mvarNomClien As String
Private mvarCtaBanc1 As String 'cta prevista de cobro

Private mvarNIFClien As String 'SOLO lectura. Para no tener que leerlo cada vez

Private mvarCodForPa As String 'forma de pago por defecto
Private mvarNomForPa As String
Private mvarTipForPa As Byte 'Tipo de pago

Private mvarTipCierre As Byte 'tipo cierre: por terminal o conjunto.
'Private mvarModoCalculo As Byte 'calcular importes de la linea con iva o no
                                '1=con IVA 0=sin IVA

'Variables asociadas a cada campo de la tabla de parametros SPATPVT
'Del terminal conectado
Private mvarContador As Long
Private mvarNomImpre As String 'impresora impresion tickets
Private mvarHayVisor As Boolean
Private mvarNumPuerto As Byte
Private mvarVelocPue As String

'Introduccion de ticket rapida
Private mvarRapida As Boolean
Private mvarFormaPagoUnica As Boolean   'Este no lo lee de parametros, si no cuando abre el terminal lo comrpueba
Private mvarImprimiDirecto As Boolean

'Desglosa IVAS al pie del TICKET
Private mvarDesglosaIVATicket As Boolean
Private mvarAbreCajon As Byte 'Boolean   0 . NO     1.- COM    2.-LPT

'Solo tendra get
Private mvarSecuenciaCajon As String


'Enero 10
'Redondeo a dos decimales
Private mvarRedondea2 As Boolean
Private mvarCalculaIVAsobrePVP As Boolean


'Control stock al finalizar la venta
Private mvarCtrstockVenta As Boolean


'Marzo 2011
Private mvarComImpresora As Byte

'Sept 2012
Private mvarAvisoGeneraFactura As Boolean
    'Avisar si una factura es al cliente de ventas contado o si un tiket es a un cliente



'Nov 2012   SOLO GET
Private mvarOfertaImporteEntregado As Boolean  'Al entrar en TPV total pondra como cantidad entregada la misma que le total
Private mvarTKocultalineaEntregado As Boolean  'Si no quiere mostrar en el tiket la linea de entregado
Private mvarTkCantidad2Decimales As Boolean  'Muestra 2 decimales en catidad --> 2 decimales tb en PVP
Private mvarTkMostrarTrabajador As Boolean  'Muestra 2 decimales en catidad --> 2 decimales tb en PVP
'Febrero 2013 SOLO GET
Private mvarProhibirTiketsNegativos As Boolean
'Mayo 2013
Private mvarProhibirFitosantiarios_a_Varios As Boolean


'Abril 2016
Private mvarProhibirTicketSocios As Boolean

'Julio 2017
Private mvarImpresoraAlbaranes As String

'Agosto 2017
Private mvarProhibirFacturasTPV As String


'En Impresion tiket directa abre un cursor cada tiket
'Para leer el pie
'EMPIEZA en 0
Private mvarCabecera(4) As String
Private mvarPie(2) As String





'Mayo 2014
' Recargo financiero en alzira
' Cargaremos la formas de pago que llevan recargo financiero
' No tiene GET ya que hay dos funciones que trabajan con esta cadena
Private mvarForpasConRecargoFinanciero As String


'Junio 2015
'Este variable se utiliza en el AriTPV
'Simplemente es para no terner que ir leyendo el numero de terminal continuamente
'Solo get
Private mvarNumeroDeTerminal As Integer


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes


'------------------------------------------------
'Propiedades del modulo CParamTPV
'------------------------------------------------

Public Property Let Cliente(ByVal vData As String)
     mvarCodClien = vData
End Property

Public Property Get Cliente() As String
     Cliente = mvarCodClien
End Property


Public Property Let NomCliente(ByVal vData As String)
     mvarNomClien = vData
End Property

Public Property Get NomCliente() As String
     NomCliente = mvarNomClien
End Property

'Solo tiene get
Public Property Get nifClien() As String
     nifClien = mvarNIFClien
End Property


'Empieza en cero
Public Property Get CabeceraTiket(Index As Integer) As String
    CabeceraTiket = mvarCabecera(Index)
End Property

Public Property Get PieTiket(Index As Integer) As String
    PieTiket = mvarPie(Index)
End Property


'cuenta prevista de cobro por si pasamos el ticket a una factura de venta FAV
Public Property Let CtaPrevCobro(ByVal vData As String)
     mvarCtaBanc1 = vData
End Property

Public Property Get CtaPrevCobro() As String
     CtaPrevCobro = mvarCtaBanc1
End Property


Public Property Let ForPago(ByVal vData As String)
     mvarCodForPa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mvarCodForPa
End Property



Public Property Let NomForPago(ByVal vData As String)
     mvarNomForPa = vData
End Property

Public Property Get NomForPago() As String
     NomForPago = mvarNomForPa
End Property


Public Property Let TipoForPago(ByVal vData As Byte)
     mvarTipForPa = vData
End Property

Public Property Get TipoForPago() As Byte
     TipoForPago = mvarTipForPa
End Property


Public Property Let TipoCierre(ByVal vData As Byte)
     mvarTipCierre = vData
End Property

Public Property Get TipoCierre() As Byte
     TipoCierre = mvarTipCierre
End Property


'Public Property Let ModoCalculo(ByVal vData As Byte)
'     mvarModoCalculo = vData
'End Property
'
'Public Property Get ModoCalculo() As Byte
'     ModoCalculo = mvarModoCalculo
'End Property



Public Property Let Contador(ByVal vData As Long)
    mvarContador = vData
End Property


Public Property Let NomImpresora(ByVal vData As String)
     mvarNomImpre = vData
End Property

Public Property Get NomImpresora() As String
     NomImpresora = mvarNomImpre
End Property


Public Property Get Contador() As Long
    Contador = mvarContador
End Property


Public Property Let HayVisor(ByVal vData As Boolean)
    mvarHayVisor = vData
End Property


Public Property Get HayVisor() As Boolean
    HayVisor = mvarHayVisor
End Property



Public Property Let NumPuerto(ByVal vData As Byte)
    mvarNumPuerto = vData
End Property


Public Property Get NumPuerto() As Byte
    NumPuerto = mvarNumPuerto
End Property



Public Property Let VelociPuerto(ByVal vData As String)
    mvarVelocPue = vData
End Property


Public Property Get VelociPuerto() As String
    VelociPuerto = mvarVelocPue
End Property



Public Property Let Rapida(ByVal vData As Boolean)
     mvarRapida = vData
End Property

Public Property Get Rapida() As Boolean
     Rapida = mvarRapida
End Property



Public Property Let FormaPagoUnica(ByVal vData As Boolean)
     mvarFormaPagoUnica = vData
End Property

Public Property Get FormaPagoUnica() As Boolean
     FormaPagoUnica = mvarFormaPagoUnica
End Property



Public Property Let ImprimiDirecto(ByVal vData As Boolean)
     mvarImprimiDirecto = vData
End Property

Public Property Get ImprimiDirecto() As Boolean
     ImprimiDirecto = mvarImprimiDirecto
End Property







Public Property Let DesglosaIVATicket(ByVal vData As Boolean)
     mvarDesglosaIVATicket = vData
End Property

Public Property Get DesglosaIVATicket() As Boolean
     DesglosaIVATicket = mvarDesglosaIVATicket
End Property


'0. NO abre cajon      1.- Por la COM (lo que habia )    2.- LPT  (nromalmente mapeando
Public Property Let AbreCajon(ByVal vData As Byte)
     mvarAbreCajon = vData
End Property
Public Property Get AbreCajon() As Byte
     AbreCajon = mvarAbreCajon
End Property




Public Property Get SecuenciaCajon() As String
    SecuenciaCajon = mvarSecuenciaCajon
End Property


'mvarRedondea2
Public Property Let Redondea2(ByVal vData As Boolean)
     mvarRedondea2 = vData
End Property

Public Property Get Redondea2() As Boolean
     Redondea2 = mvarRedondea2
End Property



Public Property Let CalculaIVAsobrePVP(ByVal vData As Boolean)
     mvarCalculaIVAsobrePVP = vData
End Property

Public Property Get CalculaIVAsobrePVP() As Boolean
     CalculaIVAsobrePVP = mvarCalculaIVAsobrePVP
End Property



Public Property Let CtrstockVenta(ByVal vData As Boolean)
     mvarCtrstockVenta = vData
End Property

Public Property Get CtrstockVenta() As Boolean
     CtrstockVenta = mvarCtrstockVenta
End Property



'En que COM esta la impresora
Public Property Let ComImpresora(ByVal vData As Byte)
     mvarComImpresora = vData
End Property

Public Property Get ComImpresora() As Byte
     ComImpresora = mvarComImpresora
End Property


'Avisar si una factura es al cliente de ventas contado o si un tiket es a un cliente
Public Property Let AvisoGeneraFactura(ByVal vData As Boolean)
     mvarAvisoGeneraFactura = vData
End Property

Public Property Get AvisoGeneraFactura() As Boolean
     AvisoGeneraFactura = mvarAvisoGeneraFactura
End Property


Public Property Let ProhibirTicketSocios(ByVal vData As Boolean)
     mvarProhibirTicketSocios = vData
End Property

Public Property Get ProhibirTicketSocios() As Boolean
     ProhibirTicketSocios = mvarProhibirTicketSocios
End Property



'SOLO GET
Public Property Get OfertaImporteEntregado() As Boolean
     OfertaImporteEntregado = mvarOfertaImporteEntregado
End Property

Public Property Get TKocultalineaEntregado() As Boolean
     TKocultalineaEntregado = mvarTKocultalineaEntregado
End Property

Public Property Get TkCantidad2Decimales() As Boolean
     TkCantidad2Decimales = mvarTkCantidad2Decimales
End Property

Public Property Get TkMostrarTrabajador() As Boolean
     TkMostrarTrabajador = mvarTkMostrarTrabajador
End Property

Public Property Get ProhibirTiketsNegativos() As Boolean
     ProhibirTiketsNegativos = mvarProhibirTiketsNegativos
End Property

Public Property Get ProhibirFitosantiarios_a_Varios() As Boolean
     ProhibirFitosantiarios_a_Varios = mvarProhibirFitosantiarios_a_Varios
End Property

Public Property Get ForpasConRecargoFinanciero() As String
     ForpasConRecargoFinanciero = mvarForpasConRecargoFinanciero
End Property


Public Property Get NumeroDeTerminal() As Integer
     NumeroDeTerminal = mvarNumeroDeTerminal
End Property


Public Property Get ImpresoraAlbaranes() As String
     ImpresoraAlbaranes = mvarImpresoraAlbaranes
End Property

Public Property Get ProhibirFacturasTPV() As Boolean
     ProhibirFacturasTPV = mvarProhibirFacturasTPV
End Property




'------------------------------------------------
'FUNCIONES del modulo CParamTPV
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
'Lee los parametros generales
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim J As Integer
On Error GoTo Err3
    
    SQL = "SELECT * FROM spatpvg WHERE codigo=1"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If Rs.EOF Then
        Leer = 1
    Else
        'cliente por defecto. Tambien leeremos forma de pago y CIF del cliene x defecto dde ventas
        mvarCodClien = Rs!codClien
        SQL = "codforpa"
        mvarNomClien = DevuelveDesdeBDNew(conAri, "sclien", "concat(nomclien,'|',nifclien,'|')", "codclien", Rs!codClien, "N", SQL)
        mvarCtaBanc1 = DBLet(Rs!ctabanc1, "T")
        mvarCodForPa = SQL 'forma de pago del cliente
        
        'Separamos el campo mvarNomClien en nombre y cif
        mvarNIFClien = Trim(RecuperaValor(mvarNomClien, 2))
        mvarNomClien = RecuperaValor(mvarNomClien, 1)
        
        
        Cad = "tipforpa"
        SQL = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", mvarCodForPa, "N", Cad)
        mvarNomForPa = SQL
        mvarTipForPa = CByte(Cad)
        
        mvarTipCierre = CByte(Rs!tipcierre)
        
        mvarRapida = Val(DBLet(Rs!rapido, "N")) = 1
        mvarFormaPagoUnica = False  'Por defecto lo dejo asi
        mvarImprimiDirecto = Val(DBLet(Rs!imprtick, "N")) = 1
        
        
        
        mvarDesglosaIVATicket = Val(DBLet(Rs!BasesImp, "N")) = 1

        
        mvarRedondea2 = (DBLet(Rs!Redondea2, "N") = 1)
        
        
        mvarCalculaIVAsobrePVP = (DBLet(Rs!IVAsobrePVP, "N") = 1)
        
        mvarCtrstockVenta = (DBLet(Rs!CtrstockVenta, "N")) = 1
        
        
        
        mvarAvisoGeneraFactura = (DBLet(Rs!AvisoGenFra, "N") = 1)
        
        
        
        
        mvarOfertaImporteEntregado = (DBLet(Rs!PonEntragado, "N") = 1)
        mvarTKocultalineaEntregado = (DBLet(Rs!TkOcultarEntregado, "N") = 1)
        mvarTkCantidad2Decimales = (DBLet(Rs!TkCantidad2Decimales, "N") = 1)
        mvarTkMostrarTrabajador = (DBLet(Rs!TkMostrarTrabajador, "N") = 1)
        mvarProhibirTiketsNegativos = (DBLet(Rs!ProhibirTiketsNegativos, "N") = 1)
        mvarProhibirFitosantiarios_a_Varios = (DBLet(Rs!prohibirFitosAVarios, "N") = 1)
        
        'Cabcera y pie
        For J = 1 To 5
            Cad = DBLet(Rs.Fields("cabtick" & CStr(J)), "T")
            mvarCabecera(J - 1) = Cad
        Next
        For J = 1 To 3
            Cad = DBLet(Rs.Fields("pietick" & CStr(J)), "T")
            mvarPie(J - 1) = Cad
        Next
        mvarForpasConRecargoFinanciero = ""
        
        mvarProhibirTicketSocios = (DBLet(Rs!ProhibirTicketSocios, "N") = 1)
        mvarProhibirFacturasTPV = (DBLet(Rs!ProhibirFacturasTPV, "N") = 1)
        Leer = 0
    End If

    Rs.Close
    
    
    If Leer = 0 Then
        'ALZIRA
        If vParamAplic.NumeroInstalacion = 1 Then FijarFormasDePagoRecargoFinanciero Rs
    End If
    
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Set Rs = Nothing
    Leer = 1
End Function




Public Function Leer2(nt As String) As Byte
'Lee los parametros del terminal
Dim Cad As String
Dim Rs As ADODB.Recordset
    
    On Error GoTo Err3
    
    SQL = "SELECT * FROM spatpvt WHERE numtermi=" & nt
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If Rs.EOF Then
        Leer2 = 1
    Else
        mvarNumeroDeTerminal = nt
        mvarNomImpre = DBLet(Rs!NomImpre, "T")
        mvarHayVisor = CBool(Rs!HayVisor)
        mvarNumPuerto = DBLet(Rs!NumPuerto, "N")
        mvarVelocPue = DBLet(Rs!velocpue, "T")
        
        
        mvarAbreCajon = DBLet(Rs!AbreCajon, "N")
        mvarSecuenciaCajon = DBLet(Rs!SecuenciaCajon)
        mvarComImpresora = DBLet(Rs!cualcom, "N")
        mvarImpresoraAlbaranes = DBLet(Rs!ImpresoraALV, "T")
        Leer2 = 0
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Set Rs = Nothing
    Leer2 = 1
End Function



'Public Function Modificar(Codigo As Byte) As Byte
''Modifica los parametros de la Aplicación
''Modifica la Tabla: spara1, BD: Ariges
'On Error GoTo EModificar
'
''    SQL = "UPDATE spatpvg SET "
''    SQL = SQL & " tipodtos= " & mvarTipoDtos
''
''    SQL = SQL & " WHERE codigo =" & Codigo & ";"
''
''    Conn.Execute SQL
''
''    Modificar = 1
'
'    Exit Function
'
'EModificar:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Modificar = 0
'End Function



Public Function ConseguirContador(vCodigo As Byte) As Long
Dim OK As Boolean
Dim C1 As Long 'contador
Dim Rs As ADODB.Recordset

    On Error GoTo Err1
    
    ConseguirContador = 1
    
    'Abrimos bloqueando
    SQL = "Select * from spatpvt WHERE numtermi=" & vCodigo & " FOR UPDATE"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        mvarContador = Rs!Contador
        C1 = Rs!Contador
        OK = True
    End If
    Rs.Close
    Set Rs = Nothing
    
    If OK Then ConseguirContador = C1 + 1
    Exit Function
    
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
End Function


Public Function IncrementarContador(vCodigo As Byte) As Boolean
Dim C1 As Long 'contador1: recibo de entrada

    On Error GoTo Err1
    
      
    'Actualizamos el contador
    C1 = Contador
    SQL = "UPDATE spatpvt set "
    C1 = C1 + 1
    mvarContador = C1
    SQL = SQL & " contador=" & C1
    
    SQL = SQL & " WHERE numtermi=" & vCodigo
    conn.Execute SQL
    IncrementarContador = True
    Exit Function
    
Err1:
    If Err.Number <> 0 Then
        IncrementarContador = False
    Else
        IncrementarContador = True
    End If
End Function



Public Function DevolverContador(vCodigo As Byte, Contador As Long) As Byte
Dim OK As Boolean
Dim C1 As Long
Dim Rs As ADODB.Recordset

    On Error GoTo Err1
    
    'Abrimos bloqueando
    SQL = "Select * from spatpvt WHERE numtermi= " & vCodigo & " FOR UPDATE " ' "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
'    Conn.Execute "Set autocommit = 0"
    DevolverContador = 1
    OK = False
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        C1 = Rs!Contador
        OK = True
    End If
    Rs.Close
    Set Rs = Nothing
    If Not OK Then
        conn.Execute "commit"
        Exit Function
    End If

    OK = False

    OK = (C1 = Contador)
    If OK Then
        'Actualizamos el contador
        SQL = "UPDATE spatpvt set "

        C1 = C1 - 1
        SQL = SQL & " contador=" & C1
        SQL = SQL & " WHERE numtermi = " & vCodigo
        conn.Execute SQL
    End If

    'Desbloqueamos
    conn.Execute "commit"
    DevolverContador = 1
    
Err1:
    If Err.Number <> 0 Then
        Mens = "Error: " & Err.Number & " : " & Err.Description
        MsgBox Mens, vbExclamation
        DevolverContador = 0
    End If
    conn.Execute "Set autocommit = 1"
End Function



'*******************************************************************************
'*******************************************************************************
'
'   Formas de pago con recargo financiero
'
Private Sub FijarFormasDePagoRecargoFinanciero(ByRef R As ADODB.Recordset)
    On Error GoTo eFijarFormasDePagoRecargoFinanciero
    
    SQL = "select codforpa from sforpa where porgasfi>0"
    R.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    While Not R.EOF
        SQL = SQL & "," & R!codforpa
        R.MoveNext
    Wend
    R.Close
    
    If SQL <> "" Then mvarForpasConRecargoFinanciero = SQL
    
    Exit Sub
eFijarFormasDePagoRecargoFinanciero:
    MuestraError Err.Number, Err.Description
End Sub

Public Function FormaDePagoConRegargoFinanciero(KFormaDePago As Integer) As Boolean
    
    FormaDePagoConRegargoFinanciero = False
    If mvarForpasConRecargoFinanciero = "" Then Exit Function

    SQL = mvarForpasConRecargoFinanciero & ","
    Mens = "," & KFormaDePago & ","
    If InStr(1, SQL, Mens) > 0 Then FormaDePagoConRegargoFinanciero = True

End Function

Public Function FormaDePagoConRecFinan_SQL() As String
    
    FormaDePagoConRecFinan_SQL = "(" & Mid(mvarForpasConRecargoFinanciero, 2) & ")"
    
End Function





