VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPpalOld 
   BackColor       =   &H00858585&
   Caption         =   "Ariges 4"
   ClientHeight    =   9525
   ClientLeft      =   165
   ClientTop       =   -990
   ClientWidth     =   17250
   Icon            =   "frmPpalOld.frx":0000
   LinkTopic       =   "MDIForm1"
   ScaleHeight     =   9525
   ScaleWidth      =   17250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17250
      _ExtentX        =   30427
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   30
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Art�culos"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Movimientos Art."
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Proveedores"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ofertas Clientes"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Clientes"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaranes Clientes"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Cliente"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas mostrador"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Proveedor"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaran Proveedor"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Factura Proveedor"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recepci�n Facturas Prov."
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantenimientos"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "N� Serie"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Avisos"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Gastos t�cnicos"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consulta precios / cliente"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Venta TPV"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar empresa"
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agenda"
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   8940
      Width           =   17250
      _ExtentX        =   30427
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmPpalOld.frx":6852
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22357
            Text            =   "asdasd"
            TextSave        =   "asdasd"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAY�S"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "N�M"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "10:25"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnConfiguracion 
      Caption         =   "C&onfiguraci�n"
      Begin VB.Menu mnConfParamGenerales 
         Caption         =   "Datos &Empresa"
         HelpContextID   =   2
      End
      Begin VB.Menu mnConfParamAplic 
         Caption         =   "Par�metros &Aplicaci�n"
      End
      Begin VB.Menu mnConTMovimiento 
         Caption         =   "Tipos &Movimiento"
      End
      Begin VB.Menu mnConfParamRpt 
         Caption         =   "Tipos de &Documentos"
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnConfManteUsuarios 
         Caption         =   "Mantenimiento &Usuarios"
         HelpContextID   =   2
      End
      Begin VB.Menu mnUsuarios 
         Caption         =   "Nuevo U&suario"
         Visible         =   0   'False
      End
      Begin VB.Menu mnCambioEmpresa 
         Caption         =   "Cambiar Em&presa"
         HelpContextID   =   2
      End
      Begin VB.Menu mnBarra17 
         Caption         =   "-"
      End
      Begin VB.Menu mnSeleccionarImpresora 
         Caption         =   "Seleccionar &Impresora"
      End
      Begin VB.Menu mnBarra12 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnAlmacen 
      Caption         =   "&Almacen"
      Begin VB.Menu mnDatosGenAlmacen 
         Caption         =   "&Datos Generales"
         Begin VB.Menu mnAlmMarcas 
            Caption         =   "&Marcas"
         End
         Begin VB.Menu mnAlmAlPropios 
            Caption         =   "Almacenes &Propios"
         End
         Begin VB.Menu mnAlmTipoUnidad 
            Caption         =   "Tipos &Unidad"
         End
         Begin VB.Menu mnTiposArticulos 
            Caption         =   "&Tipos Articulos"
         End
         Begin VB.Menu mnAlmUbicacion 
            Caption         =   "U&bicaciones"
         End
         Begin VB.Menu mnAlmFamiliaArticulo 
            Caption         =   "&Familias Art�culos"
         End
         Begin VB.Menu mnAlmCategoria 
            Caption         =   "&Categor�as"
         End
         Begin VB.Menu mnAlmArticulos 
            Caption         =   "&Art�culos"
         End
         Begin VB.Menu mnAlmNumLotes 
            Caption         =   "&Numeros de lote"
         End
      End
      Begin VB.Menu mnAlmMovimientosAlm 
         Caption         =   "&Movimientos Almacen"
         Begin VB.Menu mnAlmTraspaso 
            Caption         =   "&Traspaso Almacenes"
         End
         Begin VB.Menu mnAlmTraspasoHco 
            Caption         =   "&Hist�rico Traspaso Almacenes"
         End
         Begin VB.Menu mnAlmMovimientos 
            Caption         =   "&Movimientos Almacen"
         End
         Begin VB.Menu mnAlmMovimientosHco 
            Caption         =   "H&ist�rico Movimientos Almacen"
         End
      End
      Begin VB.Menu mnAlmConsultas 
         Caption         =   "&Consultas"
         Begin VB.Menu mnAlmMovimArticulos 
            Caption         =   "Movimientos A&rticulos"
         End
         Begin VB.Menu mnAlmListMovim 
            Caption         =   "Listado &Movimientos"
         End
         Begin VB.Menu mnAlmMovimArticulosSt 
            Caption         =   "Movimientos stock desde inventario"
         End
         Begin VB.Menu mnAlmControlStockDesdeInv 
            Caption         =   "Listado control stock"
         End
         Begin VB.Menu mnAlmListInactivos 
            Caption         =   "Listado Articulos &Inactivos"
         End
         Begin VB.Menu mnAlmListComponentes 
            Caption         =   "Listado Articulos &Componentes"
         End
         Begin VB.Menu mnAlmListValoracion 
            Caption         =   "Listado Valoraci�n &Stocks"
         End
         Begin VB.Menu mnAlmListMaxMin 
            Caption         =   "Inf. Stocks M�ximos-M�nimos"
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Inf. Stocks a una &Fecha"
            Index           =   0
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Stocks por meses"
            Index           =   1
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Alertas punto pedido"
            Index           =   2
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Informe reposici�n almacen"
            Index           =   3
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Listado stock m�nimo"
            Index           =   4
         End
         Begin VB.Menu mnAlmListadosVarios 
            Caption         =   "Movimientos lotes"
            Index           =   5
         End
      End
      Begin VB.Menu mnAlmInventario 
         Caption         =   "&Inventario"
         Begin VB.Menu mnAlmTomaInven 
            Caption         =   "&Toma de inventario"
         End
         Begin VB.Menu mnAlmEntradaInve 
            Caption         =   "&Entrada existencia real"
         End
         Begin VB.Menu mnAlmListadoInve 
            Caption         =   "&Listado diferencias"
         End
         Begin VB.Menu mnAlmActualizarInve 
            Caption         =   "Actualizar &diferencias"
         End
         Begin VB.Menu mnAlmValoracionInve 
            Caption         =   "&Valoraci�n stocks inventariados"
         End
         Begin VB.Menu mnRectifInve 
            Caption         =   "Rectificar �ltimo inventario"
            Index           =   0
         End
         Begin VB.Menu mnRectifInve 
            Caption         =   "Inventariar art�culo"
            Index           =   1
         End
         Begin VB.Menu mnRecalPrSt 
            Caption         =   "Rec�lculo precio est�ndar"
            Index           =   0
         End
         Begin VB.Menu mnRecalPrSt 
            Caption         =   "Rec�lculo precio medio ponderado"
            Index           =   1
         End
         Begin VB.Menu mnRecalPrSt 
            Caption         =   "Rec�lculo ultimo precio compra"
            Index           =   2
         End
         Begin VB.Menu mnBarra2 
            Caption         =   "-"
         End
         Begin VB.Menu mnAlmHcoInven 
            Caption         =   "&Hist�rico inventario"
         End
      End
      Begin VB.Menu mnTelematel 
         Caption         =   "Telematel"
         Index           =   0
      End
   End
   Begin VB.Menu mnFacturacion 
      Caption         =   "&Facturaci�n"
      Begin VB.Menu mnFacDatosGenerales 
         Caption         =   "Datos &Generales"
         Begin VB.Menu mnFacActividades 
            Caption         =   "Activi&dades"
         End
         Begin VB.Menu mnFacZonas 
            Caption         =   "&Zonas"
         End
         Begin VB.Menu mnFacRutas 
            Caption         =   "&Rutas"
         End
         Begin VB.Menu mnPortes 
            Caption         =   "Portes"
         End
         Begin VB.Menu mnDtoCantidad 
            Caption         =   "Descuento por cantidad"
         End
         Begin VB.Menu mnFacFormasEnvio 
            Caption         =   "Formas de &Envio"
         End
         Begin VB.Menu mnFacFormasPago 
            Caption         =   "Formas de &Pago"
         End
         Begin VB.Menu mnFacBancosPropios 
            Caption         =   "&Bancos Propios"
         End
         Begin VB.Menu mnFacSituaciones 
            Caption         =   "&Situaciones Especiales"
         End
         Begin VB.Menu mnFacAgentesCom 
            Caption         =   "Agentes &Comerciales"
         End
         Begin VB.Menu mnFacClientesV1 
            Caption         =   "Clientes &Varios"
         End
         Begin VB.Menu mnFacClientes 
            Caption         =   "Cl&ientes"
         End
         Begin VB.Menu mnFacClientesPot 
            Caption         =   "Clientes potenciales"
         End
         Begin VB.Menu mnFacCartas 
            Caption         =   "Tipos de C&artas"
         End
         Begin VB.Menu mnFacIncidencias 
            Caption         =   "&Incidencias"
         End
      End
      Begin VB.Menu mnFacInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "Clientes Inacti&vos"
            Index           =   0
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "&Altas Clientes"
            Index           =   2
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "&Etiquetas de clientes"
            Index           =   3
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "Car&tas a clientes"
            Index           =   4
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "&Etiquetas de bultos"
            Index           =   5
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "Listado tel�fonos x cliente"
            Index           =   6
         End
         Begin VB.Menu mnFacInformesVarios 
            Caption         =   "Listado cuotas telefon�a"
            Index           =   7
         End
      End
      Begin VB.Menu mnFacPreciosDtos 
         Caption         =   "&Precios y Descuentos"
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Tarifas Venta"
            Index           =   0
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Lista Precios"
            Index           =   1
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Precios &Especiales"
            Index           =   2
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Promociones"
            Index           =   3
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Descuentos Familia/Marca"
            Index           =   4
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Descuento por actividad"
            Index           =   5
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Bonificaciones Factura"
            Index           =   6
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Actualizar precios"
            Index           =   8
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Copiar precios desde compra"
            Index           =   9
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "&Control margenes tarifas"
            Index           =   11
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Correcci�n errores y actualizaci�n tarifas"
            Index           =   12
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Control error descuentos por cliente"
            Index           =   13
         End
         Begin VB.Menu mnFacTarVen 
            Caption         =   "Tarifas tax�metro"
            Index           =   14
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnFacOfert 
         Caption         =   "&Ofertas"
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Mantenimiento Ofertas"
            Index           =   0
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Grupo de Plantillas"
            Index           =   1
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "Entrada de  &Plantillas"
            Index           =   2
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "Ofertas E&fectuadas"
            Index           =   3
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Hist�rico  Ofertas"
            Index           =   5
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "&Traspaso a Hist�rico"
            Index           =   6
         End
         Begin VB.Menu mnFacOfertas 
            Caption         =   "Informe ofertas en historico"
            Index           =   7
         End
      End
      Begin VB.Menu mnFacPed 
         Caption         =   "&Pedidos"
         Begin VB.Menu mnFacPedidos 
            Caption         =   "&Mantenimiento Pedidos"
            Index           =   0
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "&Hist�rico Pedidos Anulados"
            Index           =   1
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "&Cartas Confirmacion de Pedidos"
            Index           =   3
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Informe &Pedidos por Articulo"
            Index           =   4
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Informe P&edidos por Cliente"
            Index           =   5
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Informe &Disponibilidad Stocks"
            Index           =   6
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Impresion pedidos por zona"
            Index           =   7
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Consulta precios / cliente"
            Index           =   9
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Estad�sticas consultas precios/cliente"
            Index           =   10
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Devoluci�n material"
            Index           =   12
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Pedido presupuestos"
            Index           =   13
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Agrupacion pedidos"
            Index           =   14
         End
         Begin VB.Menu mnFacPedidos 
            Caption         =   "Listado pedidos por dia"
            Index           =   15
         End
      End
      Begin VB.Menu mnFacAlbaran 
         Caption         =   "&Albaranes"
         Begin VB.Menu mnFacEntAlbaran 
            Caption         =   "&Mantenimiento Albaranes"
         End
         Begin VB.Menu mnFacAlbDevolucion 
            Caption         =   "Albaranes de devoluci�n"
         End
         Begin VB.Menu mnAlbaranesB 
            Caption         =   "Albaranes presupuestos *"
         End
         Begin VB.Menu mnFacAlbxArtic 
            Caption         =   "Informe &Albaranes por Articulo"
         End
         Begin VB.Menu mnFacIncumPlazos 
            Caption         =   "Inf. Incumplimiento Plazos &Ent."
         End
         Begin VB.Menu mnFacHcoAlbaranes 
            Caption         =   "&Hist�rico Albaranes Anulados"
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Situaci�n albaranes"
            Index           =   0
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Control de albaranes"
            Index           =   1
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Control albaranes facturados"
            Index           =   2
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Impresion albaranes transporte"
            Index           =   3
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Control direcciones de envio"
            Index           =   4
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Listado albaranes"
            Index           =   5
         End
         Begin VB.Menu mnSituaAlba 
            Caption         =   "Listado albaranes entregados"
            Index           =   6
         End
         Begin VB.Menu mnBarra5 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacPreFacturar 
            Caption         =   "&Previsi�n Facturaci�n"
         End
         Begin VB.Menu mnFacFacturarAlb 
            Caption         =   "&Facturaci�n de Albaranes"
         End
         Begin VB.Menu mnFacturarCliente 
            Caption         =   "Facturar cliente"
         End
         Begin VB.Menu mnFacAlbMostrador 
            Caption         =   "Facturas de Mo&strador"
         End
         Begin VB.Menu mnFacturarPresupuestos 
            Caption         =   "Facturar presupuestos *"
         End
         Begin VB.Menu mnFacAlbRectifica 
            Caption         =   "Facturas &Rectificativas"
         End
         Begin VB.Menu mnFacHcoFacturas 
            Caption         =   "His&t�rico Albaran/Factura"
         End
         Begin VB.Menu mnFacReImpFactu 
            Caption         =   "Re&imprimir Facturas"
         End
         Begin VB.Menu mnEnvioFactuasMail 
            Caption         =   "Enviar facturas por e&mail"
            Index           =   0
         End
         Begin VB.Menu mnEnvioFactuasMail 
            Caption         =   "Facturacion web/electr�nica"
            Index           =   1
         End
         Begin VB.Menu mnServicios 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Albaranes de servicio"
            Index           =   1
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Facturaci�n de servicios"
            Index           =   2
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Albaranes internos"
            Index           =   3
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Facturacion albaranes  internos"
            Index           =   4
         End
         Begin VB.Menu mnServicios 
            Caption         =   "Listado albaranes internos"
            Index           =   5
         End
         Begin VB.Menu mnTicket 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnTicket 
            Caption         =   "Contabilizar facturas tickets agrupados"
            Index           =   1
         End
         Begin VB.Menu mnTicket 
            Caption         =   "Listado tickets facturados"
            Index           =   2
         End
         Begin VB.Menu mnBarra9 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacContFactu 
            Caption         =   "&Contabilizar Facturas"
         End
      End
      Begin VB.Menu mnGasolineraPpal 
         Caption         =   "&Gasolinera"
         Begin VB.Menu mnGasolinera 
            Caption         =   "Albaranes gasolinera"
            Index           =   1
         End
         Begin VB.Menu mnGasolinera 
            Caption         =   "Albaranes tienda"
            Index           =   2
         End
         Begin VB.Menu mnGasolinera 
            Caption         =   "Facturacion albaranes gasolinera"
            Index           =   3
            Begin VB.Menu mnGasolFacturar 
               Caption         =   "Cambiar albaranes/facturas"
               Index           =   0
            End
            Begin VB.Menu mnGasolFacturar 
               Caption         =   "Prevision"
               Index           =   1
            End
            Begin VB.Menu mnGasolFacturar 
               Caption         =   "Combustible"
               Index           =   2
            End
            Begin VB.Menu mnGasolFacturar 
               Caption         =   "Tienda"
               Index           =   3
            End
            Begin VB.Menu mnGasolFacturar 
               Caption         =   "-"
               Index           =   4
            End
            Begin VB.Menu mnGasolFacturar 
               Caption         =   "Ajuste formas de pago"
               Index           =   5
            End
         End
         Begin VB.Menu mnGasolinera 
            Caption         =   "Importar fichero gasolinera"
            Index           =   4
         End
      End
      Begin VB.Menu mnTelefonia 
         Caption         =   "&Telefon�a"
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Albaranes de telefon�a"
            Index           =   0
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Importar fichero"
            Index           =   2
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Datos pendientes facturar"
            Index           =   3
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Consumos"
            Index           =   5
            Begin VB.Menu mnTelefonia3 
               Caption         =   "Conceptos"
               Index           =   0
            End
            Begin VB.Menu mnTelefonia3 
               Caption         =   "Descuentos"
               Index           =   1
            End
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Cuotas"
            Index           =   6
            Begin VB.Menu mnTelefonia4 
               Caption         =   "Conceptos"
               Index           =   0
            End
            Begin VB.Menu mnTelefonia4 
               Caption         =   "Descuentos"
               Index           =   1
            End
            Begin VB.Menu mnTelefonia4 
               Caption         =   "Cuotas propias cooperativa"
               Index           =   2
            End
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Cargos varios"
            Index           =   7
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Modificaci�n masiva cuotas"
            Index           =   8
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Comparativa descuentos"
            Index           =   10
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Facturacion x soporte"
            Index           =   11
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Resumen por soporte"
            Index           =   12
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "-"
            Index           =   13
         End
         Begin VB.Menu mnTelefonia2 
            Caption         =   "Datos importaci�n fichero"
            Index           =   14
         End
      End
      Begin VB.Menu mnAgua 
         Caption         =   "Agua"
         Begin VB.Menu mnAguaLin 
            Caption         =   "Contadores"
            Index           =   0
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Importar fichero"
            Index           =   1
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Facturar"
            Index           =   2
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Resumen facturaci�n"
            Index           =   3
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Listado facturaci�n por periodo"
            Index           =   4
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Listado contadores exportaci�n"
            Index           =   5
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Modificar cuota Varios"
            Index           =   6
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Declaraci�n detallada ejercicio"
            Index           =   7
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Listado tasas pendientes de cobro"
            Index           =   8
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Calibres"
            Index           =   10
         End
         Begin VB.Menu mnAguaLin 
            Caption         =   "Par�metros"
            Index           =   11
         End
      End
      Begin VB.Menu mnTratamientosRaiz 
         Caption         =   "Tratamientos"
         Begin VB.Menu mnTratamientos 
            Caption         =   "Mto materias activas"
            Index           =   0
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Mantenimiento ADR"
            Index           =   1
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Plagas"
            Index           =   2
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Flotas"
            Index           =   3
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Tratamientos"
            Index           =   4
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Partes trabajo"
            Index           =   5
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Listado fitosanitarios/campos"
            Index           =   6
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Vacio y NO visible"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnTratamientos 
            Caption         =   "Ajuste compras tratamientos"
            Index           =   9
         End
      End
      Begin VB.Menu mnObra1 
         Caption         =   "Obras"
         Begin VB.Menu mnobra 
            Caption         =   "Cap�tulos"
            Index           =   0
         End
         Begin VB.Menu mnobra 
            Caption         =   "Actuaciones"
            Index           =   1
         End
         Begin VB.Menu mnobra 
            Caption         =   "Partes de trabajo"
            Index           =   2
         End
         Begin VB.Menu mnobra 
            Caption         =   "Mto tipos �rdenes de trabajo"
            Index           =   3
         End
         Begin VB.Menu mnobra 
            Caption         =   "Reloj"
            Index           =   4
         End
         Begin VB.Menu mnobra 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnobra 
            Caption         =   "Listado compras-ventas actuacion"
            Index           =   6
         End
         Begin VB.Menu mnobra 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnobra 
            Caption         =   "Imprimir certificaci�n"
            Index           =   8
         End
      End
      Begin VB.Menu mnHuertos 
         Caption         =   "Gestion parcelas"
         Begin VB.Menu mnHuertos1 
            Caption         =   "Listado campos-hanegadas"
            Index           =   0
         End
         Begin VB.Menu mnHuertos1 
            Caption         =   "Facturaci�n derrama"
            Index           =   1
         End
      End
      Begin VB.Menu mnBarra6 
         Caption         =   "-"
      End
      Begin VB.Menu mnFacEstadistica 
         Caption         =   "&Estad�stica"
         Begin VB.Menu mnFacEstVentaCliente 
            Caption         =   "&Ventas por cliente"
         End
         Begin VB.Menu mnFacEstVentaTraba 
            Caption         =   "Ventas por &trabajador"
         End
         Begin VB.Menu mnFacEstVentaMes 
            Caption         =   "Ventas por &meses"
         End
         Begin VB.Menu mnFacEstVentaFam 
            Caption         =   "Ventas por &familia  /  Art�culo"
         End
         Begin VB.Menu mnEstVtasArituclo 
            Caption         =   "Ventas por art�culo"
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas por proveedor"
            Index           =   0
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas por &agente"
            Index           =   1
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "&Detalle facturaci�n"
            Index           =   2
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Mar&gen ventas por art�culo "
            Index           =   3
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas por tipo de precio"
            Index           =   4
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Articulos m�s vendidos"
            Index           =   5
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas familia agrupado"
            Index           =   6
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Ventas por tipo de pedido"
            Index           =   7
         End
         Begin VB.Menu mnFacEstadistica2 
            Caption         =   "Listado de costes"
            Index           =   8
         End
      End
   End
   Begin VB.Menu mnCompras 
      Caption         =   "&Compras"
      Begin VB.Menu mnComDatosGenerales 
         Caption         =   "Datos &Generales"
         Begin VB.Menu mnComProveedores 
            Caption         =   "&Proveedores"
         End
         Begin VB.Menu mnComProveVarios 
            Caption         =   "Proveedores &Varios"
         End
         Begin VB.Menu mnComDirecciones 
            Caption         =   "&Direcciones"
         End
      End
      Begin VB.Menu mnComInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnComInfVarios1 
            Caption         =   "&Etiquetas de proveedores"
            Index           =   1
         End
         Begin VB.Menu mnComInfVarios1 
            Caption         =   "&Cartas a Proveedores"
            Index           =   2
         End
         Begin VB.Menu mnComInfVarios1 
            Caption         =   "Etiquetas de bultos"
            Index           =   3
         End
      End
      Begin VB.Menu mnComPreciosDtos 
         Caption         =   "Precios y &Descuentos"
         Begin VB.Menu mnComPreProv 
            Caption         =   "P&recios Proveedor"
            Index           =   0
         End
         Begin VB.Menu mnComPreProv 
            Caption         =   "Descuentos Pro&veedor"
            Index           =   1
         End
         Begin VB.Menu mnComPreProv 
            Caption         =   "Copiar precios desde venta"
            Index           =   2
         End
         Begin VB.Menu mnComPreProv 
            Caption         =   "Actualizar precios"
            Index           =   3
         End
      End
      Begin VB.Menu mnComPedidos 
         Caption         =   "&Pedidos"
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "Mant. &Pedidos Proveedor"
            Index           =   0
         End
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "&Hist�rico Pedidos Anulados"
            Index           =   1
         End
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "List. &Material pendiente de recibir"
            Index           =   2
         End
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "Propuesta pedido"
            Index           =   4
         End
         Begin VB.Menu mnComPedidosLin 
            Caption         =   "Estad�stica reaprovisionamiento"
            Index           =   5
         End
      End
      Begin VB.Menu mnComAlbaranes 
         Caption         =   "&Albaranes"
         Begin VB.Menu mnComAlbMan 
            Caption         =   "&Mant. Albaranes Proveedor"
         End
         Begin VB.Menu mnComHcoAlbaranes 
            Caption         =   "&Hist�rico Albaranes Anulados"
         End
         Begin VB.Menu mnComPteFacturar 
            Caption         =   "List. &Pendiente de facturar"
         End
         Begin VB.Menu mnBarra7 
            Caption         =   "-"
         End
         Begin VB.Menu mnComFacturar 
            Caption         =   "&Recepci�n Facturas"
         End
         Begin VB.Menu mnComHcoFacturas 
            Caption         =   "&Hist�rico Albaran/Factura"
         End
         Begin VB.Menu mnBarra15 
            Caption         =   "-"
         End
         Begin VB.Menu mnComContFactu 
            Caption         =   "&Contabilizar Facturas"
         End
         Begin VB.Menu mnComCtrlAlb 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnComCtrlAlb 
            Caption         =   "Control albaranes"
            Index           =   1
         End
         Begin VB.Menu mnComCtrlAlb 
            Caption         =   "Control albaranes facturados"
            Index           =   2
         End
      End
      Begin VB.Menu mnProcesoLiquidacionProveedores 
         Caption         =   "Liquidaci�n proveedores"
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Cambiar precios"
            Index           =   0
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Liquidacion proveedores"
            Index           =   1
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Impresion facturas"
            Index           =   2
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Asociar albaranes compras / ventas"
            Index           =   3
         End
         Begin VB.Menu mnSociosProveedores 
            Caption         =   "Listado asociaciones albaranes"
            Index           =   4
         End
      End
      Begin VB.Menu Barra7 
         Caption         =   "-"
      End
      Begin VB.Menu mnComEstadistica 
         Caption         =   "&Estad�stica"
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Compras por &Proveedor"
            Index           =   0
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Compras por &Familia/Art�c."
            Index           =   1
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Compras por &meses"
            Index           =   2
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "&Albaranes por Proveedor"
            Index           =   3
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Informe previsi�n pagos"
            Index           =   4
         End
         Begin VB.Menu mnComEstadisticaLin 
            Caption         =   "Compras Proveedor-Marca-Familia"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnAdministracion 
      Caption         =   "A&dministraci�n"
      Begin VB.Menu mnAdmTrabajadores 
         Caption         =   "&Trabajadores"
      End
      Begin VB.Menu mnAdmGastosTec 
         Caption         =   "&Gastos T�cnicos"
      End
      Begin VB.Menu mnAdmNominas 
         Caption         =   "&Nominas y Gastos"
      End
      Begin VB.Menu mnInformesAdm 
         Caption         =   "Informes"
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Beneficio por proveedor"
            Index           =   3
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Beneficio por cliente"
            Index           =   4
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Beneficio marca-agente-proveedor"
            Index           =   5
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Informe de art�culos en promocion"
            Index           =   6
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Informe ventas articulos con dto. especial"
            Index           =   7
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Ventas trabajador / Dia"
            Index           =   8
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Ventas por forma de pago"
            Index           =   9
         End
         Begin VB.Menu mnInfoAdm 
            Caption         =   "Comparativo descuentos compra/venta"
            Index           =   10
         End
      End
      Begin VB.Menu mnAdmAgentes 
         Caption         =   "Agentes"
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Resumen ventas - agente"
            Index           =   0
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Beneficio por agente"
            Index           =   1
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Ventas agente - trabajador"
            Index           =   2
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Pagos comisiones"
            Index           =   4
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Generar pagos comisiones"
            Index           =   5
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Listado comisiones ECO"
            Index           =   7
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Listado agente-familia-marca"
            Index           =   8
         End
         Begin VB.Menu mnAdmAgen2 
            Caption         =   "Listado agente-marca-familia"
            Index           =   9
         End
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Calculo de &riesgo"
         Index           =   1
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Listado riesgo"
         Index           =   2
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Informe de previsi�n tesoreria"
         Index           =   3
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Correcci�n costes varios factura"
         Index           =   4
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Modificar coste estadistica ventas"
         Index           =   5
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Gesti�n de flotas-Maquinaria"
         Index           =   6
         Begin VB.Menu mnFlotas 
            Caption         =   "Registro "
            Index           =   0
         End
         Begin VB.Menu mnFlotas 
            Caption         =   "Mantenimiento de flotas"
            Index           =   1
         End
         Begin VB.Menu mnFlotas 
            Caption         =   "Mantenimiento de conceptos"
            Index           =   2
         End
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Comunicaci�n datos seguro"
         Index           =   7
      End
      Begin VB.Menu mnAdministra 
         Caption         =   "Informe ventas a credito"
         Index           =   8
      End
   End
   Begin VB.Menu mnMantenimientos 
      Caption         =   "&Mantenimientos"
      Begin VB.Menu mnManTiposContrato 
         Caption         =   "&Tipos de Contrato"
      End
      Begin VB.Menu mnManEntrada 
         Caption         =   "&Entrada Mantenimientos"
      End
      Begin VB.Menu mnBarra8 
         Caption         =   "-"
      End
      Begin VB.Menu mnManListado 
         Caption         =   "&Listado Mantenimientos"
      End
      Begin VB.Menu mnManRevisiones 
         Caption         =   "Listado &Revisiones Mant."
      End
      Begin VB.Menu mnManFichas 
         Caption         =   "&Fichas Mantenimientos"
      End
      Begin VB.Menu mnManAltas 
         Caption         =   "List. &Altas Mantenimientos"
      End
      Begin VB.Menu mnInfTeoMant 
         Caption         =   "Informe te�rico mantenimientos"
      End
      Begin VB.Menu mnEtiqMante 
         Caption         =   "Etiquetas de mantenimientos"
      End
      Begin VB.Menu mnBarra30 
         Caption         =   "-"
      End
      Begin VB.Menu mnCartaRenovaMante 
         Caption         =   "Carta renovaci�n"
      End
      Begin VB.Menu mnTraspasoMante 
         Caption         =   "Traspaso siguiente a actual"
      End
      Begin VB.Menu mnBarra32 
         Caption         =   "-"
      End
      Begin VB.Menu mnHcoMaten 
         Caption         =   "Hist�rico mantenimientos anulados"
      End
      Begin VB.Menu mnInfManteAnulados 
         Caption         =   "Informe mantenimientos anulados"
      End
      Begin VB.Menu mnBarra13 
         Caption         =   "-"
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "&Previsi�n Facturaci�n"
         Index           =   0
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "Fac&turaci�n  Mantenimientos"
         Index           =   1
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "Prevision facturacion Renting y servicios"
         Index           =   3
      End
      Begin VB.Menu mnManPrevFac2 
         Caption         =   "Facturacion renting y servicios"
         Index           =   4
      End
   End
   Begin VB.Menu mnReparaciones 
      Caption         =   "&Reparaciones"
      Begin VB.Menu mnRepEntReparacion 
         Caption         =   "&Mant.  Reparaciones"
      End
      Begin VB.Menu mnRepControlRep 
         Caption         =   "C&ontrol Reparaciones"
      End
      Begin VB.Menu mnRepNumSerie 
         Caption         =   "Mant. &N� Serie"
      End
      Begin VB.Menu mnRepMotivosBaja 
         Caption         =   "Motivos &baja equipos"
      End
      Begin VB.Menu mnRepMotivosPend 
         Caption         =   "Motivos &Pend. Rep."
      End
      Begin VB.Menu mnRepHistorico 
         Caption         =   "&Hist�rico de Reparaciones"
      End
      Begin VB.Menu mnManServicioAsisTecn 
         Caption         =   "Servicios asistencia t�cnica"
      End
      Begin VB.Menu mnTiposAveria 
         Caption         =   "Tipos averia"
      End
      Begin VB.Menu mnTrabaRealiz 
         Caption         =   "Trabajos realizados"
      End
      Begin VB.Menu mnMtoEuler 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnMtoEuler 
         Caption         =   "Albar�n orden de trabajo"
         Index           =   1
      End
      Begin VB.Menu mnMtoEuler 
         Caption         =   "Albar�n trabajo exterior"
         Index           =   2
      End
      Begin VB.Menu mnMtoEuler 
         Caption         =   "Albar�n de reparaci�n"
         Index           =   3
      End
      Begin VB.Menu mnMtoEuler 
         Caption         =   "Proyectos"
         Index           =   4
      End
      Begin VB.Menu Barra9 
         Caption         =   "-"
      End
      Begin VB.Menu mnRepListRepxDia 
         Caption         =   "Listado Rep. del &Dia"
      End
      Begin VB.Menu mnRepListRepxClien 
         Caption         =   "Listado Rep. por &Cliente"
      End
      Begin VB.Menu mnRepListFrecuen 
         Caption         =   "F&recuencia de reparaciones"
      End
      Begin VB.Menu mnEstadisticaReparacionTecnico 
         Caption         =   "Estad�stica reparaciones t�cnico"
      End
      Begin VB.Menu mnListadoReparacionesEfectuadas 
         Caption         =   "Listado reparaciones efectuadas"
      End
      Begin VB.Menu mnRepGarantprove 
         Caption         =   "Reparaciones garantia proveedor"
      End
      Begin VB.Menu Barra14 
         Caption         =   "-"
      End
      Begin VB.Menu mnRepAlbaranes 
         Caption         =   "Mant. &Albaranes Rep."
      End
      Begin VB.Menu mnRepPrevFact 
         Caption         =   "Pre&visi�n Facturaci�n"
      End
      Begin VB.Menu mnRepFactAlb 
         Caption         =   "&Facturaci�n Reparaciones"
      End
      Begin VB.Menu mnBarra14 
         Caption         =   "-"
      End
      Begin VB.Menu mnRepAvisos 
         Caption         =   "Av&isos de clientes"
      End
      Begin VB.Menu mnRepListAvisosPtes 
         Caption         =   "&Listado de avisos pendientes"
      End
      Begin VB.Menu mnBorrarAvisosCerrados 
         Caption         =   "Borre avisos cerrados"
      End
      Begin VB.Menu mnbarra33 
         Caption         =   "-"
      End
      Begin VB.Menu mnFrecuencias 
         Caption         =   "Frecuencias"
      End
   End
   Begin VB.Menu mnCRMmenu 
      Caption         =   "CRM"
      Begin VB.Menu mnCRM 
         Caption         =   "Mantenimiento acciones comerciales"
         Index           =   0
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Tipos acciones comerciales"
         Index           =   1
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Generar acciones comerciales"
         Index           =   2
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Impresion masiva"
         Index           =   3
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Impresion resumen CRM"
         Index           =   4
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Informe clientes por acci�n comercial"
         Index           =   5
      End
   End
   Begin VB.Menu mnproduccion 
      Caption         =   "Producci�n"
      Begin VB.Menu mnproduccion1 
         Caption         =   "�rdenes producci�n"
         Index           =   0
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Ordenes de envasado"
         Index           =   1
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Descripci�n costes tasas"
         Index           =   2
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Registro trazabilidad"
         Index           =   3
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Par�metros c�lidad"
         Index           =   4
      End
      Begin VB.Menu mnproduccion1 
         Caption         =   "Declaracion alcohol AEAT"
         Index           =   5
      End
   End
   Begin VB.Menu mnTPV 
      Caption         =   "&Punto de Venta"
      Begin VB.Menu mnTPVLinea 
         Caption         =   "Pantalla de &venta"
         Index           =   0
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "&Cierre de caja"
         Index           =   1
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "Etiquetas estanter�a"
         Index           =   2
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "&Par�metros generales TPV"
         Index           =   4
      End
      Begin VB.Menu mnTPVLinea 
         Caption         =   "Par�metros &terminales TPV"
         Index           =   5
      End
   End
   Begin VB.Menu mnUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnAgenda 
         Caption         =   "&Agenda"
      End
      Begin VB.Menu mnVerAvisos 
         Caption         =   "A&visos"
      End
      Begin VB.Menu mnLlamadas 
         Caption         =   "Llamadas"
         Index           =   0
      End
      Begin VB.Menu mnLlamadas 
         Caption         =   "Concepto llamadas"
         Index           =   1
      End
      Begin VB.Menu mnBackUp 
         Caption         =   "&Copia Seguridad local"
      End
      Begin VB.Menu mnEliminarFacturas 
         Caption         =   "&Borre Facturas y Movimientos"
      End
      Begin VB.Menu mnRevisarMultibase 
         Caption         =   "Revisar caracteres especiales"
      End
      Begin VB.Menu mnManteneLOG 
         Caption         =   "Acciones realizadas"
      End
      Begin VB.Menu mnAridocFacturas 
         Caption         =   "Traspaso Aridoc"
      End
      Begin VB.Menu mnUtiDeclaraLOM 
         Caption         =   "Lotes fitosanitarios subvencionados"
         Index           =   0
      End
      Begin VB.Menu mnUtiDeclaraLOM 
         Caption         =   "Declaraci�n ROPO"
         Index           =   1
      End
      Begin VB.Menu mnArticulos 
         Caption         =   "Acciones art�culos"
         Begin VB.Menu mnArticulos2 
            Caption         =   "Eliminar articulos"
            Index           =   0
         End
         Begin VB.Menu mnArticulos2 
            Caption         =   "Cambiar familia / marca / proveedor"
            Index           =   1
         End
         Begin VB.Menu mnArticulos2 
            Caption         =   "Cambiar codigo articulo-referencia"
            Index           =   2
         End
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Listado albaranes-pedidos anulados"
         Index           =   0
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Comprobar cuenta bancaria/NIF"
         Index           =   1
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Traspaso contados ruta"
         Index           =   2
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Eliminar presupuestos"
         Index           =   3
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Configurar PDFs ver articulo"
         Index           =   4
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Exportar albaranes de servicio"
         Index           =   5
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Exportar email csv"
         Index           =   6
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Listado credito y caucion"
         Index           =   7
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Importacion coarval"
         Index           =   8
      End
      Begin VB.Menu mnUtilidadesVarias 
         Caption         =   "Cambio cliente"
         Index           =   9
      End
      Begin VB.Menu mnBarra19 
         Caption         =   "-"
      End
      Begin VB.Menu mnUtiBuscar 
         Caption         =   "&Buscar..."
         Begin VB.Menu mnUtiBuscarErrFac 
            Caption         =   "&Errores en N� Factura clientes"
         End
         Begin VB.Menu mnUtiBuscarPteCon 
            Caption         =   "Facturas pendientes de &contabilizar"
            Begin VB.Menu mnUtiBuscarErrConCli 
               Caption         =   "&Clientes"
            End
            Begin VB.Menu mnUtiBuscarErrConPro 
               Caption         =   "&Proveedores"
            End
         End
      End
      Begin VB.Menu mnCambioPwd 
         Caption         =   "Cambiar contrase�a"
      End
      Begin VB.Menu mnBarra20 
         Caption         =   "-"
      End
      Begin VB.Menu mnUtiUsuActivos 
         Caption         =   "&Usuarios activos"
      End
      Begin VB.Menu mnUtiConnActivas 
         Caption         =   "&Conexiones activas"
      End
      Begin VB.Menu mnBarra21 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnUtiMensInt 
         Caption         =   "&Mensajeria interna"
         Visible         =   0   'False
         Begin VB.Menu mnUtiMensLin 
            Caption         =   "&Nuevo"
            Index           =   0
         End
         Begin VB.Menu mnUtiMensLin 
            Caption         =   "&Enviar/Recibir"
            Index           =   1
         End
         Begin VB.Menu mnUtiMensLin 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnUtiMensLin 
            Caption         =   "&Tipo de mensaje"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnSoporte2 
      Caption         =   "&Soporte"
      Begin VB.Menu mnSoporte 
         Caption         =   "Ayuda"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "-"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Enviar Mail"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Web Ariadna Software"
         Index           =   4
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Comprobar version operativa"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Acerca de ..."
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmPpalOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Ppal   ////////////////////////////////////
Option Explicit
Dim PrimeraVez As Boolean

Dim TieneEditorDeMenus As Boolean

Dim QueCaption As String



Private Sub SituarArriba()
    On Error Resume Next
    Me.Top = 0
    Me.Left = 0
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Activate()
Dim B As Boolean
 

   ' AvisosPendientes = False
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
       ' AvisosPendientes = TieneAvisosPendientes()
        If InstalacionEsEulerTaxco Then SituarArriba
        
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not vParam Is Nothing Then
        If vParam.Modificado Then
          'Poner datos visible del form
           
           vParam.Modificado = False
        End If
    End If
    
    PonerDatosVisiblesForm
    
    '-- Control de si se utilizan servicios o no ( si es que no no se muestra el men�) hemos fichado gente nueva para la copa
    '   el situarlo aqui hace que no haya que salir y entrar en el programa si se
    B = DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "ALI", "T") <> ""
    PuntoDeMenuVisible mnServicios(3), B
    If B And vParamAplic.NumeroInstalacion = vbTaxco Then mnServicios(3).Caption = "Albar�n de garant�as"
    PuntoDeMenuVisible mnServicios(4), B And vParamAplic.NumeroInstalacion <> 4
    PuntoDeMenuVisible mnServicios(5), B
    PuntoDeMenuVisible mnServicios(1), vParamAplic.Servicios
    PuntoDeMenuVisible mnServicios(2), vParamAplic.Servicios And vParamAplic.NumeroInstalacion <> 4  'EULER NO los factura
    B = B Or vParamAplic.Servicios
    PuntoDeMenuVisible mnServicios(0), B  'la barra
    
    
    
    'Noviembre 2019. Gasolinera
    PuntoDeMenuVisible mnGasolineraPpal, vParamAplic.TieneGasolinera
    
    
    
    
    PuntoDeMenuVisible mnFacPedidos(14), vParamAplic.NumeroInstalacion = vbFenollar
    
    
    
    
    
    'MAntenimientos y reparaciones
    'mnMantenimientos.visible = vParamAplic.Mantenimientos
    'mnReparaciones.visible = vParamAplic.Reparaciones
    PuntoDeMenuVisible mnReparaciones, vParamAplic.Reparaciones
    PuntoDeMenuVisible mnMantenimientos, vParamAplic.Mantenimientos
    
    
    '-- Eliminamos frecuencias de momento
    'mnFrecuencias.visible = vParamAplic.Frecuencias
    PuntoDeMenuVisible mnFrecuencias, vParamAplic.Frecuencias
    Me.mnbarra33.visible = mnFrecuencias.visible
    
    '--------------------
    'TElefonia
    
    PuntoDeMenuVisible mnTelefonia, vParamAplic.TieneTelefonia2 > 0
    If vParamAplic.TieneTelefonia2 > 0 Then
        
    End If

    '-- Contabilizacion tickets agrupados
    mnTicket(0).visible = vParamAplic.ContabilizarTicketAgrupados
    mnTicket(1).visible = vParamAplic.ContabilizarTicketAgrupados
    mnTicket(2).visible = vParamAplic.ContabilizarTicketAgrupados
    
        
    'Los albaranes y facturas en "B"
    'seran visibles si esta creado el tipo movimiento y tene contabilidad B
    B = False
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        B = True
        B = B And vParamAplic.ContabilidadB > 0 And HaMostradoCanal2_El_B
    Else
        B = DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "ALZ", "T") <> ""
        If B And vParamAplic.NumeroInstalacion = vbHerbelca Then
            If Val(vUsu.AlmacenPorDefecto2) < 20 Then
                B = False
            Else
                B = True
            End If
        End If
            
    End If
    
    
    PuntoDeMenuVisible Me.mnAlbaranesB, B
    PuntoDeMenuVisible mnFacturarPresupuestos, B
  
    
       
    'Produccion
    B = vParamAplic.Produccion
    If Not B Then B = vParamAplic.TieneComponentes_y_Produccion
    PuntoDeMenuVisible Me.mnproduccion, B
    PuntoDeMenuVisible Me.mnCRMmenu, vParamAplic.TieneCRM
    PuntoDeMenuVisible mnAlmListadosVarios(5), False   'de moemnto false
    
    'Flotas
    
    PuntoDeMenuVisible Me.mnAdministra(6), vParamAplic.GestionFlotas
       
    'Obras
    PuntoDeMenuVisible mnObra1, vParamAplic.HayDeparNuevo = 2
    
    'Mensajeria
    mnUtiMensInt.visible = False
    mnBarra21.visible = False
    
 '   If AvisosPendientes Then
 '       If MsgBox("Tiene avisos pendientes. �Quiere verlos ahora?", vbQuestion + vbYesNo) = vbYes Then
 '           'Mostrare la pantalla de avisos pendientes
 '           frmAlertas.Show vbModal
 '       End If
 '   End If
    '-- Descriptores especiales (Vrs 4.0.9)
    If vParamAplic.Descriptores Then
        mnAlmTipoUnidad.Caption = "Formatos"
        mnTiposArticulos.Caption = "Modelos"
        mnAlmFamiliaArticulo.Caption = "Categorias Art."
        mnAlmCategoria.visible = False
    End If
    
    'Para este usuario y esta empresa unos avlores al usuario
    vUsu.FijarOtrosValoresUsuario
    
    'mnAlmMovimArticulosSt.visible = False
    'Dim TieneDevolucionRMA As Boolean  'Si tiene el tipo de movimiento PEW''
    B = False
    If DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "PEW", "T") <> "" Then B = True
    PuntoDeMenuVisible Me.mnFacPedidos(11), B
    PuntoDeMenuVisible Me.mnFacPedidos(12), B
    B = False
    'If vParamAplic.AlmacenB >= 0 Then
    If vParamAplic.NumeroInstalacion = 2 Then
        'Tiene almacen B
        If DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "PEZ", "T") <> "" Then
            If vUsu.AlmacenPorDefecto2 = CStr(vParamAplic.AlmacenB) Then B = True
        End If
    End If
    PuntoDeMenuVisible Me.mnFacPedidos(13), B
    'QUe ponga el separador
    If B Then
        If Me.mnFacPedidos(13).visible Then Me.mnFacPedidos(11).visible = True
    End If
    
    PuntoDeMenuVisible mnFacPedidos(14), vParamAplic.NumeroInstalacion = vbFenollar
    
    
    
        'Si es empresa de b
    'Utilidades de traspaso presu a factura
    'y eliminar presu
    B = False
    'If vParamAplic.AlmacenB > 90 Then
    If vParamAplic.NumeroInstalacion = 2 Then
        If vUsu.Codigo Mod 1000 = 0 Then
            B = True
        Else
            B = Val(vUsu.AlmacenPorDefecto2) > 90
        End If
    End If
    PuntoDeMenuVisible mnUtilidadesVarias(2), B
    PuntoDeMenuVisible mnUtilidadesVarias(3), B
    PuntoDeMenuVisible mnproduccion1(5), vParamAplic.NumeroInstalacion = 5
    
    PuntoDeMenuVisible mnUtilidadesVarias(9), vParamAplic.NumeroInstalacion = vbTaxco
    
    
    
    'De momento no hay
    mnComCtrlAlb(2).visible = False
    
    'Facturacion electronica
    PuntoDeMenuVisible mnEnvioFactuasMail(1), vParamAplic.PathFacturaE <> ""
    
    'Tratamientos
    'PuntoDeMenuVisible mnTratamientosRaiz, vParamAplic.Ariagro <> "" 'si tiene conexion ariagro ALZIRA
    PuntoDeMenuVisible mnTratamientosRaiz, vParamAplic.LlevaADV  'si tiene conexion ariagro ALZIRA
    
    
    PuntoDeMenuVisible Me.mnTratamientos(7), False  'Oculto para futuros usos
    
    'Renting y servicios
    PuntoDeMenuVisible Me.mnManPrevFac2(2), vParamAplic.Renting
    PuntoDeMenuVisible Me.mnManPrevFac2(3), vParamAplic.Renting
    PuntoDeMenuVisible Me.mnManPrevFac2(4), vParamAplic.Renting
    If vParamAplic.Renting Then
        mnManPrevFac2(4).Caption = "Facturaci�n " & RentingLB & " y servicios"
        mnManPrevFac2(3).Caption = "Previsi�n " & LCase(mnManPrevFac2(4).Caption)
        
    End If
    
    PuntoDeMenuVisible mnProcesoLiquidacionProveedores, False
    
    'Tfnos x cliente
    PuntoDeMenuVisible Me.mnFacInformesVarios(6), vParamAplic.TieneTelefonia2 > 0
    PuntoDeMenuVisible Me.mnFacInformesVarios(7), vParamAplic.TieneTelefonia2 > 0
    
    PuntoDeMenuVisible Me.mnFacClientesPot, vParamAplic.ClientesPotenciales
    
    
    'Comunicacion datos ALMAGRUPO
    PuntoDeMenuVisible mnTelematel(0), vParamAplic.NumeroInstalacion = vbHerbelca
    
    
    PuntoDeMenuVisible mnAgua, vParamAplic.AguasPotables
    
    PuntoDeMenuVisible Me.mnHuertos, vParamAplic.Huertos
    
      
    'ELUER
    B = False
    If InstalacionEsEulerTaxco Then
        B = True
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            mnMtoEuler(1).Caption = "Orden de taller"
            mnMtoEuler(2).Caption = "Gestoria"

            mnFacTarVen(14).visible = True

        End If
    End If
    
    PuntoDeMenuVisible mnUtilidadesVarias(4), B 'vParamAplic.NumeroInstalacion = vbEuler
    PuntoDeMenuVisible mnMtoEuler(0), B
    PuntoDeMenuVisible mnMtoEuler(1), B
    'ALE Euler trabajo exterior    ---> Taxco  GESTORIA
    PuntoDeMenuVisible mnMtoEuler(2), B And vUsu.Nivel2 < 2
    
    PuntoDeMenuVisible mnMtoEuler(3), vParamAplic.NumeroInstalacion = vbEuler  'b
    PuntoDeMenuVisible mnMtoEuler(4), vParamAplic.NumeroInstalacion = vbEuler  'b
    PuntoDeMenuVisible mnFacEstadistica2(8), vParamAplic.NumeroInstalacion = vbEuler  'b

    If InstalacionEsEulerTaxco Then
        PuntoDeMenuVisible mnRepAlbaranes, False
        PuntoDeMenuVisible mnRepEntReparacion, False
        
        'Quito los de mantenimientos
        '
        PuntoDeMenuVisible mnRepEntReparacion, False
        PuntoDeMenuVisible mnRepControlRep, False
        PuntoDeMenuVisible mnRepNumSerie, vParamAplic.NumeroInstalacion = vbTaxco
        PuntoDeMenuVisible mnRepMotivosBaja, False
        PuntoDeMenuVisible mnRepMotivosPend, False
        PuntoDeMenuVisible mnRepHistorico, False
        PuntoDeMenuVisible mnManServicioAsisTecn, False
        PuntoDeMenuVisible mnTiposAveria, False
        PuntoDeMenuVisible mnTrabaRealiz, False
        PuntoDeMenuVisible mnMtoEuler(0), False
        
        
    End If
    
    PuntoDeMenuVisible mnUtilidadesVarias(8), vParamAplic.ImportacionesCoarval
    
        
    
    'Declaracion  fitosnaitarios
    mnUtiDeclaraLOM(0).visible = vParamAplic.LotesGeneralitat 'SUBVENCIONADOS
    mnUtiDeclaraLOM(1).visible = vParamAplic.ManipuladorFitosanitarios2
    
    If vParamAplic.CartaPortes Then
        mnFacFormasEnvio.Caption = "&Transportistas"
    Else
        mnFacFormasEnvio.Caption = "Formas de &Envio"
    End If
    'Lo pong aqui el 11 de Enero de 2011
    'Comprobar que los iconos de la barra su correspondiente
    'entrada de menu esta habilitada sino desabilitar
    PoneBarraMenus2
    
    'NUEVO 2017
    'Contabilizacion
    ComprobarFechaContabilizadas
    
    If vParamAplic.NumeroInstalacion = vbEuler Then
        CadenaDesdeOtroForm = "not fechaent is null AND 1"
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", "scaalb", CadenaDesdeOtroForm, "1")
        If Val(CadenaDesdeOtroForm) > 0 Then frmAvisosAlb.Show vbModal
        CadenaDesdeOtroForm = ""
    End If
    
    
    
    
      'DAVID enero2021
       CarpetaExportar
    
    '--
    Screen.MousePointer = vbDefault
    
    

    
End Sub

Private Sub CarpetaExportar()
On Error Resume Next
    If Dir(App.Path & "\Exportar", vbDirectory) = "" Then MkDir (App.Path & "\Exportar")
    If Err.Number <> 0 Then
        MsgBox "Error carpeta EXPORTAR." & vbCrLf & Err.Description, vbExclamation
        Err.Clear
    End If

End Sub


Private Sub PuntoDeMenuVisible(ByRef MnPuntoDMenu As Menu, B As Boolean)
    If MnPuntoDMenu.visible Then MnPuntoDMenu.visible = B
    
End Sub




Private Sub Form_Load()
'Formulario Principal

    CargaImagen

    PrimeraVez = True
    

    
    'Botones
    With Me.Toolbar1
        .ImageList = frmPpal.ImgListPpal
        .Buttons(1).Image = 1   'Articulos
        .Buttons(2).Image = 2   'Movimientos Articulos
        
        .Buttons(5).Image = 3   'Clientes
        .Buttons(6).Image = 4   'Proveedores

        .Buttons(9).Image = 5   'Ofertas Clientes
        .Buttons(10).Image = 6   'Pedidos Clientes
        .Buttons(11).Image = 7   'Albaranes Clientes
        .Buttons(12).Image = 8   'Hist. Albaranes Clientes (Facturas)

        .Buttons(13).Image = 34   'FACTURAS MOSTRADOR

        .Buttons(15).Image = 9   'Pedidos Proveedor
        .Buttons(16).Image = 10   'Albaranes Proveedor
        .Buttons(17).Image = 11   'Facturas Proveedor
        .Buttons(18).Image = 12   'Recepcion Facturas Proveedor
        
        .Buttons(21).Image = 15   'Mantenimientos
        
        .Buttons(22).Image = 16   'N� Serie
        'Si tiene PARTES seran partes
        If vParamAplic.Ariagro = "" Then
            .Buttons(23).Image = 23   'Avisos
            .Buttons(23).ToolTipText = "Aviso reparacion"
        Else
            .Buttons(23).Image = 35   'Partes trabajo
            .Buttons(23).ToolTipText = "Partes de trabajo"
        End If
        
        
        If vParamAplic.NumeroInstalacion <> 4 Then
            .Buttons(24).Image = 13 'Gastos tecnicos
            .Buttons(24).ToolTipText = "Gastos t�nicos"
        Else
            .Buttons(24).Image = 37 'Gastos tecnicos
            .Buttons(24).ToolTipText = "Reloj"
        End If
        
        
        
        .Buttons(25).Image = 22 'Consulta precio articulo
        .Buttons(26).Image = 19 'Pantalla venta del TPV
        .Buttons(27).Image = 21 'Agenda
        .Buttons(28).Image = 20 'Agenda
        
        .Buttons(30).Image = 14 'Salir
    End With
    LeerEditorMenus
    PonerDatosFormulario False
    
       
    'Fijar primer dia la semana en vbMyMonday
    'Para el calendario.
    FijarPrimerDiaSemana
    
    
    If InstalacionEsEulerTaxco Then
        Me.WindowState = 0
        Me.Width = Screen.Width - 1200
        Me.Height = Screen.Height - 3000

    Else
        Me.WindowState = 2
    End If
    
       
    
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\arifon2.dll")
    If Err.Number <> 0 Then
        Me.Picture = LoadPicture()
        Err.Clear
    End If
End Sub




Private Sub PonerDatosFormulario(DesdeCambiarEmpresa As Boolean)
Dim Config As Boolean


    If Not DesdeCambiarEmpresa Then
        Config = (vEmpresa Is Nothing) Or (vParam Is Nothing) Or (vParamAplic Is Nothing)
    
        If Config Then HabilitarSoloPrametros_o_Empresas False
    End If
    
    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
    PonerMenusNivelUsuario

    'Si no hay carpeta interaciones, no habra integraciones
'    Me.mnComprobarPendientes.Enabled = vConfig.Integraciones <> ""


    'Habilitar
    If DesdeCambiarEmpresa Then
        ReestablecerMenus
        HabilitarSoloPrametros_o_Empresas True
    End If


    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor
    

    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    HacerEND
End Sub

Private Sub HacerEND()

'Formulario Principal
Dim Cad As String


    'Elimnar bloquo BD
    Set vUsu = Nothing
    Set vConfig = Nothing
    Set vEmpresa = Nothing
    
    Set vParam = Nothing
    Set vParamAplic = Nothing
    
    
    TerminaBloquear
    
    'cerrar las conexiones
    conn.Close
    CerrarConexionConta
    End
End Sub






Private Sub mnAdmAgen2_Click(Index As Integer)
    Select Case Index
    Case 0
        'Ventas x agente
        AbrirListado2 36
        
    
    Case 1
        
        'beneficio por agente
        AbrirListado2 37
    
    Case 2
'        frmListado3.Opcion = 37    'LISTADO 3
'        frmListado3.Show vbModal
        AbrirListado3 37
    Case 4
        frmFacComisionAgen.Show vbModal
    
    Case 5
'        frmListado3.Opcion = 31
'        frmListado3.Show vbModal
        AbrirListado3 31
    Case 7
        'Comisiones ECO
        AbrirListado3 43
    Case 8
        AbrirListado3 46
    Case 9
        AbrirListado2 49
    End Select
End Sub





Private Sub mnAdmGastosTec_Click()
'Gastos T�cnicos
    frmAdmGasTec.Show vbModal
End Sub

Private Sub mnAdministra_Click(Index As Integer)
    Select Case Index
    Case 1
        'Caluclo de riesgo
        If vUsu.Nivel > 0 Then
            MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
        Else
            AbrirListado2 31
        End If
        
    Case 2
        'Lisatdo riesgo
        frmInformesNew6.OpcionListado = 1
        frmInformesNew6.Show vbModal
    Case 3
        'Informe prevision de tesoreria
'        frmListado3.Opcion = 1
'        frmListado3.Show vbModal
        AbrirListado3 1
    Case 4
        'MsgBox "Comming soon men!!!!", vbExclamation
        'Correcion de costes de articulos varios en facturas
        frmFacCosteLin.Show vbModal
        
    Case 5
        'Modificar coste estadistica ventas
'        frmListado3.Opcion = 11
'        frmListado3.Show vbModal
        AbrirListado3 11
    Case 6
        'FLOTAS. Despliega submenu
    Case 7
        AbrirListado3 40
        
    Case 8
        'Ventas a credito
        AbrirListado3 25
    End Select
        
End Sub

Private Sub mnAdmNominas_Click()
'Nominas y Gastos
    frmAdmNominas.Show vbModal
End Sub

Private Sub mnAdmTrabajadores_Click()
    If vUsu.Nivel2 = 2 Then Exit Sub
    frmAdmTrabajadores.Show vbModal
End Sub

Private Sub mnAgenda_Click()
    'MsgBox "Se ha producido un error abriendo la agenda", vbExclamation
    'FALTA###
    'frmMainCalendar.Show
    
    MsgBox "Avise soporte t�cnico. Falta OCX Codejock", vbCritical
    
    
End Sub

Private Sub mnAguaLin_Click(Index As Integer)

    Select Case Index
    Case 0
        frmAguaContadoresGR.Show vbModal
    
    Case 1
        AbrirListado3 52
        
    Case 2
        AbrirListado3 51
    
    
    Case 3
        'Resumen facturacion 53
        AbrirListado3 53
        
        
    Case 4
        'Listado para rellenar modelos 100,101,102 EPSAR
        'de facturaciones canon generalitat
        AbrirListado3 55
    Case 5
        'Listado exportacion contadores
        AbrirListado3 60
        
    Case 6
        frmListado4.Opcion = 13
        frmListado4.Show vbModal
        
    Case 7
        'Declaracion detallada ejereccio
        AbrirListado3 58
    
    Case 8
        AbrirListado3 74
    Case 10
        frmAguaCalibresGR.Show vbModal
    Case 11
        frmAguaParamGR.Show vbModal
    End Select
End Sub

Private Sub mnAlbaranesB_Click()
    AbrirFormularioGeneral "Presupuestos", "ALZ", False
'    PonerCaption False, "Presupuestos"
'    If vParamAplic.TipoFormularioClientes = 0 Then
'        If vParamAplic.HaciendoFrmulariosGrandes Then
'            frmFacEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
'            frmFacEntAlbaranesGR.hcoCodTipoM = "ALZ"
'            frmFacEntAlbaranesGR.EsHistorico = False
'            frmFacEntAlbaranesGR.Show vbModal
'        Else
'            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
'            frmFacEntAlbaranes2.hcoCodTipoM = "ALZ"
'            frmFacEntAlbaranes2.EsHistorico = False
'            frmFacEntAlbaranes2.Show vbModal
'        End If
'    Else
'        frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
'        frmFacEntAlbSAIL.hcoCodTipoM = "ALZ"
'        frmFacEntAlbSAIL.EsHistorico = False
'        frmFacEntAlbSAIL.Show vbModal
'    End If
'    PonerCaption True, ""
End Sub


'NOVIEMBRE Intentando unificar las llamadas a abri formulario
Private Sub AbrirFormularioGeneral(textoCaption As String, tipoMov As String, EnHistorico As Boolean)
    PonerCaption False, textoCaption
    If vParamAplic.TipoFormularioClientes = 0 Then
        If vParamAplic.HaciendoFrmulariosGrandes Then
            frmFacEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranesGR.hcoCodTipoM = tipoMov
            frmFacEntAlbaranesGR.EsHistorico = EnHistorico
            frmFacEntAlbaranesGR.Show vbModal
        Else
            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes2.hcoCodTipoM = tipoMov
            frmFacEntAlbaranes2.EsHistorico = EnHistorico
            frmFacEntAlbaranes2.Show vbModal
        End If
    Else
        frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbSAIL.hcoCodTipoM = tipoMov
        frmFacEntAlbSAIL.EsHistorico = EnHistorico
        frmFacEntAlbSAIL.Show vbModal
    End If
    PonerCaption True, ""
End Sub








Private Sub mnAlmActualizarInve_Click()
    AbrirListado (14)
End Sub

Private Sub mnAlmAlPropios_Click()
    frmAlmAlPropios.Show vbModal
End Sub

Private Sub mnAlmArticulos_Click()

  Dim ElNuevo As Boolean


    PonerCaption False, mnAlmArticulos.Caption
    
    ElNuevo = vParamAplic.HaciendoFrmulariosGrandes
    
    ElNuevo = True 'Siempre sale el grande
    
    
    If ElNuevo Then
        frmAlmArticulosGr.DatosADevolverBusqueda = ""
        frmAlmArticulosGr.Show vbModal
    Else
        frmAlmArticulos.DatosADevolverBusqueda = ""
        frmAlmArticulos.Show vbModal
    End If
    PonerCaption True
End Sub


Private Sub mnAlmCategoria_Click()
    'categorias de articulos
    frmAlmCategorias.Show vbModal
End Sub

Private Sub mnAlmControlStockDesdeInv_Click()

    AbrirListado3 27
End Sub

Private Sub mnAlmEntradaInve_Click()
    frmAlmInventarioGR.Show vbModal
End Sub

Private Sub mnAlmFamiliaArticulo_Click()
    frmAlmFamiliaArticulo.Show vbModal
End Sub


Private Sub mnAlmHcoInven_Click()
    frmAlmHcoInvenGR.Show vbModal
End Sub

Private Sub mnAlmListadoInve_Click()
    AbrirListado (13)
End Sub

Private Sub mnAlmListadosVarios_Click(Index As Integer)
    
    

    Select Case Index
    Case 0
            'Informe de Stocks a una Fecha
            AbrirListado (19)
    Case 1
            'Stocks por meses
            'frmListado3.Opcion = 4
            'frmListado3.Show vbModal
            AbrirListado3 4
    Case 2
            'frmListado3.Opcion = 26
            'frmListado3.Show vbModal
            AbrirListado3 26
    Case 3
            'frmListado3.Opcion = 35
            'frmListado3.Show vbModal
            AbrirListado3 35
    Case 4
            'Informe stock minimo
            AbrirListado (100)
            
    Case 5
            AbrirListado2 54
    End Select
End Sub

Private Sub mnAlmListComponentes_Click()
'Informe de articulos q estan compuestos de otros articulos
    AbrirListado (11)
End Sub

Private Sub mnAlmListInactivos_Click()
    AbrirListado (15)
End Sub

Private Sub mnAlmListMaxMin_Click()
'Informe de Stocks Maximos y Minimos
    AbrirListado (18)
End Sub

Private Sub mnAlmListMovim_Click()
    AbrirListado (9)
End Sub

Private Sub mnAlmListValoracion_Click()
    AbrirListado (17)
End Sub

Private Sub mnAlmMarcas_Click()
    frmAlmMarcas.Show vbModal
End Sub

Private Sub mnAlmMovimArticulos_Click()
    
    frmAlmMovimArticulos.Show vbModal
    
End Sub

Private Sub mnAlmMovimArticulosSt_Click()
    frmAlmMovArtSaldo.Show vbModal
End Sub

Private Sub mnAlmMovimientos_Click()
    frmAlmMovimientos.EsHistorico = False
    frmAlmMovimientos.hcoCodMovim = -1 'No carga el form al abrir
    frmAlmMovimientos.Show vbModal
End Sub

Private Sub mnAlmMovimientosHco_Click()
    frmAlmMovimientos.EsHistorico = True
    frmAlmMovimientos.hcoCodMovim = -1
    frmAlmMovimientos.Show vbModal
End Sub

Private Sub mnAlmNumLotes_Click()
'numero de lote de los art�culos
    frmAlmNumLote.Show vbModal
End Sub





Private Sub mnAlmTipoUnidad_Click()
    frmAlmTipoUnidad.Show vbModal
End Sub

Private Sub mnAlmTomaInven_Click()
    AbrirListado (12)
End Sub

Private Sub mnAlmTraspaso_Click()
    frmAlmTraspaso.EsHistorico = False
    frmAlmTraspaso.hcoCodMovim = -1
    frmAlmTraspaso.Show vbModal
End Sub

Private Sub mnAlmTraspasoHco_Click()
    frmAlmTraspaso.EsHistorico = True
    frmAlmTraspaso.hcoCodMovim = -1
    frmAlmTraspaso.Show vbModal
End Sub

Private Sub mnAlmUbicacion_Click()
    frmAlmUbicaciones.Show vbModal
End Sub

Private Sub mnAlmValoracionInve_Click()
    AbrirListado (16)
End Sub


Private Sub mnArticulos2_Click(Index As Integer)
    Select Case Index
    Case 0
        frmVarios.Opcion = 1
        frmVarios.Show vbModal
    Case 1
        AbrirListado3 49
    Case 2
        'If vUsu.Nivel > 0 Then Exit Sub
        
        'Bloquear proceso
        If BloqueoManual("CambioArt", "1") Then
            frmAlmCambRef.Show vbModal
            DesBloqueoManual "CambioArt"
        End If
        
        
        
        
    End Select
End Sub

Private Sub mnBackUp_Click()
'Copia de seguridad de toda la base de datos
    frmBackUP.Show vbModal
End Sub

Private Sub mnBorrarAvisosCerrados_Click()
    AbrirListado 83
End Sub




Private Sub mnCambioEmpresa_Click()
    
    'If Not (Me.ActiveForm Is Nothing) Then
    '    MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation
    '    Exit Sub
    'End If

    'Borramos temporal
    conn.Execute "Delete from zbloqueos where codusu = " & vUsu.Codigo


    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"
    
    frmLogin.Show vbModal

    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    conn.Close
    ConnConta.Close


    'Abre la conexi�n a BDatos:Ariges
    If AbrirConexion() = False Then
        MsgBox "La aplicaci�n no puede continuar sin acceso a los datos. ", vbCritical
        End
    Else
        Set vParam = Nothing
        Set vParamAplic = Nothing
        'Carga Parametros Generales y Contables de la empresa
        LeerParametros
    End If


    'Abrir conexi�n a la BDatos de Contabilidad para acceder a
    'Tablas: Cuentas, Tipos IVA
    If AbrirConexionConta(False) = False Then
        MsgBox "La aplicaci�n no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
        End
    End If

    
    
    Set vEmpresa = Nothing
    'LeerEmpresaParametros
    
     'Carga los Datos B�sicos de la empresa
    LeerDatosEmpresa
    
    
    'Carga los Niveles de cuentas de Contabilidad de la empresa
    LeerNivelesEmpresa
    'HaMostradoCanal2_elB = False
    
    
    
    If vParamAplic.QueEmpresaEs = 2 Then
        Me.Hide
        frmPpalGessocial.Show vbModal, Me
        Me.Show
        mnCambioEmpresa_Click
        Exit Sub
    End If
    
    
    
    
    
'    PonerDatosFormulario
    PonerDatosVisiblesForm

    'Ponemos primera vez a false
    PonerDatosFormulario True
    PrimeraVez = True
    Form_Activate

    

    Screen.MousePointer = vbDefault
End Sub


Private Sub mnCambioPwd_Click()
    frmListado4.Opcion = 4
    frmListado4.Show vbModal
End Sub

Private Sub mnCartaRenovaMante_Click()
    AbrirListado 78
End Sub

'Private Sub mnCheckVersion_Click()
''    Screen.MousePointer = vbHourglass
''    LanzaHome "webversion"
''    espera 2
''    Screen.MousePointer = vbDefault
'End Sub


Private Sub mnComAlbMan_Click()
    'Mantenimiento de Albaranes a Proveedor
    If vParamAplic.TipoFormularioClientes = 0 Then
            
        frmComEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmComEntAlbaranesGR.EsHistorico = False
        frmComEntAlbaranesGR.Show vbModal
    
        '    frmComEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
        '    frmComEntAlbaranes.EsHistorico = False
        '    frmComEntAlbaranes.Show vbModal
    
     Else
        frmComEntAlbaranSA.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmComEntAlbaranSA.EsHistorico = False
        frmComEntAlbaranSA.Show vbModal
    
    End If
End Sub



Private Sub mnComContFactu_Click()
'Contabilizar Facturas
    AbrirListado (224) 'Para pedir datos
End Sub

Private Sub mnComCtrlAlb_Click(Index As Integer)
    If Index = 1 Then
        'frmComAlbAsignar.Show vbModal
        frmComCtrDoc.Show vbModal
    End If
End Sub

Private Sub mnComDirecciones_Click()
    frmComDirecciones.Show vbModal
End Sub










Private Sub mnComEstadisticaLin_Click(Index As Integer)
    Select Case Index
    Case 0
        'Listado de compras por proveedor
        AbrirListadoOfer (310)
    Case 1
        'Listado de compras por Familia
        AbrirListadoOfer (311)
    Case 2
        frmVarios.Opcion = 11
        frmVarios.Show vbModal
    Case 3
        'Listado de alb compras por proveedor
        AbrirListadoOfer (312)
    Case 4
         'frmListado3.Opcion = 7
         'frmListado3.Show vbModal
         AbrirListado3 7
    Case 5
        AbrirListado2 50
    End Select
End Sub



Private Sub mnComFacturar_Click()
   
    frmComFacturarGR.Codprove = -1
    frmComFacturarGR.Show vbModal
    
End Sub

Private Sub mnComHcoAlbaranes_Click()
'Historico albaranes de compras a proveedores
     If vParamAplic.TipoFormularioClientes = 0 Then
        
        frmComEntAlbaranesGR.EsHistorico = True
        frmComEntAlbaranesGR.Show vbModal
    
        '    frmComEntAlbaranes.EsHistorico = True
        '    frmComEntAlbaranes.Show vbModal
       
    Else
        frmComEntAlbaranSA.EsHistorico = True
        frmComEntAlbaranSA.Show vbModal
    End If
End Sub

Private Sub mnComHcoFacturas_Click()
    If vParamAplic.TipoFormularioClientes = 0 Then
        frmComHcoFacturas2GR.hcoCodMovim = ""
        frmComHcoFacturas2GR.Show vbModal

    Else
        'SAIL
        frmComHcoFacturSA.hcoCodMovim = ""
        frmComHcoFacturSA.Show vbModal
    End If
End Sub







Private Sub mnComInfVarios1_Click(Index As Integer)
    Select Case Index
    'ESTA DENTRO DE PROVEEDORES
'    Case 0
'        'Informe de Proveedores
'        AbrirListado (58)   ': Informe Proveedores
    Case 1
        'Etiquetas de proveedores
        AbrirListadoOfer (305) '305: Informe Etiquetas de Proveedores
    
    Case 2
        'Cartas a proveedores
        AbrirListadoOfer (306) '306: Informe Cartas a Proveedores
        
    Case 3
         AbrirListado 101
    
    End Select
    
End Sub

Private Sub mnComPedidosLin_Click(Index As Integer)
    'Cuelga de COMPRAS --- PEDIDOS---
    Select Case Index
    Case 0, 1
        'Mnatenimiento de Pedidos de compras
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmComEntPedidos2.MostrarDatos = ""
            frmComEntPedidos2.EsHistorico = Index = 1
            frmComEntPedidos2.Show vbModal
        Else
            'SAIL
            frmComEntPedidosSa.MostrarDatos = ""
            frmComEntPedidosSa.EsHistorico = Index = 1
            frmComEntPedidosSa.Show vbModal
        End If
    'Case 1
    '    frmComEntPedidos.MostrarDatos = ""
    '    frmComEntPedidos.EsHistorico = True
    '    frmComEntPedidos.Show vbModal
    Case 2
        'Listado de material pendiente de recibir
        AbrirListadoOfer (307) '307: List. Materia pte recibir
    Case 4
        'Propuesta de pedido
        AbrirListado2 32
    
    Case 5
        'Propuesta de pedido
        AbrirListado3 71

    End Select
    
End Sub

Private Sub mnComPreProv_Click(Index As Integer)
    Select Case Index
    Case 0
        'precios proveedor
        frmComPreciosProv2.NuevoDato = "" 'Para que no se poing en modo insercion
        frmComPreciosProv2.Show vbModal
    Case 1
        'Dto proveedor
        frmComDtosFamMarca.Show vbModal
    
    Case 2
        'Copiar desde venta
        CadenaDesdeOtroForm = "V"
        AbrirListado2 28
    Case 3
        frmFacActPrecios2.Proveedor = True
        frmFacActPrecios2.Show vbModal
    End Select
End Sub

Private Sub mnComProveedores_Click()
'Compras. Mantenimiento de Proveedores

'    If vParamAplic.HaciendoFrmulariosGrandes Then
        frmComProveedoresGr.Show vbModal
'    Else
'        frmComProveedores.Show vbModal
'    End If



    
End Sub


Private Sub mnComProveVarios_Click()
'Proveedores varios
    frmComProveV.Show vbModal
End Sub

Private Sub mnComPteFacturar_Click()
'Listado de Albaranes pendientes de Factura
    AbrirListadoOfer (308) '308: List. Albaranes pte facturar
End Sub



Private Sub mnConfManteUsuarios_Click()
'Mantenimiento de Usuarios
    If vUsu.Nivel > 0 Then Exit Sub
    frmMantenusu2.Show vbModal
      
End Sub

Private Sub mnConfParamAplic_Click()
'Parametros de la Aplicaci�n
    Screen.MousePointer = vbHourglass
    Load frmConfParamAplic
    frmConfParamAplic.Show vbModal
    
End Sub



Private Sub mnConfParamGenerales_Click()
'Parametros generales de la Empresa

    frmConfParamGral.Show vbModal
End Sub



Private Sub mnConfParamRpt_Click()
'Parametros para informes de Crystal Report
    frmConfParamRpt.Show vbModal
End Sub

Private Sub mnConTMovimiento_Click()
'Mantenimientos de los tipos de movimientos
    frmConfTipoMov.Show vbModal
End Sub


Private Sub mnCRM_Click(Index As Integer)
    
        Select Case Index
        Case 0
            frmCRMMto.DesdeElCliente = 0 'No clien
            frmCRMMto.TipoPredefinido = 0   'Ninguno
            frmCRMMto.Show vbModal
            
        Case 1
            frmCRMtipos.Show vbModal
        
        Case 2
            frmCRMVarios.Opcion = 0
            frmCRMVarios.Show vbModal
            
        Case 3
            frmListadoOfer.OpcionListado = 406
            frmListadoOfer.Show vbModal
        Case 4
            frmCRMVarios.Opcion = 1
            frmCRMVarios.Show vbModal
        
        Case 5
            frmListado5.OpcionListado = 22
            frmListado5.Show vbModal
        End Select
        
End Sub

Private Sub mnDtoCantidad_Click()
    frmFacDtoUd.Show vbModal
End Sub



Private Sub mnEliminarFacturas_Click()
    AbrirListado 97
End Sub

Private Sub mnEnvioFactuasMail_Click(Index As Integer)
    AbrirListadoOfer 315 + Index
End Sub

Private Sub mnEstadisticaReparacionTecnico_Click()
    AbrirListado2 2
End Sub


Private Sub mnEstVtasArituclo_Click()
    'frmListado3.Opcion = 18
    'frmListado3.Show vbModal
    AbrirListado3 18
End Sub

Private Sub mnEtiqMante_Click()
    AbrirListado 79
End Sub



Private Sub mnFacActividades_Click()
    frmFacActividades.Show vbModal
End Sub

Private Sub mnFacAgentesCom_Click()
    frmFacAgentesCom.Show vbModal
End Sub

Private Sub mnFacAlbDevolucion_Click()
    
    If vParamAplic.TipoFormularioClientes = 0 Then
    
        If vParamAplic.HaciendoFrmulariosGrandes Then
            frmFacEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranesGR.hcoCodTipoM = "DEV"
            frmFacEntAlbaranesGR.EsHistorico = False
            frmFacEntAlbaranesGR.Show vbModal

        Else
            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes2.hcoCodTipoM = "DEV"
              frmFacEntAlbaranes2.EsHistorico = False
            frmFacEntAlbaranes2.Show vbModal
        End If
    End If
End Sub

Private Sub mnFacAlbMostrador_Click()
'Abre el formulario de Albaranes para introducir el Albaran de Mostrador
'y desde este generar la Factura de mostrador
    If vParamAplic.TipoFormularioClientes = 0 Then
        If vParamAplic.HaciendoFrmulariosGrandes Then
            frmFacEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranesGR.hcoCodTipoM = "ALM"
            frmFacEntAlbaranesGR.EsHistorico = False
            frmFacEntAlbaranesGR.Show vbModal

        
        Else
            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes2.hcoCodTipoM = "ALM"
            frmFacEntAlbaranes2.EsHistorico = False
            frmFacEntAlbaranes2.Show vbModal
        End If
        
    Else
            frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbSAIL.hcoCodTipoM = "ALM"
            frmFacEntAlbSAIL.EsHistorico = False
            frmFacEntAlbSAIL.Show vbModal
    End If
End Sub


Private Sub mnFacAlbRectifica_Click()
'Facturas Rectificativas
    'Abre el formulario de Albaranes para introducir el Albaran Rectificativo
    'y desde este generar la Factura Rectificativa
    If vParamAplic.TipoFormularioClientes = 0 Then
    
        If vParamAplic.HaciendoFrmulariosGrandes Then
            frmFacEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranesGR.hcoCodTipoM = "ART"
            frmFacEntAlbaranesGR.EsHistorico = False
            frmFacEntAlbaranesGR.Show vbModal

        
        Else
    
            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes2.hcoCodTipoM = "ART"
            frmFacEntAlbaranes2.EsHistorico = False
            frmFacEntAlbaranes2.Show vbModal
        End If
    Else
        frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbSAIL.hcoCodTipoM = "ART"
        frmFacEntAlbSAIL.EsHistorico = False
        frmFacEntAlbSAIL.Show vbModal
    End If
End Sub

Private Sub mnFacAlbxArtic_Click()
'Informe de Albaranes por Articulo
    AbrirListadoPed (49)
End Sub


Private Sub mnFacBancosPropios_Click()
    frmFacBancosPropios.Show vbModal
End Sub





Private Sub mnFacCartas_Click()
'Mantenimiento de Cartas
    frmFacCartasOferta.Show vbModal
End Sub


Private Sub mnFacClientes_Click()
Dim VerGrande As Boolean
'Mantenimiento de Clientes

   
   
   'Sept 2020
   ' TOOOODO GRANDE
   ' QUitamos el formulario clientresnormal (del proyecto)
    

    PonerCaption False, "Clientes"
    VerGrande = False
    'If vParamAplic.HaciendoFrmulariosGrandes Then
        VerGrande = True
    'Else
    '    If vParamAplic.NumeroInstalacion = vbTaxco Then VerGrande = True
    'End If
    Screen.MousePointer = vbHourglass
    If VerGrande Then
        frmFacClientesGr.Show vbModal
    Else
     '   frmFacClientes.Show vbModal
    End If
    PonerCaption True
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnFacClientesPot_Click()
    frmFacClienPot.Show vbModal
End Sub

Private Sub mnFacClientesV1_Click()
'Mantenimiento de Clientes Varios
    frmFacClientesV.Show vbModal
    
End Sub



Private Sub mnFacContFactu_Click()
'Contabilizar Facturas
    AbrirListado (223) 'Para pedir datos
End Sub






Private Sub mnFacEntAlbaran_Click()


    PonerCaption False, mnFacEntAlbaran.Caption

      If vParamAplic.TipoFormularioClientes = 0 Then
     
        If vParamAplic.HaciendoFrmulariosGrandes Then
            frmFacEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranesGR.hcoCodTipoM = "ALV"
            frmFacEntAlbaranesGR.EsHistorico = False
            frmFacEntAlbaranesGR.Show vbModal
        Else
            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes2.hcoCodTipoM = "ALV"
            frmFacEntAlbaranes2.EsHistorico = False
            frmFacEntAlbaranes2.Show vbModal
        End If
    
    Else
        frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbSAIL.hcoCodTipoM = "ALV"    'IIf(vParamAplic.NumeroInstalacion = vbTaxco, "ALO", "ALV")
        frmFacEntAlbSAIL.EsHistorico = False
        frmFacEntAlbSAIL.Show vbModal
    End If
    PonerCaption True
End Sub











Private Sub mnFacEstadistica2_Click(Index As Integer)
    Select Case Index
    Case 0
        'Por proveedor
        AbrirListado2 6
    Case 1
    
        'Ventas por agente
        AbrirListado2 16
    Case 2
        'Detalle facturacion clientes
        AbrirListadoOfer (231)
    Case 3
        'Estadistica margen ventas por art�culo
        AbrirListado (246)
    Case 4
        'Vtas x tipo d precio
        AbrirListado3 38
    Case 5
        'Articulos ams vendidos
        AbrirListado3 39
    Case 6
        'ALZIRA
        frmListado4.Opcion = 7
        frmListado4.Show vbModal
    
    Case 7
        'Listado pedidos por "peticion " cliente (si-no)
        AbrirListado3 63
    Case 8
        'MAYO 2019
        'AbrirListado3 69  'Costes euler
        AbrirListado2 53  'Costes euler
    End Select
End Sub

Private Sub mnFacEstVentaCliente_Click()
'Estadistica Ventas por cliente
    AbrirListadoPed (227)
    BorrarTempInformes
End Sub

Private Sub mnFacEstVentaFam_Click()
'Listado de estadistica ventas por familia de articulo
    AbrirListadoOfer (230)
End Sub

Private Sub mnFacEstVentaMes_Click()
'Estadistica Ventas por Meses
    AbrirListadoPed (229)
    
End Sub

Private Sub mnFacEstVentaTraba_Click()
'Estadistica Ventas por Trabajador
    AbrirListadoPed (228)
End Sub



Private Sub mnFacFacturarAlb_Click()
Dim B As Boolean
'Facturacion de Albaranes de Ventas
    If vParamAplic.NumeroInstalacion = vbFenollar Then
        MsgBox "No puede facturar desde este punto de men�.", vbExclamation
        Exit Sub
    End If
    
    B = False
    If vParamAplic.TipoFormularioClientes = 0 Then
        B = True
    Else
        If vParamAplic.NumeroInstalacion = vbTaxco Then B = True
    End If
    If B Then

        frmListadoPed.codClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (52)
        
    Else
        'PARA sail
        frmFacturaCliSail.Show vbModal
    End If
End Sub

Private Sub mnFacFormasPago_Click()
    frmFacFormasPago.Show vbModal
End Sub



Private Sub mnFacHcoAlbaranes_Click()
'Hist�rico de Albaranes eliminados
    If vParamAplic.TipoFormularioClientes = 0 Then
        If vParamAplic.HaciendoFrmulariosGrandes Then
            frmFacEntAlbaranesGR.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranesGR.hcoCodTipoM = "ALV"
            frmFacEntAlbaranesGR.EsHistorico = True
            frmFacEntAlbaranesGR.Show vbModal

        
        Else
            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes2.hcoCodTipoM = "ALV"
            frmFacEntAlbaranes2.EsHistorico = True
            frmFacEntAlbaranes2.Show vbModal
        End If
    Else
        frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbSAIL.hcoCodTipoM = "ALV"
        frmFacEntAlbSAIL.EsHistorico = True
        frmFacEntAlbSAIL.Show vbModal
    
    End If
End Sub

Private Sub mnFacHcoFacturas_Click()
'Hist�rico de Facturas
    PonerCaption False, "Facturas"
    frmFacHcoFacturas2.hcoCodMovim = ""
    frmFacHcoFacturas2.Show vbModal
    PonerCaption True
End Sub


Private Sub mnFacIncidencias_Click()
    frmIncidencias.Show vbModal
End Sub

Private Sub mnFacIncumPlazos_Click()
'Incumplimiento de los Plazos de Entrega
    
    AbrirListadoPed (51)
End Sub










Private Sub mnFacInformesVarios_Click(Index As Integer)
    Select Case Index
    Case 0
        'Informe de Clientes Inactivos
        AbrirListadoOfer (46) '46: Informes Clientes Inactivos
    Case 1
        'Informe de Clientes
        AbrirListadoOfer (47) '47: Informes Clientes
    Case 2
        'Informe de Altas de Nuevos Clientes
        AbrirListadoOfer (48) '48: Informes Altas Clientes

    Case 3
        'Etiquetas de clientes
        AbrirListadoOfer (90) '90: Informe Etiquetas de Clientes
    Case 4
        'Cartas a clientes
         AbrirListadoOfer (91) '91: Informe Cartas a Clientes
    Case 5
        'Listado de etiquetas de los bultos
        AbrirListado 95
    Case 6
        'Listado de telefonos por clientes
        AbrirListado3 41
    Case 7
        'Listado de telefonos por clientes
        AbrirListado3 48
    
    End Select
    
End Sub

Private Sub mnFacOfertas_Click(Index As Integer)
    'Estan todos agrupados bajo el mismo mn
    
    Select Case Index
    Case 0, 5
            'Private Sub mnFaabrcEntOfertas_Click()
         If vParamAplic.TipoFormularioClientes = 0 Then
           
            Debug.Assert False
            If False Then
            
            
               EulerParam = DevuelveDesdeBD(conAri, "pathDocs", "eulerparam", "1", "1")
            
               frmFacEntOfertasGR.DatosOferta = ""
               frmFacEntOfertasGR.Show vbModal
            Else
               frmFacEntOfertas2.DatosOferta = ""
               frmFacEntOfertas2.EsHistorico = Index = 5
               frmFacEntOfertas2.Show vbModal
            End If
        Else
            frmFacEntOferSAIL.DatosOferta = ""
            frmFacEntOferSAIL.EsHistorico = Index = 5
            frmFacEntOferSAIL.Show vbModal
        End If

    Case 1
            'Private Sub mnFacGrupoPlant_Click()
            'Mantenimiento de Grupos de Plantillas
            frmFacGrupoPlantilla.Show vbModal
    
    Case 2
            'Private Sub mnFacPlantillas_Click()
            'Mantenimiento de Plantillas
            frmFacPlantilla.Show vbModal
        
    Case 3
            ' Private Sub mnFacOfeEfectuadas_Click()
            'Listado de Ofertas Efectuadas
            AbrirListadoOfer (34) '34: Informe Ofertas Efectuadas
    
    Case 6
        
            'Private Sub mnFacTrasHist_Click()
            'Traspaso de Ofertas a las tablas de Historico
            CadenaDesdeOtroForm = ""
            frmListadoOfer.OpcionListado = 36
            AbrirListadoOfer (36) 'NO IMPRIME LISTADO, hace traspaso de Ofertas de la tabla (scapre) a (schpre)
    
    Case 7
            frmListadoOfer.OpcionListado = 409
            AbrirListadoOfer (409) 'NO IMPRIME LISTADO, hace traspaso de Ofertas de la tabla (scapre) a (schpre)
    
    End Select
End Sub

Private Sub mnFacPedidos_Click(Index As Integer)
    'Estan todos agrupados bajo el mismo mn>
    Select Case Index
    Case 0, 1
        'Mantenimiento de Pedidos  Y Hist�rico de Pedidos
        PonerCaption False, "Pedidos"
        
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmFacEntPedidos.DatosADevolverBusqueda2 = ""
            frmFacEntPedidos.EsHistorico = Index = 1
            frmFacEntPedidos.Show vbModal
        Else
            frmFacEntPedSail.EsHistorico = Index = 1
                frmFacEntPedSail.Show vbModal
        End If
        
        If vParamAplic.NumeroInstalacion = vbFenollar Then
            If HaMostradoCanal2_El_B Then Me.mnAlbaranesB.visible = True
        End If
    'Case 2  es la barra de separacion
    
    Case 3
        'Confirmar pedido   mnFacConfirmPed_Click
        AbrirListadoOfer (40)
        
    Case 4
        'Pedido por articulo
        'Private Sub mnFacPedidoxArtic_Click()
        'Informe de Pedidos por Articulo
        AbrirListadoPed (41)
        
    Case 5
        'Private Sub mnFacPedidoxClien_Click()
        'Informe de Pedidos por Cliente
        AbrirListadoPed (44)
        
        
    Case 6
        'Private Sub mnFacDispStock_Click()
        'Resumen de Disponibilidad de Stocks
        AbrirListadoPed (42)
    
    Case 7
        'Pedido por zona
        frmListado2.Opcion = 26
        frmListado2.Show vbModal
        
    Case 9
        'Precio cliente
        frmFacConsultaPrecios2.Fecha = Now
        frmFacConsultaPrecios2.Show vbModal
    Case 10
        'Estadisitcas de veces consultado  precio/cliente
        frmVarios.Opcion = 2
        frmVarios.Show vbModal
    Case 12
        frmFacEntPedidRMA.Show vbModal
    Case 13
        frmFacEntPedPresu.Show vbModal
        
    Case 14
        frmFacPedidoAgrupados.Show vbModal
    Case 15
        frmListado5.OpcionListado = 36
        frmListado5.Show vbModal
        
    End Select
    PonerCaption True
End Sub





Private Sub mnFacPreFacturar_Click()
' Previsi�n Facturacion de Albaranes
    frmListadoPed.codClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (50) 'NO IMPRIME LISTADO
End Sub



Private Sub mnFacReImpFactu_Click()
'Reimprimir Factuas ya contabilizadas
    AbrirListadoOfer 226
End Sub

Private Sub mnFacRutas_Click()
    frmFacRutas.Show vbModal
End Sub

Private Sub mnFacSituaciones_Click()
    frmFacSituaciones.Show vbModal
End Sub












'---------------------------------------------
'
'  Unico punto de menu para las tarifas venta
Private Sub mnFacTarVen_Click(Index As Integer)
        Select Case Index
    Case 0
        'Tarifas Venta
        frmFacTarifas.Show vbModal
    Case 1
        'Listado Precios
        frmFacTarifasPrecios.Show vbModal
    Case 2
        'Precios especiales
        frmFacPreciosEspecial.CadenaSituarData = ""
        frmFacPreciosEspecial.Show vbModal
    
    Case 3
        'PROMOCIONES
        frmFacPromociones.Show vbModal
    
    Case 4
        'Dots familia marca
        frmFacDtosFamMarca.Show vbModal
    
    Case 5
        'dtos por activiad
        frmFacDtosAsignar.Show vbModal
    
    Case 6
        'Bonificacines factura
        frmFacBonificacion.Show vbModal
    
    Case 8
        'Actualizar precios actuales y especiales
        frmFacActPrecios2.Proveedor = False
        frmFacActPrecios2.Show vbModal
    
    Case 9
        'Copiar desde compra
        CadenaDesdeOtroForm = ""
        AbrirListado2 28
    Case 11
        'Informe control margenes de tarifas
        AbrirListado (245)
    
    Case 12
        'Correcion
        AbrirListado 247
    Case 13
        'frmListado3.Opcion = 13
        'frmListado3.Show vbModal
        AbrirListado3 13
    Case 14
        frmTaxcoTarifa.Show vbModal
    End Select
End Sub

Private Sub mnFacturarCliente_Click()
    If vParamAplic.TipoFormularioClientes = 0 Then
    
        If False Then
           ' frmFacturacionCliMezcla.Show vbModal
        Else
    
            frmFacturacionCli.Show vbModal
        End If
    Else
        frmFacturaCliSail.ImprimirCertificacion = False
        frmFacturaCliSail.Show vbModal
    End If
End Sub

Private Sub mnFacturarPresupuestos_Click()
        frmListadoPed.codClien = "ALZ" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (52)
End Sub

Private Sub mnFacZonas_Click()
    frmFacZonas.Show vbModal
End Sub

Private Sub mnFacFormasEnvio_Click()
        
'    If vParamAplic.HaciendoFrmulariosGrandes Then
        frmFacFormasEnvio.Show vbModal
'    Else
'        frmFacFormasEnvio.Show vbModal
'    End If

End Sub

Private Sub mnFlotas_Click(Index As Integer)
    Select Case Index
    Case 0
        frmFlotaReg.DatosADevolverBusqueda = ""
        frmFlotaReg.Show vbModal
    Case 1
        frmFlotas.DatosADevolverBusqueda = ""
        frmFlotas.Show vbModal
        
    Case 2
    
        frmFlotasConceptos.DatosADevolverBusqueda = ""
        frmFlotasConceptos.Show vbModal
    End Select
End Sub

Private Sub mnFrecuencias_Click()
    frmFrecuenciasGR.Show vbModal
End Sub

Private Sub mnGasolFacturar_Click(Index As Integer)

    If Index = 0 Then
        If vUsu.Nivel > 0 Then Exit Sub
        frmListado5.OpcionListado = 34
        frmListado5.Show vbModal
    Else
        If Index = 1 Then
            frmListado5.OpcionListado = 31
            frmListado5.Show vbModal
        Else
            If Index = 5 Then
                'Ajuste formas de pago ALVIC
                frmListado5.OpcionListado = 45
                frmListado5.Show vbModal
            Else
                frmListadoPed.codClien = IIf(Index = 2, "ALD", "ALB")
                AbrirListadoPed (52)
            End If
        End If
    End If
End Sub

Private Sub mnGasolinera_Click(Index As Integer)
       
        'GASOLINERA
        If Index = 3 Then
            'Facturar
            
        ElseIf Index = 4 Then
            frmTrasAlvic.Show vbModal
        Else
            If Index = 1 Then
                AbrirFormularioGeneral "Gasolinera", "ALD", False
             Else
                AbrirFormularioGeneral "Tienda Gasol.", "ALB", False
            End If
        End If

End Sub

Private Sub mnHcoMaten_Click()
    frmManMantenimientosAnuGR.Show vbModal
End Sub


Private Sub mnHuertos1_Click(Index As Integer)
    If Index = 0 Then
        frmListado5.OpcionListado = 15
        frmListado5.Show vbModal
    Else
        AbrirListado3 72
    End If
End Sub

Private Sub mnInfManteAnulados_Click()
    AbrirListado 76
End Sub



Private Sub mnInfoAdm_Click(Index As Integer)
    Select Case Index
    Case 0
        '---------------------------
        '  CASE 0 y CASE 1 estan ahora en el submenu de Agentes dentro de administracion
        'frmListado2.Opcion = 36
        'frmListado2.Show vbModal
    Case 1
        'beneficio por agente
        'frmListado2.Opcion = 37
        'frmListado2.Show vbModal







    Case 3
        'beneficio por proveedor
        'frmListado2.Opcion = 40
        'frmListado2.Show vbModal
        AbrirListado2 40
    Case 4
        'frmListado2.Opcion = 41
        'frmListado2.Show vbModal
        AbrirListado2 41
        
    
    Case 5
        
        'Beneficio marca-agente-proveedor
        AbrirListado2 48
    Case 6
        'informe de articulos en promocion
        'frmListado3.Opcion = 5
        'frmListado3.Show vbModal
        AbrirListado3 5
    Case 7
        'informe de articulos en promocion
        'frmListado3.Opcion = 34
        'frmListado3.Show vbModal
        AbrirListado3 34
    Case 8
        'Ventas trabajaodr x dia
        'frmListado3.Opcion = 9
        'frmListado3.Show vbModal
        AbrirListado3 9
        
    Case 9
        'Listado ventas por forma de pago
        'frmListado3.Opcion = 19
        'frmListado3.Show vbModal
        AbrirListado3 19
        
    Case 10
        frmListado5.OpcionListado = 37
        frmListado5.Show vbModal
    End Select
End Sub



Private Sub mnInfTeoMant_Click()
    AbrirListado 77
End Sub




Private Sub mnListadoReparacionesEfectuadas_Click()
    AbrirListado2 1
End Sub

Private Sub mnLlamadas_Click(Index As Integer)
    Select Case Index
    Case 0
        frmLlamadas.Show vbModal
        
    Case 1
        frmLlamadasTipo.Show vbModal
    End Select
End Sub

Private Sub mnManAltas_Click()
'Listado Altas de Mantenimientos
    AbrirListado 73
End Sub

Private Sub mnManEntrada_Click()
    frmManMantenimientosGR.Show vbModal
End Sub



Private Sub mnManFichas_Click()
'Listado Fichas de Mantenimientos
    AbrirListado 72
End Sub

Private Sub mnManListado_Click()
'Listados de Mantenimientos
    AbrirListado 70
End Sub



Private Sub mnManPrevFac2_Click(Index As Integer)
    Select Case Index
    Case 0
         ' Previsi�n Facturacion de Albaranes de Mantenimiento
         '    frmListadoPed.CodClien = "ALM" 'utilizamos esta vble para pasarle el tipo de movimiento
         AbrirListadoPed (74) 'NO IMPRIME LISTADO
    Case 1
        'Facturacion de Mantenimientos
         AbrirListadoPed (75) 'NO IMPRIME LISTADO
    
    Case 3
    
        'frmListado3.Opcion = 23
        frmListado3.OtrosDatos = ""
        'frmListado3.Show vbModal
        AbrirListado3 23
    Case 4
        'FACTURACION
        'frmListado3.Opcion = 22
        frmListado3.OtrosDatos = ""
        'frmListado3.Show vbModal
        AbrirListado3 22
    End Select
End Sub

Private Sub mnManRevisiones_Click()
'Listado Revisiones de Mantenimientos
     AbrirListado 71
End Sub

Private Sub mnManServicioAsisTecn_Click()
    frmManSat.Show vbModal
End Sub



Private Sub mnManteneLOG_Click()
    Screen.MousePointer = vbHourglass
    Load frmLog
    DoEvents
    frmLog.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnManTiposContrato_Click()
    frmManTiposContrato.Show vbModal
End Sub


Private Sub mnMtoEuler_Click(Index As Integer)
    If Index > 0 Then
    
        If Index = 4 Then
            'Proyectos
            frmFacProyecto.Show vbModal
            
        Else
            'Alkbaranes.   ALE ALO ALR
            If Index = 1 Then
                frmFacEntAlbSAIL.hcoCodTipoM = "ALO"
            ElseIf Index = 2 Then
                frmFacEntAlbSAIL.hcoCodTipoM = "ALE"
            Else
                frmFacEntAlbSAIL.hcoCodTipoM = "ALR"
            End If
            frmFacEntAlbSAIL.EsHistorico = False
            frmFacEntAlbSAIL.Show vbModal
    
        End If
    End If
    
End Sub

Private Sub mnObra_Click(Index As Integer)
    Select Case Index
    Case 0
        frmObraCapitulo.Show vbModal
    Case 1
        frmObraActua.Show vbModal
    Case 2
        If InstalacionEsEulerTaxco Then
            frmEulerTrab.Show vbModal
        Else
            frmObrpartesTra.Show vbModal
        End If
    Case 3
        frmObraOT.Show vbModal
    Case 4
        frmEulerReloj.Show vbModal
    
    Case 6
        'Sept 2012
        frmObraListado.Opcion = 3
        frmObraListado.Show vbModal
    Case 8
        'Imprimir certificacion
        frmFacturaCliSail.ImprimirCertificacion = True
        frmFacturaCliSail.Show vbModal
    End Select
End Sub

Private Sub mnPortes_Click()
    frmFacPortes.Show vbModal
End Sub







Private Sub mnproduccion1_Click(Index As Integer)
    Select Case Index
    Case 0
        frmProdOrden.DatosADevolverBusqueda = ""
        frmProdOrden.Show vbModal
    Case 1
        frmProdEnvas.DatosADevolverBusqueda = ""
        frmProdEnvas.Show vbModal
    Case 2
        frmAlmDescCostesTasas.Show vbModal
    Case 3
        frmListLotes.Show vbModal
    Case 4
        frmAlmCalidad.Show vbModal

    Case 5
        'por si hay que reutilizarlo
        If vParamAplic.NumeroInstalacion = 5 Then
            frmListado5.OpcionListado = 20
            frmListado5.Show vbModal
        End If

    End Select
End Sub





Private Sub mnRecalPrSt_Click(Index As Integer)
    If vUsu.Nivel > 1 Then
        MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
        Exit Sub
    End If
    If Index = 0 Then
        frmListado3.Opcion = 6
    ElseIf Index = 1 Then
        frmListado3.Opcion = 20
    Else
        frmListado3.Opcion = 21
    End If
    frmListado3.Show vbModal
End Sub



Private Sub mnRectifInve_Click(Index As Integer)
    If vUsu.Nivel > 2 Then
        MsgBox "No tiene suficientes privilegios. Consulte al administrador del sistema. ", vbExclamation
    Else
        AbrirListado3 IIf(Index = 1, 61, 3)
    End If
End Sub

Private Sub mnRepAlbaranes_Click()
   ' If vParamAplic.TipoFormularioClientes = 0 Then
        frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
        frmFacEntAlbaranes2.hcoCodTipoM = "ALR"
        frmFacEntAlbaranes2.EsHistorico = False
        frmFacEntAlbaranes2.Show vbModal
    'End If
End Sub

Private Sub mnRepAvisos_Click()
'Avisos de averias de clientes
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        frmMensajes.OpcionMensaje = 26
        frmMensajes.Show vbModal
    Else
        frmRepAvisos.Show vbModal
    End If
End Sub

Private Sub mnRepControlRep_Click()
'Control de Reparaciones (para los Tecnicos)
    frmRepEntReparacionesGR.EntradaEquipo = ""
    frmRepEntReparacionesGR.ControlRep = True
    frmRepEntReparacionesGR.EsHistorico = False
    frmRepEntReparacionesGR.Show vbModal
End Sub

Private Sub mnRepEntReparacion_Click()
'Mantenimiento de Reparaciones
    frmRepEntReparacionesGR.EntradaEquipo = ""
    frmRepEntReparacionesGR.ControlRep = False
    frmRepEntReparacionesGR.EsHistorico = False
    frmRepEntReparacionesGR.Show vbModal
End Sub

Private Sub mnRepFactAlb_Click()
'Facturacion de Albaranes de Reparacion
    If vParamAplic.NumeroInstalacion = vbTaxco Then
        frmListadoPed.codClien = "ALO" 'utilizamos esta vble para pasarle el tipo de movimiento
    Else
        frmListadoPed.codClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
    End If
    AbrirListadoPed (52)
End Sub

Private Sub mnRepGarantprove_Click()
    frmListado2.Opcion = 30
    frmListado2.Show vbModal
End Sub

Private Sub mnRepHistorico_Click()
'Historico de las reparaciones
    frmRepEntReparacionesGR.EntradaEquipo = ""
    frmRepEntReparacionesGR.ControlRep = False
    frmRepEntReparacionesGR.EsHistorico = True
    frmRepEntReparacionesGR.Show vbModal
End Sub


Private Sub mnRepListAvisosPtes_Click()
'Listado de avisos de averias de clientes pendientes
    AbrirListado (409)
End Sub

Private Sub mnRepListFrecuen_Click()
'Listado de Frecuencia de Reparaciones
    AbrirListado (406)
End Sub

Private Sub mnRepListRepxClien_Click()
'Listado de las Reparaciones por cliente
    AbrirListado (64)
End Sub

Private Sub mnRepListRepxDia_Click()
'Listado de las Reparaciones del dia
    AbrirListado (63)
End Sub

Private Sub mnRepMotivosBaja_Click()
'Motivos baja equipos
    frmRepMotivosBaja.Show vbModal
End Sub

Private Sub mnRepMotivosPend_Click()
'Motivos Pendientes Reparar
    frmRepMotivosPend.Show vbModal
End Sub

Private Sub mnRepNumSerie_Click()
'Mantenimiento de N�s de Serie
    frmRepNumSerie2GR.Show vbModal
End Sub

Private Sub mnRepPrevFact_Click()
' Previsi�n Facturacion de Albaranes de Reparacion

    If vParamAplic.NumeroInstalacion = vbTaxco Then
        frmListadoPed.codClien = "ALO" 'utilizamos esta vble para pasarle el tipo de movimiento
    Else
        frmListadoPed.codClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
    End If
    AbrirListadoPed (50) 'NO IMPRIME LISTADO

End Sub

Private Sub mnRevisarMultibase_Click()
    AbrirListado2 3
End Sub

'Private Sub mnPedirPwd_Click()
'Dim Anterior As Boolean
'
'    Anterior = Me.mnPedirPwd.Checked
'    vConfig.PedirPasswd = Not Anterior
'    If vConfig.Grabar = 1 Then
'        Me.mnPedirPwd.Checked = Anterior
'    Else
'        Me.mnPedirPwd.Checked = Not Anterior
'    End If
'End Sub


Private Sub mnSeleccionarImpresora_Click()
    Screen.MousePointer = vbHourglass
    frmPpal.CommonDialog1.ShowPrinter
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnServicios_Click(Index As Integer)
    If Index = 0 Then Exit Sub  'La barra no puede
    Select Case Index
    Case 1, 3
        If vParamAplic.TipoFormularioClientes = 0 Then
            frmFacEntAlbaranes2.hcoCodMovim = "" 'No carga el form con datos al abrir
            If Index = 1 Then
                frmFacEntAlbaranes2.hcoCodTipoM = "ALS"
            Else
                frmFacEntAlbaranes2.hcoCodTipoM = "ALI"
            End If
            frmFacEntAlbaranes2.EsHistorico = False
            frmFacEntAlbaranes2.Show vbModal
            
        Else
            If InstalacionEsEulerTaxco Then
                
                frmFacEntAlbSAIL.hcoCodMovim = "" 'No carga el form con datos al abrir
                If Index = 1 Then
                    frmFacEntAlbSAIL.hcoCodTipoM = "ALS"
                Else
                    frmFacEntAlbSAIL.hcoCodTipoM = "ALI"
                End If
                
                frmFacEntAlbSAIL.EsHistorico = False
                frmFacEntAlbSAIL.Show vbModal
            End If
        End If
    Case 2, 4
        If Index = 2 Then
            frmListadoPed.codClien = "ALS" 'utilizamos esta vble para pasarle el tipo de movimiento
        Else
            frmListadoPed.codClien = "ALI" 'utilizamos esta vble para pasarle el tipo de movimiento
        End If
        AbrirListadoPed (52)
    
    Case 5
        'Lisatado albaranes internos
        frmListado5.OpcionListado = 13
        frmListado5.Show vbModal
        
        
        
    End Select
    
End Sub

Private Sub mnSituaAlba_Click(Index As Integer)
    Select Case Index
    Case 0, 3
        If Index = 0 Then
            frmListado2.Opcion = 23
        Else
            frmListado2.Opcion = 27
        End If
        frmListado2.Show vbModal
    Case 1
     
        frmFacAlbAsignar.Show vbModal
    Case 2
        frmFacFacAsignar.Show vbModal
    Case 4
        
'        AbrirListado3 47     TODO   QUitar el frame para que disminuya el tama�o
        frmListado4.Opcion = 11
        frmListado4.Show vbModal
        
    Case 5
    
        '
        frmListado5.OpcionListado = 24
        frmListado5.Show vbModal
        
        
        If HaMostradoCanal2_El_B Then Me.mnAlbaranesB.visible = True
            
    Case 6
        
        frmAvisosAlb.Show vbModal
    End Select
End Sub

Private Sub mnSociosProveedores_Click(Index As Integer)
    Select Case Index
    Case 0
        'Cambiar precios proveedores /socios
         AbrirListado2 7
         
    Case 1
        'Liquidacion SOCIOS
        AbrirListado2 8
        
    Case 2
        'Impresion facturas proveedores
        AbrirListado2 9
        
    Case 3
        'MsgBox "En desarrollo", vbExclamation
        'Asociar albaranes compras / vetnas
         frmComprasVentas.Show vbModal
    
    Case 4
        'listado trazabilidad
        AbrirListado2 15
    End Select
End Sub

Private Sub mnSoporte_Click(Index As Integer)

       

   
    Select Case Index
    Case 4
       
        Screen.MousePointer = vbHourglass
        LanzaHome ("websoporte")
        Screen.MousePointer = vbDefault

    Case 7
        'Acerca de
        Screen.MousePointer = vbDefault
        frmMensajes.OpcionMensaje = 3
        frmMensajes.Show vbModal
    End Select
    
End Sub

Private Sub mnTelefonia2_Click(Index As Integer)
    

    'ANTES 2013.
    'Ver seccion comenta de abajo
    Select Case Index
    Case 0
            frmFacEntAlbaranes2.hcoCodTipoM = "ALT"
            frmFacEntAlbaranes2.EsHistorico = False
            frmFacEntAlbaranes2.Show vbModal
    
    
    Case 2, 3
    
            If Index = 2 Then
                If vParamAplic.TieneTelefonia2 = 2 Then
    
'                    'importacion coarval
'                    frmTelefono1.Opcion = 5
'                    frmTelefono1.Show vbModal
'                    Exit Sub
                End If
            End If
            
            'Importar y el 4 verdatos
            If Index = 2 Then
                'Importacion
                CadenaDesdeOtroForm = ""
                frmTelefono1.Opcion = 1 'Importar
                frmTelefono1.Show vbModal
                If CadenaDesdeOtroForm = "" Then Exit Sub
                Screen.MousePointer = vbHourglass
                Espera 0.2
                DoEvents
                FuerzaCiereFormTelefonia
            End If
        
            frmTelefono1.Opcion = 0
            frmTelefono1.Show vbModal

            

    Case 7
        frmTelBolbaiteGR.QueOpcion = 3
        frmTelBolbaiteGR.Show vbModal
    Case 8
        frmListado4.Opcion = 10
        frmListado4.Show vbModal
    Case 10, 11, 12, 14
    
        ' 2.- Listado descuentos comprataiivo copera
        ' 3.- Rsumen fracion
        ' 4.- Datos face
        '
        ' 6.-  Datos importados (index=!4)
        If Index = 14 Then
            frmTelefono1.Opcion = 6
        Else
            frmTelefono1.Opcion = Index - 8 '2,3,4
        End If
        frmTelefono1.Show vbModal

    
    End Select

End Sub






Private Sub mnTelefonia3_Click(Index As Integer)
    If Index = 0 Then
            '   bolbaite
            frmTelBolbaiteGR.QueOpcion = 1
            frmTelBolbaiteGR.Show vbModal
    Else
            frmTelDtoConsumo.Show vbModal
    End If
    
End Sub

Private Sub mnTelefonia4_Click(Index As Integer)
    If Index = 0 Then
            'bolbaite
            frmTelBolbaiteGR.QueOpcion = 0
            frmTelBolbaiteGR.Show vbModal
    ElseIf Index = 1 Then
         frmTelDtoCuotas.Show vbModal
    Else
         frmTelBolbaiteGR.QueOpcion = 2
         frmTelBolbaiteGR.Show vbModal
    End If
End Sub

Private Sub mnTelematel_Click(Index As Integer)
    Select Case Index
    Case 0
        frmTelematMto.Show vbModal
    
    End Select
End Sub

Private Sub mnTicket_Click(Index As Integer)
    
    If Index > 0 Then AbrirListado2 12 + Index

    
End Sub

Private Sub mnTiposArticulos_Click()
    frmAlmTipoArticulo.Show vbModal
End Sub

Private Sub mnSalir_Click()
    HacerEND
End Sub





Private Sub mnTiposAveria_Click()
    frmtipave.Show vbModal
End Sub


Private Sub AbirTPVpantallaVenta()
'Pantalla venta del TPV
Dim nom As String

    'Antes de abrir la pantalla de venta comprobamos que podemos leer el terminal
    'nom = ComputerNameTServer

    nom = ComputerName 'Nombre PC conectado por Terminal Server / local
    
    If Trim(nom) <> "" Then
        frmFacTPVEnt.NomrePC_conectado = nom
        frmFacTPVEnt.Show
    Else
'        'Terminal con el que trabajaremos, leemos el nombre del ordenador en local
'        'si no trabajamos en terminal server
'        nom = ComputerName
'        If Trim(nom) <> "" Then
'            frmFacTPVEnt.NomrePC_conectado = nom
'            frmFacTPVEnt.Show
'        Else
            MsgBox "No se puedo establecer un terminal.", vbExclamation
'        End If
    End If
End Sub



Private Sub mnTPVLinea_Click(Index As Integer)
            '
    Select Case Index
    Case 0
        AbirTPVpantallaVenta
    Case 1
        'Cierre caja
        'Abre el informe de cierre de caja del dia en el TPV
        AbrirListadoOfer (240)
    
    Case 2
        'Etiquetas estanteria
        AbrirListado 94
    Case 4
        'Par�metros generales del TPV
        frmFacTPVParamG.Show vbModal
    Case 5
        frmFacTPVParamT.Show vbModal
    End Select
End Sub

Private Sub mnTrabaRealiz_Click()
    frmManTraReali.Show vbModal
End Sub

Private Sub mnTraspasoMante_Click()
    Screen.MousePointer = vbHourglass
    frmMensajes.OpcionMensaje = 18
    frmMensajes.Show vbModal
End Sub





Private Sub mnTratamientos_Click(Index As Integer)
    Select Case Index
    Case 0
        'If Index = 0 Then
        frmAlmMatAct.DatosADevolverBusqueda = ""
        frmAlmMatAct.Show vbModal
    Case 1
        'If Index = 1 Then
        frmAlmADR.DatosADevolverBusqueda = ""
        frmAlmADR.Show vbModal
    Case 2
        'If Index = 2 Then
        frmAlmPlagas.DatosADevolverBusqueda = ""
        frmAlmPlagas.Show vbModal
    Case 3
        frmFlotas.Show vbModal
    Case 4
        frmADVTratamientos.DatosADevolverBusqueda = False
        frmADVTratamientos.Show vbModal
    Case 5
        frmADVTraPartes.Show vbModal
        
    Case 6
        'Fitos por campos
        frmListado5.OpcionListado = 12
        frmListado5.Show vbModal
    Case 7
        'Vacio y NO visible
        
    Case 9
        frmListado5.OpcionListado = 9
        frmListado5.Show vbModal
    End Select
End Sub

Private Sub mnUtiBuscarErrConCli_Click()
'Facturas pendientes de contabilizar (CLIENTES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 6
    frmUtilidades.Show vbModal
End Sub

Private Sub mnUtiBuscarErrConPro_Click()
'Facturas pendientes de contabilizar (PROVEEDORES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 7
    frmUtilidades.Show vbModal
End Sub


Private Sub mnUtiBuscarErrFac_Click()
'Buscar errores en n� de factura (solo en facturas de clientes)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 5
    frmUtilidades.Show vbModal
End Sub



Private Sub mnUtiConnActivas_Click()
'ver las conexiones a donde apuntan
Dim Cad As String
'    cad = "Conexiones:" & vbCrLf
'    cad = cad & "------------------" & vbCrLf & vbCrLf
'    cad = cad & "Ariges: " & vbCrLf & conn.ConnectionString & vbCrLf & vbCrLf
'    cad = cad & "Conta: " & vbCrLf & ConnConta.ConnectionString & vbCrLf
'    MsgBox cad, vbInformation
    
    
    MostrarCadenasConexion
End Sub



Private Sub mnUtiDeclaraLOM_Click(Index As Integer)
    If Index = 0 Then
        frmFacLotesGeneralitat.Show vbModal
    Else
        frmUtDeclara.Show vbModal
    End If
End Sub

Private Sub mnUtilidadesVarias_Click(Index As Integer)
    Select Case Index
    Case 0
        AbrirListado3 16
    Case 1
        'Comprobar cuenta banco secciones(y contabilidades)
        
        frmListadoOfer.OpcionListado = 408
        frmListadoOfer.Show vbModal
        
    Case 2, 3
        If vUsu.Nivel > 1 Then
            MsgBox "No tienen permiso para realizar esta accion", vbExclamation
        Else
            If Index = 2 Then
                frmListado3.Opcion = 45
                frmListado3.Show vbModal
            Else
                frmListado5.OpcionListado = 2
                frmListado5.Show vbModal
            End If
        End If
        
    Case 4
        frmEulerPrecios.Show vbModal
    
    Case 5
        frmListado3.Opcion = 59
        frmListado3.Show vbModal
    Case 6
        frmListado3.Opcion = 66
        frmListado3.Show vbModal
        
    Case 7
        frmListado5.OpcionListado = 25
        frmListado5.Show vbModal
    
    Case 8
         frmListado5.OpcionListado = 26
        frmListado5.OtrosDatos = ""
        frmListado5.Show vbModal
    
    Case 9
        frmListado5.OpcionListado = 35
        frmListado5.OtrosDatos = ""
        frmListado5.Show vbModal
    
    
    End Select
End Sub







Private Sub mnUtiMensLin_Click(Index As Integer)
    'Nuevo mensaje en la utilidad de mensajeria interna
    Select Case Index
    Case 0
        frmMensaje2.Show vbModal
    Case 1
    
    
    Case 3
         frmTiposMensajes.Show vbModal
    End Select
    
End Sub

Private Sub mnUtiUsuActivos_Click()
'Muestra si hay otros usuarios conectados a la Gestion
Dim SQL As String
Dim i As Integer

    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    If CadenaDesdeOtroForm <> "" Then
        i = 1
        Me.Tag = "Los siguientes PC's est�n conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            SQL = RecuperaValor(CadenaDesdeOtroForm, i)
            If SQL <> "" Then Me.Tag = Me.Tag & "    - " & SQL & vbCrLf
            i = i + 1
        Loop Until SQL = ""
        MsgBox Me.Tag, vbExclamation
    Else
        MsgBox "Ningun usuario, adem�s de usted, conectado a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation
    End If
    CadenaDesdeOtroForm = ""
End Sub





Private Sub mnVerAvisos_Click()
    If TieneAvisosPendientes Then
        frmAlertasGR.Show vbModal
    Else
        MsgBox "No hay avisos para mostrar", vbInformation
    End If
End Sub







Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

   
    
        
    Select Case Button.Index
    Case 1 'Mantenimiento de Art�culos
        mnAlmArticulos_Click
    Case 2 'Movimientos Articulos
        mnAlmMovimArticulos_Click
        
    Case 5 'Mantenimiento Clientes
        mnFacClientes_Click
    Case 6 'Mantenimiento Proveedores
        mnComProveedores_Click
        
    Case 9 'Ofertas a Clientes
        mnFacOfertas_Click 0
    Case 10 'Pedidos a Clientes
        'mnFacEntPedidos_Click
        mnFacPedidos_Click 0
    Case 11 'Albaranes a Clientes
        mnFacEntAlbaran_Click
    
    Case 12 'Hist. Albaranes (Facturas)
        mnFacHcoFacturas_Click
        
    Case 13
        'Facturas mostrador
        mnFacAlbMostrador_Click
        
    Case 15 'Pedidos de Proveedores
        mnComPedidosLin_Click 0
    Case 16 'Albaranes de Proveedores
        mnComAlbMan_Click
    Case 17 'Facturas de Proveedores
        mnComHcoFacturas_Click
    Case 18 'Recepcion Fact. Prove
        If Me.mnComFacturar.visible And Me.mnComFacturar.Enabled Then mnComFacturar_Click
        
    Case 21 'Mantenimientos
    
        If vParamAplic.NumeroInstalacion = vbFenollar Then
            frmFacPedidoAgrupados.Cliente = -1
            frmFacPedidoAgrupados.Show vbModal
            
             If vParamAplic.NumeroInstalacion = vbFenollar Then
                If HaMostradoCanal2_El_B Then mnAlbaranesB.visible = True
            End If
            
        Else
            mnManEntrada_Click
        End If
    Case 22 'N� Serie
    
        If vParamAplic.NumeroInstalacion = vbFenollar Then
            'FENOLLAR
            
          
    
            
            frmfacFacturacion.Show vbModal
                        
            If vParamAplic.NumeroInstalacion = vbFenollar Then
                If HaMostradoCanal2_El_B Then mnAlbaranesB.visible = True
            End If
        ElseIf vParamAplic.NumeroInstalacion = vbTaxco Then
            mnMtoEuler_Click 1
        Else
        
            'Numero de serie
            mnRepNumSerie_Click
        End If
    Case 23
        If vParamAplic.Ariagro = "" Then
        
        
        
            mnRepAvisos_Click
        Else
            mnTratamientos_Click 5
        End If
    Case 24 'Gastos T�cnicos
    
        If InstalacionEsEulerTaxco Then
            mnObra_Click 4
        Else
            mnAdmGastosTec_Click
        End If
    Case 25
        'Consulta precio articulo
        mnFacPedidos_Click 9
        
    Case 26 'Entrada al TPV
        AbirTPVpantallaVenta
    Case 27
        'cambiar empresa
        mnCambioEmpresa_Click
        
    Case 28
        If vParamAplic.Frecuencias Then
            'Llamamos a frecuencias
            mnFrecuencias_Click
        Else
        
            If vParamAplic.NumeroInstalacion = vbFenollar Then
                frmfacAlbaraYFac.Show vbModal
            Else
                mnAgenda_Click
            End If
        End If
        
    Case 30 'Salir
        mnSalir_Click
    End Select
End Sub


Private Sub PonerDatosVisiblesForm()
'Escribe texto de la barra de la aplicaci�n
Dim Cad As String
    Cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    Cad = Cad & ", " & Format(Now, "d")
    Cad = Cad & " de " & Format(Now, "mmmm")
    Cad = Cad & " de " & Format(Now, "yyyy")
    Cad = "    " & Cad & "    "
    Me.StatusBar1.Panels(5).Text = Cad
    If vEmpresa Is Nothing Then
        Caption = "ARIGES" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & "   Usuario: " & vUsu.Nombre & " FALTA CONFIGURAR"
        'Panel con el nombre de la empresa
        Me.StatusBar1.Panels(2).Text = "Falta configurar"
    Else
        Caption = "ARIGES" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vEmpresa.nomempre & "  -    Usuario: " & vUsu.Nombre
        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               C�digo: " & vEmpresa.codempre
    End If
End Sub


Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim Cad As String

    
    For Each T In Me
        Cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            If LCase(Mid(T.Caption, 1, 1)) <> "-" Then T.Enabled = Habilitar
        End If
    Next
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnConfParamAplic = True
    Me.mnConfParamGenerales = True

    Me.mnSalir.Enabled = True
    Me.mnCambioEmpresa.Enabled = True
End Sub

'-------------------------------------
'Pondremos todos los que menus a visibles. Y luego ya , en f
Private Sub ReestablecerMenus()
Dim T
      For Each T In Me
        If Mid(T.Name, 1, 2) = "mn" Then T.visible = True
    Next
End Sub

Private Sub PonerMenusNivelUsuario()
Dim B As Boolean

    B = (vUsu.Nivel = 0) Or (vUsu.Nivel = 1)  'Administradores y root

    Me.mnConfParamAplic = B
    mnConfManteUsuarios = B
    
    mnUsuarios.Enabled = B
   ' mnNuevaEmpresa.Enabled = b
   ' mnPedirPwd.Enabled = b
    Me.mnUtiConnActivas.Enabled = (vUsu.Nivel = 0) 'solo para root
    

    B = vUsu.Nivel = 3  'Es usuario de consultas
    If B Then
        'Inventario
        Me.mnAlmTomaInven.Enabled = False
        Me.mnAlmEntradaInve.Enabled = False
        Me.mnAlmActualizarInve.Enabled = False
        Me.mnAlmListadoInve.Enabled = False
        Me.mnAlmValoracionInve.Enabled = False
        'Antes
        'Me.mnFacTrasHist.Enabled = False
        mnFacOfertas(6).Enabled = False
        
        
        'Facturacion Ventas
        Me.mnFacFacturarAlb.Enabled = False
        Me.mnFacContFactu.Enabled = False
        
        'Facturacion Compras
        Me.mnComFacturar.Enabled = False
        Me.mnComContFactu.Enabled = False
        
        'Reparaciones
        Me.mnRepFactAlb.Enabled = False
        
        'Mantenimientos y renting
        'Me.mnManFactAlb.Enabled = False
        mnManPrevFac2(1).Enabled = False
        mnManPrevFac2(4).Enabled = False
    End If
End Sub



Private Sub LanzaHome(Opcion As String)
Dim i As Integer
Dim Cad As String

    On Error GoTo ELanzaHome

'    LanzaHome = False
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "spara1", Opcion, "codigo", "1", "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en Par�metros de la Aplicaci�n.", vbExclamation
        Exit Sub
    End If

    If Opcion = "webversion" Then
        CadenaDesdeOtroForm = "http://" & CadenaDesdeOtroForm & "?version=" & App.Major & "." & App.Minor & "." & App.Revision
        
        If LanzaHomeGnral(CadenaDesdeOtroForm) Then Espera 2

    Else
        'Lanzamos
        
        CadenaDesdeOtroForm = "http://" & CadenaDesdeOtroForm
        LanzaVisorMimeDocumento Me.hwnd, CadenaDesdeOtroForm
    End If
'    If cad <> "" Then Shell cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
'    If vConfig.Explorador <> "" Then
'        Shell vConfig.Explorador & " " & CadenaDesdeOtroForm, vbMaximizedFocus
'        LanzaHome = True
'    End If
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, Cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub



Private Sub LeerEditorMenus()
Dim SQL As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    SQL = "Select count(*) from usuarios.appmenus where aplicacion='Ariges'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim SQL As String
Dim C As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    
    SQL = "Select * from usuarios.appmenususuario where aplicacion='Ariges' and codusu = " & Val(Right(CStr(vUsu.Codigo), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            SQL = SQL & miRsAux.Fields(3)
            If Right(miRsAux.Fields(3), 1) <> "|" Then SQL = SQL & "|"
            SQL = SQL & "�"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If SQL <> "" Then
        SQL = "�" & SQL
        For Each T In Me.Controls
            If TypeOf T Is Menu Then
                C = DevuelveCadenaMenu(T)
                C = "�" & C & "�"
                'Debug.Print C
                If InStr(1, SQL, C) > 0 Then
                    
                    '
                    T.visible = False
                End If
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index & "|"
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function



Private Sub PoneBarraMenus2()
'Para cada boton de la toolbar comprobar que el menu con el que se corresponde
'esta visible y activado, y ponerle los mismos valore que tenga el menu
Dim Activado As Boolean

    On Error GoTo 0
    
    '-----------------------------------------------------------
    'Articulos
    Me.Toolbar1.Buttons(1).visible = ComprobarBotonMenuVisible(Me.mnAlmArticulos, Activado)
    Me.Toolbar1.Buttons(1).Enabled = Activado

    'Movimientos de Articulos
    Me.Toolbar1.Buttons(2).visible = ComprobarBotonMenuVisible(Me.mnAlmMovimArticulos, Activado)
    Me.Toolbar1.Buttons(2).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Clientes
    Me.Toolbar1.Buttons(5).visible = ComprobarBotonMenuVisible(Me.mnFacClientes, Activado)
    Me.Toolbar1.Buttons(5).Enabled = Activado
    
    'Proveedores
    Me.Toolbar1.Buttons(6).visible = ComprobarBotonMenuVisible(Me.mnComProveedores, Activado)
    Me.Toolbar1.Buttons(6).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Ofertas Clientes
    Me.Toolbar1.Buttons(9).visible = ComprobarBotonMenuVisible(Me.mnFacOfertas(0), Activado)
    Me.Toolbar1.Buttons(9).Enabled = Activado
    
    'Pedidos Clientes
    Me.Toolbar1.Buttons(10).visible = ComprobarBotonMenuVisible(mnFacPedidos(0), Activado)
    Me.Toolbar1.Buttons(10).Enabled = Activado
    
    'Albaranes Clientes
    Me.Toolbar1.Buttons(11).visible = ComprobarBotonMenuVisible(Me.mnFacEntAlbaran, Activado)
    Me.Toolbar1.Buttons(11).Enabled = Activado
    
    'Facturas Clientes
    Me.Toolbar1.Buttons(12).visible = ComprobarBotonMenuVisible(Me.mnFacHcoFacturas, Activado)
    Me.Toolbar1.Buttons(12).Enabled = Activado
    
    
    
    'Si esta visible entonces SI lleva la misma serie no la muestro
    If vParamAplic.FrasMostradorSerieDistinta Then
        Me.Toolbar1.Buttons(13).visible = ComprobarBotonMenuVisible(mnFacAlbMostrador, Activado)
        Me.Toolbar1.Buttons(13).Enabled = Activado
    Else
        Me.Toolbar1.Buttons(13).visible = False
    End If
    
    '-----------------------------------------------------------
    'Pedidos Proveedor
    'Comprobar que los menus del que cuelga no este bloqueado o invisible
    Me.Toolbar1.Buttons(15).visible = ComprobarBotonMenuVisible(mnComPedidosLin(0), Activado)
    Me.Toolbar1.Buttons(15).Enabled = Activado
    
    'Albaranes Proveedor
    Me.Toolbar1.Buttons(16).visible = ComprobarBotonMenuVisible(Me.mnComAlbMan, Activado)
    Me.Toolbar1.Buttons(16).Enabled = Activado
    
    'Facturas Proveedor
    Me.Toolbar1.Buttons(17).visible = ComprobarBotonMenuVisible(Me.mnComHcoFacturas, Activado)
    Me.Toolbar1.Buttons(17).Enabled = Activado
    
    'Recepcion facturas de compras
    Me.Toolbar1.Buttons(18).visible = ComprobarBotonMenuVisible(Me.mnComFacturar, Activado)
    Me.Toolbar1.Buttons(18).Enabled = Activado


    '-----------------------------------------------------------
    'Mantenimientos
    Me.Toolbar1.Buttons(21).visible = ComprobarBotonMenuVisible(Me.mnManEntrada, Activado)
    Me.Toolbar1.Buttons(21).Enabled = Activado
    If vParamAplic.NumeroInstalacion = vbFenollar Then Me.Toolbar1.Buttons(21).ToolTipText = "Agrupaci�n de pedidos"
    'N� Serie
    Me.Toolbar1.Buttons(22).visible = ComprobarBotonMenuVisible(Me.mnRepNumSerie, Activado)
    Me.Toolbar1.Buttons(22).Enabled = Activado
    If vParamAplic.NumeroInstalacion = vbFenollar Then Me.Toolbar1.Buttons(22).ToolTipText = "FACTURACION"
    If vParamAplic.NumeroInstalacion = vbTaxco Then Me.Toolbar1.Buttons(22).ToolTipText = "ORDEN TALLER"
    
    '-----------------------------------------------------------
    'Conuslta de precio
    Me.Toolbar1.Buttons(24).visible = ComprobarBotonMenuVisible(Me.mnFacPedidos(8), Activado)
    Me.Toolbar1.Buttons(24).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Gastos tecnicos
    'Para EULER --> Reloj
    mnobra(2).Caption = "Partes de trabajo"
    If InstalacionEsEulerTaxco Then
        Me.Toolbar1.Buttons(24).visible = ComprobarBotonMenuVisible(Me.mnobra(4), Activado)
        Me.Toolbar1.Buttons(24).Enabled = Activado
        mnobra(2).Caption = "Mantenimiento tareas reloj"
    Else
        Me.Toolbar1.Buttons(24).visible = ComprobarBotonMenuVisible(Me.mnAdmGastosTec, Activado)
        Me.Toolbar1.Buttons(24).Enabled = Activado
    End If
    
    'Nuevos botones
    'TPV
    Me.Toolbar1.Buttons(26).visible = ComprobarBotonMenuVisible(mnTPVLinea(0), Activado)
    If Activado Then
        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "count(*)", "spatpvg", "1", "1")
        If CadenaDesdeOtroForm = "" Then CadenaDesdeOtroForm = "0"
        If Val(CadenaDesdeOtroForm) = 0 Then Activado = False
    End If
    Me.Toolbar1.Buttons(26).Enabled = Activado
    
    'Cambio empresa
   ' Me.Toolbar1.Buttons(27).visible = ComprobarBotonMenuVisible(mnCambioEmpresa, Activado)
    'Cambiar empresa lo dejo desde Febrero 2013 SIEMPRE visib�e
    Me.Toolbar1.Buttons(27).Enabled = True
    Me.Toolbar1.Buttons(27).visible = True
    
    'Agenda
    If vParamAplic.Frecuencias Then
        Me.Toolbar1.Buttons(28).Image = 24 'FRECUENCIAS
        Me.Toolbar1.Buttons(28).visible = ComprobarBotonMenuVisible(mnFrecuencias, Activado)
        Me.Toolbar1.Buttons(28).Enabled = Activado
        Me.Toolbar1.Buttons(28).ToolTipText = "Frecuencias"
    Else
        Me.Toolbar1.Buttons(28).Image = 20 'Agenda
        Me.Toolbar1.Buttons(28).visible = ComprobarBotonMenuVisible(mnAgenda, Activado)
        Me.Toolbar1.Buttons(28).Enabled = Activado
        Me.Toolbar1.Buttons(28).ToolTipText = "Agenda"
    End If
    
    'Avisos
    
    If vParamAplic.Ariagro = "" Then
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            Me.Toolbar1.Buttons(23).visible = ComprobarBotonMenuVisible(mnMtoEuler(1), Activado)
            QueCaption = "Hist�rico reparaciones"
        Else
            Me.Toolbar1.Buttons(23).visible = ComprobarBotonMenuVisible(mnRepAvisos, Activado)
            QueCaption = "Aviso reparacion"
        End If
        Me.Toolbar1.Buttons(23).Enabled = Activado
    Else
        'partes de trabajo
        Me.Toolbar1.Buttons(23).visible = ComprobarBotonMenuVisible(mnTratamientos(4), Activado)
        Me.Toolbar1.Buttons(23).Enabled = Activado
        QueCaption = "Partes de trabajo"
    End If
    Toolbar1.Buttons(23).ToolTipText = QueCaption
    QueCaption = ""
End Sub




Private Function ComprobarBotonMenuVisible(objMenu As Menu, Activado As Boolean) As Boolean
'Comprueba si el icono de la barra se debe activar/desactivar o si se debe poner
'visible o invisible. Para ello comprueba si su correspondiente entrada de menu
'esta activada/desactiva o visible/invisible
'(se comprueba hasta q se encuentra el false o se llega al padre)
Dim nomMenu As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Cad As String
Dim B As Boolean


    On Error GoTo EComprobar
    
    B = objMenu.visible
    Activado = objMenu.Enabled
    
    If B = False Then
        ComprobarBotonMenuVisible = B
    Else
    
        nomMenu = objMenu.Name
        
        Set RS = New ADODB.Recordset
        
        'Obtener el padre del menu
        SQL = "select padre from usuarios.appmenus where aplicacion='Ariges' and name=" & DBSet(nomMenu, "T")
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            Cad = RS.Fields(0).Value
        End If
        RS.Close
        
        B = True
        While B And Cad <> ""
                SQL = "Select name,padre from usuarios.appmenus where aplicacion='Ariges' and contador= " & Cad
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    Cad = RS!Padre
                    nomMenu = RS!Name
                End If
                RS.Close
                
                'comprobar si el padre esta bloqueado
                SQL = "Select count(*) from usuarios.appmenususuario where aplicacion='Ariges' and codusu=" & Val(Right(CStr(vUsu.Codigo), 3))
                SQL = SQL & " and tag='" & nomMenu & "|'"
                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RS.Fields(0).Value > 0 Then
                    'Esta bloqueado el menu para el usuario
                    B = False
                    Activado = False
                End If
                RS.Close
                If Cad = "0" Then Cad = "" 'terminar si llegamos a la raiz
        Wend
        ComprobarBotonMenuVisible = B
        Set RS = Nothing
    End If
    
EComprobar:
    If Err.Number <> 0 Then Err.Clear
End Function



Private Sub AbrirListado2(KOpcion As Integer)
    Screen.MousePointer = vbHourglass
    frmListado2.Opcion = KOpcion
    frmListado2.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub AbrirListado3(KOpcion As Integer)
    Screen.MousePointer = vbHourglass
    frmListado3.Opcion = KOpcion
    frmListado3.Show vbModal
    Screen.MousePointer = vbDefault
End Sub























Private Sub PonerCaption(Reestablecer As Boolean, Optional texto As String)

    If vParamAplic.NumeroInstalacion <> vbFenollar Then Exit Sub

    If Not Reestablecer Then
        QueCaption = Me.Caption
        Me.Caption = "Ariges." & Replace(texto, "&", "") & " - " & vEmpresa.nomresum
    Else
        Me.Caption = QueCaption
    End If
End Sub




Private Sub FuerzaCiereFormTelefonia()
    On Error Resume Next
    Unload frmTelefono1
    If Err.Number <> 0 Then Err.Clear
End Sub






