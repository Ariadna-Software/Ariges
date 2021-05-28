VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TelGenerador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mfichero3 As String

Private mdbFac2 As TelBaseDatos


Private mFicheroCorto As String ' lleva el nombre estricto de fichero sin PATH
'variables locales para almacenar los valores de las propiedades
Private mvarFechaFactura As Date 'copia local
Private mvarFichero As String 'copia local



'Seran los conceptos de trafico refacturables
Private mvarConceptosLlamadasRefacturables As String
Private mvarConceptosQueSonLlamadas As String 'de los refacturables, cuales son llamadas. Para cuando hace los sum
Private mvarConceptosNOSonLlamadas As String 'de los refacturables, cuales son llamadas. Para cuando hace los sum
Private mvarPrecioEstablecimiento As Currency
Private mvarPrecioMinuto As Currency

'Entra catadau
Private mvarDigitoCoarval As Integer
Private mvarCuotaMinimaSobreConsumos As Boolean
Private mvarDescuentosPorConsumo As Boolean


'Para el proceso VODAFONE
Dim Valores() As String
Dim CadenaInsert  As String
Dim NF As Integer



'Para ver si va pr telefono o agrupado
Dim EsAgrupacionTelefono As Boolean
Dim whereTelefono As String

Public Property Let Fichero(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Fichero = 5
    mvarFichero = vData
End Property


Public Property Get Fichero() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Fichero
    Fichero = mvarFichero
End Property



Public Property Let FechaFactura(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FechaFactura = 5
    mvarFechaFactura = vData
End Property


Public Property Get FechaFactura() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FechaFactura
    FechaFactura = mvarFechaFactura
End Property



Public Function cargarBaseDatosMOVISTAR(mfichero As String, ByRef LB As Label) As Boolean
    Dim NF As Integer ' soporta el identificador para apertura de fichero
    Dim leido As String ' registro leido del fichero
    Dim i As Integer ' un contadoir de registros
    
    
    
    
    '-- En función de los datos contenidos en el fichero de
    '   carga de telefonía actualiza la base de datos correspondiente
    '-- 1º comprobamos que existe el fichero a importar
    If mfichero = "" Or Dir(mfichero) = "" Then
        '-- solicitamos por pantalla que nos den el fichero de marras
    End If
    mFicheroCorto = Right(mfichero, 12)
    '-- el fichero existe y vamos a procesar sus datos, lo abrimos
    NF = FreeFile()
    Open mfichero For Input As #NF
    '- Borramos registros del mismo fichero si los hubiese
    borraDatos mFicheroCorto
    
    cargarBaseDatosMOVISTAR = True
    
    '-- El fichero por lo menos existe y grabamos sus datos
    cargaDatosPrincipales mFicheroCorto
    leido = ""
    Do While Not EOF(NF)
        Line Input #NF, leido
        i = i + 1
        LB.Caption = "Registro " & CStr(i)
        LB.Refresh
        If (i Mod 5) = 0 Then DoEvents
        
        '-- hacemos lo que tengamos que hacer con el registro
        Select Case Left(leido, 3)
            Case "010" ' cabecera soporte
                cargaRegistro010 leido
            Case "100" ' general factura
                cargaRegistro100 leido
            Case "300" ' general telefono
                cargaRegistro300 leido
            Case "310" ' cuotas teléfono
                cargaRegistro310 leido
            Case "320" ' resumen llamadas
                cargaRegistro320 leido
            Case "330" ' detalle llamadas
                cargaRegistro330 leido
            Case "340" ' varios teléfono
                cargaRegistro340 leido
            Case "800"
                'NANCY. Este tipo, NO lo procesamos. Directamente lo damos por leido
                'MsgBox ""
            Case "850" ' descuentos fectura
                cargaRegistro850 leido
            Case "990" ' totales soporte
                cargaRegistro990 leido
            Case Else ' desconocido
                '-- Registro no contemplado
                MsgBox "El registro TIPO: " & Left(leido, 3) & " no está contemplado en la aplicación.", vbCritical
                cargarBaseDatosMOVISTAR = False
        End Select
    Loop
    Close #NF
End Function

Private Sub cargaDatosPrincipales(CADENA As String)
    
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.ficheros values("
        mSQL = mSQL & .texto(CADENA) & ","  ' campo Fichero
        mSQL = mSQL & .texto(Mid(CADENA, 1, 1)) & ","   ' campo Servicio
        mSQL = mSQL & .texto(Mid(CADENA, 2, 1)) & "," ' campo Mes
        mSQL = mSQL & .texto(Mid(CADENA, 3, 1)) & ","  ' campo Ano
        mSQL = mSQL & .texto(Mid(CADENA, 4, 5)) & ","  ' campo Soporte
        mSQL = mSQL & .texto(Mid(CADENA, 10, 1)) & ","  ' campo Agrupacion
        mSQL = mSQL & .texto(Mid(CADENA, 11, 2)) & ")"  ' campo Version
        .ejecutar mSQL
    End With
End Sub


Private Sub cargaRegistro010(CADENA As String) ' cabecera de soporte
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.cabecera_de_soporte values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registrp
        mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 56, 11)) & ","   ' campo CIF
        mSQL = mSQL & .texto(Mid(CADENA, 67, 65)) & ","   ' campo Nombre
        mSQL = mSQL & .texto(Mid(CADENA, 132, 40)) & ","   ' campo Domicilio
        mSQL = mSQL & .numero(Mid(CADENA, 172, 5)) & ","   ' campo Código Postal
        mSQL = mSQL & .texto(Mid(CADENA, 177, 40)) & ","   ' campo Localidad
        mSQL = mSQL & .texto(Mid(CADENA, 217, 11)) & ","   ' campo CIF
        mSQL = mSQL & .texto(Mid(CADENA, 228, 20)) & ","   ' campo Nombre
        mSQL = mSQL & .texto(Mid(CADENA, 248, 65)) & ","   ' campo Nombre
        mSQL = mSQL & .texto(Mid(CADENA, 313, 40)) & ","   ' campo Domicilio
        mSQL = mSQL & .texto(Mid(CADENA, 353, 5)) & ","   ' campo Código Postal
        mSQL = mSQL & .texto(Mid(CADENA, 358, 40)) & ","   ' campo Localidad
        mSQL = mSQL & .texto(Mid(CADENA, 398, 20)) & ","   ' campo E-mail
        mSQL = mSQL & .texto(Mid(CADENA, 418, 5)) & ","   ' campo Código de soporte
        mSQL = mSQL & .Fecha(fechatel(Mid(CADENA, 423, 8))) & ","   ' campo Fecha de Facturación
        mSQL = mSQL & .Fecha(fechatel(Mid(CADENA, 431, 8))) & ","   ' campo Fecha de Generación registro
        mSQL = mSQL & .texto(Mid(CADENA, 439, 6)) & ","   ' campo Versión del soporte
        mSQL = mSQL & .texto(Mid(CADENA, 445, 12)) & ","   ' campo Nombre del fichero
        mSQL = mSQL & .texto(Mid(CADENA, 457, 6)) & ","   ' campo Código de agrupación asociación
        mSQL = mSQL & .texto(Mid(CADENA, 463, 3)) & ","   ' campo Moneda
        mSQL = mSQL & .numero(Mid(CADENA, 466, 8), 4) & ","  ' campo Factor de conversión
        mSQL = mSQL & .texto(Mid(CADENA, 474, 20)) & ","   ' campo Tipo Soporte
        mSQL = mSQL & .texto(Mid(CADENA, 494, 207)) & ")"  ' campo Libre
        .ejecutar mSQL
    End With
End Sub


Private Sub cargaRegistro100(CADENA As String) ' Factura
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.factura values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo Libre
        mSQL = mSQL & .Fecha(fechatel(Mid(CADENA, 56, 8))) & ","   ' campo Fecha de Factura
        mSQL = mSQL & .texto(Mid(CADENA, 64, 20)) & ","   ' campo Tipo de Factura
        mSQL = mSQL & .numero(Mid(CADENA, 84, 4)) & ","  ' campo Banco
        mSQL = mSQL & .numero(Mid(CADENA, 88, 4)) & ","   ' campo Sucursal
        mSQL = mSQL & .texto(Mid(CADENA, 92, 2)) & ","   ' campo Dígitos de control
        mSQL = mSQL & .numero(Mid(CADENA, 94, 19)) & ","   ' campo Cuenta / Nº de Tarjeta
        mSQL = mSQL & .texto(Mid(CADENA, 113, 8)) & ","   ' campo Servicio
        mSQL = mSQL & .texto(Mid(CADENA, 121, 40)) & ","   ' campo Tipo de Contrato
        mSQL = mSQL & .numero(Mid(CADENA, 161, 5)) & ","   ' campo Numero de líneas fijas
        mSQL = mSQL & .numero(Mid(CADENA, 166, 5)) & ","   ' campo Número de líneas móviles
        mSQL = mSQL & .texto(Mid(CADENA, 171, 3)) & ","   ' campo Moneda Original
        mSQL = mSQL & .numero(Mid(CADENA, 174, 16), 4) & ","  ' campo Total cuotas
        mSQL = mSQL & .numero(Mid(CADENA, 190, 16), 4) & ","  ' campo Total Servicio Medido
        mSQL = mSQL & .numero(Mid(CADENA, 206, 16), 4) & ","  ' campo Total Varios
        mSQL = mSQL & .numero(Mid(CADENA, 222, 16), 4) & ","  ' campo Total Descuentos
        mSQL = mSQL & .numero(Mid(CADENA, 238, 16), 4) & ","  ' campo Suma Total
        mSQL = mSQL & .numero(Mid(CADENA, 254, 16), 4) & ","  ' campo Total Impuestos
        mSQL = mSQL & .numero(Mid(CADENA, 270, 14), 2) & ","  ' campo Total a Pagar
        mSQL = mSQL & .texto(Mid(CADENA, 284, 6)) & ","   ' campo Impuesto
        mSQL = mSQL & .numero(Mid(CADENA, 290, 4), 2) & ","  ' campo Tipo
        mSQL = mSQL & .numero(Mid(CADENA, 294, 16), 4) & ","  ' campo Base Imponible
        mSQL = mSQL & .texto(Mid(CADENA, 310, 11)) & ","   ' campo CIF
        mSQL = mSQL & .texto(Mid(CADENA, 321, 380)) & ")"   ' campo Libre
        .ejecutar mSQL
    End With
End Sub


Private Sub cargaRegistro300(CADENA As String) ' Teléfono
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.telefono values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo N. de PreFactura
        mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo N. de telefono
        mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo N. de extensión
        mSQL = mSQL & .numero(Mid(CADENA, 56, 14), 4) & ","  ' campo Total Cuotas
        mSQL = mSQL & .numero(Mid(CADENA, 70, 14), 4) & ","  ' campo Total Servicio Medido
        mSQL = mSQL & .numero(Mid(CADENA, 84, 14), 4) & "," ' campo Total Servicio Franquiciado
        mSQL = mSQL & .numero(Mid(CADENA, 98, 14), 4) & ","  ' campo Total Varios
        mSQL = mSQL & .numero(Mid(CADENA, 112, 14), 4) & ","  ' campo Total Descuentos
        mSQL = mSQL & .numero(Mid(CADENA, 126, 14), 4) & ","  ' campo Suma Total
        mSQL = mSQL & .numero(Mid(CADENA, 140, 5), 2) & ","  ' campo Proporción de cargos
        mSQL = mSQL & .numero(Mid(CADENA, 145, 14), 4) & ","  ' campo Cargos
        mSQL = mSQL & .numero(Mid(CADENA, 159, 5), 2) & ","  ' campo Proporción Desc. Cuotas
        mSQL = mSQL & .numero(Mid(CADENA, 164, 14), 4) & ","  ' campo Descuento cuotas
        mSQL = mSQL & .numero(Mid(CADENA, 178, 5), 2) & ","  ' campo Proporcion Dto. Serv. Med
        mSQL = mSQL & .numero(Mid(CADENA, 183, 14), 4) & ","  ' campo Descuento Servicio Medido
        mSQL = mSQL & .texto(Mid(CADENA, 197, 10)) & ","   ' campo Código centro de coste
        mSQL = mSQL & .texto(Mid(CADENA, 207, 40)) & ","   ' campo Descripción centro de coste
        mSQL = mSQL & .texto(Mid(CADENA, 247, 40)) & ","   ' campo Usuario
        mSQL = mSQL & .texto(Mid(CADENA, 287, 414)) & ")"   ' campo Libre
        .ejecutar mSQL
    End With
End Sub

Private Sub cargaRegistro310(CADENA As String) ' Cuotas
    Dim i As Integer
    Dim mSQL As String
    For i = 0 To 5
    With mdbFac2
        mSQL = "insert into telefono.cuotas values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo N. de PreFactura
        mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo N. de telefono
        mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo N. de extensión
        If Mid(CADENA, 56 + (93 * i), 8) <> "00000000" Then
            '-- zona que se repite en el registro original ---------------------------------------
            mSQL = mSQL & .Fecha(fechatel(Mid(CADENA, 56 + (93 * i), 8))) & "," ' campo Fecha Inicio Cuotas
            mSQL = mSQL & .Fecha(fechatel(Mid(CADENA, 64 + (93 * i), 8))) & "," ' campo Fecha Fin Cuotas
            mSQL = mSQL & .texto(Mid(CADENA, 72 + (93 * i), 3)) & "," ' campo Código de cuota
            mSQL = mSQL & .texto(Mid(CADENA, 75 + (93 * i), 60)) & "," ' campo Descripción de la cuota
            mSQL = mSQL & .numero(Mid(CADENA, 135 + (93 * i), 14), 4) & "," ' campo Importe
            '-------------------------------------------------------------------------------------
            mSQL = mSQL & .texto(Mid(CADENA, 614, 97)) & ")"   ' campo Libre
            .ejecutar mSQL
        End If
    End With
    Next i
End Sub

Private Sub cargaRegistro320(CADENA As String) ' resumen de llamadas
    Dim i As Integer
    Dim mSQL As String
    
    
    
    For i = 0 To 5
    With mdbFac2
        mSQL = "insert into telefono.resumen_de_llamadas values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo N. de PreFactura
        mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo N. de telefono
        mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo N. de extensión
        mSQL = mSQL & .Fecha(fechatel(Mid(CADENA, 56, 8))) & ","  ' campo Fecha Inicio Periodo
        mSQL = mSQL & .Fecha(fechatel(Mid(CADENA, 64, 8))) & ","  ' campo Fecha Fin Periodo
        If Mid(CADENA, 72 + (103 * i), 3) <> "   " Then
            '-- zona que se repite en el registro original ---------------------------------------
            
            mSQL = mSQL & .texto(Mid(CADENA, 72 + (103 * i), 3)) & "," ' campo Código de tipo de tráfico
            mSQL = mSQL & .texto(Mid(CADENA, 75 + (103 * i), 35)) & "," ' campo Tipo de tráfico
            mSQL = mSQL & .numero(Mid(CADENA, 110 + (103 * i), 7)) & "," ' campo Número de llamadas
            mSQL = mSQL & .texto(Mid(CADENA, 117 + (103 * i), 10)) & "," ' campo Unidad de medida
            mSQL = mSQL & .numero(Mid(CADENA, 127 + (103 * i), 10), 2) & "," ' campo Cantidad de medida
            mSQL = mSQL & .numero(Mid(CADENA, 137 + (103 * i), 10), 2) & "," ' campo Cantidad franquiciada
            mSQL = mSQL & .numero(Mid(CADENA, 147 + (103 * i), 14), 4) & "," ' campo Importe franquiciado
            
            
            mSQL = mSQL & .numero(Mid(CADENA, 161 + (103 * i), 14), 4) & "," ' campo Importes
            
            '-------------------------------------------------------------------------------------
            mSQL = mSQL & .texto(Mid(CADENA, 690, 11)) & ")"   ' campo Libre
            .ejecutar mSQL
            
        End If

    End With
    Next i
End Sub


'**************************************************************************************
'**************************************************************************************
'Antes de entrar Bolbaite y demas
'
''''Private Sub cargaRegistro330(CADENA As String) ' detalle de llamadas
''''    Dim I As Integer
''''    Dim mSQL As String
''''    For I = 0 To 3
''''    With mdbFac2
''''        mSQL = "insert into telefono.detalle_de_llamadas (Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura"
''''        mSQL = mSQL & ",Numero_de_prefactura,Numero_de_telefono,Numero_de_extension,Codigo_de_trafico,Tipo_de_trafico"
''''        mSQL = mSQL & ",Codigo_abreviatura_llamada,Descripcion_tipo_de_llamada,Codigo_destino,Tipo_destino"
''''        mSQL = mSQL & ",Operador,Numero_llamado,Fecha,Codigo_horario,Hora_inicio,Unidad_de_medida"
''''        mSQL = mSQL & ",Cantidad_medida_originada,Cantidad_medida_recibida,Unidad_calidad_transmision"
''''        mSQL = mSQL & ",calidad_suministrada,Importe,Libre)"
''''        mSQL = mSQL & " values("
''''        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
''''        mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
''''        mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registro
''''        mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo N.Factura
''''        mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo N. de PreFactura
''''        mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo N. de telefono
''''        mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo N. de extensión
''''        mSQL = mSQL & .texto(Mid(CADENA, 56, 3)) & ","  ' campo Código de tráfico
''''        mSQL = mSQL & .texto(Mid(CADENA, 59, 30)) & ","  ' campo Tipo de tráfico
''''        '-- zona que se repite en el registro original ---------------------------------------
''''        mSQL = mSQL & .texto(Mid(CADENA, 89 + (148 * I), 3)) & "," ' campo Código de abreviatura
''''        mSQL = mSQL & .texto(Mid(CADENA, 92 + (148 * I), 30)) & "," ' campo Descripción de tipo de llamada
''''        mSQL = mSQL & .texto(Mid(CADENA, 122 + (148 * I), 3)) & "," ' campo Código de desrtino
''''        mSQL = mSQL & .texto(Mid(CADENA, 125 + (148 * I), 14)) & "," ' campo Tipo de destino
''''        mSQL = mSQL & .texto(Mid(CADENA, 139 + (148 * I), 8)) & "," ' campo Operador
''''        mSQL = mSQL & .texto(Mid(CADENA, 147 + (148 * I), 25)) & "," ' campo Número llamado
''''        mSQL = mSQL & .texto(Mid(CADENA, 172 + (148 * I), 4)) & "," ' campo Fecha (MMDD)
''''        mSQL = mSQL & .texto(Mid(CADENA, 176 + (148 * I), 2)) & "," ' campo Código horario
''''        mSQL = mSQL & .texto(Mid(CADENA, 178 + (148 * I), 4)) & "," ' campo Hora (HHMM)
''''        mSQL = mSQL & .texto(Mid(CADENA, 182 + (148 * I), 10)) & "," ' campo Unidad de medida
''''        mSQL = mSQL & .numero(Mid(CADENA, 192 + (148 * I), 7), 2) & "," ' campo Cantidad medida originada
''''        mSQL = mSQL & .numero(Mid(CADENA, 199 + (148 * I), 7), 2) & "," ' campo Cantidad medida recibida
''''        mSQL = mSQL & .texto(Mid(CADENA, 206 + (148 * I), 10)) & "," ' campo Unidad calidad trasmisión
''''        mSQL = mSQL & .numero(Mid(CADENA, 216 + (148 * I), 7), 2) & "," ' campo Calidad suministrada
''''        mSQL = mSQL & .numero(Mid(CADENA, 223 + (148 * I), 14), 4) & "," ' campo Importe
''''        '-------------------------------------------------------------------------------------
''''        mSQL = mSQL & .texto(Mid(CADENA, 681, 20)) & ")"   ' campo Libre
''''        .Ejecutar mSQL
''''    End With
''''    Next I
''''End Sub



Private Sub cargaRegistro330(CADENA As String) ' detalle de llamadas
    Dim i As Integer
    Dim mSQL As String
    Dim CodigoTrafico As String
    Dim CadenaInsert As String
    
    Dim recalcular As Boolean  'Si es bolbaite y demas recalcula los importes 17,18 (parametrizar)
    Dim i1 As Currency
    Dim Tot As Currency
    Dim OrigenContado As Currency
    Dim Aux As String
    Dim ImporteIncial As Currency
    
    Dim Libre As String  'Guardare el importe original
    
    For i = 0 To 3
        With mdbFac2
            'mSQL = "insert into telefono.detalle_de_llamadas (Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura"
            'mSQL = mSQL & ",Numero_de_prefactura,Numero_de_telefono,Numero_de_extension,Codigo_de_trafico,Tipo_de_trafico"
            'mSQL = mSQL & ",Codigo_abreviatura_llamada,Descripcion_tipo_de_llamada,Codigo_destino,Tipo_destino"
            'mSQL = mSQL & ",Operador,Numero_llamado,Fecha,Codigo_horario,Hora_inicio,Unidad_de_medida"
            'mSQL = mSQL & ",Cantidad_medida_originada,Cantidad_medida_recibida,Unidad_calidad_transmision"
            'mSQL = mSQL & ",calidad_suministrada,Importe,Libre)"
            'mSQL = mSQL & " values("
            mSQL = ", (" & .texto(mFicheroCorto) & ","   ' campo Fichero
            mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
            mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registro
            mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo N.Factura
            mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo N. de PreFactura
            mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo N. de telefono
            mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo N. de extensión
            CodigoTrafico = Trim(Mid(CADENA, 56, 3))
            mSQL = mSQL & .texto(CodigoTrafico) & ","   ' campo Código de tráfico
            mSQL = mSQL & .texto(Mid(CADENA, 59, 30)) & ","  ' campo Tipo de tráfico
            '-- zona que se repite en el registro original ---------------------------------------
            recalcular = False
            If vParamAplic.TieneTelefonia2 = 3 Then
                'mvarConceptosRefacturables = ",'17','18','w1'"
                
                CodigoTrafico = "'" & CodigoTrafico & "'"
                If InStr(1, mvarConceptosLlamadasRefacturables, CodigoTrafico) > 0 Then
                    'LAS ESPECIALES NO se refacturan
                    Aux = Mid(CADENA, 89 + (148 * i), 3)
                    If Trim(Aux) = "E" Then
                        'recalcular = True
                        
                    Else
                        recalcular = True
                    End If
                End If
                
            End If
            
            
            'Comun
            mSQL = mSQL & .texto(Mid(CADENA, 89 + (148 * i), 3)) & "," ' campo Código de abreviatura
            mSQL = mSQL & .texto(Mid(CADENA, 92 + (148 * i), 30)) & "," ' campo Descripción de tipo de llamada
            mSQL = mSQL & .texto(Mid(CADENA, 122 + (148 * i), 3)) & "," ' campo Código de desrtino
            mSQL = mSQL & .texto(Mid(CADENA, 125 + (148 * i), 14)) & "," ' campo Tipo de destino
            mSQL = mSQL & .texto(Mid(CADENA, 139 + (148 * i), 8)) & "," ' campo Operador
            mSQL = mSQL & .texto(Trim(Mid(CADENA, 147 + (148 * i), 25))) & "," ' campo Número llamado
            mSQL = mSQL & .texto(Mid(CADENA, 172 + (148 * i), 4)) & "," ' campo Fecha (MMDD)
            mSQL = mSQL & .texto(Mid(CADENA, 176 + (148 * i), 2)) & "," ' campo Código horario
            mSQL = mSQL & .texto(Mid(CADENA, 178 + (148 * i), 4)) & "," ' campo Hora (HHMM)
            mSQL = mSQL & .texto(Mid(CADENA, 182 + (148 * i), 10)) & "," ' campo Unidad de medida
            mSQL = mSQL & .numero(Mid(CADENA, 192 + (148 * i), 7), 2) & "," ' campo Cantidad medida originada
            mSQL = mSQL & .numero(Mid(CADENA, 199 + (148 * i), 7), 2) & "," ' campo Cantidad medida recibida
            mSQL = mSQL & .texto(Mid(CADENA, 206 + (148 * i), 10)) & "," ' campo Unidad calidad trasmisión
            mSQL = mSQL & .numero(Mid(CADENA, 216 + (148 * i), 7), 2) & "," ' campo Calidad suministrada
            
            
            
            Libre = Mid(CADENA, 681, 20) 'la cadena que va en la ultiposicion
            If Not recalcular Then

                mSQL = mSQL & .numero(Mid(CADENA, 223 + (148 * i), 14), 4) & "," ' campo Importe
                '-------------------------------------------------------------------------------------
                
            Else
                'RECALCULAMOS
                'Hacen una cosa muy rara. Me explico.  En duracion va un currency (2decimales)
                'Sin embargo el numero en los decimales son SEGUNDOS, no son convertido a decimal
                'Me sigo explicando
                'Ejemplo:   5.49    5 minutos y 49 segundos
                '           0.12    12 segundos
                
                'Si el importe facturado es CERO no hacemos nada mas
                Aux = .numero(Mid(CADENA, 223 + (148 * i), 14), 4) 'importe linea
                ImporteIncial = TransformaPuntosComas(Aux)
                
                If ImporteIncial <> 0 Then
                    
                    Libre = .numero(Mid(CADENA, 223 + (148 * i), 14), 4)
                    'Cantidad medida orignalmente
                    Aux = .numero(Mid(CADENA, 192 + (148 * i), 7), 2)
                    OrigenContado = TransformaPuntosComas(Aux)
                    
                    i1 = Int(OrigenContado)  'minutos
                    Tot = mvarPrecioEstablecimiento + (i1 * mvarPrecioMinuto)       'etablecimiento + minutos en ctos
                    
                    i1 = OrigenContado - i1  'segundos
                    i1 = (i1 * 100) / 60
                    i1 = i1 * mvarPrecioMinuto  'importe segundos
                    Tot = Tot + i1
                    Tot = Tot / 100
                    
                    
                Else
                    
                    Tot = 0
                End If
                
                'Cambia este trozo UNICAMENTE
                'mSQL = mSQL & .numero(Mid(CADENA, 223 + (148 * I), 14), 4) & "," ' campo Importe
                mSQL = mSQL & .numero(Tot) & "," ' campo Importe
                
                
                
            End If
            
            
            mSQL = mSQL & .texto(Libre) & ")"   ' campo Libre
            
            '.Ejecutar mSQL
            CadenaInsert = CadenaInsert & mSQL
            
        End With
    Next i
    CadenaInsert = Mid(CadenaInsert, 2)
    mSQL = "insert into telefono.detalle_de_llamadas (Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura"
    mSQL = mSQL & ",Numero_de_prefactura,Numero_de_telefono,Numero_de_extension,Codigo_de_trafico,Tipo_de_trafico"
    mSQL = mSQL & ",Codigo_abreviatura_llamada,Descripcion_tipo_de_llamada,Codigo_destino,Tipo_destino"
    mSQL = mSQL & ",Operador,Numero_llamado,Fecha,Codigo_horario,Hora_inicio,Unidad_de_medida"
    mSQL = mSQL & ",Cantidad_medida_originada,Cantidad_medida_recibida,Unidad_calidad_transmision"
    mSQL = mSQL & ",calidad_suministrada,Importe,Libre)"
    mSQL = mSQL & " values " & CadenaInsert
    mdbFac2.ejecutar mSQL
End Sub

Private Sub cargaRegistro340(CADENA As String) ' varios
    Dim i As Integer
    Dim mSQL As String
    For i = 0 To 4
    With mdbFac2
        mSQL = "insert into telefono.varios values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registro
        mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo N. de PreFactura
        mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo N. de telefono
        mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo N. de extensión
        '-- zona que se repite en el registro original ---------------------------------------
        mSQL = mSQL & .texto(Mid(CADENA, 56 + (117 * i), 3)) & "," ' campo Código de Vario
        mSQL = mSQL & .texto(Mid(CADENA, 59 + (117 * i), 100)) & "," ' campo Descripción de Vario
        mSQL = mSQL & .numero(Mid(CADENA, 159 + (117 * i), 14), 4) & "," ' campo Importe
        '-------------------------------------------------------------------------------------
        mSQL = mSQL & .texto(Mid(CADENA, 641, 60)) & ")"   ' campo Libre
        .ejecutar mSQL
    End With
    Next i
End Sub

Private Sub cargaRegistro850(CADENA As String) ' Descuentos por factura
    Dim i As Integer
    Dim mSQL As String
    For i = 0 To 4
    With mdbFac2
        mSQL = "insert into telefono.descuentos_por_factura values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo Libre
        '----------- Zona que se repite
        mSQL = mSQL & .texto(Mid(CADENA, 56 + (120 * i), 3)) & "," ' campo Código de descuento
        mSQL = mSQL & .texto(Mid(CADENA, 59 + (120 * i), 40)) & "," ' campo Descripción de descuento
        mSQL = mSQL & .texto(Mid(CADENA, 99 + (120 * i), 3)) & "," ' campo Código de tipo de descuento
        mSQL = mSQL & .texto(Mid(CADENA, 102 + (120 * i), 1)) & "," ' campo Indicador tramo / base
        mSQL = mSQL & .texto(Mid(CADENA, 103 + (120 * i), 25)) & "," ' campo Tramo / base
        mSQL = mSQL & .numero(Mid(CADENA, 128 + (120 * i), 14), 4) & "," ' campo Importe
        mSQL = mSQL & .numero(Mid(CADENA, 142 + (120 * i), 5), 2) & "," ' campo Porcentaje descuento
        mSQL = mSQL & .texto(Mid(CADENA, 147 + (120 * i), 1)) & "," ' campo Indicador descuento global
        mSQL = mSQL & .numero(Mid(CADENA, 148 + (120 * i), 14), 4) & "," ' campo Descuento global
        mSQL = mSQL & .numero(Mid(CADENA, 162 + (120 * i), 14), 4) & "," ' campo Importe descuento
        '-----------------
        mSQL = mSQL & .texto(Mid(CADENA, 656, 45)) & ")"   ' campo Libre (??)
        
        .ejecutar mSQL
    End With
    Next i
End Sub

Private Sub cargaRegistro990(CADENA As String) ' totales del soporte
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.totales_del_soporte values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(CADENA, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(CADENA, 4, 6)) & ","   ' campo Secuencia de registro
        mSQL = mSQL & .texto(Mid(CADENA, 10, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 24, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 38, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(CADENA, 47, 9)) & ","   ' campo Libre
        mSQL = mSQL & .numero(Mid(CADENA, 56, 24), 4) & ","  ' campo Total Cuotas
        mSQL = mSQL & .numero(Mid(CADENA, 80, 24), 4) & ","  ' campo Total varios
        mSQL = mSQL & .numero(Mid(CADENA, 104, 24), 4) & "," ' campo Total descuentos
        mSQL = mSQL & .numero(Mid(CADENA, 128, 24), 4) & ","  ' campo Total servicio medido
        mSQL = mSQL & .numero(Mid(CADENA, 152, 24), 4) & ","  ' campo Suma Total
        mSQL = mSQL & .numero(Mid(CADENA, 176, 24), 4) & ","  ' campo Total impuestos
        mSQL = mSQL & .numero(Mid(CADENA, 200, 22), 2) & ","  ' campo Total a pagar
        mSQL = mSQL & .numero(Mid(CADENA, 222, 8)) & ","   ' campo Total registros
        mSQL = mSQL & .numero(Mid(CADENA, 230, 8)) & ","   ' campo Total clientes
        mSQL = mSQL & .numero(Mid(CADENA, 238, 8)) & ","   ' campo Total SERVICIOS
        mSQL = mSQL & .numero(Mid(CADENA, 246, 8)) & ","   ' campo Total facturas
        mSQL = mSQL & .numero(Mid(CADENA, 254, 8)) & ","   ' campo Total prefacturas
        mSQL = mSQL & .numero(Mid(CADENA, 262, 8)) & ","   ' campo Total teléfonos fijos
        mSQL = mSQL & .numero(Mid(CADENA, 270, 8)) & ","   ' campo Total teléfonos móviles
        mSQL = mSQL & .texto(Mid(CADENA, 278, 5)) & ","   ' campo Código de soporte
        mSQL = mSQL & .texto(Mid(CADENA, 283, 6)) & ","   ' campo Código de agrupación
        mSQL = mSQL & .texto(Mid(CADENA, 289, 412)) & ")"   ' campo Libre (??)
        .ejecutar mSQL
    End With
End Sub

Private Sub Class_Initialize()

    Set mdbFac2 = New TelBaseDatos
    mdbFac2.Tipo = "MYSQL"
End Sub

Private Function fechatel(mfecha As String) As Date
    '-- Cambia una fecha pasada en formato AAAAMMDD a un tipo DATE
    fechatel = CDate(Mid(mfecha, 7, 2) & "/" & Mid(mfecha, 5, 2) & "/" & Mid(mfecha, 1, 4))
End Function

Private Sub borraDatos(Fichero As String)
    Dim mSQL As String
    '-- Se encarga de borrar los datos contenidos en todas las tablas para un fichero concreto
    With mdbFac2
        mSQL = "delete from telefono.ficheros where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.cabecera_de_soporte where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.factura where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.telefono where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.cuotas where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.resumen_de_llamadas where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.detalle_de_llamadas where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.resumen_de_llamadas where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.varios where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.descuentos_por_factura where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
        mSQL = "delete  from telefono.totales_del_soporte where Fichero = " & .texto(Fichero)
        .ejecutar mSQL
    End With
End Sub


Public Function EmitirFacturas_(mfichero As String, mfecha As Date, ByRef Lbl As Label, LaCompanyia As Byte) As Boolean
    If LaCompanyia = 3 Then
        'VODAFONE
        
        EmitirFacturas_ = EmitirFacturasVODAFONE(mfichero, mfecha, Lbl, LaCompanyia)
    Else
        EmitirFacturas_ = EmitirFacturasMOVI_ORANGE(mfichero, mfecha, Lbl, LaCompanyia)
        
    End If
End Function

Public Function DevuelveElNombreFicheroTraspasado() As String
    DevuelveElNombreFicheroTraspasado = mvarFichero 'mvarFicheromFicheroCorto
End Function

Private Function EmitirFacturasMOVI_ORANGE(mfichero As String, mfecha As Date, ByRef Lbl As Label, LaCompanyia As Byte) As Boolean
    Dim mSQL As String
    Dim msql2 As String
    Dim mrs0 As ADODB.Recordset
    Dim mrs1 As ADODB.Recordset
    Dim mrs2 As ADODB.Recordset
    Dim mAno As Integer '-- valor del año de facturación
    Dim AntTelefono As String '-- lleva el último teléfono leido
    Dim mLinea As Integer ' lleva el número de línea
    
    Dim mSumConsumos As Currency
    Dim mSumConsumosNoLlamadas As Currency   'Cuando el mvarDescuentosPorConsumo desdoblaremos para separar NO llamadas
    Dim mSumCuotas As Currency
    Dim mSumDescuentos As Currency
    Dim mSumEspecial As Currency
    Dim mTotalPrevio As Currency
    Dim mTotal As Currency
    Dim mDescuento As Currency
    Dim mPorcentaje As Integer
    Dim mBase As Currency
    Dim mCuota As Currency
    Dim mCuotaMinimia As Currency
    Dim UltLineaCuota As Integer
    Dim NF As Integer
    Dim i As Integer
    Dim mFecha_inicio_periodo As Date
    Dim mFecha_fin_periodo As Date
    
    Dim ContadorFAI_Incial As Long  'Para reestablecerlo al final del proceso
    Dim vC As CTiposMov
    Dim vI As CTiposMov
    
    
    
    Dim b As Boolean
    
    Dim Nombre As String
    Dim Apell1 As String
    Dim Apell2 As String
    Dim mFactura3 As Long  ' lleva el número de factura en proceso
    Dim mSerie As String
    Dim mValorIva As Currency
    Dim iNSERTS As String
    
    
    
    
    
    Dim PorTelefono As String
    Dim TelefonosFactura As String
    Dim AntCliente As Long
    Dim AntAgrupacion As Integer
    Dim CreaFac As Boolean
    Dim ColTelefonos As Collection
    Dim numAgrupados As Integer
    Dim J As Integer
    
    Dim DescCuotaMinima As String
    Dim DtoConsumo As Currency
    Dim ImporteTemporal As Currency
    Dim ImporteTemporalSin As Currency
    Dim ImporteTMP As Currency
    
    EmitirFacturasMOVI_ORANGE = True  'Con errores
    NF = -1
    On Error GoTo eEmitirFacturasMOVI_ORANGE
    
    
    
    '-- Este método crea las facturas correspondientes a un determinado
    '   fichero, asignándoles una fecha de factura determinada. Ambos
    '   valores se le pasan por parámetros
    '-- Por si nos pasan ruta completa modificamos el nombre de fichero
    If Len(mfichero) > 12 Then mfichero = Right(mfichero, 12)
        
    mSQL = "select distinct(Fichero) from tel_cab_factura where not Fichero in (select Fichero from tel_fichtraspasados)"
    Set mrs0 = mdbFac2.cursor(mSQL)
    i = 0
    While Not mrs0.EOF
        If mrs0.Fields(0) <> mfichero Then i = i + 1
        mrs0.MoveNext
    Wend
    mrs0.Close
    
    If i > 1 Then
        MsgBox "No pueden haber mas de un fichero en proceso", vbExclamation
        Exit Function
    End If
    
    
    
    Set vC = New CTiposMov
    If Not vC.Leer("FAT") Then Exit Function
                       'LAS INTERNAS
    
    'Habria que ver el CODIGO COARVAL para catadau:  9000722
    If mvarDigitoCoarval > 0 Then vC.Contador = (mvarDigitoCoarval * 1000000) + vC.Contador

    
    
    Set vI = New CTiposMov
    If Not vI.Leer("FAI") Then Exit Function
    ContadorFAI_Incial = vI.Contador
    vI.Contador = vI.Contador + 1000  'para evitar problemas luego
    
    '-- Borramos cualquier generación previa de ese fichero
    'Cojo el numero de contador mas bajo del fichero
    Lbl.Caption = "Restablecer contador"
    Lbl.Refresh
    mSQL = "select serie,min(NumFact) Minimo from tel_cab_factura where Fichero = " & DBSet(mfichero, "T") & " GROUP BY 1"
    Set mrs0 = mdbFac2.cursor(mSQL)
    mSQL = ""
    While Not mrs0.EOF
        If mrs0!Serie = Trim(vC.LetraSerie) Then
            vC.Contador = mrs0!Minimo - 1
            If mvarDigitoCoarval > 0 Then vC.Contador = Mid(CStr(vC.Contador), 2)
            
            'Dejo el que toca: incrementar hace c++ and pdate
            vC.Contador = vC.Contador - 1
            vC.IncrementarContador vC.TipoMovimiento   'Como hace el update con c+1
            
            If mvarDigitoCoarval > 0 Then vC.Contador = (mvarDigitoCoarval * 1000000) + vC.Contador
            
            
        ElseIf mrs0!Serie = Trim(vI.LetraSerie) Then
            

            msql2 = ""
        Else
           mSQL = "Error obteniedo letra serie"
        End If
        mrs0.MoveNext
    Wend
    mrs0.Close
    If mSQL <> "" Then
        MsgBox mSQL, vbExclamation
        Exit Function
    End If
    
    
    Lbl.Caption = "Proceso borre"
    Lbl.Refresh
    borrarFacturas2 mfichero, Lbl
    
    
    'SEPTIEMBRE 2013
    'NO borra los datos ya que se pueden meter a mano albaranes
    'Ultimo acceso. Los albaranes ALT van "todo o nada". Es decir solo se crean para genrar la fra
    'Y la referencia es el nombre del fichero
    'conn.Execute "DELETE from slialb where codtipom='ALT';"
    'conn.Execute "DELETE from scaalb where codtipom='ALT';"

    
    
    mSQL = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtiTelefonia, "T")
    mSQL = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", mSQL)
    If mSQL = "" Then
        MsgBox "Error obteniendo valor IVA", vbExclamation
        Exit Function
    End If
    mValorIva = CCur(mSQL)
    
    '-- Calculamos Año de facturación, lo necesitamos para el contador de facturas
    mAno = Year(mfecha)
    '-- 0 Tomamos como base la tabla de teléfonos para dar de alta las cabeceras de factura
    
    
    
    mSQL = "select * from telefono.telefono" & _
            " where Fichero = " & mdbFac2.texto(mfichero) & _
            " order by Numero_de_telefono"
      'Facturacion agrupada antes-despes
    'mSQL = "select distinct(substring(telef,1,9)) elTelefono from vodafonetxt.dcul_cab order by 1"
    mSQL = "SELECT idtelefono, Numero_de_telefono,codclien,agrupacion from telefono.telefono "
    mSQL = mSQL & " LEFT JOIN sclientfno on Numero_de_telefono = IdTelefono"
    mSQL = mSQL & " WHERE Fichero = " & mdbFac2.texto(mfichero) & " ORDER BY  codclien,agrupacion,Numero_de_telefono"
    
    
    
    i = 0
    mrs0.Open mSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    NF = FreeFile()
    Open App.Path & "\emitefac.log" For Output As #NF
    Print #NF, "Facturas emitidas de fichero " & mfichero & " y fecha " & Format(mfecha, "dd/mm/yyyy")
    
    
    
    
    ' TelefonosFactura AntCliente Agrupa
    Set ColTelefonos = New Collection
    AntCliente = -1
    AntTelefono = ""
    TelefonosFactura = ""
    UltLineaCuota = 0
    While Not mrs0.EOF
    
        
    
        If IsNull(mrs0!codClien) Then
            Print #NF, DBLet(mrs0!Numero_de_telefono, "T")
            UltLineaCuota = UltLineaCuota + 1
        Else
            If mrs0!codClien <> AntCliente Then
                If AntCliente >= 0 Then
                    mSQL = AntTelefono & "|" & numAgrupados & "|1|" & TelefonosFactura & "|"
                    ColTelefonos.Add mSQL
                End If
                AntCliente = mrs0!codClien
                TelefonosFactura = mrs0!idtelefono
                AntAgrupacion = mrs0!Agrupacion
                AntTelefono = mrs0!idtelefono
                numAgrupados = 1
            Else
                If mrs0!idtelefono = AntTelefono Then
                    'no hacemos nada
                    
                Else
                    If AntAgrupacion = 0 Then
                        'Se inserta directemente para facturar
                        mSQL = AntTelefono & "|" & numAgrupados & "|0|" & TelefonosFactura & "|"
                        ColTelefonos.Add mSQL
                        AntTelefono = mrs0!idtelefono
                        AntAgrupacion = mrs0!Agrupacion
                        TelefonosFactura = mrs0!idtelefono
                        numAgrupados = 1
                    Else
                        If mrs0!Agrupacion = AntAgrupacion Then
                            'Insertamos el que hay
                            TelefonosFactura = TelefonosFactura & "@" & mrs0!idtelefono
                            AntTelefono = mrs0!idtelefono
                            numAgrupados = numAgrupados + 1
                        Else
                            'Añadmios el telefono
                            Stop
                            'hemos pasado de agrupacion 1 a cero. Eso no pued ser por el SQL
                            
                        End If
                    End If
                End If
            End If
        End If  'si es nulo
        mrs0.MoveNext
    Wend
    mrs0.Close
    
    'El ultimo
    If AntTelefono = "" Then Err.Raise 513, , "No existe ultimo telefono facturable"
    mSQL = IIf(numAgrupados > 1, 1, 0)
    mSQL = AntTelefono & "|" & numAgrupados & "|" & mSQL & "|" & TelefonosFactura & "|"
    ColTelefonos.Add mSQL

    
    
    
    If UltLineaCuota > 0 Then Err.Raise 513, , "Error leyendo telefnos"
        
        
    
    
    
    
    
    
    For i = 1 To ColTelefonos.Count
        If (i Mod 20) = 1 Then DoEvents
       
        
        mSQL = ColTelefonos.Item(i)
        
        AntTelefono = RecuperaValor(mSQL, 2)
        numAgrupados = CInt(AntTelefono)
        AntAgrupacion = RecuperaValor(mSQL, 3)
        If CInt(AntAgrupacion) > 1 Then Err.Raise 513, , "Error leyendo TAG agrupacion: " & mSQL & " (" & i & ")"
        
        
        AntTelefono = RecuperaValor(mSQL, 1)
        TelefonosFactura = RecuperaValor(mSQL, 4)
        If TelefonosFactura = "" Then Err.Raise 513, , "Error leyendo telefonos a facturar cliente : " & mSQL & " (" & i & ")"
        TelefonosFactura = TelefonosFactura & "|"
        If numAgrupados > 1 Then TelefonosFactura = Replace(TelefonosFactura, "@", "|")
                  
        Lbl.Caption = i & "/" & ColTelefonos.Count & " " & AntTelefono    ' RTef!eltelefono
        Lbl.Refresh
     
    
    
    
        'Cogera los datos de la sclien
        mSQL = "select nifclien nifclien,nomclien,domclien,codpobla,pobclien,proclien,operador,CodDirec,sclientfno.codclien,clienppal"
        mSQL = mSQL & " from sclien,sclientfno  where sclien.codclien=sclientfno.codclien AND "
        mSQL = mSQL & " IdTelefono = " & mdbFac2.texto(AntTelefono)
        mrs0.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If mrs0!nifClien = vParam.CifEmpresa Then
            vI.Contador = vI.Contador + 1
            mFactura3 = vI.Contador
            mSerie = Trim(vI.LetraSerie)
            
            'El iva
            iNSERTS = 0
        Else
            vC.Contador = vC.Contador + 1
            mFactura3 = vC.Contador
            mSerie = Trim(vC.LetraSerie)
            'El iva
            iNSERTS = mValorIva
        End If
        
        If numAgrupados > 1 Then
            AntTelefono = Format(mrs0!codClien, "000000") & "/" & "01"  'de momento solo acepta una agrupacion
        Else
            'Es una agrupacion del cliente. Lleva varios telefonos juntos
            'NO hacemos nada puesto que AntTelefono=telefonosfactura
            
        End If

        
        
        
        '  mFactura2 = obtenNumeroFactura(vC.LetraSerie, mAno)  YA LO HA CONSEGUIDO ARRIBA conseguircontador
        mSQL = "insert into tel_cab_factura (Serie,Ano,NumFact,Fichero,Fecha,Telefono,TipoIVA,TotalPrevio,BaseImponible,Cuota,Total, EsAgupacion  ,ClienPpal"
        mSQL = mSQL & ",NIF,Apellido1,Apellido2,Nombre,Direccion,CodPostal,Poblacion,Provincia,companyia ) VALUES ("
        
        mSQL = mSQL & mdbFac2.texto(mSerie) & ","  ' Serie
        mSQL = mSQL & mdbFac2.numero(mAno) & "," ' Año
        mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
        
        mSQL = mSQL & mdbFac2.texto(mfichero) & "," ' Fichero
        mSQL = mSQL & mdbFac2.Fecha(mfecha) & "," ' Fecha
        mSQL = mSQL & mdbFac2.texto(AntTelefono) & " ," 'mrs0!Numero_de_telefono) & "," ' Telefono
        
        mSQL = mSQL & mdbFac2.numero(iNSERTS) & "," ' Si es INTERNA va a cero, sino mValorIva
        mSQL = mSQL & mdbFac2.numero(0) & "," ' Total previo (será recalculada)
        mSQL = mSQL & mdbFac2.numero(0) & "," ' Base Imponible (será recalculada)
        mSQL = mSQL & mdbFac2.numero(0) & "," ' Cuota (será recalculada)
        mSQL = mSQL & mdbFac2.numero(0) & ","
        
        
        'EsAgupacion
        mSQL = mSQL & mdbFac2.numero(IIf(numAgrupados > 1, 1, 0)) & ","
        ',ClienPpal
        If IsNull(mrs0!clienppal) Then
            mSQL = mSQL & "null"
        Else
            mSQL = mSQL & mrs0!clienppal
        End If
        mSQL = mSQL & ","
        mSQL = mSQL & mdbFac2.texto(mrs0!nifClien)
                
                
                
        If DBLet(mrs0!CodDirec, "T") <> "" Then
            Apell2 = "codclien = " & mrs0!codClien & " AND coddirec = " & mrs0!CodDirec
            Apell1 = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Apell2 & " AND 1", "1")
            If Apell1 <> "" Then
                'TIENE 'departameto', es decir, va a otro nombre
                'cierro el rs
                
                mrs0.Close
                
                Apell1 = "Select nomdirec as nomclien,domdirec domclien,codpobla,"
                Apell1 = Apell1 & " pobdirec pobclien,prodirec proclien"
                Apell1 = Apell1 & " from sdirec where " & Apell2
                mrs0.Open Apell1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Apell1 = "": Apell2 = ""
                If mrs0.EOF Then Err.Raise 513, , "Error leyendo dpto /dir : " & mrs0.Source
            Else
                Print #NF, "Codirec: " & miRsAux!eltelefono & " " & Apell2
            End If
        End If
        SeparaNomclienNomApellidos mrs0!NomClien, Nombre, Apell1, Apell2
        
        
        
        mSQL = mSQL & " ,  " & mdbFac2.texto(Apell1)
        mSQL = mSQL & " ,  " & mdbFac2.texto(Apell2)
        mSQL = mSQL & " ,  " & mdbFac2.texto(Nombre)
        mSQL = mSQL & " ,  " & mdbFac2.texto(mrs0!domclien)
        mSQL = mSQL & " ,  " & mdbFac2.texto(mrs0!codpobla)
        mSQL = mSQL & " ,  " & mdbFac2.texto(mrs0!pobclien)
        mSQL = mSQL & " ,  " & mdbFac2.texto(mrs0!proclien)
        
            

        If LaCompanyia = 2 Then
            Apell1 = "ORA"
        Else
            Apell1 = "MOV"
        End If
        mSQL = mSQL & ", " & DBSet(Apell1, "T") & ")"
        mdbFac2.ejecutar mSQL



        'tel_cab_factura_agr (Serie,Ano,NumFact,Telefono)
        mSQL = ""
        For J = 1 To numAgrupados
            Apell1 = RecuperaValor(TelefonosFactura, J)
            If Apell1 = "" Then Err.Raise 513, , "Error leyendo telefonos agrupados cliente : " & miRsAux!codClien & " " & TelefonosFactura & " (n:" & J & ")"
            mSQL = mSQL & ", (" & mdbFac2.texto(mSerie) & ","  ' Serie
            mSQL = mSQL & mdbFac2.numero(Year(mfecha)) & "," ' Año
            mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
            mSQL = mSQL & mdbFac2.texto(Apell1) & ")"
        Next J
        mSQL = Mid(mSQL, 2) 'QUitamos la primera coma
        mSQL = "insert into tel_cab_factura_agr (Serie,Ano,NumFact,Telefono) VALUES " & mSQL
        
        mdbFac2.ejecutar mSQL



        mrs0.Close


    Next i
    
    
     'DescCuota CodCuota tel_desc_cuotas ARI00000
    DescCuotaMinima = DevuelveDesdeBD(conAri, "DescCuota", "tel_desc_cuotas", "CodCuota", "ARI00000", "T")
                    
    
    
    'Ajustamos contadores
    vC.Contador = vC.Contador - 1
    
    If mvarDigitoCoarval > 0 Then vC.Contador = Mid(CStr(vC.Contador), 2)
    
    vC.IncrementarContador vC.TipoMovimiento

    
    '- Ya tenemos las facturas y ahora montamos los consumos
    mSQL = "select * from tel_cab_factura" & _
            " where Fichero = " & mdbFac2.texto(mfichero)
    Set mrs0 = mdbFac2.cursor(mSQL)
    If Not mrs0.EOF Then
        '-- Hay facturas a procesar, las leemos secuencialmente
        mrs0.MoveFirst
        i = 0
        While Not mrs0.EOF
            
            i = i + 1
           
            Lbl.Caption = "Lineas factura " & CStr(i)
            Lbl.Refresh
            DoEvents
            
            EsAgrupacionTelefono = Val(mrs0!EsAgupacion) = 1
            If EsAgrupacionTelefono Then
                whereTelefono = "serie = '" & mrs0!Serie & "' AND ano = " & mrs0!Ano & " AND numfact =" & mrs0!NumFact & " AND 1"
                whereTelefono = DevuelveDesdeBD(conAri, "GROUP_CONCAT( concat('''',telefono,'''') separator ',')", "tel_cab_factura_agr", whereTelefono, "1 GROUP BY numfact")
                If whereTelefono = "" Then Err.Raise 513, , "Sin telefonos para: " & mrs0!Serie & " - " & mrs0!Ano & " - " & mrs0!NumFact
                whereTelefono = "  #campotele# IN  (" & whereTelefono & ")  "
                
            Else
                'Lo que habia. No hacemos nada
                ' Leemos las cuotas correspondientes a ese número de teléfono
                whereTelefono = " #campotele# like '" & mrs0!Telefono & "%'"
            End If
            
                
          
            mFactura3 = mrs0!NumFact
            mSerie = mrs0!Serie
            mSQL = " AND " & Replace(whereTelefono, "#campotele#", "Numero_de_telefono")
            '-- Vamos a montar cuotas
            mSQL = "SELECT cuotas.Fichero, cuotas.Numero_de_telefono, cuotas.Codigo_de_cuota, cuotas.Descripcion_de_cuota, Sum(cuotas.Importe) AS SumaDeImporte, cuotas.Fecha_inicio_cuotas, cuotas.Fecha_final_cuotas" & _
                    " FROM telefono.cuotas" & _
                    " WHERE Fichero = " & mdbFac2.texto(mfichero) & _
                    mSQL & " GROUP BY cuotas.Fichero, cuotas.Numero_de_telefono, cuotas.Codigo_de_cuota, cuotas.Descripcion_de_cuota, cuotas.Fecha_inicio_cuotas, cuotas.Fecha_final_cuotas"

            Set mrs1 = mdbFac2.cursor(mSQL)
            mLinea = 0
            UltLineaCuota = 0
            If Not mrs1.EOF Then
                mrs1.MoveFirst
                
                iNSERTS = "" 'VELOCIDAD
                                
                
                While Not mrs1.EOF
                    '-- Ahora damos de alta los valores de cuota correspondiente
                    mLinea = mLinea + 1 ' incrementamos la línea correspondiente
                    UltLineaCuota = mLinea
                    mSQL = "insert into tel_lin_factura_cuotas values("
                    mSQL = mSQL & mdbFac2.texto(mSerie) & "," ' Serie
                    mSQL = mSQL & mdbFac2.numero(mAno) & "," ' Año
                    mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
                    mSQL = mSQL & mdbFac2.numero(mLinea) & "," ' Número de línea
                    mSQL = mSQL & mdbFac2.Fecha(mrs1!Fecha_inicio_cuotas) & "," ' Fecha inicio del consumo
                    mSQL = mSQL & mdbFac2.Fecha(mrs1!Fecha_final_cuotas) & "," ' Fecha final del consumo
                    mSQL = mSQL & mdbFac2.texto(mrs1!codigo_de_cuota) & "," ' código de la cuota correspondiente
                    mSQL = mSQL & mdbFac2.texto(mrs1!Descripcion_de_cuota) & "," ' Fecha
                    'MAYO 2013. Metemos el 0 al final para el tipo de IVA(0=normal  1=exento)
                    mSQL = mSQL & mdbFac2.numero(mrs1!SumaDeImporte) & ",0," ' importe de la cuota correspondiente
                    'Abirl 2020. Pude agrupar por elefono
                    mSQL = mSQL & mdbFac2.texto(mrs1!Numero_de_telefono) & ")"
                    iNSERTS = iNSERTS & ", " & Mid(mSQL, 13)
                    
                    mdbFac2.ejecutar mSQL
                    '-- Calculamos si hay algún descuento aplicable
                    mDescuento = buscaDescuentos(0, mrs1!SumaDeImporte, mrs1!codigo_de_cuota, mPorcentaje)
                    If mDescuento <> 0 Then
                        '-- Hay descuento aplicable
                        mSQL = "insert into tel_lin_factura_descuentos values("
                        mSQL = mSQL & mdbFac2.texto(mSerie) & "," ' Serie
                        mSQL = mSQL & mdbFac2.numero(mAno) & "," ' Año
                        mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
                        mSQL = mSQL & mdbFac2.numero(siguienteLineaDeDescuentos(mSerie, mAno, mFactura3)) & "," ' Número de línea
                        mSQL = mSQL & mdbFac2.texto(mrs1!Descripcion_de_cuota) & "," ' Concepto
                        mSQL = mSQL & mdbFac2.numero(mPorcentaje) & "," ' Pordentaje
                        mSQL = mSQL & mdbFac2.numero(mrs1!SumaDeImporte) & "," ' Base
                        mSQL = mSQL & mdbFac2.numero(mDescuento) & "," ' Importe
                        'Abirl 2020. Pude agrupar por elefono
                        mSQL = mSQL & mdbFac2.texto(mrs1!Numero_de_telefono) & ")"
                        mdbFac2.ejecutar mSQL
                    End If
                    mrs1.MoveNext
                Wend
            End If
            '-- Leemos todos los consumos correspondientes al teléfono de la factura
            mSQL = " AND " & Replace(whereTelefono, "#campotele#", "Numero_de_telefono")
            mSQL = "SELECT resumen_de_llamadas.Fichero, resumen_de_llamadas.Numero_de_telefono, resumen_de_llamadas.Codigo_de_tipo_de_trafico, resumen_de_llamadas.Tipo_de_trafico, Sum(resumen_de_llamadas.Numero_de_llamadas) AS SumaDeNumero_de_llamadas, Sum(resumen_de_llamadas.Cantidad_medida) AS SumaDeCantidad_medida, Sum(resumen_de_llamadas.Cantidad_franquiciada) AS SumaDeCantidad_franquiciada, Sum(resumen_de_llamadas.Importe_franquiciado) AS SumaDeImporte_franquiciado, Sum(resumen_de_llamadas.Importes) AS SumaDeImportes, resumen_de_llamadas.Fecha_inicio_periodo, resumen_de_llamadas.Fecha_final_periodo, resumen_de_llamadas.Unidad_de_medida" & _
                    " FROM telefono.resumen_de_llamadas" & _
                    " WHERE Fichero = " & mdbFac2.texto(mfichero) & _
                    mSQL & "" & _
                    " GROUP BY resumen_de_llamadas.Fichero, resumen_de_llamadas.Numero_de_telefono, resumen_de_llamadas.Codigo_de_tipo_de_trafico, resumen_de_llamadas.Tipo_de_trafico, resumen_de_llamadas.Fecha_inicio_periodo, resumen_de_llamadas.Fecha_final_periodo, resumen_de_llamadas.Unidad_de_medida"


'            mSQL = "SELECT * from Resumen_llamadas_x_telefono" & _
'                    " WHERE Fichero = " & mdbTel.Texto(mfichero) & _
'                    " AND Numero_de_telefono = " & mdbTel.Texto(mrs0!Telefono)
            Set mrs1 = mdbFac2.cursor(mSQL)
            mLinea = 0
            If Not mrs1.EOF Then
                mrs1.MoveFirst
                
                iNSERTS = "" 'VELOCIDAD
                
                While Not mrs1.EOF
                    '-- Pasa por aqui por cada línea de consumo de este teléfono
                    mLinea = mLinea + 1 ' incrementamos la línea correspondiente
                    'mSQL = "insert into tel_lin_factura_consumos values("
                    mSQL = "("
                    mSQL = mSQL & mdbFac2.texto(mSerie) & "," ' Serie
                    mSQL = mSQL & mdbFac2.numero(mAno) & "," ' Año
                    mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
                    mSQL = mSQL & mdbFac2.numero(mLinea) & "," ' Número de línea
                    mSQL = mSQL & mdbFac2.Fecha(mrs1!Fecha_inicio_periodo) & "," ' Fecha inicio del consumo
                    mSQL = mSQL & mdbFac2.Fecha(mrs1!Fecha_final_periodo) & "," ' Fecha final del consumo
                    mSQL = mSQL & mdbFac2.texto(mrs1!Codigo_de_tipo_de_trafico) & "," ' Código del tipo de consumo
                    mSQL = mSQL & mdbFac2.texto(mrs1!Tipo_de_trafico) & "," ' Descripción del tipo de consumo
                    mSQL = mSQL & mdbFac2.numero(mrs1!SumaDeNumero_de_llamadas) & "," ' Núemro de llamadas
                    mSQL = mSQL & mdbFac2.texto(mrs1!Unidad_de_medida) & "," ' Unidad de medida del consumo
                    mSQL = mSQL & mdbFac2.numero(mrs1!SumaDeCantidad_medida) & "," ' Cantidad consumida
                    
                  
                    
                    'MAYO 2013
                    'Metemos el 0 para tipo de iva(0Normal,1Exento)
                    mSQL = mSQL & mdbFac2.numero(mrs1!SumaDeImportes) & ",0,"
                    'Abirl 2020. Pude agrupar por elefono
                    mSQL = mSQL & mdbFac2.texto(mrs1!Numero_de_telefono) & ")"
                
                    mFecha_inicio_periodo = mrs1!Fecha_inicio_periodo
                    mFecha_fin_periodo = mrs1!Fecha_final_periodo
                    
                    
                    iNSERTS = iNSERTS & ", " & mSQL
                    
                    'mdbFac2.Ejecutar mSQL
                    '-- Calculamos si hay algún descuento aplicable
                    
                    If mvarDescuentosPorConsumo Then
                        'NUevo SETP 2020   Catadau
                        mDescuento = "0"
                        
                    Else
                        'Lo que habia
                        mDescuento = buscaDescuentos(1, mrs1!SumaDeImportes, mrs1!Codigo_de_tipo_de_trafico, mPorcentaje)
                    End If
                    If mDescuento <> 0 Then
                        '-- Hay descuento aplicable
                        mSQL = "insert into tel_lin_factura_descuentos values("
                        mSQL = mSQL & mdbFac2.texto(mSerie) & "," ' Serie
                        mSQL = mSQL & mdbFac2.numero(mAno) & "," ' Año
                        mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
                        mSQL = mSQL & mdbFac2.numero(siguienteLineaDeDescuentos(mSerie, mAno, mFactura3)) & "," ' Número de línea
                        mSQL = mSQL & mdbFac2.texto(mrs1!Tipo_de_trafico) & "," ' Concepto
                        mSQL = mSQL & mdbFac2.numero(mPorcentaje) & "," ' Pordentaje
                        mSQL = mSQL & mdbFac2.numero(mrs1!SumaDeImportes) & "," ' Base
                        mSQL = mSQL & mdbFac2.numero(mDescuento) & "," ' Importe
                        'Abirl 2020. Pude agrupar por elefono
                        mSQL = mSQL & mdbFac2.texto(mrs1!Numero_de_telefono) & ")"
                    
                           mdbFac2.ejecutar mSQL
                    End If
                    mrs1.MoveNext
                Wend
                
                If iNSERTS <> "" Then
                    iNSERTS = Mid(iNSERTS, 2)
                    iNSERTS = "insert into tel_lin_factura_consumos values " & iNSERTS
                    mdbFac2.ejecutar iNSERTS
                Else
                    MsgBox "Sin consumo!!!  Avise a Ariadna Software" & vbCrLf & "El proceso continua", vbExclamation
                End If
                
                mrs1.Close
                
                
                
            End If
            
            
            
            
            '-- Ahora comprobamos si en la tabla de varios tiene algun cargo aplicable
            mSQL = "select * from telefono.varios"
            mSQL = mSQL & " where Fichero = " & mdbFac2.texto(mfichero)
            mSQL = mSQL & " AND " & Replace(whereTelefono, "#campotele#", "Numero_de_telefono")   '" and Numero_de_telefono = " & mdbFac2.texto(mrs0!Telefono)
            Set mrs2 = mdbFac2.cursor(mSQL)
            If Not mrs2.EOF Then
                '-- tiene apuntes ese teléfono / fichero en la tabla de varios
                mrs2.MoveFirst
                While Not mrs2.EOF
                    '-- comprobamos si la tabla de cargos varios incluye ese concepto lo que
                    '   indicaría que demos reflejar ese cargo en factura.
                    mSQL = "select * from tel_cargo_varios"
                    mSQL = mSQL & " where CodigoVario = " & mdbFac2.texto(mrs2!Codigo_de_vario)
                    'Octubre 2013
                    mSQL = mSQL & " AND CargarEnFactura = 1"
                    Set mrs1 = mdbFac2.cursor(mSQL)
                    If Not mrs1.EOF Then
                        mSQL = "insert into tel_lin_factura_especial values("
                        mSQL = mSQL & mdbFac2.texto(mSerie) & "," ' Serie
                        mSQL = mSQL & mdbFac2.numero(mAno) & "," ' Año
                        mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
                        mSQL = mSQL & mdbFac2.numero(siguienteLineaDeEspeciales(mSerie, mAno, mFactura3)) & "," ' Número de línea
                        mSQL = mSQL & mdbFac2.texto(mrs2!Descripcion_de_vario) & "," ' Concepto del cargo especial
                        mSQL = mSQL & mdbFac2.numero(mrs2!Importe) & ","   ' Importe
                        'Abirl 2020. Pude agrupar por elefono
                        mSQL = mSQL & mdbFac2.texto(mrs2!Numero_de_telefono) & ")"
                        mdbFac2.ejecutar mSQL
                        mrs1.Close
                    End If
                    mrs2.MoveNext
                Wend
            End If
            mrs2.Close
            
        
            '-- Ahora comprobamos si han llegado al mínimo de llamadas o no
            mSumCuotas = SumCuotas(mrs0!Serie, mAno, mrs0!NumFact)
            mSumConsumos = SumConsumos(mrs0!Serie, mAno, mrs0!NumFact)
                        
            '-- Calculamos los descuentos de grupo
            SumDescuentoGrupo mrs0!Serie, mAno, mrs0!NumFact
            '-- Y ahora actualizamos los campos de totales de la factura correspondiente
            mSumEspecial = SumEspecial(mrs0!Serie, mAno, mrs0!NumFact)
            mSumDescuentos = SumDescuentos(mrs0!Serie, mAno, mrs0!NumFact)
            mSumConsumosNoLlamadas = 0
            
            
            Debug.Print whereTelefono
            
            
            If mvarDescuentosPorConsumo Then
                        
                    'NUevo SETP 2020   Catadau
                    mSQL = " SElect idtelefono,cuotaminima from sclientfno WHERE cuotaminima>0 AND " & Replace(whereTelefono, "#campotele#", "idtelefono")
                    
                    'mrs1.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    Set mrs1 = mdbFac2.cursor(mSQL)
                    mSQL = ""
                    While Not mrs1.EOF
                        'AND Codigo_de_tipo_de_trafico in (1,2,5,6)
                        mSQL = "  AND Numero_de_telefono IN  ('" & mrs1!idtelefono & "') GROUP BY 1"   'quito la ,
                        mSQL = " FROM telefono.resumen_de_llamadas where Fichero = " & mdbFac2.texto(mfichero) & mSQL
                        mSQL = "SELECT  Codigo_de_tipo_de_trafico ,  Sum(resumen_de_llamadas.Importes) AS SumaDeImportes " & mSQL
                        Set mrs2 = mdbFac2.cursor(mSQL)
                        ImporteTemporal = 0     'LLAMADAS
                        ImporteTemporalSin = 0    'RESTO
                        
                        While Not mrs2.EOF
                            
                            If DBLet(mrs2!Codigo_de_tipo_de_trafico, "N") = 1 Or DBLet(mrs2!Codigo_de_tipo_de_trafico, "N") = 5 Or DBLet(mrs2!Codigo_de_tipo_de_trafico, "N") = 5 Or DBLet(mrs2!Codigo_de_tipo_de_trafico, "N") = 5 Then
                                ImporteTemporal = ImporteTemporal + mrs2!SumaDeImportes
                            Else
                                ImporteTemporalSin = ImporteTemporalSin + mrs2!SumaDeImportes
                            End If
                            mrs2.MoveNext
                        Wend
                        mrs2.Close
                        
                        
                                                      
                        'Ajustamos los consumos con los descuentos
                        If ImporteTemporal > 0 Then
                            'tramos descuentos
                            DtoConsumo = 0
                            If ImporteTemporal >= 20 Then
                                DtoConsumo = 45
                            ElseIf ImporteTemporal >= 10 Then
                                DtoConsumo = 40
                            ElseIf ImporteTemporal >= 6 Then
                                DtoConsumo = 35
                            Else
                                If ImporteTemporal > 0 Then DtoConsumo = 10
                            End If
                              
                            If DtoConsumo > 0 Then
    
                                
                                mSQL = "UPDATE telefono.resumen_de_llamadas SET Tipo_de_trafico =concat(Tipo_de_trafico,' ' ,'" & DtoConsumo & "%'),"
                                DtoConsumo = 1 - (DtoConsumo / 100)
                                
                              '  If mSumConsumos <> ImporteTemporal Then Stop
                              '  If ImporteTemporal2 <> 0 Then Stop
                                
                                ImporteTMP = Round2(ImporteTemporal * DtoConsumo, 2)
                                
                                mSumConsumos = mSumConsumos - ImporteTemporal + ImporteTMP
                                'El importe para el mimino se queda en importetmp
                                ImporteTemporal = ImporteTMP
                                
                                mSQL = mSQL & " importes=round((importes * " & DBSet(DtoConsumo, "N") & "),4)"
                                mSQL = mSQL & " WHERE Fichero = " & mdbFac2.texto(mfichero)
                                mSQL = mSQL & " AND Codigo_de_tipo_de_trafico in (1,2,5,6) AND importes>0 AND Numero_de_telefono =" & DBSet(mrs1!idtelefono, "T")
                                conn.Execute mSQL
                                
                                
                                'Y en la factura
                                mSQL = "UPDATE tel_lin_factura_consumos SET DescTipoTrafico =concat(DescTipoTrafico,' ' ,'" & DtoConsumo & "%'),"
                                mSQL = mSQL & " importe=round((importe * " & DBSet(DtoConsumo, "N") & "),4)"
                                mSQL = mSQL & " WHERE serie = " & mdbFac2.texto(mSerie) & " AND ano = " & mdbFac2.numero(mAno) & " AND numfact = " & mdbFac2.numero(mFactura3)
                                mSQL = mSQL & " AND CodTipoTrafico in (1,2,5,6) AND importe>0 AND telefono =" & DBSet(mrs1!idtelefono, "T")
                                conn.Execute mSQL
                               
                                
                            
                            End If
                          
                       End If
                       'para el conusmo minimo cuentan tanto llamads como mensajes y demas
                       ImporteTemporal = ImporteTemporal + ImporteTemporalSin
                       ImporteTMP = mrs1!cuotaminima - ImporteTemporal
                        ImporteTemporalSin = 0
                       If ImporteTMP > 0 Then
                              mSumCuotas = mSumCuotas + ImporteTMP
                              mLinea = UltLineaCuota + 1 ' incrementamos la línea correspondiente
                              UltLineaCuota = UltLineaCuota + 1
                              mSQL = " (" & mdbFac2.texto(mSerie) & "," & mdbFac2.numero(mAno) & "," & mdbFac2.numero(mFactura3) & "," & mdbFac2.numero(mLinea) & ","
                              mSQL = mSQL & mdbFac2.Fecha(mfecha) & "," & mdbFac2.Fecha(mfecha) & "," & mdbFac2.texto("ARI00000")
                              mSQL = mSQL & "," & mdbFac2.texto(DescCuotaMinima & " " & mrs1!idtelefono) & "," & mdbFac2.numero(ImporteTMP) & ",0," & mdbFac2.texto(mrs1!idtelefono) & ")"
                              mSQL = "insert into tel_lin_factura_cuotas values" & mSQL
                              mdbFac2.ejecutar mSQL
                        End If
                        
                                     
                        mrs1.MoveNext
                    Wend
                    mrs1.Close
            End If 'dto por consumo
            
            
            
            
            
            'TOTALES
            mTotalPrevio = mSumCuotas + mSumConsumos + mSumEspecial + mSumConsumosNoLlamadas
            mTotal = mTotalPrevio - mSumDescuentos

            
            'NUEVO SEP 2020
            'Para el resto (no llevan mvarCuotaMinimaSobreConsumos , dejamos lom que habia)
            ' De monmento NO han protestado, Casi todos van por cuotas, pero cuando agrupas, el consumo minimo deberia SSER por  cada linea
            ' si se quejan este trozo deberá ir en el de arriba, separando campos
            If Not mvarCuotaMinimaSobreConsumos Then
            
                 mSQL = " SElect idtelefono,cuotaminima from sclientfno WHERE cuotaminima>0 AND " & Replace(whereTelefono, "#campotele#", "idtelefono")
                 Set mrs1 = mdbFac2.cursor(mSQL)
                 If Not mrs1.EOF Then
                 
                    
                     mSQL = ""
                     While Not mrs1.EOF
                         mCuotaMinimia = mrs1!cuotaminima
                         
                         If mvarCuotaMinimaSobreConsumos Then
                             b = mSumConsumos < mCuotaMinimia
                         Else
                             b = mTotal < mCuotaMinimia  'lo que habia
                         End If
                         If b Then
                             
                             
                             If mvarCuotaMinimaSobreConsumos Then
                                 mCuotaMinimia = mCuotaMinimia - mSumConsumos   'lo que habia
                                 mTotal = mTotal + mCuotaMinimia
                             Else
                                 'HAy que añadir el minimo
                                 MsgBox "Error leyendo datos cuto minimia", vbCritical: End
                                 mSumEspecial = mCuotaMinimia  'Reutilizo la variable
                                 mCuotaMinimia = mCuotaMinimia - mTotal   'lo que habia
                                 mTotal = mSumEspecial
                             End If
                            
                             
                             
                             
                             
                             mLinea = UltLineaCuota + 1 ' incrementamos la línea correspondiente
                             mSQL = mSQL & ", (" & mdbFac2.texto(mSerie) & "," & mdbFac2.numero(mAno) & "," & mdbFac2.numero(mFactura3) & "," & mdbFac2.numero(mLinea) & ","
                             mSQL = mSQL & mdbFac2.Fecha(mfecha) & "," & mdbFac2.Fecha(mfecha) & "," & mdbFac2.texto("ARI00000")
                             mSQL = mSQL & "," & mdbFac2.texto(DescCuotaMinima & " " & mrs1!idtelefono) & "," & mdbFac2.numero(mCuotaMinimia) & ",0," & mdbFac2.texto(mrs1!idtelefono) & ")"
                             
                         End If
                         mrs1.MoveNext
                     Wend
                     If mSQL <> "" Then
                         mSQL = Mid(mSQL, 2)
                         mSQL = "insert into tel_lin_factura_cuotas values" & mSQL
                         mdbFac2.ejecutar mSQL
                     End If
                     
                End If
                mrs1.Close
           End If 'mvarCuotaMinimaSobreConsumos
                    
                    
                    
                    


            mBase = CCur(Format(mTotal, "0.00"))
            If mrs0!Serie = Trim(vI.LetraSerie) Then
                mCuota = 0
            Else
                mCuota = CCur(Format((mBase * mValorIva) / 100#, "0.00")) ' antes 18
            End If
            mTotal = mBase + mCuota
            mSQL = "update tel_cab_factura SET " & _
                    " TotalPrevio = " & mdbFac2.numero(mTotalPrevio) & "," & _
                    " BaseImponible = " & mdbFac2.numero(mBase) & "," & _
                    " Cuota = " & mdbFac2.numero(mCuota) & "," & _
                    " Total = " & mdbFac2.numero(mTotal) & _
                    " where Serie = " & mdbFac2.texto(mrs0!Serie) & _
                    " and Ano = " & mdbFac2.numero(mAno) & _
                    " and NumFact = " & mdbFac2.numero(mrs0!NumFact)
            mdbFac2.ejecutar mSQL
            '-- considerando un tipo de IVA de 16 calculamos la base
            mrs0.MoveNext
        Wend
    End If
    Set vC = Nothing
    If vI.Contador <> ContadorFAI_Incial Then
        vI.Contador = ContadorFAI_Incial - 1
        vI.IncrementarContador vI.TipoMovimiento
    End If
    
    
    
    
    Set vI = Nothing
    Close #NF
    NF = -1
    EmitirFacturasMOVI_ORANGE = False
    
    
    
eEmitirFacturasMOVI_ORANGE:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    
    Set mrs0 = Nothing
    Set mrs1 = Nothing
    Set mrs2 = Nothing
    Set vI = Nothing
    Set vC = Nothing
    Set ColTelefonos = Nothing
    If NF > 0 Then Close #NF
End Function

Private Sub borrarFacturas2(mfichero As String, ByRef L As Label)
    '-- Este método interno sirve para borrar todas las facturas correspondientes
    '   a un fichero seleccionado y deja los contadores actualizados en consecuencia
    Dim mrs0 As ADODB.Recordset
    Dim mSQL As String
    Dim mSerie As String
    Dim mAno As String
    Dim mNumero As Long
    '-- Por si nos pasan ruta completa modificamos el nombre de fichero
    If Len(mfichero) > 12 Then mfichero = Right(mfichero, 12)
    '-- Comenzamos a buscar todas las facturas correspondientes a este fichero
    Set mrs0 = mdbFac2.cursor("select * from tel_cab_factura where Fichero = " & mdbFac2.texto(mfichero) & " ORDER BY NumFact")
                
    If Not mrs0.EOF Then
        mrs0.MoveFirst
        mSerie = mrs0!Serie
        mAno = mrs0!Ano
        
        
        
        While Not mrs0.EOF
            L.Caption = "DEL " & mSerie & mrs0!NumFact & " - " & mAno
            L.Refresh
            With mdbFac2
                '-- Borramos todas las lineas de la factura correspondiente
                mSQL = " where Serie = " & .texto(mrs0!Serie)
                mSQL = mSQL & " and Ano = " & .numero(mrs0!Ano)
                mSQL = mSQL & " and NumFact = " & .numero(mrs0!NumFact)
                .ejecutar "delete from tel_lin_factura_consumos" & mSQL
                .ejecutar "delete from tel_lin_factura_descuentos" & mSQL
                .ejecutar "delete from tel_lin_factura_especial" & mSQL
                .ejecutar "delete from tel_lin_factura_cuotas" & mSQL
                .ejecutar "delete from tel_cab_factura_agr" & mSQL
            End With
            mrs0.MoveNext
        Wend
        '-- Borradas todas las lineas de todas las facturas, borramos cabeceras
        mdbFac2.ejecutar "delete from tel_cab_factura where Fichero = " & mdbFac2.texto(mfichero)
    End If
    

    '-- Por si acaso repaso, es que falla incomprensiblemente
'    mdbFac.ejecutar "delete from tel_lin_factura_cuotas where Serie = " & mdbFac.Texto(mSerie) & " and Ano = " & mdbFac.numero(mAno) & " and NumFact > " & mdbFac.numero(mNumero)
'    mdbFac.ejecutar "delete from tel_lin_factura_consumos where Serie = " & mdbFac.Texto(mSerie) & " and Ano = " & mdbFac.numero(mAno) & " and NumFact > " & mdbFac.numero(mNumero)
'    mdbFac.ejecutar "delete from tel_lin_factura_descuentos where Serie = " & mdbFac.Texto(mSerie) & " and Ano = " & mdbFac.numero(mAno) & " and NumFact > " & mdbFac.numero(mNumero)
'    mdbFac.ejecutar "delete from tel_lin_factura_especial where Serie = " & mdbFac.Texto(mSerie) & " and Ano = " & mdbFac.numero(mAno) & " and NumFact > " & mdbFac.numero(mNumero)
End Sub

Private Sub Class_Terminate()
    Set mdbFac2 = Nothing
End Sub





'
'
Private Function SumConsumos(Serie As String, Ano As Integer, NumFact As Long) As Currency
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    mSQL = "select sum(Importe) from tel_lin_factura_consumos"
    mSQL = mSQL & " where serie = " & mdbFac2.texto(Serie)
    mSQL = mSQL & " and ano = " & mdbFac2.numero(Ano)
    mSQL = mSQL & " and numfact = " & mdbFac2.numero(NumFact)
    Set mrs = mdbFac2.cursor(mSQL)
    If IsNull(mrs.Fields(0)) Then
        SumConsumos = 0
    Else
        SumConsumos = mrs.Fields(0)
    End If
    mrs.Close
    Set mrs = Nothing
End Function
'

Private Function SumCuotas(Serie As String, Ano As Integer, NumFact As Long) As Currency
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    mSQL = "select sum(Importe) from tel_lin_factura_cuotas"
    mSQL = mSQL & " where serie = " & mdbFac2.texto(Serie)
    mSQL = mSQL & " and ano = " & mdbFac2.numero(Ano)
    mSQL = mSQL & " and numfact = " & mdbFac2.numero(NumFact)
    Set mrs = mdbFac2.cursor(mSQL)
    If IsNull(mrs.Fields(0)) Then
        SumCuotas = 0
    Else
        SumCuotas = mrs.Fields(0)
    End If
    mrs.Close
    Set mrs = Nothing
End Function


Private Function SumDescuentos(Serie As String, Ano As Integer, NumFact As Long) As Currency
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    mSQL = "select sum(Importe) from tel_lin_factura_descuentos"
    mSQL = mSQL & " where serie = " & mdbFac2.texto(Serie)
    mSQL = mSQL & " and ano = " & mdbFac2.numero(Ano)
    mSQL = mSQL & " and numfact = " & mdbFac2.numero(NumFact)
    Set mrs = mdbFac2.cursor(mSQL)
    If IsNull(mrs.Fields(0)) Then
        SumDescuentos = 0
    Else
        SumDescuentos = mrs.Fields(0)
    End If
    mrs.Close
    Set mrs = Nothing
End Function


Private Function SumEspecial(Serie As String, Ano As Integer, NumFact As Long) As Currency
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    mSQL = "select sum(Importe) from tel_lin_factura_especial"
    mSQL = mSQL & " where serie = " & mdbFac2.texto(Serie)
    mSQL = mSQL & " and ano = " & mdbFac2.numero(Ano)
    mSQL = mSQL & " and numfact = " & mdbFac2.numero(NumFact)
    Set mrs = mdbFac2.cursor(mSQL)
    If IsNull(mrs.Fields(0)) Then
        SumEspecial = 0
    Else
        SumEspecial = mrs.Fields(0)
    End If
    mrs.Close
    Set mrs = Nothing
End Function

Private Function buscaDescuentos(Tipo As Integer, Importe As Currency, Codigo As String, ByRef Porcentaje As Integer) As Currency
    '-- buscaDescuentos:
    '   Tipo, indica 0 = Cuotas, 1= Consumos
    '   Importe = Importe sobre el que se aplicará el descuento
    '   Código = Indica el código del tipo de cuota o consumo
    '   Esta función devuelve 0 sin no hay ningún valor aplicable y el valor
    '   del descuento en el caso de que haya alguno que aplicar
    '   El campo porcentaje viene por referencia, con lo que devuielve el porcentaje aplicado
    '-- MOD:06/09/2005
    '   Se ha incluido en la tabla tel_desc_consumos una nueva columna (Grupo)
    '   Grupo = 0, el descuento no se agrupa a otros se aplica como hasta ahora
    '   Grupo <> 0, el descuento se aplica si el importe de los consumos del grupo
    '   supera el mínimo del que se trate
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    mSQL = "select Porcentaje"
    If Tipo = 0 Then
        '-- Cuotas
        mSQL = mSQL & " from tel_desc_cuotas"
        mSQL = mSQL & " where CodCuota = " & mdbFac2.texto(Codigo)
    Else
        '-- Consumos
        mSQL = mSQL & " from tel_desc_consumos"
        mSQL = mSQL & " where CodTipoTrafico = " & mdbFac2.texto(Codigo)
        mSQL = mSQL & " and Importe < " & mdbFac2.numero(Importe)
        mSQL = mSQL & " and Grupo = 0" ' no es un descuento agrupado
        mSQL = mSQL & " order by Importe DESC"
    End If
    Set mrs = mdbFac2.cursor(mSQL)
    If Not mrs.EOF Then
        '-- Calculamos el descuento
        Porcentaje = mrs.Fields(0)
        buscaDescuentos = (Importe * mrs.Fields(0) / 100#)
    Else
        buscaDescuentos = 0
        Porcentaje = 0
    End If
End Function

Private Function buscaDescuentosGrupo(Tipo As Integer, Importe As Currency, Codigo As String, ByRef Porcentaje As Integer) As Currency
    '-- buscaDescuentosGrupo:
    '   Tipo, indica 0 = Cuotas, 1= Consumos
    '   Importe = Importe sobre el que se aplicará el descuento
    '   Código = Indica el código del tipo de cuota o consumo
    '   Esta función devuelve 0 sin no hay ningún valor aplicable y el valor
    '   del descuento en el caso de que haya alguno que aplicar
    '   El campo porcentaje viene por referencia, con lo que devuielve el porcentaje aplicado
    '-- MOD:06/09/2005
    '   Se ha incluido en la tabla tel_desc_consumos una nueva columna (Grupo)
    '   Grupo = 0, el descuento no se agrupa a otros se aplica como hasta ahora
    '   Grupo <> 0, el descuento se aplica si el importe de los consumos del grupo
    '   supera el mínimo del que se trate
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    mSQL = "select Porcentaje"
    If Tipo = 0 Then
        '-- Cuotas
        mSQL = mSQL & " from tel_desc_cuotas"
        mSQL = mSQL & " where CodCuota = " & mdbFac2.texto(Codigo)
    Else
        '-- Consumos
        mSQL = mSQL & " from tel_desc_consumos"
        mSQL = mSQL & " where CodTipoTrafico = " & mdbFac2.texto(Codigo)
        mSQL = mSQL & " and Importe < " & mdbFac2.numero(Importe)
        mSQL = mSQL & " order by Importe DESC"
    End If
    Set mrs = mdbFac2.cursor(mSQL)
    If Not mrs.EOF Then
        '-- Calculamos el descuento
        Porcentaje = mrs.Fields(0)
        buscaDescuentosGrupo = (Importe * mrs.Fields(0) / 100#)
    Else
        buscaDescuentosGrupo = 0
        Porcentaje = 0
    End If
End Function

Private Function SumDescuentoGrupo(Serie As String, Ano As Integer, NumFact2 As Long) As Currency
    '-- SumDescuentoGrupo:
    '   Calcula si los descuentos agrupados son aplicables o no y en caso
    '   afirmativo da de alta las líneas de descuento que corresponden.
    Dim mSQL As String
    Dim msql2 As String
    Dim Aux As String
    Dim mrs As ADODB.Recordset
    Dim mrs2 As ADODB.Recordset
    Dim mrs3 As ADODB.Recordset
    Dim ImporteProv As Currency
    Dim DescuentoProv As Currency
    Dim PorcentajeProv As Integer
    Dim CodTrafico As String
    Dim UnoDeLosTelefonos As String
    mSQL = "select distinct(Grupo) from tel_desc_consumos where Grupo <> 0"
    Set mrs = mdbFac2.cursor(mSQL)
    UnoDeLosTelefonos = ""
    If Not mrs.EOF Then
        mrs.MoveFirst
        While Not mrs.EOF
            ImporteProv = 0
            DescuentoProv = 0
            PorcentajeProv = 0
            mSQL = "select distinct(CodTipoTrafico) from tel_desc_consumos where Grupo = " & mdbFac2.numero(mrs.Fields(0))
            Set mrs2 = mdbFac2.cursor(mSQL)
            If Not mrs2.EOF Then
                mrs2.MoveFirst
                While Not mrs2.EOF
                    CodTrafico = mrs2!codtipotrafico ' vale cualquiera como representante
                    mSQL = "select sum(Importe),min(telefono) from tel_lin_factura_consumos"
                    mSQL = mSQL & " where serie = " & mdbFac2.texto(Serie)
                    mSQL = mSQL & " and ano = " & mdbFac2.numero(Ano)
                    mSQL = mSQL & " and numfact = " & mdbFac2.numero(NumFact2)
                    mSQL = mSQL & " and CodTipoTrafico = " & mdbFac2.texto(mrs2!codtipotrafico)
                    Set mrs3 = mdbFac2.cursor(mSQL)
                    If Not IsNull(mrs3.Fields(0)) Then
                        ImporteProv = ImporteProv + mrs3.Fields(0)
                        If UnoDeLosTelefonos = "" Then UnoDeLosTelefonos = mrs3.Fields(1)
                    End If
                    mrs3.Close
                    Set mrs3 = Nothing
                    mrs2.MoveNext
                Wend
            End If
            If ImporteProv Then
                DescuentoProv = buscaDescuentosGrupo(1, ImporteProv, CodTrafico, PorcentajeProv)
                If DescuentoProv <> 0 Then
                    '-- Hay descuento aplicable
                    mSQL = "insert into tel_lin_factura_descuentos values("
                    mSQL = mSQL & mdbFac2.texto(Serie) & "," ' Serie
                    mSQL = mSQL & mdbFac2.numero(Ano) & "," ' Año
                    mSQL = mSQL & mdbFac2.numero(NumFact2) & "," ' Número de factura propiamente dicho
                    mSQL = mSQL & mdbFac2.numero(siguienteLineaDeDescuentos(Serie, Ano, NumFact2)) & ","  ' Número de línea
                    '-- Busacamos la descripcion del descuento en: tel_desc_nombres_grupo
                    msql2 = "select * from tel_desc_nombres_grupo where grupo = " & mdbFac2.numero(mrs.Fields(0))
                    Set mrs3 = mdbFac2.cursor(msql2)
                    If Not mrs3.EOF Then
                        Aux = mrs3!Nombre
                    Else
                        Aux = "DESCUENTO AGRUPADO (" & CStr(mrs.Fields(0)) & ")"
                    End If
                    mSQL = mSQL & mdbFac2.texto(Aux) & "," ' Concepto
                    mSQL = mSQL & mdbFac2.numero(PorcentajeProv) & "," ' Pordentaje
                    mSQL = mSQL & mdbFac2.numero(ImporteProv) & "," ' Base
                    mSQL = mSQL & mdbFac2.numero(DescuentoProv) & "," & UnoDeLosTelefonos & ")" ' Importe
                    mdbFac2.ejecutar mSQL
                End If
            End If
            mrs2.Close
            Set mrs2 = Nothing
            mrs.MoveNext
        Wend
    End If
    mrs.Close
    Set mrs = Nothing
End Function

Private Function siguienteLineaDeDescuentos(Serie As String, Ano As Integer, NFac As Long) As Integer
    '-- siguienteLineaDeDescuentos:
    '   Devuelve el siguiente número de línea que corresponde a los parámetros pasados
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    mSQL = "select MAX(NumLin) from tel_lin_factura_descuentos"
    mSQL = mSQL & " where Serie = " & mdbFac2.texto(Serie)
    mSQL = mSQL & " and Ano = " & mdbFac2.numero(Ano)
    mSQL = mSQL & " and NumFact = " & mdbFac2.numero(NFac)
    Set mrs = mdbFac2.cursor(mSQL)
    If Not IsNull(mrs.Fields(0)) Then
        siguienteLineaDeDescuentos = mrs.Fields(0) + 1
    Else
        siguienteLineaDeDescuentos = 1
    End If
End Function

Private Function siguienteLineaDeEspeciales(Serie As String, Ano As Integer, NFac As Long) As Integer
    '-- siguienteLineaDeDescuentos:
    '   Devuelve el siguiente número de línea que corresponde a los parámetros pasados
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    mSQL = "select MAX(NumLin) from tel_lin_factura_especial"
    mSQL = mSQL & " where Serie = " & mdbFac2.texto(Serie)
    mSQL = mSQL & " and Ano = " & mdbFac2.numero(Ano)
    mSQL = mSQL & " and NumFact = " & mdbFac2.numero(NFac)
    Set mrs = mdbFac2.cursor(mSQL)
    If Not IsNull(mrs.Fields(0)) Then
        siguienteLineaDeEspeciales = mrs.Fields(0) + 1
    Else
        siguienteLineaDeEspeciales = 1
    End If
End Function


'




Private Function EstaElTelefono(Telefono As String) As Boolean
    '-- Comprueba si el teléfono pasado se encuentra en la base de datos
    '   de GesSocial, en caso afirmativo devuelve true, en caso contrario FALSE
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    'mSQL = "select * from telefonos where IdTelefono = " & mdbFac2.texto(Telefono)
    'DAVID
    mSQL = "select * from sclientfno where IdTelefono = " & mdbFac2.texto(Telefono)
    
    Set mrs = mdbFac2.cursor(mSQL)
    If Not mrs.EOF Then EstaElTelefono = True Else EstaElTelefono = False
End Function

Private Function NifDelTelefono(Telefono As String) As String
    Dim mSQL As String
    Dim mrs As ADODB.Recordset
    
    mSQL = "select * from telefonos where "
    'NUEVO DAVID
    mSQL = "select nifclien as NIF from sclien,sclientfno  where sclien.codclien=sclientfno.codclien  AND "
    
    
    mSQL = mSQL & " IdTelefono = " & mdbFac2.texto(Telefono)
    Set mrs = mdbFac2.cursor(mSQL)
    If Not mrs.EOF Then
        NifDelTelefono = mrs!NIF
    Else
        NifDelTelefono = ""
    End If
    mrs.Close
End Function




Public Function calculaInformeDescuentos(mfichero As String) As Boolean
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim i As Integer
    Dim i2 As Integer
    Dim C As String
    Dim Mc As String
    Dim Grupo As Integer
    Dim ID As Long
    Dim ID2 As Long
    Dim Tipo As String
    '--
    Dim cop_descripcion As String
    Dim cop_porcentaje As Double
    Dim cop_base As Double
    Dim cop_importe As Double
    Dim Reg As Integer
    '-- Lo primero comprobar coherencias
    If Not compruebaCoherenciaDescuentos Then
        MsgBox "La relacion entre descuentos de telefónica y los nuestros contiene incoherencias", vbCritical
        Exit Function
    End If
    '-- Este proceso actualiza las tablas tmp_inf_descuentos y tmp_inf_descuentos2
    '-- 1.- Nos basamos en los descuentos que nos hace telefónica, recorremos los
    '   registros y vamos dando de alta los correspondientes en tmp_inf_*
    '-- Borramos los valores hasta el momento
    SQL = "delete from tmp_inf_descuentos where fichero = " & mdbFac2.texto(mfichero)
    mdbFac2.ejecutar SQL
    SQL = "delete from tmp_inf_descuentos2 where fichero = " & mdbFac2.texto(mfichero)
    mdbFac2.ejecutar SQL
    '--
    SQL = "select * from telefono.descuentos_por_factura" & _
            " where fichero = " & mdbFac2.texto(mfichero) & _
            " and Codigo_de_descuento <> ''"
    Set RS = New ADODB.Recordset
    Set RS = mdbFac2.cursor(SQL)
    If Not RS.EOF Then
        Reg = 0
        RS.MoveFirst
        While Not RS.EOF
            Reg = Reg + 1
            '-- Ya podemos dar de alta nuestro registro
            ID = tmpSiguienteID()
            SQL = "insert into tmp_inf_descuentos(ID,fichero,tel_descripcion,tel_porcentaje,tel_base,tel_importe)"
            SQL = SQL & " values("
            SQL = SQL & mdbFac2.numero(ID) & ","
            SQL = SQL & mdbFac2.texto(mfichero) & ","
            SQL = SQL & mdbFac2.texto(RS!Descripcion_de_descuento) & ","
            SQL = SQL & mdbFac2.numero(RS!Porcentaje_descuento) & ","
            SQL = SQL & mdbFac2.numero(RS!Importe) & ","
            SQL = SQL & mdbFac2.numero(Abs(RS!Importe_Descuento)) & ")"
            mdbFac2.ejecutar SQL
            '-- Ahora buscamos las correspondencias para ver los descuentos que hacemos nosotros
            SQL = "select * from tel_relacion_descuentos" & _
                    " where coddesctel = " & mdbFac2.texto(RS!Codigo_de_descuento)
            Set rs2 = New ADODB.Recordset
            Set rs2 = mdbFac2.cursor(SQL)
            If Not rs2.EOF Then
                Grupo = 0
                i2 = 0
                Tipo = rs2!Tipo
                For i = 1 To Len(rs2!reldescpro)
                    C = Mid(rs2!reldescpro, i, 1)
                    If C = "/" Then
                        Select Case Tipo
                            Case "T"
                                '-- buscammos si pertenece a un grupo y a cual
                                SQL = "select * from tel_desc_consumos" & _
                                        " where CodTipoTrafico = " & mdbFac2.texto(Mc)
                                Set rs3 = New ADODB.Recordset
                                Set rs3 = mdbFac2.cursor(SQL)
                                If Not rs3.EOF Then Grupo = rs3!Grupo Else Grupo = 0
                            Case "C"
                                Grupo = 0
                        End Select
                        If Grupo = 0 Then
                            If Tipo = "T" Then
                                '-- No forma parte de un grupo lo calculamos individual
                                SQL = "select * from telefono.detalle_de_llamadas " & _
                                        " where fichero = " & mdbFac2.texto(mfichero) & _
                                        " and Codigo_de_trafico = " & mdbFac2.texto(Mc)
                                Set rs3 = New ADODB.Recordset
                                Set rs3 = mdbFac2.cursor(SQL)
                                If Not rs3.EOF Then cop_descripcion = rs3!Tipo_de_trafico _
                                    Else cop_descripcion = ""
                            Else
                                '-- es una cuota
                                SQL = "select * from telefono.cuotas " & _
                                        " where fichero = " & mdbFac2.texto(mfichero) & _
                                        " and Codigo_de_cuota = " & mdbFac2.texto(Mc)
                                Set rs3 = New ADODB.Recordset
                                Set rs3 = mdbFac2.cursor(SQL)
                                If Not rs3.EOF Then cop_descripcion = rs3!Descripcion_de_cuota _
                                    Else cop_descripcion = ""
                            End If
                        Else
                            '-- Forma parte de un grupo, buscamos el nombre del grupo
                            SQL = "select * from tel_desc_nombres_grupo where grupo = " & mdbFac2.numero(Grupo)
                            Set rs3 = mdbFac2.cursor(SQL)
                            If Not rs3.EOF Then cop_descripcion = rs3!Nombre _
                                Else cop_descripcion = ""
                            '-- Ahora confirmamos que ese grupo no se ha procesado ya
                            If cop_descripcion <> "" Then
                                SQL = "select * from tmp_inf_descuentos" & _
                                        " where ID = " & mdbFac2.numero(ID) & _
                                        " and cop_descripcion = " & mdbFac2.texto(cop_descripcion)
                                Set rs3 = mdbFac2.cursor(SQL)
                                If Not rs3.EOF Then cop_descripcion = ""
                                SQL = "select * from tmp_inf_descuentos2" & _
                                        " where ID = " & mdbFac2.numero(ID) & _
                                        " and cop_descripcion = " & mdbFac2.texto(cop_descripcion)
                                Set rs3 = mdbFac2.cursor(SQL)
                                If Not rs3.EOF Then cop_descripcion = ""
                            End If
                        End If
                        '-- Tanto si forma parte de un grupo como si no nos movemos por descripciones
                        If cop_descripcion <> "" Then
                            '-- Ahora a buscar importes de descuentos
                            SQL = "select a.descuento, sum(a.Base), sum(a.Importe)" & _
                                    " from tel_lin_factura_descuentos as a, tel_cab_factura as b" & _
                                    " where b.fichero = " & mdbFac2.texto(mfichero) & _
                                    " and a.concepto = " & mdbFac2.texto(cop_descripcion) & _
                                    " and a.serie = b.serie and a.ano = b.ano and a.numfact = b.numfact" & _
                                    " group by a.descuento"
                             Set rs3 = mdbFac2.cursor(SQL)
                             If Not rs3.EOF Then
                                rs3.MoveFirst
                                While Not rs3.EOF
                                    cop_porcentaje = rs3.Fields(0)
'                                    If i2 = 0 Then
'                                        sql = "update tmp_inf_descuentos set"
'                                        sql = sql & " cop_descripcion = " & mdbFac.Texto(cop_descripcion) & ","
'                                        sql = sql & " cop_porcentaje = " & mdbFac.numero(cop_porcentaje) & ","
'                                        sql = sql & " cop_base = " & mdbFac.numero(rs3.Fields(1)) & ","
'                                        sql = sql & " cop_importe = " & mdbFac.numero(rs3.Fields(2))
'                                        sql = sql & " where ID = " & mdbFac.numero(ID)
'                                    Else
                                        ID2 = tmpsiguienteID2(ID)
                                        SQL = "insert into tmp_inf_descuentos2(ID,ID2,fichero,cop_descripcion,cop_porcentaje,cop_base,cop_importe)"
                                        SQL = SQL & " values("
                                        SQL = SQL & mdbFac2.numero(ID) & ","
                                        SQL = SQL & mdbFac2.numero(ID2) & ","
                                        SQL = SQL & mdbFac2.texto(mfichero) & ","
                                        SQL = SQL & mdbFac2.texto(cop_descripcion) & ","
                                        SQL = SQL & mdbFac2.numero(cop_porcentaje) & ","
                                        SQL = SQL & mdbFac2.numero(rs3.Fields(1)) & ","
                                        SQL = SQL & mdbFac2.numero(rs3.Fields(2)) & ")"
'                                    End If
                                    mdbFac2.ejecutar SQL
                                    i2 = i2 + 1
                                    rs3.MoveNext
                                Wend
                             End If
                        End If
                        Mc = ""
                    Else
                        Mc = Mc & C
                    End If
                Next i
            End If
            RS.MoveNext
        Wend
    End If
    '-- Ya están actualizados todos los registros, ahora vamos a por los totales
    Dim tot_cop_base As Currency
    Dim tot_cop_importe As Currency
    SQL = "select sum(cop_base), sum(cop_importe), ID from tmp_inf_descuentos2"
    SQL = SQL & " where fichero = " & mdbFac2.texto(mfichero)
    SQL = SQL & " group by ID"
    Set RS = mdbFac2.cursor(SQL)
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            SQL = "update tmp_inf_descuentos set "
            SQL = SQL & " cop_descripcion = 'TOTAL'" & ","
            SQL = SQL & " cop_base = " & mdbFac2.numero(RS.Fields(0)) & ","
            SQL = SQL & " cop_importe = " & mdbFac2.numero(RS.Fields(1))
            SQL = SQL & " where ID = " & mdbFac2.numero(RS.Fields(2))
            SQL = SQL & " and fichero = " & mdbFac2.texto(mfichero)
            mdbFac2.ejecutar SQL
            RS.MoveNext
        Wend
    End If
    '-- Por último los nombrecitos a ver si salen bien.
    Carga_Inf_Nombres mfichero
End Function

Private Function compruebaCoherenciaDescuentos() As Boolean
    '-- Esta function comprueba si las relaciones parametrizadas en relacion_descuentos
    '   son correctas, fundamentalmente si el descuento pertenece a un grupo que todos
    '   los tráficos asignados a ese grupo estan en el mismo registro de relación
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Dim C As String
    Dim Mc As String
    Dim Grupo As Integer
    '-- Por defecto consideramos que las cosas van bien
    compruebaCoherenciaDescuentos = True
    Grupo = 0
    '--
    SQL = "select * from tel_relacion_descuentos"
    Set RS = New ADODB.Recordset
    Set RS = mdbFac2.cursor(SQL)
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            If RS!Tipo = "T" Then ' solo tiene sentido en el caso de tráfico
                '-- descomponemos la relación de tipos de tráfico asociada
                For i = 1 To Len(RS!reldescpro)
                    C = Mid(RS!reldescpro, i, 1)
                    If C = "/" Then
                        '-- buscammos si pertenece a un grupo y a cual
                        SQL = "select * from tel_desc_consumos" & _
                                " where CodTipoTrafico = " & mdbFac2.texto(RS!coddesctel)
                        Set rs2 = New ADODB.Recordset
                        Set rs2 = mdbFac2.cursor(SQL)
                        If Not rs2.EOF Then
                            If (Grupo <> 0) And (rs2!Grupo <> Grupo) Then
                                compruebaCoherenciaDescuentos = False
                            End If
                            Grupo = rs2!Grupo
                        End If
                        Mc = ""
                    Else
                        Mc = Mc & C
                    End If
                Next i
            End If
            RS.MoveNext
        Wend
    End If
    Set RS = Nothing
    Set rs2 = Nothing
End Function

Private Function tmpSiguienteID() As Long
    Dim RS As ADODB.Recordset
    Dim SQL As String
    SQL = "select max(ID) from tmp_inf_descuentos"
    Set RS = New ADODB.Recordset
    Set RS = mdbFac2.cursor(SQL)
    If Not IsNull(RS.Fields(0)) Then
        tmpSiguienteID = RS.Fields(0) + 1
    Else
        tmpSiguienteID = 1
    End If
    RS.Close
    Set RS = Nothing
End Function

Private Function tmpsiguienteID2(ID As Long) As Long
    Dim RS As ADODB.Recordset
    Dim SQL As String
    SQL = "select max(ID2) from tmp_inf_descuentos2" & _
            " where ID = " & mdbFac2.numero(ID)
    Set RS = New ADODB.Recordset
    Set RS = mdbFac2.cursor(SQL)
    If Not IsNull(RS.Fields(0)) Then
        tmpsiguienteID2 = RS.Fields(0) + 1
    Else
        tmpsiguienteID2 = 1
    End If
    RS.Close
    Set RS = Nothing
End Function
    
Private Sub Carga_Inf_Nombres(mfichero As String)
    Dim RS As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim SQL As String
    Dim C As String
    Dim Mc As String
    Dim i As Integer
    Dim n1 As String
    Dim N2 As String
    Dim Importe As Currency
    '--
    SQL = "delete from tmp_inf_nombres where fichero = " & mdbFac2.texto(mfichero)
    mdbFac2.ejecutar SQL
    '-- leemos secuencialmente los valores
    SQL = "select * from tel_relacion_nombres"
    Set RS = New ADODB.Recordset
    Set RS = mdbFac2.cursor(SQL)
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            SQL = "select * from telefono.descuentos_por_factura" & _
                    " where fichero = " & mdbFac2.texto(mfichero) & _
                    " and Codigo_de_descuento = " & mdbFac2.texto(RS!coddesctel)
            Set rs2 = New ADODB.Recordset
            Set rs2 = mdbFac2.cursor(SQL)
            If Not rs2.EOF Then n1 = rs2!Descripcion_de_descuento Else n1 = ""
            For i = 1 To Len(RS!reldescpro)
                C = Mid(RS!reldescpro, i, 1)
                If C = "/" Then
                    Select Case RS!Tipo
                        Case "T"
                            SQL = "select * from telefono.detalle_de_llamadas " & _
                                    " where fichero = " & mdbFac2.texto(mfichero) & _
                                    " and Codigo_de_trafico = " & mdbFac2.texto(Mc)
                            Set rs3 = New ADODB.Recordset
                            Set rs3 = mdbFac2.cursor(SQL)
                            If Not rs3.EOF Then N2 = rs3!Tipo_de_trafico _
                                Else N2 = ""
                        Case "C"
                            SQL = "select * from telefono.cuotas " & _
                                    " where fichero = " & mdbFac2.texto(mfichero) & _
                                    " and Codigo_de_cuota = " & mdbFac2.texto(Mc)
                            Set rs3 = New ADODB.Recordset
                            Set rs3 = mdbFac2.cursor(SQL)
                            If Not rs3.EOF Then N2 = rs3!Descripcion_de_cuota _
                                Else N2 = ""
                    End Select
                    '-- Y ahora damos de alta el registro correspondiente
                    If N2 <> "" And Not IsNull(N2) Then
                        If RS!Tipo = "T" Then
                            SQL = "select sum(a.Importe)" & _
                                   " from tel_lin_factura_consumos as a, tel_cab_factura as b" & _
                                   " where b.fichero = " & mdbFac2.texto(mfichero) & _
                                   " and a.DescTipoTrafico = " & mdbFac2.texto(N2) & _
                                   " and a.serie = b.serie and a.ano = b.ano and a.numfact = b.numfact"
                        Else
                            SQL = "select sum(a.Importe)" & _
                                   " from tel_lin_factura_cuotas as a, tel_cab_factura as b" & _
                                   " where b.fichero = " & mdbFac2.texto(mfichero) & _
                                   " and a.DescCuota = " & mdbFac2.texto(N2) & _
                                   " and a.serie = b.serie and a.ano = b.ano and a.numfact = b.numfact"
                        End If
                        Set rs3 = mdbFac2.cursor(SQL)
                        If Not IsNull(rs3.Fields(0)) Then
                            Importe = rs3.Fields(0)
                        Else
                            Importe = 0
                        End If
                        SQL = "insert into tmp_inf_nombres(fichero,n1,n2,importe)"
                        SQL = SQL & " values("
                        SQL = SQL & mdbFac2.texto(mfichero) & ","
                        SQL = SQL & mdbFac2.texto(n1) & ","
                        SQL = SQL & mdbFac2.texto(N2) & ","
                        SQL = SQL & mdbFac2.numero(Importe) & ")"
                        mdbFac2.ejecutar SQL
                    End If
                    N2 = ""
                    Mc = ""
                Else
                    Mc = Mc & C
                End If
            Next i
            n1 = ""
            RS.MoveNext
        Wend
    End If
End Sub




'ByVal para que pueda modificar el valor si modificar en origen
Private Sub SeparaNomclienNomApellidos(ByVal NomClien As String, nom As String, Ap1 As String, Ap2 As String)
Dim i As Integer
Dim TieneComa As Boolean

    i = InStr(1, NomClien, ",")
    
    TieneComa = i > 0
    
    If TieneComa Then
        'El nombre va detras de la coma
        nom = Mid(NomClien, i + 1)
        NomClien = Mid(NomClien, 1, i - 1)
        
        
        i = InStrRev(NomClien, " ")
        If i = 0 Then
            Ap1 = NomClien
            Ap2 = ""
        Else
            Ap1 = Mid(NomClien, 1, i - 1)
            Ap2 = Mid(NomClien, i + 1)
        End If
        
        
    Else
        
        Ap1 = ""
        Ap2 = ""
        
        i = Len(NomClien) / 3
        If i = 0 Then i = 1
        i = InStr(i, NomClien, " ")
        
        If i = 0 Then
            i = InStr(1, NomClien, " ")
            If i = 0 Then
                NomClien = "NO encontrado"
                Exit Sub
            End If
        End If
        
        nom = Mid(NomClien, 1, i - 1)
        NomClien = Trim(Mid(NomClien, i + 1))
        
        i = InStr(5, NomClien, " ")
        If i = 0 Then
            Ap1 = NomClien
            Ap2 = ""
        Else
            Ap1 = Mid(NomClien, 1, i - 1)
            Ap2 = Mid(NomClien, i + 1)
        End If
        
        
    End If
    
    
End Sub







Public Function RecalcularImporteLlamadasCoperativa(ByRef L As Label, Operadora As Byte)
    Dim RS As ADODB.Recordset
    Dim cad As String
    Dim Aux As String
    Dim Impor As Currency
    Dim imp2 As Currency
    Dim Aux2 As String
    Dim i As Integer
    
    'Entran con espacios en blanco
    L.Caption = "Ajustando numeros llamados"
    L.Refresh
    cad = "update telefono.detalle_de_llamadas set Numero_llamado =trim(Numero_llamado) where fichero='" & mFicheroCorto & "'"
    conn.Execute cad
    Espera 0.25
    
    'LLAMADAS DE COOPERATIVA
    Set RS = New ADODB.Recordset
    
    If Operadora = 1 Then
        'SOLO PARA MOVISTAR. Las llamadas entre telefonos en ORANGE, ya vienen a 0
    
     L.Caption = "Leyendo registros llamada"
     L.Refresh
     If mvarConceptosQueSonLlamadas <> "" Then
         cad = "select Numero_llamado from telefono.detalle_de_llamadas where fichero='" & mFicheroCorto & "' and Numero_llamado<>'' and"
         cad = cad & " Numero_llamado in (select idtelefono from sclientfno where operador = " & Operadora & ") group by 1"
         Set RS = New ADODB.Recordset
         RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         While Not RS.EOF
             L.Caption = RS.Fields(0)
             L.Refresh
             
             
             'Primero el importl
             cad = "update telefono.detalle_de_llamadas set libre=importe"
             cad = cad & " WHERE Numero_llamado ='" & RS.Fields(0) & "' AND fichero='" & mFicheroCorto & "'"
             cad = cad & " and importe<>0 AND libre=''"
             cad = cad & " and Codigo_de_trafico in (" & mvarConceptosQueSonLlamadas & ")"
             conn.Execute cad
           
           
             'Me estoy guardando el importe anterior en LIBRE
             cad = "update telefono.detalle_de_llamadas set importe=0"
             cad = cad & " WHERE Numero_llamado ='" & RS.Fields(0) & "' AND fichero='" & mFicheroCorto & "'"
             cad = cad & " and importe<>0 "
             cad = cad & " and Codigo_de_trafico in (" & mvarConceptosQueSonLlamadas & ")"
             conn.Execute cad
             
             
             
         
             RS.MoveNext
         Wend
         RS.Close
         
     End If
    End If
    
    
   'Ajuste cuotas que no se facturan
   DoEvents

   
   'Borraremos el resumen de llamada para los conceptos refacturables
   If mvarConceptosLlamadasRefacturables <> "" Then
        DoEvents
        L.Caption = "Eliminar datos recfac."
        L.Refresh
        
        'OK, tiene para refacturar
        'Valores por defecto. Para poder ordenar debe ser numerico el campo del where
        Aux = "Fichero =" & DBSet(mFicheroCorto, "T") & " AND 1"
        cad = "concat(numero_de_factura,'|',min(Fecha_inicio_periodo),'|',max(Fecha_final_periodo),'|')"
        cad = DevuelveDesdeBD(conAri, cad, "telefono.resumen_de_llamadas", Aux, "1 GROUP BY  numero_de_factura", "N")
        'Me lo guardo
        
        
        'Faltaria ver que si el concepto NO es sexagesimal(horas), la suma va de otra forma
        'falta ver si la suma de llamadas que va en "sexagesimal" y la suma de mb que va en decimal la hace bien
        
        Aux = "DELETE FROM telefono.resumen_de_llamadas WHERE Fichero =" & DBSet(mFicheroCorto, "T") & " AND "
        Aux = Aux & " Codigo_de_tipo_de_trafico in (" & mvarConceptosLlamadasRefacturables & ")"
        mdbFac2.ejecutar Aux
   
        
        'Las sumas se haran de dos formas.
        'Si son mesajes o trafico va sum(importe y duracion) diratemente
        'pero si son Segundos, el sum es distinto
        If mvarConceptosNOSonLlamadas <> "" Then
            
            Aux = "INSERT INTO telefono.resumen_de_llamadas (Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura,"
            Aux = Aux & "Numero_de_prefactura,Numero_de_telefono,Numero_de_extension,Fecha_inicio_periodo,"
            Aux = Aux & "Fecha_final_periodo,Codigo_de_tipo_de_trafico,Tipo_de_trafico,Numero_de_llamadas,"
            Aux = Aux & "Unidad_de_medida,Cantidad_medida,Cantidad_franquiciada,Importe_franquiciado,Importes,Libre)"
            
            Aux = Aux & "  SELECT '" & mFicheroCorto & "',320,1,'" & RecuperaValor(cad, 1) & "',''"
            Aux = Aux & ",Numero_de_telefono,'','" & RecuperaValor(cad, 2) & "','" & RecuperaValor(cad, 3) & "',"
            Aux = Aux & "Codigo_de_trafico,Tipo_de_trafico,count(*),Unidad_de_medida,sum(Cantidad_medida_originada)"
            Aux = Aux & ",0,0,sum(importe),'' FROM telefono.detalle_de_llamadas WHERE  Fichero =" & DBSet(mFicheroCorto, "T")
            Aux = Aux & " and Codigo_de_trafico in (" & mvarConceptosNOSonLlamadas & ")  and fecha<>'0000'"
            Aux = Aux & " group by Codigo_de_trafico,Numero_de_telefono order by Numero_de_telefono,Codigo_de_trafico"
            mdbFac2.ejecutar Aux
        
        End If
        
        
        
        If mvarConceptosQueSonLlamadas <> "" Then
            'La suma de la cantidad es distinta.
            '2.50 + 2.50 no son 5
            ' son 2 + 2 = 4
            '   50 + 50=100
            '   100= 1min 40 ->>  total 5 +40
            L.Caption = "Calculo resumen consumo"
            L.Refresh
        
            Aux = "  SELECT '" & mFicheroCorto & "',320,1,'" & RecuperaValor(cad, 1) & "',''"
            Aux = Aux & ",Numero_de_telefono,'','" & RecuperaValor(cad, 2) & "','" & RecuperaValor(cad, 3) & "',"
            Aux = Aux & "Codigo_de_trafico,Tipo_de_trafico,count(*),Unidad_de_medida,"
            'sum(Cantidad_medida_originada)"
            'Los segundos se suma distinto
            'Aux = Aux & "(floor(Cantidad_medida_originada) *60) + ((Cantidad_medida_originada - floor(Cantidad_medida_originada))*100) as calculados"
            Aux2 = "sum(floor(Cantidad_medida_originada) * 60 +(Cantidad_medida_originada-floor(Cantidad_medida_originada)) *100 ) as Calculados"

            Aux = Aux & Aux2
            
            Aux = Aux & ",0,0,sum(importe) totimpor,'' FROM telefono.detalle_de_llamadas WHERE  Fichero =" & DBSet(mFicheroCorto, "T")
            Aux = Aux & " and Codigo_de_trafico in (" & mvarConceptosQueSonLlamadas & ")  and fecha<>'0000'"
            Aux = Aux & " group by Codigo_de_trafico,Numero_de_telefono order by Numero_de_telefono,Codigo_de_trafico"
            RS.Open Aux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Aux = "INSERT INTO telefono.resumen_de_llamadas (Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura,"
            Aux = Aux & "Numero_de_prefactura,Numero_de_telefono,Numero_de_extension,Fecha_inicio_periodo,"
            Aux = Aux & "Fecha_final_periodo,Codigo_de_tipo_de_trafico,Tipo_de_trafico,Numero_de_llamadas,"
            Aux = Aux & "Unidad_de_medida,Cantidad_medida,Cantidad_franquiciada,Importe_franquiciado,Importes,Libre)"
            
            cad = "" 'cuidado. Esta variable estaba para el numero de factura telefonia, fechainic y fin
            While Not RS.EOF
                Aux2 = ", ("
                For i = 0 To 12
                    Aux2 = Aux2 & DBSet(RS.Fields(i), "T", "N") & ","
                Next
                            
                Impor = DBLet(RS!Calculados, "N")
                imp2 = Impor \ 60
                Impor = Impor - (imp2 * 60)
                Impor = imp2 + (Round(Impor / 100, 2))
                    
                    
                Aux2 = Aux2 & DBSet(Impor, "N")
                
                ',0,0,sum(importe),''
                Aux2 = Aux2 & ",0,0," & DBSet(RS!totimpor, "N") & ",''"
                
                cad = cad & Aux2 & ")"
                
                
                If Len(cad) > 25000 Then
                    cad = Mid(cad, 2)
                    cad = Aux & " VALUES " & cad
                    conn.Execute cad
                    cad = ""
                End If
                RS.MoveNext
            Wend
            RS.Close
            
            If cad <> "" Then
                cad = Mid(cad, 2)
                cad = Aux & " VALUES " & cad
                conn.Execute cad
            
            End If
        End If
    End If
    
    
    
    'En bolbaite. Ya vermeos como lo parametrizo
    'pOr si acaso
    If vParamAplic.TieneTelefonia2 = 3 Then
        If Operadora = 1 Then
            L.Caption = "Concepto W1"
            L.Refresh

            cad = "UPDATE telefono.resumen_de_llamadas set codigo_de_tipo_de_trafico='20' WHERE Fichero =" & DBSet(mFicheroCorto, "T")
            
            cad = cad & " AND codigo_de_tipo_de_trafico='W1'"
            conn.Execute cad
            
        End If
    End If
    
    
    
    
    'Cargos/Cuotas de varios
    'tel_cargo_varios
    
    
    
    
    
    
    
    Set RS = Nothing
    
    
End Function



'TipoFicheroOperador
'   1. MOV    2.  ORAN  ....   3VODA
Public Function AjusteCuotasNuevas2(TipoFicheroOperador As Byte, mfichero As String, ByRef L As Label) As Boolean
Dim vSQL As String
Dim cad As String
Dim Origen As String
Dim Rc As ADODB.Recordset
Dim i As Integer
Dim Importe2 As Currency
Dim Importe3 As Currency
Dim idC As Long

    AjusteCuotasNuevas2 = False

    L.Caption = "Ajuste quotas"
    L.Refresh
    
    'Valores por defecto. Para poder ordenar debe ser numerico el campo del where
    If TipoFicheroOperador = 2 Then mfichero = mFicheroCorto
    
    If TipoFicheroOperador = 3 Then
        'Vodafone refacturado
        vSQL = DevuelveDesdeBD(conAri, "vodafoneRefacturado", "spara2", "1", "1")
        If Val(vSQL) = 1 Then
            mfichero = mFicheroCorto
            mvarFichero = mFicheroCorto
        End If
    End If
    
    vSQL = "Fichero =" & DBSet(mfichero, "T") & " AND 1"
    
    

        'Resto operadores
        cad = "concat(numero_de_factura,'|',min(Fecha_inicio_cuotas),'|',min(Fecha_final_cuotas),'|')"
        Origen = DevuelveDesdeBD(conAri, cad, "telefono.cuotas", vSQL, "1 GROUP BY  numero_de_factura", "N")
        If Origen = "" Then
            'Probamos en detalle lamadas
             
            vSQL = "Fichero =" & DBSet(mfichero, "T") & " AND 1"
        
            cad = "concat(numero_de_factura,'|',min(Fecha_inicio_periodo),'|',min(Fecha_final_periodo),'|')"
            Origen = DevuelveDesdeBD(conAri, cad, "telefono.resumen_de_llamadas", vSQL, "1 GROUP BY  numero_de_factura", "N")
        End If

    L.Caption = "Ajuste cuotas personalizable (I)"
    L.Refresh
    
    Set Rc = New ADODB.Recordset
    
    'Vemos las acciones tipo1.   ELIMINAR
    cad = DevuelveACcionesTipo(1, 0, Rc, TipoFicheroOperador)
    If cad <> "" Then
        cad = "DELETE FROM telefono.cuotas WHERE Fichero =" & DBSet(mfichero, "T") & " AND Codigo_de_cuota IN (" & cad & ")"
        conn.Execute cad
    End If
    
    L.Caption = "Ajuste cuotas personalizable (II)"
    L.Refresh
    'Acciones tipo 2, cuota mes entera
    '  tipo: 0. Se cobra el mes entero.

    
    cad = "Select * from tel_descuentoscuotas WHERE accion=2 AND Mensual =0 " 'todo el mes
    cad = cad & " AND operadora =" & TipoFicheroOperador
'    vSQL = "1"
'    If EsFicheroOrange_ Then vSQL = "2"
'    Cad = Cad & vSQL
    
    Rc.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not Rc.EOF
        cad = TransformaComasPuntos(CStr(Rc!Valor))
        cad = "UPDATE telefono.cuotas SET importe = " & cad
        cad = cad & " WHERE Fichero =" & DBSet(mfichero, "T") & " AND Codigo_de_cuota = " & DBSet(Rc!codigo_de_cuota, "T")
        conn.Execute cad
        Rc.MoveNext
    Wend
    Rc.Close
    
    
     'Acciones tipo 2, cuota mes entera
    '  tipo: 1.  Se prorratea los dias,
    L.Caption = "Ajuste cuotas personalizable (III)"
    L.Refresh
    cad = DevuelveACcionesTipo(2, 1, Rc, TipoFicheroOperador)
    If cad <> "" Then
        MsgBox "FALTA acciones 2. Cuota mes entera. No deberia haber ocurrido.", vbCritical
        MsgBox "Avise a Ariadna !!!!!!", vbCritical
        Exit Function
    End If
    
    
    'Acciones tipo 2, cuota mes entera
    '  tipo: 1.  Se prorratea los dias,
    L.Caption = "Ajuste cuotas personalizable (III)"
    L.Refresh
    cad = DevuelveACcionesTipo(3, 1, Rc, TipoFicheroOperador)
    If cad <> "" Then
        L.Caption = "Prorrateo de cuotas"
        L.Refresh
        
        'CUOTAS PRORRATEABLES pro fecha de baja u alta
        
        vSQL = "select Codigo_de_cuota,Numero_de_telefono,datediff(Fecha_final_cuotas,Fecha_inicio_cuotas) dias,Importe from telefono.cuotas"
        vSQL = vSQL & " WHERE Fichero =" & DBSet(mfichero, "T") & " AND Codigo_de_cuota IN (" & cad & ") ORDER BY 1"
        Rc.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        vSQL = "" 'Aqui llevare la cuota
        Importe2 = 0 '% para guardar
        While Not Rc.EOF
            
            If vSQL <> Rc!codigo_de_cuota Then
                vSQL = DevuelveDesdeBD(conAri, "Valor", "tel_descuentoscuotas", "codigo_de_cuota", CStr(Rc!codigo_de_cuota), "T")
                If vSQL = "" Then vSQL = "100"
                Importe2 = CCur(vSQL)
                    
                vSQL = Rc!codigo_de_cuota
            End If
            i = Rc!Dias
            If i > 30 Then i = 30
            If i < 0 Then i = 0
            i = 30 'El % de dias ya lo trae la cuota
            Importe3 = ((i / 30) * Rc!Importe * Importe2) / 100
            Importe3 = Round(Importe3, 2)
            
            
            
            cad = "UPDATE telefono.cuotas SET libre=" & TransformaComasPuntos(CStr(Rc!Importe))
            cad = cad & ",importe = " & TransformaComasPuntos(CStr(Importe3))
            cad = cad & " WHERE Fichero =" & DBSet(mfichero, "T") & " AND Codigo_de_cuota = " & DBSet(Rc!codigo_de_cuota, "T")
            cad = cad & " AND Numero_de_telefono =" & DBSet(Rc!Numero_de_telefono, "T")
            conn.Execute cad
            Rc.MoveNext
        Wend
        Rc.Close
        
        

        
        
        
    End If
    
    
    'INSERTAMOS LAS QUE NO ESTAN
    L.Caption = "insertar cuotas personalizable nuevas"
    L.Refresh
    
    'Primero comoprobaré si hay inserciones nuevas
    'operadora=1 MOVISTAR
'    vSQL = "1"
'    If EsFicheroOrange_ Then vSQL = "2"
    vSQL = "operadora = " & TipoFicheroOperador
    
    cad = "Fichero =" & DBSet(mfichero, "T") & " and not codigo_de_cuota in (select codigo_de_cuota from tel_descuentoscuotas where " & vSQL & ") and 1"
    cad = DevuelveDesdeBD(conAri, "count(*)", "telefono.cuotas", cad, "1", "N")
    If cad <> "" Then
        If Val(cad) > 0 Then cad = ""
    End If
    If cad = "" Then
        cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2) "
        cad = cad & " select " & vUsu.Codigo & ",1,0,Codigo_de_cuota,Descripcion_de_cuota"
        cad = cad & "  from telefono.cuotas where Fichero =" & DBSet(mfichero, "T") & " group by Codigo_de_cuota"
        conn.Execute cad
        
        'vSQL = "1"
        'If EsFicheroOrange_ Then vSQL = "2"
        vSQL = TipoFicheroOperador
        cad = "insert ignore into tel_descuentoscuotas(operadora,codigo_de_cuota,Descripcion_de_cuota,accion,Mensual,Valor)"
        cad = cad & " select " & vSQL & ",Codigo_de_cuota,Descripcion_de_cuota,0,0,sum(importe)/count(*) from telefono.cuotas "
        cad = cad & " where Fichero =" & DBSet(mfichero, "T") & " group by Codigo_de_cuota"
        conn.Execute cad
                    
    End If
    

    'VARIOS
    L.Caption = "Cargos VARIOS"
    L.Refresh
    cad = "Fichero =" & DBSet(mfichero, "T") & " and Codigo_de_vario<>'' and not Codigo_de_vario "
    'vSQL = "1"
    'If EsFicheroOrange_ Then vSQL = "2"
    vSQL = "operadora = " & TipoFicheroOperador '& vSQL
    cad = cad & " in (select CodigoVario from tel_cargo_varios where " & vSQL & ") and 1"
    cad = DevuelveDesdeBD(conAri, "count(*)", "telefono.varios", cad, "1", "N")
    If cad <> "" Then
        If Val(cad) > 0 Then cad = ""
    End If
    If cad = "" Then
        'Codigo_de_vario Descripcion_de_vario

        cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2) "
        cad = cad & " select " & vUsu.Codigo & ",2,0,Codigo_de_vario,Descripcion_de_vario"
        cad = cad & "  from telefono.varios where Fichero =" & DBSet(mfichero, "T")
        cad = cad & " AND Codigo_de_vario<>"""" AND Importe<>0 group by Codigo_de_vario"
        conn.Execute cad
        

        cad = "insert ignore tel_cargo_varios(CodigoVario,Descripcion,CargarEnFactura,Operadora)"
        cad = cad & " select Codigo_de_vario,Descripcion_de_vario,0,1 from telefono.varios "
        cad = cad & " where Codigo_de_vario<>'' AND importe<>0"
        cad = cad & " AND Fichero =" & DBSet(mfichero, "T") & " group by Codigo_de_vario"
        conn.Execute cad
        
    End If


    DoEvents
    L.Caption = "Lineas duplicadas en fichero"
    L.Refresh


    'Abril 2020
    'Para que no de error NO controlado
    cad = "select Numero_de_telefono,count(*) from telefono.telefono where Fichero = " & mdbFac2.texto(mfichero)
    cad = cad & " group by Numero_de_telefono having count(*)>1"
    Rc.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not Rc.EOF
        cad = cad & vbCrLf & Rc.Fields(0) & "    " & Rc.Fields(1)
        Rc.MoveNext
    Wend
    Rc.Close
    If cad <> "" Then
        cad = L.Caption & vbCrLf & cad
        MsgBox cad, vbExclamation
        L.Caption = ""
        Set Rc = Nothing
        Exit Function
    End If
    L.Caption = "Telefonos sin consumo / inactivos / err. operadora / Plazos"
    L.Refresh
    
    
    
    'Despues de ajustar las cuotas, mostraremos los telefonos que no tienen CUOTAS
    'Por si quiere quitar alguno. Para que no le lleguen la factura
    'Primero. Metemos todos los telelfonos del fichero
    cad = "DELETE FROM tmpnseries WHERE codusu = " & vUsu.Codigo
    conn.Execute cad
    Espera 0.15
    
    
    
  
    
    
    
    
    
    'INACTIVOS. Pondremos numlinea =2
    cad = "select " & vUsu.Codigo & ",Numero_de_telefono,0 from telefono.telefono" & _
            " where Fichero = " & mdbFac2.texto(mfichero) & " order by Numero_de_telefono"
    cad = "INSERT INTO tmpnseries(codusu,codartic,numlinea) " & cad
    conn.Execute cad
    
    
    Espera 0.25
    
    'INACTIVOS. Pondremos numlinea =2
    cad = "select * from tmpnseries where codusu =" & vUsu.Codigo & " and codartic in (select IdTelefono from sclientfno where inactivo=1)"
    Rc.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rc.EOF
        cad = "UPDATE tmpnseries set numlinea= 2 where codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(Rc!codArtic, "T")
        Rc.MoveNext
        ejecutar cad, False
        Espera 0.1
    Wend
    Rc.Close
    
    'Febrero 2014
    'Los que en la fiche del telefono, el operador que llevan es OTRO que el del fichero
    cad = "select * from tmpnseries where codusu =" & vUsu.Codigo & " and codartic in (select IdTelefono from sclientfno where inactivo=0 and operador <>" & TipoFicheroOperador & " )"
    Rc.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rc.EOF
        cad = "UPDATE tmpnseries set numlinea= 5 where codusu = " & vUsu.Codigo & " AND codartic = " & DBSet(Rc!codArtic, "T")
        Rc.MoveNext
        ejecutar cad, False
        Espera 0.1
    Wend
    Rc.Close
    
    
    'Borramos los que tienen detalle _llamadas , cutoas y varios  FEB 2014 -->
    cad = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.Codigo & " AND numlinea<2 AND codartic IN ("
    vSQL = "Select distinct(Numero_de_telefono) FROM telefono.cuotas WHERE Fichero = " & mdbFac2.texto(mfichero)
    vSQL = cad & vSQL & ")"
    conn.Execute vSQL
    vSQL = "Select distinct(Numero_de_telefono) FROM telefono.resumen_de_llamadas  n WHERE Fichero = " & mdbFac2.texto(mfichero)
    vSQL = cad & vSQL & ")"
    conn.Execute vSQL
    vSQL = "Select distinct(Numero_de_telefono) FROM telefono.varios WHERE Fichero = " & mdbFac2.texto(mfichero)
    vSQL = cad & vSQL & ")"
    conn.Execute vSQL
    'vodafone
    If TipoFicheroOperador = 3 Then
        vSQL = "Select distinct(cuenta_no) FROM vodafonetxt.fac_cab WHERE fac_no = " & mdbFac2.texto(mfichero)
        vSQL = cad & vSQL & ")"
        conn.Execute vSQL
    End If
    
    i = 0
    vSQL = DevuelveDesdeBD(conAri, "count(*)", "tmpnseries", "codusu", CStr(vUsu.Codigo))
    If Val(vSQL) > 0 Then i = 1
    
    If i = 0 Then
        ' en vodafo refacturado, el fichero será CSV...
        If TipoFicheroOperador = 3 Then 'vodafone
            'vSQL = DevuelveDesdeBD(conAri, "vodafoneRefacturado", "spara2", "1", "1")
            'If vSQL = "1" Then
                'vodafone  refacturado .Comprobaremos Lineas que NO vienen en el Fichero y estan activas
            vSQL = "operador=3 and inactivo=0 and idtelefono not in (select cuenta_no from vodafonetxt.fac_cab) AND 1"
            vSQL = DevuelveDesdeBD(conAri, "count(*)", "sclientfno ", vSQL, "1")
            If Val(vSQL) > 0 Then i = 1
            
            
        End If
        
        If TipoFicheroOperador = 1 Then 'vodafone
            'vSQL = DevuelveDesdeBD(conAri, "vodafoneRefacturado", "spara2", "1", "1")
            'If vSQL = "1" Then
                'vodafone  refacturado .Comprobaremos Lineas que NO vienen en el Fichero y estan activas
            vSQL = "( SELECT Numero_de_telefono from telefono.telefono WHERE fichero =" & DBSet(mfichero, "T") & ")"
            vSQL = "operador=1 and inactivo=0 and idtelefono not in " & vSQL & " AND 1"
            vSQL = DevuelveDesdeBD(conAri, "count(*)", "sclientfno ", vSQL, "1")
            If Val(vSQL) > 0 Then i = 1
            
            
        End If
        
        
    End If
    
    
    
    
    
    If vParamAplic.TelefoniaVtaPlazos Then
        vSQL = "operador = " & TipoFicheroOperador & " and ArtPlazos<>'' and PlazosMeses >0 AND 1"
        vSQL = DevuelveDesdeBD(conAri, "count(*)", "sclientfno", vSQL, "1")
        If Val(vSQL) > 0 Then i = i + 2
    End If
    
    
    
            
    
    If i > 0 Then
            'Para ver que mostramos en la siguiente pantalla
            vSQL = IIf(i = 1, 0, 1)    'el 1 significa que veremos el lw de vta plazos
            'CadenaDesdeOtroForm   1 Si muestra o no lw plazos      2 Operador   3 Nomfichero
            CadenaDesdeOtroForm = vSQL & "|" & TipoFicheroOperador & "|" & mfichero & "|"
                
            'Lanzamos el visor donde sal
            frmVarios3.Opcion = 3
            frmVarios3.Show vbModal
    
            
            If CadenaDesdeOtroForm = "" Then
                'MAL, ha cancelado
                L.Caption = ""
                L.Refresh

                Exit Function
            End If
            
            cad = "Select * from tmpnseries WHERE codusu = " & vUsu.Codigo
            Rc.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = ""
            While Not Rc.EOF
                
                'Febrero 2013
                'Borramos tb los datos que vienen del fichero
                L.Caption = "Quitar datos: " & Rc!codArtic
                L.Refresh
                
                cad = " WHERE Fichero =" & DBSet(mfichero, "T") & " AND Numero_de_telefono = " & DBSet(Rc!codArtic, "T")
                
                conn.Execute "DELETE FROM telefono.cuotas" & cad
                
                conn.Execute "DELETE FROM telefono.resumen_de_llamadas" & cad
                
                conn.Execute "DELETE from telefono.varios" & cad
                
                conn.Execute "DELETE FROM telefono.telefono" & cad
                Rc.MoveNext
            Wend
            Rc.Close
            
            DoEvents
                
    End If
    
    
    
    
    
    
    
    
    
    'Si quedan datos, los
    DoEvents
    L.Caption = "Ajuste cuotas propias"
    L.Refresh
    
    'Veremos is tien cuentas a insertar
    vSQL = DevuelveDesdeBD(conAri, "count(*)", "sclientfnoCuotas", "1", "1")
    If Val(vSQL) > 0 Then
    
        If TipoFicheroOperador = 3 Then
                'VODFONE
                
                cad = "max(dcul_cab_id)"
                Origen = DevuelveDesdeBD(conAri, cad, "vodafonetxt.dcul_cab", "1", "1", "N")
                If Origen = "" Then Origen = "0"
                idC = Val(Origen) + 1
        
                cad = "concat(min(fecha_inicio),'|',min(fecha_fin),'|',ciclo,'|')"
                Origen = DevuelveDesdeBD(conAri, cad, "vodafonetxt.dcul_cab", "fact_no", mfichero, "T")
                If Origen = "" Then Err.Raise 513, , "Imposible obtener fecha inicio/fin cuotas VOD"
                
                
                vSQL = "INSERT INTO vodafonetxt.dcul_cab (tipo_cuota,telef,importe,fecha_inicio,fecha_fin,fact_no,descuento,dcul_cab_id,ciclo)"
                vSQL = vSQL & " SELECT Descripcion,idtelefono,precio,'" & RecuperaValor(Origen, 1) & "','" & RecuperaValor(Origen, 2) & "',"
                 vSQL = vSQL & mdbFac2.texto(mfichero) & ",0.00,@rownum:=@rownum+1 AS rownum ,'" & RecuperaValor(Origen, 3) & "'"
                vSQL = vSQL & " FROM sclientfnoCuotas ,(SELECT @rownum:=" & idC & ") r WHERE "
            'Ene 2014  Fichero
            vSQL = vSQL & " idtelefono IN (select  distinct substring(telef,1,9) FROM vodafonetxt.dcul_cab where fact_no = " & mdbFac2.texto(mfichero) & " order by telef)"
        
        Else
            'lo que habia
            vSQL = "INSERT INTO telefono.cuotas (Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura,"
            vSQL = vSQL & "Numero_de_prefactura,Numero_de_telefono,Numero_de_extension , Fecha_inicio_cuotas, "
            vSQL = vSQL & " Fecha_final_cuotas, Codigo_de_cuota, Descripcion_de_cuota,"
            vSQL = vSQL & "Importe, Libre)  SELECT '" & mfichero & "',9999 + numlinea,1,'" & RecuperaValor(Origen, 1)
            vSQL = vSQL & "','',idtelefono,'','" & RecuperaValor(Origen, 2) & "','" & RecuperaValor(Origen, 3) & "',"
            vSQL = vSQL & "concat('D',numlinea),descripcion,precio,'ARIGES' FROM sclientfnoCuotas  WHERE "
            'Ene 2014  Fichero
            vSQL = vSQL & " IdTelefono IN (select distinct Numero_de_telefono FROM telefono.telefono where Fichero = " & mdbFac2.texto(mfichero) & " order by Numero_de_telefono)"
            
        End If
        mdbFac2.ejecutar vSQL

    End If





    AjusteCuotasNuevas2 = True


End Function

Public Sub ComprobarConceptosFacturacion(QueOperador As Byte)
Dim C As String

   

    If mfichero3 = "" Then mfichero3 = Fichero
 
    C = "Fichero =" & DBSet(mfichero3, "T") & " and not Codigo_de_tipo_de_trafico in (select Codigo_de_tipo_de_trafico from tel_conceptosllamadas where operadora=" & QueOperador & ") and 1"
    C = DevuelveDesdeBD(conAri, "count(*)", "telefono.resumen_de_llamadas", C, "1", "N")
    If Val(C) > 0 Then
        
    
    
        C = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2) "
        C = C & " select " & vUsu.Codigo & "," & QueOperador & " 1,0,Codigo_de_cuota,Descripcion_de_cuota"
        C = C & " where Fichero =" & DBSet(mfichero3, "T") & " group by Codigo_de_cuota"
        mdbFac2.ejecutar C
    
    
        C = "INSERT INTO tel_conceptosllamadas (operadora,Codigo_de_tipo_de_trafico,Tipo_de_trafico,refacturar)"
        C = C & " select " & QueOperador & " operadora,Codigo_de_tipo_de_trafico,Tipo_de_trafico,0 refacturar from telefono.resumen_de_llamadas"
        C = C & " WHERE Fichero = '" & mfichero3 & "'  group by Codigo_de_tipo_de_trafico order by Codigo_de_tipo_de_trafico"
        mdbFac2.ejecutar C
    End If

End Sub


Private Function DevuelveACcionesTipo(Accion As Byte, mes As Byte, ByRef RS As ADODB.Recordset, QueOperador As Byte) As String
Dim C As String
    C = "Select * from tel_descuentoscuotas WHERE accion=" & Accion
    If Accion > 1 Then C = C & " AND Mensual =" & mes  'si es cuota entera o prrorateable
    C = C & " AND operadora =" & QueOperador
'    If EsOrange Then
'        C = C & "2"
'    Else
'        C = C & "1"
'    End If
    RS.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = ""
    While Not RS.EOF
        C = C & ", " & DBSet(RS!codigo_de_cuota, "T")
        RS.MoveNext
    Wend
    RS.Close
  
    If C <> "" Then DevuelveACcionesTipo = Mid(C, 2) 'quito la primera coma
End Function

Public Function LeerParametrosFacturacionTelefonica(Companyia As Integer) As Boolean
Dim ReFacturaCutoas As Boolean

    On Error GoTo eLeerParametrosFacturacionTelefonica
    LeerParametrosFacturacionTelefonica = False
    mvarConceptosLlamadasRefacturables = ""
    mvarConceptosQueSonLlamadas = ""
    mvarDigitoCoarval = -1
    mvarCuotaMinimaSobreConsumos = False
    mvarDescuentosPorConsumo = False
    ReFacturaCutoas = False
    If vParamAplic.TieneTelefonia2 = 3 Then
        'BLOBAITE - anna. navarr
        ReFacturaCutoas = True
            
            
        mvarPrecioEstablecimiento = 12  'centimos de euro
        mvarPrecioMinuto = 9    'centimos de euro
        
        mvarDescuentosPorConsumo = False
        mvarCuotaMinimaSobreConsumos = False
        
        If miRsAux Is Nothing Then Set miRsAux = New ADODB.Recordset
        mvarFichero = "select Establecimiento ,PrecioUnMinuto,EstablecimientoORAN ,PrecioUnMinutoORAN , DigitoCoarval,CuotaMinimaSobreConsumos,DescuentosPorConsumo from spara2"
        miRsAux.Open mvarFichero, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
        
            If DBLet(miRsAux!DigitoCoarval, "N") > 0 Then mvarDigitoCoarval = miRsAux!DigitoCoarval
            
            
            mvarCuotaMinimaSobreConsumos = DBLet(miRsAux!CuotaMinimaSobreConsumos, "N") = 1
            mvarDescuentosPorConsumo = DBLet(miRsAux!DescuentosPorConsumo, "N") = 1
            
            If Companyia = 2 Then
                'EstablecimientoORAN PrecioUnMinutoORAN
                mvarPrecioEstablecimiento = miRsAux!EstablecimientoORAN  'centimos de euro
                mvarPrecioMinuto = miRsAux!PrecioUnMinutoORAN    'centimos de euro
            Else
                mvarPrecioEstablecimiento = DBLet(miRsAux!Establecimiento, "N") 'centimos de euro
                mvarPrecioMinuto = DBLet(miRsAux!PrecioUnMinuto, "N")   'centimos de euro
            End If
            
            
            
        End If
        miRsAux.Close
        
        mvarFichero = ""
    Else
        mvarFichero = DevuelveDesdeBD(conAri, "DigitoCoarval", "spara2", "1", "1")
        If Val(mvarFichero) > 0 Then mvarDigitoCoarval = Val(mvarFichero)
        mvarFichero = ""
        
    End If
    
    
    
    ''Conceptos refacturables
    If ReFacturaCutoas Then
        Set miRsAux = New ADODB.Recordset
        mvarFichero = "select Codigo_de_tipo_de_trafico,TraficoSegundos from tel_conceptosllamadas  where"
        mvarFichero = mvarFichero & " operadora = " & Companyia & " AND refacturar=1 order by Codigo_de_tipo_de_trafico"
        miRsAux.Open mvarFichero, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'MOVISTAR
            ''los mensajes son 20,21
            'mvarConceptosLlamadasRefacturables = "'17','18','W1','20','21'"
            mvarConceptosLlamadasRefacturables = mvarConceptosLlamadasRefacturables & ", " & DBSet(miRsAux!Codigo_de_tipo_de_trafico, "T")
            ' '17,18
            If DBLet(miRsAux!TraficoSegundos, "N") > 0 Then
                mvarConceptosQueSonLlamadas = mvarConceptosQueSonLlamadas & ", " & DBSet(miRsAux!Codigo_de_tipo_de_trafico, "T")
            Else
                mvarConceptosNOSonLlamadas = mvarConceptosNOSonLlamadas & ", " & DBSet(miRsAux!Codigo_de_tipo_de_trafico, "T")
            End If
            
            
            
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        'Quitamos la primera coma
        If mvarConceptosLlamadasRefacturables <> "" Then mvarConceptosLlamadasRefacturables = Mid(mvarConceptosLlamadasRefacturables, 2)
        If mvarConceptosQueSonLlamadas <> "" Then mvarConceptosQueSonLlamadas = Mid(mvarConceptosQueSonLlamadas, 2)
        If mvarConceptosNOSonLlamadas <> "" Then mvarConceptosNOSonLlamadas = Mid(mvarConceptosNOSonLlamadas, 2)
        mvarFichero = ""
    End If
    
    
    
    LeerParametrosFacturacionTelefonica = True
    Set miRsAux = Nothing
    Exit Function
eLeerParametrosFacturacionTelefonica:
    MuestraError Err.Number, Err.Description
    mvarFichero = ""
End Function








'---------------------------------------------------------------------------------
'ORANGE



Public Function DevuelveNombreFicheroOrange(NombreFicheroOrange As String) As String
Dim fin As Boolean

    On Error GoTo eDevuelveNombreFicheroOrange
    DevuelveNombreFicheroOrange = ""
    NF = FreeFile
    
    'El fichero viene sin vbcrlf
    Open NombreFicheroOrange For Input As #NF
    mvarFichero = ""
    While Not EOF(NF)
        Line Input #NF, mfichero3
        mvarFichero = mvarFichero & mfichero3
    Wend
    Close (NF)
    
    'Para que devuelve el nombre del fichero
    ComunNombreFicheroOrange
    
    
    DevuelveNombreFicheroOrange = mfichero3
    
    mvarFichero = ""
    mfichero3 = ""
    
    Exit Function
eDevuelveNombreFicheroOrange:
    MuestraError Err.Number
End Function



Private Sub ComunNombreFicheroOrange()
Dim i As Integer


    ''número de factura:;A10020017274-0813
    mfichero3 = ""
    
    NF = InStr(1, mvarFichero, "mero de factura:")
    If NF = 0 Then Exit Sub
    
    NF = NF + 16 'numero de facutra
    
    'buscamos el guion delimitador
    i = InStr(NF, mvarFichero, "-")
    If i = 0 Then Exit Sub
    
    If i - NF > 13 Then Exit Sub 'El telefono es varchar12
    
    mfichero3 = Mid(mvarFichero, NF + 1, i - NF - 1)

End Sub


Public Function cargarBaseDatosOrange(NombreFicheroOrange As String, ByRef LB As Label) As Boolean
Dim fin As Boolean
Dim i As Long
Dim J As Long
Dim K As Long
Dim Tramo As String
Dim Tipo As Byte
Dim IdLinDetalle As Long

    On Error GoTo eCargarBaseDatosOrange
    cargarBaseDatosOrange = False
    NF = FreeFile
    
    'El fichero viene sin vbcrlf
    Open NombreFicheroOrange For Input As #NF
    mvarFichero = ""
    While Not EOF(NF)
        Line Input #NF, mfichero3
        mvarFichero = mvarFichero & mfichero3
    Wend
    Close (NF)
    
    ComunNombreFicheroOrange  'devolvera el nombre del fichero
    mFicheroCorto = mfichero3
    mfichero3 = ""
    
    'Cargamos datos del
    If Not LB Is Nothing Then
        LB.Caption = "Eliminado datos anteriores"
        LB.Refresh
    End If
    borraDatos mFicheroCorto

    
    '-- El fichero por lo menos existe y grabamos sus datos
    'cargaDatosPrincipales mFicheroCorto   --> DEBERIAMOS CARGALO
    
    
    'A dia de hoy el fichero no viene con vbcrlf, si no con vblflf
    'Se trata de poder dividir las secciones del fichero y ver si las
    'estamos tratando todas.
    'Es dceir, puede que la seccion la reconozcamos, pero no hagamos nada con ella
    
    
    
        
    Tramo = DevuelveDesdeBD(conAri, "max(id)", "telefono.detalle_de_llamadas", "1", "1")
    If Tramo = "" Then Tramo = "0"
    IdLinDetalle = Val(Tramo) + 1
    
    
    i = 1
    Do
        K = InStr(i, mvarFichero, Chr(10) & Chr(10))
        If K > 0 Then
        
            If K > 40 Then
                Tramo = Mid(mvarFichero, i, K - i - 1)
                J = 0
                Do
                    If Mid(Tramo, 1, 1) = vbLf Then
                        Tramo = Mid(Tramo, 2)
                    Else
                        J = 1
                    End If
                Loop Until J = 1
                
                Tipo = ProcesamosTramoOrange(Tramo)
                
                LB.Caption = "Fase : " & Tipo
                LB.Refresh
                
                
                If Tipo = 0 Then
                    'MsgBox "TIPO=0", vbExclamation
                                 
                    LB.Caption = "Fase :0 ---> Proceso en marcha"
                    LB.Refresh
                    Espera 1
                
                Else
'

                    
                    If Tipo = 1 Then
                        'NADA
                    Else
                        If Tipo > 1 And Tipo < 10 Then  '2 al 9 detalles de llamada
                            InsertarTramoDetalleLlamdas IdLinDetalle, Tipo, Tramo, LB
                        Else
                            OrangeInsertBdTelefono Tipo, Tramo
                        End If
                    End If
                End If
            Else
                'Stop
            End If
            
            i = K + 2
        End If
    Loop Until K = 0
            
            
            
            
    'Lo recalculamos pq viene a 2 decimales y en el detalle a 4
    'Hay error
    LB.Caption = "Obtencion resumen detalle"
    LB.Refresh
            
    Tramo = "insert into telefono.resumen_de_llamadas(Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura,"
    Tramo = Tramo & "Numero_de_telefono,Codigo_de_tipo_de_trafico,Tipo_de_trafico,Numero_de_llamadas,"
    Tramo = Tramo & "Unidad_de_medida,Cantidad_medida,Importes)"

    Tramo = Tramo & " select fichero,'320','1',fichero"
    Tramo = Tramo & ",Numero_de_telefono,Codigo_de_trafico,Tipo_de_trafico,count(*),Unidad_de_medida , "
    Tramo = Tramo & "Sum(Cantidad_medida_originada), Sum(Importe) FROM telefono.detalle_de_llamadas "
    Tramo = Tramo & " WHERE fichero = " & DBSet(mFicheroCorto, "T")
    Tramo = Tramo & " group by Numero_de_telefono,Codigo_de_trafico"
    Tramo = Tramo & " order by Numero_de_telefono"
            
    conn.Execute Tramo
            
            
    'AJUSTE en la tabla
    'del periodo de facturacion
    'Obtenemos de la tabla el periodo
    LB.Caption = "Obtencion periodo facturacion"
    LB.Refresh
    Espera 0.5
    
    'resumen_de_llamadas Fecha_inicio_periodo Fecha_final_periodo
    Tramo = DevuelveDesdeBD(conAri, "Nombre_1", "telefono.cabecera_de_soporte", "fichero", mFicheroCorto, "T")
    If Tramo = "" Then
        MsgBox "No es posible localizar peridodo de facturacion", vbExclamation
        Exit Function
    End If
    
    
    Tramo = Replace(Tramo, "-", "")
    J = InStr(1, Tramo, " ")
    
    Tramo = "'" & Format(Mid(Tramo, 1, J - 1), FormatoFecha) & "' , Fecha_final_periodo = '" & Format(Trim(Mid(Tramo, J + 1)), FormatoFecha) & "'"
    Tramo = "UPDATE telefono.resumen_de_llamadas SET Fecha_inicio_periodo = " & Tramo
    Tramo = Tramo & " WHERE fichero = " & DBSet(mFicheroCorto, "T")
    
    conn.Execute Tramo
    
    Tramo = Replace(Tramo, "telefono.resumen_de_llamadas", "telefono.cuotas")
    '
    Tramo = Replace(Tramo, "Fecha_inicio_periodo", "Fecha_inicio_cuotas")
    Tramo = Replace(Tramo, "Fecha_final_periodo", "Fecha_final_cuotas")
    conn.Execute Tramo

    
        
        
    
        
    
    
   cargarBaseDatosOrange = True
   Exit Function
eCargarBaseDatosOrange:
    MuestraError Err.Number
End Function


Private Function ProcesamosTramoOrange(ByRef Tramo As String) As Byte
Dim C As String
Dim J As Integer

    ProcesamosTramoOrange = 0

    'Segun como empieza el tramo.
    J = InStr(1, Tramo, ":")
    If J = 0 Then Exit Function
        
    C = Mid(Tramo, 1, J - 1)
    
    
    If C = "datos del cliente" Then
        ProcesamosTramoOrange = 20   'Estan ok, sabemos que estan, pero no los procesamos
        Exit Function
    End If
    If C = "número de factura" Then
        ProcesamosTramoOrange = 21   'Estan ok, sabemos que estan, pero no los procesamos
        Exit Function
    End If
    
    If Mid(C, 1, 17) = "detalle de las ll" Then
        'Veremos si procesamos las lineas de detalle
        C = Mid(Tramo, J + 1)
        J = InStr(1, C, vbLf)
        If J = 0 Then
            'NO puedo saber de que es el detalle
            ProcesamosTramoOrange = 0
        
        Else
            C = LCase(Mid(C, 1, J - 1))
            
            If C = "nacional" Then
                ProcesamosTramoOrange = 2 'Detalle llamadas nacional
            
            ElseIf C = "internacional" Then
                ProcesamosTramoOrange = 3 'Detalle llamadas internacional
                
            ElseIf C = "datos" Then
                ProcesamosTramoOrange = 4 'Datos
            
            ElseIf C = "roaming" Then
                ProcesamosTramoOrange = 5 'Roamning
            
            ElseIf C = "rpv" Then
                ProcesamosTramoOrange = 6 'rpv
                
            ElseIf C = "servicios especiales" Then
                ProcesamosTramoOrange = 7 'servicios especiales
                
            ElseIf C = "mensajes" Then
                ProcesamosTramoOrange = 8 'mensajes
            
            ElseIf C = "datos roaming" Then
                ProcesamosTramoOrange = 9 'datos roaming
                
                
            End If
        
        
        
        
        End If
    Else
        If Mid(C, 1, 17) = "resumen servicios" Then
            ProcesamosTramoOrange = 1   'Estan ok, sabemos que estan, pero no los procesamos
        ElseIf Mid(C, 1, 17) = "resumen de cargos" Then
            ProcesamosTramoOrange = 22   'Metera cuantos telefonos tratamos
        Else
            ProcesamosTramoOrange = 0 'NO LO tratamos
        End If
    
    End If
    
        
        
    

End Function





Private Sub OrangeInsertBdTelefono(Tipo As Byte, Tramo As String)
Dim Aux As String
Dim i As Integer
Dim F As Date
Dim H As Integer
Dim CadenaInsert As String

Dim Minima As String
Dim Importe As Currency
Dim linea As String
    
    Select Case Tipo
    Case 21
        'Numero de facutura
        'cabecera_de_soporte
        F = Now
        i = InStr(1, Tramo, "fecha de factura:") '18lenth
        If i > 0 Then
            Aux = Mid(Tramo, i + 18, 10)
            If IsDate(Aux) Then F = CDate(Aux)
        End If
        If F = Now Then MsgBox "No se encuetra fecha factura ORANGE"
        
        
        'Insertaremos en telefono.factura telefono.cabecera_de_soporte  telefono.ficheros
        If Month(F) > 9 Then
            i = Month(F) - 9
            Aux = Mid("ABC", i, 1)
        Else
            Aux = Month(F)
        End If
            
        Aux = DBSet(mFicheroCorto, "T") & ",'O','" & Aux & "','" & Year(F) & "')"
        Aux = "INSERT INTO telefono.Ficheros(Fichero,Servicio,Mes,Ano) VALUES (" & Aux
        conn.Execute Aux
    

        i = InStr(1, Tramo, "F/CIF:")
        If i > 0 Then
            Aux = Mid(Tramo, i + 7, 10)
            Aux = Mid(Aux, 1, Len(Aux) - 1)
        Else
            MsgBox "No se encuetra el CIF"
        End If
        Aux = DBSet(mFicheroCorto, "T") & ",'1O','1','" & Aux & "','"
        'Periodo de facturacion
        i = InStr(1, Tramo, "odo facturado:")
        If i > 0 Then
            H = InStr(i + 1, Tramo, vbLf)
            If H = 0 Then i = 0
        End If
            
        If i > 0 Then Aux = Aux & Mid(Tramo, i + 15, 23)
        
        Aux = Aux & "')"
        Aux = "INSERT INTO telefono.cabecera_de_soporte(Fichero,Tipo_de_Registro,Secuencia_de_Registro,CIF,Nombre_1) VALUES (" & Aux
        conn.Execute Aux
        
        Aux = DBSet(mFicheroCorto, "T") & ",'100','1'," & DBSet(mFicheroCorto, "T") & "," & DBSet(F, "F") & ")"
        Aux = "INSERT INTO telefono.factura(Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_factura,Fecha_de_factura) VALUES (" & Aux
        conn.Execute Aux
    
    
    Case 20
    
            
            
    
    Case 22
    
        'Minima = DevuelveDesdeBD(conAri, "DescCuota", "tel_desc_cuotas", "CodCuota", "ARI00000", "T")
        'If Minima = "" Then Err.Raise 513, , "Error obteniendo cuota minima"
        
        'telefono(Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura,Numero_de_telefono,Suma_total)
        H = InStr(1, Tramo, vbLf)
        H = InStr(H + 1, Tramo, vbLf)
        If H = 0 Then
            Tramo = ""
        Else
            Tramo = Mid(Tramo, H + 1)
        End If
    
    
        'resumen de cargos por cada número:
        'número de teléfono;tarifa;cuota;nacionales;internacionales;rpv;tempo grupo;roaming;m. texto;especiales;GPRS;GPRS Roaming;importe (euros);
        '609491517;Sol Personaliz.  RPV;0,00;0,00;0,00;0,00;0,00;0,00;1,95;0,00;0,00;0,00;1,95;
    
        'cuota;nacionales;internacionales;rpv;tempo grupo;roaming;m. texto;especiales;GRPS;GPRS Roaming
    
        CadenaInsert = ""
        i = 0
        While Tramo <> ""
            H = InStr(1, Tramo, vbLf)
            If H = 0 Then
                Tramo = ""
                Aux = ""
            Else
                Aux = Mid(Tramo, 1, H - 1)
                Tramo = Mid(Tramo, H + 1)
            End If
            If Aux <> "" Then
                Aux = Replace(Aux, ";", "|")
                linea = Aux
                
                i = i + 1                                                '300 pq el de movistar viene un 300
                CadenaInsert = CadenaInsert & ", ('" & mFicheroCorto & "','300'," & i & ",'" & mFicheroCorto & "','"
                'Numero de telefono
                CadenaInsert = CadenaInsert & RecuperaValor(Aux, 1) & "'," & TransformaComasPuntos(RecuperaValor(Aux, 13)) & ")"
            
            
                'INSERT COMUN
                If Len(CadenaInsert) > 30000 Then
                        CadenaInsert = Mid(CadenaInsert, 2)
                        Aux = "insert into telefono.telefono(Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura,Numero_de_telefono,Suma_total) VALUES "
                        Aux = Aux & CadenaInsert
                        conn.Execute Aux
                        CadenaInsert = ""
                End If
            
                'CUOTA MINIMA
                'Desde la poscion 3 a la 10 van los distintos cargos.
                'Si la suma no coincide es que se la ha sumado un cargo de consumo minimo
                
                
                Importe = 0
                For H = 3 To 12
                    Aux = RecuperaValor(linea, H)
                    If Aux <> "0,00" Then Importe = Importe + CCur(Aux)
                Next
                Aux = RecuperaValor(linea, 13)
                If Aux <> "0,00" Then
                    Importe = CCur(Aux) - Importe
                    If Importe <> 0 Then
                        Aux = "INSERT INTO telefono.cuotas(Fichero,Tipo_de_registro,Secuencia_de_registro,"
                        Aux = Aux & "Numero_de_factura,Numero_de_telefono,Codigo_de_cuota,"
                        Aux = Aux & "Descripcion_de_cuota,Importe) VALUES (" & DBSet(mFicheroCorto, "T") & ",'310',1,"
                        Aux = Aux & DBSet(mFicheroCorto, "T") & ",'" & RecuperaValor(linea, 1) & "','AR0','Consumo mínim',"
                        Aux = Aux & TransformaComasPuntos(CStr(Importe)) & ")"
                        
                        'NO INSERTAMOS LA CUOTA MINIMA
                        'conn.Execute Aux
                    End If
                    End If
            End If
        Wend
        'INSERT COMUN
        If Len(CadenaInsert) > 0 Then
                CadenaInsert = Mid(CadenaInsert, 2)
                Aux = "insert into telefono.telefono(Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura,Numero_de_telefono,Suma_total) VALUES "
                Aux = Aux & CadenaInsert
                conn.Execute Aux
        End If
        
    End Select

End Sub




Private Sub InsertarTramoDetalleLlamdas(ByRef IDLinea As Long, Tipo As Byte, Tramo As String, ByRef LB2 As Label)
Dim H As Long
Dim linea As String
Dim Aux As String
Dim TotalLlamada As Currency

Dim Ud As String
Dim Libre As String
Dim Valor As Currency
Dim Valor2 As Currency
    
Dim LenFich As Long
Dim LlevoF As Long

Dim Descripcion_tipo_de_llamada_y_tipo_destino As String
Dim DetaL

    On Error GoTo eInsertarTramoDetalleLlamdas
    
    '-->Fecha mmdd   hora hhmm
    'Fichero,id,Numero_de_telefono,Codigo_de_trafico,Tipo_de_trafico,Numero_llamado,Fecha,Hora_inicio,Unidad_de_medida,  Cantidad_medida_originada Importe
    'quito las dos primeras lineas que son expliativas
    
    LenFich = Len(Tramo)
    LB2.Tag = LB2.Caption
    
    H = InStr(1, Tramo, vbLf)
    H = InStr(H + 1, Tramo, vbLf)
    If H = 0 Then
        Tramo = ""
    Else
        Tramo = Mid(Tramo, H + 1)
    End If
    
        
    CadenaInsert = ""
    While Tramo <> ""
        H = InStr(1, Tramo, vbLf)
        If H = 0 Then
            'JUnio 18
            'Antes no iba fino si solo habia una linea
            Tramo = Tramo & vbLf
            H = InStr(1, Tramo, vbLf)
        End If
        If H = 0 Then
            Tramo = ""
            linea = ""
        Else
            linea = Mid(Tramo, 1, H - 1)
            Tramo = Mid(Tramo, H + 1)
            
            LlevoF = LlevoF + Len(linea)
            LB2.Caption = LB2.Tag & "     " & LlevoF & " / " & LenFich
            LB2.Refresh
        End If
        If linea <> "" Then
            If InStr(1, linea, "|") > 0 Then
                MsgBox "Separador incorrecto", vbExclamation
                MsgBox "Error, LINEA: " & linea
            End If
            linea = Replace(linea, ";", "|") & "|"  'pong el ultimo pipe
            
            CadenaInsert = CadenaInsert & ", ('" & mFicheroCorto & "'," & IDLinea
            IDLinea = IDLinea + 1
            
            'INSERT en BD
            'Fichero,id,Numero_de_telefono,Codigo_de_trafico,Tipo_de_trafico,Unidad_de_medida,Numero_llamado,Fecha,Hora_inicio,  Cantidad_medida_originada Importe
            
            
            'Todos los tipos tran los mismos datos
            'Número de Telefóno;Fecha;Número destino;Tipo llamada;Destino/Operador;Especiales;Inicio;Duración/Vol;Tarifa;Import
            'Número de Telefóno;Fecha;Número destino;Tipo llamada;Destino/Operador;Especiales;Inicio;Duración/Vol;Tarifa;Import
            'Número de Telefóno;Fecha;Número destino;Tipo llamada;Destino/Operador;Especiales;Inicio;Duración/Vol;Tarifa;Importe
            'Número de Telefóno;Fecha;Número destino;Tipo llamada;Destino/Operador;Especiales;Inicio;Duración/Vol;Tarifa;Importe
            'Número de Telefóno;Fecha;Número destino;Tipo llamada;Destino/Operador;Especiales;Inicio;Duración/Vol;Tarifa;Importe
            CadenaInsert = CadenaInsert & ",'" & RecuperaValor(linea, 1) & "',"
            
            If Tipo = 2 Then
                CadenaInsert = CadenaInsert & "'1','NACIONAL',"
            ElseIf Tipo = 3 Then
                CadenaInsert = CadenaInsert & "'2','INTERNACIONAL',"
            ElseIf Tipo = 4 Then
                'DATOS
                CadenaInsert = CadenaInsert & "'3','DATOS',"
            ElseIf Tipo = 5 Then
                'DATOS
                
                CadenaInsert = CadenaInsert & "'3','DATOS',"
            ElseIf Tipo = 6 Then
                'DATOS
                CadenaInsert = CadenaInsert & "'5','RPV',"
            ElseIf Tipo = 7 Then
                'Llamadas especiales
                
                Aux = RecuperaValor(linea, 5)
                If Aux = "SMS/MMS Premium" Then
                    
                    
                    CadenaInsert = CadenaInsert & "'11','SMS/MMS Premium',"
                Else
                    
                    
                    CadenaInsert = CadenaInsert & "'6','ESPECIALES',"
                End If
            ElseIf Tipo = 8 Then
                CadenaInsert = CadenaInsert & "'7','Mensajes',"
            ElseIf Tipo = 9 Then
                CadenaInsert = CadenaInsert & "'8','Datos roaming',"
                
                
                
                
            End If
            CadenaInsert = CadenaInsert & "'" & RecuperaValor(linea, 3)
            Aux = RecuperaValor(linea, 2) 'fecha
            Aux = Format(Aux, "mmdd")
            CadenaInsert = CadenaInsert & "','" & Aux & "','"
            Aux = RecuperaValor(linea, 7)
            Aux = Mid(Aux, 1, 2) & Mid(Aux, 4, 2)
            CadenaInsert = CadenaInsert & Aux & "',"
            'Cantidad
            
            
            
            
            Aux = RecuperaValor(linea, 6) 'unidada
            Libre = ""
            TotalLlamada = DevuelveUnidad(Tipo, RecuperaValor(linea, 8), Aux)
            If Tipo <> 4 Then
                If Tipo = 8 Then
                    Aux = "SMS"
                ElseIf Tipo = 7 Then
                    If Aux = "" Then Aux = "Segundos"
                ElseIf Tipo = 9 Then
                    If Aux = "" Then Aux = "KB"
                Else
                    Aux = "Segundos"
                End If
            End If
            
            CadenaInsert = CadenaInsert & "'" & Aux & "'," & TransformaComasPuntos(CStr(TotalLlamada)) & ","
            
            'Importe
            Aux = RecuperaValor(linea, 10)
           
            
            
            
            
            
            'Se refacturan
            If Tipo = 2 Then
                'Se refactura
                
                If CCur(Aux) <> 0 Then
                    'Refacturamos
                    Libre = TransformaComasPuntos(Aux)
                    
                    Valor = Int(CCur(TotalLlamada))
                    Valor2 = (Valor * mvarPrecioMinuto) + mvarPrecioEstablecimiento
                
                    Valor = CCur(TotalLlamada) - Valor
                    Valor = Valor * 100
                    Valor = (Valor * mvarPrecioMinuto) / 60
                    
                    
                    Valor2 = Valor2 + Valor
                    Aux = CStr(Round((Valor2 / 100), 4))
                    
                End If
                
                                
                                
            ElseIf Tipo = 6 Then
                 Libre = TransformaComasPuntos(Aux)
                 Aux = "0,0"
            End If
            
            
            
            CadenaInsert = CadenaInsert & TransformaComasPuntos(Aux) & ",'" & Libre & "',"
            
            
            Descripcion_tipo_de_llamada_y_tipo_destino = "null,null)"
            Aux = Trim(RecuperaValor(linea, 6) & " " & RecuperaValor(linea, 4))
            If Aux <> "" Then
                Descripcion_tipo_de_llamada_y_tipo_destino = DBSet(Aux, "T") & ","
                Aux = Trim(RecuperaValor(linea, 5))
                Descripcion_tipo_de_llamada_y_tipo_destino = Descripcion_tipo_de_llamada_y_tipo_destino & DBSet(Aux, "T") & ")"
            End If
            CadenaInsert = CadenaInsert & Descripcion_tipo_de_llamada_y_tipo_destino
            
            
            'INSERT COMUN
            If Len(CadenaInsert) > 30000 Then
            
            
                    ' aqui FALTA###  Descripcion_tipo_de_llamada y tipo_destino en BD.detalle_de_llamadas
            
            
            
                    CadenaInsert = Mid(CadenaInsert, 2)
                    Aux = "insert into telefono.detalle_de_llamadas(Fichero,id,Numero_de_telefono,Codigo_de_trafico,Tipo_de_trafico,Numero_llamado,fecha,Hora_inicio, Unidad_de_medida, Cantidad_medida_originada,Importe,libre,Descripcion_tipo_de_llamada , tipo_destino) VALUES "
                    Aux = Aux & CadenaInsert
                    conn.Execute Aux
                    CadenaInsert = ""
            End If
        End If
    Wend
    LB2.Caption = LB2.Tag & " FIN"
    LB2.Refresh
    
    If Len(CadenaInsert) > 0 Then
        CadenaInsert = Mid(CadenaInsert, 2)
        Aux = "insert into telefono.detalle_de_llamadas(Fichero,id,Numero_de_telefono,Codigo_de_trafico,Tipo_de_trafico,Numero_llamado,fecha,Hora_inicio,Unidad_de_medida,  Cantidad_medida_originada,Importe,libre,Descripcion_tipo_de_llamada , tipo_destino) VALUES "
        Aux = Aux & CadenaInsert
        conn.Execute Aux
        CadenaInsert = ""
    End If
    
    
    
eInsertarTramoDetalleLlamdas:
    If Err.Number <> 0 Then
        Aux = "Error procesando tramo. " & vbCrLf & linea
        MsgBox Aux, vbCritical
        Err.Clear
        Err.Raise 513, , "Revise fichero a procesar"
    End If
End Sub



Private Function DevuelveUnidad(ByVal Tipo As Byte, ByVal Aux As String, ByRef Ud As String) As Currency
Dim H As Integer
Dim K As Integer
Dim Cambialo As Byte

    If Tipo = 7 Then
        
        Cambialo = 0
        If LCase(Right(Aux, 2)) = "gb" Then
            Cambialo = 1
        ElseIf LCase(Right(Aux, 2)) = "mb" Then
            Cambialo = 1
        ElseIf LCase(Right(Aux, 2)) = "kb" Then
            Cambialo = 1
        ElseIf LCase(Right(Aux, 5)) = "bytes" Then
            Cambialo = 1
        End If
        If Cambialo Then Tipo = 4
        'If Cambialo = 0 Then St op
    End If
    If Tipo = 4 Or Tipo = 9 Then
        DevuelveUnidad = 0
        Aux = Trim(Aux)
        If LCase(Right(Aux, 2)) = "gb" Then
            'MsgBox "GIGAS"
            Ud = "Gigabytes"
            Aux = Trim(Mid(Aux, 1, Len(Aux) - 2))
            If Aux = "" Then Aux = "0"
            DevuelveUnidad = CCur(Aux)
        End If
        If LCase(Right(Aux, 2)) = "mb" Then
            
            Ud = "Megabytes"
            Aux = Trim(Mid(Aux, 1, Len(Aux) - 2))
            If Aux = "" Then Aux = "0"
            DevuelveUnidad = CCur(Aux)
        End If
        If LCase(Right(Aux, 2)) = "kb" Then
            
            Ud = "Kilobytes"
            Aux = Trim(Mid(Aux, 1, Len(Aux) - 2))
            If Aux = "" Then Aux = "0"
            DevuelveUnidad = CCur(Aux)
        End If
        If LCase(Right(Aux, 5)) = "bytes" Then
            Ud = "bytes"
            Aux = Trim(Mid(Aux, 1, Len(Aux) - 5))
            If Aux = "" Then Aux = "0"
            DevuelveUnidad = CCur(Aux)
        End If
        
        
    
    Else
            DevuelveUnidad = 0
            H = InStr(1, Aux, "h")
            If H > 0 Then
                
                K = Val(Mid(Aux, 1, H - 1))
                Aux = Trim(Mid(Aux, H + 1))
                DevuelveUnidad = DevuelveUnidad + (60 * K)
            End If
            H = InStr(1, Aux, "m")
            If H > 0 Then
                K = Val(Mid(Aux, 1, H - 1))
                Aux = Trim(Mid(Aux, H + 1))
                DevuelveUnidad = DevuelveUnidad + K
            End If
            
            H = InStr(1, Aux, "s")
            If H > 0 Then
                 K = Val(Mid(Aux, 1, H - 1))
                 Aux = Trim(Mid(Aux, H + 1))
                 DevuelveUnidad = DevuelveUnidad + (K / 100)
            End If
            
            'Tipo=7
            If Tipo = 7 Then Ud = "minutos"
            
            If Aux <> "" Then
                MsgBox "error separando horas"
                Stop
            End If
    
      
    End If
End Function










'***********************************************************************************
'***********************************************************************************
'
'
'   VODAFONE    VODAFONE     VODAFONE     VODAFONE    VODAFONE     VODAFONE
'
'
'***********************************************************************************
'***********************************************************************************
Public Function cargarBaseDatosVODAFONE(mfichero As String, ByRef LB As Label) As Boolean
Dim AvanCoop As Boolean
    
    'Mayo 2020
    ' Parametro para leer o bien los ficheros tal y como vienen de la web de vodafone,
    ' o bien los que procesa una empresa AVANCOOP
    
    CadenaInsert = DevuelveDesdeBD(conAri, "vodafoneRefacturado", "spara2", "1", "1")
    AvanCoop = CadenaInsert = "1"
    CadenaInsert = ""
    If AvanCoop Then
        cargarBaseDatosVODAFONE = cargarBaseDatosVODAFONE_Refacturado(mfichero, LB)
    Else
        cargarBaseDatosVODAFONE = cargarBaseDatosVODAFONENormal(mfichero, LB)
        
    End If





    'Octubre 2020
    'Falta insertar en telefono.telefono todos los que vienen en el traspaso
    CadenaInsert = ""
    whereTelefono = ""
    
        'mvarFichero mfichero2
        CadenaInsert = "SELECT " & DBSet(mFicheroCorto, "T") & ",320,fac_cab_id ," & DBSet(mFicheroCorto, "T") & ",cuenta_no,0"
        CadenaInsert = CadenaInsert & " FROM vodafonetxt.fac_cab WHERE  fac_no =" & DBSet(mFicheroCorto, "T")
        
        
        whereTelefono = "INSERT IGNORE INTO telefono.telefono(Fichero,Tipo_de_registro,Secuencia_de_registro,"
        whereTelefono = whereTelefono & " Numero_de_factura,Numero_de_telefono,Suma_total) "
        CadenaInsert = whereTelefono & CadenaInsert
        conn.Execute CadenaInsert
    

    CadenaInsert = ""
    whereTelefono = ""
    



End Function




Private Function cargarBaseDatosVODAFONENormal(mfichero As String, ByRef LB As Label) As Boolean

    Dim leido As String ' registro leido del fichero
    Dim i As Long ' un contadoir de registros
    Dim GuardarNombreFichero As String
    Dim tabla As String

    
    
    On Error GoTo ecargarBaseDatosVODAFONE
    
    
    cargarBaseDatosVODAFONENormal = False
    '-- el fichero existe y vamos a procesar sus datos, lo abrimos

    mFicheroCorto = Right(mfichero, 12)
  
    EliminarTodoElFichero mFicheroCorto, LB

    
    NF = FreeFile()
    Open mfichero For Input As #NF
    
    '- Borramos registros del mismo fichero si los hubiese
    borraDatos mFicheroCorto
    
    
    
    '-- El fichero por lo menos existe y grabamos sus datos
    cargaDatosPrincipales mFicheroCorto
    
    
    
    '--- Para no tener que pasar variables a todas y cada un
    'ya que las funciones las he hecho copiando pegando
    'del .NET
    
    
    GuardarNombreFichero = mFicheroCorto
    
    CadenaInsert = ""
    
    Do While Not EOF(NF)
        Line Input #NF, leido
        i = i + 1
        
        
        
        LB.Caption = "Registro " & CStr(i)
        LB.Refresh
        If (i Mod 40) = 0 Then DoEvents
        
        '-- hacemos lo que tengamos que hacer con el registro
        If leido <> "" Then
            
            Valores = Split(leido, Chr(9))
                       
                       
                       
                
          
        
        
                       
                       
                       
            Select Case Left(leido, 8)
                Case "FACT_DAT":
                        ProcesaFacCab
                        tabla = "Fac_Cab(ciclo,Fac_No,Cuenta_No,Fecha_Emision,Titular,Nif_Cif,Lugar_Emision,Fecha_Inicio ,Fecha_Fin ,No_Lineas )"
      
                Case "IDCL_DAT":
                        ProcesaIdclCab
                        tabla = "Idcl_Cab(ciclo,Fac_No,Nombre_Emp ,Atencion_De ,Dir1 ,Dir2 ,Dir3 ,Dir4 ,Dir5 )"
                Case "AVPG_DAT":
                        ProcesaAvpgCab
                        tabla = "Avpg_Cab(ciclo,Fac_No,Forma_Pago ,Ent_Bancaria ,No_Cuenta ,Cta_Ingreso ,Refer ,No_Tarjeta ,Fecha_Vto ) "
                Case "RCFE_DAT":
                        ProcesaRcfeCab
                        tabla = "Rcfe_Cab(ciclo,Fac_No ,Cat_Concep ,Tipo_Concep ,Importe )"
                Case "RIMP_DAT":
                        ProcesaRimpCab
                        tabla = "Rimp_Cab(ciclo,Fac_No ,Base_Imp ,Tipo_Imp ,Impuesto ,Ttal  )"
                Case "ECTA_DAT":
                        ProcesaEctaCab
                        tabla = "Ecta_Cab(ciclo,Fac_No ,Saldo_Pend ,Total_Factura ,Abono_Pend ,Total_Pagar    )        "
               
                Case "RCLL_DAT":
                        ProcesaRcllCab
                        tabla = "Rcll_Cab(ciclo,Fac_No ,Cat_Cons ,Tipo_Cons,Numero ,Minutos_Kb ,Importe ,Descuento,Importe_Neto )"
                        '            Ciclo,FacNo,CatCons,TipoCons,Numero,MinutosKb,Importe ,Descuento,ImporteNeto
                Case "RCME_DAT":
                        ProcesarRcmeCab
                        tabla = "Rcme_Cab(ciclo,Fac_No,Cat_Cons ,Tipo_Cons ,Numero ,Minutos_Kb,Importe ,Descuento ,Importe_Neto)"
        
                Case "RCDA_DAT":
                        ProcesarRcdaCab
                        tabla = "Rcda_Cab(ciclo,Fact_No ,Cat_Cons,Tipo_Cons ,Numero ,Minutos_Kb ,Importe,Descuento,Importe_Neto )"
        
                Case "RCTL_DAT":
                        ProcesarRctlCab
                        tabla = "Rctl_Cab(ciclo,Fact_No ,Cat_Cuota ,Tipo_Cuota ,Num_Cuota,Importe ,Descuento,importe_neto )"
                Case "RDES_DAT":
                        ProcesarRdesCab
                        tabla = "Rdes_Cab(ciclo,Fact_No ,Nivel_Linea ,Nivel_Cuenta ,Importe)"

                Case "DCFC_DAT":
                        ProcesarDcfcCab
                        tabla = "Dcfc_Cab(ciclo,Fact_No ,Cuenta ,Cuotas ,Descuentos ,Importe   )      "
        
                Case "DCFV_DAT":
                        ProcesarDcfvCab
                        tabla = "Dcfv_Cab(ciclo,Fact_No ,Telef ,P_Precios_Voz ,P_Precios_Datos ,Tipo ,Importe ,Descuento,Importe_Neto )"

                Case "DCFM_DAT":
                        ProcesarDcfmCab
                        tabla = "Dcfm_Cab(ciclo,Fact_No  ,Telef  ,P_Precios_Voz  ,P_Precios_Datos  ,Tipo ,Importe  ,Descuento  ,Importe_Neto  )"
        
                Case "DCFD_DAT":
                        ProcesarDcfdCab
                        tabla = "Dcfd_Cab(ciclo,Fact_No  ,Telef ,P_Precios_Voz ,P_Precios_Datos ,Tipo ,Importe  ,Descuento  ,Importe_Neto  )"
                                        
                Case "DCFT_DAT":
                        ProcesarDcftCab
                        tabla = "Dcft_Cab(ciclo,Fact_No  ,Telef  ,P_Precios_Voz  ,P_Precios_Datos ,Tipo  ,Importe  ,Descuento  ,Importe_Neto    )      "

                Case "DCUC_DAT":
                        ProcesarDcucCab
                        tabla = "Dcuc_Cab(ciclo,Fact_No ,Cuenta  ,Tipo_Cuota  ,Importe  ,Descuento  ,Fecha_Inicio ,Fecha_Fin) "
                   
                Case "DCUL_DAT":
                        ProcesarDculCab
                        tabla = "Dcul_Cab(ciclo,Fact_No ,Telef ,Tipo_Cuota ,Importe ,Descuento ,Fecha_Inicio ,Fecha_Fin )"

                Case "DDNC_DAT":
                        ProcesarDdncCab
                        tabla = "Ddnc_Cab(ciclo,Fact_No ,Cuenta ,Tipo_Desc,Cantidad ,Porcentaje,Importe ,Fecha_Inicio,Fecha_Fin )"

                Case "DDNL_DAT":
                        ProcesarDdnlCab
                        tabla = "Ddnl_Cab(ciclo,Fact_No ,Telef ,Tipo_Desc ,Cantidad ,Porcentaje ,Importe,Fecha_Inicio ,Fecha_Fin )     "
                    
        
        
                Case "DNTV_DAT":
                     ProcesarDntvCab
                     tabla = "Dntv_Cab(ciclo,Fact_No,Telef_Ext ,Fecha_Hora ,No_Recep ,Tipo ,Destino ,Duracion ,Durac_Ses "
                    tabla = tabla & ",Volumen,Pasos ,Tipo_Fact ,Tarifa,Descripcion ,Canal ,Ext_Origen ,Importe ,Importe_Neto) "
        
                Case "DNTD_DAT":
                     ProcesarDntdCab
                     tabla = "Dntd_Cab(ciclo,Fact_no ,Telef_ext ,Fecha_hora ,No_recep ,Tipo ,Destino ,Duracion ,Duracion_ses"
                     tabla = tabla & ",Volumen ,Pasos ,Tipo_fact ,Tarifa ,Descripcion ,Canal ,Ext_origen ,Importe ,Importe_neto )"
                     
                Case "DSVA_DAT":
                     ProcesarDsvaCab
                     tabla = "Dsva_Cab(ciclo,Fact_No ,Telef_Ext ,Fecha_Hora ,No_Recep ,Tipo ,Destino ,Duracion ,Durac_Ses "
                     tabla = tabla & ",Volumen ,Pasos ,Tipo_Fact ,Tarifa ,Descripcion ,Canal ,Ext_Origen ,Importe ,Importe_Neto )"
                
                Case "DOIM_DAT":
                     ProcesarDoimCab
                     tabla = "Doim_Cab(ciclo,Fact_No  ,Tipo ,Tipo_Imp ,Descripcion ,Telef,Fecha_Hora ,Producto  ,Num_Llamadas,Base_Imp"
                     tabla = tabla & ",Porcentaje ,Importe ,Descuento ,Importe_Neto  ,Fecha_Inicio ,Fecha_Fin) "
        
                'NUEVO Dic 2014
                Case "DCFC_DAT":
                    mFicheroCorto = ""
                    
                'Seprtiembre 2019
                'Seran igual que el fichero DNTV_CAB
                Case "DNAD_DAT"
                    ProcesarDntvCab
                    tabla = "Dntv_Cab(ciclo,Fact_No,Telef_Ext ,Fecha_Hora ,No_Recep ,Tipo ,Destino ,Duracion ,Durac_Ses "
                    tabla = tabla & ",Volumen,Pasos ,Tipo_Fact ,Tarifa,Descripcion ,Canal ,Ext_Origen ,Importe ,Importe_Neto) "
                Case Else ' desconocido
                    '-- Registro no contemplado
                    '
                    Debug.Print leido
                    If Mid(leido, 6, 3) = "CAB" Then
                        'Son las cabeceras
                        '
                         
                         If CadenaInsert <> "" Then
                            conn.Execute CadenaInsert
                            mFicheroCorto = ""
                         End If
                         
                         
                         CadenaInsert = ""
                         
                    Else
                        MsgBox "El registro TIPO: " & Left(leido, 3) & " no está contemplado en la aplicación.", vbCritical
                    End If
                    
                    mFicheroCorto = ""
            End Select
            
            If mFicheroCorto <> "" Then
                    
                
                
                
                
                If CadenaInsert = "" Then
                    CadenaInsert = "insert into vodafonetxt." & tabla & " VALUES "
                Else
                    CadenaInsert = CadenaInsert & ", "
                End If
                CadenaInsert = CadenaInsert & mFicheroCorto & ")"
                
                If Len(CadenaInsert) > 50000 Then
                  
                
                    conn.Execute CadenaInsert
                    CadenaInsert = ""
                    mFicheroCorto = ""
                End If
            End If
        End If
    
    Loop
    
    If CadenaInsert <> "" Then conn.Execute CadenaInsert
    
    Close #NF
    mFicheroCorto = GuardarNombreFichero
    
    
    
    
    'Haremos las comprobaciones aqui
    Set miRsAux = New ADODB.Recordset
    LB.Caption = "Comprobar numeros"
    LB.Refresh
    Espera 1
    IntentaEliminarFicheroLog  'por si acaso existe
    
    
    'Telefonos "RAROS que viene"
    Dim Rtr As ADODB.Recordset
    Set Rtr = New ADODB.Recordset
    
    CadenaInsert = "select * from vodafonetxt.cambionumeros WHERE false" 'Por si no esta creada la estrucutura
    If ejecutar(CadenaInsert, True) Then
        CadenaInsert = "select * from vodafonetxt.cambionumeros"
        Rtr.Open CadenaInsert, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rtr.EOF
            LB.Caption = "Cambio numeros:" & Rtr!tfno_fich & " -> " & Rtr!tfno_real
            LB.Refresh
            
            'ACtualiza N tablas
            tabla = "dcfd_cab|dcfm_cab|dcft_cab|dcfv_cab|ddnl_cab|dcul_cab|dntv_cab|dntd_cab|"
            For NF = 1 To 8
                leido = "telef"
                If NF = 7 Or NF = 8 Then leido = leido & "_ext"
                
                
                CadenaInsert = RecuperaValor(tabla, NF)
                CadenaInsert = "UPDATE vodafonetxt." & CadenaInsert & " SET " & leido & "=" & DBSet(Rtr!tfno_real, "T")
                
                CadenaInsert = CadenaInsert & " WHERE " & leido & " = " & DBSet(Rtr!tfno_fich, "T")
                
                ejecutar CadenaInsert, False
            
            Next
            Rtr.MoveNext
        Wend
    
    End If
    Set Rtr = Nothing
    LB.Caption = "Comprobaciones"
    LB.Refresh
    
    NF = FreeFile
    GuardarNombreFichero = App.Path & "\emitefac.log"
    Open GuardarNombreFichero For Output As #NF
    
    i = 0
    If Not ComprobarCuotasVodafone(LB) Then i = 1
    
    If Not ComprobarTraficoVodafone(LB, False) Then i = 1
    
    If Not ComprobarNumerosDeTelefonoVODAFONE(LB) Then i = 1
    
    
    
    
    Close #NF
    LB.Caption = "Finalizando importacion"
    Set miRsAux = New ADODB.Recordset
    
    If i = 0 Then
        IntentaEliminarFicheroLog
        cargarBaseDatosVODAFONENormal = True
    Else
        MsgBox "Errores.  Vea el archivo: " & GuardarNombreFichero, vbExclamation
        Shell "notepad " & GuardarNombreFichero, vbMaximizedFocus
    End If
    Exit Function
    
ecargarBaseDatosVODAFONE:
    MuestraError Err.Number, Err.Description
    
End Function


Private Sub IntentaEliminarFicheroLog()
    On Error Resume Next
    Kill App.Path & "\emitefac.log"
    Err.Clear
End Sub

Private Sub BorrarDatosTMPVodafone(VariabelCadena As String, ByRef L As Label)
Dim J As Integer
    
    VariabelCadena = "avpg_cab|dcfd_cab|dcfc_cab |dcfm_cab |dcft_cab |dcfv_cab |dcuc_cab |dcul_cab |ddnc_cab |ddnl_cab|"
    VariabelCadena = VariabelCadena & "dntd_cab |dntv_cab |doim_cab |dsva_cab |ecta_cab|fac_cab|"
    VariabelCadena = VariabelCadena & "idcl_cab |rcda_cab |rcfe_cab|rcll_cab |rcme_cab |rctl_cab |rdes_cab |rimp_cab|"
    
    J = 1
    While J <> 0
        J = InStr(1, VariabelCadena, "|")
        If J <> 0 Then
            L.Caption = "Tabla: " & Trim(Mid(VariabelCadena, 1, J - 1))
            L.Refresh
            conn.Execute "DELETE FROM vodafonetxt." & Trim(Mid(L.Caption, 7))
            VariabelCadena = Mid(VariabelCadena, J + 1)
        End If
    Wend
End Sub

Private Function TieneBDVodafone() As Boolean
    On Error Resume Next
    conn.Execute "DELETE FROM vodafonetxt.avpg_cab"
    If Err.Number <> 0 Then
        Err.Clear
        TieneBDVodafone = False   'no tiene vodafone
    Else
        TieneBDVodafone = True
    End If
    
End Function


'Procesar lineas fichero VODAFONE


Private Sub ProcesaFacCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(4))
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(7), "T")
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(9))
    mFicheroCorto = mFicheroCorto & "," & Val(Valores(10))
    

End Sub
Private Sub ProcesaIdclCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(7), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(8), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(9), "T")
   

End Sub

Private Sub ProcesaAvpgCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(7), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(8), "T")
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(9))

End Sub
Private Sub ProcesaRcfeCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(5))

End Sub
Private Sub ProcesaRimpCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(3))
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(5))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(6))
    
End Sub
Private Sub ProcesaEctaCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(3))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(4))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(5))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(6))

End Sub
Private Sub ProcesaRcllCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(9))
    
End Sub
Private Sub ProcesarRcmeCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(9))
    
End Sub
Private Sub ProcesarRcdaCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(9))
   
End Sub
Private Sub ProcesarRctlCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(6))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(8))
    
End Sub
Private Sub ProcesarRdesCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(3))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(4))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(5))

End Sub
Private Sub ProcesarDcfcCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(4))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(5))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(6))
    
End Sub
Private Sub ProcesarDcfvCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(9))
   
End Sub


Private Sub ProcesarDcfmCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(9))

End Sub

Private Sub ProcesarDcfdCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(9))
    
End Sub
Private Sub ProcesarDcftCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(9))

End Sub
Private Sub ProcesarDcucCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(5))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(6))
    If Valores(7) = "" Then Valores(7) = Format(Now, "dd/mm/yyyy")
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(7))
    If Valores(8) = "" Then Valores(7) = Format(Now, "dd/mm/yyyy")
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(8))

End Sub
Private Sub ProcesarDculCab()

    
    mFicheroCorto = ""
    If Valores(4) = "" Then
        If Valores(5) = "0.0000" Then
            If Valores(5) = "0.0000" Then Exit Sub
        End If
    End If
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(5))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(6))
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(8))


End Sub
Private Sub ProcesarDdncCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(5))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(6))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(9))
    
End Sub
Private Sub ProcesarDdnlCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(5))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(6))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(7))
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(8))
    mFicheroCorto = mFicheroCorto & "," & ProcesarFecha(Valores(9))

End Sub
Private Sub ProcesarDntvCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & ProcesarFechayHora(Valores(4))
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(7), "T")
    
    If (Trim(Valores(8)) = "") Then Valores(8) = 0
    mFicheroCorto = mFicheroCorto & "," & Val(Valores(8))
    If (Trim(Valores(9)) = "") Then Valores(9) = 0
    mFicheroCorto = mFicheroCorto & "," & Val(Valores(9))
    If (Trim(Valores(10)) = "") Then Valores(10) = 0
    mFicheroCorto = mFicheroCorto & "," & Val(Valores(10))

    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(11), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(12), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(13), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(14), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(15), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(16), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(17))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(18))
    
End Sub

Private Sub ProcesarDsvaCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & ProcesarFechayHora(Valores(4))
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(7), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(8), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(9), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(10), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(11), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(12), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(13), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(14), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(15), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(16), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(17))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(18))
   
End Sub
Private Sub ProcesarDntdCab()

    
    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & ProcesarFechayHora(Valores(4))
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(7), "T")

    If (Trim(Valores(8)) = "") Then Valores(8) = 0
    mFicheroCorto = mFicheroCorto & "," & Val(Valores(8))
    If (Trim(Valores(9)) = "") Then Valores(9) = 0
    mFicheroCorto = mFicheroCorto & "," & Val(Valores(9))
    If (Trim(Valores(10)) = "") Then Valores(10) = 0
    mFicheroCorto = mFicheroCorto & "," & Val(Valores(10))

    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(11))
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(12), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(13), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(14), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(15), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(16), "T")
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(17))
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(18))
    
End Sub

Private Sub ProcesarDoimCab()
Dim s As String

    mFicheroCorto = "(" & DBSet(Valores(1), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(2), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(3), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(4), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(5), "T")
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(6), "T")
    s = DBLet(Valores(7), "T")
    s = Replace(s, "'", "")
    '//Justo para esta fecha hora, la fecha viene en formato MYSQL
    '//2013-04-12 17:56  -> 20130412 17:56
    s = Replace(s, "-", "")
    Valores(7) = s
    mFicheroCorto = mFicheroCorto & "," & ProcesarFechayHora(s)
    mFicheroCorto = mFicheroCorto & "," & DBSet(Valores(8), "T")
    
    If (Trim(Valores(9)) = "") Then Valores(9) = 0
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(9))
    
    If (Trim(Valores(10)) = "") Then Valores(10) = 0
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(10))
    If (Trim(Valores(11)) = "") Then Valores(11) = 0
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(11))
    
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(12))

    If (Trim(Valores(13)) = "") Then Valores(13) = 0
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(13))
    
    mFicheroCorto = mFicheroCorto & "," & ValorDecimal(Valores(14))
    
    
    If (Trim(Valores(15)) <> "") Then
        
       
        mFicheroCorto = mFicheroCorto & ",'" & Valores(15) & " 00:00'"
   
    Else
        mFicheroCorto = mFicheroCorto & "," & ProcesarFechayHora(Valores(7))
    End If
    
    If (Trim(Valores(16)) <> "") Then
        mFicheroCorto = mFicheroCorto & ",'" & Valores(16) & " 00:00'"
    Else
        mFicheroCorto = mFicheroCorto & "," & ProcesarFechayHora(Valores(7))
    End If

End Sub




Private Function ProcesarFecha(Valor As String) As String
    
    
    ProcesarFecha = "'" & Mid(Valor, 1, 4) & "-" & Mid(Valor, 5, 2) & "-" & Mid(Valor, 7, 2) & "'"
    
End Function

Private Function ProcesarFechayHora(Valor As String) As String
    
    
    ProcesarFechayHora = "'" & Mid(Valor, 1, 4) & "-" & Mid(Valor, 5, 2) & "-" & Mid(Valor, 7, 2) & " "
    
    ProcesarFechayHora = ProcesarFechayHora & Trim(Mid(Valor, 10, 5)) & ":00'"
    'return new DateTime(val(valor.Substring(0, 4))
    '    , val(valor.Substring(4, 2))
    '    , val(valor.Substring(6, 2))
    '    , val(valor.Substring(9, 2))
    '    , val(valor.Substring(12, 2))
    '    ,0)
End Function

Private Function ValorDecimal(Valor As String) As String
    
    'valor = valor.Replace(".", ",")
    'return Decimal.Parse(valor)
    If Valor = "" Then Valor = "0"
    ValorDecimal = Valor
End Function


Private Function ComprobarCuotasVodafone(ByRef L As Label) As Boolean
Dim C As String
Dim i As Integer

    ComprobarCuotasVodafone = True 'en ppio estan bien
    
    CadenaInsert = "select tipo_cuota from vodafonetxt.Dcul_Cab where fact_no='" & mFicheroCorto & "' group by 1"
    L.Caption = "Obtener cuotas"
    L.Refresh
    miRsAux.Open CadenaInsert, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaInsert = ""
    While Not miRsAux.EOF
        CadenaInsert = CadenaInsert & Mid(miRsAux!tipo_cuota, 1, 50) & "|"
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    NumRegElim = 0
    
    While CadenaInsert <> ""
        i = InStr(1, CadenaInsert, "|")
        If i = 0 Then
            If CadenaInsert <> "" Then Err.Raise 513, , "Comprobando cuotas vodafone(1)"
        Else
            C = Mid(CadenaInsert, 1, i - 1)
            L.Caption = C
            L.Refresh
            
            CadenaInsert = Mid(CadenaInsert, i + 1)
            miRsAux.Open "Select * from Tel_desc_cuotas where DescCuota=" & DBSet(C, "T", "N"), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                ComprobarCuotasVodafone = False
                
                If NumRegElim = 0 Then
                    'Vamos a ver que cuota le toca
                    miRsAux.Close
                    
                    miRsAux.Open "select substring(codcuota,5) from Tel_desc_cuotas where codcuota like 'VODC%' order by 1 desc", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
                        
                End If
                NumRegElim = NumRegElim + 1
                
                
                'INSERTAMOS LA CUOTA
                Print #NF, "Cuota: " & C
                C = "('VODC" & Format(NumRegElim, "0000") & "'," & DBSet(C, "T") & ",0,0)"
                C = "INSERT INTO tel_desc_cuotas(CodCuota,DescCuota,Porcentaje,PorcentajeOperador) VALUES " & C
                conn.Execute C
            Else
                'ok
            End If
            miRsAux.Close
        End If
    Wend
    
End Function


Private Function ComprobarTraficoVodafone(ByRef L As Label, DesdeRefacturado As Boolean) As Boolean
Dim C As String
Dim Conceptos As String
Dim i As Integer

'Borrar cuando se pueda
'Dim Para_tel_conceptosllamadas As String
'Dim Aux2 As String


    ComprobarTraficoVodafone = True 'en ppio estan bien
    
    
'    'AVANCOOP
'    'Para taxco, los consumos tambien se pueden refacturar. Con lo cual, los generará autmaticamente
'    Para_tel_conceptosllamadas = ""
        
    
    
    
    
    
    
    
    Conceptos = "|"
    For i = 1 To 4
    
        CadenaInsert = RecuperaValor("Dntv_Cab|Dsva_Cab|Dntd_cab|Doim_Cab|", i)
        
        L.Caption = "Comprobar concepto " & CadenaInsert
        L.Refresh
        
        CadenaInsert = "select " & IIf(i = 4, "Producto", "tipo") & " from vodafonetxt." & CadenaInsert & " where fact_no='" & mFicheroCorto & "' group by 1"
   
        miRsAux.Open CadenaInsert, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not miRsAux.EOF
            If InStr(1, Conceptos, "|" & miRsAux.Fields(0) & "|") = 0 Then Conceptos = Conceptos & Mid(miRsAux.Fields(0), 1, 50) & "|"
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    Next i
    
    NumRegElim = 0
    
    If Conceptos <> "|" Then Conceptos = Mid(Conceptos, 2)
    While Conceptos <> ""
        'Debug.Print Conceptos
        i = InStr(1, Conceptos, "|")
        If i = 0 Then
            If Conceptos <> "" Then Err.Raise 513, , "Comprobando cuotas trafico vodafone(1)"
        Else
            
            C = Mid(Conceptos, 1, i - 1)
            L.Caption = C
            L.Refresh
            Conceptos = Mid(Conceptos, i + 1)
            If C <> "" Then
               
 '               If DesdeRefacturado Then
 '                   Aux2 = DevuelveDesdeBD(conAri, "Tipo_de_trafico", "tel_conceptosllamadas", "Tipo_de_trafico", C, "T")
 '                   If Aux2 = "" Then
 '                       'operadora,Codigo_de_tipo_de_trafico,Tipo_de_trafico,refacturar,TraficoSegundos
 '                       Aux2 = " (3,"
 '                       If InStr(1, UCase(C), "SMS") > 0 Then
 '                           Aux2 = Aux2 & "'SMS'," & DBSet(C, "T") & ",0,0)"
 '                       Else
 '                           Aux2 = Aux2 & "'Mov'," & DBSet(C, "T") & ",0,1)"
 '                       End If
 '                       If Para_tel_conceptosllamadas <> "" Then Para_tel_conceptosllamadas = Para_tel_conceptosllamadas & ","
 '                       Para_tel_conceptosllamadas = Para_tel_conceptosllamadas & Aux2
 '
 '                   End If
 '               End If
                miRsAux.Open "Select * from Tel_desc_consumos where DescTipoTrafico =" & DBSet(C, "T"), conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    ComprobarTraficoVodafone = False
                    
                    If NumRegElim = 0 Then
                        'Vamos a ver que cuota le toca
                        miRsAux.Close
                        
                        miRsAux.Open "select substring(CodTipoTrafico,5) from Tel_desc_consumos where CodTipoTrafico like 'VODT%' order by 1 desc", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If Not miRsAux.EOF Then NumRegElim = Val(miRsAux.Fields(0))
                            
                    End If
                    NumRegElim = NumRegElim + 1
                    
                    
                    'INSERTAMOS LA CUOTA de trafico
                    Print #NF, "Trafico: " & C
                    C = " VALUES ('VODT" & Format(NumRegElim, "0000") & "',0,0,0," & DBSet(C, "T") & ",0)"
                    C = "INSERT INTO tel_desc_consumos(CodTipoTrafico,Porcentaje,Importe,Grupo,DescTipoTrafico,PorcentajeOperador) " & C
                    
                    conn.Execute C
                Else
                    'ok
                End If
                miRsAux.Close
            End If
        End If
    Wend
  '  If Para_tel_conceptosllamadas <> "" Then
  '      C = "INSERT IGNORE INTO tel_conceptosllamadas(operadora,Codigo_de_tipo_de_trafico,Tipo_de_trafico,refacturar,TraficoSegundos) VALUES " & Para_tel_conceptosllamadas
  '      conn.Execute C
  '  End If
End Function




Private Function ComprobarNumerosDeTelefonoVODAFONE(ByRef L As Label) As Boolean
Dim RT As ADODB.Recordset

    ComprobarNumerosDeTelefonoVODAFONE = True 'en ppio estan bien
        
    CadenaInsert = "select mid(telef,1,9) from vodafonetxt.dcul_cab where fact_no='" & mFicheroCorto & "' group by 1"
    L.Caption = "Obtener telefonos"
    L.Refresh
    Set RT = New ADODB.Recordset
    RT.Open "Select * from sclientfno", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    miRsAux.Open CadenaInsert, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        L.Caption = miRsAux.Fields(0)
        L.Refresh
        RT.Find "idtelefono = '" & miRsAux.Fields(0) & "'", , adSearchForward, 1
        If RT.EOF Then
            If CadenaInsert <> "" Then
                'Primeraa linea
                Print #NF, "": Print #NF, "": Print #NF, "": Print #NF, "TELEFONO NO EXISTEN"
                CadenaInsert = ""
            End If
            Print #NF, ".- " & miRsAux.Fields(0)
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    RT.Close
    Set RT = Nothing
    
    If CadenaInsert = "" Then ComprobarNumerosDeTelefonoVODAFONE = False
    
    

End Function









Private Function EmitirFacturasVODAFONE(mfichero As String, mfecha As Date, ByRef Lbl As Label, LaCompanyia As Byte) As Boolean
Dim mSQL As String
Dim vC As CTiposMov
Dim vI As CTiposMov
Dim ContadorFAI_Incial  As Long
Dim mValorIva As Currency
Dim mFactura3 As Long  ' lleva el número de factura en proceso
Dim mSerie As String
Dim Nombre As String
Dim Apell1 As String
Dim Apell2 As String
Dim RTef As ADODB.Recordset

'ENERO 2015
'Hay cutoas refacturables en VODAFONE
Dim RsCuotasRefacturar As ADODB.Recordset

Dim TotalFactura  As Currency
Dim TotalConsumo As Currency
Dim TotalConsumoExento As Currency
Dim TotalDescuentos As Currency
Dim TotalCuotas As Currency
Dim TotalPrevio As Currency
Dim ImporteAjusteMinimo As Currency
Dim ImporteparaElCalculoDelMinimo As Currency
Dim i As Integer
Dim PorTelefono As String
Dim TelefonosFactura As String
Dim AntCliente As Long
Dim AntAgrupacion As Integer
Dim CreaFac As Boolean
Dim AntTelefono As String
Dim ColTelefonos As Collection
Dim numAgrupados As Integer
Dim J As Integer


    On Error GoTo eEmitirFacturasVODAFONE
    
    NF = -1
    EmitirFacturasVODAFONE = False    'false es que va bien que no hya errores
    
    '-- Por si nos pasan ruta completa modificamos el nombre de fichero
    mfichero = Replace(mfichero, ".dat", "")
    If Len(mfichero) > 12 Then mfichero = Right(mfichero, 12)
    mfichero3 = mfichero
    
    
    Set miRsAux = New ADODB.Recordset
    Set RTef = New ADODB.Recordset
    Set RsCuotasRefacturar = New ADODB.Recordset
    
    
    If False Then
        mSQL = "select distinct(substring(telef,1,9)) elTelefono from vodafonetxt.dcul_cab"
        mSQL = mSQL & " where  not substring(telef,1,9) IN (select IdTelefono from sclientfno)"
        i = 0
        RTef.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        mSQL = ""
        While Not RTef.EOF
            mSQL = mSQL & "-" & RTef!eltelefono
            i = i + 1
            RTef.MoveNext
        Wend
        RTef.Close
        
        If i > 0 Then Err.Raise 513, , "Telefonos NO existe:" & vbCrLf & mSQL
        
    End If
    
    
    Lbl.Caption = "Leyendo datos telefono-agrupacion"
    Lbl.Refresh
    
    
    Set vC = New CTiposMov
    If Not vC.Leer("FAT") Then Exit Function
    
    
    Set vI = New CTiposMov                          'LAS INTERNAS
    If Not vI.Leer("FAI") Then Exit Function
    ContadorFAI_Incial = vI.Contador
    vI.Contador = vI.Contador + 1000  'para evitar problemas luego

        
    mSQL = DevuelveDesdeBD(conAri, "codigiva", "sartic", "codartic", vParamAplic.ArtiTelefonia, "T")
    mSQL = DevuelveDesdeBD(conConta, "porceiva", "tiposiva", "codigiva", mSQL)
    If mSQL = "" Then Err.Raise 513, , "Error obteniendo valor IVA"
        
    mValorIva = CCur(mSQL)
        
    'Facturacion agrupada antes-despes
    'mSQL = "select distinct(substring(telef,1,9)) elTelefono from vodafonetxt.dcul_cab order by 1"
    mSQL = "select idtelefono, telef,codclien,agrupacion from vodafonetxt.dcul_cab"
    mSQL = mSQL & " left join sclientfno on substring(telef,1,9) = IdTelefono order by codclien,agrupacion,idtelefono"


    
    

    IntentaEliminarFicheroLog
    NF = FreeFile()
    Open App.Path & "\emitefac.log" For Output As #NF
    i = 0
    RTef.Open mSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    ' TelefonosFactura AntCliente Agrupa
    Set ColTelefonos = New Collection
    AntCliente = -1
    AntTelefono = ""
    TelefonosFactura = ""
    
    While Not RTef.EOF
    
        
    
        If IsNull(RTef!codClien) Then Err.Raise 513, , "NO VINCULADO : " & DBLet(RTef!telef, "T")
            
        
        If RTef!codClien <> AntCliente Then
            If AntCliente >= 0 Then
                mSQL = AntTelefono & "|" & numAgrupados & "|1|" & TelefonosFactura & "|"
                ColTelefonos.Add mSQL
            End If
            AntCliente = RTef!codClien
            TelefonosFactura = RTef!idtelefono
            AntAgrupacion = RTef!Agrupacion
            AntTelefono = RTef!idtelefono
            numAgrupados = 1
        Else
            If RTef!idtelefono = AntTelefono Then
                'no hacemos nada
                
            Else
                If AntAgrupacion = 0 Then
                    'Se inserta directemente para facturar
                    mSQL = AntTelefono & "|" & numAgrupados & "|0|" & TelefonosFactura & "|"
                    ColTelefonos.Add mSQL
                    AntTelefono = RTef!idtelefono
                    AntAgrupacion = RTef!Agrupacion
                    TelefonosFactura = RTef!idtelefono
                    numAgrupados = 1
                Else
                    If RTef!Agrupacion = AntAgrupacion Then
                        'Insertamos el que hay
                        TelefonosFactura = TelefonosFactura & "@" & RTef!idtelefono
                        AntTelefono = RTef!idtelefono
                        numAgrupados = numAgrupados + 1
                    Else
                        'Añadmios el telefono
                        Stop
                        'hemos pasado de agrupacion 1 a cero. Eso no pued ser por el SQL
                        
                    End If
                End If
            End If
        End If
        
        RTef.MoveNext
    Wend
    RTef.Close
    
    'El ultimo
    If AntTelefono = "" Then Err.Raise 513, , "No existe ultimo telefono facturable"
    mSQL = IIf(numAgrupados > 1, 1, 0)
    mSQL = AntTelefono & "|" & numAgrupados & "|" & mSQL & "|" & TelefonosFactura & "|"
    ColTelefonos.Add mSQL
        
        
        
        
        
        
    Lbl.Caption = "Insertando en cab_tel"
    Lbl.Refresh
    DoEvents
        
        
        
    For i = 1 To ColTelefonos.Count
        If (i Mod 20) = 1 Then DoEvents
        
        
        mSQL = ColTelefonos.Item(i)
        
        AntTelefono = RecuperaValor(mSQL, 2)
        numAgrupados = CInt(AntTelefono)
        AntAgrupacion = RecuperaValor(mSQL, 3)
        If CInt(AntAgrupacion) > 1 Then Err.Raise 513, , "Error leyendo TAG agrupacion: " & mSQL & " (" & i & ")"
        
        
        AntTelefono = RecuperaValor(mSQL, 1)
        TelefonosFactura = RecuperaValor(mSQL, 4)
        If TelefonosFactura = "" Then Err.Raise 513, , "Error leyendo telefonos a facturar cliente : " & mSQL & " (" & i & ")"
        TelefonosFactura = TelefonosFactura & "|"
        If numAgrupados > 1 Then TelefonosFactura = Replace(TelefonosFactura, "@", "|")
                  
        Lbl.Caption = i & "/" & ColTelefonos.Count & " " & AntTelefono    ' RTef!eltelefono
        Lbl.Refresh
        
        
        'Cogera los datos de la sclien
        mSQL = "select nifclien nifclien,nomclien,domclien,codpobla,pobclien,proclien,operador,CodDirec,sclientfno.codclien,clienppal"
        mSQL = mSQL & " from sclien,sclientfno  where sclien.codclien=sclientfno.codclien AND "
        mSQL = mSQL & " IdTelefono = " & mdbFac2.texto(AntTelefono)
        
        miRsAux.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE SER EOF
        If miRsAux.EOF Then Err.Raise 513, , "No existe NIF para telefono: " & AntTelefono
        If IsNull(miRsAux!nifClien) Then Err.Raise 513, , "NIF nulo para telefono: " & AntTelefono

        'AHora sera el texto que ponga en la factura
        If numAgrupados > 1 Then
            AntTelefono = Format(miRsAux!codClien, "000000") & "/" & "01"  'de momento solo acepta una agrupacion
        Else
            'Es una agrupacion del cliente. Lleva varios telefonos juntos
            'NO hacemos nada puesto que AntTelefono=telefonosfactura
            
        End If
        
        
        If miRsAux!nifClien = vParam.CifEmpresa Then
            vI.Contador = vI.Contador + 1
            mFactura3 = vI.Contador
            mSerie = Trim(vI.LetraSerie)
        Else
            vC.Contador = vC.Contador + 1
            mFactura3 = vC.Contador
            mSerie = Trim(vC.LetraSerie)
        End If
        '  mFactura2 = obtenNumeroFactura(vC.LetraSerie, mAno)  YA LO HA CONSEGUIDO ARRIBA conseguircontador
        mSQL = "insert into tel_cab_factura (Serie,Ano,NumFact,Fichero,Fecha,Telefono,NIF"
        mSQL = mSQL & ",companyia,TipoIVA , TotalPrevio,BaseImponible,Cuota,Total, EsAgupacion  ,ClienPpal"
        mSQL = mSQL & ",Apellido1,Apellido2,Nombre,Direccion,CodPostal,Poblacion,Provincia) VALUES ("
        
        
        mSQL = mSQL & mdbFac2.texto(mSerie) & ","  ' Serie
        mSQL = mSQL & mdbFac2.numero(Year(mfecha)) & "," ' Año
        mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
        
        mSQL = mSQL & mdbFac2.texto(mfichero) & "," ' Fichero
        mSQL = mSQL & mdbFac2.Fecha(mfecha) & "," ' Fecha
        mSQL = mSQL & mdbFac2.texto(AntTelefono) & "," ' Telefono   ponia RTef!eltelefono
        
        
        mSQL = mSQL & mdbFac2.texto(miRsAux!nifClien) & ","
        'compañia
        mSQL = mSQL & " 'VOD',"
        If miRsAux!nifClien = vParam.CifEmpresa Then
            mSQL = mSQL & "0," ' exento
        Else
            mSQL = mSQL & mdbFac2.numero(mValorIva) & "," ' IVA antes 18
        End If
        mSQL = mSQL & mdbFac2.numero(0) & "," ' Total previo (será recalculada)
        mSQL = mSQL & mdbFac2.numero(0) & "," ' Base Imponible (será recalculada)
        mSQL = mSQL & mdbFac2.numero(0) & "," ' Cuota (será recalculada)
        mSQL = mSQL & mdbFac2.numero(0) & "," ' Total (idem)
        
        'EsAgupacion
        mSQL = mSQL & mdbFac2.numero(IIf(numAgrupados > 1, 1, 0)) & ","
        ',ClienPpal
        If IsNull(miRsAux!clienppal) Then
            mSQL = mSQL & "null"
        Else
            mSQL = mSQL & miRsAux!clienppal
        End If
        mSQL = mSQL & "," ' Total (idem)
        
        
        
        
        
        
        
        
        
        If DBLet(miRsAux!CodDirec, "T") <> "" Then
            Apell2 = "codclien = " & miRsAux!codClien & " AND coddirec = " & miRsAux!CodDirec
            Apell1 = DevuelveDesdeBD(conAri, "nomdirec", "sdirec", Apell2 & " AND 1", "1")
            If Apell1 <> "" Then
                'TIENE 'departameto', es decir, va a otro nombre
                'cierro el rs
                
                miRsAux.Close
                
                Apell1 = "Select nomdirec as nomclien,domdirec domclien,codpobla,"
                Apell1 = Apell1 & " pobdirec pobclien,prodirec proclien"
                Apell1 = Apell1 & " from sdirec where " & Apell2
                miRsAux.Open Apell1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Apell1 = "": Apell2 = ""
                If miRsAux.EOF Then Err.Raise 513, , "Error leyendo dpto /dir : " & miRsAux.Source
            Else
                Print #NF, "Codirec: " & miRsAux!eltelefono & " " & Apell2
            End If
        End If
        
        SeparaNomclienNomApellidos miRsAux!NomClien, Nombre, Apell1, Apell2
        mSQL = mSQL & mdbFac2.texto(Apell1) & ","
        mSQL = mSQL & mdbFac2.texto(Apell2) & ","
        mSQL = mSQL & mdbFac2.texto(Nombre) & ","
        mSQL = mSQL & mdbFac2.texto(miRsAux!domclien) & ","
        mSQL = mSQL & mdbFac2.texto(miRsAux!codpobla) & ","
        mSQL = mSQL & mdbFac2.texto(miRsAux!pobclien) & ","
        mSQL = mSQL & mdbFac2.texto(miRsAux!proclien) & ""
        mSQL = mSQL & ")"
        mdbFac2.ejecutar mSQL
        miRsAux.Close
        
        'tel_cab_factura_agr (Serie,Ano,NumFact,Telefono)
        mSQL = ""
        For J = 1 To numAgrupados
            Apell1 = RecuperaValor(TelefonosFactura, J)
            If Apell1 = "" Then Err.Raise 513, , "Error leyendo telefonos agrupados cliente : " & miRsAux!codClien & " " & TelefonosFactura & " (n:" & J & ")"
            mSQL = mSQL & ", (" & mdbFac2.texto(mSerie) & ","  ' Serie
            mSQL = mSQL & mdbFac2.numero(Year(mfecha)) & "," ' Año
            mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
            mSQL = mSQL & mdbFac2.texto(Apell1) & ")"
        Next J
        mSQL = Mid(mSQL, 2) 'QUitamos la primera coma
        mSQL = "insert into tel_cab_factura_agr (Serie,Ano,NumFact,Telefono) VALUES " & mSQL
        
        mdbFac2.ejecutar mSQL
        
        
        
        
        
    Next

        
   'Ajustamos contadores
    vC.Contador = vC.Contador - 1
    vC.IncrementarContador vC.TipoMovimiento
    'Ver contador FAIs
    
    If vI.Contador <> ContadorFAI_Incial Then
        vI.Contador = ContadorFAI_Incial - 1
        vI.IncrementarContador vI.TipoMovimiento
    End If

    
    
   
    mSQL = "Select CodCuota,refacturar,valor from tel_desc_cuotas where refacturar >0 "
    RsCuotasRefacturar.Open mSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    '- Ya tenemos las facturas y ahora montamos los consumos
    mSQL = "concat(Apellido1,'',Apellido2,' ',Nombre) elNombre,telefono"
    mSQL = "select Serie,Ano,NumFact,Fichero,TipoIVA," & mSQL & ",EsAgupacion from tel_cab_factura where Fichero = " & mdbFac2.texto(mfichero)
    
    RTef.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Apell1 = DevuelveDesdeBD(conAri, "DescCuota", "tel_desc_cuotas", "CodCuota", "ARI00000", "T")
    i = 0
    While Not RTef.EOF
            
            
            
            'TelefonosFactura
            
            
            i = i + 1
            'If I = 70 Then Stop
            Lbl.Caption = "Lineas factura " & CStr(i) & " " & RTef.Fields(6)
            Lbl.Refresh
            DoEvents
            TotalFactura = 0

            
            'Si va agrupado o por un telefono solo
            EsAgrupacionTelefono = Val(RTef!EsAgupacion) = 1
            If EsAgrupacionTelefono Then
                whereTelefono = "serie = '" & RTef!Serie & "' AND ano = " & RTef!Ano & " AND numfact =" & RTef!NumFact & " AND 1"
                whereTelefono = DevuelveDesdeBD(conAri, "GROUP_CONCAT( concat('''',telefono,'''') separator ',')", "tel_cab_factura_agr", whereTelefono, "1 GROUP BY numfact")
                If whereTelefono = "" Then Err.Raise 513, , "Sin telefonos para: " & RTef!Serie & " - " & RTef!Ano & " - " & RTef!NumFact
                whereTelefono = "  #campotele# IN  (" & whereTelefono & ")  "
                
            Else
                'Lo que habia. No hacemos nada
                ' Leemos las cuotas correspondientes a ese número de teléfono
                whereTelefono = " #campotele# like '" & RTef!Telefono & "%'"
            End If

            
            
            
            
            
        
            'Cuotas
            TotalCuotas = CargarCuotasIva(RTef, RsCuotasRefacturar)
            TotalFactura = TotalFactura + TotalCuotas
            
            'Consumos
            TotalConsumo = CargarTraficoNuevo(RTef, TotalConsumoExento)
            TotalFactura = TotalFactura + TotalConsumo
            TotalFactura = TotalFactura + TotalConsumoExento
            
            'fatlta ver algun concepto con descuento
            TotalDescuentos = CargarDescuentos(RTef)
        
            'El mimino es sobre cuotas+consumo-descuentos
            ImporteparaElCalculoDelMinimo = TotalFactura - TotalDescuentos
            
            'Mimino ALZIRA
            'CalcularImporteMinimoConsumoNuevo
            'If vParamAplic.TieneTelefonia2 = 1 Then
            
            'fatta ver aqui
            'Agrupacion de telefono NO lleva consumo minimo
            If Not EsAgrupacionTelefono Then
                mSQL = Replace(whereTelefono, "#campotele#", "idtelefono")
                mSQL = "Select cuotaminima,fechaalta from sclientfno where " & mSQL
                miRsAux.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                ImporteAjusteMinimo = 0
                mValorIva = 0
                If Not miRsAux.EOF Then
                    mValorIva = CCur(RTef!TipoIVA)
                    '//MAYO 2013. Junio 2013. Si no han pasado 30 dias desde la fecha de alta
                    If Not IsNull(miRsAux!fechaalta) Then
                        If DateDiff("d", miRsAux!fechaalta, Now) > 30 Then
                            'Tiene mas de un mes de permenencia
                            
                            ImporteAjusteMinimo = CalcularImporteMinimoConsumoNuevo(RTef, CCur(DBLet(miRsAux!cuotaminima, "N")), ImporteparaElCalculoDelMinimo, Apell1)
                        End If
                    End If
                    
                Else
                    Print #NF, "No se encuentra tfno:  " & RTef!Telefono
                End If
                miRsAux.Close
            End If


            'Base Normal
            TotalPrevio = Round(TotalCuotas + TotalConsumo - TotalDescuentos, 2)
            TotalCuotas = Round(TotalPrevio + ImporteAjusteMinimo, 2)
               
            'Cuota IVA
            If RTef!Serie = Trim(vI.LetraSerie) Then
                
                TotalConsumo = 0 'INTERNA .   NO lleva IVA
            Else
                TotalConsumo = Round((TotalCuotas * mValorIva) / 100, 2)
            End If
            
            'cab.Total = cab.BaseImponible + cab.Cuota + cab.Base_exenta ;
            TotalDescuentos = TotalCuotas + TotalConsumo + TotalConsumoExento
            
            mSQL = "UPDATE tel_cab_factura set TotalPrevio=" & DBSet(TotalPrevio, "N")
            mSQL = mSQL & ", BaseImponible =" & DBSet(TotalCuotas, "N")
            mSQL = mSQL & ", Cuota =" & DBSet(TotalConsumo, "N")
            mSQL = mSQL & ", Total =" & DBSet(TotalDescuentos, "N")
            If TotalConsumoExento <> 0 Then mSQL = mSQL & ", base_exenta =" & DBSet(TotalConsumoExento, "N")
            mSQL = mSQL & " WHERE serie = '" & RTef!Serie & "' and ano = " & RTef!Ano & " and numfact =" & RTef!NumFact
            conn.Execute mSQL
            RTef.MoveNext
    Wend
    RTef.Close
        

    
    
    'AHORA HAY QUE INSERTAR EN BD telefono, los detalles y las cuotas
    VodafoneInsertarDetalleLlamadas Lbl
    
    
    
eEmitirFacturasVODAFONE:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Emit. facturas vodafone ", Err.Description
        EmitirFacturasVODAFONE = False
    End If
    Set miRsAux = Nothing
    Set RTef = Nothing
    Set RsCuotasRefacturar = Nothing
    Set vI = Nothing
    Set vC = Nothing
    Set ColTelefonos = Nothing
    If NF >= 0 Then Close #NF
    
End Function


Private Function CargarCuotasIva(ByRef RR As ADODB.Recordset, RecordSetConCuotasRefacturables As ADODB.Recordset) As Currency
Dim i As Integer
Dim Importe As Currency
Dim Cuota As String
Dim J As Integer

            CargarCuotasIva = 0
            
            
            'CadenaInsert = "Select substring(telef,1,9),substring(tipo_cuota,1,50) lacuota,sum(importe) importe from vodafonetxt.Dcul_Cab WHERE "
            'CadenaInsert = CadenaInsert & Replace(whereTelefono, "#campotele#", "substring(telef,1,9)") & " GROUP BY 1,2"
            
            CadenaInsert = "Select substring(telef,1,9),substring(tipo_cuota,1,50) lacuota,importe, date(fecha_inicio) fini,date(fecha_fin) ffin"
            CadenaInsert = CadenaInsert & " FROM vodafonetxt.Dcul_Cab WHERE "
            CadenaInsert = CadenaInsert & Replace(whereTelefono, "#campotele#", "substring(telef,1,9)")
            
            
            miRsAux.Open CadenaInsert, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            i = 0
            CadenaInsert = ""
            While Not miRsAux.EOF
                i = i + 1
                Importe = miRsAux!Importe
                
                
                If Importe <> 0 Then
                    Cuota = CodigoDecripcionCuotaVoda(0, DBLet(miRsAux!lacuota, "T"))
                    J = InStr(1, Cuota, "','")
                    'No puede ser 0
                    Cuota = Mid(Cuota, 1, J - 1)
                    J = InStr(1, Cuota, ",'")
                    Cuota = Mid(Cuota, J + 2)
                    RecordSetConCuotasRefacturables.Find "CodCuota=" & DBSet(Cuota, "T"), , adSearchForward, 1
                    If Not RecordSetConCuotasRefacturables.EOF Then
                        J = Val(RecordSetConCuotasRefacturables!refacturar)
                       
                        If J = 1 Then
                            'Refacturacion incrementando un importe
                            Importe = Importe + RecordSetConCuotasRefacturables!Valor
                            
                        ElseIf J = 2 Then
                            'PORCENTUAL
                            Importe = (miRsAux!Importe * RecordSetConCuotasRefacturables!Valor) / 100
                            Importe = miRsAux!Importe + Round(Importe, 2)
                        ElseIf J = 3 Then
                            'Precio fijado en CUOTA
                            Importe = RecordSetConCuotasRefacturables!Valor
                        Else
                            Err.Raise 513, , "Modo refacturacion INCORRECTO"
                        End If
                    End If
                End If
                'tel_lin_factura_cuotas(Serie,Ano,NumFact,NumLin,CodCuota,DescCuota,Importe,TipoIva)
                CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i
                CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(0, DBLet(miRsAux!lacuota, "T"))
                CadenaInsert = CadenaInsert & DBSet(Importe, "N") & ",0," & DBSet(miRsAux.Fields(0), "T")
                '               FechaInicio FechaFin
                CadenaInsert = CadenaInsert & "," & DBSet(miRsAux!FIni, "F") & "," & DBSet(miRsAux!Ffin, "F") & ")"
                CargarCuotasIva = CargarCuotasIva + Importe

                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If CadenaInsert <> "" Then
                CadenaInsert = Mid(CadenaInsert, 2)
                'añado el peridod
                CadenaInsert = "INSERT INTO tel_lin_factura_cuotas(Serie,Ano,NumFact,NumLin,CodCuota,DescCuota,Importe,TipoIva,telefono,FechaInicio,FechaFin) VALUES " & CadenaInsert
                conn.Execute CadenaInsert
            End If

 

End Function


Private Function CargarTraficoNuevo(ByRef RR As ADODB.Recordset, ByRef consumoSinIVA As Currency) As Currency
Dim i As Integer
Dim cad As String
Dim InsertamosCuota As Boolean
Dim AuxSQL As String
        
        consumoSinIVA = 0
        CargarTraficoNuevo = 0
        CadenaInsert = ""
        
        
        
        
        
        AuxSQL = Replace(whereTelefono, "#campotele#", "substring(telef_ext,1,9)")
        
        cad = "Select substring(telef_ext,1,9),substring(tipo,1,50) lacuota, sum(duracion) SumaDeDuracion,"
        cad = cad & "sum(volumen) SumaDeVolumen ,sum(Importe) SumaDeCoste , count(*) NumLlamadas "
        cad = cad & " from vodafonetxt.dntv_cab WHERE " & AuxSQL & " GROUP BY 1,2"
    
        i = 0
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF

            i = i + 1
            'Serie,Ano,NumFact,NumLin,CodTipoTrafico,DescTipoTrafico,
            'NumLlamadas,Unidad,Cantidad,Importe,TipoIVA)
            CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i
            CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(1, miRsAux!lacuota)
            CadenaInsert = CadenaInsert & DBSet(miRsAux!NumLlamadas, "N") & ","
            If DBLet(miRsAux!SumaDeDuracion, "N") = 0 Then
                'SOn KB
                If InStr(1, miRsAux!lacuota, "SMS") > 0 Then
                    CadenaInsert = CadenaInsert & "'SMS',1"
                Else
                    CadenaInsert = CadenaInsert & "'KBytes'," & DBSet(miRsAux!SumaDeVolumen, "N", "N")
                End If
            Else
                'Son segundos
                CadenaInsert = CadenaInsert & "'Segundos'," & DBSet(miRsAux!SumaDeDuracion, "N", "N")
            End If
            CadenaInsert = CadenaInsert & "," & DBSet(miRsAux!SumaDeCoste, "N") & ",0,'1900-01-01'," & DBSet(miRsAux.Fields(0), "T") & ")"
            CargarTraficoNuevo = CargarTraficoNuevo + miRsAux!SumaDeCoste

            miRsAux.MoveNext
        
        
        Wend
        miRsAux.Close
        
        
        
        '//- MMS
'            var rs2 = from m in ctx2.DsvaCabs
'                      Where m.TelefExt.StartsWith(Telefono)
'                      group m by new { m.TelefExt, m.FactNo, m.Tipo }
'                      into g select new
'                      {
'                          g.Key.TelefExt,
'                          g.Key.FactNo,
'                          g.Key.Tipo,
'                          SumaDeCoste = g.Sum(m => m.Importe),
'                          NumLlamadas = g.Count()
'                      };
        cad = "Select substring(telef_ext,1,9),substring(tipo,1,50) lacuota, sum(importe) SumaDeCoste,count(*) NumLlamadas "
        cad = cad & " from vodafonetxt.Dsva_Cab WHERE " & AuxSQL & "  GROUP BY 1,2"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
        

            
'                cons.TipoIVA = 0;  //IVA normal
'                cons.NumLlamadas = dr.NumLlamadas;
'                cons.CodTipoTrafico = BuscarCodigoTrafico(dr.Tipo, ctx);
'                cons.DescTipoTrafico = dr.Tipo;
'                cons.Unidad = "Mensajes";
'                cons.Cantidad = dr.NumLlamadas;
'                cons.Importe = dr.SumaDeCoste;
'                totalTrafico += cons.Importe;
'
            i = i + 1
            'Serie,Ano,NumFact,NumLin,CodTipoTrafico,DescTipoTrafico,
            'NumLlamadas,Unidad,Cantidad,Importe,TipoIVA)
            CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i
            CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(1, miRsAux!lacuota)
            CadenaInsert = CadenaInsert & DBSet(miRsAux!NumLlamadas, "N") & ","
            CadenaInsert = CadenaInsert & "'Mensajes'," & DBSet(miRsAux!NumLlamadas, "N") & "," & DBSet(miRsAux!SumaDeCoste, "N", "N")
            CadenaInsert = CadenaInsert & ",0,'1900-01-01'," & DBSet(miRsAux.Fields(0), "T") & ")"
            CargarTraficoNuevo = CargarTraficoNuevo + miRsAux!SumaDeCoste

            miRsAux.MoveNext
        
        
        Wend
        miRsAux.Close
        
        
        
        
        '//-- Mas trafico
        '    var rs3 = from l in ctx2.Dntd_cabs
        '              Where L.Telef_ext.StartsWith(Telefono)
        '              group l by new { l.Telef_ext, l.Fact_no, l.Tipo }
        '              into g select new
        '              {
        '                  g.Key.Telef_ext,
        '                  g.Key.Fact_no,
        '                  g.Key.Tipo,
        '                  SumaDeDuracion = g.Sum(l => l.Duracion_ses),
        '                  SumaDeVolumen = g.Sum(l => l.Volumen),
        '                  SumaDeCoste = g.Sum(l => l.Importe),
        '                  NumLlamadas = g.Count()
        '              };
        cad = "Select substring(telef_ext,1,9),substring(tipo,1,50) lacuota, sum(duracion) SumaDeDuracion,"
        cad = cad & "sum(volumen) SumaDeVolumen ,sum(Importe) SumaDeCoste , count(*) NumLlamadas "
        cad = cad & " from vodafonetxt.Dntd_cab WHERE " & AuxSQL & " GROUP BY 1,2"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
                'cons.NumLlamadas = dr.NumLlamadas;
                'cons.CodTipoTrafico = BuscarCodigoTrafico(dr.Tipo, ctx);
                'cons.DescTipoTrafico = dr.Tipo;
                'cons.Unidad = "KB";
                'cons.Cantidad = CntVodafoneTxt.ValorDecimal(dr.SumaDeDuracion.ToString());
                'cons.Importe = CntVodafoneTxt.ValorDecimal(dr.SumaDeCoste.ToString());
                'totalTrafico += cons.Importe;
            
            'Serie,Ano,NumFact,NumLin,CodTipoTrafico,DescTipoTrafico,
            'NumLlamadas,Unidad,Cantidad,Importe,TipoIVA)
            i = i + 1
            CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i
            CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(1, miRsAux!lacuota)
            CadenaInsert = CadenaInsert & DBSet(miRsAux!NumLlamadas, "N") & ",'KB',"
                      
            If IsNull(miRsAux!SumaDeDuracion) Then
                CadenaInsert = CadenaInsert & "1"
            Else
                CadenaInsert = CadenaInsert & DBSet(miRsAux!SumaDeDuracion, "N")
            End If
            CadenaInsert = CadenaInsert & "," & DBSet(miRsAux!SumaDeCoste, "N") & ",0,'1900-01-01'," & DBSet(miRsAux.Fields(0), "T") & ")"
            CargarTraficoNuevo = CargarTraficoNuevo + miRsAux!SumaDeCoste

            miRsAux.MoveNext
        
        
        Wend
        miRsAux.Close
            
                
        '    //DAVID
        '     //-- Mas trafico
        '     var rs4 = from ld in ctx2.DoimCabs
        '               Where ld.Telef.StartsWith(Telefono)
        '               select ld;
        '
        '
        cad = "Select substring(telef,1,9),substring(producto,1,50) lacuota,  importe , fecha_inicio,num_llamadas"
        cad = cad & " from vodafonetxt.Doim_Cab WHERE " & Replace(AuxSQL, "telef_ext", "telef") & " GROUP BY 1,2"
        miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
       
       
        '        cons.TipoIVA = 1;  //IVA exento
        '
        '        cons.NumLlamadas = (Int32)dc0.NumLlamadas;
        '        cons.CodTipoTrafico = BuscFechaInicioarCodigoTrafico(dc0.Producto , ctx);
        '        cons.DescTipoTrafico = dc0.Producto  ;
        '        cons.Unidad = "UD";
        '        cons.FechaInicio = dc0.;
        '        cons.Cantidad = CntVodafoneTxt.ValorDecimal("1");
        '        cons.Importe = dc0.Importe;
        '        consumoSinIVA += cons.Importe;
                    'Serie,Ano,NumFact,NumLin,CodTipoTrafico,DescTipoTrafico,
            'NumLlamadas,Unidad,Cantidad,Importe,TipoIVA)
            InsertamosCuota = True
            If DBLet(miRsAux!Importe, "N") = 0 Then
                If DBLet(miRsAux!lacuota, "T") = "" Then InsertamosCuota = False
            End If
            If InsertamosCuota Then
                i = i + 1
                CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i
                CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(1, miRsAux!lacuota)
                CadenaInsert = CadenaInsert & DBSet(miRsAux!num_llamadas, "N") & ","
                
                CadenaInsert = CadenaInsert & "'UD',1"
                
                CadenaInsert = CadenaInsert & "," & DBSet(miRsAux!Importe, "N") & ",1,'1900-01-01'," & DBSet(miRsAux.Fields(0), "T") & ")"
                If Not IsNull(miRsAux!Importe) Then consumoSinIVA = consumoSinIVA + miRsAux!Importe
            End If
            miRsAux.MoveNext
        
        
        Wend
        miRsAux.Close
            
        'Insertamos
        If CadenaInsert <> "" Then
            CadenaInsert = Mid(CadenaInsert, 2)
            CadenaInsert = "INSERT INTO tel_lin_factura_consumos(Serie,Ano,NumFact,NumLin,CodTipoTrafico,DescTipoTrafico,NumLlamadas,Unidad,Cantidad,Importe,TipoIVA,FechaInicio,telefono) VALUES " & CadenaInsert
            conn.Execute CadenaInsert
        End If
End Function



Private Function CodigoDecripcionCuotaVoda(Tipo As Byte, Descrip As String) As String
Dim cad As String

    If Tipo = 0 Then
        'CUOTAS
        'DescCuota CodCuota  tel_desc_cuotas
      
        cad = DevuelveDesdeBD(conAri, "CodCuota", "tel_desc_cuotas", "DescCuota", Descrip, "T")
        
        If cad = "" Then
            Print #NF, "Cuota incorrecta: " & Descrip
            cad = "INCORR"
        End If
    Else
        'tipo=1
        
        cad = DevuelveDesdeBD(conAri, "CodTipoTrafico", "tel_desc_consumos", "DescTipoTrafico", Descrip, "T")
        If cad = "" Then
            Print #NF, "Cuota trafico incorrecta: " & Descrip
            cad = "INCORR"
        End If
    End If
    CodigoDecripcionCuotaVoda = "," & DBSet(cad, "T") & "," & DBSet(Descrip, "T") & ","
End Function





  
Private Function CargarDescuentos(ByRef RR As ADODB.Recordset) As Currency
Dim i As Integer
Dim RN As ADODB.Recordset
Dim cad As String
Dim TieneDto As Boolean
Dim Importe As Currency

            CargarDescuentos = 0
            i = 0
            '// Calcular los descuentos asociados a cuotas
            'var rs = from l in ctx.Tel_lin_factura_cuotas
            '         where l.Serie == serie && l.Ano == ano & l.NumFact == numfac
            '         select l;
            
            Set RN = New ADODB.Recordset
            'Tel_lin_factura_descuento(Serie,Ano,NumFact,NumLin,Concepto,Descuento,Base,Importe
            CadenaInsert = ""
            cad = "Select codcuota,sum(importe) importe from Tel_lin_factura_cuotas WHERE serie ='" & RR!Serie & "' AND ano = " & RR!Ano & " AND numfact = " & RR!NumFact & " GROUP BY 1"
            RN.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            If Not RN.EOF Then
                cad = ""
                While Not RN.EOF
                    cad = cad & ", " & DBSet(RN!CodCuota, "T")
                    RN.MoveNext
                Wend
                RN.MoveFirst
                
                cad = Mid(cad, 2) 'No puede ser EOF
                cad = "Select * from Tel_desc_cuotas WHERE codcuota IN (" & cad & ")"
                miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
                'foreach (ModAriGesTel.Tel_lin_factura_cuota lc in rs)
                '{
                '    // Vemos si hay algun descuento posible.
                '    ModAriGesTel.Tel_desc_cuota d = (from desc in ctx.Tel_desc_cuotas
                '                                     where desc.CodCuota == lc.CodCuota
                '                                     select desc).FirstOrDefault<ModAriGesTel.Tel_desc_cuota>();
                    
                 While Not RN.EOF
                        miRsAux.Find "codcuota = " & DBSet(RN!CodCuota, "T"), , adSearchForward, 1
                    
                        TieneDto = False
                        If Not miRsAux.EOF Then
                            If DBLet(miRsAux!Porcentaje, "N") <> 0 Then TieneDto = True
                        End If
                        
                        If TieneDto Then
'                            if (d.Porcentaje != 0)
'                            {
'                                // Hay descuento
'                                decimal descuento = (lc.Importe * d.Porcentaje) / 100;
'                                ModAriGesTel.Tel_lin_factura_descuento tld = new ModAriGesTel.Tel_lin_factura_descuento();
'                                tld.Serie = serie;
'                                tld.Ano = ano;
'                                tld.NumFact = numfac;
'                                tld.NumLin = ++i;
'                                tld.Concepto = d.DescCuota;
'                                tld.Base = lc.Importe;
'                                tld.Descuento = (short)d.Porcentaje;
'                                tld.Importe = descuento;
'                                totalDescuentos += descuento;
'                               Serie,Ano,NumFact,NumLin,Concepto,Descuento,Base,Importe
                            Importe = (miRsAux!Porcentaje * RN!Importe) / 100
                            CargarDescuentos = CargarDescuentos + Importe
                            i = i + 1
                            CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i & "," & DBSet(miRsAux!DescCuota, "T")
                            CadenaInsert = CadenaInsert & "," & DBSet(miRsAux!Porcentaje, "N") & "," & DBSet(RN!Importe, "N") & "," & DBSet(Importe, "N") & ")"
                                
                        End If
                        RN.MoveNext
                Wend
                miRsAux.Close
            End If 'rn.eof
            RN.Close
            
            
            
            '// Calcular los descuentos asociados a tráfico.
            'var rs2 = from l in ctx.Tel_lin_factura_consumos
            '          where l.Serie == serie && l.Ano == ano && l.NumFact == numfac
            '          select l;
            'foreach (ModAriGesTel.Tel_lin_factura_consumo lt in rs2)
            cad = "Select CodTipoTrafico,sum(importe) importe from Tel_lin_factura_consumos WHERE serie ='" & RR!Serie & "' AND ano = " & RR!Ano & " AND numfact = " & RR!NumFact & " GROUP BY 1"
            RN.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            If Not RN.EOF Then
                cad = ""
                While Not RN.EOF
                    cad = cad & ", " & DBSet(RN!codtipotrafico, "T")
                    RN.MoveNext
                Wend
                RN.MoveFirst
                
                cad = Mid(cad, 2) 'No puede ser EOF
                cad = "Select * from Tel_desc_consumos WHERE CodTipoTrafico IN (" & cad & ")"
                miRsAux.Open cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            
                
            
            
                '// Comprobamos is hay algún descuento de tráfico disponible.
                'ModAriGesTel.Tel_desc_consumo dc = (from des in ctx.Tel_desc_consumos
                '                                    where des.CodTipoTrafico == lt.CodTipoTrafico
                '                                    select des).FirstOrDefault<ModAriGesTel.Tel_desc_consumo>();
                While Not RN.EOF
                        miRsAux.Find "CodTipoTrafico = " & DBSet(RN!codtipotrafico, "T"), , adSearchForward, 1
                    
                        TieneDto = False
                        If Not miRsAux.EOF Then
                            If DBLet(miRsAux!Porcentaje, "N") > 0 Then TieneDto = True
                        End If
                        
                        If TieneDto Then
                            
                            'if (dc.Porcentaje != 0)
                            '{
                            '    // Hay descuento
                            '    decimal descuento = (lt.Importe * dc.Porcentaje) / 100;
                            '    ModAriGesTel.Tel_lin_factura_descuento tld = new ModAriGesTel.Tel_lin_factura_descuento();
                            '    tld.Serie = serie;
                            '    tld.Ano = ano;
                            '    tld.NumFact = numfac;
                            '    tld.NumLin = ++i;
                            '    tld.Concepto = dc.DescTipoTrafico;
                            '    tld.Base = lt.Importe;
                            '    tld.Descuento = (short)dc.Porcentaje;
                            '    tld.Importe = descuento;
                            Importe = (miRsAux!Porcentaje * RN!Importe) / 100
                            CargarDescuentos = CargarDescuentos + Importe
                            i = i + 1
        
                            CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i & "," & DBSet(miRsAux!DescTipoTrafico, "T")
                            CadenaInsert = CadenaInsert & "," & DBSet(miRsAux!Porcentaje, "N") & "," & DBSet(RN!Importe, "N") & "," & DBSet(Importe, "N") & ")"
                        End If
                        
                        RN.MoveNext
                Wend
                miRsAux.Close
            End If 'rn.eof
            RN.Close

            If CadenaInsert <> "" Then
                CadenaInsert = Mid(CadenaInsert, 2)
                CadenaInsert = "INSERT INTO Tel_lin_factura_descuentos(Serie,Ano,NumFact,NumLin,Concepto,Descuento,Base,Importe) VALUES " & CadenaInsert
                conn.Execute CadenaInsert
            End If
        
End Function





'-----------------------------------------------------------------
Public Sub EliminarTodoElFichero(mfich As String, ByRef L As Label)
    L.Caption = "tmptfno"
    L.Refresh
    borraDatos mfich
    
    L.Caption = "vodtmp"
    L.Refresh
    
    'Para que no de ERROR de vodafone
    If TieneBDVodafone Then
        CadenaInsert = ""
        BorrarDatosTMPVodafone CadenaInsert, L
    End If
        
        
    L.Caption = "Tmp facturas"
    L.Refresh
    borrarFacturas2 mfich, L


End Sub





 Private Function CalcularImporteMinimoConsumoNuevo(ByRef RR As ADODB.Recordset, ImporMin As Currency, ImporteConsumo As Currency, DescripcionCuotaMinima As String) As Currency
Dim ImporAnyadir As Currency
        'SOLO ALZIRA. La cuota minima es el concepto ARI00000
        CalcularImporteMinimoConsumoNuevo = 0
        ImporAnyadir = ImporMin - ImporteConsumo
        If ImporAnyadir > 0 Then
            
                miRsAux.Close    'estaba abierto
                
                CadenaInsert = "Select max(numlin) from tel_lin_factura_cuotas WHERE serie = '"
                CadenaInsert = CadenaInsert & RR!Serie & "' and ano = " & RR!Ano & " and numfact= " & RR!NumFact
                miRsAux.Open CadenaInsert, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                CadenaInsert = "0"
                If Not miRsAux.EOF Then CadenaInsert = DBLet(miRsAux.Fields(0), "N")
                'miRsAux.Close lo cierra la funcion MADRE
                
                CadenaInsert = Val(CadenaInsert) + 1
                
                   
                'tel_lin_factura_cuotas(Serie,Ano,NumFact,NumLin,CodCuota,DescCuota,Importe,TipoIva)
                CadenaInsert = " ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & CadenaInsert
                
                ''DescripcionCuotaMinima
                'CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(0, "ARI00000")
                CadenaInsert = CadenaInsert & ",'ARI00000'," & DBSet(DescripcionCuotaMinima, "T") & ","
                
                
                CadenaInsert = CadenaInsert & DBSet(ImporAnyadir, "N") & ",0)"
                CalcularImporteMinimoConsumoNuevo = ImporAnyadir
                
                
                CadenaInsert = "INSERT INTO tel_lin_factura_cuotas(Serie,Ano,NumFact,NumLin,CodCuota,DescCuota,Importe,TipoIva) VALUES " & CadenaInsert
                conn.Execute CadenaInsert
        End If
        
End Function


Private Sub VodafoneInsertarDetalleLlamadas(ByRef L As Label)
Dim SQ As String
Dim ID As Long

        L.Caption = "Detalle llamada"
        L.Refresh
        Set miRsAux = New ADODB.Recordset
        SQ = "Select max(id) from telefono.Detalle_de_llamadas"
        miRsAux.Open SQ, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        ID = 0
        If Not miRsAux.EOF Then ID = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        
        CadenaInsert = ""
        SQ = "Select  Fact_No  , substring(telef_ext,1,9) telef_ext ,  Tipo ,  Destino,  no_recep  ,  fecha_hora ,duracion,   Importe from vodafonetxt.dntv_cab"
        miRsAux.Open SQ, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
                
                'detalle_de_llamadas(FactNo,Numero_de_telefono,Numero_de_telefono Codigo_de_trafico
                ' Tipo_de_trafico Descripcion_tipo_de_llamada Codigo_destino
                ' Codigo_destino Tipo_destino Numero_llamado Fecha Hora_inicio Unidad_de_medida Importe
                '
'                Detalle_de_llamada dl = new Detalle_de_llamada();
'                dl.Fichero = dr.FactNo;
'                dl.Numero_de_telefono = VerdaderoNumero(dr.TelefExt);
'                dl.Codigo_de_trafico = dr.Tipo;
'                dl.Tipo_de_trafico = dr.Tipo;
'                dl.Descripcion_tipo_de_llamada = dr.Tipo;
'                dl.Codigo_destino = dr.Destino;
'                dl.Tipo_destino = dr.Destino;
'                dl.Numero_llamado = dr.NoRecep;
'                dl.Fecha = String.Format("{0:MMdd}", dr.FechaHora);
'                dl.Hora_inicio = String.Format("{0:HHmm}", dr.FechaHora);
'                dl.Unidad_de_medida = "Segundos";
'                dl.Cantidad_medida_originada = dr.Duracion;
'                dl.Importe = dr.Importe;
               
            ID = ID + 1
            SQ = ", (" & ID & "," & DBSet(miRsAux!Fact_No, "T") & "," & DBSet(miRsAux!telef_ext, "T") & "," & DBSet(miRsAux!Tipo, "T")
            SQ = SQ & "," & DBSet(miRsAux!Tipo, "T") & "," & DBSet(miRsAux!Tipo, "T") & "," & DBSet(miRsAux!Destino, "T", "S") & ","
            SQ = SQ & DBSet(miRsAux!Destino, "T", "S") & "," & DBSet(miRsAux!no_recep, "T", "S") & ",'" & Format(miRsAux!fecha_hora, "mmdd") & "','"
            SQ = SQ & Format(miRsAux!fecha_hora, "hhnn") & "','Segundos'," & DBSet(miRsAux!Duracion, "N") & "," & DBSet(miRsAux!Importe, "N") & ")"
            CadenaInsert = CadenaInsert & SQ
            If Len(CadenaInsert) > 30000 Then
                
                L.Caption = "Registro: " & ID
                L.Refresh
                DevuelveSQLDetalleLlamadas SQ
                SQ = "INSERT INTO  telefono.detalle_de_llamadas (" & SQ & ") VALUES "
                SQ = SQ & Mid(CadenaInsert, 2)
                conn.Execute SQ
                
                CadenaInsert = ""
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        If CadenaInsert <> "" Then
            
            L.Caption = "Finalizando " & ID
            L.Refresh
            DevuelveSQLDetalleLlamadas SQ
            SQ = "INSERT INTO  telefono.detalle_de_llamadas (" & SQ & ") VALUES "
            SQ = SQ & Mid(CadenaInsert, 2)
            conn.Execute SQ
        End If
        
        
        
        
        'AHora insertamos en RESUMEN de llamadas, sobre todo para los PERIODOS
        'fecha_inicio fecha_fin
        L.Caption = "Periodo de facturacion"
        L.Refresh
        SQ = "Select * from vodafonetxt.fac_cab"
        miRsAux.Open SQ, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            MsgBox "No existe registro: fac_cab. Imposible establecer periodo facturacion"
            CadenaInsert = DBSet(Now, "F") & "," & DBSet(Now, "F")
        Else
            'fecha_inicio  fecha_fin
            CadenaInsert = DBSet(miRsAux!fecha_inicio, "F") & "," & DBSet(miRsAux!fecha_fin, "F")
            Espera 1
        End If
        miRsAux.Close
        
        'mvarFichero mfichero2
        SQ = "SELECT " & DBSet(mfichero3, "T") & ",320,1," & DBSet(mfichero3, "T") & ",'',Numero_de_telefono,''," & CadenaInsert
        SQ = SQ & ",Codigo_de_trafico,Tipo_de_trafico,count(*),Unidad_de_medida,sum(Cantidad_medida_originada)"
        
        SQ = SQ & ",0,0,sum(importe),'' FROM telefono.detalle_de_llamadas WHERE  Fichero =" & DBSet(mfichero3, "T")
        SQ = SQ & " and   fecha<>'0000'"
        SQ = SQ & " group by Codigo_de_trafico,Numero_de_telefono order by Numero_de_telefono,Codigo_de_trafico"
    
        
        CadenaInsert = "INSERT INTO telefono.resumen_de_llamadas (Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura"
        CadenaInsert = CadenaInsert & ",Numero_de_prefactura , Numero_de_telefono, Numero_de_extension, Fecha_inicio_periodo"
        CadenaInsert = CadenaInsert & ",Fecha_final_periodo , Codigo_de_tipo_de_trafico, Tipo_de_trafico, Numero_de_llamadas"
        CadenaInsert = CadenaInsert & ",Unidad_de_medida,Cantidad_medida,Cantidad_franquiciada,Importe_franquiciado,Importes,Libre)"
             
        SQ = CadenaInsert & SQ
        conn.Execute SQ





'





        Set miRsAux = Nothing
End Sub


Private Sub DevuelveSQLDetalleLlamadas(ByRef C As String)
    C = "id,fichero,Numero_de_telefono, Codigo_de_trafico,Tipo_de_trafico"
    C = C & ",Descripcion_tipo_de_llamada,Codigo_destino ,Tipo_destino, Numero_llamado "
    C = C & ",Fecha, Hora_inicio ,Unidad_de_medida ,Cantidad_medida_originada,Importe"
End Sub




'*********************************************************************************************************
'*********************************************************************************************************
'
' Refacturar cuotas VODAFONE
'
'*********************************************************************************************************
'*********************************************************************************************************
Public Function RefacturarCuotasVODAFONE(mfichero As String, ByRef L As Label)
Dim vSQL As String
Dim cad As String
Dim Origen As String
Dim Rc As ADODB.Recordset
Dim i As Integer
Dim Importe2 As Currency


    L.Caption = "Refacturar cuotas VODAFONE"
    L.Refresh
    
    mFicheroCorto = Right(mfichero, 12)

    
   'tel_desc_cuotas where refacturar>0
    
    
End Function


'AVANCOOP . Es un csv
Private Function cargarBaseDatosVODAFONE_Refacturado(mfichero As String, ByRef LB As Label) As Boolean


    Dim leido As String ' registro leido del fichero
    Dim i As Long ' un contadoir de registros
    Dim GuardarNombreFichero As String
    Dim tabla As String
    Dim K As Long
    Dim Rtr As ADODB.Recordset
    Dim Aux As String
    Dim J As Integer
    Dim fec As Date
    Dim Duracion As Currency
    Dim Volumen As Currency
    Dim LlevaSaltoLinea_VBCRLF As Byte
    Dim ElFichero As Collection
    Dim CicloFacturacion As Date
    
    On Error GoTo ecargarBaseDatosVODAFONE
    
    
    cargarBaseDatosVODAFONE_Refacturado = False
    NF = -1
   
   
   
    'En el nombre tiene que venir el MES a procesar
   mFicheroCorto = Dir(mfichero, vbArchive)
   If mFicheroCorto = "" Then Err.Raise 513, , "Fichero no EXISTE"
   
   If InStr(1, LCase(mFicheroCorto), ".csv") = 0 Then Err.Raise 513, , "No es fichero csv"
   
   
   'Debe tener esta forma
   ' ..\..\*F46029260_202005*.csv
   mFicheroCorto = Mid(mFicheroCorto, 1, Len(mFicheroCorto) - 4)
   mFicheroCorto = Right(mFicheroCorto, 6)  'busco 202005
   If Not IsNumeric(mFicheroCorto) Then Err.Raise 513, , "Falta año-mes en el nombre"
   
   NF = Val(Mid(mFicheroCorto, 1, 4))
   If NF < 2020 Or NF > Year(Now) + 1 Then Err.Raise 513, , "Error año fichero YYYYMM"
   NumRegElim = NF
   
   NF = Val(Mid(mFicheroCorto, 5, 2))
   If NF < 1 Or NF > 12 Then Err.Raise 513, , "Error mes fichero MM"
   
   'Inicio fin cic
   tabla = DevuelveDesdeBD(conAri, "diaInicioFacturacion", "spara2", "1", "1", "T")
   If tabla = "" Then Err.Raise 513, , "Falta parametro dia inicio facturacion"
   
   'Cuando no va a mes natural, factura desde el anterior. Ejemplo. Si dia 15 y fichero MAYO, seria de 15 abril a 14 mayo
   If tabla <> "1" Then Err.Raise 513, , "Periodos facturacion mes completo"
   
   'En enero facturaremos diciembre del año anterior
   NF = NF - 1
   If NF <= 0 Then
        NF = 12
        NumRegElim = NumRegElim - 1
   End If
   tabla = Right("00" & tabla, 2) & "/" & Right("00" & NF, 2) & "/" & NumRegElim
   CicloFacturacion = CDate(tabla)
   
   
   mFicheroCorto = "CSV" & mFicheroCorto
   tabla = DevuelveDesdeBD(conAri, "Fecha", "tel_fichtraspasados", "Fichero", mFicheroCorto, "T")
   If tabla <> "" Then Err.Raise 513, , "El fichero fue traspasado el " & tabla
     
  
   
   
   
   
   
   
  
  
  
    LB.Caption = "Preparando datos"
    LB.Refresh
    conn.Execute "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
  
    EliminarTodoElFichero mFicheroCorto, LB

    tabla = CStr(FileLen(mfichero))
    
    NF = FreeFile()
    Open mfichero For Input As #NF
    
    
    
    '- Borramos registros del mismo fichero si los hubiese
    borraDatos mFicheroCorto
    
    
    
    '-- El fichero por lo menos existe y grabamos sus datos
    cargaDatosPrincipales mFicheroCorto
    
    
    
    '--- Para no tener que pasar variables a todas y cada un
    'ya que las funciones las he hecho copiando pegando
    'del .NET
    
    
    
    
    GuardarNombreFichero = mFicheroCorto
    
    CadenaInsert = ""
    NumRegElim = 0
    whereTelefono = ""
    
    LlevaSaltoLinea_VBCRLF = 0   '0 No sabemos si lleva o no lleva saltos de linea
    
    
    Set ElFichero = New Collection
    

    



    Do While Not EOF(NF)
        LB.Caption = "Leido " & NumRegElim & " de " & tabla
        LB.Refresh
        
        
        '##No ponemos input line
        Input #NF, leido
        
        NumRegElim = NumRegElim + Len(leido)
        If LlevaSaltoLinea_VBCRLF = 0 Then
            If InStr(leido, vbLf) > 0 Then
                LlevaSaltoLinea_VBCRLF = 1 'No lleva salto de linea, lleva vblf
            Else
                LlevaSaltoLinea_VBCRLF = 2 'Si que lleva salto de linea, luego cada lido es una linea
            End If
        End If
        
        If LlevaSaltoLinea_VBCRLF = 2 Then
            'Cada lienea la meto directo en el leido
            ElFichero.Add leido
        Else
            
            'Si el trozo leido NO lleva vblf entonces añadimos coma
            leido = leido & ","
            
            whereTelefono = whereTelefono & leido
            Do
                K = InStr(1, whereTelefono, vbLf)
                If K > 0 Then
                    Aux = Mid(whereTelefono, 1, K - 1)
                    whereTelefono = Mid(whereTelefono, K + 1)
                    ElFichero.Add Aux
                End If
            Loop Until K = 0
        End If
     Loop
     Close #NF
     
      If ElFichero.Count <= 1 Then Err.Raise 513, , "Fichero con una linea"
      'primera linea tiene los encabezados
      For K = 2 To ElFichero.Count
        
          
        
        
            LB.Caption = "Proceso lin " & K & " de " & ElFichero.Count
            LB.Refresh
        
            leido = ElFichero.Item(K)
            
            
            Valores = Split(leido, ";")
            J = UBound(Valores)
            If J < 6 Then
                Err.Raise 513, , "Menos de 6 campos en la linea del CSV"
            Else
                If J > 10 Then Err.Raise 513, , "mas de 10 campos en la linea del CSV"
            End If
            
            'En el fichero:
            'Consumo Datos 1      Consumo Voz  2
            'Cuotas  3           Consumo Mensajes  4
            If Valores(0) = "Consumo Datos" Then
                J = 1
                
            ElseIf Valores(0) = "Consumo Voz" Then
                J = 2
            
            ElseIf Valores(0) = "Cuotas" Then
                J = 3
              
            
            ElseIf Valores(0) = "Consumo Mensajes" Then
                J = 4
            
            Else
                J = 0 'MAL
                Err.Raise 513, , "Campos soportados: Consumo Datos- Consumo Voz - Cuotas - Consumo Mensajes "
            End If
            
            If J = 3 Then
                fec = Now
            Else
                Aux = Mid(Valores(3), 7, 2) & "/" & Mid(Valores(3), 5, 2) & "/" & Mid(Valores(3), 1, 4) & " " & Mid(Valores(3), 10)
                fec = CDate(Aux)
            End If
            
            'codusu,codigo1,campo1,nombre1,nombre2,nombre3,importe1,importe2,importe3,obser,fecha1
            'codusu,codigo1,campo1,nombre1,
            CadenaInsert = CadenaInsert & ", (" & vUsu.Codigo & "," & i & "," & J & "," & DBSet(Valores(1), "T") & ","
            Aux = Trim(Valores(4) & " " & Valores(5))
            'nombre2 nombre
            Valores(2) = CStr(RevisaCaracterMultibase(Valores(2)))
            CadenaInsert = CadenaInsert & DBSet(Valores(2), "T", "S") & "," & DBSet(Aux, "T", "S") & ","
            
            If J > 1 Then
                If J = 4 Then
                
                    'If Valores(6) <> "" Then Stop
                
                    Volumen = 1
                    Duracion = 0
                Else
                    Volumen = CCur(Valores(7))
                    Duracion = CCur(Valores(8))
                End If
            Else
            
                Volumen = 0
                Duracion = 0
            End If
            
            'importe1 importe2
            CadenaInsert = CadenaInsert & DBSet(Volumen, "N", "N") & "," & DBSet(Duracion, "N", "S") & ","
            'importe3  y observa fecha
            If J > 1 Then
                Aux = Trim(Valores(9) & " " & Valores(10))
            Else
                Aux = ""
            End If
            If Aux <> "" Then Aux = CStr(RevisaCaracterMultibase(Aux))
            
            'Duracion = CCur(Valores(6)) / 10000
            Duracion = CCur(Valores(6))
            If Valores(6) <> "" Then
                If Valores(6) <> "0" Then
                    If Valores(6) <> "0,0000" Then Debug.Print Valores(6)    ': Stop
                End If
            End If
            CadenaInsert = CadenaInsert & DBSet(Duracion, "N", "N") & "," & DBSet(Aux, "T", "S") & "," & DBSet(fec, "FH") & ")"
        
            If Len(CadenaInsert) > 10000 Then
            
                CadenaInsert = Mid(CadenaInsert, 2)
                CadenaInsert = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,nombre3,importe1,importe2,importe3,obser,fecha3) VALUES " & CadenaInsert
                conn.Execute CadenaInsert
                CadenaInsert = ""
            End If

    Next K
 
    Set ElFichero = Nothing
    
    If Len(CadenaInsert) > 0 Then
        CadenaInsert = Mid(CadenaInsert, 2)
        CadenaInsert = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2,nombre3,importe1,importe2,importe3,obser,fecha1)  VALUES " & CadenaInsert
        conn.Execute CadenaInsert
    End If
    
    NF = -1
    
    
    
    
    CadenaInsert = ""
    
    Set Rtr = New ADODB.Recordset
    
    
    'Primera pasada fichero:
    ' Metemos en dcuil
    
    'Resumen cuotas, consumos
    
    LB.Caption = "Datos BD vodafone dcul_cab (1/4)"
    LB.Refresh
    Espera 0.5
    
'    CadenaInsert = "insert into vodafonetxt.dcul_cab(tipo_cuota,telef,importe,fecha_inicio,fecha_fin,fact_no,descuento,dcul_cab_id,ciclo)"
'    CadenaInsert = CadenaInsert & " select nombre2 tipo_cuota,nombre1 telef ,sum(importe3) importe ,"
'    J = DiasMes(Month(CicloFacturacion), Year(CicloFacturacion))
'    fec = Right("00" & J, 2) & Format(CicloFacturacion, "/mm/yyyy")
'    CadenaInsert = CadenaInsert & Format(CicloFacturacion, "yyyymmdd") & " fechaini , "
'    CadenaInsert = CadenaInsert & Format(fec, "yyyymmdd") & " fechafin, '" & mFicheroCorto & "' factno, 0 dto ,@rownum:=@rownum+1 AS rownum,"
'    CadenaInsert = CadenaInsert & Format(CicloFacturacion, "yyyymmdd") & " ciclo"
'    CadenaInsert = CadenaInsert & " from tmpinformes,(SELECT @rownum:=0) r where codusu=" & vUsu.Codigo & "  AND campo1=3 group by 1,2"
'
    
    'Nov 2020
    CadenaInsert = "insert into vodafonetxt.dcul_cab(tipo_cuota,telef,importe,fecha_inicio,fecha_fin,fact_no,descuento,dcul_cab_id,ciclo)"
    CadenaInsert = CadenaInsert & " select nombre2 tipo_cuota,nombre1 telef ,importe3 importe ,"
    J = DiasMes(Month(CicloFacturacion), Year(CicloFacturacion))
    fec = Right("00" & J, 2) & Format(CicloFacturacion, "/mm/yyyy")
    CadenaInsert = CadenaInsert & " substring( obser,1,8) fechaini , substring(obser,10) fechafin,'"
    CadenaInsert = CadenaInsert & mFicheroCorto & "' factno, 0 dto ,@rownum:=@rownum+1 AS rownum,"
    CadenaInsert = CadenaInsert & Format(CicloFacturacion, "yyyymmdd") & " ciclo"
    CadenaInsert = CadenaInsert & " from tmpinformes,(SELECT @rownum:=0) r where codusu=" & vUsu.Codigo & " AND importe3<>0 AND campo1=3 "
    
    conn.Execute CadenaInsert
    
    LB.Caption = "Datos BD vodafone fac_cab (2/4)"
    LB.Refresh
    CadenaInsert = "INSERT INTO vodafonetxt.fac_cab ( titular,no_lineas,nif_cif,lugar_emision,fecha_inicio,fecha_fin,fecha_emision,fac_no,fac_cab_id,cuenta_no,ciclo) "
    CadenaInsert = CadenaInsert & "SELECT " & DBSet(vEmpresa.nomempre, "T") & " titular, count(*) no_lineas," & DBSet(vParam.CifEmpresa, "T") & " nif_cif,"
    CadenaInsert = CadenaInsert & DBSet(vParam.DomicilioEmpresa, "T", "N") & " lugar_emision, "
        
    J = DiasMes(Month(CicloFacturacion), Year(CicloFacturacion))
    fec = Right("00" & J, 2) & Format(CicloFacturacion, "/mm/yyyy")
    CadenaInsert = CadenaInsert & DBSet(CicloFacturacion, "F") & " fechaini , "
    CadenaInsert = CadenaInsert & DBSet(fec, "F") & " fechafin,"
    CadenaInsert = CadenaInsert & DBSet(Now, "F")
    CadenaInsert = CadenaInsert & ",'" & mFicheroCorto & "' factno,@rownum:=@rownum+1 AS rownum , nombre1 ,"
    CadenaInsert = CadenaInsert & Format(CicloFacturacion, "yyyymmdd") & " ciclo"
    CadenaInsert = CadenaInsert & " from tmpinformes,(SELECT @rownum:=0) r where codusu=" & vUsu.Codigo & "  group by nombre1"
    conn.Execute CadenaInsert
    
    LB.Caption = "Datos BD vodafone dntd_cab (3/4)"
    LB.Refresh

    
    CadenaInsert = "INSERT INTO vodafonetxt.dntd_cab(dntd_cab_id,ciclo,fact_no,telef_ext,fecha_hora,tipo,volumen,pasos,tipo_fact,importe,importe_neto) "
    CadenaInsert = CadenaInsert & " SELECT @rownum:=@rownum+1 AS rownum  ," & Format(CicloFacturacion, "'yyyymmdd'") & ",'" & mFicheroCorto & "' factno, "
    CadenaInsert = CadenaInsert & " nombre1, fecha3 , nombre2 tipo, round(importe1/1024,4) volumen,0  pasos,'KB' tipo_fact, 0 ,0  "
    CadenaInsert = CadenaInsert & " from tmpinformes,(SELECT @rownum:=0) r where codusu=" & vUsu.Codigo & "  AND campo1=1 "
    conn.Execute CadenaInsert
   
    
    
    
    
    LB.Caption = "Datos BD vodafone dntv_cab (4/4)"
    LB.Refresh
  
    CadenaInsert = "INSERT INTO vodafonetxt.dntv_cab(dntv_cab_id,ciclo,fact_no,tipo,telef_ext,no_recep"
    CadenaInsert = CadenaInsert & " ,fecha_hora,importe_neto,importe,duracion,destino) "
    CadenaInsert = CadenaInsert & " SELECT @rownum:=@rownum+1 AS rownum  ," & Format(CicloFacturacion, "'yyyymmdd'") & ",'" & mFicheroCorto & "' factno, "
    CadenaInsert = CadenaInsert & " nombre2 tipo, nombre1 telef_ext, nombre3 recep, fecha3 , importe3,importe3,coalesce(importe2,0) duracion, '' "
    CadenaInsert = CadenaInsert & " from tmpinformes,(SELECT @rownum:=0) r where codusu=" & vUsu.Codigo & "  AND campo1 in (2,4) "
    conn.Execute CadenaInsert
   
    
    
    
    
    'Haremos las comprobaciones aqui
    Set miRsAux = New ADODB.Recordset
    LB.Caption = "Comprobar numeros"
    LB.Refresh
    Espera 1
    IntentaEliminarFicheroLog  'por si acaso existe
    
    
    'Telefonos "RAROS que viene"
    Set Rtr = New ADODB.Recordset
    
    CadenaInsert = "select * from vodafonetxt.cambionumeros WHERE false" 'Por si no esta creada la estrucutura
    If ejecutar(CadenaInsert, True) Then
        CadenaInsert = "select * from vodafonetxt.cambionumeros"
        Rtr.Open CadenaInsert, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rtr.EOF
            LB.Caption = "Cambio numeros:" & Rtr!tfno_fich & " -> " & Rtr!tfno_real
            LB.Refresh
            
            'ACtualiza N tablas
            tabla = "dcfd_cab|dcfm_cab|dcft_cab|dcfv_cab|ddnl_cab|dcul_cab|dntv_cab|dntd_cab|"
            For NF = 1 To 8
                leido = "telef"
                If NF = 7 Or NF = 8 Then leido = leido & "_ext"
                
                
                CadenaInsert = RecuperaValor(tabla, NF)
                CadenaInsert = "UPDATE vodafonetxt." & CadenaInsert & " SET " & leido & "=" & DBSet(Rtr!tfno_real, "T")
                
                CadenaInsert = CadenaInsert & " WHERE " & leido & " = " & DBSet(Rtr!tfno_fich, "T")
                
                ejecutar CadenaInsert, False
            
            Next
            Rtr.MoveNext
        Wend
    
    End If
    Set Rtr = Nothing
    LB.Caption = "Comprobaciones"
    LB.Refresh
    
    NF = FreeFile
    GuardarNombreFichero = App.Path & "\emitefac.log"
    Open GuardarNombreFichero For Output As #NF
    
    NumRegElim = i 'Numreo registros en tmp
    
    J = 0
    If Not ComprobarCuotasVodafone(LB) Then J = 1
    
    If Not ComprobarTraficoVodafone(LB, True) Then J = 1
    
    If Not ComprobarNumerosDeTelefonoVODAFONE(LB) Then J = 1
    
    Close #NF
    
    
    
    
    
    'Vamos a borrar la tmpinfomres
    
    
    NF = 0
    Do
        NF = NF + 1
        LB.Caption = "Finalizando importacion " & NF & " / 10"
        LB.Refresh
        
        If NF = 1 Then
            tabla = DevuelveDesdeBD(conAri, "max(codigo1)", "tmpinformes", "codusu", vUsu.Codigo)
            If tabla = "" Then tabla = "0"
            NumRegElim = Val(tabla) + 1
            i = NumRegElim / 10
        End If
        
        conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo & " AND codigo1 <=" & CStr(NF * i)
    Loop Until NF = 10
    
    If J = 0 Then
        IntentaEliminarFicheroLog
        cargarBaseDatosVODAFONE_Refacturado = True
    Else
        MsgBox "Errores.  Vea el archivo: " & GuardarNombreFichero, vbExclamation
        Shell "notepad " & GuardarNombreFichero, vbMaximizedFocus
    End If
    
    Set ElFichero = Nothing
    
    Exit Function
    
ecargarBaseDatosVODAFONE:
    MuestraError Err.Number, , Err.Description
    Set Rtr = Nothing
    If NF > 0 Then Close #NF
    Set ElFichero = Nothing
End Function






'*********************************************************************************************************
'*********************************************************************************************************
'
' VENTA PLAZOS
'
'*********************************************************************************************************
'*********************************************************************************************************

Public Function ComprobarTelefonosVentaPlazos(ByRef L As Label, Operadora As Byte)

'
'    Dim Rs As ADODB.Recordset
'    Dim cad As String
'    Dim Aux As String
'    Dim Impor As Currency
'    Dim imp2 As Currency
'    Dim Aux2 As String
'    Dim i As Integer
'
'    'Entran con espacios en blanco
'    L.Caption = "Comprobando"
'    L.Refresh
'    cad = "update telefono.detalle_de_llamadas set Numero_llamado =trim(Numero_llamado) where fichero='" & mFicheroCorto & "'"
'    conn.Execute cad
'    Espera 0.25
'
'    'LLAMADAS DE COOPERATIVA
'    Set Rs = New ADODB.Recordset
'
'
'
'     L.Caption = "Leyendo registros llamada"
'     L.Refresh
'A10011554341
'
'
'

End Function
