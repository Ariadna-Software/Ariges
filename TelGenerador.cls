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
Private mfichero2 As String

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



'Para el proceso VODAFONE
Dim Valores() As String
Dim CadenaInsert  As String
Dim NF As Integer

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
    Dim Leido As String ' registro leido del fichero
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
    

    
    '-- El fichero por lo menos existe y grabamos sus datos
    cargaDatosPrincipales mFicheroCorto
    Do While Not EOF(NF)
        Line Input #NF, Leido
        i = i + 1
        LB.Caption = "Registro " & CStr(i)
        LB.Refresh
        If (i Mod 5) = 0 Then DoEvents
        
        '-- hacemos lo que tengamos que hacer con el registro
        Select Case Left(Leido, 3)
            Case "010" ' cabecera soporte
                cargaRegistro010 Leido
            Case "100" ' general factura
                cargaRegistro100 Leido
            Case "300" ' general telefono
                cargaRegistro300 Leido
            Case "310" ' cuotas teléfono
                cargaRegistro310 Leido
            Case "320" ' resumen llamadas
                cargaRegistro320 Leido
            Case "330" ' detalle llamadas
                cargaRegistro330 Leido
            Case "340" ' varios teléfono
                cargaRegistro340 Leido
            Case "850" ' descuentos fectura
                cargaRegistro850 Leido
            Case "990" ' totales soporte
                cargaRegistro990 Leido
            Case Else ' desconocido
                '-- Registro no contemplado
                MsgBox "El registro TIPO: " & Left(Leido, 3) & " no está contemplado en la aplicación.", vbCritical
        End Select
    Loop
    Close #NF
End Function

Private Sub cargaDatosPrincipales(Cadena As String)
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.ficheros values("
        mSQL = mSQL & .texto(Cadena) & ","  ' campo Fichero
        mSQL = mSQL & .texto(Mid(Cadena, 1, 1)) & ","   ' campo Servicio
        mSQL = mSQL & .texto(Mid(Cadena, 2, 1)) & "," ' campo Mes
        mSQL = mSQL & .texto(Mid(Cadena, 3, 1)) & ","  ' campo Ano
        mSQL = mSQL & .texto(Mid(Cadena, 4, 5)) & ","  ' campo Soporte
        mSQL = mSQL & .texto(Mid(Cadena, 10, 1)) & ","  ' campo Agrupacion
        mSQL = mSQL & .texto(Mid(Cadena, 11, 2)) & ")"  ' campo Version
        .ejecutar mSQL
    End With
End Sub


Private Sub cargaRegistro010(Cadena As String) ' cabecera de soporte
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.cabecera_de_soporte values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(Cadena, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(Cadena, 4, 6)) & ","   ' campo Secuencia de registrp
        mSQL = mSQL & .texto(Mid(Cadena, 10, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 24, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 38, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 47, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 56, 11)) & ","   ' campo CIF
        mSQL = mSQL & .texto(Mid(Cadena, 67, 65)) & ","   ' campo Nombre
        mSQL = mSQL & .texto(Mid(Cadena, 132, 40)) & ","   ' campo Domicilio
        mSQL = mSQL & .numero(Mid(Cadena, 172, 5)) & ","   ' campo Código Postal
        mSQL = mSQL & .texto(Mid(Cadena, 177, 40)) & ","   ' campo Localidad
        mSQL = mSQL & .texto(Mid(Cadena, 217, 11)) & ","   ' campo CIF
        mSQL = mSQL & .texto(Mid(Cadena, 228, 20)) & ","   ' campo Nombre
        mSQL = mSQL & .texto(Mid(Cadena, 248, 65)) & ","   ' campo Nombre
        mSQL = mSQL & .texto(Mid(Cadena, 313, 40)) & ","   ' campo Domicilio
        mSQL = mSQL & .texto(Mid(Cadena, 353, 5)) & ","   ' campo Código Postal
        mSQL = mSQL & .texto(Mid(Cadena, 358, 40)) & ","   ' campo Localidad
        mSQL = mSQL & .texto(Mid(Cadena, 398, 20)) & ","   ' campo E-mail
        mSQL = mSQL & .texto(Mid(Cadena, 418, 5)) & ","   ' campo Código de soporte
        mSQL = mSQL & .Fecha(fechatel(Mid(Cadena, 423, 8))) & ","   ' campo Fecha de Facturación
        mSQL = mSQL & .Fecha(fechatel(Mid(Cadena, 431, 8))) & ","   ' campo Fecha de Generación registro
        mSQL = mSQL & .texto(Mid(Cadena, 439, 6)) & ","   ' campo Versión del soporte
        mSQL = mSQL & .texto(Mid(Cadena, 445, 12)) & ","   ' campo Nombre del fichero
        mSQL = mSQL & .texto(Mid(Cadena, 457, 6)) & ","   ' campo Código de agrupación asociación
        mSQL = mSQL & .texto(Mid(Cadena, 463, 3)) & ","   ' campo Moneda
        mSQL = mSQL & .numero(Mid(Cadena, 466, 8), 4) & ","  ' campo Factor de conversión
        mSQL = mSQL & .texto(Mid(Cadena, 474, 20)) & ","   ' campo Tipo Soporte
        mSQL = mSQL & .texto(Mid(Cadena, 494, 207)) & ")"  ' campo Libre
        .ejecutar mSQL
    End With
End Sub


Private Sub cargaRegistro100(Cadena As String) ' Factura
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.factura values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(Cadena, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(Cadena, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(Cadena, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(Cadena, 24, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 38, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 47, 9)) & ","   ' campo Libre
        mSQL = mSQL & .Fecha(fechatel(Mid(Cadena, 56, 8))) & ","   ' campo Fecha de Factura
        mSQL = mSQL & .texto(Mid(Cadena, 64, 20)) & ","   ' campo Tipo de Factura
        mSQL = mSQL & .numero(Mid(Cadena, 84, 4)) & ","  ' campo Banco
        mSQL = mSQL & .numero(Mid(Cadena, 88, 4)) & ","   ' campo Sucursal
        mSQL = mSQL & .texto(Mid(Cadena, 92, 2)) & ","   ' campo Dígitos de control
        mSQL = mSQL & .numero(Mid(Cadena, 94, 19)) & ","   ' campo Cuenta / Nº de Tarjeta
        mSQL = mSQL & .texto(Mid(Cadena, 113, 8)) & ","   ' campo Servicio
        mSQL = mSQL & .texto(Mid(Cadena, 121, 40)) & ","   ' campo Tipo de Contrato
        mSQL = mSQL & .numero(Mid(Cadena, 161, 5)) & ","   ' campo Numero de líneas fijas
        mSQL = mSQL & .numero(Mid(Cadena, 166, 5)) & ","   ' campo Número de líneas móviles
        mSQL = mSQL & .texto(Mid(Cadena, 171, 3)) & ","   ' campo Moneda Original
        mSQL = mSQL & .numero(Mid(Cadena, 174, 16), 4) & ","  ' campo Total cuotas
        mSQL = mSQL & .numero(Mid(Cadena, 190, 16), 4) & ","  ' campo Total Servicio Medido
        mSQL = mSQL & .numero(Mid(Cadena, 206, 16), 4) & ","  ' campo Total Varios
        mSQL = mSQL & .numero(Mid(Cadena, 222, 16), 4) & ","  ' campo Total Descuentos
        mSQL = mSQL & .numero(Mid(Cadena, 238, 16), 4) & ","  ' campo Suma Total
        mSQL = mSQL & .numero(Mid(Cadena, 254, 16), 4) & ","  ' campo Total Impuestos
        mSQL = mSQL & .numero(Mid(Cadena, 270, 14), 2) & ","  ' campo Total a Pagar
        mSQL = mSQL & .texto(Mid(Cadena, 284, 6)) & ","   ' campo Impuesto
        mSQL = mSQL & .numero(Mid(Cadena, 290, 4), 2) & ","  ' campo Tipo
        mSQL = mSQL & .numero(Mid(Cadena, 294, 16), 4) & ","  ' campo Base Imponible
        mSQL = mSQL & .texto(Mid(Cadena, 310, 11)) & ","   ' campo CIF
        mSQL = mSQL & .texto(Mid(Cadena, 321, 380)) & ")"   ' campo Libre
        .ejecutar mSQL
    End With
End Sub


Private Sub cargaRegistro300(Cadena As String) ' Teléfono
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.telefono values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(Cadena, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(Cadena, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(Cadena, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(Cadena, 24, 14)) & ","   ' campo N. de PreFactura
        mSQL = mSQL & .texto(Mid(Cadena, 38, 9)) & ","   ' campo N. de telefono
        mSQL = mSQL & .texto(Mid(Cadena, 47, 9)) & ","   ' campo N. de extensión
        mSQL = mSQL & .numero(Mid(Cadena, 56, 14), 4) & ","  ' campo Total Cuotas
        mSQL = mSQL & .numero(Mid(Cadena, 70, 14), 4) & ","  ' campo Total Servicio Medido
        mSQL = mSQL & .numero(Mid(Cadena, 84, 14), 4) & "," ' campo Total Servicio Franquiciado
        mSQL = mSQL & .numero(Mid(Cadena, 98, 14), 4) & ","  ' campo Total Varios
        mSQL = mSQL & .numero(Mid(Cadena, 112, 14), 4) & ","  ' campo Total Descuentos
        mSQL = mSQL & .numero(Mid(Cadena, 126, 14), 4) & ","  ' campo Suma Total
        mSQL = mSQL & .numero(Mid(Cadena, 140, 5), 2) & ","  ' campo Proporción de cargos
        mSQL = mSQL & .numero(Mid(Cadena, 145, 14), 4) & ","  ' campo Cargos
        mSQL = mSQL & .numero(Mid(Cadena, 159, 5), 2) & ","  ' campo Proporción Desc. Cuotas
        mSQL = mSQL & .numero(Mid(Cadena, 164, 14), 4) & ","  ' campo Descuento cuotas
        mSQL = mSQL & .numero(Mid(Cadena, 178, 5), 2) & ","  ' campo Proporcion Dto. Serv. Med
        mSQL = mSQL & .numero(Mid(Cadena, 183, 14), 4) & ","  ' campo Descuento Servicio Medido
        mSQL = mSQL & .texto(Mid(Cadena, 197, 10)) & ","   ' campo Código centro de coste
        mSQL = mSQL & .texto(Mid(Cadena, 207, 40)) & ","   ' campo Descripción centro de coste
        mSQL = mSQL & .texto(Mid(Cadena, 247, 40)) & ","   ' campo Usuario
        mSQL = mSQL & .texto(Mid(Cadena, 287, 414)) & ")"   ' campo Libre
        .ejecutar mSQL
    End With
End Sub

Private Sub cargaRegistro310(Cadena As String) ' Cuotas
    Dim i As Integer
    Dim mSQL As String
    For i = 0 To 5
    With mdbFac2
        mSQL = "insert into telefono.cuotas values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(Cadena, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(Cadena, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(Cadena, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(Cadena, 24, 14)) & ","   ' campo N. de PreFactura
        mSQL = mSQL & .texto(Mid(Cadena, 38, 9)) & ","   ' campo N. de telefono
        mSQL = mSQL & .texto(Mid(Cadena, 47, 9)) & ","   ' campo N. de extensión
        If Mid(Cadena, 56 + (93 * i), 8) <> "00000000" Then
            '-- zona que se repite en el registro original ---------------------------------------
            mSQL = mSQL & .Fecha(fechatel(Mid(Cadena, 56 + (93 * i), 8))) & "," ' campo Fecha Inicio Cuotas
            mSQL = mSQL & .Fecha(fechatel(Mid(Cadena, 64 + (93 * i), 8))) & "," ' campo Fecha Fin Cuotas
            mSQL = mSQL & .texto(Mid(Cadena, 72 + (93 * i), 3)) & "," ' campo Código de cuota
            mSQL = mSQL & .texto(Mid(Cadena, 75 + (93 * i), 60)) & "," ' campo Descripción de la cuota
            mSQL = mSQL & .numero(Mid(Cadena, 135 + (93 * i), 14), 4) & "," ' campo Importe
            '-------------------------------------------------------------------------------------
            mSQL = mSQL & .texto(Mid(Cadena, 614, 97)) & ")"   ' campo Libre
            .ejecutar mSQL
        End If
    End With
    Next i
End Sub

Private Sub cargaRegistro320(Cadena As String) ' resumen de llamadas
    Dim i As Integer
    Dim mSQL As String
    
    
    
    For i = 0 To 5
    With mdbFac2
        mSQL = "insert into telefono.resumen_de_llamadas values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(Cadena, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(Cadena, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(Cadena, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(Cadena, 24, 14)) & ","   ' campo N. de PreFactura
        mSQL = mSQL & .texto(Mid(Cadena, 38, 9)) & ","   ' campo N. de telefono
        mSQL = mSQL & .texto(Mid(Cadena, 47, 9)) & ","   ' campo N. de extensión
        mSQL = mSQL & .Fecha(fechatel(Mid(Cadena, 56, 8))) & ","  ' campo Fecha Inicio Periodo
        mSQL = mSQL & .Fecha(fechatel(Mid(Cadena, 64, 8))) & ","  ' campo Fecha Fin Periodo
        If Mid(Cadena, 72 + (103 * i), 3) <> "   " Then
            '-- zona que se repite en el registro original ---------------------------------------
            
            mSQL = mSQL & .texto(Mid(Cadena, 72 + (103 * i), 3)) & "," ' campo Código de tipo de tráfico
            mSQL = mSQL & .texto(Mid(Cadena, 75 + (103 * i), 35)) & "," ' campo Tipo de tráfico
            mSQL = mSQL & .numero(Mid(Cadena, 110 + (103 * i), 7)) & "," ' campo Número de llamadas
            mSQL = mSQL & .texto(Mid(Cadena, 117 + (103 * i), 10)) & "," ' campo Unidad de medida
            mSQL = mSQL & .numero(Mid(Cadena, 127 + (103 * i), 10), 2) & "," ' campo Cantidad de medida
            mSQL = mSQL & .numero(Mid(Cadena, 137 + (103 * i), 10), 2) & "," ' campo Cantidad franquiciada
            mSQL = mSQL & .numero(Mid(Cadena, 147 + (103 * i), 14), 4) & "," ' campo Importe franquiciado
            
            
            mSQL = mSQL & .numero(Mid(Cadena, 161 + (103 * i), 14), 4) & "," ' campo Importes
            
            '-------------------------------------------------------------------------------------
            mSQL = mSQL & .texto(Mid(Cadena, 690, 11)) & ")"   ' campo Libre
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



Private Sub cargaRegistro330(Cadena As String) ' detalle de llamadas
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
            mSQL = mSQL & .numero(Mid(Cadena, 1, 3)) & "," ' campo Tipo de registro
            mSQL = mSQL & .numero(Mid(Cadena, 4, 6)) & ","   ' campo Secuencia de registro
            mSQL = mSQL & .texto(Mid(Cadena, 10, 14)) & ","   ' campo N.Factura
            mSQL = mSQL & .texto(Mid(Cadena, 24, 14)) & ","   ' campo N. de PreFactura
            mSQL = mSQL & .texto(Mid(Cadena, 38, 9)) & ","   ' campo N. de telefono
            mSQL = mSQL & .texto(Mid(Cadena, 47, 9)) & ","   ' campo N. de extensión
            CodigoTrafico = Trim(Mid(Cadena, 56, 3))
            mSQL = mSQL & .texto(CodigoTrafico) & ","   ' campo Código de tráfico
            mSQL = mSQL & .texto(Mid(Cadena, 59, 30)) & ","  ' campo Tipo de tráfico
            '-- zona que se repite en el registro original ---------------------------------------
            recalcular = False
            If vParamAplic.TieneTelefonia2 = 3 Then
                'mvarConceptosRefacturables = ",'17','18','w1'"
                
                CodigoTrafico = "'" & CodigoTrafico & "'"
                If InStr(1, mvarConceptosLlamadasRefacturables, CodigoTrafico) > 0 Then
                    'LAS ESPECIALES NO se refacturan
                    Aux = Mid(Cadena, 89 + (148 * i), 3)
                    If Trim(Aux) = "E" Then
                        'recalcular = True
                        
                    Else
                        recalcular = True
                    End If
                End If
                
            End If
            
            
            'Comun
            mSQL = mSQL & .texto(Mid(Cadena, 89 + (148 * i), 3)) & "," ' campo Código de abreviatura
            mSQL = mSQL & .texto(Mid(Cadena, 92 + (148 * i), 30)) & "," ' campo Descripción de tipo de llamada
            mSQL = mSQL & .texto(Mid(Cadena, 122 + (148 * i), 3)) & "," ' campo Código de desrtino
            mSQL = mSQL & .texto(Mid(Cadena, 125 + (148 * i), 14)) & "," ' campo Tipo de destino
            mSQL = mSQL & .texto(Mid(Cadena, 139 + (148 * i), 8)) & "," ' campo Operador
            mSQL = mSQL & .texto(Mid(Cadena, 147 + (148 * i), 25)) & "," ' campo Número llamado
            mSQL = mSQL & .texto(Mid(Cadena, 172 + (148 * i), 4)) & "," ' campo Fecha (MMDD)
            mSQL = mSQL & .texto(Mid(Cadena, 176 + (148 * i), 2)) & "," ' campo Código horario
            mSQL = mSQL & .texto(Mid(Cadena, 178 + (148 * i), 4)) & "," ' campo Hora (HHMM)
            mSQL = mSQL & .texto(Mid(Cadena, 182 + (148 * i), 10)) & "," ' campo Unidad de medida
            mSQL = mSQL & .numero(Mid(Cadena, 192 + (148 * i), 7), 2) & "," ' campo Cantidad medida originada
            mSQL = mSQL & .numero(Mid(Cadena, 199 + (148 * i), 7), 2) & "," ' campo Cantidad medida recibida
            mSQL = mSQL & .texto(Mid(Cadena, 206 + (148 * i), 10)) & "," ' campo Unidad calidad trasmisión
            mSQL = mSQL & .numero(Mid(Cadena, 216 + (148 * i), 7), 2) & "," ' campo Calidad suministrada
            
            
            
            Libre = Mid(Cadena, 681, 20) 'la cadena que va en la ultiposicion
            If Not recalcular Then

                mSQL = mSQL & .numero(Mid(Cadena, 223 + (148 * i), 14), 4) & "," ' campo Importe
                '-------------------------------------------------------------------------------------
                
            Else
                'RECALCULAMOS
                'Hacen una cosa muy rara. Me explico.  En duracion va un currency (2decimales)
                'Sin embargo el numero en los decimales son SEGUNDOS, no son convertido a decimal
                'Me sigo explicando
                'Ejemplo:   5.49    5 minutos y 49 segundos
                '           0.12    12 segundos
                
                'Si el importe facturado es CERO no hacemos nada mas
                Aux = .numero(Mid(Cadena, 223 + (148 * i), 14), 4) 'importe linea
                ImporteIncial = TransformaPuntosComas(Aux)
                
                If ImporteIncial <> 0 Then
                    
                    Libre = .numero(Mid(Cadena, 223 + (148 * i), 14), 4)
                    'Cantidad medida orignalmente
                    Aux = .numero(Mid(Cadena, 192 + (148 * i), 7), 2)
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

Private Sub cargaRegistro340(Cadena As String) ' varios
    Dim i As Integer
    Dim mSQL As String
    For i = 0 To 4
    With mdbFac2
        mSQL = "insert into telefono.varios values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(Cadena, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(Cadena, 4, 6)) & ","   ' campo Secuencia de registro
        mSQL = mSQL & .texto(Mid(Cadena, 10, 14)) & ","   ' campo N.Factura
        mSQL = mSQL & .texto(Mid(Cadena, 24, 14)) & ","   ' campo N. de PreFactura
        mSQL = mSQL & .texto(Mid(Cadena, 38, 9)) & ","   ' campo N. de telefono
        mSQL = mSQL & .texto(Mid(Cadena, 47, 9)) & ","   ' campo N. de extensión
        '-- zona que se repite en el registro original ---------------------------------------
        mSQL = mSQL & .texto(Mid(Cadena, 56 + (117 * i), 3)) & "," ' campo Código de Vario
        mSQL = mSQL & .texto(Mid(Cadena, 59 + (117 * i), 100)) & "," ' campo Descripción de Vario
        mSQL = mSQL & .numero(Mid(Cadena, 159 + (117 * i), 14), 4) & "," ' campo Importe
        '-------------------------------------------------------------------------------------
        mSQL = mSQL & .texto(Mid(Cadena, 641, 60)) & ")"   ' campo Libre
        .ejecutar mSQL
    End With
    Next i
End Sub

Private Sub cargaRegistro850(Cadena As String) ' Descuentos por factura
    Dim i As Integer
    Dim mSQL As String
    For i = 0 To 4
    With mdbFac2
        mSQL = "insert into telefono.descuentos_por_factura values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(Cadena, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(Cadena, 4, 6)) & ","   ' campo Secuencia de registr0
        mSQL = mSQL & .texto(Mid(Cadena, 10, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 24, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 38, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 47, 9)) & ","   ' campo Libre
        '----------- Zona que se repite
        mSQL = mSQL & .texto(Mid(Cadena, 56 + (120 * i), 3)) & "," ' campo Código de descuento
        mSQL = mSQL & .texto(Mid(Cadena, 59 + (120 * i), 40)) & "," ' campo Descripción de descuento
        mSQL = mSQL & .texto(Mid(Cadena, 99 + (120 * i), 3)) & "," ' campo Código de tipo de descuento
        mSQL = mSQL & .texto(Mid(Cadena, 102 + (120 * i), 1)) & "," ' campo Indicador tramo / base
        mSQL = mSQL & .texto(Mid(Cadena, 103 + (120 * i), 25)) & "," ' campo Tramo / base
        mSQL = mSQL & .numero(Mid(Cadena, 128 + (120 * i), 14), 4) & "," ' campo Importe
        mSQL = mSQL & .numero(Mid(Cadena, 142 + (120 * i), 5), 2) & "," ' campo Porcentaje descuento
        mSQL = mSQL & .texto(Mid(Cadena, 147 + (120 * i), 1)) & "," ' campo Indicador descuento global
        mSQL = mSQL & .numero(Mid(Cadena, 148 + (120 * i), 14), 4) & "," ' campo Descuento global
        mSQL = mSQL & .numero(Mid(Cadena, 162 + (120 * i), 14), 4) & "," ' campo Importe descuento
        '-----------------
        mSQL = mSQL & .texto(Mid(Cadena, 656, 45)) & ")"   ' campo Libre (??)
        
        .ejecutar mSQL
    End With
    Next i
End Sub

Private Sub cargaRegistro990(Cadena As String) ' totales del soporte
    Dim mSQL As String
    With mdbFac2
        mSQL = "insert into telefono.totales_del_soporte values("
        mSQL = mSQL & .texto(mFicheroCorto) & ","  ' campo Fichero
        mSQL = mSQL & .numero(Mid(Cadena, 1, 3)) & "," ' campo Tipo de registro
        mSQL = mSQL & .numero(Mid(Cadena, 4, 6)) & ","   ' campo Secuencia de registro
        mSQL = mSQL & .texto(Mid(Cadena, 10, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 24, 14)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 38, 9)) & ","   ' campo Libre
        mSQL = mSQL & .texto(Mid(Cadena, 47, 9)) & ","   ' campo Libre
        mSQL = mSQL & .numero(Mid(Cadena, 56, 24), 4) & ","  ' campo Total Cuotas
        mSQL = mSQL & .numero(Mid(Cadena, 80, 24), 4) & ","  ' campo Total varios
        mSQL = mSQL & .numero(Mid(Cadena, 104, 24), 4) & "," ' campo Total descuentos
        mSQL = mSQL & .numero(Mid(Cadena, 128, 24), 4) & ","  ' campo Total servicio medido
        mSQL = mSQL & .numero(Mid(Cadena, 152, 24), 4) & ","  ' campo Suma Total
        mSQL = mSQL & .numero(Mid(Cadena, 176, 24), 4) & ","  ' campo Total impuestos
        mSQL = mSQL & .numero(Mid(Cadena, 200, 22), 2) & ","  ' campo Total a pagar
        mSQL = mSQL & .numero(Mid(Cadena, 222, 8)) & ","   ' campo Total registros
        mSQL = mSQL & .numero(Mid(Cadena, 230, 8)) & ","   ' campo Total clientes
        mSQL = mSQL & .numero(Mid(Cadena, 238, 8)) & ","   ' campo Total SERVICIOS
        mSQL = mSQL & .numero(Mid(Cadena, 246, 8)) & ","   ' campo Total facturas
        mSQL = mSQL & .numero(Mid(Cadena, 254, 8)) & ","   ' campo Total prefacturas
        mSQL = mSQL & .numero(Mid(Cadena, 262, 8)) & ","   ' campo Total teléfonos fijos
        mSQL = mSQL & .numero(Mid(Cadena, 270, 8)) & ","   ' campo Total teléfonos móviles
        mSQL = mSQL & .texto(Mid(Cadena, 278, 5)) & ","   ' campo Código de soporte
        mSQL = mSQL & .texto(Mid(Cadena, 283, 6)) & ","   ' campo Código de agrupación
        mSQL = mSQL & .texto(Mid(Cadena, 289, 412)) & ")"   ' campo Libre (??)
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



Private Function EmitirFacturasMOVI_ORANGE(mfichero As String, mfecha As Date, ByRef Lbl As Label, LaCompanyia As Byte) As Boolean
    Dim mSQL As String
    Dim mSQL2 As String
    Dim mrs0 As ADODB.Recordset
    Dim mrs1 As ADODB.Recordset
    Dim mrs2 As ADODB.Recordset
    Dim mAno As Integer '-- valor del año de facturación
    Dim antTelefono As String '-- lleva el último teléfono leido
    Dim mLinea As Integer ' lleva el número de línea
    
    Dim mSumConsumos As Currency
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
    
    Dim Nombre As String
    Dim Apell1 As String
    Dim Apell2 As String
    
    Dim mFactura3 As Long  ' lleva el número de factura en proceso
    Dim mSerie As String
    
    Dim mValorIva As Currency
    
    
    Dim iNSERTS As String
    
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
    Set vI = New CTiposMov                          'LAS INTERNAS
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
            vC.IncrementarContador vC.TipoMovimiento   'Como hace el update con c+1
        ElseIf mrs0!Serie = Trim(vI.LetraSerie) Then
            

            mSQL2 = ""
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
    
            
    Set mrs0 = mdbFac2.cursor(mSQL)
    i = 0
    If Not mrs0.EOF Then
        NF = FreeFile()
        Open App.Path & "\emitefac.log" For Output As #NF
        Print #NF, "Facturas emitidas de fichero " & mfichero & " y fecha " & Format(mfecha, "dd/mm/yyyy")
        mrs0.MoveFirst
        While Not mrs0.EOF
            i = i + 1
            
            Lbl.Caption = "Cabecera facturas " & CStr(i)
            Lbl.Refresh
            DoEvents
        
            
        
        
        
            If mrs0!Numero_de_telefono <> antTelefono Then
                '-- Un nuevo número, una nueva factura
                '-- Primero comprobar si es interna
                
                'FALTA####
                If NifDelTelefono(mrs0!Numero_de_telefono) = vParam.CifEmpresa Then
                    vI.Contador = vI.Contador + 1
                    mFactura3 = vI.Contador
                    
                    mSerie = Trim(vI.LetraSerie)
                Else
                    vC.Contador = vC.Contador + 1
                    mFactura3 = vC.Contador
                    mSerie = Trim(vC.LetraSerie)
                End If
                '  mFactura2 = obtenNumeroFactura(vC.LetraSerie, mAno)  YA LO HA CONSEGUIDO ARRIBA conseguircontador
                mSQL = "insert into tel_cab_factura (Serie,Ano,NumFact,Fichero,Fecha,Telefono,TipoIVA,TotalPrevio,BaseImponible,Cuota,Total) values("
                
                mSQL = mSQL & mdbFac2.texto(mSerie) & ","  ' Serie
                mSQL = mSQL & mdbFac2.numero(mAno) & "," ' Año
                mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
                
                mSQL = mSQL & mdbFac2.texto(mfichero) & "," ' Fichero
                mSQL = mSQL & mdbFac2.Fecha(mfecha) & "," ' Fecha
                mSQL = mSQL & mdbFac2.texto(mrs0!Numero_de_telefono) & "," ' Telefono
                mSQL = mSQL & mdbFac2.numero(mValorIva) & "," ' IVA antes 18
                mSQL = mSQL & mdbFac2.numero(0) & "," ' Total previo (será recalculada)
                mSQL = mSQL & mdbFac2.numero(0) & "," ' Base Imponible (será recalculada)
                mSQL = mSQL & mdbFac2.numero(0) & "," ' Cuota (será recalculada)
                mSQL = mSQL & mdbFac2.numero(0) & ")" ' Total (idem)
                mdbFac2.ejecutar mSQL
                mLinea = 0 ' se inicializan las líneas para ese numero de telefono
                '-- buscamos si el teléfono está en la base de datos
                If Not EstaElTelefono(mrs0!Numero_de_telefono) Then
                    Print #NF, "Falta el teléfono " & mrs0!Numero_de_telefono & " en Ariges"
                    EmitirFacturasMOVI_ORANGE = True
                Else
             
                    
                   
                    'Cogera los datos de la sclien
                    mSQL = "select nifclien nifclien,nomclien,domclien,codpobla,pobclien,proclien,operador"
                    mSQL = mSQL & " from sclien,sclientfno  where sclien.codclien=sclientfno.codclien AND "
                    mSQL = mSQL & " IdTelefono = " & mdbFac2.texto(mrs0!Numero_de_telefono)
                    
                    
                    Set mrs1 = New ADODB.Recordset
                    Set mrs1 = mdbFac2.cursor(mSQL)
                    If Not mrs1.EOF Then
                    
                        SeparaNomclienNomApellidos mrs1!Nomclien, Nombre, Apell1, Apell2
                    
                    
                        mSQL = "update tel_cab_factura set "
                        mSQL = mSQL & " NIF = " & mdbFac2.texto(mrs1!nifClien) & ","
                        mSQL = mSQL & " Apellido1 = " & mdbFac2.texto(Apell1) & ","
                        mSQL = mSQL & " Apellido2 = " & mdbFac2.texto(Apell2) & ","
                        mSQL = mSQL & " Nombre = " & mdbFac2.texto(Nombre) & ","
                        mSQL = mSQL & " Direccion = " & mdbFac2.texto(mrs1!domclien) & ","
                        mSQL = mSQL & " CodPostal = " & mdbFac2.texto(mrs1!codpobla) & ","
                        mSQL = mSQL & " Poblacion = " & mdbFac2.texto(mrs1!pobclien) & ","
                        mSQL = mSQL & " Provincia = " & mdbFac2.texto(mrs1!proclien)
                        
                        

                        If LaCompanyia = 2 Then
                            Apell1 = "ORA"
                        Else
                            Apell1 = "MOV"
                        End If
                        mSQL = mSQL & ", companyia = " & DBSet(Apell1, "T")
                        
                        mSQL = mSQL & " where Telefono = " & mdbFac2.texto(mrs0!Numero_de_telefono)
                        mSQL = mSQL & " and Serie = " & mdbFac2.texto(mSerie)
                        mSQL = mSQL & " and Ano = " & mdbFac2.numero(mAno)
                        mSQL = mSQL & " and NumFact = " & mdbFac2.numero(mFactura3)
                        mdbFac2.ejecutar mSQL
                    End If
                End If
            End If
            antTelefono = mrs0!Numero_de_telefono
            mrs0.MoveNext
        Wend
        Close #NF
    End If
    
    'Ajustamos contadores
    vC.Contador = vC.Contador - 1
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
          
          
          
           'If mrs0!Telefono = "654396236" Then Stop
          
            mFactura3 = mrs0!NumFact
            mSerie = mrs0!Serie
            '-- Vamos a montar cuotas
            mSQL = "SELECT cuotas.Fichero, cuotas.Numero_de_telefono, cuotas.Codigo_de_cuota, cuotas.Descripcion_de_cuota, Sum(cuotas.Importe) AS SumaDeImporte, cuotas.Fecha_inicio_cuotas, cuotas.Fecha_final_cuotas" & _
                    " FROM telefono.cuotas" & _
                    " WHERE Fichero = " & mdbFac2.texto(mfichero) & _
                    " AND Numero_de_telefono = " & mdbFac2.texto(mrs0!Telefono) & _
                    " GROUP BY cuotas.Fichero, cuotas.Numero_de_telefono, cuotas.Codigo_de_cuota, cuotas.Descripcion_de_cuota, cuotas.Fecha_inicio_cuotas, cuotas.Fecha_final_cuotas"

'            mSQL = "SELECT * from Resumen_cuotas_x_telefono" & _
'                    " WHERE Fichero = " & mdbTel.Texto(mfichero) & _
'                    " AND Numero_de_telefono = " & mdbTel.Texto(mrs0!Telefono)
            Set mrs1 = mdbFac2.cursor(mSQL)
            mLinea = 0
            UltLineaCuota = 0
            If Not mrs1.EOF Then
                mrs1.MoveFirst
                
                iNSERTS = "" 'VELOCIDAD
                                
                'If mrs0!Telefono = "677013615" Then Stop
                
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
                    mSQL = mSQL & mdbFac2.numero(mrs1!SumaDeImporte) & ",0)" ' importe de la cuota correspondiente
                    
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
                        mSQL = mSQL & mdbFac2.numero(mDescuento) & ")" ' Importe
                        mdbFac2.ejecutar mSQL
                    End If
                    mrs1.MoveNext
                Wend
            End If
            '-- Leemos todos los consumos correspondientes al teléfono de la factura
            mSQL = "SELECT resumen_de_llamadas.Fichero, resumen_de_llamadas.Numero_de_telefono, resumen_de_llamadas.Codigo_de_tipo_de_trafico, resumen_de_llamadas.Tipo_de_trafico, Sum(resumen_de_llamadas.Numero_de_llamadas) AS SumaDeNumero_de_llamadas, Sum(resumen_de_llamadas.Cantidad_medida) AS SumaDeCantidad_medida, Sum(resumen_de_llamadas.Cantidad_franquiciada) AS SumaDeCantidad_franquiciada, Sum(resumen_de_llamadas.Importe_franquiciado) AS SumaDeImporte_franquiciado, Sum(resumen_de_llamadas.Importes) AS SumaDeImportes, resumen_de_llamadas.Fecha_inicio_periodo, resumen_de_llamadas.Fecha_final_periodo, resumen_de_llamadas.Unidad_de_medida" & _
                    " FROM telefono.resumen_de_llamadas" & _
                    " WHERE Fichero = " & mdbFac2.texto(mfichero) & _
                    " AND Numero_de_telefono = " & mdbFac2.texto(mrs0!Telefono) & _
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
                    mSQL = mSQL & mdbFac2.numero(mrs1!sumadeimportes) & ",0)" ' Importe del consumo
                    mFecha_inicio_periodo = mrs1!Fecha_inicio_periodo
                    mFecha_fin_periodo = mrs1!Fecha_final_periodo
                    
                    
                    iNSERTS = iNSERTS & ", " & mSQL
                    
                    'mdbFac2.Ejecutar mSQL
                    '-- Calculamos si hay algún descuento aplicable
                    mDescuento = buscaDescuentos(1, mrs1!sumadeimportes, mrs1!Codigo_de_tipo_de_trafico, mPorcentaje)
                    If mDescuento <> 0 Then
                        '-- Hay descuento aplicable
                        mSQL = "insert into tel_lin_factura_descuentos values("
                        mSQL = mSQL & mdbFac2.texto(mSerie) & "," ' Serie
                        mSQL = mSQL & mdbFac2.numero(mAno) & "," ' Año
                        mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
                        mSQL = mSQL & mdbFac2.numero(siguienteLineaDeDescuentos(mSerie, mAno, mFactura3)) & "," ' Número de línea
                        mSQL = mSQL & mdbFac2.texto(mrs1!Tipo_de_trafico) & "," ' Concepto
                        mSQL = mSQL & mdbFac2.numero(mPorcentaje) & "," ' Pordentaje
                        mSQL = mSQL & mdbFac2.numero(mrs1!sumadeimportes) & "," ' Base
                        mSQL = mSQL & mdbFac2.numero(mDescuento) & ")" ' Importe
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
            End If
            '-- Ahora comprobamos si en la tabla de varios tiene algun cargo aplicable
            mSQL = "select * from telefono.varios"
            mSQL = mSQL & " where Fichero = " & mdbFac2.texto(mfichero)
            mSQL = mSQL & " and Numero_de_telefono = " & mdbFac2.texto(mrs0!Telefono)
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
                        mSQL = mSQL & mdbFac2.texto(mrs0!Serie) & "," ' Serie
                        mSQL = mSQL & mdbFac2.numero(mAno) & "," ' Año
                        mSQL = mSQL & mdbFac2.numero(mrs0!NumFact) & "," ' Número de factura propiamente dicho
                        mSQL = mSQL & mdbFac2.numero(siguienteLineaDeEspeciales(vC.LetraSerie, mAno, mrs0!NumFact)) & "," ' Número de línea
                        mSQL = mSQL & mdbFac2.texto(mrs2!Descripcion_de_vario) & "," ' Concepto del cargo especial
                        mSQL = mSQL & mdbFac2.numero(mrs2!Importe) & ")"   ' Importe
                        mdbFac2.ejecutar mSQL
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
            mTotalPrevio = mSumCuotas + mSumConsumos + mSumEspecial
            mTotal = mTotalPrevio - mSumDescuentos


            mCuotaMinimia = 0
            mSQL = "select cuotaminima from sclientfno where idtelefono=" & mdbFac2.texto(mrs0!Telefono)
            mrs2.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not mrs2.EOF Then
                If Not IsNull(mrs2.Fields(0)) Then mCuotaMinimia = mrs2.Fields(0)
            End If
            mrs2.Close

            
            If mCuotaMinimia > 0 Then
                If mTotal < mCuotaMinimia Then
                    'HAy que añadir el minimo
                    mSumEspecial = mCuotaMinimia  'Reutilizo la variable
                    mCuotaMinimia = mCuotaMinimia - mTotal
                    mTotal = mSumEspecial
                    'DescCuota CodCuota tel_desc_cuotas ARI00000
                    mSQL = DevuelveDesdeBD(conAri, "DescCuota", "tel_desc_cuotas", "CodCuota", "ARI00000", "T")
                    
                    
                    mLinea = UltLineaCuota + 1 ' incrementamos la línea correspondiente
                    
                    mSQL = "," & mdbFac2.texto(mSQL) & "," & mdbFac2.numero(mCuotaMinimia) & ",0)"
                    mSQL = mdbFac2.Fecha(mfecha) & "," & mdbFac2.Fecha(mfecha) & "," & mdbFac2.texto("ARI00000") & mSQL
                    
                    mSQL = mdbFac2.texto(mSerie) & "," & mdbFac2.numero(mAno) & "," & mdbFac2.numero(mFactura3) & "," & mdbFac2.numero(mLinea) & "," & mSQL
                    mSQL = "insert into tel_lin_factura_cuotas values(" & mSQL
                    
                    mdbFac2.ejecutar mSQL
                    
                    
                    
                    
                    
                    
                    
                    
                End If
            End If


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
            End With
            mrs0.MoveNext
        Wend
        '-- Borradas todas las lineas de todas las facturas, borramos cabeceras
        mdbFac2.ejecutar "delete from tel_cab_factura where Fichero = " & mdbFac2.texto(mfichero)
    End If
    '-- Y ahora habra que revisar contadores
    ' FALTA###
    ' FALTA###
    ' FALTA###
    ' FALTA###
    

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

Private Function buscaDescuentos(Tipo As Integer, Importe As Currency, codigo As String, ByRef Porcentaje As Integer) As Currency
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
        mSQL = mSQL & " where CodCuota = " & mdbFac2.texto(codigo)
    Else
        '-- Consumos
        mSQL = mSQL & " from tel_desc_consumos"
        mSQL = mSQL & " where CodTipoTrafico = " & mdbFac2.texto(codigo)
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

Private Function buscaDescuentosGrupo(Tipo As Integer, Importe As Currency, codigo As String, ByRef Porcentaje As Integer) As Currency
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
        mSQL = mSQL & " where CodCuota = " & mdbFac2.texto(codigo)
    Else
        '-- Consumos
        mSQL = mSQL & " from tel_desc_consumos"
        mSQL = mSQL & " where CodTipoTrafico = " & mdbFac2.texto(codigo)
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
    Dim mSQL2 As String
    Dim Aux As String
    Dim mrs As ADODB.Recordset
    Dim mrs2 As ADODB.Recordset
    Dim mrs3 As ADODB.Recordset
    Dim ImporteProv As Currency
    Dim DescuentoProv As Currency
    Dim PorcentajeProv As Integer
    Dim CodTrafico As String
    mSQL = "select distinct(Grupo) from tel_desc_consumos where Grupo <> 0"
    Set mrs = mdbFac2.cursor(mSQL)
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
                    CodTrafico = mrs2!CodTipoTrafico ' vale cualquiera como representante
                    mSQL = "select sum(Importe) from tel_lin_factura_consumos"
                    mSQL = mSQL & " where serie = " & mdbFac2.texto(Serie)
                    mSQL = mSQL & " and ano = " & mdbFac2.numero(Ano)
                    mSQL = mSQL & " and numfact = " & mdbFac2.numero(NumFact2)
                    mSQL = mSQL & " and CodTipoTrafico = " & mdbFac2.texto(mrs2!CodTipoTrafico)
                    Set mrs3 = mdbFac2.cursor(mSQL)
                    If Not IsNull(mrs3.Fields(0)) Then
                        ImporteProv = ImporteProv + mrs3.Fields(0)
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
                    mSQL2 = "select * from tel_desc_nombres_grupo where grupo = " & mdbFac2.numero(mrs.Fields(0))
                    Set mrs3 = mdbFac2.cursor(mSQL2)
                    If Not mrs3.EOF Then
                        Aux = mrs3!Nombre
                    Else
                        Aux = "DESCUENTO AGRUPADO (" & CStr(mrs.Fields(0)) & ")"
                    End If
                    mSQL = mSQL & mdbFac2.texto(Aux) & "," ' Concepto
                    mSQL = mSQL & mdbFac2.numero(PorcentajeProv) & "," ' Pordentaje
                    mSQL = mSQL & mdbFac2.numero(ImporteProv) & "," ' Base
                    mSQL = mSQL & mdbFac2.numero(DescuentoProv) & ")" ' Importe
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
    Dim Id As Long
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
            Id = tmpSiguienteID()
            SQL = "insert into tmp_inf_descuentos(ID,fichero,tel_descripcion,tel_porcentaje,tel_base,tel_importe)"
            SQL = SQL & " values("
            SQL = SQL & mdbFac2.numero(Id) & ","
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
                                        " where ID = " & mdbFac2.numero(Id) & _
                                        " and cop_descripcion = " & mdbFac2.texto(cop_descripcion)
                                Set rs3 = mdbFac2.cursor(SQL)
                                If Not rs3.EOF Then cop_descripcion = ""
                                SQL = "select * from tmp_inf_descuentos2" & _
                                        " where ID = " & mdbFac2.numero(Id) & _
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
                                        ID2 = tmpsiguienteID2(Id)
                                        SQL = "insert into tmp_inf_descuentos2(ID,ID2,fichero,cop_descripcion,cop_porcentaje,cop_base,cop_importe)"
                                        SQL = SQL & " values("
                                        SQL = SQL & mdbFac2.numero(Id) & ","
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

Private Function tmpsiguienteID2(Id As Long) As Long
    Dim RS As ADODB.Recordset
    Dim SQL As String
    SQL = "select max(ID2) from tmp_inf_descuentos2" & _
            " where ID = " & mdbFac2.numero(Id)
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
Private Sub SeparaNomclienNomApellidos(ByVal Nomclien As String, Nom As String, Ap1 As String, Ap2 As String)
Dim i As Integer
Dim TieneComa As Boolean

    i = InStr(1, Nomclien, ",")
    
    TieneComa = i > 0
    
    If TieneComa Then
        'El nombre va detras de la coma
        Nom = Mid(Nomclien, i + 1)
        Nomclien = Mid(Nomclien, 1, i - 1)
        
        
        i = InStrRev(Nomclien, " ")
        If i = 0 Then
            Ap1 = Nomclien
            Ap2 = ""
        Else
            Ap1 = Mid(Nomclien, 1, i - 1)
            Ap2 = Mid(Nomclien, i + 1)
        End If
        
        
    Else
        
        Ap1 = ""
        Ap2 = ""
        
        i = InStr((Len(Nomclien) / 3), Nomclien, " ")
        
        If i = 0 Then
            i = InStr(1, Nomclien, " ")
            If i = 0 Then
                Nomclien = "NO encontrado"
                Exit Sub
            End If
        End If
        
        Nom = Mid(Nomclien, 1, i - 1)
        Nomclien = Trim(Mid(Nomclien, i + 1))
        
        i = InStr(5, Nomclien, " ")
        If i = 0 Then
            Ap1 = Nomclien
            Ap2 = ""
        Else
            Ap1 = Mid(Nomclien, 1, i - 1)
            Ap2 = Mid(Nomclien, i + 1)
        End If
        
        
    End If
    
    
End Sub







Public Function RecalcularImporteLlamadasCoperativa(ByRef L As Label, Operadora As Byte)
    Dim RS As ADODB.Recordset
    Dim Cad As String
    Dim Aux As String
    Dim Impor As Currency
    Dim imp2 As Currency
    Dim Aux2 As String
    Dim i As Integer
    
    'Entran con espacios en blanco
    L.Caption = "Ajustando numeros llamados"
    L.Refresh
    Cad = "update telefono.detalle_de_llamadas set Numero_llamado =trim(Numero_llamado) where fichero='" & mFicheroCorto & "'"
    conn.Execute Cad
    Espera 0.25
    
    'LLAMADAS DE COOPERATIVA
    Set RS = New ADODB.Recordset
    
    If Operadora = 1 Then
        'SOLO PARA MOVISTAR. Las llamadas entre telefonos en ORANGE, ya vienen a 0
    
     L.Caption = "Leyendo registros llamada"
     L.Refresh
     If mvarConceptosQueSonLlamadas <> "" Then
         Cad = "select Numero_llamado from telefono.detalle_de_llamadas where fichero='" & mFicheroCorto & "' and Numero_llamado<>'' and"
         Cad = Cad & " Numero_llamado in (select idtelefono from sclientfno where operador = " & Operadora & ") group by 1"
         Set RS = New ADODB.Recordset
         RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         While Not RS.EOF
             L.Caption = RS.Fields(0)
             L.Refresh
             
             
             'Primero el importl
             Cad = "update telefono.detalle_de_llamadas set libre=importe"
             Cad = Cad & " WHERE Numero_llamado ='" & RS.Fields(0) & "' AND fichero='" & mFicheroCorto & "'"
             Cad = Cad & " and importe<>0 AND libre=''"
             Cad = Cad & " and Codigo_de_trafico in (" & mvarConceptosQueSonLlamadas & ")"
             conn.Execute Cad
           
           
             'Me estoy guardando el importe anterior en LIBRE
             Cad = "update telefono.detalle_de_llamadas set importe=0"
             Cad = Cad & " WHERE Numero_llamado ='" & RS.Fields(0) & "' AND fichero='" & mFicheroCorto & "'"
             Cad = Cad & " and importe<>0 "
             Cad = Cad & " and Codigo_de_trafico in (" & mvarConceptosQueSonLlamadas & ")"
             conn.Execute Cad
             
             
             
         
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
        Cad = "concat(numero_de_factura,'|',min(Fecha_inicio_periodo),'|',max(Fecha_final_periodo),'|')"
        Cad = DevuelveDesdeBD(conAri, Cad, "telefono.resumen_de_llamadas", Aux, "1 GROUP BY  numero_de_factura", "N")
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
            
            Aux = Aux & "  SELECT '" & mFicheroCorto & "',320,1,'" & RecuperaValor(Cad, 1) & "',''"
            Aux = Aux & ",Numero_de_telefono,'','" & RecuperaValor(Cad, 2) & "','" & RecuperaValor(Cad, 3) & "',"
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
        
            Aux = "  SELECT '" & mFicheroCorto & "',320,1,'" & RecuperaValor(Cad, 1) & "',''"
            Aux = Aux & ",Numero_de_telefono,'','" & RecuperaValor(Cad, 2) & "','" & RecuperaValor(Cad, 3) & "',"
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
            
            Cad = "" 'cuidado. Esta variable estaba para el numero de factura telefonia, fechainic y fin
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
                
                Cad = Cad & Aux2 & ")"
                
                
                If Len(Cad) > 25000 Then
                    Cad = Mid(Cad, 2)
                    Cad = Aux & " VALUES " & Cad
                    conn.Execute Cad
                    Cad = ""
                End If
                RS.MoveNext
            Wend
            RS.Close
            
            If Cad <> "" Then
                Cad = Mid(Cad, 2)
                Cad = Aux & " VALUES " & Cad
                conn.Execute Cad
            
            End If
        End If
    End If
    
    
    
    'En bolbaite. Ya vermeos como lo parametrizo
    'pOr si acaso
    If vParamAplic.TieneTelefonia2 = 3 Then
        If Operadora = 1 Then
            L.Caption = "Concepto W1"
            L.Refresh

            Cad = "UPDATE telefono.resumen_de_llamadas set codigo_de_tipo_de_trafico='20' WHERE Fichero =" & DBSet(mFicheroCorto, "T")
            
            Cad = Cad & " AND codigo_de_tipo_de_trafico='W1'"
            conn.Execute Cad
            
        End If
    End If
    
    
    
    
    'Cargos/Cuotas de varios
    'tel_cargo_varios
    
    
    
    
    
    
    
    Set RS = Nothing
    
    
End Function



'TipoFicheroOperador
'   1. MOV    2.  ORAN  ....
Public Function AjusteCuotasNuevas2(TipoFicheroOperador As Byte, mfichero As String, ByRef L As Label)
Dim vSQL As String
Dim Cad As String
Dim Origen As String
Dim Rc As ADODB.Recordset
Dim i As Integer
Dim Importe2 As Currency
Dim Importe3 As Currency

    L.Caption = "Ajuste quotas"
    L.Refresh
    
   

    'Valores por defecto. Para poder ordenar debe ser numerico el campo del where
    If TipoFicheroOperador = 2 Then mfichero = mFicheroCorto
    
    vSQL = "Fichero =" & DBSet(mfichero, "T") & " AND 1"
    
    Cad = "concat(numero_de_factura,'|',min(Fecha_inicio_cuotas),'|',min(Fecha_final_cuotas),'|')"
    Origen = DevuelveDesdeBD(conAri, Cad, "telefono.cuotas", vSQL, "1 GROUP BY  numero_de_factura", "N")
    If Origen = "" Then
        'Probamos en detalle lamadas
         
        vSQL = "Fichero =" & DBSet(mfichero, "T") & " AND 1"
    
        Cad = "concat(numero_de_factura,'|',min(Fecha_inicio_periodo),'|',min(Fecha_final_periodo),'|')"
        Origen = DevuelveDesdeBD(conAri, Cad, "telefono.resumen_de_llamadas", vSQL, "1 GROUP BY  numero_de_factura", "N")
    End If
    
    L.Caption = "Ajuste cuotas personalizable (I)"
    L.Refresh
    
    Set Rc = New ADODB.Recordset
    
    'Vemos las acciones tipo1.   ELIMINAR
    Cad = DevuelveACcionesTipo(1, 0, Rc, TipoFicheroOperador)
    If Cad <> "" Then
        Cad = "DELETE FROM telefono.cuotas WHERE Fichero =" & DBSet(mfichero, "T") & " AND Codigo_de_cuota IN (" & Cad & ")"
        conn.Execute Cad
    End If
    
    L.Caption = "Ajuste cuotas personalizable (II)"
    L.Refresh
    'Acciones tipo 2, cuota mes entera
    '  tipo: 0. Se cobra el mes entero.

    
    Cad = "Select * from tel_descuentoscuotas WHERE accion=2 AND Mensual =0 " 'todo el mes
    Cad = Cad & " AND operadora =" & TipoFicheroOperador
'    vSQL = "1"
'    If EsFicheroOrange_ Then vSQL = "2"
'    Cad = Cad & vSQL
    
    Rc.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not Rc.EOF
        Cad = TransformaComasPuntos(CStr(Rc!Valor))
        Cad = "UPDATE telefono.cuotas SET importe = " & Cad
        Cad = Cad & " WHERE Fichero =" & DBSet(mfichero, "T") & " AND Codigo_de_cuota = " & DBSet(Rc!codigo_de_cuota, "T")
        conn.Execute Cad
        Rc.MoveNext
    Wend
    Rc.Close
    
    
     'Acciones tipo 2, cuota mes entera
    '  tipo: 1.  Se prorratea los dias,
    L.Caption = "Ajuste cuotas personalizable (III)"
    L.Refresh
    Cad = DevuelveACcionesTipo(2, 1, Rc, TipoFicheroOperador)
    If Cad <> "" Then
        MsgBox "FALTA acciones 2. Cuota mes entera"
        Stop
    End If
    
    
    'Acciones tipo 2, cuota mes entera
    '  tipo: 1.  Se prorratea los dias,
    L.Caption = "Ajuste cuotas personalizable (III)"
    L.Refresh
    Cad = DevuelveACcionesTipo(3, 1, Rc, TipoFicheroOperador)
    If Cad <> "" Then
        L.Caption = "Prorrateo de cuotas"
        L.Refresh
        
        'CUOTAS PRORRATEABLES pro fecha de baja u alta
        
        vSQL = "select Codigo_de_cuota,Numero_de_telefono,datediff(Fecha_final_cuotas,Fecha_inicio_cuotas) dias,Importe from telefono.cuotas"
        vSQL = vSQL & " WHERE Fichero =" & DBSet(mfichero, "T") & " AND Codigo_de_cuota IN (" & Cad & ") ORDER BY 1"
        Rc.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
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
            
            
            
            Cad = "UPDATE telefono.cuotas SET libre=" & TransformaComasPuntos(CStr(Rc!Importe))
            Cad = Cad & ",importe = " & TransformaComasPuntos(CStr(Importe3))
            Cad = Cad & " WHERE Fichero =" & DBSet(mfichero, "T") & " AND Codigo_de_cuota = " & DBSet(Rc!codigo_de_cuota, "T")
            Cad = Cad & " AND Numero_de_telefono =" & DBSet(Rc!Numero_de_telefono, "T")
            conn.Execute Cad
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
    
    Cad = "Fichero =" & DBSet(mfichero, "T") & " and not codigo_de_cuota in (select codigo_de_cuota from tel_descuentoscuotas where " & vSQL & ") and 1"
    Cad = DevuelveDesdeBD(conAri, "count(*)", "telefono.cuotas", Cad, "1", "N")
    If Cad <> "" Then
        If Val(Cad) > 0 Then Cad = ""
    End If
    If Cad = "" Then
        Cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2) "
        Cad = Cad & " select " & vUsu.codigo & ",1,0,Codigo_de_cuota,Descripcion_de_cuota"
        Cad = Cad & "  from telefono.cuotas where Fichero =" & DBSet(mfichero, "T") & " group by Codigo_de_cuota"
        conn.Execute Cad
        
        'vSQL = "1"
        'If EsFicheroOrange_ Then vSQL = "2"
        vSQL = TipoFicheroOperador
        Cad = "insert ignore into tel_descuentoscuotas(operadora,codigo_de_cuota,Descripcion_de_cuota,accion,Mensual,Valor)"
        Cad = Cad & " select " & vSQL & ",Codigo_de_cuota,Descripcion_de_cuota,0,0,sum(importe)/count(*) from telefono.cuotas "
        Cad = Cad & " where Fichero =" & DBSet(mfichero, "T") & " group by Codigo_de_cuota"
        conn.Execute Cad
        
    End If
    

    'VARIOS
    L.Caption = "Cargos VARIOS"
    L.Refresh
    Cad = "Fichero =" & DBSet(mfichero, "T") & " and Codigo_de_vario<>'' and not Codigo_de_vario "
    'vSQL = "1"
    'If EsFicheroOrange_ Then vSQL = "2"
    vSQL = "operadora = " & TipoFicheroOperador '& vSQL
    Cad = Cad & " in (select CodigoVario from tel_cargo_varios where " & vSQL & ") and 1"
    Cad = DevuelveDesdeBD(conAri, "count(*)", "telefono.varios", Cad, "1", "N")
    If Cad <> "" Then
        If Val(Cad) > 0 Then Cad = ""
    End If
    If Cad = "" Then
        'Codigo_de_vario Descripcion_de_vario

        Cad = "INSERT INTO tmpinformes(codusu,codigo1,campo1,nombre1,nombre2) "
        Cad = Cad & " select " & vUsu.codigo & ",2,0,Codigo_de_vario,Descripcion_de_vario"
        Cad = Cad & "  from telefono.varios where Fichero =" & DBSet(mfichero, "T")
        Cad = Cad & " AND Codigo_de_vario<>"""" AND Importe<>0 group by Codigo_de_vario"
        conn.Execute Cad
        

        Cad = "insert ignore tel_cargo_varios(CodigoVario,Descripcion,CargarEnFactura,Operadora)"
        Cad = Cad & " select Codigo_de_vario,Descripcion_de_vario,0,1 from telefono.varios "
        Cad = Cad & " where Codigo_de_vario<>'' AND importe<>0"
        Cad = Cad & " AND Fichero =" & DBSet(mfichero, "T") & " group by Codigo_de_vario"
        conn.Execute Cad
        
    End If


    DoEvents
    L.Caption = "Telefonos sin consumo / inactivos / err. operadora"
    L.Refresh


    'Despues de ajustar las cuotas, mostraremos los telefonos que no tienen CUOTAS
    'Por si quiere quitar alguno. Para que no le lleguen la factura
    'Primero. Metemos todos los telelfonos del fichero
    Cad = "DELETE FROM tmpnseries WHERE codusu = " & vUsu.codigo
    conn.Execute Cad
    Espera 0.15
    
    
    Cad = "select " & vUsu.codigo & ",Numero_de_telefono,0 from telefono.telefono" & _
            " where Fichero = " & mdbFac2.texto(mfichero) & " order by Numero_de_telefono"
    Cad = "INSERT INTO tmpnseries(codusu,codartic,numlinea) " & Cad
    conn.Execute Cad
    Espera 0.25
    
    'INACTIVOS. Pondremos numlinea =2
    Cad = "select * from tmpnseries where codusu =" & vUsu.codigo & " and codartic in (select IdTelefono from sclientfno where inactivo=1)"
    Rc.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rc.EOF
        Cad = "UPDATE tmpnseries set numlinea= 2 where codusu = " & vUsu.codigo & " AND codartic = " & DBSet(Rc!codArtic, "T")
        Rc.MoveNext
        ejecutar Cad, False
        Espera 0.1
    Wend
    Rc.Close
    
    'Febrero 2014
    'Los que en la fiche del telefono, el operador que llevan es OTRO que el del fichero
    Cad = "select * from tmpnseries where codusu =" & vUsu.codigo & " and codartic in (select IdTelefono from sclientfno where inactivo=0 and operador <>" & TipoFicheroOperador & " )"
    Rc.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rc.EOF
        Cad = "UPDATE tmpnseries set numlinea= 5 where codusu = " & vUsu.codigo & " AND codartic = " & DBSet(Rc!codArtic, "T")
        Rc.MoveNext
        ejecutar Cad, False
        Espera 0.1
    Wend
    Rc.Close
    
    
    'Borramos los que tienen detalle _llamadas , cutoas y varios  FEB 2014 -->
    Cad = "DELETE FROM tmpnseries WHERE codusu =" & vUsu.codigo & " AND numlinea<2 AND codartic IN ("
    vSQL = "Select distinct(Numero_de_telefono) FROM telefono.cuotas WHERE Fichero = " & mdbFac2.texto(mfichero)
    vSQL = Cad & vSQL & ")"
    conn.Execute vSQL
    vSQL = "Select distinct(Numero_de_telefono) FROM telefono.resumen_de_llamadas WHERE Fichero = " & mdbFac2.texto(mfichero)
    vSQL = Cad & vSQL & ")"
    conn.Execute vSQL
    vSQL = "Select distinct(Numero_de_telefono) FROM telefono.varios WHERE Fichero = " & mdbFac2.texto(mfichero)
    vSQL = Cad & vSQL & ")"
    conn.Execute vSQL
    

    
    
    'SI hay telefonos, los muestro por si quiere eliminarlos para no facturarlos
    
    vSQL = DevuelveDesdeBD(conAri, "count(*)", "tmpnseries", "codusu", CStr(vUsu.codigo))
    If Val(vSQL) > 0 Then
            'Lanzamos el visor donde sal
            frmVarios3.Opcion = 3
            frmVarios3.Show vbModal
    
            Cad = "Select * from tmpnseries WHERE codusu = " & vUsu.codigo
            Rc.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            While Not Rc.EOF
                
                'Febrero 2013
                'Borramos tb los datos que vienen del fichero
                L.Caption = "Quitar datos: " & Rc!codArtic
                L.Refresh
                
                Cad = " WHERE Fichero =" & DBSet(mfichero, "T") & " AND Numero_de_telefono = " & DBSet(Rc!codArtic, "T")
                
                conn.Execute "DELETE FROM telefono.cuotas" & Cad
                
                conn.Execute "DELETE FROM telefono.resumen_de_llamadas" & Cad
                
                conn.Execute "DELETE from telefono.varios" & Cad
                
                conn.Execute "DELETE FROM telefono.telefono" & Cad
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
    
        
        vSQL = "INSERT INTO telefono.cuotas (Fichero,Tipo_de_registro,Secuencia_de_registro,Numero_de_factura,"
        vSQL = vSQL & "Numero_de_prefactura,Numero_de_telefono,Numero_de_extension , Fecha_inicio_cuotas, "
        vSQL = vSQL & " Fecha_final_cuotas, Codigo_de_cuota, Descripcion_de_cuota,"
        vSQL = vSQL & "Importe, Libre)  SELECT '" & mfichero & "',9999 + numlinea,1,'" & RecuperaValor(Origen, 1)
        vSQL = vSQL & "','',idtelefono,'','" & RecuperaValor(Origen, 2) & "','" & RecuperaValor(Origen, 3) & "',"
        vSQL = vSQL & "concat('D',numlinea),descripcion,precio,'ARIGES' FROM sclientfnoCuotas  WHERE "
        'Ene 2014  Fichero
        vSQL = vSQL & " IdTelefono IN (select distinct Numero_de_telefono FROM telefono.telefono where Fichero = " & mdbFac2.texto(mfichero) & " order by Numero_de_telefono)"
        
        
        mdbFac2.ejecutar vSQL

    End If








End Function

Public Sub ComprobarConceptosFacturacion()
Dim C As String

 
    C = "Fichero =" & DBSet(mfichero2, "T") & " and not Codigo_de_tipo_de_trafico in (select Codigo_de_tipo_de_trafico from tel_conceptosllamadas where operadora=1) and 1"
    C = DevuelveDesdeBD(conAri, "count(*)", "telefono.resumen_de_llamadas", C, "1", "N")
    If Val(C) > 0 Then
        
    
    
        C = "INSERT INTO tmpinformes(codusu,codigo1,campo1,campo2,nombre1,nombre2) "
        C = C & " select " & vUsu.codigo & ",1,0,Codigo_de_cuota,Descripcion_de_cuota"
        C = C & " where Fichero =" & DBSet(mfichero2, "T") & " group by Codigo_de_cuota"
        mdbFac2.ejecutar C
    
    
        C = "INSERT INTO tel_conceptosllamadas (operadora,Codigo_de_tipo_de_trafico,Tipo_de_trafico,refacturar)"
        C = C & " select 1 operadora,Codigo_de_tipo_de_trafico,Tipo_de_trafico,0 refacturar from telefono.resumen_de_llamadas"
        C = C & " WHERE Fichero = '" & mfichero2 & "'  group by Codigo_de_tipo_de_trafico order by Codigo_de_tipo_de_trafico"
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
    
    On Error GoTo eLeerParametrosFacturacionTelefonica
    LeerParametrosFacturacionTelefonica = False
    mvarConceptosLlamadasRefacturables = ""
    mvarConceptosQueSonLlamadas = ""
    If vParamAplic.TieneTelefonia2 = 3 Then
        'BLOBAITE
        
        'Conceptos refacturables
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
            
            
        mvarPrecioEstablecimiento = 12  'centimos de euro
        mvarPrecioMinuto = 9    'centimos de euro
        
        mvarFichero = "select Establecimiento ,PrecioUnMinuto,EstablecimientoORAN ,PrecioUnMinutoORAN from spara2"
        miRsAux.Open mvarFichero, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
        
            
            If Companyia = 2 Then
                'EstablecimientoORAN PrecioUnMinutoORAN
                mvarPrecioEstablecimiento = miRsAux!EstablecimientoORAN  'centimos de euro
                mvarPrecioMinuto = miRsAux!PrecioUnMinutoORAN    'centimos de euro
            Else
                mvarPrecioEstablecimiento = miRsAux!Establecimiento  'centimos de euro
                mvarPrecioMinuto = miRsAux!PrecioUnMinuto    'centimos de euro
            End If
        End If
        miRsAux.Close
        
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
        Line Input #NF, mfichero2
        mvarFichero = mvarFichero & mfichero2
    Wend
    Close (NF)
    
    'Para que devuelve el nombre del fichero
    ComunNombreFicheroOrange
    
    
    DevuelveNombreFicheroOrange = mfichero2
    
    mvarFichero = ""
    mfichero2 = ""
    
    Exit Function
eDevuelveNombreFicheroOrange:
    MuestraError Err.Number
End Function



Private Sub ComunNombreFicheroOrange()
Dim i As Integer


    ''número de factura:;A10020017274-0813
    mfichero2 = ""
    
    NF = InStr(1, mvarFichero, "mero de factura:")
    If NF = 0 Then Exit Sub
    
    NF = NF + 16 'numero de facutra
    
    'buscamos el guion delimitador
    i = InStr(NF, mvarFichero, "-")
    If i = 0 Then Exit Sub
    
    If i - NF > 13 Then Exit Sub 'El telefono es varchar12
    
    mfichero2 = Mid(mvarFichero, NF + 1, i - NF - 1)

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
        Line Input #NF, mfichero2
        mvarFichero = mvarFichero & mfichero2
    Wend
    Close (NF)
    
    ComunNombreFicheroOrange  'devolvera el nombre del fichero
    mFicheroCorto = mfichero2
    mfichero2 = ""
    
    'Cargamos datos del
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
'                    If InStr(1, Tramo, "consumo en servicios especia") > 0 Then Stop
'                    If InStr(1, Tramo, "mínimo") > 0 Then Stop
'                    If InStr(1, Tramo, "anquilidad") > 0 Then Stop
'                    If InStr(1, Tramo, "606399489") > 0 Then Stop
                    
                    
                    If Tipo = 1 Then
                        'NADA
                    Else
                        If Tipo > 1 And Tipo < 10 Then  '2 al 9 detalles de llamada
                            InsertarTramoDetalleLlamdas IdLinDetalle, Tipo, Tramo
                        Else
                            OrangeInsertBdTelefono Tipo, Tramo
                        End If
                    End If
                End If
            
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




Private Sub InsertarTramoDetalleLlamdas(ByRef IDLinea As Long, Tipo As Byte, Tramo As String)
Dim H As Long
Dim linea As String
Dim Aux As String
Dim TotalLlamada As Currency

Dim Ud As String
Dim Libre As String
Dim Valor As Currency
Dim Valor2 As Currency
    
    '-->Fecha mmdd   hora hhmm
    'Fichero,id,Numero_de_telefono,Codigo_de_trafico,Tipo_de_trafico,Numero_llamado,Fecha,Hora_inicio,Unidad_de_medida,  Cantidad_medida_originada Importe
    'quito las dos primeras lineas que son expliativas
    
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
            Tramo = ""
            linea = ""
        Else
            linea = Mid(Tramo, 1, H - 1)
            Tramo = Mid(Tramo, H + 1)
        End If
        If linea <> "" Then
            If InStr(1, linea, "|") > 0 Then
                MsgBox "Separador incorrecto", vbExclamation
                Stop
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
            
            'If Tipo <> 2 And Tipo <> 4 And Tipo <> 7 Then Stop
            
            
            Aux = RecuperaValor(linea, 6) 'unidada
            Libre = ""
            TotalLlamada = DevuelveUnidad(Tipo, RecuperaValor(linea, 8), Aux)
            If Tipo <> 4 Then
                If Tipo = 8 Then
                    Aux = "SMS"
                ElseIf Tipo = 7 Then
                    If Aux = "" Then Aux = "Segundos"
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
            
            
            
            CadenaInsert = CadenaInsert & TransformaComasPuntos(Aux) & ",'" & Libre & "')"
            
            
            
            
        
            
            
            
            
            'INSERT COMUN
            If Len(CadenaInsert) > 30000 Then
                    CadenaInsert = Mid(CadenaInsert, 2)
                    Aux = "insert into telefono.detalle_de_llamadas(Fichero,id,Numero_de_telefono,Codigo_de_trafico,Tipo_de_trafico,Numero_llamado,fecha,Hora_inicio, Unidad_de_medida, Cantidad_medida_originada,Importe,libre) VALUES "
                    Aux = Aux & CadenaInsert
                    conn.Execute Aux
                    CadenaInsert = ""
            End If
        End If
    Wend
    
    If Len(CadenaInsert) > 0 Then
        
        CadenaInsert = Mid(CadenaInsert, 2)
        Aux = "insert into telefono.detalle_de_llamadas(Fichero,id,Numero_de_telefono,Codigo_de_trafico,Tipo_de_trafico,Numero_llamado,fecha,Hora_inicio,Unidad_de_medida,  Cantidad_medida_originada,Importe,libre) VALUES "
        Aux = Aux & CadenaInsert
        conn.Execute Aux
        CadenaInsert = ""
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
        'If Cambialo = 0 Then Stop
    End If
    If Tipo = 4 Or Tipo = 9 Then
        DevuelveUnidad = 0
        Aux = Trim(Aux)
        If LCase(Right(Aux, 2)) = "gb" Then
            'MsgBox "GIGAS"
            Ud = "Gigabytes"
            Aux = Trim(Mid(Aux, 1, Len(Aux) - 2))
            DevuelveUnidad = CCur(Aux)
        End If
        If LCase(Right(Aux, 2)) = "mb" Then
            
            Ud = "Megabytes"
            Aux = Trim(Mid(Aux, 1, Len(Aux) - 2))
            DevuelveUnidad = CCur(Aux)
        End If
        If LCase(Right(Aux, 2)) = "kb" Then
            
            Ud = "Kilobytes"
            Aux = Trim(Mid(Aux, 1, Len(Aux) - 2))
            
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

    Dim Leido As String ' registro leido del fichero
    Dim i As Long ' un contadoir de registros
    Dim GuardarNombreFichero As String
    Dim tabla As String

    
    
    On Error GoTo ecargarBaseDatosVODAFONE
    
    
    cargarBaseDatosVODAFONE = False
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
        Line Input #NF, Leido
        i = i + 1
        
        
        
        LB.Caption = "Registro " & CStr(i)
        LB.Refresh
        If (i Mod 40) = 0 Then DoEvents
        
        '-- hacemos lo que tengamos que hacer con el registro
        If Leido <> "" Then
            
            Valores = Split(Leido, Chr(9))
                       
                       
                       
                
          
        
        
                       
                       
                       
            Select Case Left(Leido, 8)
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
                Case Else ' desconocido
                    '-- Registro no contemplado
                    '
                    If Mid(Leido, 6, 3) = "CAB" Then
                        'Son las cabeceras
                        '
                         
                         If CadenaInsert <> "" Then
                            conn.Execute CadenaInsert
                            mFicheroCorto = ""
                         End If
                         
                         
                         CadenaInsert = ""
                         
                    Else
                        MsgBox "El registro TIPO: " & Left(Leido, 3) & " no está contemplado en la aplicación.", vbCritical
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
                    'If InStr(1, CadenaInsert, "600400257") > 0 Then Stop
                
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
    LB.Caption = "Comprobaciones"
    LB.Refresh
    Espera 1
    IntentaEliminarFicheroLog  'por si acaso existe
    
    NF = FreeFile
    GuardarNombreFichero = App.Path & "\emitefac.log"
    Open GuardarNombreFichero For Output As #NF
    
    i = 0
    If Not ComprobarCuotasVodafone(LB) Then i = 1
    
    If Not ComprobarTraficoVodafone(LB) Then i = 1
    
    If Not ComprobarNumerosDeTelefonoVODAFONE(LB) Then i = 1
    
    Close #NF
    LB.Caption = "Finalizando importacion"
    Set miRsAux = New ADODB.Recordset
    
    If i = 0 Then
        IntentaEliminarFicheroLog
        cargarBaseDatosVODAFONE = True
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


Private Function ComprobarTraficoVodafone(ByRef L As Label) As Boolean
Dim C As String
Dim Conceptos As String
Dim i As Integer

    ComprobarTraficoVodafone = True 'en ppio estan bien
    
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
    
        i = InStr(1, Conceptos, "|")
        If i = 0 Then
            If Conceptos <> "" Then Err.Raise 513, , "Comprobando cuotas trafico vodafone(1)"
        Else
            
            C = Mid(Conceptos, 1, i - 1)
            L.Caption = C
            L.Refresh
            
            Conceptos = Mid(Conceptos, i + 1)
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
    Wend

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


    On Error GoTo eEmitirFacturasVODAFONE
    
    NF = -1
    EmitirFacturasVODAFONE = False
    
    '-- Por si nos pasan ruta completa modificamos el nombre de fichero
    If Len(mfichero) > 12 Then mfichero = Right(mfichero, 12)
    mfichero2 = mfichero
    
    
    Set miRsAux = New ADODB.Recordset
    Set RTef = New ADODB.Recordset
    Set RsCuotasRefacturar = New ADODB.Recordset
    
 
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
        
    mSQL = "select distinct(substring(telef,1,9)) elTelefono from vodafonetxt.dcul_cab order by 1"
             
    IntentaEliminarFicheroLog
    NF = FreeFile()
    Open App.Path & "\emitefac.log" For Output As #NF
    i = 0
    RTef.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RTef.EOF
        i = i + 1
        Lbl.Caption = i & "  " & RTef!eltelefono
        Lbl.Refresh
        
        If (i Mod 50) = 0 Then DoEvents
                
                
                'Cogera los datos de la sclien
                mSQL = "select nifclien nifclien,nomclien,domclien,codpobla,pobclien,proclien,operador,CodDirec,sclientfno.codclien"
                mSQL = mSQL & " from sclien,sclientfno  where sclien.codclien=sclientfno.codclien AND "
                mSQL = mSQL & " IdTelefono = " & mdbFac2.texto(RTef!eltelefono)
                
                miRsAux.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                'NO PUEDE SER EOF
                
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
                mSQL = mSQL & ",companyia,TipoIVA , TotalPrevio,BaseImponible,Cuota,Total"
                mSQL = mSQL & ",Apellido1,Apellido2,Nombre,Direccion,CodPostal,Poblacion,Provincia) VALUES ("
                
                
                mSQL = mSQL & mdbFac2.texto(mSerie) & ","  ' Serie
                mSQL = mSQL & mdbFac2.numero(Year(mfecha)) & "," ' Año
                mSQL = mSQL & mdbFac2.numero(mFactura3) & "," ' Número de factura propiamente dicho
                
                mSQL = mSQL & mdbFac2.texto(mfichero) & "," ' Fichero
                mSQL = mSQL & mdbFac2.Fecha(mfecha) & "," ' Fecha
                mSQL = mSQL & mdbFac2.texto(RTef!eltelefono) & "," ' Telefono
                
                
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
                    Else
                        Print #NF, "Codirec: " & RTef!eltelefono & " " & Apell2
                    End If
                End If
                
                SeparaNomclienNomApellidos miRsAux!Nomclien, Nombre, Apell1, Apell2
                mSQL = mSQL & mdbFac2.texto(Apell1) & ","
                mSQL = mSQL & mdbFac2.texto(Apell2) & ","
                mSQL = mSQL & mdbFac2.texto(Nombre) & ","
                mSQL = mSQL & mdbFac2.texto(miRsAux!domclien) & ","
                mSQL = mSQL & mdbFac2.texto(miRsAux!codpobla) & ","
                mSQL = mSQL & mdbFac2.texto(miRsAux!pobclien) & ","
                mSQL = mSQL & mdbFac2.texto(miRsAux!proclien) & ")"
            
                
                
                
                mdbFac2.ejecutar mSQL
                
                        
 
        miRsAux.Close
        
        
        RTef.MoveNext
    Wend
    RTef.Close
        
   'Ajustamos contadores
    vC.Contador = vC.Contador - 1
    vC.IncrementarContador vC.TipoMovimiento
    'Ver contador FAIs
    If vI.Contador <> ContadorFAI_Incial Then
        vI.Contador = ContadorFAI_Incial - 1
        vI.IncrementarContador vI.TipoMovimiento
    End If
    
    
    'ENERO 2015
    'Hay cuotas VODAFONE refacturables
    mSQL = "Select CodCuota,refacturar,valor from tel_desc_cuotas where refacturar >0 "
    RsCuotasRefacturar.Open mSQL, conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    '- Ya tenemos las facturas y ahora montamos los consumos
    mSQL = "concat(Apellido1,'',Apellido2,' ',Nombre) elNombre,telefono"
    mSQL = "select Serie,Ano,NumFact,Fichero,TipoIVA," & mSQL & " from tel_cab_factura where Fichero = " & mdbFac2.texto(mfichero)
    
    RTef.Open mSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Apell1 = DevuelveDesdeBD(conAri, "DescCuota", "tel_desc_cuotas", "CodCuota", "ARI00000", "T")
    i = 0
    While Not RTef.EOF
            
            'If RTef!Telefono = "672305329" Then Stop
            
            
            i = i + 1
        
            Lbl.Caption = "Lineas factura " & CStr(i) & " " & RTef.Fields(6)
            Lbl.Refresh
            DoEvents
            TotalFactura = 0


        
            'Cuotas
            TotalCuotas = CargarCuotasIva(RTef, RsCuotasRefacturar)
            TotalFactura = TotalFactura + TotalCuotas
            
            'Consumos
            TotalConsumo = CargarTraficoNuevo(RTef, TotalConsumoExento)
            TotalFactura = TotalFactura + TotalConsumo
            TotalFactura = TotalFactura + TotalConsumoExento
            
            
            TotalDescuentos = CargarDescuentos(RTef)
        
            'El mimino es sobre cuotas+consumo-descuentos
            ImporteparaElCalculoDelMinimo = TotalFactura - TotalDescuentos
            
            'Mimino ALZIRA
            'CalcularImporteMinimoConsumoNuevo
            'If vParamAplic.TieneTelefonia2 = 1 Then
            mSQL = "Select cuotaminima,fechaalta from sclientfno where idtelefono = '" & RTef!Telefono & "'"
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
 


            'Base Normal
            TotalPrevio = Round(TotalCuotas + TotalConsumo - TotalDescuentos, 2)
            TotalCuotas = Round(TotalPrevio + ImporteAjusteMinimo, 2)
            
            
            'Exenta
            'cab.Base_exenta = totalConsumoExento;
               
            'Cuota IVA
            TotalConsumo = Round((TotalCuotas * mValorIva) / 100, 2)
            'cab.Cuota = (cab.BaseImponible * cab.TipoIVA) / 100;
            '// redondear los valores a 2 decimales
            ' cab.Cuota = decimal.Round(cab.Cuota, 2, MidpointRounding.AwayFromZero);
            'cab.BaseImponible = decimal.Round(cab.BaseImponible, 2, MidpointRounding.AwayFromZero);
            
            
            'cab.Total = cab.BaseImponible + cab.Cuota + cab.Base_exenta ;
            TotalDescuentos = TotalCuotas + TotalConsumo + TotalConsumoExento
            
            mSQL = "UPDATE tel_cab_factura set TotalPrevio=" & DBSet(TotalPrevio, "N")
            mSQL = mSQL & ", BaseImponible =" & DBSet(TotalCuotas, "N")
            mSQL = mSQL & ", Cuota =" & DBSet(TotalConsumo, "N")
            mSQL = mSQL & ", Total =" & DBSet(TotalDescuentos, "N")
            mSQL = mSQL & " WHERE serie = '" & RTef!Serie & "' and ano = " & RTef!Ano & " and numfact =" & RTef!NumFact
            conn.Execute mSQL
            RTef.MoveNext
    Wend
    RTef.Close
        

    
    
    'AHORA HAY QUE INSERTAR EN BD telefono, los detalles y las cuotas
    VodafoneInsertarDetalleLlamadas Lbl
    
    
    
eEmitirFacturasVODAFONE:
    If Err.Number <> 0 Then MuestraError Err.Number, "Emit. facturas vodafone ", Err.Description
    Set miRsAux = Nothing
    Set RTef = Nothing
    Set RsCuotasRefacturar = Nothing
    Set vI = Nothing
    Set vC = Nothing
    If NF >= 0 Then Close #NF
    
End Function


Private Function CargarCuotasIva(ByRef RR As ADODB.Recordset, RecordSetConCuotasRefacturables As ADODB.Recordset) As Currency
Dim i As Integer
Dim Importe As Currency
Dim Cuota As String
Dim J As Integer

            CargarCuotasIva = 0
            
            ' Leemos las cuotas correspondientes a ese número de teléfono
            i = 0
            CadenaInsert = "Select substring(tipo_cuota,1,50) lacuota,sum(importe) importe from vodafonetxt.Dcul_Cab WHERE telef like '" & RR!Telefono & "%' GROUP BY 1"
            miRsAux.Open CadenaInsert, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            CadenaInsert = ""
            While Not miRsAux.EOF
                i = i + 1
                Importe = miRsAux!Importe
                
                
                If Importe <> 0 Then
                    Cuota = CodigoDecripcionCuotaVoda(0, DBLet(miRsAux!LaCuota, "T"))
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
                            
                        Else
                            Err.Raise 513, , "Modo refacturacion INCORRECTO"
                        End If
                    End If
                End If
                'tel_lin_factura_cuotas(Serie,Ano,NumFact,NumLin,CodCuota,DescCuota,Importe,TipoIva)
                CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i
                CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(0, DBLet(miRsAux!LaCuota, "T"))
                CadenaInsert = CadenaInsert & DBSet(Importe, "N") & ",0)"
                CargarCuotasIva = CargarCuotasIva + Importe

                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If CadenaInsert <> "" Then
                CadenaInsert = Mid(CadenaInsert, 2)
                CadenaInsert = "INSERT INTO tel_lin_factura_cuotas(Serie,Ano,NumFact,NumLin,CodCuota,DescCuota,Importe,TipoIva) VALUES " & CadenaInsert
                conn.Execute CadenaInsert
            End If

 

End Function


Private Function CargarTraficoNuevo(ByRef RR As ADODB.Recordset, ByRef consumoSinIVA As Currency) As Currency
Dim i As Integer
Dim Cad As String
        
        
        
        consumoSinIVA = 0
        CargarTraficoNuevo = 0
        CadenaInsert = ""
        'Para todos
        'tel_lin_factura_consumos(Serie,Ano,NumFact,NumLin,CodTipoTrafico,DescTipoTrafico,NumLlamadas,Unidad,Cantidad,Importe,TipoIVA,FechaInicio)
        
        i = 0
        Cad = "Select substring(tipo,1,50) lacuota, sum(duracion) SumaDeDuracion,"
        Cad = Cad & "sum(volumen) SumaDeVolumen ,sum(Importe) SumaDeCoste , count(*) NumLlamadas "
        Cad = Cad & " from vodafonetxt.dntv_cab WHERE telef_ext like '" & RR!Telefono & "%' GROUP BY 1"
        
'            var rs0 = from l in ctx2.DntvCabs
'                      Where L.TelefExt.StartsWith(Telefono)
'                      group l by new { l.TelefExt, l.FactNo, l.Tipo }
'                      into g select new
'                      {
'                          g.Key.TelefExt,
'                          g.Key.FactNo,
'                          g.Key.Tipo,
'                          SumaDeDuracion = g.Sum(l => l.Duracion),
'                          SumaDeVolumen = g.Sum(l => l.Volumen),
'                          SumaDeCoste = g.Sum(l => l.Importe),
'                          NumLlamadas = g.Count()

        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
        
        '        cons.NumLlamadas = dr.NumLlamadas;
        '        cons.CodTipoTrafico = BuscarCodigoTrafico(dr.Tipo, ctx);
        '        cons.DescTipoTrafico = dr.Tipo;
        '        if (dr.SumaDeDuracion == 0)
        '        {
        '            cons.Unidad = "KBytes";
        '            cons.Cantidad = dr.SumaDeVolumen;
        '        }
        '        Else
        '        {
        '            cons.Unidad = "Segundos";
        '            cons.Cantidad = dr.SumaDeDuracion;
        '        }
        '        cons.Importe = dr.SumaDeCoste;
        '        totalTrafico += cons.Importe;
            i = i + 1
            'Serie,Ano,NumFact,NumLin,CodTipoTrafico,DescTipoTrafico,
            'NumLlamadas,Unidad,Cantidad,Importe,TipoIVA)
            CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i
            CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(1, miRsAux!LaCuota)
            CadenaInsert = CadenaInsert & DBSet(miRsAux!NumLlamadas, "N") & ","
            If DBLet(miRsAux!SumaDeDuracion, "N") = 0 Then
                'SOn KB
                CadenaInsert = CadenaInsert & "'KBytes'," & DBSet(miRsAux!SumaDeVolumen, "N", "N")
            Else
                'Son segundos
                CadenaInsert = CadenaInsert & "'Segundos'," & DBSet(miRsAux!SumaDeDuracion, "N", "N")
            End If
            CadenaInsert = CadenaInsert & "," & DBSet(miRsAux!SumaDeCoste, "N") & ",0,'1900-01-01')"
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
        Cad = "Select substring(tipo,1,50) lacuota, sum(importe) SumaDeCoste,count(*) NumLlamadas "
        Cad = Cad & " from vodafonetxt.Dsva_Cab WHERE telef_ext like '" & RR!Telefono & "%' GROUP BY 1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(1, miRsAux!LaCuota)
            CadenaInsert = CadenaInsert & DBSet(miRsAux!NumLlamadas, "N") & ","
            CadenaInsert = CadenaInsert & "'Mensajes'," & DBSet(miRsAux!NumLlamadas, "N") & "," & DBSet(miRsAux!SumaDeCoste, "N", "N")
            CadenaInsert = CadenaInsert & ",0,'1900-01-01')"
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
        Cad = "Select substring(tipo,1,50) lacuota, sum(duracion) SumaDeDuracion,"
        Cad = Cad & "sum(volumen) SumaDeVolumen ,sum(Importe) SumaDeCoste , count(*) NumLlamadas "
        Cad = Cad & " from vodafonetxt.Dntd_cab WHERE telef_ext like '" & RR!Telefono & "%' GROUP BY 1"
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(1, miRsAux!LaCuota)
            CadenaInsert = CadenaInsert & DBSet(miRsAux!NumLlamadas, "N") & ","
            
            CadenaInsert = CadenaInsert & "'KB'," & DBSet(miRsAux!SumaDeDuracion, "N", "N")
            
            CadenaInsert = CadenaInsert & "," & DBSet(miRsAux!SumaDeCoste, "N") & ",0,'1900-01-01')"
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
        Cad = "Select substring(producto,1,50) lacuota,  importe , fecha_inicio,num_llamadas"
        Cad = Cad & " from vodafonetxt.Doim_Cab WHERE telef like '" & RR!Telefono & "%' "
        miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            i = i + 1
            CadenaInsert = CadenaInsert & ", ('" & RR!Serie & "'," & RR!Ano & "," & RR!NumFact & "," & i
            CadenaInsert = CadenaInsert & CodigoDecripcionCuotaVoda(1, miRsAux!LaCuota)
            CadenaInsert = CadenaInsert & DBSet(miRsAux!num_llamadas, "N") & ","
            
            CadenaInsert = CadenaInsert & "'UD',1"
            
            CadenaInsert = CadenaInsert & "," & DBSet(miRsAux!Importe, "N") & ",1," & DBSet(miRsAux!fecha_inicio, "F") & ")"
            consumoSinIVA = miRsAux!Importe

            miRsAux.MoveNext
        
        
        Wend
        miRsAux.Close
            
        'Insertamos
        If CadenaInsert <> "" Then
            CadenaInsert = Mid(CadenaInsert, 2)
            CadenaInsert = "INSERT INTO tel_lin_factura_consumos(Serie,Ano,NumFact,NumLin,CodTipoTrafico,DescTipoTrafico,NumLlamadas,Unidad,Cantidad,Importe,TipoIVA,FechaInicio) VALUES " & CadenaInsert
            conn.Execute CadenaInsert
        End If
End Function



Private Function CodigoDecripcionCuotaVoda(Tipo As Byte, Descrip As String) As String
Dim Cad As String

    If Tipo = 0 Then
        'CUOTAS
        'DescCuota CodCuota  tel_desc_cuotas
      
        Cad = DevuelveDesdeBD(conAri, "CodCuota", "tel_desc_cuotas", "DescCuota", Descrip, "T")
        
        If Cad = "" Then
            Print #NF, "Cuota incorrecta: " & Descrip
            Cad = "INCORR"
        End If
    Else
        'tipo=1
        
        Cad = DevuelveDesdeBD(conAri, "CodTipoTrafico", "tel_desc_consumos", "DescTipoTrafico", Descrip, "T")
        If Cad = "" Then
            Print #NF, "Cuota trafico incorrecta: " & Descrip
            Cad = "INCORR"
        End If
    End If
    CodigoDecripcionCuotaVoda = "," & DBSet(Cad, "T") & "," & DBSet(Descrip, "T") & ","
End Function





  
Private Function CargarDescuentos(ByRef RR As ADODB.Recordset) As Currency
Dim i As Integer
Dim RN As ADODB.Recordset
Dim Cad As String
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
            Cad = "Select codcuota,sum(importe) importe from Tel_lin_factura_cuotas WHERE serie ='" & RR!Serie & "' AND ano = " & RR!Ano & " AND numfact = " & RR!NumFact & " GROUP BY 1"
            RN.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            If Not RN.EOF Then
                Cad = ""
                While Not RN.EOF
                    Cad = Cad & ", " & DBSet(RN!CodCuota, "T")
                    RN.MoveNext
                Wend
                RN.MoveFirst
                
                Cad = Mid(Cad, 2) 'No puede ser EOF
                Cad = "Select * from Tel_desc_cuotas WHERE codcuota IN (" & Cad & ")"
                miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
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
            Cad = "Select CodTipoTrafico,sum(importe) importe from Tel_lin_factura_consumos WHERE serie ='" & RR!Serie & "' AND ano = " & RR!Ano & " AND numfact = " & RR!NumFact & " GROUP BY 1"
            RN.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            If Not RN.EOF Then
                Cad = ""
                While Not RN.EOF
                    Cad = Cad & ", " & DBSet(RN!CodTipoTrafico, "T")
                    RN.MoveNext
                Wend
                RN.MoveFirst
                
                Cad = Mid(Cad, 2) 'No puede ser EOF
                Cad = "Select * from Tel_desc_consumos WHERE CodTipoTrafico IN (" & Cad & ")"
                miRsAux.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            
                
            
            
                '// Comprobamos is hay algún descuento de tráfico disponible.
                'ModAriGesTel.Tel_desc_consumo dc = (from des in ctx.Tel_desc_consumos
                '                                    where des.CodTipoTrafico == lt.CodTipoTrafico
                '                                    select des).FirstOrDefault<ModAriGesTel.Tel_desc_consumo>();
                While Not RN.EOF
                        miRsAux.Find "CodTipoTrafico = " & DBSet(RN!CodTipoTrafico, "T"), , adSearchForward, 1
                    
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
Dim Id As Long

        L.Caption = "Detalle llamada"
        L.Refresh
        Set miRsAux = New ADODB.Recordset
        SQ = "Select max(id) from telefono.Detalle_de_llamadas"
        miRsAux.Open SQ, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Id = 0
        If Not miRsAux.EOF Then Id = DBLet(miRsAux.Fields(0), "N")
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
               
            Id = Id + 1
            SQ = ", (" & Id & "," & DBSet(miRsAux!FACT_NO, "T") & "," & DBSet(miRsAux!telef_ext, "T") & "," & DBSet(miRsAux!Tipo, "T")
            SQ = SQ & "," & DBSet(miRsAux!Tipo, "T") & "," & DBSet(miRsAux!Tipo, "T") & "," & DBSet(miRsAux!Destino, "T", "S") & ","
            SQ = SQ & DBSet(miRsAux!Destino, "T", "S") & "," & DBSet(miRsAux!no_recep, "T", "S") & ",'" & Format(miRsAux!fecha_hora, "mmdd") & "','"
            SQ = SQ & Format(miRsAux!fecha_hora, "hhnn") & "','Segundos'," & DBSet(miRsAux!duracion, "N") & "," & DBSet(miRsAux!Importe, "N") & ")"
            CadenaInsert = CadenaInsert & SQ
            If Len(CadenaInsert) > 30000 Then
                
                L.Caption = "Registro: " & Id
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
            
            L.Caption = "Finalizando " & Id
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
        SQ = "SELECT " & DBSet(mfichero2, "T") & ",320,1," & DBSet(mfichero2, "T") & ",'',Numero_de_telefono,''," & CadenaInsert
        SQ = SQ & ",Codigo_de_trafico,Tipo_de_trafico,count(*),Unidad_de_medida,sum(Cantidad_medida_originada)"
        
        SQ = SQ & ",0,0,sum(importe),'' FROM telefono.detalle_de_llamadas WHERE  Fichero ='CI0643137077'"
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
Dim Cad As String
Dim Origen As String
Dim Rc As ADODB.Recordset
Dim i As Integer
Dim Importe2 As Currency


    L.Caption = "Refacturar cuotas VODAFONE"
    L.Refresh
    
    mFicheroCorto = Right(mfichero, 12)

    
   'tel_desc_cuotas where refacturar>0
    
    
End Function
