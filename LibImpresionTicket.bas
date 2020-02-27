Attribute VB_Name = "LibImpresionTicket"
Option Explicit


'David
'Llamara a esta funcion. Si el tipo de documento 32 (tickets) pone impresion directa, lo dejamos como esta, si no...
' hay que hacer a traves del rpt
Public Sub ImprimirTicketDirecto(NumTicket As String, FechaTicket As Date, Optional Entregado As Currency, Optional Cambio As Currency)   ' (RAFA/ALZIRA 05092006)
Dim Directo As Boolean
Dim cadParam As String
Dim numParam As Byte
Dim cadNomRPT As String
Dim NomImpre As String

    Directo = True
    If Not PonerParamRPT2(32, cadParam, numParam, cadNomRPT, Directo, pPdfRpt, pRptvMultiInforme) Then Directo = True
    'NO lleva pRptvMultiInforme
    
    ' ----  [07/10/2009] [LAURA] : se poner general para impresion directa y crystal reports
    ' -- Establecemos la impresora de ticket
    If vParamTPV.NomImpresora <> "" Then
        If Printer.DeviceName <> vParamTPV.NomImpresora Then
            'guardamos la impresora que habia
            NomImpre = Printer.DeviceName
            'establecemos la de ticket
            EstablecerImpresora vParamTPV.NomImpresora
        End If
    End If
    ' ---- []
    

    If Directo Then
        '-- Impresion directa
        ImprimirElTicketDirecto2 NumTicket, FechaTicket, Not vParamTPV.Redondea2, Entregado, Cambio
        
    Else
        'Establecemos la impresora de ticket
'        If vParamTPV.NomImpresora <> "" Then
'            If Printer.DeviceName <> vParamTPV.NomImpresora Then
'                'guardamos la impresora que habia
'                NomImpre = Printer.DeviceName
'                'establecemos la de ticket
'                EstablecerImpresora vParamTPV.NomImpresora
'            End If
'        End If
    
        '-- Con crystal
        With frmImprimir
            .FormulaSeleccion = " {scafac.codtipom} = 'FTI'" & _
                " and {scafac.numfactu} = " & CStr(NumTicket) & _
                " and {scafac.fecfactu} = Date(" & Year(FechaTicket) & "," & Month(FechaTicket) & "," & Day(FechaTicket) & ")"
                
            .OtrosParametros = ""
            .NumeroParametros = 0
            .SoloImprimir = True
            .EnvioEMail = False
            .Opcion = 93
            .Titulo = "Ticket"
            .NombreRPT = cadNomRPT
            .NombrePDF = pPdfRpt
            .ConSubInforme = True
            .Show vbModal
         
         End With
        
        
        
        
        'sI ABRE EL CAJON
        If vParamTPV.AbreCajon > 0 Then ImprimePorLaCom "", vParamTPV.AbreCajon
              
              
'        'Volver la impresora a la predeterminada
'        If NomImpre <> "" Then EstablecerImpresora NomImpre
    End If
    
    
    ' ----  [07/10/2009] [LAURA] : se poner general para impresion directa y crystal reports
    ' -- Volver la impresora a la predeterminada
    If NomImpre <> "" Then EstablecerImpresora NomImpre
    ' ----- []
End Sub




'Obligo la fecha. Antes NO y la cogia de rsventa
'Public Sub ImprimirTicketDirecto(NumTicket As String, NumAlbTicket1 As String, FechaTicket As Date)  ' (RAFA/ALZIRA 05092006)
Public Sub ImprimirElTicketDirecto2(NumTicket As String, FechaTicket As Date, Precio4Decimales As Boolean, Optional Entregado As Currency, Optional Cambio As Currency)   ' (RAFA/ALZIRA 05092006)

'    Dim NomImpre As String
  '  Dim FechaT As Date
    'Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim SQL As String
    Dim Lin As String ' línea de impresión
    Dim I As Integer
    Dim N As Integer
    Dim ImporteIva As Currency
    Dim EnEfectivo As Boolean
    Dim ClienteMostrador As Boolean
On Error GoTo EImpTickD
    
    '-- Obtenemos cabeceras y pies en un recordset (rs1)
    'SQL = "select * from spatpvg"
    'Set rs11 = New ADODB.Recordset
    'rs11.Open SQL, conn, adOpenForwardOnly
    
    
    'If Not rs11.EOF Then
    If Not vParamTPV Is Nothing Then
        ' Ahora buscamos las cabecera de ticket
        SQL = "select * from scafac where codtipom = 'FTI'" & _
                " and numfactu = " & CStr(NumTicket) & _
                " and fecfactu = '" & Format(FechaTicket, "yyyy-mm-dd") & "'"
        Set rs2 = New ADODB.Recordset
        rs2.Open SQL, conn, adOpenForwardOnly
        If Not rs2.EOF Then
            'Veremos si imprime con 4 decimales, o no
            If Precio4Decimales Then
                'Veremos si HA puesto en paramtpv que solo imprima dos
                If vParamTPV.TkCantidad2Decimales Then Precio4Decimales = False
        
        
            End If
        
            '-- Consultamos la forma de pago pa 2 cosas
            '   Para imprimirla en el pie y para en el caso de contado mostrar entregado
            '   y cambio.
            SQL = "select * from sforpa where codforpa = " & CStr(rs2!codforpa)
            Set rs4 = New ADODB.Recordset
            rs4.Open SQL, conn, adOpenForwardOnly
            If Not rs4.EOF Then
                If rs4!tipforpa = 0 Then EnEfectivo = True
            End If
            '-- Montar las líneas
            SQL = "select * from slifac where codtipom = 'FTI'" & _
                    " and numfactu = " & CStr(NumTicket) & _
                    " and fecfactu = '" & Format(FechaTicket, "yyyy-mm-dd") & "'"
            Set rs3 = New ADODB.Recordset
            rs3.Open SQL, conn, adOpenForwardOnly
            If Not rs3.EOF Then
                '-- Impresión de la cabecera

                For I = 1 To 5
                   ' If Not IsNull(rs11.Fields("cabtick" & CStr(I))) Then
                   '     Lin = LineaCentrada(rs11.Fields("cabtick" & CStr(I)))
                   '     If Lin <> "" Then Printer.Print Lin
                   ' End If
                   Lin = Trim(vParamTPV.CabeceraTiket(I - 1))
                   If Lin <> "" Then Printer.Print LineaCentrada(Lin)
                Next I
                
                

                'FACTURA SIMPLIFICADA
                Printer.Print ""
                Printer.Print LineaCentrada("FACTURA SIMPLIFICADA")
                Printer.Print ""
                
                Lin = CuadraParteI(20, "Número:" & Format(NumTicket, "0000000"))
                SQL = CuadraParteD(20, " Fecha " & Format(FechaTicket, "dd/mm/yyyy hh:mm"))
                Lin = "Número:" & Format(NumTicket, "0000000")
                SQL = "Fecha: " & Format(FechaTicket, "dd/mm/yyyy hh:mm")
                I = Len(Lin & SQL)
                If I < 40 And I > 0 Then
                    I = 40 - I
                    Lin = Lin & Space(I)
                End If
                
                
                Printer.Print Lin & SQL
                
                
                
                
                
                ' ----
                
                rs3.MoveFirst
                Printer.Print ""
                ClienteMostrador = False
                If rs2!codClien = vParamTPV.Cliente Then
                    'Es cliente mostrador. Veamos si ha identificado al cliente o no
                    Lin = Trim(DBLet(rs2!nifClien, "T"))
                    If Lin = "" Then Lin = vParamTPV.nifClien  'Si no tiene NIF es cli varios
                    If Lin = vParamTPV.nifClien Then ClienteMostrador = True
                        
                
                End If
                
                'Cliente mostrador
                If ClienteMostrador Then
                    Lin = CuadraParteI(40, "CLIENTE: " & Format(rs2!codClien, "0000") & "  " & rs2!NomClien)
                    Printer.Print Lin
                Else
                    SQL = Trim(rs2!NomClien & " (" & rs2!nifClien & ")")
                    If Len(SQL) <= 40 Then
                        Lin = CuadraParteI(40, SQL)
                    Else
                        Lin = "Cliente: " & rs2!NomClien
                        Printer.Print Lin
                        Lin = "NIF: " & rs2!nifClien
                    End If
                    Printer.Print Lin
                    'Domicilio
                    Lin = DBLet(rs2!domclien, "T")
                    If Lin <> "" Then Printer.Print Lin
                    'Poblacion
                    Lin = Trim(DBLet(rs2!codpobla, "T") & " - " & DBLet(rs2!pobclien))
                    If Len(Lin) > 2 Then Printer.Print Lin
                    
                    
                    
                End If
               
                Printer.Print ""
                Lin = LineaCentrada("IVA INCLUIDO")
                Printer.Print Lin
                Lin = String(40, "-")
                Printer.Print Lin
                Lin = CuadraParteI(16, "DESCRIPCION") & _
                        CuadraParteD(6, " CANT") & _
                        CuadraParteD(8, "PVP") & _
                        CuadraParteD(10, "IMPORTE")
                Printer.Print Lin
                Lin = String(40, "-")
                Printer.Print Lin
                While Not rs3.EOF
                    '-- Una línea de impresión
                    'FALTA###
                    
                    If Precio4Decimales Then
                        Lin = CuadraParteI(16, Mid(rs3!NomArtic, 1, 16)) & _
                                CuadraParteD(4, Format(rs3!cantidad, "#0")) & _
                                CuadraParteD(10, Format(rs3!precioiv, "#,##0.0000")) & _
                                CuadraParteD(10, Format(Round2(rs3!cantidad * rs3!precioiv, 2), "###,##0.00"))
                    
                    
                    Else
                        'Linea normal
                        Lin = CuadraParteI(16, Mid(rs3!NomArtic, 1, 16)) & _
                                CuadraParteD(6, Format(rs3!cantidad, "##0.00")) & _
                                CuadraParteD(8, Format(rs3!precioiv, "#,##0.00")) & _
                                CuadraParteD(10, Format(Round2(rs3!cantidad * rs3!precioiv, 2), "###,##0.00"))
                    
                    
                    
                        If vParamTPV.TkCantidad2Decimales Then
                            If rs3!cantidad < 0 Then
                                'Para MOIXENT si el es un abono no pone cantidad precio
                                Lin = CuadraParteI(16, Mid(rs3!NomArtic, 1, 16)) & _
                                    CuadraParteD(6, " ") & _
                                    CuadraParteD(8, " ") & _
                                    CuadraParteD(10, Format(Round2(rs3!cantidad * rs3!precioiv, 2), "###,##0.00"))
                            End If
                        End If
                    End If
                    Printer.Print Lin
                    rs3.MoveNext
                Wend
                '-- Impresion del total
                Printer.Print String(40, " ")
                Lin = CuadraParteI(20, "Total ticket: ") & CuadraParteD(20, Format(rs2!TotalFac, "###,###,#0.00"))
                Printer.Print Lin
                
                'Si deglosa IVAS
                '2012 Diciembre  SIEMPRE se desglosaran los iva, obligado
                'If vParamTPV.DesglosaIVATicket Then
                If True Then
                    'Linea en blanco
                    'Lin = String(40, " ")
                    Printer.Print ""
                    
                    'Los tpios de IVA
                    Printer.Print "Detalle desglose IVA"
                    
                    For I = 1 To 3
                        If Not IsNull(rs2.Fields("porciva" & CStr(I))) Then
                            'Lleva TIPO IVA
                            SQL = Format(DBLet(rs2.Fields("porciva" & CStr(I)), "N"), "0.00") & "%"
                            Lin = CuadraParteD(6, SQL)
                            'base
                            ImporteIva = DBLet(rs2.Fields("baseimp" & CStr(I)), "N")
                            SQL = Format(ImporteIva, "0.00")
                            Lin = Lin & CuadraParteD(10, SQL)
                            
                            'iva
                            SQL = Format(DBLet(rs2.Fields("imporiv" & CStr(I)), "N"), "0.00")
                            ImporteIva = ImporteIva + DBLet(rs2.Fields("imporiv" & CStr(I)), "N")
                            Lin = Lin & CuadraParteD(10, SQL)
                            'total
                            
                            SQL = Format(ImporteIva, FormatoImporte)
                            Lin = Lin & CuadraParteD(14, SQL)
                            Printer.Print Lin
                        End If
                    Next I
                    Printer.Print ""
                End If
                '-- (RAFA 15/05/2008) -- Para Quatretonda
                '-- Imprimir la forma de pago
                Lin = CuadraParteI(40, "Forma de pago: " & DBLet(rs4!nomforpa, "T"))
                Printer.Print Lin
                If EnEfectivo Then
                    'Si ha puesto imprimir entragado
                    If Not vParamTPV.TKocultalineaEntregado Then
                        '-- Si han pagado en efectivo mostramos entregado y cambio.
                        Printer.Print String(40, " ")
                        SQL = Format(Entregado, "0.00")
                        Lin = CuadraParteI(20, "Entregado: " & SQL)
                        SQL = Format(Cambio, "0.00")
                        Lin = Lin & CuadraParteD(20, "Cambio: " & SQL)
                        Printer.Print Lin
                    End If
                End If
                
                'Nov 2012
                'Le atendio
                If vParamTPV.TkMostrarTrabajador Then
                    Lin = Trim(vUsu.Nombre)
                    If Lin <> "" Then
                        Lin = "Le atendio: " & Lin
                        If Len(Lin) > 40 Then Lin = Mid("Le atendio: " & vUsu.Login, 1, 40)
                        
                        Printer.Print Lin
                    End If
                    
                End If
                    
                
                '-- Impresion del pie
                Printer.Print String(40, " ")
                For I = 1 To 3
                    'If Not IsNull(rs11.Fields("pietick" & CStr(I))) Then
                    '    Lin = LineaCentrada(rs11.Fields("pietick" & CStr(I)))
                    '    If Lin <> "" Then Printer.Print Lin
                    'End If
                    Lin = Trim(vParamTPV.PieTiket(I - 1))
                    If Lin <> "" Then Printer.Print LineaCentrada(Lin)
                    
                Next I
                For I = 1 To 8
                    Printer.Print String(40, " ")
                Next I
                
                '-- Fin de impresión
                Printer.NewPage
                Printer.EndDoc

                
                
                'Abrir cajon
                            'El primer numero es el numero de caracteres de secuencia.
                            'Ej: 5|27|p|0|25|250|
                            '   Son 5:  27;p;0;25:250
                If vParamTPV.AbreCajon > 0 Then
                    
    
                        'De momento lo pongo a piñon
                        'N = RecuperaValor(vParamTPV.SecuenciaCajon, 1)
                        'Lin = ""
                        'For i = 1 To N
                        '    SQL = RecuperaValor(vParamTPV.SecuenciaCajon, i + 1)
                        '    If IsNumeric(SQL) Then
                        '        Lin = Lin & Chr(SQL)
                        '    Else
                        '        Lin = Lin & SQL
                        '    End If
                        'Next i
                        'Printer.Print Lin
                        ImprimePorLaCom "", vParamTPV.AbreCajon
                End If
                
                
                
            Else
                MsgBox "No se han encontrado lineas del ticket " & CStr(NumTicket) & " de " & Format(FechaTicket, "dd/mm/yyyy"), vbCritical
            End If
            rs3.Close
        Else
            MsgBox "No se ha encontrado el ticket " & CStr(NumTicket) & " de " & Format(FechaTicket, "dd/mm/yyyy"), vbCritical
        End If
        rs2.Close
    Else
        MsgBox "Faltan los parámetros para la impresión del ticket", vbCritical
    End If
    'rs11.Close
    
    ' ----  [07/10/2009] [LAURA] : se poner general para impresion directa y crystal reports
'    'Volver la impresora a la predeterminada
'    EstablecerImpresora NomImpre
    ' ----  []
    
    Exit Sub
EImpTickD:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir ticket."
End Sub

'TipoPuerto :   1 COM      2: LPT
Public Sub ImprimePorLaCom(CADENA As String, TipoPuerto As Byte)
    On Error GoTo EI
    
    Dim nFicSalCajon As Integer
    Dim Puerto As String
    
    'Marzo 2011
    'Puerto = "COM1"
    If TipoPuerto = 2 Then
        Puerto = "LPT" & vParamTPV.ComImpresora
    Else
        'Lo que habia
        Puerto = "COM" & vParamTPV.ComImpresora
    End If
    nFicSalCajon = FreeFile
    
    Open Puerto For Output As #nFicSalCajon
    'En ppio esta secuencia es STANDRD
    Print #nFicSalCajon, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
    Close nFicSalCajon
    
    Exit Sub
EI:
    CADENA = "Error en " & Puerto & ": " & vbCrLf & vbCrLf & Err.Description
    MsgBox CADENA, vbCritical
End Sub


'Private Sub CortaPapel()
'    Printer.Print Chr(29) & Chr(56) & Chr(49)
''    Printer.EndDoc
'End Sub




Private Function LineaCentrada(Lin As String) As String
    Dim queda As Integer
    Dim parte As Integer
    queda = 40 - Len(Lin)
    parte = queda / 2
    If parte Then
        LineaCentrada = String(parte, " ") & Lin & String(queda - parte, " ")
    Else
        LineaCentrada = Lin
    End If
End Function

Private Function CuadraParteD(Longitud As Integer, CADENA As String) As String
    CuadraParteD = Right(String(Longitud, " ") & CADENA, Longitud)
End Function

Private Function CuadraParteI(Longitud As Integer, CADENA As String) As String
    CuadraParteI = Left(CADENA & String(Longitud, " "), Longitud)
End Function

