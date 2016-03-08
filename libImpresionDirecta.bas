Attribute VB_Name = "libImpresionDirecta"
Option Explicit



Private Const LineasPorHoja = 13
Private Const MargenIzdo = 6   'Si las pruebas las estoy haciendo o no. Pruebas=6  Real=0
                
                
Private Const ModoImpresion = 2
    ' 0 .- Solo en modo DEBUG. No envia a la impresora
    ' 1 .- Objeto PRINTER
    ' 2 .- Direcatamente sobre LPT
        
    '  Diferencia IMPORTANTE.
    ' SI imprimimos directamente seleccionando la fuente en la impresora hay 36 LINEAS
    ' ni una ni mas ni una menos
    ' Sin embargo con el TPRINTER podemos llegar a las 37 lineas
    ' .....  como suena. ASIN ES!!!!!
        
Dim Cabecera As Collection
Dim Lineas As Collection
Dim Importes As Collection
                    
Dim rs1 As ADODB.Recordset
Dim LasObservaciones As String
Dim NF As Integer
                
    
                
                
                
Private Sub AccionesIniciales()
    
    If ModoImpresion = 1 Then
            Printer.Font = "Courier New"
            Printer.FontSize = "10"
    ElseIf ModoImpresion = 2 Then
        NF = FreeFile
        'Open "d:\t1.txt" For Output As #NF
        Open "LPT1" For Output As #NF
        
        
    End If
    LasObservaciones = ""
End Sub
                
                
                
                
                
                
                
                
'************************************************************
'************************************************************
'
'       Impresion directa. Para facturas, albaranes
'
'
'
'       De momento para 4tonda
'
'           COn lo cual:  El papel es el mismo para todo

Public Sub ImprimirDirectoAlb(cadSelect As String)
    Dim NomImpre As String
  '  Dim FechaT As Date
    Dim rsIVA As ADODB.Recordset
    Dim vFactu As CFactura
    
    Dim SQL As String
    Dim Lin As String ' línea de impresión
    Dim i As Integer
    
    
    
On Error GoTo EImpD
    

        
        AccionesIniciales
        
        
        Set rs1 = New ADODB.Recordset
        
       
        
        'Tipos de IVA
        Set rsIVA = New ADODB.Recordset
        rsIVA.Open "Select * from tiposiva", ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
        
        
        
        
        'Cabecera del albaran
        SQL = "select * from scaalb WHERE " & cadSelect
        rs1.Open SQL, conn, adOpenForwardOnly
        
        
        Lin = Space(MargenIzdo + 45) & "ALB.   " & rs1!codtipom & Format(rs1!NumAlbar, "000000") & Space(12) & Format(rs1!FechaAlb, "dd/mm/yyyy")
        Set Cabecera = New Collection
        'EN la impresora se alineara la linea roja del cabezal con la linea superiror del papel impreso (en verde)
        'Añadairemos una linea en blanco
        Cabecera.Add " "
        Cabecera.Add Lin
        Cabecera.Add Space(MargenIzdo + 45)
        
        'Lineas 2 a 7 , datos cliente  nomclien  domclien  codpobla  pobclien  proclien  nifclien
        CargaEncabezado2 0, rs1
        
        
        'Leo estos valores para el final del albaran dtoppago dtognral
        Set vFactu = New CFactura
        vFactu.DtoPPago = rs1!DtoPPago
        vFactu.DtoGnral = rs1!DtoGnral
        vFactu.Cliente = rs1!codclien

        If Not vFactu.CalcularDatosFactura(cadSelect, "scaalb", "slialb", False) Then
            MsgBox "MAAAL"
        End If
        
        
        'Veo el campo observaciones
        LasObservaciones = DBLet(rs1!observa01, "T")
        
        'Cerramos el rs
        rs1.Close
        
        
        SQL = "select slialb.*,codigiva,numserie,fecvigen from slialb,sartic where slialb.codartic=sartic.codartic AND "
        SQL = SQL & Replace(cadSelect, "scaalb", "slialb") & " ORDER by numlinea"
        rs1.Open SQL, conn, adOpenForwardOnly
        
        
        Set Lineas = New Collection
        While Not rs1.EOF
            
            'Las lineas correspondientes
            Lin = Right(Space(16) & rs1!codArtic, 16)  '16 es la longiyud de codartic
            Lin = Space(MargenIzdo) & Lin
            Lin = Lin & " " & Left(rs1!NomArtic & Space(30), 30)
            
            Lin = Lin & Right(Space(9) & Format(rs1!Cantidad, FormatoCantidad), 9) & Space(2)
            Lin = Lin & Right(Space(10) & Format(rs1!precioar, FormatoPrecio), 10)
            'El IVA.
            rsIVA.Find "codigiva = " & rs1!codigiva, , adSearchForward, 1
            If rsIVA.EOF Then
                Lin = Lin & " * "
            Else
                Lin = Lin & " " & Format(rsIVA!PorceIVA, "00")
            End If
            Lin = Lin & Right(Space(15) & Format(rs1!ImporteL, FormatoPrecio), 15)
            Lineas.Add Lin
            'El numero de serie
            Lin = DBLet(rs1!numSerie, "T")
            If Lin <> "" Then
                If Not IsNull(rs1!fecvigen) Then Lin = Lin & " (" & Format(rs1!fecvigen, "dd/mm/yy") & ")"
                
                If Not IsNull(rs1!numLote) Then Lin = Lin & "  NºLote: " & rs1!numLote
                
                'Esto estaba
                Lin = Space(10) & "N. Reg: " & Lin
                
                'Modificacion Febrero 2012
                'Fecha vigencia y nºlote
                
                Lineas.Add Lin
            End If
            rs1.MoveNext
            
            
        Wend
        
        
        rs1.Close
        rsIVA.Close
        
        'Añado las obseraciones
        If LasObservaciones <> "" Then
            Lineas.Add " "  'Un espacio en blanco
            Lineas.Add Space(MargenIzdo) & "Observaciones"
            Lineas.Add Space(MargenIzdo) & "  -" & LasObservaciones
        End If
        
        'TRozo final de los importes
        AjusteLineasImportes
        
        'Linea uno. SEGURO QUE LA IMPRIME
        '--------------------------------
        'Campo BAse imponible. Empieza hasta el 41, si alineamos a la derecha
        Lin = Format(vFactu.TotalFac, FormatoImporte)
        Lin = LineaImportes(vFactu.BaseIVA1, vFactu.PorceIVA1, vFactu.ImpIVA1, vFactu.PorceIVA1RE, vFactu.ImpIVA1RE, Lin)
        Importes.Add Lin
        
        If vFactu.BaseIVA2 <> 0 Then
            Lin = LineaImportes(vFactu.BaseIVA2, vFactu.PorceIVA2, vFactu.ImpIVA2, vFactu.PorceIVA2RE, vFactu.ImpIVA2RE, "")
        Else
            Lin = ""
        End If
        Importes.Add Lin
        
        If vFactu.BaseIVA3 <> 0 Then
            Lin = LineaImportes(vFactu.BaseIVA3, vFactu.PorceIVA3, vFactu.ImpIVA3, vFactu.PorceIVA3RE, vFactu.ImpIVA3RE, "")
        Else
            Lin = ""
        End If
        Importes.Add Lin
                
        
        
        'Ya tenemos todos los datos
        'Ahora manadmos a la impresora
        ImprimeEnPapel
        
        
        
EImpD:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir directo."
        Err.Clear
    End If
    
    
    Set Cabecera = Nothing
    Set Lineas = Nothing
    Set Importes = Nothing
    Set rsIVA = New ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Exit Sub
    
End Sub
        
Private Sub AjusteLineasImportes()
    'Linea en blaco deonde van los cuadrados de BImpo, porceta....
    Set Importes = New Collection
    Importes.Add " "

    If ModoImpresion = 2 Then
        'SOlo tiro uno p'abajo
    Else
        Importes.Add " "
    End If
End Sub


Private Sub ImprimeEnPapel()
    Dim i As Integer
    Dim J As Integer
    Dim PagActual As Integer
    Dim Lin As String
    Dim Impor As Currency
    Dim NumeroPaginas As Integer
        'AHORA IMPRIMIMOS.
        'TEnemos cargada las lineas
        NumeroPaginas = ((Lineas.Count - 1) \ LineasPorHoja) + 1
        i = 0
        PagActual = 1
        For J = 1 To Lineas.Count
            
            
            If i = 0 Then
                '***********************************************************
                'Imprimir cabecera
                For i = 1 To Cabecera.Count
                    ImprimeLaLinea Cabecera(i)
                Next i
                i = 0
                'Si hay mas de una hoja pongo tambien el numero de hoja
                If NumeroPaginas > 1 Then
                    Lin = Space(MargenIzdo + 45) & "Pag: " & PagActual & " / " & NumeroPaginas
                    ImprimeLaLinea Lin
                Else
                    ImprimeLaLinea " "
                End If
                ImprimeLaLinea " "
                ImprimeLaLinea " "
                
                PagActual = PagActual + 1
            End If
            
            ImprimeLaLinea Lineas(J)
            i = i + 1
            
            'Si es la ultima linea NO hacemos nada
            If J < Lineas.Count Then
                If i = LineasPorHoja Then
                    ImprimeLaLinea " ": ImprimeLaLinea " ":
                    If ModoImpresion = 1 Then ImprimeLaLinea " "
                    ImprimeLaLinea Space(50) & "** ** **" 'los importes
                    'Linea en blaco deonde van los cauadrados de BImpo, porceta....
                    ' y las lineas finales
                    'Ha rellenado todas. Si hay mas lineas que imprimir entonces
                    
                        
                    For i = 1 To 5
                        ImprimeLaLinea " "
                    Next i
        
                    i = 0
                End If
            End If
        
        Next
        
        
        'Para situar el cabezal en la impresion
        If i < LineasPorHoja And i <> 0 Then
            'Ha impreso i lineas
            'Hasta las 10 que caben...
            i = LineasPorHoja - i
            While i > 0
                ImprimeLaLinea ""
                i = i - 1
            Wend
            
        End If
        
        'Los importes
        For J = 1 To Importes.Count
            ImprimeLaLinea Importes.item(J)
        Next
        
        'Final hoja
        '--------------------
        If ModoImpresion = 1 Then
            Printer.EndDoc
        Else
            If ModoImpresion = 2 Then
                'Re situo el papel donde le toca

                For J = 1 To 3
                    ImprimeLaLinea " "
                Next
            
            
            
            
                Close (NF)
            End If
        End If
        
    'Volver la impresora a la predeterminada
    'EstablecerImpresora NomImpre
    
End Sub


Private Function LineaImportes(BaseIVA As Currency, PorceIVA As Currency, ImpIVA As Currency, IvaRE As Currency, ImpIVARE As Currency, TotalFac As String) As String
Dim Lin As String
    
        Lin = Space(17) & Format(BaseIVA, FormatoImporte)
        Lin = Right(Lin, 17) '17 es la longiyud de bas imponible
        Lin = Space(MargenIzdo + 16) & Lin
        Lin = Lin & "  " & Right(Space(5) & Format(PorceIVA, FormatoPorcen), 5)
         Lin = Lin & " "
        Lin = Lin & Right(Space(11) & Format(ImpIVA, FormatoImporte), 11)
        If IvaRE = 0 Then
            'No lleva % retencion
            Lin = Lin & Space(17)
        Else
            'SI LLEVA
            
        End If
        
        LineaImportes = Lin & Right(Space(16) & TotalFac, 16)
        
        
End Function


'Como los campos del albaran y de la factura son los mismos...
' Paso Opcion por si acaso tengo que hacer algo a las facturas o a los albaranes...
Private Sub CargaEncabezado2(Opcion As Byte, ByRef RS As ADODB.Recordset)
Dim L As String
        L = Space(35) & Format(RS!codclien, "000") & Space(15)
        L = Mid(L, 1, (MargenIzdo + 45)) & RS!Nomclien
        'linea 4
        Cabecera.Add L
        Cabecera.Add Space(MargenIzdo + 45) & DBLet(RS!domclien, "T")
        Cabecera.Add Space(MargenIzdo + 45) & RS!pobclien
        Cabecera.Add Space(MargenIzdo + 45) & Format(RS!codpobla, "00000") & " " & RS!proclien
        Cabecera.Add Space(MargenIzdo + 45) & "C.I.F.: " & RS!NIFClien
        L = Space(MargenIzdo + 2) & vEmpresa.nomempre & Space(40)
        L = Mid(L, 1, MargenIzdo + 45) & "Forma de pago: " & DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "Codforpa", RS!codforpa)
        Cabecera.Add L
        Cabecera.Add Space(MargenIzdo + 2) & vParam.DomicilioEmpresa
        L = Space(MargenIzdo + 2) & vParam.CPostal & " " & vParam.Poblacion & " " & vParam.Provincia
        Cabecera.Add L
        L = Space(MargenIzdo + 2) & "Tfno: " & vParam.Telefono & " " & vParam.CifEmpresa
        Cabecera.Add L
        
End Sub

Private Sub ImprimeLaLinea(Linea As String)
    Debug.Print Linea
    If ModoImpresion = 0 Then Exit Sub  'Solo debug
    If ModoImpresion = 1 Then
        Printer.Print Linea
    Else
        Print #NF, Linea
    End If
    
End Sub








'------------------------------------------------------
' FACTURAS TPV

Public Sub ImprimirDirectoFact(cadSelect As String)
    Dim NomImpre As String
  '  Dim FechaT As Date

    Dim rsIVA As ADODB.Recordset
    Dim vFactu As CFactura
    
    Dim SQL As String
    Dim Lin As String ' línea de impresión
    Dim TieneObsAlbaran As Integer
    
    
    
On Error GoTo EImpD
    
    'Establecemos la impresora de ticket
'    If vParamTPV.NomImpresora <> "" Then
'        If Printer.DeviceName <> vParamTPV.NomImpresora Then
'            'guardamos la impresora que habia
'            NomImpre = Printer.DeviceName
'            'establecemos la de ticket
'            EstablecerImpresora vParamTPV.NomImpresora
'        End If
'    End If
        
        
        AccionesIniciales
        
        Set rs1 = New ADODB.Recordset
        
        
        
        
        
        
        'Lineas de albaranes
        'SQL:
            
        'Guardo las obseraviaciones
        SQL = " FROM scafac INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom AND "
        SQL = SQL & " scafac.numfactu=scafac1.numfactu AND scafac.fecfactu=scafac1.fecfactu "
        SQL = SQL & " WHERE " & cadSelect
        
        rs1.Open "Select observa1,scafac1.numalbar " & SQL, conn, adOpenForwardOnly
        TieneObsAlbaran = 0
        While Not rs1.EOF
            SQL = DBLet(rs1!observa1, "T")
            Lin = "[" & Format(rs1!NumAlbar, "000000") & "]"
            rs1.MoveNext
            If Not rs1.EOF Then TieneObsAlbaran = 1 'Para que pinte el numero de albaran
            If SQL <> "" Then
                If TieneObsAlbaran = 1 Then SQL = Lin & "   " & SQL
                LasObservaciones = LasObservaciones & "- " & SQL & "|"
            End If
        Wend
        rs1.Close
        
        
        
        
        'El select para las lineas de albaran
        SQL = " FROM ((scafac INNER JOIN scafac1 ON ((scafac.codtipom=scafac1.codtipom) AND "
        SQL = SQL & " (scafac.numfactu=scafac1.numfactu)) AND (scafac.fecfactu=scafac1.fecfactu)) "
        SQL = SQL & " INNER JOIN slifac ON ((((scafac1.codtipom=slifac.codtipom) AND "
        SQL = SQL & " (scafac1.numfactu=slifac.numfactu)) AND (scafac1.fecfactu=slifac.fecfactu)) AND "
        SQL = SQL & " (scafac1.codtipoa=slifac.codtipoa)) AND (scafac1.numalbar=slifac.numalbar)) "
        SQL = SQL & " INNER JOIN sartic ON slifac.codartic=sartic.codartic"
        'Y el albaran
        SQL = SQL & " AND " & cadSelect

        
        
        
        'Tipos de IVA
        Set rsIVA = New ADODB.Recordset
        rsIVA.Open "Select * from tiposiva", ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
        
        
        
        
        'Cabecera del albaran
        Lin = "select * from scafac WHERE " & cadSelect
        rs1.Open Lin, conn, adOpenForwardOnly
        
        
        Lin = Space(MargenIzdo + 45) & "FAC.   " & rs1!codtipom & Format(rs1!NumFactu, "000000") & Space(12) & Format(rs1!FecFactu, "dd/mm/yyyy")
        Set Cabecera = New Collection
        'EN la impresora se alineara la linea roja del cabezal con la linea superiror del papel impreso (en verde)
        'Añadairemos una linea en blanco
        Cabecera.Add " "
        Cabecera.Add Lin
        Cabecera.Add Space(MargenIzdo + 45)
        
        'Lineas 2 a 7 , datos cliente  nomclien  domclien  codpobla  pobclien  proclien  nifclien
        CargaEncabezado2 1, rs1
        
        
        'Leo estos valores para el final del albaran dtoppago dtognral
        Set vFactu = New CFactura
        vFactu.DtoPPago = rs1!DtoPPago
        vFactu.DtoGnral = rs1!DtoGnral
        vFactu.Cliente = rs1!codclien
        vFactu.NumFactu = rs1!NumFactu
        vFactu.FecFactu = rs1!FecFactu
        vFactu.codtipom = rs1!codtipom
        
        'Cerramos el rs
        rs1.Close
        
        
        
        Lin = "select slifac.*,codigiva,numserie " & SQL
        Lin = Lin & " ORDER BY numalbar,numlinea"
        rs1.Open Lin, conn, adOpenForwardOnly
        
        
        Set Lineas = New Collection
        While Not rs1.EOF
            
            'Las lineas correspondientes
            Lin = Right(Space(16) & rs1!codArtic, 16)  '16 es la longiyud de codartic
            Lin = Space(MargenIzdo) & Lin
            Lin = Lin & " " & Left(rs1!NomArtic & Space(30), 30)
            
            Lin = Lin & Right(Space(9) & Format(rs1!Cantidad, FormatoCantidad), 9) & Space(2)
            Lin = Lin & Right(Space(10) & Format(rs1!precioar, FormatoPrecio), 10)
            'El IVA.
            rsIVA.Find "codigiva = " & rs1!codigiva, , adSearchForward, 1
            If rsIVA.EOF Then
                Lin = Lin & " * "
            Else
                Lin = Lin & " " & Format(rsIVA!PorceIVA, "00")
            End If
            Lin = Lin & Right(Space(15) & Format(rs1!ImporteL, FormatoPrecio), 15)
            Lineas.Add Lin
            'El numero de serie
            Lin = DBLet(rs1!numSerie, "T")
            If Lin <> "" Then
                Lin = Space(14) & " N. Reg: " & Space(12) & Lin
                Lineas.Add Lin
            End If
            rs1.MoveNext
            
            
        Wend
        rs1.Close
        rsIVA.Close
        
        'Las observaciones de la factura
        'Las tenemos cargadas, empipadas, en LasObservaciones
        If LasObservaciones <> "" Then
            Lineas.Add " "  'Un espacio en blanco
            Lineas.Add Space(MargenIzdo) & "Observaciones"
            While LasObservaciones <> ""
                TieneObsAlbaran = InStr(1, LasObservaciones, "|")
                If TieneObsAlbaran = 0 Then
                    LasObservaciones = ""
                Else
                    SQL = Mid(LasObservaciones, 1, TieneObsAlbaran - 1)
                    LasObservaciones = Mid(LasObservaciones, TieneObsAlbaran + 1)
                    Lineas.Add Space(MargenIzdo) & SQL
                End If
            Wend
        End If
        
        'Los importes. Los cargo desde la factura
        If Not CargarImportesDesdeFactura(vFactu, Lin) Then
            If Not vFactu.CalcularDatosFactura(cadSelect, "scafac", "slifac", False) Then
                MsgBox "Importes factura NO encontrados NI calculados", vbExclamation
            Else
                MsgBox "Importes factura NO encontrados. Se han calculado para la impresion", vbExclamation
            End If
        End If
        
        'TRozo final de los importes
        AjusteLineasImportes
        'Linea uno. SEGURO QUE LA IMPRIME
        '--------------------------------
        
        'Voy a cargar todos los datos de  importes de la factura
        
        Lin = Format(vFactu.TotalFac, FormatoImporte)
       
        
        Lin = LineaImportes(vFactu.BaseIVA1, vFactu.PorceIVA1, vFactu.ImpIVA1, vFactu.PorceIVA1RE, vFactu.ImpIVA1RE, Lin)
        Importes.Add Lin
        
        If vFactu.BaseIVA2 <> 0 Then
            Lin = LineaImportes(vFactu.BaseIVA2, vFactu.PorceIVA2, vFactu.ImpIVA2, vFactu.PorceIVA2RE, vFactu.ImpIVA2RE, "")
        Else
            Lin = ""
        End If
        Importes.Add Lin
        
        If vFactu.BaseIVA3 <> 0 Then
            Lin = LineaImportes(vFactu.BaseIVA3, vFactu.PorceIVA3, vFactu.ImpIVA3, vFactu.PorceIVA3RE, vFactu.ImpIVA3RE, "")
        Else
            Lin = ""
        End If
        Importes.Add Lin
                
        
        
        'Ya tenemos todos los datos
        'Ahora manadmos a la impresora
        ImprimeEnPapel
        
        
        
EImpD:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir directo."
        Err.Clear
    End If
    
    
    Set Cabecera = Nothing
    Set Lineas = Nothing
    Set Importes = Nothing
    Set rsIVA = New ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Exit Sub
    
End Sub





'------------------------------------------------------
' REimpresion de facturas. Pone lo del albaran y eso



Public Sub ReImprimirDirectoFact(cadSelect As String)
    
  '  Dim FechaT As Date

    Dim vFactu As CFactura
    Dim Grupo As String
    Dim SQL As String
    Dim Lin As String ' línea de impresión
    Dim i As Integer
    Dim NumeroPaginas  As Integer
    Dim Importe As Currency
    Dim Albaran As String
On Error GoTo EImpD
    
        
        
      
        
        
        Set rs1 = New ADODB.Recordset
        
        AccionesIniciales
        
        
        
        'Cogeremos. los albaranes de las facturas y los articulos que tengan nºregistro
        'SQL:
        SQL = "Select scafac.*,slifac.*,CodTraba,FechaAlb,numSerie"
        SQL = SQL & " FROM ((scafac INNER JOIN scafac1 ON ((scafac.codtipom=scafac1.codtipom) AND "
        SQL = SQL & " (scafac.numfactu=scafac1.numfactu)) AND (scafac.fecfactu=scafac1.fecfactu)) "
        SQL = SQL & " INNER JOIN slifac ON ((((scafac1.codtipom=slifac.codtipom) AND "
        SQL = SQL & " (scafac1.numfactu=slifac.numfactu)) AND (scafac1.fecfactu=slifac.fecfactu)) AND "
        SQL = SQL & " (scafac1.codtipoa=slifac.codtipoa)) AND (scafac1.numalbar=slifac.numalbar)) "
        SQL = SQL & " INNER JOIN sartic ON slifac.codartic=sartic.codartic"
        
        'Y el albaran
        SQL = SQL & " AND " & cadSelect
        
        rs1.Open SQL, conn, adOpenForwardOnly
        
        
        Lin = Space(MargenIzdo + 45) & "FAC.   " & rs1!codtipom & Format(rs1!NumFactu, "000000") & Space(12) & Format(rs1!FecFactu, "dd/mm/yyyy")
        Set Cabecera = New Collection
        'EN la impresora se alineara la linea roja del cabezal con la linea superiror del papel impreso (en verde)
        'Añadairemos una linea en blanco
        Cabecera.Add " "
        Cabecera.Add Lin
        Cabecera.Add Space(MargenIzdo + 45)
        
        'Lineas 2 a 7 , datos cliente  nomclien  domclien  codpobla  pobclien  proclien  nifclien
        CargaEncabezado2 1, rs1
        
        
        'Leo estos valores para el final del albaran dtoppago dtognral
        Set vFactu = New CFactura
        vFactu.DtoPPago = rs1!DtoPPago
        vFactu.DtoGnral = rs1!DtoGnral
        vFactu.Cliente = rs1!codclien
        vFactu.NumFactu = rs1!NumFactu
        vFactu.FecFactu = rs1!FecFactu
        vFactu.codtipom = rs1!codtipom
        
        
        'En sql tendremos los numeros de lote
        SQL = ""
        Grupo = ""
        'vamos imprimiendo los albaranes
        Set Lineas = New Collection
        i = 0
        While Not rs1.EOF
            Lin = rs1!codTipoa & Format(rs1!NumAlbar, "0000000")
            If Lin <> Grupo Then
                If Grupo <> "" Then LineaAlbaranFactura Albaran, Importe, SQL, i
                
            
                Grupo = Lin
                Lin = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", rs1!CodTraba)
                If Lin <> "" Then Lin = " Venta realizada por " & Lin
                Albaran = "Albarán: " & Grupo & " de fecha " & Format(rs1!FechaAlb, "dd/mm/yyyy") & " " & Lin
                'Faltara añadir el importe
                Importe = 0
                
                SQL = "|" 'Llevaremos los nº de lote en este albaran
            
            End If
            'El numero de serie
            Lin = DBLet(rs1!numSerie, "T")
            If Lin <> "" Then
                If InStr(1, SQL, "|" & Lin & "|") = 0 Then SQL = SQL & Lin & "|"
                    
            End If
            Importe = Importe + rs1!ImporteL
            rs1.MoveNext
        Wend
        rs1.Close
        LineaAlbaranFactura Albaran, Importe, SQL, i
        

        
        'Los importes. Los cargo desde la factura
        If Not CargarImportesDesdeFactura(vFactu, Lin) Then
            If Not vFactu.CalcularDatosFactura(cadSelect, "scafac", "slifac", False) Then
                MsgBox "Importes factura NO encontrados NI calculados", vbExclamation
            Else
                MsgBox "Importes factura NO encontrados. Se han calculado para la impresion", vbExclamation
            End If
        End If
        
        
        'TRozo final de los importes
        AjusteLineasImportes
        
        'Linea uno. SEGURO QUE LA IMPRIME
        '--------------------------------
        'Campo BAse imponible. Empieza hasta el 41, si alineamos a la derecha
        Lin = Format(vFactu.TotalFac, FormatoImporte)
        Lin = LineaImportes(vFactu.BaseIVA1, vFactu.PorceIVA1, vFactu.ImpIVA1, vFactu.PorceIVA1RE, vFactu.ImpIVA1RE, Lin)
        Importes.Add Lin
        
        If vFactu.BaseIVA2 <> 0 Then
            Lin = LineaImportes(vFactu.BaseIVA2, vFactu.PorceIVA2, vFactu.ImpIVA2, vFactu.PorceIVA2RE, vFactu.ImpIVA2RE, "")
        Else
            Lin = ""
        End If
        Importes.Add Lin
        
        If vFactu.BaseIVA3 <> 0 Then
            Lin = LineaImportes(vFactu.BaseIVA3, vFactu.PorceIVA3, vFactu.ImpIVA3, vFactu.PorceIVA3RE, vFactu.ImpIVA3RE, "")
        Else
            Lin = ""
        End If
        Importes.Add Lin
                
        
        
        'Ya tenemos todos los datos
        'Ahora manadmos a la impresora
        'NumeroPaginas = ((i - 1) \ LineasPorHoja) + 1
        'If I > 13 Then Stop
        ImprimeEnPapel
        
        
        
EImpD:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir directo."
        Err.Clear
    End If
    
    
    Set Cabecera = Nothing
    Set Lineas = Nothing
    Set Importes = Nothing
    Set rs1 = New ADODB.Recordset
    Exit Sub
    
End Sub


Private Sub LineaAlbaranFactura(L As String, Importe As Currency, ArticulosConNumeroSerie As String, ByRef ContadorDeLineas As Integer)
Dim i As Integer
        L = Space(MargenIzdo) & L & Space(30)
        L = Mid(L, 1, 78)
        L = L & Right(Space(15) & Format(Importe, FormatoImporte), 15)
        Lineas.Add L
        ContadorDeLineas = ContadorDeLineas + 1
        
        If ArticulosConNumeroSerie <> "|" Then
            ArticulosConNumeroSerie = Mid(ArticulosConNumeroSerie, 2)
            i = 1
            Lineas.Add ""
            ContadorDeLineas = ContadorDeLineas + 1
            
            While i <> 0
                i = InStr(1, ArticulosConNumeroSerie, "|")
                If i > 0 Then
                    L = Mid(ArticulosConNumeroSerie, 1, i - 1)
                    ArticulosConNumeroSerie = Mid(ArticulosConNumeroSerie, i + 1)
                    L = Space(14) & " N. Reg: " & Space(12) & L
                    Lineas.Add L
                    ContadorDeLineas = ContadorDeLineas + 1
                End If
            Wend
        End If
End Sub


Private Function CargarImportesDesdeFactura(ByRef F As CFactura, ByRef auxiliar As String) As Boolean
    CargarImportesDesdeFactura = False
    auxiliar = "Select * from scafac where codtipom=" & DBSet(F.codtipom, "T")
    auxiliar = auxiliar & " AND numfactu=" & DBSet(F.NumFactu, "N")
    auxiliar = auxiliar & " AND fecfactu=" & DBSet(F.FecFactu, "F")
    rs1.Open auxiliar, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If rs1.EOF Then
        
    
    
    Else
        CargarImportesDesdeFactura = True
        
        F.BaseIVA1 = DBLet(rs1!baseimp1, "N")
        F.PorceIVA1 = DBLet(rs1!porciva1, "N")
        F.ImpIVA1 = DBLet(rs1!imporiv1, "N")
        F.PorceIVA1RE = DBLet(rs1!porciva1re, "N")
        F.ImpIVA1RE = DBLet(rs1!imporiv1re, "N")
        
        
        
        F.BaseIVA2 = DBLet(rs1!baseimp2, "N")
        F.PorceIVA2 = DBLet(rs1!porciva2, "N")
        F.ImpIVA2 = DBLet(rs1!imporiv2, "N")
        F.PorceIVA2RE = DBLet(rs1!porciva2re, "N")
        F.ImpIVA2RE = DBLet(rs1!imporiv2re, "N")
        
        F.BaseIVA3 = DBLet(rs1!baseimp3, "N")
        F.PorceIVA3 = DBLet(rs1!porciva3, "N")
        F.ImpIVA3 = DBLet(rs1!imporiv3, "N")
        F.PorceIVA3RE = DBLet(rs1!porciva3re, "N")
        F.ImpIVA3RE = DBLet(rs1!imporiv3re, "N")
        
        F.TotalFac = rs1!TotalFac
            
    
    End If
    rs1.Close
End Function

