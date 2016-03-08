Attribute VB_Name = "libAlmagrupoComun"
Option Explicit


Dim miSQL As String

Public Function DatosGeneradosExportacionAlmagrupo(EnviarStocks As Boolean, EnviarConsumos As Boolean, ByRef L As Label) As Boolean
Dim NF1 As Integer
Dim NF2 As Integer
Dim J As Integer
Dim Anyo As Integer
Dim mes As Byte
Dim Asociado As String



On Error GoTo eDatosGeneradosExportacionAlmagrupo
    DatosGeneradosExportacionAlmagrupo = False
    
    Set miRsAux = New ADODB.Recordset

    miSQL = "Select count(*) from   salmagrupo  " 'todo lo que este en la tabla
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim = 0 Then
        If Not L Is Nothing Then
            MsgBox "ningun dato generado", vbExclamation
            Exit Function
        End If
    End If
    
    
    miSQL = DevuelveDesdeBD(1, "idasociado", "salmagrupoparam", "1", "1")
    If miSQL = "" Then
        MsgBox "Error codigo asociado", vbExclamation
        Exit Function
    End If
    Asociado = miSQL
  
    
    
    

    
    
        NF1 = -1: NF2 = -1
        If EnviarConsumos Then
            NF1 = FreeFile
            miSQL = FijaNombreFichero(mes, Anyo, True, Asociado)
            Open miSQL For Output As #NF1
        End If
        
         If EnviarStocks Then
            If NF1 >= 0 Then
                NF2 = NF1 + 1
            Else
                NF2 = FreeFile
            End If
            miSQL = FijaNombreFichero(mes, Anyo, False, Asociado)
            Open miSQL For Output As #NF2
        End If
        
    
    
       miSQL = "Select salmagrupo.*,referprov,codtelem,nomunbre,sunida.codunida,unicajas,NomArtic,preciouc from  salmagrupo,sartic,sunida  where salmagrupo.codartic=sartic.codartic  and sartic.codunida=sunida.codunida"
       'miSQL = miSQL & " AND  mes =" & mes & " AND anyo = " & Anyo & " ORDER By cifproveedor,salmagrupo.codartic"
       miSQL = miSQL & " ORDER By cifproveedor,salmagrupo.codartic"
       miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
       While Not miRsAux.EOF
           
            'If miRsAux!codArtic = "0010000389" Then Stop
           
           If NF1 > 0 Then
                'compras
                miSQL = DBLet(miRsAux!cifproveedor, "T")
                If miSQL = "" Then miSQL = "S"
                If miSQL <> "S" Then
                
                    miSQL = miRsAux!cifproveedor & "#"
                    miSQL = miSQL & miRsAux!nomproveedor & "#"
                    miSQL = miSQL & miRsAux!Anyo & "#"
                    miSQL = miSQL & miRsAux!mes & "#"
                    miSQL = miSQL & miRsAux!codArtic & "#"
                    miSQL = miSQL & miRsAux!referprov & "#"
                    miSQL = miSQL & Trim(DBLet(miRsAux!codtelem, "T")) & "#"
                    'Nombre marca
                    miSQL = miSQL & Mid(DBLet(miRsAux!nomprovhabitual, "T"), 1, 14) & "#"
                    
                    
                    miSQL = miSQL & miRsAux!NomArtic & "#"
                    miSQL = miSQL & AlmaGrupoTextoUds & "#"
                    miSQL = miSQL & miRsAux!udscompra & "#"
                    miSQL = miSQL & Abs(miRsAux!Importe) '& "#"  'el precio siempre en positivo
                    
                    Print #NF1, miSQL
                End If
           End If
           
           
           If NF2 > 0 Then
            'STOCK
                miSQL = DBLet(miRsAux!cifproveedor, "T")
                If miSQL = "S" Then
                    miSQL = miRsAux!codArtic & "#"
                    miSQL = miSQL & miRsAux!referprov & "#"
                    miSQL = miSQL & Trim(DBLet(miRsAux!codtelem, "T")) & "#"
                    miSQL = miSQL & Mid(miRsAux!nomprovhabitual, 1, 14) & "#"
                    miSQL = miSQL & miRsAux!NomArtic & "#"
                    miSQL = miSQL & miRsAux!stock & "#"
                    miSQL = miSQL & AlmaGrupoTextoUds & "#"
                    'Cantidad de precio
                    miSQL = miSQL & "1" & "#"  'nosotros siempre precio unitario
                    'precio
                    
                    If IsNull(miRsAux!Importe) Then
                        miSQL = miSQL & DBLet(miRsAux!precioUC, "N") & "#"   'por si se creo a mano
                    Else
                        miSQL = miSQL & miRsAux!Importe & "#"
                    End If
                    miSQL = miSQL & miRsAux!cifprovhabitual & "#"
                    miSQL = miSQL & miRsAux!nomprovhabitual '& "#"
                    
                    Print #NF2, miSQL
                End If
           End If
           
           
           miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        If NF1 > 0 Then Close #NF1
        If NF2 > 0 Then Close #NF2

    DatosGeneradosExportacionAlmagrupo = True
eDatosGeneradosExportacionAlmagrupo:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function


Private Function FijaNombreFichero(ByRef mes, ByRef Anyo, Compras As Boolean, Asociado As String)
    FijaNombreFichero = App.Path & "\Temp\ALMA" & Asociado
    If Compras Then
        FijaNombreFichero = FijaNombreFichero & "COM"
    Else
        FijaNombreFichero = FijaNombreFichero & "STO"
    End If
    FijaNombreFichero = FijaNombreFichero & Format(Now, "yyyymmdd") & ".txt"

End Function


Private Function AlmaGrupoTextoUds() As String
    'Codigos de familia superirores a 20 son UNIDADES
    If miRsAux!codunida > 20 Then
        AlmaGrupoTextoUds = "UN"
    Else
        'Si no pone nada --> UNIDAD
        If DBLet(miRsAux!nomunbre, "T") = "" Then
            AlmaGrupoTextoUds = "UN"
        Else
            AlmaGrupoTextoUds = Mid(miRsAux!nomunbre & "  ", 1, 2)
        End If
    End If
End Function




Public Sub ExportarDatosFTP(ByRef LB As Label)
   
    If Not CrearFicheroLotesFTP Then Exit Sub
    'sql tendra la IP
    Shell App.Path & "\Ariftp.bat", vbMinimizedFocus
    If Not LB Is Nothing Then
        LB.Caption = "Envio FTP"
        LB.Refresh
    End If
    Espera 2
    MataLanzaFTP
End Sub

Private Function CrearFicheroLotesFTP() As Boolean
Dim Nf As Integer

    On Error GoTo eCrearFicheroLotesFTP
    CrearFicheroLotesFTP = False
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from salmagrupoparam", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    miSQL = miRsAux!ip
    'NO PUEDE SE EOF
    Nf = FreeFile
    Open App.Path & "\ftp.dat" For Output As #Nf
    Print #Nf, miRsAux!Usuario
    Print #Nf, miRsAux!Clave
    miRsAux.Close
    Print #Nf, "lcd """ & App.Path & "\temp"""
    Print #Nf, "ascii"
    
    Print #Nf, "mput *.txt"
    Print #Nf, "close"
    Print #Nf, "bye"
   
    Close #Nf
    
    Nf = FreeFile
    Open App.Path & "\Ariftp.bat" For Output As #Nf
    Print #Nf, "cd """ & App.Path & """"
    Print #Nf, "cls"
    Print #Nf, "FTP -i -s:ftp.dat " & miSQL & ""
   ' Print #Nf, ""
    'Print #Nf, "pause "
   ' Print #Nf, "pause 0"
    Close #Nf
    CrearFicheroLotesFTP = True
eCrearFicheroLotesFTP:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
End Function



Private Sub MataLanzaFTP()
    On Error Resume Next
    Kill App.Path & "\Ariftp.bat"
    Kill App.Path & "\ftp.dat"
    Err.Clear
End Sub








'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'
'   GENERACION DE LOS DATOS
'
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

Public Function GeneraDatosAlmagrupo(QueOpcion As Byte, ByRef LB As Label) As Boolean
Dim Nf As Integer
Dim Cad As String
Dim F As Date
Dim Fechas As Collection
Dim Meses As Byte
Dim J As Integer

    
    conn.Execute "DELETE FROM salmagrupo"   'Borramos los datos
    If Not CarpetaLogAlmagrupo Then Exit Function

    'Opciones.  QUEOPCION
    ' 1- Proceso diario  (consumo + stocks de 2 meses)
    ' 2- Proceso bianual (solo consumos)
    
    
    'Preparamos el fichero LOG
    Nf = FreeFile
    Open App.Path & "\Logftp\" & Format(Now, "yyyy_mm_dd_hhnnss") & ".LOG" For Output As #Nf
    Print #Nf, "Incio proceso: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf
    
    If QueOpcion = 1 Then
        Meses = 3
    Else
        Meses = 24
    End If
    
    Set Fechas = New Collection
    F = Now   'desde el dia anterior a la fecha proceso PREGUNTAR
    
   
    
    For J = 1 To Meses
        Cad = Format(F, "dd/mm/yyyy")
        F = CDate("01/" & Format(F, "mm/yyyy"))
        Cad = Format(F, "dd/mm/yyyy") & "|" & Cad & "|"
        F = DateAdd("d", -1, F)
        Fechas.Add Cad
    Next
    
    'Ya tengo todos los meses a presentar
    For J = 1 To Fechas.Count
        Cad = Fechas.Item(J)
        
        PonerLabel LB, Replace(Cad, "|", "  ")
        If (J Mod 3) = 2 Then DoEvents
    
    
        If QueOpcion = 1 Then
            'Solo para el envio diario, solo la primera vez
            If J = 1 Then
                Print #Nf, "Genera Stocks: " & Replace(Cad, "|", "  ")
                Print #Nf, String(40, "-")

                If Not GeneraDatosAlmagrupoStocks(QueOpcion, Nf, LB) Then Exit For
            End If
        End If
        
        
        Print #Nf, "Consumos periodo: " & Replace(Cad, "|", "  ")
        Print #Nf, String(40, "-")
        Cad = RecuperaValor(Cad, 1)
        F = CDate(Cad)
        Cad = Fechas.Item(J)
        If Not GeneraDatosAlmagrupoMesConsumos(F, CDate(RecuperaValor(Cad, 2)), QueOpcion, Nf, LB) Then Exit For
        
    Next
    
    Cad = "Fin proceso: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    If J <= Fechas.Count Then
        Print #Nf, Cad & " con ERRORES"
    Else
        'Vamos a ver totales
        Print #Nf, Cad
        Cad = DevuelveDesdeBD(conAri, "count(*)", "salmagrupo", "cifproveedor = 'S' AND 1", "1")
        If Cad = "" Then Cad = "0"
        Print #Nf, "Total stocks: " & Cad
        
        Cad = DevuelveDesdeBD(conAri, "count(*)", "salmagrupo", "cifproveedor <> 'S' AND 1", "1")
        If Cad = "" Then Cad = "0"
        Print #Nf, "Total consumos: " & Cad
        
        
    End If
    
    
    
    
    Close Nf
End Function

Private Sub PonerLabel(ByRef L As Label, ByRef T As String)
    If Not L Is Nothing Then
        L.Caption = T
        L.Refresh
    End If
End Sub

Private Function GeneraDatosAlmagrupoMesConsumos(FIni As Date, Ffin As Date, LaOpcion As Byte, Fichero As Integer, ByRef L As Label) As Boolean
Dim I As Currency
Dim Aux2 As String
Dim Fin As Boolean
Dim Insert As String
Dim Contador As Long
    On Error GoTo eGeneraDatosAlmagrupo

    GeneraDatosAlmagrupoMesConsumos = False
    
    Set miRsAux = New ADODB.Recordset

    
    
    'Es un proceso que se hara todos los dias.
    'Por ello ahora veremos el consumo del mes a fecha de hoy
    'Cargaremos en RT los datos del mes -año ordenado por nifprove
    'Cifprovee='S' es el stcok
    
    
    
    'cifproveedor,mes,anyo,codartic,udscompra,importe,stock,cifprovhabitual,
    miSQL = "Select nifprove elprove,month(fecfactu),year(fecfactu),slifpc.codartic,sum(cantidad) cantidad,"
    miSQL = miSQL & "sum(importel) elimporte,0,sartic.codprove habitual ,nomprove"
    miSQL = miSQL & " from slifpc,sartic,sfamia,sprove WHERE slifpc.codartic = sartic.codartic "
    miSQL = miSQL & " AND sartic.codfamia =sfamia.codfamia AND sprove.codprove=slifpc.codprove "
    
    miSQL = miSQL & " AND comunica=1 AND artvario =0 "       'Familias que se comunican
    
    'NIO hace falta que tengan codigo telematel
    'miSQL = miSQL & " AND codtelem<>''"     'que tenga codigo telematel
    'En el tag tengo la fecha a declarar
    

    miSQL = miSQL & " AND fecfactu >='" & Format(FIni, FormatoFecha)
    miSQL = miSQL & "' AND fecfactu <='" & Format(Ffin, FormatoFecha) & "'"
    miSQL = miSQL & " group by slifpc.codartic,nifprove"
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Insert = ""
    Contador = 0
    While Not miRsAux.EOF
            'nO ACEPTA multiSELECT
            'Esta ordenado codartic,codproveventa
            
           
           
                'insert salmagrupo (cifproveedor,mes,anyo,codartic,udscompra,importe,stock,cifprovhabitual,
                'miSQL = "INSERT INTO salmagrupo (cifprovhabitual, nomprovhabitual,cifproveedor,nomproveedor,mes,anyo,codartic,udscompra,importe,stock) " & miSQL
                Aux2 = "nomprove"
                miSQL = DevuelveDesdeBD(conAri, "nifprove", "sprove", "codprove", miRsAux!habitual, "N", Aux2)
    
                If miSQL = "" Then
                    'ERROR en el NIF
                    Err.Raise 513, "Nif erroneo: " & miRsAux!habitual
                    
                    
                    
                    
                Else
                    Contador = Contador + 1
                    Aux2 = DBSet(Aux2, "T")
                    Aux2 = DBSet(miSQL, "T") & "," & Aux2
                    
                    'cifprovhabitual, nomprovhabitual,cifproveedor,nomproveedor
                    miSQL = ", (" & Aux2 & "," & DBSet(miRsAux!elprove, "T") & "," & DBSet(miRsAux!nomprove, "T")
                    'mes,anyo,codartic,udscompra,importe,stock
                    miSQL = miSQL & "," & Month(FIni) & "," & Year(FIni) & "," & DBSet(miRsAux!codArtic, "T")
                    miSQL = miSQL & "," & DBSet(miRsAux!Cantidad, "N") & "," & DBSet(miRsAux!elimporte, "N") & ",0)"
                    Insert = Insert & miSQL
                    
                    If Len(Insert) > 10000 Then
                        
                        Insert = " VALUES " & Mid(Insert, 2) 'quito el primer punto
                        Insert = "INSERT INTO salmagrupo (cifprovhabitual, nomprovhabitual,cifproveedor,nomproveedor,mes,anyo,codartic,udscompra,importe,stock) " & Insert
                        conn.Execute Insert
                        Insert = ""
                    End If
                End If

            
            
            miRsAux.MoveNext
    Wend
    miRsAux.Close

    
    If Insert <> "" Then
        Insert = " VALUES " & Mid(Insert, 2) 'quito el primer punto
        Insert = "INSERT INTO salmagrupo (cifprovhabitual, nomprovhabitual,cifproveedor,nomproveedor,mes,anyo,codartic,udscompra,importe,stock) " & Insert
        conn.Execute Insert
        Insert = ""
    End If
    
    GeneraDatosAlmagrupoMesConsumos = True
    Aux2 = "Total registros: " & Contador
    Print #Fichero, Aux2
    
eGeneraDatosAlmagrupo:
    If Err.Number <> 0 Then
        Aux2 = Err.Description
        Print #Fichero, Aux2
        'MuestraError Err.Number, "AVISE SOPORTE TECNICO"
    End If
    Set miRsAux = Nothing

End Function



Private Function GeneraDatosAlmagrupoStocks(LaOpcion As Byte, Fichero As Integer, ByRef L As Label) As Boolean
Dim I As Currency
Dim Aux2 As String
Dim Fin As Boolean
Dim Insert As String


    On Error GoTo eGeneraDatosAlmagrupo

    GeneraDatosAlmagrupoStocks = False
    
    Set miRsAux = New ADODB.Recordset

    
    
    'Ahora va lo bueno, el stock del mes
    PonerLabel L, "STOCK"

    
    
    miSQL = "select salmac.codartic,canstock,PorcenComunica,nomartic,nifprove,nomprove,rotacion,preciouc from salmac ,sartic,sfamia,sprove"
    miSQL = miSQL & " where salmac.codartic=sartic.codartic and sprove.codprove=sartic.codprove"
    miSQL = miSQL & " and sartic.codfamia=sfamia.codfamia"
    miSQL = miSQL & " and codalmac=1 AND artvario =0 "
    'No hace falt que tengan codigo telemalte
    'miSQL = miSQL & " and codtelem<>'' "
    miSQL = miSQL & " and comunica=1 and canstock>0"
    miSQL = miSQL & " ORDER BY salmac.codartic"
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    'insert salmagrupo (cifproveedor,mes,anyo,codartic,udscompra,importe,stock,cifprovhabitual
    miSQL = ""
    While Not miRsAux.EOF
        PonerLabel L, miRsAux!NomArtic
        
        If DBLet(miRsAux!Rotacion, "N") = 0 Then
            'Los que nos son de rotacion declaro el %
            
            
            I = (DBLet(miRsAux!PorcenComunica, "N") / 100)
        Else
            'De rotacion NO declaro NADA
            'i = (DBLet(miRsAux!PorcenComunica, "N") / 100)
            I = 0
        End If
        I = miRsAux!CanStock * I
        I = CInt(I)
        
        
        'Solo declaro stocks positivos
        If I > 0 Then
                    
                NumRegElim = NumRegElim + 1
                    
                'Nuevo en salmagrupo
                'insert salmagrupo (cifproveedor,mes,anyo,codartic,udscompra,importe,stock,cifprovhabitual,ConcomprasPeriodo,
                'miSQL = miSQL & ", ('S'," & Month(FIni) & "," & Year(FIni) & "," & DBSet(miRsAux!codArtic, "T")
                miSQL = miSQL & ", ('S',1,1," & DBSet(miRsAux!codArtic, "T")
                miSQL = miSQL & ",NULL," & DBSet(miRsAux!precioUC, "N") & "," & DBSet(I, "N") & "," & DBSet(miRsAux!nifProve, "T") & "," & DBSet(miRsAux!nomprove, "T") & ")"
                
                If NumRegElim > 50 Then
                    miSQL = Mid(miSQL, 2)
                    miSQL = "insert INTO salmagrupo (cifproveedor,mes,anyo,codartic,udscompra,importe,stock,cifprovhabitual,nomprovhabitual) VALUES " & miSQL
                    conn.Execute miSQL
                    NumRegElim = 0
                    miSQL = ""
                End If
                

            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close

     If NumRegElim > 0 Then
        miSQL = Mid(miSQL, 2)
        miSQL = "insert INTO salmagrupo (cifproveedor,mes,anyo,codartic,udscompra,importe,stock,cifprovhabitual,nomprovhabitual) VALUES " & miSQL
        conn.Execute miSQL
    End If
    GeneraDatosAlmagrupoStocks = True
    
eGeneraDatosAlmagrupo:
    If Err.Number <> 0 Then
        Aux2 = Err.Description
        Print #Fichero, Aux2
        'MuestraError Err.Number, "AVISE SOPORTE TECNICO"
    End If
    Set miRsAux = Nothing

End Function






Private Function CarpetaLogAlmagrupo() As Boolean
On Error GoTo eCarpetaLogAlmagrupo
    
    CarpetaLogAlmagrupo = False
    If Dir(App.Path & "\Logftp", vbDirectory) = "" Then MkDir App.Path & "\Logftp"
    CarpetaLogAlmagrupo = True
    Exit Function
eCarpetaLogAlmagrupo:
    MuestraError Err.Number, Err.Description
    
End Function



