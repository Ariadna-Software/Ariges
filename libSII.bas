Attribute VB_Name = "libSII"
Option Explicit



'***********************************************************************************************
' Avisos facturas sin contabilizar


Public Sub ComprobarFechaContabilizadas()
    
    'Solo para nuevas contabilidad
    If Not vParamAplic.ContabilidadNueva Then Exit Sub
    
    'Solo usuarios con nivel 0-1
    If vUsu.Nivel > 1 Then Exit Sub
    
   
    
    
    ' Veremos si ya ha dado el mensaje, si se tiene dar , y si es asi, darlo
    DarAvisoContabilizadas
    
    
End Sub
 

Private Sub DarAvisoContabilizadas()
Dim Cad As String
Dim FecUltAviso As Date
Dim Horas As Long
Dim VerSiDamosAviso As Boolean
Dim Mensaje As String
Dim TicketAgrupado As String
    

    Mensaje = ""
    
    
    TicketAgrupado = DevuelveDesdeBD(conAri, "contador", "stipom", "codtipom", "FTG", "T")
    If TicketAgrupado <> "" Then
        If Val(TicketAgrupado) > 0 Then
            'Tiene tickets AGRUPADOS. No damos mensajes de tickets pendientes de contabilizar
            TicketAgrupado = " codtipom <> 'FTI' AND "
        Else
            TicketAgrupado = ""
        End If
    End If
    
    If Not vParamAplic.SII_Tiene Then
        
        Cad = TicketAgrupado & "  codtipom <> 'FAI' AND fecfactu>=" & DBSet(vEmpresa.FechaIni, "F") & " AND codtipom<>'FAZ' AND  intconta "
        Cad = DevuelveDesdeBD(conAri, "min(fecfactu)", "scafac", Cad, "0")
        If Cad <> "" Then
            Horas = DateDiff("d", CDate(Cad), Now)
            If Horas > 1 Then Mensaje = "Cliente: " & Cad & vbCrLf
        End If
        
        If Mensaje = "" Then
            Cad = "fecrecep>=" & DBSet(vEmpresa.FechaIni, "F") & " AND  intconta "
            Cad = DevuelveDesdeBD(conAri, "min(fecrecep)", "scafpc", Cad, "0")
            If Cad <> "" Then
                Horas = DateDiff("d", CDate(Cad), Now)
                If Horas > 1 Then Mensaje = Mensaje & "Proveedor: " & Cad & vbCrLf
            End If
        End If
                
        If Mensaje <> "" Then
            'Damos mensaje
           ' Mensaje = "Facturas pendientes de contabilizar." & vbCrLf & vbCrLf & Mensaje
           ' MsgBox Mensaje, vbInformation
        End If
            
    Else
        '****************************  Tiene SII
        'Veremos un poco mas el mensaje de facturas contabilizadas
        Cad = ""
        If TicketAgrupado <> "" Then Cad = TicketAgrupado
        Cad = Cad & "fecfactu>=" & DBSet(vParamAplic.Sii_Finicio, "F") & " AND intconta "
        Cad = DevuelveDesdeBD(conAri, "min(fecfactu)", "scafac", Cad, "0")
        If Cad <> "" Then
            Horas = DateDiff("d", CDate(Cad), Now)
            If Horas > 1 Then Mensaje = "O"
        End If
        If Mensaje = "" Then
            
            Cad = "fecrecep>=" & DBSet(vParamAplic.Sii_Finicio, "F") & " AND intconta "
            'Solo aviso intconta.  Las facturas suben al sii con fecha integracion contable
            Cad = " intconta "
            Cad = DevuelveDesdeBD(conAri, "min(fecrecep)", "scafpc", Cad, "0")
            If Cad <> "" Then
                Horas = DateDiff("d", CDate(Cad), Now)
                If Horas > 1 Then Mensaje = Mensaje & "0"
            End If
        End If
    End If
    
    
    
    
    If Mensaje <> "" Then
        'Damos mensaje
        If vUsu.Nivel2 = 2 Then
            'No hacemos nada
            
        Else

             frmSiiAvisos.Show vbModal
            
        End If
    End If


    If vParamAplic.NumeroInstalacion = vbHerbelca Then
        If vParamAplic.DiasCaducidadPuntos > 0 Then
            Mensaje = PonerTrabajadorConectado("")
            If Mensaje <> "" Then
                Mensaje = " codtraba =" & Mensaje & " and fecaviso>20010101 and datediff(now(),fecaviso)>15 AND 1"
                Mensaje = DevuelveDesdeBD(conAri, "codtraba", "straba", Mensaje, "1")
                If Mensaje <> "" Then
                
                    Mensaje = "UPDATE straba set fecaviso=" & DBSet(Now, "F") & " WHERE codtraba = " & Mensaje
                    ejecutar Mensaje, False
                
                    Mensaje = "¿Desea realizar el proceso de caducar puntos?"
                    If MsgBox(Mensaje, vbQuestion + vbYesNo) = vbYes Then Mensaje = ""
                    
                    
                    
                    If Mensaje = "" Then
                        frmMensajes.cadWhere = ""
                        frmMensajes.OpcionMensaje = 31
                        frmMensajes.Show vbModal
                    End If
                    
                End If
            End If
        End If
    End If
    


End Sub




'FechaPresentacion:  Normalmente sera now()
Public Function UltimaFechaCorrectaSII(DiasAVisoSII As Integer, FechaPresentacion As Date) As Date
Dim DiaSemanaPresen As Integer
Dim DiaSemanaUltimoDiaPresentar As Integer
Dim F As Date

Dim Resta As Integer

    If DiasAVisoSII > 5 Then
        
        UltimaFechaCorrectaSII = DateAdd("d", -DiasAVisoSII, FechaPresentacion)
        

    Else
        DiaSemanaPresen = Weekday(FechaPresentacion, vbMonday)
       
                
                
        If DiaSemanaPresen >= 6 Then
            'Si presento el sabado o el domingo tengo mas dias
            If DiaSemanaPresen = 6 Then
                Resta = DiasAVisoSII
            Else
                Resta = DiasAVisoSII + 1
            End If
        Else
            F = DateAdd("d", -DiasAVisoSII, FechaPresentacion)
            DiaSemanaUltimoDiaPresentar = Weekday(F, vbMonday)
            
            If DiaSemanaUltimoDiaPresentar > DiaSemanaPresen Then
                Resta = DiasAVisoSII + 2
            
            Else
                'Directamente la resta son 4
                Resta = DiasAVisoSII
            End If
        End If
        UltimaFechaCorrectaSII = DateAdd("d", -Resta, FechaPresentacion)
    End If

    UltimaFechaCorrectaSII = Format(UltimaFechaCorrectaSII, "dd/mm/yyyy")

End Function
'************** RUTINA COPMPROBACION
'   Dim fin As Boolean
'    fin = False
'
'    Dim F As Date
'    Dim F2 As Date
'    Dim Cad As String
'    Dim c2 As String
'    Dim I As Integer
'
'    Do
'        Cad = ""
'        For I = 1 To 28
'            F = CDate(Format(I, "00") & "/02/2018")
'
'            F2 = UltimaFechaCorrectaSII(3, F)
'
'
'            c2 = F & "  " & Weekday(F, vbMonday) & " --> "
'            c2 = c2 & F2 & "  " & Weekday(F2, vbMonday)
'            Cad = Cad & c2 & vbCrLf
'        Next
'
'        MsgBox Cad, vbExclamation
'
'
'
'
'    Loop Until fin
