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
Dim cad As String
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
        cad = TicketAgrupado & "fecfactu>=" & DBSet(vEmpresa.FechaIni, "F") & " AND  intconta "
        cad = DevuelveDesdeBD(conAri, "min(fecfactu)", "scafac", cad, "0")
        If cad <> "" Then
            Horas = DateDiff("d", CDate(cad), Now)
            If Horas > 1 Then Mensaje = "Cliente: " & cad & vbCrLf
        End If
        
        If Mensaje = "" Then
            cad = "fecrecep>=" & DBSet(vEmpresa.FechaIni, "F") & " AND  intconta "
            cad = DevuelveDesdeBD(conAri, "min(fecrecep)", "scafpc", cad, "0")
            If cad <> "" Then
                Horas = DateDiff("d", CDate(cad), Now)
                If Horas > 1 Then Mensaje = Mensaje & "Proveedor: " & cad & vbCrLf
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
        
        cad = "fecfactu>=" & DBSet(vParamAplic.Sii_Finicio, "F") & " AND intconta "
        cad = DevuelveDesdeBD(conAri, "min(fecfactu)", "scafac", cad, "0")
        If cad <> "" Then
            Horas = DateDiff("d", CDate(cad), Now)
            If Horas > 1 Then Mensaje = "O"
        End If
        If Mensaje = "" Then
            cad = "fecrecep>=" & DBSet(vParamAplic.Sii_Finicio, "F") & " AND intconta "
            cad = DevuelveDesdeBD(conAri, "min(fecrecep)", "scafpc", cad, "0")
            If cad <> "" Then
                Horas = DateDiff("d", CDate(cad), Now)
                If Horas > 1 Then Mensaje = Mensaje & "0"
            End If
        End If
    End If
    If Mensaje <> "" Then
        'Damos mensaje
        frmSiiAvisos.Show vbModal
    End If

    
    
    


End Sub
