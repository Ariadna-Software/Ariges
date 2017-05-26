Attribute VB_Name = "libSII"
Option Explicit



'***********************************************************************************************
' Avisos facturas sin contabilizar


Public Sub ComprobarFechaContabilizadas()
    
    'Solo para nuevas contabilidad
    If Not vParamAplic.ContabilidadNueva Then Exit Sub
    
    'Solo usuarios con nivel 0-1
    If vUsu.Nivel > 1 Then Exit Sub
    
    ComprobarTablaFechas
    
    
    ' Veremos si ya ha dado el mensaje, si se tiene dar , y si es asi, darlo
    DarAvisoContabilizadas
    
    
End Sub
 
Private Sub ComprobarTablaFechas()
    On Error Resume Next
    
    conn.Execute "Select * from usuarios.wavisoscontabilizacion where false"
    If Err.Number <> 0 Then
        Err.Clear
        CrearTableTablasFechas
    End If
    
    
    
End Sub

Private Sub CrearTableTablasFechas()
Dim cad As String
    
    cad = "CREATE TABLE usuarios.wavisoscontabilizacion ("
    cad = cad & "login varchar(20) NOT NULL DEFAULT '0',"
    cad = cad & "aplicacion tinyint(4) NOT NULL DEFAULT '0',"
    cad = cad & "codempre smallint(1) unsigned NOT NULL DEFAULT '0',"
    cad = cad & "ultaviso datetime DEFAULT NULL,"
    cad = cad & "PRIMARY KEY (`login`,`aplicacion`,`codempre`)"
    cad = cad & ") ENGINE=MyISAM ;"
    
    
    ejecutar cad, True
End Sub


Private Sub DarAvisoContabilizadas()
Dim cad As String
Dim FecUltAviso As Date
Dim Horas As Long
Dim VerSiDamosAviso As Boolean
Dim Mensaje As String
    '       ariges
    cad = "aplicacion = 1 AND codempre = " & vEmpresa.codempre & " AND login "
    cad = DevuelveDesdeBD(conAri, "ultaviso", "usuarios.wavisoscontabilizacion", cad, vUsu.Login, "T")
    If cad = "" Then
       FecUltAviso = DateAdd("yyyy", -1, Now)
    Else
        FecUltAviso = CDate(cad)
    End If
    
    VerSiDamosAviso = False
    If Year(FecUltAviso) - Year(Now) > 1 Then
        VerSiDamosAviso = True
    Else
        'Si hay mas de un dia de diferencia
        Horas = DateDiff("d", FecUltAviso, Now)
        If Horas > 0 Then
            VerSiDamosAviso = True
        Else
            
            Horas = DateDiff("h", FecUltAviso, Now)
            If Horas > 20 Then VerSiDamosAviso = True
        End If
    End If
    
    If Not VerSiDamosAviso Then Exit Sub
    Mensaje = ""
    
    cad = "fecfactu>=" & DBSet(vEmpresa.FechaIni, "F") & " AND  intconta "
    cad = DevuelveDesdeBD(conAri, "min(fecfactu)", "scafac", cad, "0")
    If cad <> "" Then
        Horas = DateDiff("d", CDate(cad), Now)
        If Horas > 7 Then Mensaje = "Cliente: " & cad & vbCrLf
    End If
    
    cad = "fecrecep>=" & DBSet(vEmpresa.FechaIni, "F") & " AND  intconta "
    cad = DevuelveDesdeBD(conAri, "min(fecrecep)", "scafpc", cad, "0")
    If cad <> "" Then
        Horas = DateDiff("d", CDate(cad), Now)
        If Horas > 7 Then Mensaje = Mensaje & "Proveedor: " & cad & vbCrLf
    End If
        
    If Mensaje <> "" Then
        'Damos mensaje
        Mensaje = "Facturas pendientes de contabilizar." & vbCrLf & vbCrLf & Mensaje
        MsgBox Mensaje, vbInformation
    End If
    
    
    cad = "replace into usuarios.wavisoscontabilizacion(`login`,`aplicacion`,`codempre`,`ultaviso`) values ("
    cad = cad & DBSet(vUsu.Login, "T") & ",'1'," & vEmpresa.codempre & "," & DBSet(Now, "FH") & ")"
    ejecutar cad, False
End Sub
