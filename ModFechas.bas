Attribute VB_Name = "ModFechas"
Option Explicit

Dim ComprobarFecha

'=== DAVID (estaban en Modulo:bus) (NO LA USO!!!)
Public Function DiasMes(mes As Byte, Anyo As Integer) As Integer
    Select Case mes
    Case 2
        If (Anyo Mod 4) = 0 Then
            DiasMes = 29
        Else
            DiasMes = 28
        End If
    Case 1, 3, 5, 7, 8, 10, 12
        DiasMes = 31
    Case Else
        DiasMes = 30
    End Select
End Function


'=== DAVID (estaban en Modulo:bus)
'Public Function EsFechaOK(ByRef T As TextBox) As Boolean
''Dim cad As String
''
''    cad = T.Text
''    If InStr(1, cad, "/") = 0 Then
''        If Len(T.Text) = 8 Then
''            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
''        Else
''            If Len(T.Text) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
''        End If
''    End If
''
''    If IsDate(cad) Then
''        EsFechaOK = True
''        T.Text = Format(cad, "dd/mm/yyyy")
''    Else
''        EsFechaOK = False
''    End If
''EsFechaOK = EsFechaOKString
'End Function

'=== DAVID (estaban en Modulo:bus, antes era ESFechaOKString)
Public Function EsFechaOK(T As String) As Boolean
Dim Cad As String
Dim mes As String, dia As String
    
    Cad = T
    If InStr(1, Cad, "/") = 0 Then
       'debe ser una cadena tipo:020105 y la convertimos a 02/01/05
       If Not IsNumeric(Cad) Then
            EsFechaOK = False
            Exit Function
       End If
        
      '==== Anade: Laura 04/02/2005 =============
        If Len(Cad) < 6 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el dia es correcto, valores entre 1-31
        dia = Mid(Cad, 1, 2)
        If dia < 1 Or dia > 31 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el mes es correcto, valores entre 1-12
        mes = Mid(Cad, 3, 2)
        If mes < 1 Or mes > 12 Then
            EsFechaOK = False
            Exit Function
        End If
      '============================================
        
        If Len(T) = 8 Then
            Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        Else
            If Len(T) = 6 Then Cad = Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5)
        End If
    Else
        dia = Mid(Cad, 1, 2)
        mes = Mid(Cad, 4, 2)
    End If
    
    If IsDate(Cad) Then
        EsFechaOK = True
        T = Format(Cad, "dd/mm/yyyy")
      '==== A�ade: Laura 08/02/2005
        If Month(T) <> Val(mes) Then EsFechaOK = False
        If Day(T) <> Val(dia) Then EsFechaOK = False
      '====
    Else
        EsFechaOK = False
    End If
End Function


'=== DAVID (estaba en Modulo:bus)
Public Function EsHoraOK(T As String) As Boolean
Dim Cad As String
    
    Cad = T
    If InStr(1, Cad, ":") = 0 Then
        Select Case Len(T)
            Case 8
                Cad = Mid(Cad, 1, 2) & ":" & Mid(Cad, 3, 2) & ":" & Mid(Cad, 5)
            Case 6
                Cad = Mid(Cad, 1, 2) & ":" & Mid(Cad, 3, 2) & ":" & Mid(Cad, 5)
            Case 4
                Cad = Mid(Cad, 1, 2) & ":" & Mid(Cad, 3, 2) & ":00"
        End Select
    End If
    
    If IsDate(Cad) Then
        EsHoraOK = True
        T = Format(Cad, "hh:mm:ss")
    Else
        EsHoraOK = False
    End If
End Function


'==== LAURA
Public Sub PonerFormatoFecha(ByRef T As TextBox)
Dim Cad As String

    Cad = T.Text
    If Cad <> "" Then
        If Not EsFechaOK(Cad) Then
            MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
            Cad = "mal"
        End If
        If Cad <> "" And Cad <> "mal" Then
            T.Text = Cad
        Else
            T.Text = ""
            PonerFoco T
        End If
    End If
End Sub


'---- David Mao 2016
Public Function EsFechaHoraOK(T As String) As Boolean
Dim EspacioEnBlanco As Integer
Dim LaFecha As String
Dim LaHora As String
Dim Mensaje As String

    'Devolvera OK si:
    ' dd/mm/yyyy hh:nn:ss
    
    'Es decir. Hay un espacio en blanco
    ' Lo de antes del espacio en blanco es una fecha
    ' Lo de despues es una hora
    '
    EsFechaHoraOK = False
    Mensaje = ""
    EspacioEnBlanco = InStr(1, T, " ")
    If EspacioEnBlanco = 0 Then
        Mensaje = "Falta separacion fecha-hora(Espacio)"
    Else
        LaFecha = Trim(Mid(T, 1, EspacioEnBlanco))
        LaHora = Trim(Mid(T, EspacioEnBlanco))
        
        If Not EsFechaOK(LaFecha) Then
            Mensaje = "Fecha incorrecta: " & LaFecha
        Else
            If Not EsHoraOK(LaHora) Then Mensaje = "Fecha incorrecta: " & LaFecha
        End If
                
    End If
    If Mensaje <> "" Then
        MsgBox Mensaje, vbExclamation
    Else
        T = LaFecha & " " & LaHora
        EsFechaHoraOK = True
    End If
End Function


'==== LAURA
Public Sub PonerFormatoHora(ByRef T As TextBox)
Dim Cad As String

        Cad = T.Text
        If Cad <> "" Then
            If Not EsHoraOK(Cad) Then
                MsgBox "Hora incorrecta. (hh:mm:ss)", vbExclamation
                Cad = "mal"
            End If
            If Cad <> "" And Cad <> "mal" Then
                T.Text = Cad
            Else
                T.Text = ""
                PonerFoco T
            End If
        End If
End Sub


'==== LAURA
Public Function EsFechaPosterior(FIni As String, Ffin As String, MError As Boolean, Optional Men As String) As Boolean
'Comprueba que la Fecha Fin es posterior a la Fecha de Inicio
'Si se pasa un cadena Men, se muestra esta como Mensaje de Error
On Error Resume Next

    EsFechaPosterior = True
    If Trim(FIni) <> "" And Trim(Ffin) <> "" Then
        If CDate(FIni) >= CDate(Ffin) Then
            EsFechaPosterior = False
            If MError Then
                If Men <> "" Then
                    MsgBox Men, vbInformation
                Else
                    MsgBox "La Fecha Fin debe ser posterior a la Fecha Inicio", vbInformation
                End If
            End If
        Else
            EsFechaPosterior = True
        End If
    End If
End Function


'==== LAURA
'==== Fec. ult. modif.: 20/06/2008
Public Function EsFechaIgualPosterior(FIni As String, Ffin As String, MError As Boolean, Optional Men As String) As Boolean
'Comprueba que la Fecha Fin es igual o posterior a la Fecha de Inicio
'Si se pasa un cadena Men, se muestra esta como Mensaje de Error
'(IN) -> FIni: fecha inicio
'(IN) -> FFin: fecha fin
'(IN) -> MError: mostrar mensaje de error si/no
'(IN) -> Men: cadena mensaje de error
'(OUT) -> true: FFin >= Fini

    On Error GoTo ErrFec

'    EsFechaIgualPosterior = True
    
    If Trim(FIni) <> "" And Trim(Ffin) <> "" Then
        If CDate(FIni) > CDate(Ffin) Then
            EsFechaIgualPosterior = False
            
            If MError Then 'mostrar error
                If Men <> "" Then
                    'mostrar mensaje especifico q pasamos como parametro
                    MsgBox Men, vbInformation
                Else
                    'mostrar mensaje general
                    MsgBox "La Fecha Fin debe ser igual o posterior a la Fecha Inicio", vbInformation
                End If
            End If
        Else
            EsFechaIgualPosterior = True
        End If
    Else
        EsFechaIgualPosterior = True
    End If
    
    Exit Function
    
ErrFec:
    MuestraError Err.Number, "", Err.Description
End Function



'Marzo 2015
' Duplico la funcion para no tocarlo en todos los sitios
'
Public Function EsFechaIgualPosteriorDavid(F1 As String, F2 As String) As Boolean

    On Error GoTo ErrFec

   'f1 no puede ser ""
   
    If F2 = "" Then
        EsFechaIgualPosteriorDavid = False
    Else
        If CDate(F1) >= CDate(F2) Then
            EsFechaIgualPosteriorDavid = True
            
        Else
            EsFechaIgualPosteriorDavid = False
        End If
    End If
    
    Exit Function
    
ErrFec:
    MuestraError Err.Number, "", Err.Description
End Function



'==== LAURA
Public Function EntreFechas(FIni As String, FechaComp As String, Ffin As String) As Boolean
Dim b As Boolean
    b = False
    If FIni <> "" And Ffin <> "" Then
        If EsFechaIgualPosterior(FIni, FechaComp, False) And EsFechaIgualPosterior(FechaComp, Ffin, False) Then
            b = True
        End If
    ElseIf FIni = "" And Ffin <> "" Then
        If EsFechaIgualPosterior(FechaComp, Ffin, False) Then
            b = True
        End If
    ElseIf FIni <> "" And Ffin = "" Then
        If EsFechaIgualPosterior(FIni, FechaComp, False) Then
            b = True
        End If
    End If
    EntreFechas = b
End Function

'==== LAURA
Public Function CalculaSemana(Fecha As Date) As Integer
  'Antes
    'CalculaSemana = DatePart("ww", Fecha)
    'Ahora. Copiado de ariges.  11-Ene-2011
    'Enero 2013.   Primera semana es aquella que tiene un jueves.  vbFirstFourDays
    'CalculaSemana = DatePart("ww", Fecha, vbMonday, vbFirstFullWeek)
    CalculaSemana = DatePart("ww", Fecha, vbMonday, vbFirstFourDays)
    If CalculaSemana >= 52 Then
        If Month(Fecha) = 1 Then CalculaSemana = 0
    Else
        If CalculaSemana = 1 Then
            If Month(Fecha) = 12 Then CalculaSemana = 53
        End If
    End If
End Function




'==== LAURA
Public Function EsMesOK(vMes As Integer) As Boolean

    If vMes >= 1 And vMes <= 12 Then
        EsMesOK = True
    Else
        EsMesOK = False
    End If
End Function
