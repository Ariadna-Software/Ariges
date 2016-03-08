Attribute VB_Name = "ExpresionBusqueda"
Option Explicit

Public Function SeparaCampoBusqueda(Tipo As String, campo As String, CADENA As String, ByRef DevSQL As String, Optional paraRPT) As Byte
Dim Cad As String
Dim Aux As String
Dim CH As String
Dim Fin As Boolean
Dim I, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
Cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    '==== Laura: 11/07/05
    If IsNumeric(CADENA) Then
        CADENA = CStr(ImporteFormateado(CADENA))
        CADENA = TransformaComasPuntos(CADENA)
    End If
    '====================
    I = CararacteresCorrectos(CADENA, "N")
    If I > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo numerico
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not IsNumeric(Cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = campo & " >= " & Cad & " AND " & campo & " <= " & Aux
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    I = 1
                    Cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        CH = Mid(CADENA, I, 1)
                        If CH = ">" Or CH = "<" Or CH = "=" Then
                            Cad = Cad & CH
                            Else
                                Aux = Mid(CADENA, I)
                                Fin = True
                        End If
                        I = I + 1
                        If I > Len(CADENA) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If Cad = "" Then Cad = " = "
                    DevSQL = campo & " " & Cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(CADENA, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo de fechas
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not EsFechaOK(Cad) Or Not EsFechaOK(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        
'        If Not Left(campo, 1) = "{" Then
'                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
'                Else
'                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
'                End If
        
        If paraRPT Then
            Cad = "Date(" & Year(Cad) & "," & Month(Cad) & "," & Day(Cad) & ")"
            Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
            DevSQL = campo & " >=" & Cad & " AND " & campo & " <= " & Aux
        Else
            Cad = Format(Cad, FormatoFecha)
            Aux = Format(Aux, FormatoFecha)
            'En my sql es la ' no el #
            'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
            DevSQL = campo & " >='" & Cad & "' AND " & campo & " <= '" & Aux & "'"
        End If
        '----
        'ELSE
    Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    CH = Mid(CADENA, I, 1)
                    If CH = ">" Or CH = "<" Or CH = "=" Then
                        Cad = Cad & CH
                        Else
                            Aux = Mid(CADENA, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(CADENA) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOK(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                If Not Left(campo, 1) = "{" Then
                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
                Else
                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
                End If
                If Cad = "" Then Cad = " = "
                DevSQL = campo & " " & Cad & " " & Aux
            End If
    End If
    
  Case "H"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(CADENA, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        'Intervalo de fechas
        Cad = Mid(CADENA, 1, I - 1)
        Aux = Mid(CADENA, I + 1)
        If Not EsFechaOK(Cad) Or Not EsFechaOK(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        
'        If Not Left(campo, 1) = "{" Then
'                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
'                Else
'                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
'                End If
        
        If paraRPT Then
            Cad = "Date(" & Year(Cad) & "," & Month(Cad) & "," & Day(Cad) & ")"
            Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
            DevSQL = campo & " >=" & Cad & " AND " & campo & " <= " & Aux
        Else
            Cad = Format(Cad, FormatoFecha)
            Aux = Format(Aux, FormatoFecha)
            'En my sql es la ' no el #
            'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
            DevSQL = campo & " >='" & Cad & "' AND " & campo & " <= '" & Aux & "'"
        End If
        '----
        'ELSE
    Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    CH = Mid(CADENA, I, 1)
                    If CH = ">" Or CH = "<" Or CH = "=" Then
                        Cad = Cad & CH
                        Else
                            Aux = Mid(CADENA, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(CADENA) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOK(Aux) Then
                    'Veo si es una hora
                    If EsHoraOK(Aux) Then MsgBox "Debe especificar fecha/hora", vbExclamation
                    
                    Exit Function
                    
                    
                Else
                
                
                
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If Not Left(campo, 1) = "{" Then
                        Aux = Format(Aux, FormatoFecha)
                        DevSQL = campo & " >= '" & Aux & " 00:00:00' AND " & campo & " <= '" & Aux & " 23:59:59'"
                    Else
                        Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
                    
                    
                        DevSQL = campo & " " & Cad & " " & Aux
                    End If
                End If
            End If
    End If
    
Case "T"
    '---------------- TEXTO ------------------
    I = CararacteresCorrectos(CADENA, "T")
    If I = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If CADENA = ">>" Or CADENA = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    
    'Comprobamos si es LIKE o NOT LIKE
    Cad = Mid(CADENA, 1, 2)
    If Cad = "<>" Then
        CADENA = Mid(CADENA, 3)
        If Left(campo, 1) <> "{" Then
            'No es consulta seleccion para Report.
            DevSQL = campo & " NOT LIKE '"
        Else
            'Consulta de seleccion para Crystal Report
            DevSQL = "NOT (" & campo & " LIKE """ & CADENA & """)"
        End If
    Else
        If Left(campo, 1) <> "{" Then
        'NO es para report
            DevSQL = campo & " LIKE '"
        Else  'Es para report
            I = InStr(1, CADENA, "*")
            'Poner Consulta de seleccion para Crystal Report
            If I > 0 Then
                DevSQL = campo & " LIKE """ & CADENA & """"
            Else
                DevSQL = campo & " = """ & CADENA & """"
            End If
        End If
    End If
    
    
    'Cambiamos el * por % puesto que en ADO es el caraacter para like
    I = 1
    Aux = CADENA
    If Not Left(campo, 1) = "{" Then
      'No es para report
       While I <> 0
           I = InStr(1, Aux, "*")
           If I > 0 Then
                Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
            End If
        Wend
    End If
    
    'Cambiamos el ? por la _ pue es su omonimo
    I = 1
    While I <> 0
        I = InStr(1, Aux, "?")
        If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
    Wend
    
    
    'Poner el valor de la expresion
    If Left(campo, 1) <> "{" Then
        'No es consulta seleccion para Report.
        DevSQL = DevSQL & Aux & "'"
    'Else
        'Consulta de seleccion para Crystal Report
        'DevSQL = DevSQL & CADENA & """)"
    End If
    
    '=========
    'ANTES
'    If cad = "<>" Then
'        '====David
'        'Aux = Mid(CADENA, 3)
'        'LAura
'        Aux = Mid(Aux, 3)
'        '====
'        If Left(Campo, 1) <> "{" Then
'            'Mo es consulta seleccion para Report.
'            DevSQL = Campo & " NOT LIKE '" & Aux & "'"
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " <> " & Aux & ""
'        End If
'    Else
'        If Left(Campo, 1) <> "{" Then
'            DevSQL = Campo & " LIKE '" & Aux & "'"
'        ElseIf Left(Aux, 4) = "like" Then
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " " & Aux
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " = """ & Aux & """"
'        End If
'    End If
    
    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    I = InStr(1, CADENA, "<>")
    If I = 0 Then
        'IGUAL A valor
        Cad = " = "
        Else
            'Distinto a valor
        Cad = " <> "
    End If
    'Verdadero o falso
    I = InStr(1, CADENA, "V")
    If I > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = campo & " " & Cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Private Function CararacteresCorrectos(vcad As String, Tipo As String) As Byte
Dim I As Integer
Dim CH As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " ", "-"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "a" To "z"
            Case "A" To "Z"
            Case "0" To "9"
            Case "*", "%", "?", "_", "\", "/", ":", ".", " " ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "<", ">"
            Case "Ñ", "ñ"
            Case "-", "+", ",", """" 'Añade Laura
            'Abril 2014
            Case "[", "]"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
    
Case "F"
    'Tipo Fecha. Aceptamos Numeros , "/" ,":"
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next I

Case "B"
    'Numeros , "/" ,":"
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next I
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function


Public Function QuitarCaracterEnter(vcad As String) As String
Dim I As Integer

    Do
        I = InStr(1, vcad, Chr(13))
        If I > 0 Then 'Hay ENTER
            vcad = Mid(vcad, 1, I - 1) & Mid(vcad, I + 2)
        End If
    Loop Until I = 0
    QuitarCaracterEnter = vcad
End Function




Public Function QuitarCaracterNULL(vcad As String) As String
Dim I As Integer

    Do
        I = InStr(1, vcad, vbNullChar)
        If I > 0 Then 'Hay null
            vcad = Mid(vcad, 1, I - 1) & Mid(vcad, I + 2)
        End If
    Loop Until I = 0
    QuitarCaracterNULL = vcad
End Function




'======== Añade: Laura
Public Function ContieneCaracterBusqueda(CADENA As String) As Boolean
'Comprueba si la cadena contiene algun caracter especial de busqueda
' >,>,>=,: , ....
'si encuentra algun caracter de busqueda devuelve TRUE y sale
Dim b As Boolean
Dim I As Integer
Dim CH As String


    'Febrero 2012, el 29
    'NULL
    If UCase(CADENA) = "NULL" Then
        ContieneCaracterBusqueda = True
        Exit Function
    End If

    'For i = 1 To Len(cadena)
    I = 1
    b = False
    Do
        CH = Mid(CADENA, I, 1)
        Select Case CH
            Case "<", ">", ":", "="
                b = True
            Case "*", "%", "?", "_", "\", ":" ', "."
                b = True
            Case Else
                b = False
        End Select
    'Next i
        I = I + 1
    Loop Until (b = True) Or (I > Len(CADENA))
    ContieneCaracterBusqueda = b
End Function

