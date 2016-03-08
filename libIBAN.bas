Attribute VB_Name = "libIBAN"
Option Explicit




'A partir de una cuenta banco formateada y todos los numeros juntos (chr(20))
'  devuelve DOS(2) caracteres del IBAN
'  calculados como dice la formula
'  i=ctabanco_con ES... mod 97
'  i = 98-i
' format(i,"00"             'para que copie                     'es lo que devuelve
'
'Puede NO poner pais. Sera ES
Public Function DevuelveIBAN2(PAIS As String, ByVal CtaBancoFormateada As String, DosCaracteresIBAN As String) As Boolean
Dim Aux As String
Dim N As Long
Dim CadenaPais As String
On Error GoTo EDevuelveIBAN
    DevuelveIBAN2 = False
    DosCaracteresIBAN = ""
    
    
    
    If PAIS = "" Then
        PAIS = "ES"
    Else
        If Len(PAIS) <> 2 Then
            PAIS = "ES"
        Else
            PAIS = UCase(PAIS)
        End If
    End If
    
    
    'Ejemplo mio: 20770294901101867914  IBAN: 41
    'Construir el IBAn:
    'A la derecha de la cuenta se pone
    '   el ES00-->   20770294961101915202ES00 ->92
    'Se transforma las letras ES a numeros.
    ' E=14 S=28
    '           ->>> 20770294961101915202 142800
    If PAIS = "ES" Then
        CadenaPais = "1428"
    Else
        N = Asc(Mid(PAIS, 1, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CStr(N)
        N = Asc(Mid(PAIS, 2, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CadenaPais & CStr(N)
    End If
    'Se le añaden 2 ceros al final
    CadenaPais = CadenaPais & "00"
    'Esta es la cadena para ES. SiCadenaPais  fuera otro pais es aqui donde hay que cambiar
    CtaBancoFormateada = CtaBancoFormateada & "142800"
    Aux = ""
    While CtaBancoFormateada <> ""
        If Len(CtaBancoFormateada) >= 6 Then
            Aux = Aux & Mid(CtaBancoFormateada, 1, 6)
            CtaBancoFormateada = Mid(CtaBancoFormateada, 7)
        Else
            Aux = Aux & CtaBancoFormateada
            CtaBancoFormateada = ""
        End If
        
        N = CLng(Aux)
        N = N Mod 97
        
        Aux = CStr(N)
    Wend
        
    N = 98 - N
    
    DosCaracteresIBAN = Format(N, "00")
    DevuelveIBAN2 = True
    Exit Function
EDevuelveIBAN:
    CadenaPais = Err.Description
    CadenaPais = Err.Number & "   " & CadenaPais
    MsgBox "Devuelve IBAN. " & vbCrLf & CadenaPais, vbExclamation
    Err.Clear
End Function



'DADO un txtIBAN de 4 posciones los dos primeros podriamos comprobar el pais en una tabla
' y los dos sugientes son numericos
Public Function IBAN_Correcto(IBAN As String) As Boolean
Dim Aux As String
    IBAN_Correcto = False
    Aux = ""
    If Len(IBAN) <> 4 Then
        Aux = "Longitud incorrecta"
    Else
        If IsNumeric(Mid(Aux, 3, 2)) Then
            Aux = "Digitos 3 y 4 deben ser numericos"
        Else
            'Podriamos comprobar lista de paises
    
        End If
    End If
    If Aux <> "" Then
        MsgBox Aux, vbExclamation
    Else
        IBAN_Correcto = True
    End If
End Function



'A partir de una cadena, con letras y numeros convertira
'en mod 97,10 Norma ISO 7064
'Para ello los caracteres se pasan a dos digitos
Public Function CadenaTextoMod97(CADENA As String) As String
Dim I As Integer
Dim C As String
Dim N As Long

    CADENA = Trim(CADENA)
    C = ""
    'Substitucion de texto por caracteres
    For I = 1 To Len(CADENA)
        N = Asc(Mid(CADENA, I, 1))
        If N >= 48 Then
            If N <= 57 Then
                'Es numerico 0..9
                'C = C & CStr(N)
            Else
                If N < 65 Or N > 90 Then
                    'MAL. No es un caracter ASCII entre A..Z  (10..35)
                    N = 0
                Else
                    N = N - 55  'el ascci menos 55 (0...35)
                End If
            End If
        End If
        If N = 0 Then
            CadenaTextoMod97 = "Caracter NO valido: " & Mid(CADENA, I, 1) & " --- " & CADENA
            Exit Function
        Else
            If N >= 48 Then
                'Es un numero
                C = C & Chr(N)
            Else
                C = C & CStr(N)
            End If
        End If
        
    Next
    
    
    
    'Ya tengo C que es numerica
    CADENA = C
    C = ""
    While CADENA <> ""
        If Len(CADENA) >= 6 Then
            C = C & Mid(CADENA, 1, 6)
            CADENA = Mid(CADENA, 7)
        Else
            C = C & CADENA
            CADENA = ""
        End If
        
        N = CLng(C)
        N = N Mod 97
        
        C = CStr(N)
    Wend
        
    N = 98 - N
    CadenaTextoMod97 = Format(N, "00")
End Function
