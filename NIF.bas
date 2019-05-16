Attribute VB_Name = "NIF"
 

Public Function Comprobar_NIF(NIF As String) As Boolean

    '-- Comprobación general de NIF

    If Len(NIF) <> 9 Then

        Comprobar_NIF = False

        Exit Function

    Else

        If IsNumeric(Mid(NIF, 1, 1)) Then

            '-- Comienza por número

            If IsNumeric(Mid(NIF, 9, 1)) Then

                Comprobar_NIF = False

                Exit Function

            Else

                Comprobar_NIF = Comprobar_NIF_PersonaFisica(NIF)

            End If

        Else

            '-- comienza por letra

            If IsNumeric(Mid(NIF, 9, 1)) Then

                '-- Acaba en número
                'If InStr(1, "ABCDEFGHPQSNU", Mid(NIF, 1, 1)) <> 0 Then
                'If InStr(1, "ABCDEFGHPQSNUV", Mid(NIF, 1, 1)) <> 0 Then   Marzo 2019
                If InStr(1, "ABCDEFGJHPQSNUV", Mid(NIF, 1, 1)) <> 0 Then
                    '-- Es una sociedad , juridica.......

                    Comprobar_NIF = Comprobar_NIF_Sociedad(NIF)

                ElseIf InStr(1, "T", Mid(NIF, 1, 1)) <> 0 Then

                    '-- Es un NIF antiguo que no lleva comprobación

                    Comprobar_NIF = True

                End If

            Else

                '-- Acaba en letra

                If InStr(1, "ABCDEFGHPQRSN", Mid(NIF, 1, 1)) <> 0 Then

                    '-- Es una sociedad , congregacion religiosa(R)....

                    Comprobar_NIF = Comprobar_NIF_Sociedad(NIF)

                ElseIf InStr(1, "MXY", Mid(NIF, 1, 1)) <> 0 Then

                    '-- Es un extranjero

                    Comprobar_NIF = Comprobar_NIF_PersonaExtranjera(NIF)

                ElseIf InStr(1, "KL", Mid(NIF, 1, 1)) <> 0 Then

                    '-- Es un NIF antiguo que no lleva comprobación

                    Comprobar_NIF = True

                End If

            End If

        End If

    End If

End Function

Public Function Comprobar_NIF_PersonaFisica(NIF As String) As Boolean

    Dim mCadena As String

    Dim mLetra As String

    Dim m23 As Integer

    mCadena = "TRWAGMYFPDXBNJZSQVHLCKE"

    '-- Tomamos el NIF propiamente dicho y calculamos el módulo 23

    m23 = Val(Mid(NIF, 1, 8)) Mod 23

    mLetra = Mid(mCadena, m23 + 1, 1)

    '-- Validamos que la letra es correcta

    If Mid(NIF, 9, 1) = mLetra Then

        Comprobar_NIF_PersonaFisica = True

    Else

        Comprobar_NIF_PersonaFisica = False

    End If

End Function

 

Public Function Comprobar_NIF_PersonaExtranjera(NIF As String) As Boolean

    Dim mCadena As String

    Dim mLetra As String

    Dim m23 As Integer

    mCadena = "DTRWAGMYFPXBNJZSQVHLCKE" 'ANTES

    mCadena = "TRWAGMYFPDXBNJZSQVHLCKE"

    '-- Tomamos el NIF propiamente dicho y calculamos el módulo 23

     Aux = Mid(NIF, 2, 7)
    If Mid(NIF, 1, 1) = "Y" Then Aux = "1" & Aux

    m23 = Val(Aux) Mod 23

    mLetra = Mid(mCadena, m23 + 1, 1)

    '-- Validamos que la letra es correcta

    If Mid(NIF, 9, 1) = mLetra Then

        Comprobar_NIF_PersonaExtranjera = True

    Else

        Comprobar_NIF_PersonaExtranjera = False

    End If

End Function

 

Public Function Comprobar_NIF_Sociedad(NIF As String) As Boolean

    Dim mCadena As String

    Dim mLetra As String

    Dim vNif As String

    Dim mN2 As String

    Dim i, i2 As Integer

    Dim Suma, Control As Long

    mCadena = "ABCDEFGHIJV"

    vNif = Mid(NIF, 2, 7)

    '-- Sumamos las cifras pares

    For i = 2 To Len(vNif) Step 2

        Suma = Suma + Val(Mid(vNif, i, 1))

    Next i

    '-- Ahora las impares * 2, y sumando las cifras del resultado

    For i = 1 To Len(vNif) Step 2

        mN2 = CStr(Val(Mid(vNif, i, 1)) * 2)

        For i2 = 1 To Len(mN2)

            Suma = Suma + Val(Mid(mN2, i2, 1))

        Next i2

    Next i

    '-- Ya tenemos la suma y calculamos el control

    Control = 10 - (Suma Mod 10)

    If Control = 10 Then Control = 0

    mLetra = Mid(NIF, 9, 1)

    If IsNumeric(mLetra) Then

        If Val(mLetra) = Control Then

            Comprobar_NIF_Sociedad = True

        Else

            Comprobar_NIF_Sociedad = False

        End If

    Else

        If Control = 0 Then Control = 10

        If mLetra = Chr(64 + Control) Then

            Comprobar_NIF_Sociedad = True

        Else

            Comprobar_NIF_Sociedad = False

        End If

    End If

End Function





'========= AÑADE: LAURA   ==========================
'========= Fecha: 20/8/04 ==========================
'Public Function ValidarNIF_old(ByRef Txt As TextBox) As Boolean
'    ValidarNIF_old = True
'    With Txt
'        If .Text <> "" Then
'            If Not Comprobar_NIF(.Text) Then
'                MsgBox "El NIF no es correcto.", vbExclamation
''                PonerFoco Txt
'                ValidarNIF_old = False
'            End If
'        End If
'    End With
'End Function


Public Function ValidarNIF(valorNIF As String) As Boolean

    ValidarNIF = True
    If valorNIF <> "" Then
        If Not Comprobar_NIF(valorNIF) Then
            MsgBox "El NIF no es correcto.", vbExclamation
            ValidarNIF = False
        End If
    End If
End Function
