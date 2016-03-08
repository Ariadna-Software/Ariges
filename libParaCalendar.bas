Attribute VB_Name = "libParaCalendar"
Public vbMyMonday    '///  Primer dia semana no es vbMonday.  Sera ---> vbMyMonday
                     'Error reportado por microsoft.
                     'La constante no esta asociada al componente, con lo cual
                     'en unas versiones funciona correctaamente y en otras no
                     



'Funcionamiento.
'Habra un fichero en app.path que tendra la extension .dia
'Lo otro sera el numero que asignaremos a la variable
Public Sub FijarPrimerDiaSemana()
Dim C As String
    On Error GoTo EFijarPrimerDiaSemana
    vbMyMonday = vbMonday
    C = Dir(App.Path & "\*.dia", vbArchive)
    If C <> "" Then
        C = Mid(C, 1, (InStr(1, C, ".") - 1))
        vbMyMonday = CInt(C)
    End If
    Exit Sub
EFijarPrimerDiaSemana:
    Err.Clear
End Sub

