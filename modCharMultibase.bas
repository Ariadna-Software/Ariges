Attribute VB_Name = "modCharMultibase"



Public Function RevisaCaracterMultibase(CADENA As String) As String
Dim I As Integer
Dim J As Integer
Dim L As String
Dim C As String

    L = ""
    For I = 1 To Len(CADENA)
        C = Mid(CADENA, I, 1)
        J = Asc(C)
        If J > 125 Then
            Select Case J
            Case 128
                C = "Ç"
            Case 164  'ñ minuscula
                C = "ñ"
            Case 165
                'Es la Ñ
                C = "Ñ"
            Case 166
                C = "ª"
            Case 167, 186
                C = "º"
            Case 194
                C = ""
            Case 209
            
            Case Else
                
            End Select
        End If
        L = L & C
    Next I
    
    
    
' CAMBIOS EN MySQL a MySQL by MASL 08092009


       If InStr(L, "Ã‘") Then L = Replace(L, "Ã‘", "Ñ")
       If InStr(L, "Âª") Then L = Replace(L, "Âª", "ª")
       If InStr(L, "Âº") Then L = Replace(L, "Âº", "º")
       If InStr(L, "Ã‘") Then L = Replace(L, "Ã‘", "Ñ")
       If InStr(L, "Â§") Then L = Replace(L, "Â§", "º")
       If InStr(L, "š") Then L = Replace(L, "š", "Ü")
       If InStr(L, "Ã³") Then L = Replace(L, "Ã³", "ó")
       If InStr(L, "Ã­") Then L = Replace(L, "Ã­", "í")
       If InStr(L, "Ãº") Then L = Replace(L, "Ãº", "ú")
       
       'Septiebre 2020
       If InStr(L, "Ã“") Then L = Replace(L, "Ã“", "Ó")
       If InStr(L, "Ã‰") Then L = Replace(L, "Ã‰", "É")
       
       
       
       
       'este debe ser el ultimo
       If InStr(L, "Ã") Then L = Replace(L, "Ã", "Á")





RevisaCaracterMultibase = L

End Function
