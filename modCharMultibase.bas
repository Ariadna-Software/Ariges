Attribute VB_Name = "modCharMultibase"



Public Function RevisaCaracterMultibase(CADENA As String) As String
Dim i As Integer
Dim J As Integer
Dim L As String
Dim C As String

    L = ""
    For i = 1 To Len(CADENA)
        C = Mid(CADENA, i, 1)
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
    Next i
    
    
    
' CAMBIOS EN MySQL a MySQL by MASL 08092009


       If InStr(CADENA, "Ã‘") Then L = Replace(CADENA, "Ã‘", "Ñ")
       If InStr(CADENA, "Âª") Then L = Replace(CADENA, "Âª", "ª")
       If InStr(CADENA, "Âº") Then L = Replace(CADENA, "Âº", "º")
       If InStr(CADENA, "Ã‘") Then L = Replace(CADENA, "Ã‘", "Ñ")
       If InStr(CADENA, "Â§") Then L = Replace(CADENA, "Â§", "º")
       If InStr(CADENA, "š") Then L = Replace(CADENA, "š", "Ü")

RevisaCaracterMultibase = L

End Function
