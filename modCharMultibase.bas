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
                C = "�"
            Case 164  '� minuscula
                C = "�"
            Case 165
                'Es la �
                C = "�"
            Case 166
                C = "�"
            Case 167, 186
                C = "�"
            Case 194
                C = ""
            Case 209
            
            Case Else
                
            End Select
        End If
        L = L & C
    Next I
    
    
    
' CAMBIOS EN MySQL a MySQL by MASL 08092009


       If InStr(L, "Ñ") Then L = Replace(L, "Ñ", "�")
       If InStr(L, "ª") Then L = Replace(L, "ª", "�")
       If InStr(L, "º") Then L = Replace(L, "º", "�")
       If InStr(L, "Ñ") Then L = Replace(L, "Ñ", "�")
       If InStr(L, "§") Then L = Replace(L, "§", "�")
       If InStr(L, "�") Then L = Replace(L, "�", "�")
       If InStr(L, "ó") Then L = Replace(L, "ó", "�")
       If InStr(L, "í") Then L = Replace(L, "í", "�")
       If InStr(L, "ú") Then L = Replace(L, "ú", "�")
       
       'Septiebre 2020
       If InStr(L, "Ó") Then L = Replace(L, "Ó", "�")
       If InStr(L, "É") Then L = Replace(L, "É", "�")
       
       
       
       
       'este debe ser el ultimo
       If InStr(L, "�") Then L = Replace(L, "�", "�")





RevisaCaracterMultibase = L

End Function
