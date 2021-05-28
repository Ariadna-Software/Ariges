Attribute VB_Name = "modEnumResIcons"
'---------------------------------------------------------------------------------------
' Module      : modEnumResIcons.bas
' DateTime    : 03/04/2004 21.52
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Project     : EnumResource.vbp
' Purpose     : Load resources ICON from a EXE/DLL library
' Descritpion : This project show how to load Windows XP (32bpp) icons format
'               from executable files.
' Comments    : Read the 'frmEnumRes.frm' Comments section.
'               Also, you can find other info on README.TXT file
'---------------------------------------------------------------------------------------
Option Explicit

' *** MEU: per a dir-li en quin ImageList possa el resultat i el tamany ***
Public opcio As Integer
Public tamany As Integer
' *************************************************************
Public ghmodule As Long
Public giSize As Integer
Public giColorDepth As Integer
Public gbAllSizeFormat As Boolean

Public arrSize As Long

Private arIcon(1 To 4, 1 To 4)
Private Const SIZE_16 = 1
Private Const SIZE_24 = 2
Private Const SIZE_32 = 3
Private Const SIZE_48 = 4
Private Const COLOR_4 = 1
Private Const COLOR_16 = 2
Private Const COLOR_24 = 3
Private Const COLOR_32 = 4

Public Declare Function LoadLibraryEx Lib "kernel32.DLL" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Const DONT_RESOLVE_DLL_REFERENCES = &H1
Public Const LOAD_LIBRARY_AS_DATAFILE = &H2
Public Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLibModule As Long) As Long
Public Declare Function EnumResourceNames Lib "KERNEL32" Alias "EnumResourceNamesA" (ByVal ghmodule As Long, ByVal lpType As ResType, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'String management
Public Declare Function StrLen Lib "KERNEL32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function StrCpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Const DIFFERENCE = 11
Public Enum ResType ' Resource Types
    RT_FIRST = 1&
    RT_CURSOR = 1&
    RT_BITMAP = 2&
    RT_ICON = 3&
    RT_MENU = 4&
    RT_DIALOG = 5&
    RT_STRING = 6&
    RT_FONTDIR = 7&
    RT_FONT = 8&
    RT_ACCELERATOR = 9&
    RT_RCDATA = 10&
    RT_MESSAGETABLE = (11)
    RT_GROUP_CURSOR = (RT_CURSOR + DIFFERENCE)  ' (12)
    RT_GROUP_ICON = (RT_ICON + DIFFERENCE)      ' (14)
    RT_VERSION = (16)
    'RT_DLGINCLUDE = (17)
    'RT_PLUGPLAY = (19)
    'RT_VXD = (20)
    'RT_ANICURSOR = (21)
    'RT_ANIICON = (22)
    'RT_HTML = (23)
    RT_LAST = (16)
End Enum

' IMAGES management
'Const IMAGE_BITMAP = 0
'Const IMAGE_ICON = 1
'Const IMAGE_CURSOR = 2
'Const IMAGE_ENHMETAFILE = 3
'Private Const LR_COLOR As Long = &H2
'Private Const LR_COPYDELETEORG As Long = &H8
'Private Const LR_COPYFROMRESOURCE As Long = &H4000
'Private Const LR_COPYRETURNORG As Long = &H4
'Private Const LR_CREATEDIBSECTION As Long = &H2000
'Private Const LR_DEFAULTCOLOR As Long = &H0
'Private Const LR_DEFAULTSIZE As Long = &H40
'Private Const LR_LOADFROMFILE As Long = &H10
Private Const LR_LOADMAP3DCOLORS As Long = &H1000
'Private Const LR_LOADTRANSPARENT As Long = &H20
'Private Const LR_MONOCHROME As Long = &H1
'Private Const LR_SHARED As Long = &H8000
'Private Const LR_VGACOLOR As Long = &H80

'/ mie
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
'Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
'Private Declare Function GetLastError Lib "kernel32.dll" () As Long
'Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32.DLL" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FindResource Lib "kernel32.DLL" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function FindResourceByNum Lib "KERNEL32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long

Private Declare Function LoadResource Lib "kernel32.DLL" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As Long
Private Declare Function LockResource Lib "KERNEL32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "KERNEL32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "KERNEL32" (ByVal hResData As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type PICTDESC
    cbSizeofStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

'Private Declare Function GetBitmapDimensionEx Lib "gdi32.dll" (ByVal hBitmap As Long, lpDimension As SIZE) As Long
Private Type Size
    cx As Long
    cy As Long
End Type




'---------------------------------------------------------------------------------------
' Procedure   : GetPictureRes
' DateTime    : 04/04/2004 17.47
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Load the resource from library
' Descritpion : There is many resource types, we can check which you need (ICON)
' Comments    : I have keep the other 'Case' for future use (if you find useful...)
'---------------------------------------------------------------------------------------

Public Function GetPictureRes(ByVal sResType As String, ByVal sResName As String, ByVal iSize As Integer, ByVal iColorDepth As Integer) As StdPicture
'    Recupera la risorsa
    Dim hData As Long
    Dim arr() As Byte, vRet As Variant
    Select Case sResType
        Case "1", "3" ' Hardware dependent cursor or icon.
            vRet = GetDataArray(sResType, sResName, iSize, iColorDepth)
            If CStr(vRet) = "0" Then
                Set GetPictureRes = Nothing
                Exit Function
            Else
                arr = vRet
                hData = CreateIconFromResourceEx(arr(0), UBound(arr) + 1, CLng(sResType) - 1, &H30000, 0, 0, LR_LOADMAP3DCOLORS)
            End If
            
        Case "2"  ' Bitmap
            hData = LoadImage(ghmodule, sResName, 0, 0, 0, LR_LOADMAP3DCOLORS)
        Case "12" ' Hardware independent cursor
            hData = LoadImage(ghmodule, sResName, 2, 0, 0, LR_LOADMAP3DCOLORS)
        Case "14" ' Hardware independent icon
            hData = LoadImage(ghmodule, sResName, 1, 0, 0, LR_LOADMAP3DCOLORS)
    End Select
    If hData = 0 Then Exit Function
    
    Set GetPictureRes = IconToPicture(hData)
    
End Function

'---------------------------------------------------------------------------------------
' Procedure   : IconToPicture
' DateTime    : 04/04/2004 17.46
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Export the resource from ICON to PICTURE
' Descritpion :
' Comments    :
'---------------------------------------------------------------------------------------

Private Function IconToPicture(ByVal hIcon As Long) As StdPicture
    
    If hIcon = 0 Then Exit Function
    Dim oNewPic As Picture
    Dim tPicConv As PICTDESC
    Dim IGuid As GUID
    With tPicConv
        .cbSizeofStruct = Len(tPicConv)
        .PicType = vbPicTypeIcon
        .hImage = hIcon
    End With
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .data4(0) = &H8B
        .data4(1) = &HBB
        .data4(2) = &H0
        .data4(3) = &HAA
        .data4(4) = &H0
        .data4(5) = &H30
        .data4(6) = &HC
        .data4(7) = &HAB
    End With
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
    Set IconToPicture = oNewPic
End Function

Public Function GetDataArray(ByVal ResType As String, ByVal ResName As String, ByVal iSize As Integer, ByVal iColorDepth As Integer) As Variant
    Dim hRsrc As Long
    Dim hGlobal As Long
    Dim arrData() As Byte
    Dim lpData As Long
    'Dim arrSize As Long    ' Global: to get image size
    If IsNumeric(ResType) Then hRsrc = FindResourceByNum(ghmodule, ResName, CLng(ResType))
    If hRsrc = 0 Then hRsrc = FindResource(ghmodule, ResName, ResType)
    If hRsrc = 0 Then Exit Function
    hGlobal = LoadResource(ghmodule, hRsrc)
    lpData = LockResource(hGlobal)
    
    arrSize = SizeofResource(ghmodule, hRsrc)
    Dim iNDXSize As Integer, iNDXColor As Integer
    Select Case iSize
        Case 16
            iNDXSize = 1
        Case 24
            iNDXSize = 2
        Case 32
            iNDXSize = 3
        Case 48
            iNDXSize = 4
    End Select
    Select Case iColorDepth
        Case 4
            iNDXColor = 1
        Case 16
            iNDXColor = 2
        Case 24
            iNDXColor = 3
        Case 32
            iNDXColor = 4
    End Select
    
    
    If Not gbAllSizeFormat Then
        ' Load images that only match the color depth
        If arrSize <> arIcon(iNDXSize, iNDXColor) Then
            GetDataArray = 0
            Exit Function
        End If
    End If
    
    If arrSize = 0 Then
        GetDataArray = 0
        Exit Function
    End If
    
    ReDim arrData(arrSize - 1)
    Call CopyMemory(arrData(0), ByVal lpData, arrSize)
    Call FreeResource(hGlobal)
    GetDataArray = arrData

End Function

Public Function EnumResNameProc(ByVal ghmodule As Long, ByVal lpszType As ResType, ByVal lpszName As Long, ByVal lParam As Long) As Long
    Dim sNumber As String, IsNum As Boolean
    
    If (lpszName > &HFFFF&) Or (lpszName < 0) Then
        sNumber = PtrToVBString(lpszName)
        IsNum = False
    Else
        sNumber = CStr(lpszName)
        IsNum = True
    End If
    
    ' 16x16
    arIcon(SIZE_16, COLOR_4) = 296
    arIcon(SIZE_16, COLOR_16) = 1384
    arIcon(SIZE_16, COLOR_24) = 872
    arIcon(SIZE_16, COLOR_32) = 1128
    ' 24x24
    arIcon(SIZE_24, COLOR_4) = 488
    arIcon(SIZE_24, COLOR_16) = 1736
    arIcon(SIZE_24, COLOR_24) = 1864
    arIcon(SIZE_24, COLOR_32) = 2440
    ' 32x32
    arIcon(SIZE_32, COLOR_4) = 744
    arIcon(SIZE_32, COLOR_16) = 2216
    arIcon(SIZE_32, COLOR_24) = 3240
    arIcon(SIZE_32, COLOR_32) = 4264
    ' 48x48
    arIcon(SIZE_48, COLOR_4) = 1640
    arIcon(SIZE_48, COLOR_16) = 3752
    arIcon(SIZE_48, COLOR_24) = 7336
    arIcon(SIZE_48, COLOR_32) = 9640
    
    
    ' *** MEU: per ara li pose ací el tamany i els colors ***
    'giSize = 16
    'giSize = 24
    giSize = tamany
    giColorDepth = 32
    ' *******************************************************
        
    If IsNum Then
        If lpszType = RT_ICON Then
            LoadIconRes lpszType, sNumber, giSize, giColorDepth
        End If
    End If
    EnumResNameProc = 1
End Function
Private Function PtrToVBString(ByVal lpszBuffer As Long) As String
    Dim Buffer As String, LenBuffer As Long
    LenBuffer = StrLen(lpszBuffer)
    Buffer = String(LenBuffer + 1, 0)
    StrCpy Buffer, lpszBuffer
    PtrToVBString = Left(Buffer, LenBuffer)
End Function




'---------------------------------------------------------------------------------------
' Procedure   : LoadIconRes
' DateTime    : 04/04/2004 17.50
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Get each image format from the icon resource
' Descritpion :
' Comments    : Use GetPictureRes to get the image
'---------------------------------------------------------------------------------------

Public Sub LoadIconRes(ByVal sResType As ResType, ByVal sResNumber As String, ByVal iSize As Integer, ByVal iColorDepth As Integer)
    Dim sResName As String
    Dim hPicture As StdPicture
    
    sResName = sResNumber
    If IsNumeric(sResName) Then sResName = "#" & sResName
    
    ' load the icon that match the Size (lSize)
    Set hPicture = GetPictureRes(sResType, sResName, iSize, iColorDepth)
    
    ' Here I use a Image control to load icons and get the width of it
    ' Brrrrrrrrr.... Isn't good programming, but work!
    ' Image control is good to retrieve the size of icon, because is autosize ;-))
    ' Feel free to choose another method! ;-))))))))))
    Dim H As Long, W As Long
'    Set frmEnumRes.Image1.Picture = hPicture
'    w = frmEnumRes.Image1.Width / Screen.TwipsPerPixelX
'    h = frmEnumRes.Image1.Height / Screen.TwipsPerPixelY
    W = iSize
    H = iSize
       
    If Not hPicture Is Nothing Then
            

            
            'NORMAL. LO que habia
            If opcio = 1 Then
            '    frmPpaln.ImageList1.ListImages.Add , sResName & " " & CStr(arrSize) & " " & W & "x" & H, hPicture
            ElseIf opcio = 2 Then
                frmPpal.ImgListComun2.ListImages.Add , sResName & " " & CStr(arrSize) & " " & W & "x" & H, hPicture
            '++
            ElseIf opcio = 3 Then
                frmPpal.imgListComun_BN2.ListImages.Add , sResName & " " & CStr(arrSize) & " " & W & "x" & H, hPicture
            ElseIf opcio = 4 Then
                frmPpal.imgListComun_OM2.ListImages.Add , sResName & " " & CStr(arrSize) & " " & W & "x" & H, hPicture
              ElseIf opcio = 5 Then
                frmPpal.imgListComun16.ListImages.Add , sResName & " " & CStr(arrSize) & " " & W & "x" & H, hPicture
            ElseIf opcio = 6 Then
                frmPpal.imgListComun_BN16.ListImages.Add , sResName & " " & CStr(arrSize) & " " & W & "x" & H, hPicture
            ElseIf opcio = 7 Then
                frmPpal.imgListComun_OM16.ListImages.Add , sResName & " " & CStr(arrSize) & " " & W & "x" & H, hPicture
        
            End If
            

        
    End If
    
End Sub
