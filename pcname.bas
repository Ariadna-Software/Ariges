Attribute VB_Name = "pcname"
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Public Const MAX_COMPUTERNAME_LENGTH = 255


'------------------------------------------------------------------------
'Declaraciones Para obtener nombre del PC conectado por TErminal Server
'------------------------------------------------------------------------
Private Const WTS_CURRENT_SERVER_HANDLE = 0&

Private Enum WTS_INFO_CLASS
  WTSInitialProgram
  WTSApplicationName
  WTSWorkingDirectory
  WTSOEMId
  WTSSessionId
  WTSUserName
  WTSWinStationName
  WTSDomainName
  WTSConnectState
  WTSClientBuildNumber
  WTSClientName
  WTSClientDirectory
  WTSClientProductId
  WTSClientHardwareId
  WTSClientAddress
  WTSClientDisplay
  WTSClientProtocolType
End Enum


Private Declare Function GetCurrentProcessId Lib "Kernel32.dll" () As Long
Private Declare Function ProcessIdToSessionId Lib "Kernel32.dll" (ByVal dwProcessId As Long, ByRef pSessionId As Long) As Long


Private Declare Function WTSQuerySessionInformation _
    Lib "wtsapi32.dll" Alias "WTSQuerySessionInformationA" ( _
    ByVal hServer As Long, ByVal SessionID As Long, _
    ByVal WTSInfoClass As WTS_INFO_CLASS, _
    ByRef ppBuffer As Long, _
    ByRef pBytesReturned As Long _
    ) As Long

Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" ( _
    ByVal pMemory As Long)
    
    Private Declare Function lstrlenA Lib "kernel32" ( _
    ByVal lpString As String) As Long

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As Long) As Long




'------------------------------------------------------------------------
'------------------------------------------------------------------------
' Lanza visores predeterminados por MIME
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long






Public Function LanzaVisorMimeDocumento(Formhwnd As Long, Archivo As String)
    Call ShellExecute(Formhwnd, "Open", Archivo, "", "", 1)
End Function


Public Function lanzaImpresionShellDirecta(Formhwnd As Long, Archivo As String)
    Call ShellExecute(Formhwnd, "print", Archivo, vbNullString, vbNullString, 1)
End Function


'------------------------------------------------------------------------
'------------------------------------------------------------------------



Private Function ComputerNameL() As String
    'Devuelve el nombre del equipo actual
    Dim sComputerName As String
    Dim ComputerNameLength As Long
    
    sComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
    ComputerNameLength = MAX_COMPUTERNAME_LENGTH
    Call GetComputerName(sComputerName, ComputerNameLength)
     ComputerNameL = Mid(sComputerName, 1, ComputerNameLength)
    
End Function




'===========================================
'===== LAURA            Fecha: 18/01/06
'===========================================

'Public Function ComputerNameTServer() As String
''lee por terminal server el nombre del pc de la conexion
''para ello lee de un fichero en la maquina local nombre.ini el nombre del pc q se conecto
''accediendo mediante \\tsclien\c\nombre.ini
'    Dim NomFich As String
'    Dim NF As Integer
'    Dim cad As String
'
'    NomFich = "\\tsclient\c\nombre.ini"
'
'    On Error GoTo ECompuName
'
'    If Dir(NomFich, vbArchive) <> "" Then
'        NF = FreeFile
'        Open NomFich For Input As #NF
'        Line Input #NF, cad
'        Close #NF
'    Else
'        cad = ""
'        MsgBox "No se ha podido encontrar el fichero C:\nombre.ini en la maquina local.", vbExclamation
'    End If
'
'    ComputerNameTServer = cad
'    MsgBox "Nombre PC: " & cad, vbInformation
'
'ECompuName:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener nombre del PC por Terminal Server.", Err.Description
'End Function




'=================================================================
'===== LAURA            Fecha: 18/01/06
'===== Funciones para obtener Computer Name desde Terminal Server
'=================================================================

Private Function PointerToStringA(ByVal lpStringA As Long) As String
   Dim nLen As Long
   Dim sTemp As String

   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         sTemp = String(nLen, vbNullChar)
         lstrcpy sTemp, ByVal lpStringA
         PointerToStringA = sTemp
      End If
   End If
End Function



Private Function GetComputerNameTS() As String
'Devuelve el nombre del PC de la sesion de Terminal Server
    Dim RetVal As Long
    Dim lpBuffer As Long
    Dim Count As Long
    Dim p As Long
    Dim QueryInfo As String
    Dim CurrentSessionId As Long
    Dim CurrentProcessId As Long

                                   
     CurrentProcessId = GetCurrentProcessId()
     RetVal = ProcessIdToSessionId(CurrentProcessId, CurrentSessionId)
'     MsgBox "Current Process: " & CurrentProcessId
'     MsgBox "Current Session ID: " & CurrentSessionId
     
                                   
    RetVal = WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE, _
                CurrentSessionId, _
                WTSClientName, _
                lpBuffer, _
                Count)
                
                                   
    If RetVal Then
        ' WTSEnumerateProcesses was successful.

        p = lpBuffer
        QueryInfo = PointerToStringA(p)
        
        ' Free the memory buffer.
        WTSFreeMemory lpBuffer

     Else
        ' Error occurred calling WTSEnumerateProcesses.
        ' Check Err.LastDllError for error code.
        If Err.LastDllError <> 1151 Then
            '1151: ERROR_APP_WRONG_OS = The specified program is not a Windows or MS-DOS program.
            'En el SERVER no hay instalado:Requires Windows Server "Longhorn", Windows Server 2003, or Windows 2000 Server.
            
            MsgBox "An error occurred calling WTSQuerySessionInformation.  " & _
            "Check the Platform SDK error codes in the MSDN Documentation " & _
            "for more information.", vbCritical, "ERROR " & Err.LastDllError
        End If
    End If
   
    GetComputerNameTS = QuitarCaracterNULL(QueryInfo)
'    If QueryInfo = "" Then QueryInfo = ComputerName
'    GetComputerNameTS = QueryInfo
End Function



Public Function ComputerName() As String
    Dim nom As String
    
    'Por Terminal Server
    nom = GetComputerNameTS
    
    'Si no conectado por TServer mirar en local
    If nom = "" Then nom = ComputerNameL
    ComputerName = nom
End Function
