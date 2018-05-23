Attribute VB_Name = "ModuloBase2"
'*********************************************************
'Declaração de variáveis para fechar
'todos os aplicativos
'*********************************************************
Private Declare Function GetShortPathName Lib "KERNEL32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Dim lRebootNow As Boolean
Dim cRegistry As String
Dim cPath As String
'Dim hFile As Long
Dim nLoop As Long

Public Declare Function CloseHandle Lib "KERNEL32" (ByVal hFile As Long) As Long
'Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type
Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
Public Const MaxLFNPath = 260
Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MaxLFNPath
cShortFileName As String * 14
End Type
Public WFD As WIN32_FIND_DATA, hItem&, hFile&
Public FileSpec$
Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const MAX_PATH As Integer = 260
Public Const TH32CS_SNAPPROCESS = &H2
Public Const PROCESS_TERMINATE = &H1
Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * MAX_PATH
End Type
Public Declare Function CreateToolhelpSnapshot Lib "KERNEL32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "KERNEL32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "KERNEL32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function GetExitCodeProcess Lib "KERNEL32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function TerminateProcess Lib "KERNEL32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "KERNEL32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Const ANYSIZE_ARRAY = 1
Type LARGE_INTEGER
lowpart As Long
highpart As Long
End Type
Type LUID_AND_ATTRIBUTES
pLuid As LARGE_INTEGER
Attributes As Long
End Type
Type TOKEN_PRIVILEGES
PrivilegeCount As Long
Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Public Const TOKEN_ADJUST_PRIVILEGES = 32
Public Const TOKEN_QUERY = 8
Public Const SE_PRIVILEGE_ENABLED As Long = 2
Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Declare Function GetCurrentProcess Lib "KERNEL32" () As Long
Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Global xAnswer As Variant

'*********************************************************
'Declaração de variáveis para ocultar
'a barra de tarefas do windows
'*********************************************************
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

'********************************************************************
'Declaração de variáveis para desabilitar/habilitar
'teclas de atalho
'********************************************************************
Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

'****************************************************************
'Windows API/Global Declarations
'****************************************************************
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4

Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type

Public Declare Function WSAGetLastError Lib "wsock32" () As Long

Public Declare Function WSAStartup Lib "wsock32" _
  (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
   
Public Declare Function WSACleanup Lib "wsock32" () As Long

Public Declare Function gethostname Lib "wsock32" _
  (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "wsock32" _
  (ByVal szHost As String) As Long
   
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

#If Win16 Then
 
  Type Retang
    esquerda As Integer
    topo As Integer
    direita As Integer
    baixo As Integer
   End Type
    
   Declare Sub ClipCursor Lib "User" (lpRetang As Retang)
   Declare Sub GetWindowRect Lib "User" (ByVal hwnd _
        As Integer, lpRetang As Retang)
   Declare Function GetDesktopWindow Lib "User" () As Integer
   
#Else

    Type Retang
     esquerda As Long
     topo As Long
     direita As Long
     baixo As Long
    End Type
    
    Declare Sub ClipCursor Lib "user32" (lpRetang As Retang)
    Declare Sub GetWindowRect Lib "user32" (ByVal hwnd _
        As Integer, lpRetang As Retang)
    
    Declare Function GetDesktopWindow Lib "user32" () As Long
    
#End If

Public Const SW_SHOW As Long = 5
Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal _
        lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory _
        As String, ByVal nShowCmd As Long) As Long
     
     
Public Declare Function ShowWindow Lib "user32" _
       (ByVal hwnd As Long, _
       ByVal nCmdShow As Long) As Long
Public Const SW_HIDE As Long = 0

Public Declare Function GetCurrentProcessId Lib _
       "KERNEL32" () As Long

Public Const RSP_SIMPLE_SERVICE As Long = 1
Public Const RSP_UNREGISTER_SERVICE As Long = 0

'Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function EnumChildWindows Lib "user32" _
   (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, _
   ByVal lParam As Long) As Long
   
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
   (ByVal hwnd As Long, ByVal lpClassName As String, _
   ByVal nMaxCount As Long) As Long

Public Declare Function EnableWindow Lib "user32" _
   (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Public StartButtonhWnd As Long

'Public Declare Function ExitWindowsEx Lib "user32" _
'       (ByVal uFlags As Long, _
'       ByVal dwReserved As Long) As Long

'Public Const EWX_LOGOFF As Long = 0 'Faz Logoff do usuário.
'Public Const EWX_SHUTDOWN As Long = 1 'Desligar o computador.
'Public Const EWX_REBOOT As Long = 2 'Reiniciar o computador.
'Public Const EWX_FORCE As Long = 4 'Força a ação desejada.


'====================================================

Public Function GetIPAddress() As String

   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim HOST      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If
    
   'If gethostname(sHostName, 256) = SOCKET_ERROR Then
   '   GetIPAddress = ""
   '   MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
   '           " has occurred. Unable to successfully get Host Name."
   '   SocketsCleanup
   '   Exit Function
   'End If
   
   'sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
    
  'to extract the returned IP address, we have to copy
  'the HOST structure and its members
   CopyMemory HOST, lpHost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4
   
  'create an array to hold the result
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   
  'and with the array, build the actual address,
  'appending a period between members
   For i = 1 To HOST.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
  
  'the routine adds a period to the end of the
  'string, so remove it here
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function

Public Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub

Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function

Public Function HiByte(ByVal wParam As Integer) As Byte
  
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function


Public Function LoByte(ByVal wParam As Integer) As Byte

  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function

Public Sub ShutDownWindows95(ByVal pExitType As Byte)
  Dim Success As Boolean

  Success = ExitWindowsEx(pExitType, 0)

  If Success Then
    MsgBox "Shutting down Windows NOW!"
    End
  Else
    MsgBox "Function failed..."
  End If

End Sub

Public Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) _
   As Long
   Dim RetVal As Long
   Dim WinClassBuf As String * 255
   Dim WinClass As String
   
   RetVal = GetClassName(lhWnd, WinClassBuf, 255)
   WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
   If WinClass = "Button" Then
      StartButtonhWnd = lhWnd
      RetVal = EnableWindow(StartButtonhWnd, False)
      EnumChildProc = False  ' Stop looking
   Else
      EnumChildProc = True   ' Keep looking
   End If
End Function

Public Function StripNulls(OriginalStr As String) As String
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function

'Desabilita/habilita teclas
Sub DisableCtrlAltDelete(bDisabled As Boolean)
Dim X As Long
X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

'Fecha o aplicativo escolhido
Public Sub KillProgramInMemory(cProgram As String)
    Dim cPathRemoveFile As String
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim rProcess As Long
    Dim tPID As Long
    Dim tMID As Long
    Dim lExitCode As Long
    Dim hProcess As Long
    Dim cProcess As String
    cProgram = UCase$(cProgram)
    
    If Right$(cProgram, 1) = "\" Then
        cPathRemoveFile = cProgram
        FileSpec$ = "*.*"
        Exit Sub
    End If
    
    cPathRemoveFile = ""
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    
    If hSnapShot <> 0 Then
        uProcess.dwSize = Len(uProcess)
        rProcess = ProcessFirst(hSnapShot, uProcess)
        Do While rProcess
            tPID = uProcess.th32ProcessID
            tMID = uProcess.th32ModuleID
            cPathRemoveFile = RemoveChr0(uProcess.szExeFile)
            If cPathRemoveFile <> "" And UCase$(Right$(cPathRemoveFile, Len(cProgram) + 1)) = "\" + cProgram Then
                While cPathRemoveFile <> "" And Right$(cPathRemoveFile, 1) <> "\"
                cPathRemoveFile = Left$(cPathRemoveFile, Len(cPathRemoveFile) - 1)
            Wend
            cPathRemoveFile = ""
            hProcess = OpenProcess(PROCESS_TERMINATE, CLng(False), CLng(uProcess.th32ProcessID))
            If hProcess <> 0 Then
                If GetExitCodeProcess(hProcess, lExitCode) <> 0 Then
                    xAnswer = TerminateProcess(hProcess, lExitCode)
                End If
            End If
        End If
        rProcess = ProcessNext(hSnapShot, uProcess)
        Loop
        Call CloseHandle(hSnapShot)
    End If
End Sub

'Fecha o aplicativo escohido
Public Function GetShortFileName(ByVal cFileName As String) As String
    Dim cShortPath As String
    Dim nLen As Long
    cShortPath = String$(165, 0)
    nLen = GetShortPathName(cFileName, cShortPath, 164)
    GetShortFileName = Left$(cShortPath, nLen)
End Function

Public Function RemoveChr0(cString As String)
    While Right(cString, 1) = Chr$(0)
    cString = Left(cString, Len(cString) - 1)
    Wend
    RemoveChr0 = cString
End Function


