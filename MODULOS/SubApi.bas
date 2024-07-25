Attribute VB_Name = "SubWindows"
Option Explicit
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
' DECLARACION VARIABLES GLOBALES
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Global TimeToEnd As Boolean

Global WinVersion As Integer
Global SoundAvailable As Integer

Global pXCenter As Long
Global pYCenter As Long
Global pRadius As Long

'LookUp table with relative coordinates
Global LookUp(1 To 2, 1 To 36) As Long
Global lhPrinter As Long
Global lpFSHigh As Long

'Variable Para la clase
Global ftp As cFTP

Public msgIP As String

Private ofn As OPENFILENAME

Public VisibleFrame As Frame

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Const GWL_HINSTANCE = (-6)
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const HCBT_ACTIVATE = 5
Public Const WH_CBT = 5

Public Const INTERNET_CONNECTION_MODEM As Long = &H1
Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Public Const INTERNET_RAS_INSTALLED As Long = &H10
Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
Public Const INTERNET_CONNECTION_LAN As Long = &H2
Public Const INTERNET_CONNECTION_PROXY As Long = &H4
Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000

Private Const BUFFER_LEN = 256
Private Const FLAG_ICC_FORCE_CONNECTION = &H1

Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWNA = 8

Const WSADESCRIPTION_LEN = 257
Const WSASYS_STATUS_LEN = 129
Const ERROR_SUCCESS As Long = &H0

Public Const TWIPS = 1
Public Const PIXELS = 3
Public Const RES_INFO = 2
Public Const MINIMIZED = 1
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Const WF_CPU286 = &H2&
Public Const WF_CPU386 = &H4&
Public Const WF_CPU486 = &H8&
Public Const WF_STANDARD = &H10&
Public Const WF_ENHANCED = &H20&
Public Const WF_80x87 = &H400&

Public Const SM_MOUSEPRESENT = 19

Public Const GFSR_SYSTEMRESOURCES = &H0
Public Const GFSR_GDIRESOURCES = &H1
Public Const GFSR_USERRESOURCES = &H2

Public Const MF_POPUP = &H10
Public Const MF_BYPOSITION = &H400
Public Const MF_SEPARATOR = &H800

Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCAND = &H8800C6

'  flag values for uFlags parameter
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const PI As Single = 3.141593
Public Const ANSI_CHARSET As Long = 0
Public Const FF_DONTCARE As Long = 0
Public Const CLIP_LH_ANGLES As Long = &H10
Public Const CLIP_DEFAULT_PRECIS As Long = 0
Public Const OUT_TT_ONLY_PRECIS As Long = 7
Public Const PROOF_QUALITY As Long = 2
Public Const TRUETYPE_FONTTYPE As Long = &H4
Public Const DC_MAXEXTENT = 5
Public Const DC_MINEXTENT = 4
Public Const DC_PAPERNAMES = 16
Public Const DC_PAPERS = 2
Public Const DC_PAPERSIZE = 3
Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH = 260

' Constantes para las teclas y otros
Private Const KEYEVENTF_KEYUP = &H2
Private Const KEYEVENTF_EXTENDEDKEY = &H1

Private Const PBM_SETBKCOLOR  As Long = (&H2000& + 1)
Private Const PBM_SETBARCOLOR As Long = (&H400 + 9)

Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_EXPLORER = &H80000
Private Const SB_GETRECT = (MF_BYPOSITION + 10)

'Constante para pasar que indica que se abre el archivo en modo lectura
Private Const OF_READ = &H0&

Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Public Const IP_NO_RESOURCES = (11000 + 6)
Public Const IP_BAD_OPTION = (11000 + 7)
Public Const IP_HW_ERROR = (11000 + 8)
Public Const IP_PACKET_TOO_BIG = (11000 + 9)
Public Const IP_REQ_TIMED_OUT = (11000 + 10)
Public Const IP_BAD_REQ = (11000 + 11)
Public Const IP_BAD_ROUTE = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Public Const IP_PARAM_PROBLEM = (11000 + 15)
Public Const IP_SOURCE_QUENCH = (11000 + 16)
Public Const IP_OPTION_TOO_BIG = (11000 + 17)
Public Const IP_BAD_DESTINATION = (11000 + 18)
Public Const IP_ADDR_DELETED = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Public Const IP_MTU_CHANGE = (11000 + 21)
Public Const IP_UNLOAD = (11000 + 22)
Public Const IP_ADDR_ADDED = (11000 + 23)
Public Const IP_GENERAL_FAILURE = (11000 + 50)
Public Const MAX_IP_STATUS = 11000 + 50
Public Const IP_PENDING = (11000 + 255)
Public Const PING_TIMEOUT = 200 'default was 200 /400
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128

'Constantes
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40

'-------------------------------------
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204

' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

'
Private Const NCBASTAT As Long = &H33
Private Const NCBNAMSZ As Long = 16
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const NCBRESET As Long = &H32

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
' TIPOS
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Type MYVERSION
     lMajorVersion As Long
     lMinorVersion As Long
     lExtraInfo As Long
End Type

Type OSVERSIONINFO
     dwOSVersionInfoSize As Long
     dwMajorVersion As Long
     dwMinorVersion As Long
     dwBuildNumber As Long
     dwPlatformId As Long
     szCSDVersion As String * 128    '  Cadena de mantenimiento para uso de PSS
End Type

Type DOCINFO
     pDocName As String
     pOutputFile As String
     pDatatype As String
End Type

Type RECT
     Left As Integer
     Top As Integer
     Right As Integer
     Bottom As Integer
End Type

Public Type SystemInfo
    dwOemId As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

'----------------------------------------
'Tipos para obtener la ip del Computador
'----------------------------------------
Const Max_IP = 5
Public Const MAX_COMPUTERNAME_LENGTH As Long = 31

Type IPINFO
     dwAddr      As Long
     dwIndex     As Long
     dwMask      As Long
     dwBCastAddr As Long
     dwReasmSize As Long
     unused1     As Integer
     unused2     As Integer
End Type

Type MIB_IPADDRTABLE
     dEntrys         As Long
     mIPInfo(Max_IP) As IPINFO
End Type

Type IP_Array
     mBuffer   As MIB_IPADDRTABLE
     BufferLen As Long
End Type

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS

Public Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

'-------------------------------------------------
' Estructura SHFILEOPSTRUCT o para usar con el Api
'-------------------------------------------------
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

'-------------------------------------
Public Type NOTIFYICONDATA
    cbsize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Type NET_CONTROL_BLOCK 'NCB
        ncb_command As Byte
        ncb_retcode As Byte
        ncb_lsn As Byte
        ncb_num As Byte
        ncb_buffer As Long
        ncb_length As Integer
        ncb_callname As String * NCBNAMSZ
        ncb_name As String * NCBNAMSZ
        ncb_rto As Byte
        ncb_sto As Byte
        ncb_post As Long
        ncb_lana_num As Byte
        ncb_cmd_cplt As Byte
        ncb_reserve(9) As Byte 'Reserved, must be 0
        ncb_event As Long
End Type
 
Private Type ADAPTER_STATUS
        adapter_address(5) As Byte
        rev_major As Byte
        reserved0 As Byte
        adapter_type As Byte
        rev_minor As Byte
        duration As Integer
        frmr_recv As Integer
        frmr_xmit As Integer
        iframe_recv_err As Integer
        xmit_aborts As Integer
        xmit_success As Long
        recv_success As Long
        iframe_xmit_err As Integer
        recv_buff_unavail As Integer
        t1_timeouts As Integer
        ti_timeouts As Integer
        Reserved1 As Long
        free_ncbs As Integer
        max_cfg_ncbs As Integer
        max_ncbs As Integer
        xmit_buf_unavail As Integer
        max_dgram_size As Integer
        pending_sess As Integer
        max_cfg_sess As Integer
        max_sess As Integer
        max_sess_pkt_size As Integer
        name_count As Integer
End Type
 
Private Type NAME_BUFFER
        Name As String * NCBNAMSZ
        name_num As Integer
        name_flags As Integer
End Type
 
Private Type ASTAT
        adapt As ADAPTER_STATUS
        NameBuff(30) As NAME_BUFFER
End Type
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: URLMON
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
                                                                                    ByVal szURL As String, _
                                                                                    ByVal szFileName As String, _
                                                                                    ByVal dwReserved As Long, _
                                                                                    ByVal lpfnCB As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: NETAPI32
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function Netbios Lib "netapi32" (pncb As NET_CONTROL_BLOCK) As Byte
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: KERNEL32
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                          lpKeyName As Any, _
                                                                                          ByVal lpDefault As String, _
                                                                                          ByVal lpRetunedString As String, _
                                                                                          ByVal nSize As Long, _
                                                                                          ByVal lpFileName As String) As Long
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, _
                                                                            lpKeyName As Any, _
                                                                            ByVal lpDefault As String, _
                                                                            ByVal lpReturnedString As String, _
                                                                            ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                                                                ByVal nSize As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SystemInfo)
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
                                                                ByVal dwDesiredAccess As Long, _
                                                                ByVal dwShareMode As Long, _
                                                                lpSecurityAttributes As Any, _
                                                                ByVal dwCreationDisposition As Long, _
                                                                ByVal dwFlagsAndAttributes As Long, _
                                                                ByVal hTemplateFile As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, _
                                           ByVal lpBuffer As String, _
                                           ByVal nNumberOfBytesToWrite As Long, _
                                           lpNumberOfBytesWritten As Long, _
                                           lpOverlapped As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
' Api lOpen para abrir un archivo
Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, _
                                                      ByVal iReadWrite As Long) As Long
' Api lclose para cerrar el archivo
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long

' Api GetFileSize para averiguar el tamaño
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

'Funciones de unidades de red
Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
    
''' Maps a character string to a UTF-16 (wide character) string
Public Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long _
    ) As Long
    
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: USER32
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                    ByVal nIndex As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                                                          ByVal lpfn As Long, _
                                                                          ByVal hmod As Long, _
                                                                          ByVal dwThreadId As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, _
                                                        ByVal lHelpFile As String, _
                                                        ByVal wCommand As Long, _
                                                        ByVal dwData As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, _
                                              ByVal wFlags As Long, _
                                              ByVal X As Long, _
                                              ByVal y As Long, _
                                              ByVal nReserved As Long, _
                                              ByVal hwnd As Long, _
                                              lpReserved As Any) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, _
                                          ByVal nPos As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                         ByVal hdc As Long) As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                       ByVal hWndInsertAfter As Long, _
                                       ByVal X As Long, _
                                       ByVal y As Long, _
                                       ByVal cx As Long, _
                                       ByVal cy As Long, _
                                       ByVal wFlags As Long)
                                       
Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal CadenaAConvertir As String, _
                                                            ByVal CadenaConvertida As String) As Long
Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal CadenaAConvertir As String, _
                                                            ByVal CadenaConvertida As String) As Long
'Declaración del Api keybd_event para la presión de tecla
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
                                              ByVal bScan As Byte, _
                                              ByVal dwFlags As Long, _
                                              ByVal dwExtraInfo As Long)
                                                            
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                         ByVal hWndNewParent As Long) As Long
                                                 
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                ByVal wMsg As Long, _
                                                                ByVal wParam As Long, _
                                                                lParam As Any) As Long
'Establece la región
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                                            ByVal hRgn As Long, _
                                            ByVal bRedraw As Boolean) As Long

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: GDI32
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal U As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal cp As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
' Crea la región
Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, _
                                        ByVal x1 As Long, _
                                        ByVal y1 As Long, _
                                        ByVal X2 As Long, _
                                        ByVal Y2 As Long, _
                                        ByVal X3 As Long, _
                                        ByVal Y3 As Long) As Long
                                        
' Crea la región
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, _
                                                 ByVal y1 As Long, _
                                                 ByVal X2 As Long, _
                                                 ByVal Y2 As Long, _
                                                 ByVal X3 As Long, _
                                                 ByVal Y3 As Long) As Long
                                                                                                  
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: SHELL32
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                       ByVal lpOperation As String, _
                                                                       ByVal lpFile As String, _
                                                                       ByVal lpParameters As String, _
                                                                       ByVal lpDirectory As String, _
                                                                       ByVal nShowCmd As Long) As Long

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'Declaración Api SHFileOperation
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: WININET
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: WINMM
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: WINSPOOL (Impresoras)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, lpDevMode As Any) As Long
Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: WRITE/READ Puerto Paralelo
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Sub erik Lib "port1.dll" (ByVal p As Integer, ByVal D As Integer)
Declare Function cochex Lib "port1.dll" (ByVal p As Integer, ByVal D As Integer) As Integer
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: COMDLG32 Abri/Guardar File
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, _
                                                                        ByVal lpszTitle As String, _
                                                                        ByVal cbBuf As Integer) As Integer
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: COMCTL32 Abri/Guardar File
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Sub InitCommonControls Lib "Comctl32" ()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: IPHlpApi para obtener IP
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: ICMP para Ping de IP
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Public Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
   
Public Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: WSOCK32
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function gethostbyname Lib "WSOCK32" (ByVal Hostname As String) As Long
'Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones API: IPHLPAPI.DLL
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Private Declare Function GetRTTAndHopCount Lib "iphlpapi.dll" (ByVal lDestIPAddr As Long, _
'                                                               ByRef lHopCount As Long, _
'                                                               ByVal lMaxHops As Long, _
'                                                               ByRef lRTT As Long) As Long
    
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Funciones propias del sistema
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function pvGetFileTitle(sFileName As String) As String
    Dim Buffer As String
    Buffer = String(255, 0)
    GetFileTitle sFileName, Buffer, Len(Buffer)
    Buffer = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
    pvGetFileTitle = Buffer
End Function


Public Function DeviceColors(hdc As Long) As Single
Const PLANES = 14
Const BITSPIXEL = 12
    DeviceColors = 2 ^ (GetDeviceCaps(hdc, PLANES) * GetDeviceCaps(hdc, BITSPIXEL))
End Function

Public Function GetSysIni(section, key)
Dim RetVal As String, AppName As String, worked As Integer
    RetVal = String$(255, 0)
    worked = GetPrivateProfileString(section, key, "", RetVal, Len(RetVal), "System.ini")
    If worked = 0 Then
        GetSysIni = "desconocido"
    Else
        GetSysIni = LeftStrg(RetVal, InStr(RetVal, Chr(0)) - 1)
    End If
End Function

Public Function GetWinIni(section, key)
Dim RetVal As String, AppName As String, worked As Integer
    RetVal = String$(255, 0)
    worked = GetProfileString(section, key, "", RetVal, Len(RetVal))
    If worked = 0 Then
        GetWinIni = "desconocido"
    Else
        GetWinIni = LeftStrg(RetVal, InStr(RetVal, Chr(0)) - 1)
    End If
End Function

Public Function SystemDirectory() As String
Dim WinPath As String
    WinPath = String$(145, Chr(0))
    SystemDirectory = LeftStrg(WinPath, GetSystemDirectory(WinPath, 145))
End Function

Public Function WindowsDirectory() As String
Dim WinPath As String
    WinPath = String$(145, Chr(0))
    WindowsDirectory = LeftStrg(WinPath, GetWindowsDirectory(WinPath, 145))
End Function

Public Function WindowsVersion() As MYVERSION
Dim myOS As OSVERSIONINFO, WinVer As MYVERSION
Dim lResult As Long
    myOS.dwOSVersionInfoSize = Len(myOS) 'debe ser 148
    lResult = GetVersionEx(myOS)
   'Rellena el tipo de usuario con la información pertinente
    WinVer.lMajorVersion = myOS.dwMajorVersion
    WinVer.lMinorVersion = myOS.dwMinorVersion
    WinVer.lExtraInfo = myOS.dwPlatformId
    WindowsVersion = WinVer
End Function

Public Sub AdjustScrollBars(FrmTarget As Form)
Dim sHeight As Single, sWidth As Single
Dim objCount As Object
Dim scrHScroll As Control, scrVScroll As Control
For Each objCount In FrmTarget.Controls
    If TypeName(objCount) = "HScrollBar" Then
       Set scrHScroll = objCount
       If scrHScroll.Visible = True Then sHeight = scrHScroll.Height
    ElseIf TypeName(objCount) = "VScrollBar" Then
       Set scrVScroll = objCount
       If scrVScroll.Visible = True Then sWidth = scrVScroll.width
    End If
Next objCount
If Not IsEmpty(scrHScroll) Then
   scrHScroll.Top = FrmTarget.ScaleHeight - sHeight
   scrHScroll.width = FrmTarget.ScaleWidth - sWidth
End If
If Not IsEmpty(scrVScroll) Then
   scrVScroll.Left = FrmTarget.ScaleWidth - sWidth
   scrVScroll.Height = FrmTarget.ScaleHeight - sHeight
End If
End Sub

Public Function TextoWindowsADos(ByVal Cadena As String) As String
Dim strBuffer As String
Dim Resultado As Long
   strBuffer = String$(Len(Cadena), " ")
   Resultado = CharToOem(Cadena, strBuffer)
   TextoWindowsADos = strBuffer
End Function

Public Function TextoDosAWindows(ByVal Cadena As String) As String
Dim strBuffer As String
Dim Resultado As Long
   strBuffer = String$(Len(Cadena), " ")
   Resultado = OemToChar(Cadena, strBuffer)
   TextoDosAWindows = strBuffer
End Function

Public Sub Impresora_Rollo(PrinterBuffers As String)
'"LPT1"
Dim lReturn As Long
Dim lpcWritten As Long
Dim sWrittenData As String
Dim lDoc As Long
Dim MyDocInfo As DOCINFO
Dim PosCrLf As Long
Dim TextoImprimir As String
  'MsgBox Printer.DeviceName & ":" & vbCrLf & PrinterBuffers
   lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
   If lReturn = 0 Then
      MsgBox "The Printer Name you typed wasn't recognized."
      Exit Sub
   End If
   MyDocInfo.pDocName = "AAAAAA"
   MyDocInfo.pOutputFile = vbNullString
   MyDocInfo.pDatatype = vbNullString
   lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
   Call StartPagePrinter(lhPrinter)
  'Empieza la impresion linea a linea
   TextoImprimir = PrinterBuffers
  'Retiramos letras incorrectas
   TextoImprimir = Sin_Signos_Especiales(TextoImprimir)
   Do While Len(TextoImprimir) <> 0
      PosCrLf = InStr(TextoImprimir, vbCrLf)
      sWrittenData = TextoWindowsADos(MidStrg(TextoImprimir, 1, PosCrLf) & vbCrLf)
      If sWrittenData = "" Then sWrittenData = " "
      lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, Len(sWrittenData), lpcWritten)
      TextoImprimir = MidStrg(TextoImprimir, PosCrLf + 2, Len(TextoImprimir))
      'MsgBox "================>" & vbCrLf & sWrittenData & "================>" & vbCrLf & TextoImprimir
   Loop
  'Terminar la impresion y cerrar los documentos abiertos
   lReturn = EndPagePrinter(lhPrinter)
   lReturn = EndDocPrinter(lhPrinter)
   lReturn = ClosePrinter(lhPrinter)
End Sub

Public Function WritePrinterText(lhPrinter As Long, lpcWritten As Long, PrinterBuffers As String, Optional Centrar As Boolean, Optional Ancho_Pag As Byte) As Long
   PrinterBuffers = Sin_Signos_Especiales(PrinterBuffers)
   If Centrar And Len(PrinterBuffers) < Ancho_Pag Then
      PrinterBuffers = String$((Ancho_Pag - Len(PrinterBuffers)) / 2, " ") & PrinterBuffers
   End If
   WritePrinterText = WritePrinter(lhPrinter, ByVal PrinterBuffers, Len(PrinterBuffers), lpcWritten)
End Function

'Muestra el cuadro de dialogo para abrir archivos:
Public Function Abrir_Archivo(hwnd As Long, Dialogo As Directorio_Dialogo, TipoOperacion As Integer) As String
    On Local Error Resume Next
    Dim ofn As OPENFILENAME
    Dim A As Long
    Dim PosDir As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hwnd
    ofn.hInstance = App.hInstance
    If RightStrg(Dialogo.Filter, 1) <> "|" Then Dialogo.Filter = Dialogo.Filter + "|"
'''    For A = 1 To Len(Dialogo.Filter)
'''        If MidStrg(Dialogo.Filter, A, 1) = "|" Then MidStrg(Dialogo.Filter, A, 1) = Chr(0)
'''    Next
        Dialogo.Title = "Abrir documento"
        Dialogo.FilterIndex = 0
        Dialogo.FinDir = Dialogo.InitDir
        ofn.lpstrFilter = Dialogo.Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = Dialogo.InitDir
        If Not Dialogo.Filename = vbNullString Then ofn.lpstrFile = Dialogo.Filename & Space$(254 - Len(Dialogo.Filename))
        ofn.nFilterIndex = Dialogo.FilterIndex
        ofn.lpstrTitle = Dialogo.Title
        ofn.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        If TipoOperacion = OpenFile Then A = GetOpenFileName(ofn) Else A = GetSaveFileName(ofn)
        If A Then
             Abrir_Archivo = TrimStrg(ofn.lpstrFile)
'''             If Asc(MidStrg(Abrir_Archivo, Len(Abrir_Archivo), 1)) = 0 Then
'''                Abrir_Archivo = MidStrg(Abrir_Archivo, 1, Len(Abrir_Archivo) - 1)
'''             End If
             Dialogo.File = Abrir_Archivo
             'If VBA.RightStrg(VBA.TrimStrg(Abrir_Archivo), 1) = Chr(0) Then Abrir_Archivo = VBA.LeftStrg(VBA.TrimStrg(ofn.lpstrFile), Len(VBA.TrimStrg(ofn.lpstrFile)) - 1)
             PosDir = Len(Dialogo.File)
             While PosDir > 1
                If MidStrg(Dialogo.File, PosDir, 1) = "\" Then
                   Dialogo.FinDir = MidStrg(Dialogo.File, 1, PosDir)
                   Dialogo.File = MidStrg(Dialogo.File, PosDir + 1, TrimStrg(Len(Dialogo.File)))
                   PosDir = 1
                End If
                PosDir = PosDir - 1
             Wend
        Else
             Abrir_Archivo = ""
             Dialogo.File = Abrir_Archivo
        End If
End Function

'Muestra si existe la carpeta
Public Function DirectoryFileExists(sSource As String) As Boolean

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   
   hFile = FindFirstFile(sSource, WFD)
   DirectoryFileExists = hFile <> INVALID_HANDLE_VALUE
   
   Call FindClose(hFile)

End Function

'Muestra el cuadro de dialogo para guardar archivos:
Public Function Guardar_Archivo(hwnd As Long, Dialogo As Directorio_Dialogo) As String
    On Local Error Resume Next
    Dim ofn As OPENFILENAME
    Dim A As Long
    Dim PosDir As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hwnd
    ofn.hInstance = App.hInstance
    If RightStrg(Dialogo.Filter, 1) <> "|" Then Dialogo.Filter = Dialogo.Filter + "|"
'''    For A = 1 To Len(Dialogo.Filter)
'''        If MidStrg(Dialogo.Filter, A, 1) = "|" Then MidStrg(Dialogo.Filter, A, 1) = Chr(0)
'''    Next
        Dialogo.Title = "Guardar documento"
        Dialogo.FilterIndex = 2
        Dialogo.FinDir = Dialogo.InitDir
        ofn.lpstrFilter = Dialogo.Filter
        ofn.lpstrFile = Space(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = Dialogo.InitDir
        If Not Dialogo.Filename = vbNullString Then ofn.lpstrFile = Dialogo.Filename & Space(254 - Len(Dialogo.Filename))
        ofn.nFilterIndex = Dialogo.FilterIndex
        ofn.lpstrTitle = Dialogo.Title
        ofn.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT Or OFN_EXPLORER
        A = GetSaveFileName(ofn)
        If A Then
             Guardar_Archivo = TrimStrg(ofn.lpstrFile)
             Dialogo.File = Guardar_Archivo
             'If VBA.RightStrg(TrimStrg(Guardar_Archivo), 1) = Chr(0) Then Guardar_Archivo = VBA.LeftStrg(TrimStrg(ofn.lpstrFile), Len(TrimStrg(ofn.lpstrFile)) - 1) & GetExtension(ofn.lpstrFilter, ofn.nFilterIndex)
             PosDir = Len(Dialogo.File)
             While PosDir > 1
                If MidStrg(Dialogo.File, PosDir, 1) = "\" Then
                   Dialogo.FinDir = MidStrg(Dialogo.File, 1, PosDir)
                   Dialogo.File = MidStrg(Dialogo.File, PosDir + 1, TrimStrg(Len(Dialogo.File)))
                   PosDir = 1
                End If
                PosDir = PosDir - 1
             Wend
        Else
             Guardar_Archivo = ""
             Dialogo.File = Guardar_Archivo
        End If
End Function

'Extrae la extension seleccionada del filtro:
Public Function GetExtension(sfilter As String, pos As Long) As String
    Dim Ext() As String
    Ext = Split(sfilter, vbNullChar)
    If pos = 1 And Ext(pos) <> "*.*" Then
        GetExtension = "." & Replace(Ext(pos), "*.", "")
        Exit Function
    End If
    If pos = 1 And Ext(pos) = "*.*" Then
        GetExtension = vbNullString
        Exit Function
    End If
    If InStr(Ext(pos + 1), "*.*") Then
       GetExtension = vbNullString
    Else
       GetExtension = "." & Replace(Ext(pos + 1), "*.", "")
    End If
End Function

Public Sub Pulsar_Tecla(Tecla As Long)
    Call keybd_event(Tecla, 0, 0, 0)
    Call keybd_event(Tecla, 0, KEYEVENTF_KEYUP, 0)
End Sub

Public Function Tamano_Archivo(FilePath As String) As Long
Dim Handle As Long
    Handle = lOpen(FilePath, OF_READ)
    Tamano_Archivo = GetFileSize(Handle, lpFSHigh) / 1024
    lclose Handle
End Function

Public Function Unidades_De_Red() As String
Dim I As Long
Dim ret As Long
Dim Unidad_Disp As String
Dim Unidad_No_Disp As String
'
Unidad_Disp = ""
Unidad_No_Disp = ""

ret = GetLogicalDrives()
If ret Then
    For I = 0 To 25
        ' Si el bit es cero, es que no existe la unidad o no está mapeada
        If (ret And 2 ^ I) = 0 Then
            ' Mostrar el nombre de la unidad disponible
            Unidad_Disp = Unidad_Disp & Chr$(I + 65) & ":" & vbCrLf
            'Combo2.AddItem Chr$(i + 65) & ":"
        Else
            ' Mostrar el nombre de la unidad ocupada
            Unidad_No_Disp = Unidad_No_Disp & Chr$(I + 65) & ":" & vbCrLf
            'Combo1.AddItem Chr$(i + 65) & ":"
        End If
    Next
    ' Mostrar la primera letra disponible
'''    If Combo2.ListCount > 0 Then
'''        Combo2.ListIndex = 0
'''    End If
'''    ' Mostrar la primera letra ocupada
'''    If Combo1.ListCount > 0 Then
'''        Combo1.ListIndex = 0
'''    End If
End If
Unidades_De_Red = Unidad_Disp & vbCrLf & Unidad_No_Disp
End Function

Public Function Unidades_De_Red_Disponibles() As String
Dim I As Long
Dim ret As Long
Dim S As String
Dim Unidades_Disponibles As String
'
Unidades_Disponibles = ""
I = 260
S = String$(I, Chr$(0))
ret = GetLogicalDriveStrings(I, S)
' Si el valor devuelto es mayor que el tamaño del buffer, es que el buffer debe ser mayor
If ret > I Then
    I = ret + 2
    S = String$(I, Chr$(0))
    ret = GetLogicalDriveStrings(I, S)
End If
'
If ret Then
    ' Quitar los caracteres extras
    S = LeftStrg(S, ret)
    Do
        I = InStr(S, Chr$(0))
        If I Then
            Unidades_Disponibles = Unidades_Disponibles & LeftStrg(S, I - 1) & vbCrLf
            'Combo1.AddItem LeftStrg(S, I - 1)
            S = MidStrg(S, I + 1)
        End If
    Loop While I
    ' Mostrar la primera letra disponible
'    If Combo1.ListCount > 0 Then
'        Combo1.ListIndex = 0
'    End If
End If
Unidades_De_Red_Disponibles = Unidades_Disponibles
End Function

Public Sub SetBackColor(objObject As Object, ByVal BackColor As Long)
    SendMessage objObject.hwnd, SB_SETBKCOLOR, 0, ByVal BackColor
End Sub
 
'Cambia el color del Value de la barra, si no se especifica el color por defecto utiliza el color Vederde
Public Sub Color_Progreso(ByVal HWND_Prog As Long, Optional ByVal color As Long = vbGreen)
    Call SendMessage(HWND_Prog, PBM_SETBARCOLOR, 0&, ByVal color)
End Sub
  
' Cambia el color del fondo del Progress, si no se especifica el color por defecto utiliza el color Rojo
Public Sub Color_Fondo(ByVal HWND_Prog As Long, _
                       Optional ByVal color As Long = vbRed)
      
    Call SendMessage(HWND_Prog, PBM_SETBKCOLOR, 0&, ByVal color)
  
End Sub

Public Sub Redondear_Cuadro(El_Form As Form, Radio As Long)

Dim Region As Long
Dim ret As Long
Dim Ancho As Long
Dim alto As Long
Dim old_Scale As Integer
    
    ' guardar la escala
    old_Scale = El_Form.ScaleMode
    
    ' cambiar la escala a pixeles
    El_Form.ScaleMode = vbPixels
    
    'Obtenemos el ancho y alto de la region del Form
    Ancho = El_Form.ScaleWidth
    alto = El_Form.ScaleHeight

    'Pasar el ancho alto del formualrio y el valor de redondeo .. es decir el radio
    Region = CreateRoundRectRgn(0, 0, Ancho, alto, Radio, Radio)

    ' Aplica la región al formulario
    ret = SetWindowRgn(El_Form.hwnd, Region, True)
    
    ' restaurar la escala
    El_Form.ScaleMode = old_Scale

End Sub

'-----------------------------------------------------------------
'Obtener IP del PC
'-----------------------------------------------------------------
Public Function ConvertAddressToString(longAddr As Long) As String
Dim myByte(3) As Byte
Dim cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(cnt)) + "."
    Next cnt
    ConvertAddressToString = LeftStrg(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Function Get_WAN_IP() As Boolean
Dim strString As String
Dim ret As Long, Tel As Long
Dim bBytes() As Byte
Dim TempList() As String
Dim TempIP As String
Dim Tempi As Long
Dim Listing As MIB_IPADDRTABLE
Dim L3 As String
Dim dwLen As Long
Dim RSeg As Boolean

On Error GoTo END1

    RSeg = True
    IP_PC.InterNet = GetNetConnectString()
    RSeg = IP_PC.InterNet
    'RSeg = Ping_PC("8.8.8.8")
    'If IP_PC.InterNet Then RSeg = Ping_PC("dns.google.com")
   'MsgBox "InterNet: " & IP_PC.InterNet & vbCrLf & RSeg
    IP_PC.Max_IP = -1
    IP_PC.IP_PC = "0.0.0.0"
   'Averiguamos el nombre del PC
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = LeftStrg(strString, dwLen)
    IP_PC.Nombre_PC = strString
    
   'Determinamos la IP del Ordenador
    GetIpAddrTable ByVal 0&, ret, True
    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
    GetIpAddrTable bBytes(0), ret, False
    CopyMemory Listing.dEntrys, bBytes(0), 4
    For Tel = 0 To Listing.dEntrys - 1
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
    Next Tel
    
    For Tempi = 0 To Listing.dEntrys - 1
       'MsgBox "TempList(" & Tempi & "): " & TempList(Tempi)
        L3 = LeftStrg(TempList(Tempi), 3)
        If L3 <> "169" And L3 <> "127" Then IP_PC.Max_IP = IP_PC.Max_IP + 1
    Next Tempi
   'MsgBox "Max_IP: " & IP_PC.Max_IP
    If IP_PC.Max_IP >= 0 Then
       ReDim IP_PC.Lista_IPs(0 To IP_PC.Max_IP) As String
       IP_PC.Max_IP = -1
       TempIP = "" ' TempList(0)
       For Tempi = 0 To Listing.dEntrys - 1
           L3 = LeftStrg(TempList(Tempi), 3)
           If L3 <> "169" And L3 <> "127" Then
              IP_PC.Max_IP = IP_PC.Max_IP + 1
              IP_PC.Lista_IPs(IP_PC.Max_IP) = TempList(Tempi)
              IP_PC.IP_PC = TempList(Tempi)
           End If
          '123456789012345
          '000.000.000.000
          ' If InStr(TempList(Tempi), ".27.") Then RSeg = False
          ' If InStr(TempList(Tempi), ".56.") Then RSeg = False
           'MsgBox "Seg: " & TempList(Tempi)
       Next Tempi
    Else
       ReDim IP_PC.Lista_IPs(0 To 0) As String
       IP_PC.Max_IP = 0
       IP_PC.Lista_IPs(IP_PC.Max_IP) = "Sin Conexion"
       IP_PC.InterNet = False
       RSeg = False
    End If
   'MsgBox "Get_WAN_IP: " & IP_PC.Conexion_WAN
    IP_PC.MAC_PC = Mi_MAC_Local
    IP_PC.WAN_PC = Mi_IP_Publica
   'MsgBox IP_PC.IP_PC
    Get_WAN_IP = RSeg
    Exit Function
END1:
    ReDim IP_PC.Lista_IPs(0 To 0) As String
    IP_PC.Max_IP = 0
    IP_PC.Lista_IPs(IP_PC.Max_IP) = "Sin Conexion"
    IP_PC.InterNet = False
   'AdoStrCnnMySQL = ""
   'MsgBox IP_PC.IP_PC
    Get_WAN_IP = RSeg
End Function

Public Function GetStatusCode(Status As Long) As String
   Select Case Status
      Case IP_SUCCESS:               msg = "ip success"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   GetStatusCode = CStr(Status)
End Function

Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Public Function Ping(szAddress As String, ECHO As ICMP_ECHO_REPLY) As Long
   Dim hPort As Long
   Dim dwAddress As Long
   Dim sDataToSend As String
   Dim iOpt As Long
   
   sDataToSend = "Echo This"
   dwAddress = AddressStringToLong(szAddress)
   
   Call SocketsInitialize
   hPort = IcmpCreateFile()
   
   If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO, Len(ECHO), PING_TIMEOUT) Then
      Ping = ECHO.RoundTripTime
   Else
      Ping = ECHO.Status * -1
   End If
  'MsgBox szAddress & vbCrLf & dwAddress & vbCrLf & ECHO.RoundTripTime & vbCrLf & ECHO.status
   Call IcmpCloseHandle(hPort)
   Call SocketsCleanup
End Function
   
Function AddressStringToLong(ByVal tmp As String) As Long


   Dim I As Integer
   Dim parts(1 To 4) As String
   
   I = 0
   While InStr(tmp, ".") > 0
      I = I + 1
      parts(I) = Mid(tmp, 1, InStr(tmp, ".") - 1)
      tmp = Mid(tmp, InStr(tmp, ".") + 1)
   Wend
   
   I = I + 1
   parts(I) = tmp
   
   If I <> 4 Then
      AddressStringToLong = 0
      Exit Function
   End If
   AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
                         Right("00" & Hex(parts(3)), 2) & _
                         Right("00" & Hex(parts(2)), 2) & _
                         Right("00" & Hex(parts(1)), 2))
   
''''
''''   Dim I As Integer
''''   Dim parts(1 To 4) As String
''''   i = 0
''''   While InStr(tmp, ".") > 0
''''      i = i + 1
''''      parts(i) = MidStrg(tmp, 1, InStr(tmp, ".") - 1)
''''      tmp = MidStrg(tmp, InStr(tmp, ".") + 1)
''''   Wend
''''   i = i + 1
''''   parts(i) = tmp
''''   If i <> 4 Then
''''      AddressStringToLong = 0
''''      Exit Function
''''   End If
''''   AddressStringToLong = Val("&H" & RightStrg("00" & Hex(parts(4)), 2) & _
''''                         RightStrg("00" & Hex(parts(3)), 2) & _
''''                         RightStrg("00" & Hex(parts(2)), 2) & _
''''                         RightStrg("00" & Hex(parts(1)), 2))
End Function

Public Function SocketsCleanup() As Boolean
    Dim X As Long
    X = WSACleanup()
    If X <> 0 Then
        MsgBox "Windows Sockets error " & TrimStrg(Str$(X)) & _
               " occurred in Cleanup.", vbExclamation
        SocketsCleanup = False
    Else
        SocketsCleanup = True
    End If
End Function

Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSADATA
    Dim X As Integer
    Dim szLoByte As String, szHiByte As String, szBuf As String
    
    X = WSAStartup(WS_VERSION_REQD, WSAD)
    
    If X <> 0 Then
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
        SocketsInitialize = False
        Exit Function
    End If
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
       (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
        HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        
        szHiByte = TrimStrg(Str$(HiByte(WSAD.wVersion)))
        szLoByte = TrimStrg(Str$(LoByte(WSAD.wVersion)))
        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        szBuf = szBuf & " is not supported by Windows " & _
                          "Sockets for 32 bit Windows environments."
        MsgBox szBuf, vbExclamation
        SocketsInitialize = False
        Exit Function
    End If
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        szBuf = "This application requires a minimum of " & _
                 TrimStrg(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox szBuf, vbExclamation
        SocketsInitialize = False
        Exit Function
    End If
    SocketsInitialize = True
End Function

Public Function Ping_PC(My_IP_PC As String) As Boolean
Dim DosPuntos As Integer
Dim Coma As Integer
Dim Resultado As Boolean
'Estructura con la información para hacer el Ping
Dim ECHO As ICMP_ECHO_REPLY
Dim TargetIP As String ' dirección ip
Dim ret As String
Dim PingIP As String

    RatonReloj
   'MsgBox IP_PC & vbCrLf & IsNumeric(Replace(IP_PC, ".", "")) & vbCrLf & recuperar_IP(IP_PC)
    Resultado = False
    If IP_PC.InterNet Then
        PingIP = My_IP_PC
        DosPuntos = InStr(PingIP, ":")
        Coma = InStr(PingIP, ",")
        If (DosPuntos + Coma) > 2 Then PingIP = TrimStrg(MidStrg(PingIP, DosPuntos + 1, Coma - DosPuntos - 1)) Else PingIP = TrimStrg(My_IP_PC)
          
        If IsNumeric(Replace(PingIP, ".", "")) Then TargetIP = PingIP Else TargetIP = recuperar_IP(PingIP)
        
       'MsgBox My_IP_PC & vbCrLf & PingIP & vbCrLf & TargetIP
        
        If Len(TargetIP) > 0 Then
            ret = "Respuesta desde " & TargetIP & ": "
            Call Ping(TargetIP, ECHO)
           'Estado del Ping
            If ECHO.Status = 0 Then
                ret = ret & "Tiempo: " & ECHO.RoundTripTime & " ms(1) "
                Resultado = True
            Else
                Call Ping(TargetIP, ECHO)
                If ECHO.Status = 0 Then
                   ret = ret & "Tiempo: " & ECHO.RoundTripTime & " ms(2) "
                   Resultado = True
                Else
                   Call Ping(TargetIP, ECHO)
                   If ECHO.Status = 0 Then
                      ret = ret & "Tiempo: " & ECHO.RoundTripTime & " ms(3) "
                      Resultado = True
                   Else
                     'Error al hacer ping
                      ret = ret & "Not successful "
    ''                  Call Ping(TargetIP, ECHO)
    ''                  If ECHO.Status = 0 Then
    ''                     ret = ret & "Tiempo: " & ECHO.RoundTripTime & " ms(4) "
    ''                     Resultado = True
    ''                  Else
    ''                  End If
                   End If
                End If
            End If
            ret = ret & "bytes=" & ECHO.DataSize & " "
            If ECHO.DataSize = 0 And ECHO.RoundTripTime = 0 Then Resultado = False
        Else
           'Error
            ret = ret & "Not successful, not exist the hosting"
        End If
        
    '    Progreso_Barra.Mensaje_Box = ret
    '    Progreso_Esperar False
       'MsgBox Progreso_Barra.Mensaje_Box
        With ECHO
             ret = ret & vbCrLf _
                 & "Address            : " & .Address & vbCrLf _
                 & "Status             : " & .Status & vbCrLf _
                 & "RoundTripTime      : " & .RoundTripTime & vbCrLf _
                 & "DataSize           : " & .DataSize & vbCrLf _
                 & "Reserved           : " & .Reserved & vbCrLf _
                 & "DataPointer        : " & .DataPointer & vbCrLf _
                 & "Data               : " & .Data & vbCrLf _
                 & "Options Flags      : " & .Options.Flags & vbCrLf _
                 & "Options OptionsData: " & .Options.OptionsData & vbCrLf _
                 & "Options OptionsSize: " & .Options.OptionsSize & vbCrLf _
                 & "Options Tos        : " & .Options.Tos & vbCrLf _
                 & "Options Ttl        : " & .Options.Ttl
        End With
       'MsgBox ret, vbInformation
        IP_PC.Status = IP_PC.Status & ret
    End If
    Ping_PC = Resultado
End Function

Public Function recuperar_IP(ByVal Nombre_Host As String) As String
Dim ErrorNo As String
Dim lHost As Long, T_Host As HOSTENT, ipDir As Long
Dim tIP() As Byte, I As Integer, sIP As String
      
    If Not Inicializar_Socket() Then
        recuperar_IP = ""
        Exit Function
    End If
    Nombre_Host = TrimStrg(Nombre_Host)
    lHost = gethostbyname(Nombre_Host)
    If lHost = 0 Then
        recuperar_IP = ""
        ErrorNo = "Error: " & Nombre_Host & ", fuera de linea o sin internet"
       'Progreso_Barra.Mensaje_Box = ErrorNo
       'Progreso_Esperar False
       'MsgBox ErrorNo
        Remover_Socket
        Control_Procesos Normal, ErrorNo, "Conexion"
        Exit Function
    End If
    CopyMemoryIP T_Host, lHost, Len(T_Host)
    CopyMemoryIP ipDir, T_Host.hAddrList, 4
    ReDim tIP(1 To T_Host.hLen)
    CopyMemoryIP tIP(1), ipDir, T_Host.hLen
     
    For I = 1 To T_Host.hLen
        sIP = sIP & tIP(I) & "."
    Next
    recuperar_IP = MidStrg(sIP, 1, Len(sIP) - 1)
      
    Remover_Socket
End Function
  
Public Function Inicializar_Socket() As Boolean
    Dim W As WSADATA, slb As String, shb As String
      
    If WSAStartup(&H101, W) <> ERROR_SUCCESS Then
        MsgBox "Error"
        Inicializar_Socket = False
        Exit Function
    End If
    If W.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "Error"
        Inicializar_Socket = False
        Exit Function
    End If
    If LB(W.wVersion) < WS_VERSION_MAJOR Or (LB(W.wVersion) = WS_VERSION_MAJOR And HB(W.wVersion) < WS_VERSION_MINOR) Then
        shb = CStr(HB(W.wVersion))
        slb = CStr(LB(W.wVersion))
        MsgBox "Error"
        Inicializar_Socket = False
        Exit Function
    End If
    Inicializar_Socket = True
End Function
  
Public Function HB(ByVal wParam As Integer)
    HB = wParam \ &H100 And &HFF&
End Function

Public Function LB(ByVal wParam As Integer)
    LB = wParam And &HFF&
End Function

Public Sub Remover_Socket()
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
End Sub

''' Return byte array with VBA "Unicode" string encoded in UTF-8
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function


''' Return length of byte array or zero if uninitialized
Private Function BytesLength(abBytes() As Byte) As Long
    ' Trap error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
End Function

''' Return VBA "Unicode" string from byte array encoded in UTF-8
Public Function Utf8BytesToString(abUtf8Array() As Byte) As String
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
    Utf8BytesToString = ""
    ' Catch uninitialized input array
    nBytes = BytesLength(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)
End Function

Public Sub Redondear_Formulario(El_Form As Form, Radio As Long)
Dim Region As Long
Dim ret As Long
Dim Ancho As Long
Dim alto As Long

'Obtenemos el ancho y alto de la region del Form
Ancho = El_Form.width / Screen.TwipsPerPixelX
alto = El_Form.Height / Screen.TwipsPerPixelY

'Le pasamos el ancho alto del formualrio y el valor de _
 redondeo es decir el radio

Region = CreateRoundRectRgn(0, 0, Ancho, alto, Radio, Radio)

' Aplica la región al formulario
ret = SetWindowRgn(El_Form.hwnd, Region, True)

End Sub

Public Function GetNetConnectString() As Boolean
Dim dwFlags As Long
Dim msg As String
Dim ConexionOK As Boolean
    msg = ""
    If InternetGetConnectedState(dwFlags, 0&) Then ConexionOK = True Else ConexionOK = False
    If dwFlags And INTERNET_CONNECTION_CONFIGURED Then msg = msg & "Network. "
    If dwFlags And INTERNET_CONNECTION_LAN Then msg = msg & "LAN. "
    If dwFlags And INTERNET_CONNECTION_PROXY Then msg = msg & "Proxy Server. "
    If dwFlags And INTERNET_CONNECTION_MODEM Then msg = msg & "Modem. "
    If dwFlags And INTERNET_CONNECTION_OFFLINE Then msg = msg & "Offline. "
    If dwFlags And INTERNET_CONNECTION_MODEM_BUSY Then msg = msg & "non-Internetconnection. "
    If dwFlags And INTERNET_RAS_INSTALLED Then msg = msg & "Remote Access Services. "
    If Not ConexionOK Then msg = msg & "No conectado a Internet"
    IP_PC.Status = TrimStrg(msg)
    GetNetConnectString = ConexionOK
End Function

''Public Function GetMACAddress(sDelimiter As String) As String
'''retrieve the MAC Address for the network controller
'''installed, returning a formatted string
''Dim tmp As String
''Dim pASTAT As Long
''Dim NCB As NET_CONTROL_BLOCK
''Dim AST As ASTAT
''Dim cnt As Long
''
'''The IBM NetBIOS 3.0 specifications defines four basic
'''NetBIOS environments under the NCBRESET command. Win32
'''follows the OS/2 Dynamic Link Routine (DLR) environment.
'''This means that the first NCB issued by an application
'''must be a NCBRESET, with the exception of NCBENUM.
'''The Windows NT implementation differs from the IBM
'''NetBIOS 3.0 specifications in the NCB_CALLNAME field.
''NCB.ncb_command = NCBRESET
''Call Netbios(NCB)
''
'''To get the Media Access Control (MAC) address for an
'''ethernet adapter programmatically, use the Netbios()
'''NCBASTAT command and provide a "*" as the name in the
'''NCB.ncb_CallName field (in a 16-chr string).
''NCB.ncb_callname = "* "
''NCB.ncb_command = NCBASTAT
''
'''For machines with multiple network adapters you need to
'''enumerate the LANA numbers and perform the NCBASTAT
'''command on each. Even when you have a single network
'''adapter, it is a good idea to enumerate valid LANA numbers
'''first and perform the NCBASTAT on one of the valid LANA
'''numbers. It is considered bad programming to hardcode the
'''LANA number to 0 (see the comments section below).
''NCB.ncb_lana_num = 0
''NCB.ncb_length = Len(AST)
''
''pASTAT = HeapAlloc(GetProcessHeap(), _
''HEAP_GENERATE_EXCEPTIONS Or _
''HEAP_ZERO_MEMORY, _
''NCB.ncb_length)
''
''If pASTAT <> 0 Then
''   NCB.ncb_buffer = pASTAT
''   Call Netbios(NCB)
''   CopyMemory AST, NCB.ncb_buffer, Len(AST)
''  'convert the byte array to a string
''   GetMACAddress = MakeMacAddress(AST.adapt.adapter_address(), sDelimiter)
''   HeapFree GetProcessHeap(), 0, pASTAT
''Else
''   Debug.Print "memory allocation failed!"
''   Exit Function
''End If
''
''End Function

Public Function MakeMacAddress(b() As Byte, sDelim As String) As String
Dim cnt As Long
Dim buff As String
On Local Error GoTo MakeMac_error
 
'so far, MAC addresses are
'exactly 6 segments in size (0-5)
If UBound(b) = 5 Then
  'concatenate the first five values
  'together and separate with the
  'delimiter char
   For cnt = 0 To 4
       buff = buff & Right$("00" & Hex(b(cnt)), 2) & sDelim
   Next
     
  'and append the last value
   buff = buff & Right$("00" & Hex(b(5)), 2)
End If 'UBound(b)
 
MakeMacAddress = buff
 
MakeMac_exit:
Exit Function
 
MakeMac_error:
MakeMacAddress = "(error building MAC address)"
Resume MakeMac_exit
End Function

Public Function Mi_MAC_Local() As String
Dim colNetAdapters, objWMIService, objitem As Object
Dim strComputer As String
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    For Each objitem In colNetAdapters
        Mi_MAC_Local = objitem.MACAddress
    Next
End Function

Public Function Mi_IP_Publica() As String
Dim cTemp As String
Dim arTemp() As String
Dim URL As String
Dim IP As String
Dim PosIpI As Long
Dim PosIpF As Long
IP = RutaSysBases & "\TEMP\IP_" & CodigoUsuario & ".txt"
cTemp = "NO IP WAN"
'URL = "http://miip.es"
URL = "https://whatismyipaddress.com/"
'URL = "http://myip.es" 'AUN NO FUNCIONA
If Dir(IP) <> "" Then Kill IP
Call URLDownloadToFile(0, URL, IP, 0, 0)
If Dir(IP) <> "" Then
   cTemp = CreateObject("Scripting.FileSystemObject").OpenTextFile(IP).ReadAll
   cTemp = Replace(cTemp, """", "'")
   PosIpI = InStr(cTemp, "https://whatismyipaddress.com/ip/")
    If PosIpI > 0 Then
       cTemp = MidStrg(cTemp, PosIpI, Len(cTemp))
       PosIpF = InStr(cTemp, "'")
       cTemp = MidStrg(cTemp, 1, PosIpF - 1)
       cTemp = Replace(cTemp, "https://whatismyipaddress.com/ip/", "")
       cTemp = TrimStrg(MidStrg(cTemp, 1, 15))
    End If
    Kill IP
End If
Mi_IP_Publica = cTemp
End Function

Public Function GetUrlSource(sURL As String) As String
Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
Dim hInternet As Long, hSession As Long, lReturn As Long

   'get the handle of the current internet connection
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
   'get the handle of the url
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
   'if we have the handle, then start reading the web page
    If hInternet Then
       'get the first chunk & buffer it.
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
       'if there's more data then keep reading it into the buffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
 
   'close the URL
    iResult = InternetCloseHandle(hInternet)
    sData = Trim(Replace(sData, Chr(0), ""))
    GetUrlSource = sData
End Function

