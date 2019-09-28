VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Winsock Initialization and termination
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADATA) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long

'Server side Winsock API functions
Private Declare Function WSABind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Private Declare Function WSAListen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function WSAAccept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long

'String functions
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'Socket Functions
Private Declare Function WSAConnect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function WSACloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

'Data transfer functions
Private Declare Function WSARecv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long
Private Declare Function WSASend Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long

'Network byte ordering functions
Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Private Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'End point information
Private Declare Function GetSockName Lib "ws2_32.dll" Alias "getsockname" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Private Declare Function GetPeerName Lib "ws2_32.dll" Alias "getpeername" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long

'Hostname resolving functions
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

'ICMP functions
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long

'Winsock API functions for resolving hostnames and IP's
Private Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal BufLen As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
'Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
'Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Private Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long

'Memory copy and move functions
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)

'Window creation and destruction functions
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

'Messaging functions
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

'Memory allocation functions
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

'..
Private Const WSADescription_Len = 257
Private Const WSASYS_Status_Len = 129

'Maximum queue length specifiable by listen.
Private Const SOMAXCONN = &H7FFFFFFF

'Windows Socket types
Private Const SOCK_STREAM = 1     'Stream socket

''Address family
'Private Const AF_INET = 2          'Internetwork: UDP, TCP, etc.
'
''Socket Protocol
'Private Const IPPROTO_TCP = 6     'tcp
Public Enum AddressFamily
    '
    AF_UNSPEC = 0      '/* unspecified */
'/*
' * Although  AF_UNSPEC  is  defined for backwards compatibility, using
' * AF_UNSPEC for the "af" parameter when creating a socket is STRONGLY
' * DISCOURAGED.    The  interpretation  of  the  "protocol"  parameter
' * depends  on the actual address family chosen.  As environments grow
' * to  include  more  and  more  address families that use overlapping
' * protocol  values  there  is  more  and  more  chance of choosing an
' * undesired address family when AF_UNSPEC is used.
' */
    AF_UNIX = 1        '/* local to host (pipes, portals) */
    AF_INET = 2        '/* internetwork: UDP, TCP, etc. */
    AF_IMPLINK = 3     '/* arpanet imp addresses */
    AF_PUP = 4         '/* pup protocols: e.g. BSP */
    AF_CHAOS = 5       '/* mit CHAOS protocols */
    AF_NS = 6          '/* XEROX NS protocols */
    AF_IPX = AF_NS     '/* IPX protocols: IPX, SPX, etc. */
    AF_ISO = 7         '/* ISO protocols */
    AF_OSI = AF_ISO    '/* OSI is ISO */
    AF_ECMA = 8        '/* european computer manufacturers */
    AF_DATAKIT = 9     '/* datakit protocols */
    AF_CCITT = 10      '/* CCITT protocols, X.25 etc */
    AF_SNA = 11        '/* IBM SNA */
    AF_DECnet = 12     '/* DECnet */
    AF_DLI = 13        '/* Direct data link interface */
    AF_LAT = 14        '/* LAT */
    AF_HYLINK = 15     '/* NSC Hyperchannel */
    AF_APPLETALK = 16  '/* AppleTalk */
    AF_NETBIOS = 17    '/* NetBios-style addresses */
    AF_VOICEVIEW = 18  '/* VoiceView */
    AF_FIREFOX = 19    '/* Protocols from Firefox */
    AF_UNKNOWN1 = 20   '/* Somebody is using this! */
    AF_BAN = 21        '/* Banyan */
    AF_ATM = 22        '/* Native ATM Services */
    AF_INET6 = 23      '/* Internetwork Version 6 */
    AF_CLUSTER = 24    '/* Microsoft Wolfpack */
    AF_12844 = 25      '/* IEEE 1284.4 WG AF */
    AF_MAX = 26
    '
End Enum
'
'/*
' * Protocols
' */
Public Enum SocketProtocol
    IPPROTO_IP = 0             '/* dummy for IP */
    IPPROTO_ICMP = 1           '/* control message protocol */
    IPPROTO_IGMP = 2           '/* internet group management protocol */
    IPPROTO_GGP = 3            '/* gateway^2 (deprecated) */
    IPPROTO_TCP = 6            '/* tcp */
    IPPROTO_PUP = 12           '/* pup */
    IPPROTO_UDP = 17           '/* user datagram protocol */
    IPPROTO_IDP = 22           '/* xns idp */
    IPPROTO_ND = 77            '/* UNOFFICIAL net disk proto */
    IPPROTO_RAW = 255          '/* raw IP packet */
    IPPROTO_MAX = 256
End Enum

'Data type conversion constants
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

'Fixed memory flag for GlobalAlloc
Private Const GMEM_FIXED = &H0

'Winsock error offset
Private Const WSABASEERR = 10000

' Other constants
Private Const ERROR_SUCCESS              As Long = 0
Private Const WS_VERSION_REQD            As Long = &H101
Private Const WS_VERSION_MAJOR           As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR           As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD           As Long = 1
Private Const DATA_SIZE = 32
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128

Private Enum IP_STATUS
    IP_STATUS_BASE = 11000
    IP_SUCCESS = 0
    IP_BUF_TOO_SMALL = (11000 + 1)
    IP_DEST_NET_UNREACHABLE = (11000 + 2)
    IP_DEST_HOST_UNREACHABLE = (11000 + 3)
    IP_DEST_PROT_UNREACHABLE = (11000 + 4)
    IP_DEST_PORT_UNREACHABLE = (11000 + 5)
    IP_NO_RESOURCES = (11000 + 6)
    IP_BAD_OPTION = (11000 + 7)
    IP_HW_ERROR = (11000 + 8)
    IP_PACKET_TOO_BIG = (11000 + 9)
    IP_REQ_TIMED_OUT = (11000 + 10)
    IP_BAD_REQ = (11000 + 11)
    IP_BAD_ROUTE = (11000 + 12)
    IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
    IP_TTL_EXPIRED_REASSEM = (11000 + 14)
    IP_PARAM_PROBLEM = (11000 + 15)
    IP_SOURCE_QUENCH = (11000 + 16)
    IP_OPTION_TOO_BIG = (11000 + 17)
    IP_BAD_DESTINATION = (11000 + 18)
    IP_ADDR_DELETED = (11000 + 19)
    IP_SPEC_MTU_CHANGE = (11000 + 20)
    IP_MTU_CHANGE = (11000 + 21)
    IP_UNLOAD = (11000 + 22)
    IP_ADDR_ADDED = (11000 + 23)
    IP_GENERAL_FAILURE = (11000 + 50)
    MAX_IP_STATUS = 11000 + 50
    IP_PENDING = (11000 + 255)
    PING_TIMEOUT = 255
End Enum


'Winsock messages that will go to the window handler
Private Enum WSAMessage
    FD_READ = &H1&      'Data is ready to be read from the buffer
    FD_CONNECT = &H10&  'Connection esatblished
    FD_CLOSE = &H20&    'Connection closed
    FD_ACCEPT = &H8&    'Connection request pending
    FD_WRITE = &H2& 'Data written
End Enum

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long
  ' Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type


'Winsock Data structure
Private Type WSADATA
    wVersion       As Integer                       'Version
    wHighVersion   As Integer                       'High Version
    szDescription  As String * WSADescription_Len   'Description
    szSystemStatus As String * WSASYS_Status_Len    'Status of system
    iMaxSockets    As Integer                       'Maximum number of sockets allowed
    iMaxUdpDg      As Integer                       'Maximum UDP datagrams
    lpVendorInfo   As Long                          'Vendor Info
End Type

'HostEnt Structure
Private Type HOSTENT
    hName     As Long       'Host Name
    hAliases  As Long       'Alias
    hAddrType As Integer    'Address Type
    hLength   As Integer    'Length
    hAddrList As Long       'Address List
End Type

'Socket Address structure
Private Type SOCKADDR_IN
    sin_family       As Integer 'Address familly
    sin_port         As Integer 'Port
    sin_addr         As Long    'Long address
    sin_zero(1 To 8) As Byte
End Type

'End Point of connection information
Public Enum IPEndPointFields
    LOCAL_HOST          'Local hostname
    LOCAL_HOST_IP       'Local IP
    LOCAL_PORT          'Local port
    REMOTE_HOST         'Remote hostname
    REMOTE_HOST_IP      'Remote IP
    REMOTE_PORT         'Remote port
End Enum

'Basic Winsock error results.
Private Enum WSABaseErrors
    INADDR_NONE = &HFFFF
    SOCKET_ERROR = -1
    INVALID_SOCKET = -1
End Enum

'Winsock error constants
Private Enum WSAErrorConstants
'Windows Sockets definitions of regular Microsoft C error constants
    WSAEINTR = (WSABASEERR + 4)
    WSAEBADF = (WSABASEERR + 9)
    WSAEACCES = (WSABASEERR + 13)
    WSAEFAULT = (WSABASEERR + 14)
    WSAEINVAL = (WSABASEERR + 22)
    WSAEMFILE = (WSABASEERR + 24)
'Windows Sockets definitions of regular Berkeley error constants
    WSAEWOULDBLOCK = (WSABASEERR + 35)
    WSAEINPROGRESS = (WSABASEERR + 36)
    WSAEALREADY = (WSABASEERR + 37)
    WSAENOTSOCK = (WSABASEERR + 38)
    WSAEDESTADDRREQ = (WSABASEERR + 39)
    WSAEMSGSIZE = (WSABASEERR + 40)
    WSAEPROTOTYPE = (WSABASEERR + 41)
    WSAENOPROTOOPT = (WSABASEERR + 42)
    WSAEPROTONOSUPPORT = (WSABASEERR + 43)
    WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
    WSAEOPNOTSUPP = (WSABASEERR + 45)
    WSAEPFNOSUPPORT = (WSABASEERR + 46)
    WSAEAFNOSUPPORT = (WSABASEERR + 47)
    WSAEADDRINUSE = (WSABASEERR + 48)
    WSAEADDRNOTAVAIL = (WSABASEERR + 49)
    WSAENETDOWN = (WSABASEERR + 50)
    WSAENETUNREACH = (WSABASEERR + 51)
    WSAENETRESET = (WSABASEERR + 52)
    WSAECONNABORTED = (WSABASEERR + 53)
    WSAECONNRESET = (WSABASEERR + 54)
    WSAENOBUFS = (WSABASEERR + 55)
    WSAEISCONN = (WSABASEERR + 56)
    WSAENOTCONN = (WSABASEERR + 57)
    WSAESHUTDOWN = (WSABASEERR + 58)
    WSAETOOMANYREFS = (WSABASEERR + 59)
    WSAETIMEDOUT = (WSABASEERR + 60)
    WSAECONNREFUSED = (WSABASEERR + 61)
    WSAELOOP = (WSABASEERR + 62)
    WSAENAMETOOLONG = (WSABASEERR + 63)
    WSAEHOSTDOWN = (WSABASEERR + 64)
    WSAEHOSTUNREACH = (WSABASEERR + 65)
    WSAENOTEMPTY = (WSABASEERR + 66)
    WSAEPROCLIM = (WSABASEERR + 67)
    WSAEUSERS = (WSABASEERR + 68)
    WSAEDQUOT = (WSABASEERR + 69)
    WSAESTALE = (WSABASEERR + 70)
    WSAEREMOTE = (WSABASEERR + 71)
'Extended Windows Sockets error constant definitions
    WSASYSNOTREADY = (WSABASEERR + 91)
    WSAVERNOTSUPPORTED = (WSABASEERR + 92)
    WSANOTINITIALISED = (WSABASEERR + 93)
    WSAEDISCON = (WSABASEERR + 101)
    WSAENOMORE = (WSABASEERR + 102)
    WSAECANCELLED = (WSABASEERR + 103)
    WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
    WSAEINVALIDPROVIDER = (WSABASEERR + 105)
    WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
    WSASYSCALLFAILURE = (WSABASEERR + 107)
    WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
    WSATYPE_NOT_FOUND = (WSABASEERR + 109)
    WSA_E_NO_MORE = (WSABASEERR + 110)
    WSA_E_CANCELLED = (WSABASEERR + 111)
    WSAEREFUSED = (WSABASEERR + 112)
    WSAHOST_NOT_FOUND = 11001
    WSATRY_AGAIN = 11002
    WSANO_RECOVERY = 11003
    WSANO_DATA = 11004
    FD_SETSIZE = 64
End Enum
Private Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
Private Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dst As Any, ByVal iLen&)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_NEUTRAL = &H0
Private Const SUBLANG_DEFAULT = &H1
Private Const ERROR_BAD_USERNAME = 2202&
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'################# Class Variables ########################################

'Some variables that are used to store the state of this object.
Private m_intConnectionPort As Integer    'Just remember the port that we connected to
Private m_lngConnectionSocket As Long     'Just remember the socket where the connection is set up for
Private Type ConnectionInfo
    lngSocket As Long
    strRemoteHost As String
    strRemoteIP As String
    lngRemotePort As Long
End Type

Private m_Connections() As ConnectionInfo 'Long  ' As Connection  'When we are connected to the server this object will be created for handeling this connection.
Private m_ConnectionCount As Long

'We need a memory block for resolving the hostname. these variables will be used for handeling those.
Private lngMemoryHandle    As Long      'Handle of the allocated memory block object
Private lngMemoryPointer   As Long      'Address of the memory block

'We are going to subclass a generated window and add 2 new message numbers for the communication
Private objSuperClass As cSubclass               'This is the subclass handler
Implements iSubclass                   'This is the interface to the subclass handler


Private lngWinsockMessage As Long            'General Winsock message
Private lngResolveHostMessage As Long   'Winsock resolve host message
Private lngWindowHandle As Long              'Window handle
Private m_lngSocket As Long  'The socket number where we are connected to
Private Buffer() As Byte
Private BufferLength As Long
Private SentLength As Long
Private m_Listening As Boolean
Public LastAcceptedSocket As Long

'Purpose: This event will be called when the connection is closed by the server or the client itself.
Public Event OnClose(ByVal lngSocket As Long)
'Purpose: This event will be called when the connect method executed successfully.
Public Event ConnectionRequest(ByVal lngSocket As Long)
'Purpose: This event will be called when the connect method executed successfully.
Public Event OnConnect(ByVal lngSocket As Long)
'Purpose: This event will be called when data arived. After this you have to call the .Connection.Recv method to get the data.
Public Event DataArrival(ByVal lngSocket As Long)
Public Event SendProgress(ByVal lngSocket As Long, ByVal bytesSent As Long, ByVal bytesRemaining As Long)

'Purpose: This event will be called whenever there is a problem.
Public Event OnError(ByVal lngRetCode As Long, ByVal strDescription As String)


'Purpose:
' The moment an instance of this object is created we can set up a new socket for communication.
Private Sub Class_Initialize()
On Error GoTo ErrorHandler
Dim WSAInfo As WSADATA
    
   'Create a blank, invisible window
            lngWindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

   'Create a unique message number for our messages
            lngWinsockMessage = RegisterWindowMessage(App.EXEName & ".lngWinsockMessage")
            lngResolveHostMessage = RegisterWindowMessage(App.EXEName & ".ResolveHostMessage")
    
   'Start up the subclassing
            Set objSuperClass = New cSubclass        'Create a cSuperClass instance
            With objSuperClass
              .Subclass lngWindowHandle, Me
              .AddMsg lngWindowHandle, lngWinsockMessage, MSG_AFTER
               .AddMsg lngWindowHandle, lngResolveHostMessage, MSG_AFTER
            End With
     
   'Start version 1.1 of the Microsoft Winsock API
            WSAStartup &H101, WSAInfo
    
            m_ConnectionCount = 0
        ReDim m_Connections(m_ConnectionCount)
            
 Exit Sub
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Class_Initialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Purpose:
' When this object is closed we just make sure that we close the socket, clean up the connection and stop the subclassing.
Private Sub Class_Terminate()
On Error GoTo ErrorHandler
Dim ret As Long
   'Close all clients that are connected
            CloseAll
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
   'Cleanup Winsock
            WSACleanup
   'clean up the connection object
            Erase m_Connections
                m_ConnectionCount = 0
   'Stop subclassing
            Set objSuperClass = Nothing  'Destroy the subclasser
   'Also remove the window
            DestroyWindow lngWindowHandle
Exit Sub
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Class_Terminate on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub
 
'Purpse:
' You can get the name of the client computer.
Public Function GetLocalHost() As String
On Error GoTo ErrorHandler
Const LOCAL_HOST_BUFF As Integer = 256
    Dim strHostNameBuf As String * LOCAL_HOST_BUFF
    Dim lngResult As Long
    
    lngResult = gethostname(strHostNameBuf, LOCAL_HOST_BUFF)
    
    If lngResult = SOCKET_ERROR Then
        GetLocalHost = vbNullString
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
        Err.Raise lngErrorCode, "CSocketMaster.GetLocalHostName", GetErrorDescription(lngErrorCode)
    Else
        GetLocalHost = Left(strHostNameBuf, InStr(1, strHostNameBuf, vbNullChar) - 1)
    End If
 
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetLocalHost on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
'Purpse:
' You can get the IP adress of the client computer.
Public Function GetLocalIP(ByVal lngSocket As Long) As String
On Error GoTo ErrorHandler
    Dim lngResult As Long
    Dim lngPtrToIP As Long
    Dim strLocalHost As String
    Dim arrIpAddress(1 To 4) As Byte
    Dim Count As Integer
    Dim udtHostent As HOSTENT
    Dim strIpAddress As String
    
    strLocalHost = GetLocalHost
    
    lngResult = gethostbyname(strLocalHost)
    
    If lngResult = 0 Then
        GetLocalIP = vbNullString
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
        Err.Raise lngErrorCode, "CSocketMaster.GetLocalIP", GetErrorDescription(lngErrorCode)
    Else
         CopyMemory udtHostent, ByVal lngResult, LenB(udtHostent)
         CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
         CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
       
        For Count = 1 To 4
            strIpAddress = strIpAddress & arrIpAddress(Count) & "."
        Next
       
        strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
        GetLocalIP = strIpAddress
    End If
Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetLocalIP on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
'Purpse:
' You can get the connection port on the client computer.
Public Function GetLocalPort(ByVal lngSocket As Long) As Long
On Error GoTo ErrorHandler
    Dim udtSockAddr As SOCKADDR_IN
    Dim lngResult As Long
    
    lngResult = GetSockName(lngSocket, udtSockAddr, LenB(udtSockAddr))
    
    If lngResult = SOCKET_ERROR Then
        GetLocalPort = SOCKET_ERROR
    Else
        GetLocalPort = IntegerToUnsigned(ntohs(udtSockAddr.sin_port))
    End If
Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetLocalPort on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
'Purpse:
' You can get the name of the remote computer.
Public Function GetRemoteHost(ByVal lngSocket As Long) As String
On Error GoTo ErrorHandler
    Dim C As Long
        C = ConnectionBySocket(lngSocket)
    If C <> -1 Then
        GetRemoteHost = m_Connections(C).strRemoteHost
    End If
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemoteHost on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

'Purpse:
' You can get the IP adress of the remote computer.
Public Function GetRemoteIP(ByVal lngSocket As Long) As String
On Error GoTo ErrorHandler
    Dim C As Long
        C = ConnectionBySocket(lngSocket)
    If C <> -1 Then
        GetRemoteIP = m_Connections(C).strRemoteIP
    End If
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemoteIP on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

'Purpse:
' You can get the port number of the remote computer.
Public Function GetRemotePort(ByVal lngSocket As Long) As Long
On Error GoTo ErrorHandler
    Dim C As Long
        C = ConnectionBySocket(lngSocket)
    If C <> -1 Then
        GetRemotePort = m_Connections(C).lngRemotePort
    End If
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemotePort on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

Public Function GetSocket() As Long
On Error GoTo ErrorHandler
    GetSocket = m_lngSocket
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in GetRemotePort on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

'Purpose:
' Send data to the remote computer.
Public Function Send(ByVal lngSocket As Long, ByVal strData As String) As Long
On Error GoTo ErrorHandler

    mSend lngSocket, strData
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Send on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

'Purpose:
' Send data to the remote computer.
Public Function SendByte(ByVal lngSocket As Long, byteData() As Byte) As Long
On Error GoTo ErrorHandler
    mSendByte lngSocket, byteData
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in SendByte on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
'Purpose:
' After receiving an OnDataArive event we can get the data with this.
Public Function Recv(ByVal lngSocket As Long, ByRef strBuffer As String, Optional ByVal bLength As Long) As Long
On Error GoTo ErrorHandler
           Const MAX_BUFFER_LENGTH As Long = 8192 'Normal= 8192  'MAX = 65536
'  Dim arrBuffer(1 To MAX_BUFFER_LENGTH)   As Byte
  Dim arrBuffer()   As Byte
  Dim lngBytesReceived                    As Long
  Dim strTempBuffer                       As String
                If bLength = 0 Then
                    bLength = MAX_BUFFER_LENGTH
                End If
                ReDim arrBuffer(bLength - 1) As Byte
    'Call the recv Winsock API function in order to read data from the buffer
'             lngBytesReceived = WSARecv(lngSocket, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)
             lngBytesReceived = WSARecv(lngSocket, arrBuffer(0), bLength, 0&)

             If lngBytesReceived > 0 Then
        'If we have received some data, convert it to the Unicode
        'string that is suitable for the Visual Basic String data type
                 strTempBuffer = StrConv(arrBuffer, vbUnicode)

        'Remove unused bytes
                 strBuffer = Left$(strTempBuffer, lngBytesReceived)
             End If
        
             Recv = lngBytesReceived
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in Recv on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function

Private Sub Remove_m_Connection(ByVal Index As Long)
If m_ConnectionCount > 0 Then
'    Dim ByteLen As Byte
'
'    '// String pointers are 4 bytes
'    ByteLen = LenB(m_Connections(0))
'
'    '// The copymemory operation is not necessary unless
'    '// we are working with an array element that is not
'    '// at the end of the array
'    If Index < UBound(m_Connections) Then
'        '// Copy the block of string pointers starting at
'        ' the position after the
'        '// removed item back one spot.
'        CopyMemory ByVal VarPtr(m_Connections(Index)), ByVal _
'            VarPtr(m_Connections(Index + 1)), (ByteLen) * _
'            (UBound(m_Connections) - Index)
'    End If
'
'    '// If we are removing the last array element
'    '// just deinitialize the array
'    '// otherwise chop the array down by one.
'    If UBound(m_Connections) = LBound(m_Connections) Then
'        Erase m_Connections
'    Else
'        ReDim Preserve m_Connections(UBound(m_Connections) - 1)
'    End If
'        m_ConnectionCount = m_ConnectionCount - 1
    Dim tmps() As ConnectionInfo
    Dim C As Long, i As Long
            ReDim tmps(m_ConnectionCount - 1)
    For C = 0 To m_ConnectionCount - 1
        If C <> Index Then
                LSet tmps(i) = m_Connections(C)
            i = i + 1
        End If
    Next
        m_ConnectionCount = m_ConnectionCount - 1
        Erase m_Connections
        ReDim m_Connections(m_ConnectionCount)
    For C = 0 To i - 1
        LSet m_Connections(C) = tmps(C)
    Next
        Erase tmps
End If
End Sub
'Purpose:
' Close the connection.
Public Sub CloseSocket(ByVal lngSocket As Long)
On Error GoTo ErrorHandler
Dim C As Long, Index As Long
            Index = -1
        For C = 0 To m_ConnectionCount - 1
            If m_Connections(C).lngSocket = lngSocket Then
                Index = C
                Exit For
            End If
        Next
    If LastAcceptedSocket = lngSocket Then
        LastAcceptedSocket = 0
    End If
If Index <> -1 Then
     WSACloseSocket m_Connections(Index).lngSocket
    Remove_m_Connection Index
    Exit Sub
End If
     WSACloseSocket lngSocket
 Exit Sub
ErrorHandler:
   Err.Raise vbObjectError Or Err, "Connection", "Connection :: Error in CloseSocket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub
 
'Purpose:
' This method will set the socket in listening mode.
' The moment a client requests a connection the event ConnectionRequest will be trigered.
' You must then call the .Accept method to accept this connection.
Public Function Listen(ByVal lngLocalPort As Long) As Boolean
On Error GoTo ErrorHandler
Dim udtSocketAddress As SOCKADDR_IN
Dim lngReturnValue   As Long
Dim lngAddress       As Long
                m_Listening = False

    'Create a new socket
             m_lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    'Let the socket send the messages to the subclassed window handle
        lngReturnValue = WSAAsyncSelect(m_lngSocket, lngWindowHandle, lngWinsockMessage, FD_CONNECT Or FD_WRITE Or FD_READ Or FD_CLOSE Or FD_ACCEPT)
             Listen = m_lngSocket
             m_intConnectionPort = lngLocalPort

    'If Not m_lngSocket > 0 Then Exit Function 'Check the socket handle
    'If Not lngLocalPort > 0 Then Exit Function 'Check the port number
    
             With udtSocketAddress
                 .sin_addr = inet_addr("0.0.0.0")                    'We will accept any IP
                 .sin_port = htons(UnsignedToInteger(lngLocalPort))  'The port to listen on
                 .sin_family = AF_INET                               'The internet family (TCP/IP, UDP etc)
             End With
        
    'Call the bind method
             lngReturnValue = WSABind(m_lngSocket, udtSocketAddress, LenB(udtSocketAddress))

    'Call the listen function
            lngReturnValue = WSAListen(m_lngSocket, SOMAXCONN)
             If lngReturnValue <> 0 Then
                RaiseEvent OnError(vbObjectError + 1, "Could not start listening on port " & lngLocalPort)
             End If
             
        m_Listening = True
  
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Listen on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
'Purpose:
' You have to call this method to accept a client that wants to connect to this server.
Public Function Accept(ByVal lngSocketHandle As Long) As Long
On Error GoTo ErrorHandler
Dim udtSocketAddress As SOCKADDR_IN
Dim lReturn As Long
    'Call the accept Winsock API function in order to create a new socket
             LastAcceptedSocket = WSAAccept(lngSocketHandle, udtSocketAddress, LenB(udtSocketAddress))
        
    'Let the socket send the messages to the subclassed window handle (see class_initialize)
'lReturn = WSAAsyncSelect(LastAcceptedSocket, lngWindowHandle, lngWinsockMessage, FD_CONNECT Or FD_WRITE Or FD_READ Or FD_CLOSE Or FD_ACCEPT)
lReturn = WSAAsyncSelect(LastAcceptedSocket, lngWindowHandle, lngWinsockMessage, FD_WRITE Or FD_READ Or FD_CLOSE)

       
'    Dim m_C As New Connection
'             m_C.Socket = Accept
'             m_Connections.Add m_C
        ReDim Preserve m_Connections(m_ConnectionCount)
            m_Connections(m_ConnectionCount).lngSocket = LastAcceptedSocket
Dim lngLocalPort As Long, strLocalIP As String
Dim lngRemortPort As Long, strRemoteIP As String, strRemoteHost As String
            GetLocalInfo LastAcceptedSocket, lngLocalPort, strLocalIP
            GetRemoteInfo LastAcceptedSocket, lngRemortPort, strRemoteIP, strRemoteHost
                m_Connections(m_ConnectionCount).lngRemotePort = lngRemortPort
                m_Connections(m_ConnectionCount).strRemoteIP = strRemoteIP
                m_Connections(m_ConnectionCount).strRemoteHost = strRemoteHost
        m_ConnectionCount = m_ConnectionCount + 1
    
        Accept = LastAcceptedSocket
 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in Accept on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
Private Function GetLocalInfo(ByVal lngSocket As Long, ByRef lngLocalPort As Long, ByRef strLocalIP As String) As Boolean
GetLocalInfo = False
Dim lngResult As Long
Dim udtSockAddr As SOCKADDR_IN

lngResult = GetSockName(lngSocket, udtSockAddr, LenB(udtSockAddr))

If lngResult = SOCKET_ERROR Then
    lngLocalPort = 0
    strLocalIP = ""
Else
    GetLocalInfo = True
    lngLocalPort = IntegerToUnsigned(ntohs(udtSockAddr.sin_port))
    strLocalIP = StringFromPointer(inet_ntoa(udtSockAddr.sin_addr))
End If

End Function

'Retrieves remote info from a connected socket.
'If succeeds returns TRUE and loads the arguments.
'If fails returns FALSE and arguments are not loaded.
Private Function GetRemoteInfo(ByVal lngSocket As Long, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String) As Boolean
GetRemoteInfo = False
Dim lngResult As Long
Dim udtSockAddr As SOCKADDR_IN

lngResult = GetPeerName(lngSocket, udtSockAddr, LenB(udtSockAddr))

If lngResult = 0 Then
    GetRemoteInfo = True
    GetRemoteInfoFromSI udtSockAddr, lngRemotePort, strRemoteHostIP, strRemoteHost
Else
   lngRemotePort = 0
   strRemoteHostIP = ""
   strRemoteHost = ""
End If
End Function

'Gets remote info from a sockaddr_in structure.
Private Sub GetRemoteInfoFromSI(ByRef udtSockAddr As SOCKADDR_IN, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String)

'Dim lngResult As Long
'Dim udtHostent As HOSTENT

lngRemotePort = IntegerToUnsigned(ntohs(udtSockAddr.sin_port))
strRemoteHostIP = StringFromPointer(inet_ntoa(udtSockAddr.sin_addr))
'lngResult = api_gethostbyaddr(udtSockAddr.sin_addr, 4&, AF_INET)

'If lngResult <> 0 Then
'    api_CopyMemory udtHostent, ByVal lngResult, LenB(udtHostent)
'    strRemoteHost = StringFromPointer(udtHostent.hName)
'Else
'    strRemoteHost = ""
'End If

End Sub
'Purpose:
'Make sure that all clients are disconnected.
Public Function CloseAll()
On Error GoTo ErrorHandler
Dim l As Long
    ' Close all connections
'             If m_Connections.Count > 0 Then
'                For l = m_Connections.Count To 1 Step -1
'                   m_Connections(l).CloseSocket
'                   m_Connections.Remove l
'                Next
'             End If
                For l = m_ConnectionCount - 1 To 0 Step -1
                    If m_Connections(l).lngSocket <> 0 Then
                        CloseSocket m_Connections(l).lngSocket
                    End If
                Next
            Erase m_Connections
                m_ConnectionCount = 0
            
    ' Close the listening socket
             WSACloseSocket m_lngSocket
             m_lngSocket = 0

 Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in CloseAll on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
  

'Purpose:
' This will give you the socket number where the server is listening on.
Public Property Get ListenSocket() As Long
On Error GoTo ErrorHandler

           ListenSocket = m_lngSocket

 Exit Property
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in ListenSocket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property
Public Property Get ConnectSocket() As Long
On Error GoTo ErrorHandler

           ConnectSocket = m_lngConnectionSocket

 Exit Property
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in ListenSocket on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property

Private Function ConnectionByIndex(ByVal Index As Long) As Long
Dim C As Long
    For C = 0 To m_ConnectionCount - 1
        If C = Index Then
                ConnectionByIndex = m_Connections(C).lngSocket
            Exit For
        End If
    Next
End Function
Private Function ConnectionBySocket(ByVal lngSocket As Long) As Long
Dim C As Long
        ConnectionBySocket = -1
    For C = 0 To m_ConnectionCount - 1
        If m_Connections(C).lngSocket = lngSocket Then
            ConnectionBySocket = C
            Exit For
        End If
    Next
End Function
'Purpose:
' This property will return the number of connected clients.
Public Property Get ConnectionCount() As Long
On Error GoTo ErrorHandler
           ConnectionCount = m_ConnectionCount
 Exit Property
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in m_ConnectionCount on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Property

'Purpose:
' This method will do the actual connect to the server.
' Before sending data make sure that te onConnect event was fired (and onDisconnect not).
' This method will return before the actuall connection was set up.
Public Function Connect(ByVal strRemoteHost As String, ByVal intRemotePort As Integer) As Long
On Error GoTo ErrorHandler
Dim lngSocket As Long

    'Create a new socket
             lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
             m_lngSocket = lngSocket
    
    'Let the socket send the messages to the subclassed window handle (see class_initialize)
             WSAAsyncSelect lngSocket, lngWindowHandle, lngWinsockMessage, FD_CONNECT Or FD_WRITE Or FD_READ Or FD_CLOSE Or FD_ACCEPT
             Connect = lngSocket
    'Save the connection parameters
             m_intConnectionPort = intRemotePort
             m_lngConnectionSocket = lngSocket
    
    'Resolve the host name and if that was successful then give the socket a connect command.
             ResolveHost strRemoteHost
    
    
Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in Connect on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
  
Private Sub ResolveHost(strHostName As String)
On Error GoTo ErrorHandler
Dim lngAddress As Long

    'Try and resolve the address. This will work if it was an IP we were given
             lngAddress = inet_addr(strHostName)
    
    'We were unable to resolve it so we will have to go for the long way
             If lngAddress = INADDR_NONE Then
        'Allocate 1Kb of fixed memory
                 lngMemoryHandle = GlobalAlloc(GMEM_FIXED, 1024)
        
                 If lngMemoryHandle > 0 Then
            'Lock the memory block just to get the address
                     lngMemoryPointer = GlobalLock(lngMemoryHandle)

                     If lngMemoryPointer = 0 Then
                'Memory allocation error
                         Call GlobalFree(lngMemoryHandle)
                         Exit Sub
                     Else
                'Unlock the memory block
                         GlobalUnlock (lngMemoryHandle)
                     End If
                 Else
            'Memory allocation error
                     Exit Sub
                 End If
        
        'Get the host by the name. This is an Asynchroneous call. This means
        'that the call will not freeze the app. It will post a message
        'to the WindowProc when it has finished.
                 WSAAsyncGetHostByName lngWindowHandle, lngResolveHostMessage, strHostName, ByVal lngMemoryPointer, 1024
             Else
        'A valid IP adress was suplied as a hostname so we can connect immediately.
                 ResolveHostMessage (lngAddress)
             End If

Exit Sub
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in ResolveHost on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub

'Purpuse:
' Give a connect message to the winsock.
Private Sub ResolveHostMessage(lngAddress As Long)
On Error GoTo ErrorHandler
Dim udtSocketAddress As SOCKADDR_IN
Dim ret As Long, lngErrorCode As Long
   ' We will only connect if we have a valid IP adress (converted to a long)
            If lngAddress <> INADDR_NONE Then
               With udtSocketAddress
                  .sin_addr = lngAddress
                  .sin_port = htons(UnsignedToInteger(CLng(m_intConnectionPort)))
                  .sin_family = AF_INET
               End With
      'Call the connect API function
               ret = WSAConnect(m_lngConnectionSocket, udtSocketAddress, LenB(udtSocketAddress))
            End If
'    If ret = -1 Then
'        lngErrorCode = WSAGetLastError
'        If lngErrorCode = 0 Then
'            lngErrorCode = GetLastError
'            If lngErrorCode = 0 Then lngErrorCode = Err.LastDllError
'            RaiseEvent OnError(lngErrorCode, GetErrorDescription(lngErrorCode))
'
'        Else
'            RaiseEvent OnError(lngErrorCode, GetErrorDescription(lngErrorCode))
'        End If
'    End If
Exit Sub
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericClient", "GenericClient :: Error in ResolveHostMesssage on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub


Function Get_Error_Message(ByVal errCode As Long) As String
'    errCode=GetLastError
    Dim errBuffer As String
    'Create a string buffer
    errBuffer = Space(256)
    'Format the message string
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, errCode, LANG_NEUTRAL, errBuffer, 256, ByVal 0&
    'Show the message
        If InStr(errBuffer, Chr(0)) Then
            Buffer = Mid$(errBuffer, 1, InStr(errBuffer, Chr(0)) - 1)
        End If
'    MsgBox Buffer
    Get_Error_Message = errBuffer
End Function

Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
On Error GoTo ErrorHandler
'We only use the _After
'wParam is the socket
'lParam is the winsock message
'uMsg is the windows message type
Dim lngErrorCode As Long

   'was an error passed on
            lngErrorCode = HiWord(lParam)
            If lngErrorCode > 0 Then
               RaiseEvent OnError(lngErrorCode, GetErrorDescription(lngErrorCode))
            End If

   'Was it a winsock or resolvehost message?
Select Case uMsg
   Case lngWinsockMessage
      'What winsock message was posted
            Select Case lParam
                  Case FD_ACCEPT      'There is a connection request pending
                              RaiseEvent ConnectionRequest(m_lngSocket)
                              lReturn = 0
                  Case FD_CONNECT     'The connection is established
                              RaiseEvent OnConnect(m_lngSocket)
                              lReturn = 0
                Case FD_WRITE
                    '
                    'This message means that the socket in a write-able
                    'state, that is, buffer for outgoing data of the transport
                    'service is empty and ready to receive data to send through
                    'the network.
                    If BufferLength > 0 Then
                        Call mSendByte(wParam, Buffer)
                        RaiseEvent SendProgress(wParam, SentLength, BufferLength)
                    Else
                        RaiseEvent SendProgress(wParam, SentLength, 0)
                    End If
                              lReturn = 0
                  
                  Case FD_READ        'We are ready to read from the winsock buffer
                    ' We needed this as a solution to handle faster incomming data than what we can process
                     Static intBufferCounter As Long
                     intBufferCounter = intBufferCounter + 1
                     If intBufferCounter > 1 Then Exit Sub
ContinueWithBuffer:
                     RaiseEvent DataArrival(wParam)
                     intBufferCounter = intBufferCounter - 1
                     If intBufferCounter > 0 Then GoTo ContinueWithBuffer
        
        
                          lReturn = 0
                  
                   Case FD_CLOSE       'The other end point wishes to close the connection
                        CloseSocket wParam
                              RaiseEvent OnClose(wParam)
                              lReturn = 0
               End Select
   Case lngResolveHostMessage
      'A resolve message was posted. This only happens if we connect to a name instead of an IP adress.
      Dim udtHost           As HOSTENT
      Dim lngIpAddrPtr      As Long
      Dim lngHostAddress    As Long
               If Not lngErrorCode > 0 Then
         'Extract the host name from the memory block
                  RtlMoveMemory udtHost, ByVal lngMemoryPointer, Len(udtHost)
                  RtlMoveMemory lngIpAddrPtr, ByVal udtHost.hAddrList, 4
                  RtlMoveMemory lngHostAddress, ByVal lngIpAddrPtr, 4
         'Free the allocated memory block
                  Call GlobalUnlock(lngMemoryHandle)
                  Call GlobalFree(lngMemoryHandle)
               Else
                  lngHostAddress = INADDR_NONE
               End If
      'Now connect to the resolved host.
               ResolveHostMessage lngHostAddress
End Select


 Exit Sub
ErrorHandler:
   Err.Raise vbObjectError Or Err, "GenericServer", "GenericServer :: Error in iSubclass_After on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub
 
'Purpose:
' Convert an unsigned long to an integer.
Private Function UnsignedToInteger(Value As Long) As Integer
On Error GoTo ErrorHandler

             If Value < 0 Or Value >= OFFSET_2 Then Error 6  'Overflow
    
             If Value <= MAXINT_2 Then
                 UnsignedToInteger = Value
             Else
                 UnsignedToInteger = Value - OFFSET_2
             End If

Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in UnsignedToInteger on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
'Purpose:
' Convert an integer to an unsigned long.
Private Function IntegerToUnsigned(Value As Integer) As Long
On Error GoTo ErrorHandler
 
             If Value < 0 Then
                 IntegerToUnsigned = Value + OFFSET_2
             Else
                 IntegerToUnsigned = Value
             End If
    
Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in IntegerToUnsigned on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
'Purpose:
' Create a string from a pointer
Private Function StringFromPointer(ByVal lngPointer As Long) As String
On Error GoTo ErrorHandler


  Dim strTemp As String
  Dim lRetVal As Long
    
             strTemp = String$(lstrlen(ByVal lngPointer), 0)    'prepare the strTemp buffer
             lRetVal = lstrcpy(ByVal strTemp, ByVal lngPointer) 'copy the string into the strTemp buffer
             If lRetVal Then StringFromPointer = strTemp        'return the string

Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in StringFromPointer on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function



'Purpose:
' Return the Hi Word of a long value.
Private Function HiWord(lngValue As Long) As Long
On Error GoTo ErrorHandler

             If (lngValue And &H80000000) = &H80000000 Then
                 HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
             Else
                 HiWord = (lngValue And &HFFFF0000) \ &H10000
             End If
    
Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in HiWord on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
'Purpose:
' Send data to the specified socket.
Private Function mSend(ByVal lngSocket As Long, ByVal strData As String) As Long
On Error GoTo ErrorHandler
  Dim arrBuffer()     As Byte

    'Convert the data string to a byte array
             arrBuffer() = StrConv(strData, vbFromUnicode)
    'Call the send Winsock API function in order to send data
'             mSend = WSASend(lngSocket, arrBuffer(0), Len(strData), 0&)
            
    mSend = mSendByte(lngSocket, arrBuffer)
            
Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in mSend on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
Function MidByte(b_in() As Byte, ByVal bstart As Long, Optional ByVal bLength As Long) As Byte()
Dim b() As Byte, bLen As Long
        bLen = UBound(b_in)
            bstart = bstart - 1
    If bstart > bLen Then
        bstart = bLen + 1
    End If
    If bLength = 0 Then
        bLength = (bLen - bstart) + 1
    End If
                If bLength = 0 Then
                    MidByte = b
                    Exit Function
                End If
        ReDim b(bLength - 1)
CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(b_in(bstart)), (LenB(b_in(0))) * bLength
    MidByte = b
End Function

'Purpose:
' Send data to the specified socket.
Private Function mSendByte(ByVal lngSocket As Long, byteData() As Byte) As Long
On Error GoTo ErrorHandler
Dim lngRetValue     As Long         'value returned by send/sendto Winsock API function
    'Call the send Winsock API function in order to send data
                BufferLength = UBound(byteData) + 1
         lngRetValue = WSASend(lngSocket, byteData(0), BufferLength, 0&)
        
        SentLength = lngRetValue
      
 If lngRetValue = SOCKET_ERROR Then
'        'If a value of Err.LastDllError is WSAEWOULDBLOCK, that means
'        'that the Winsock's buffer for outgoing data is full and cannot
'        'accept data to send. In this case we ignore this error and do
'        'not empty local buffer (m_strSendBuffer).
'        '
        
            Buffer = byteData
    
'        If Not Err.LastDllError = WSAEWOULDBLOCK Then
'            'Debug.Print "Error occurred: " & Err.LastDllError & " - " & GetErrorDescription(Err.LastDllError)
'            Err.Raise Err.LastDllError, "Socket.SendData", GetErrorDescription(Err.LastDllError)
'        Else
'            'Debug.Print "WSAEWOULDBLOCK"
'        End If
        '
Else
    '
    'The data were sent successfully. Raise the OnSendProgress or
    'OnSendComplete event to let the client app know.
    '
    'Debug.Print "SendData - Bytes sent: " & lngRetValue
    '
    If BufferLength > lngRetValue Then
        '
'            m_strSendBuffer = Mid$(m_strSendBuffer, lngRetValue + 1)
        '
        '// Copy the block of string pointers starting at
    ' the position after the
    '// removed item back one spot.
        BufferLength = (BufferLength - lngRetValue) + 1
'            ReDim buffer(BufferLength)
'        CopyMemory ByVal VarPtr(buffer(0)), ByVal VarPtr(byteData(lngRetValue + 1)), _
'               (LenB(byteData(0))) * BufferLength
         
'         buffer = MidByte(byteData, lngRetValue + 1, BufferLength)
         Buffer = MidByte(byteData, lngRetValue + 1)
    Else
        '
        Erase Buffer
        BufferLength = 0
        '
        '---------------------------------------------
        'Modified: 23-AUG-2002
        '---------------------------------------------
        'RaiseEvent OnSendComplete
        Call PostMessage(lngWindowHandle, lngWinsockMessage, lngSocket, FD_WRITE)
        '---------------------------------------------
    End If
    '
'    RaiseEvent SendProgress(lngSocket, lngRetValue, BufferLength)
    '
End If


    mSendByte = lngRetValue
    
    
Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in mSendByte on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
 
'Purpose:
' Get the error description of a socket error.
Private Function GetErrorDescription(ByVal lngErrorCode As Long) As String
On Error GoTo ErrorHandler

  Dim strDesc As String
    
             Select Case lngErrorCode
        Case WSAEACCES
                     strDesc = "Permission denied."
        Case WSAEADDRINUSE
                     strDesc = "Address already in use."
        Case WSAEADDRNOTAVAIL
                     strDesc = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT
                     strDesc = "Address family not supported by protocol family."
        Case WSAEALREADY
                     strDesc = "Operation already in progress."
        Case WSAECONNABORTED
                     strDesc = "Software caused connection abort."
        Case WSAECONNREFUSED
                     strDesc = "Connection refused."
        Case WSAECONNRESET
                     strDesc = "Connection reset by peer."
        Case WSAEDESTADDRREQ
                     strDesc = "Destination address required."
        Case WSAEFAULT
                     strDesc = "Bad address."
        Case WSAEHOSTDOWN
                     strDesc = "Host is down."
        Case WSAEHOSTUNREACH
                     strDesc = "No route to host."
        Case WSAEINPROGRESS
                     strDesc = "Operation now in progress."
        Case WSAEINTR
                     strDesc = "Interrupted function call."
        Case WSAEINVAL
                     strDesc = "Invalid argument."
        Case WSAEISCONN
                     strDesc = "Socket is already connected."
        Case WSAEMFILE
                     strDesc = "Too many open files."
        Case WSAEMSGSIZE
                     strDesc = "Message too long."
        Case WSAENETDOWN
                     strDesc = "Network is down."
        Case WSAENETRESET
                     strDesc = "Network dropped connection on reset."
        Case WSAENETUNREACH
                     strDesc = "Network is unreachable."
        Case WSAENOBUFS
                     strDesc = "No buffer space available."
        Case WSAENOPROTOOPT
                     strDesc = "Bad protocol option."
        Case WSAENOTCONN
                     strDesc = "Socket is not connected."
        Case WSAENOTSOCK
                     strDesc = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP
                     strDesc = "Operation not supported."
        Case WSAEPFNOSUPPORT
                     strDesc = "Protocol family not supported."
        Case WSAEPROCLIM
                     strDesc = "Too many processes."
        Case WSAEPROTONOSUPPORT
                     strDesc = "Protocol not supported."
        Case WSAEPROTOTYPE
                     strDesc = "Protocol wrong type for socket."
        Case WSAESHUTDOWN
                     strDesc = "Cannot send after socket shutdown."
        Case WSAESOCKTNOSUPPORT
                     strDesc = "Socket type not supported."
        Case WSAETIMEDOUT
                     strDesc = "Connection timed out."
        Case WSATYPE_NOT_FOUND
                     strDesc = "Class type not found."
        Case WSAEWOULDBLOCK
                     strDesc = "Resource temporarily unavailable."
        Case WSAHOST_NOT_FOUND
                     strDesc = "Host not found."
        Case WSANOTINITIALISED
                     strDesc = "Successful WSAStartup not yet performed."
        Case WSANO_DATA
                     strDesc = "Valid name, no data record of requested type."
        Case WSANO_RECOVERY
                     strDesc = "This is a nonrecoverable error."
        Case WSASYSCALLFAILURE
                     strDesc = "System call failure."
        Case WSASYSNOTREADY
                     strDesc = "Network subsystem is unavailable."
        Case WSATRY_AGAIN
                     strDesc = "Nonauthoritative host not found."
        Case WSAVERNOTSUPPORTED
                     strDesc = "Winsock.dll version out of range."
        Case WSAEDISCON
                     strDesc = "Graceful shutdown in progress."
        Case Else
                     strDesc = "Unknown error."
             End Select
    
             GetErrorDescription = strDesc
    
Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in GetErrorDescription on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function


Private Sub SocketsCleanup()
On Error GoTo ErrorHandler
    
            WSACleanup

Exit Sub
ErrorHandler:
   Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in SocketsCleanup on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Sub


Private Function SocketsInitialize() As Boolean
On Error GoTo ErrorHandler
Dim WSAD            As WSADATA

            SocketsInitialize = False

            If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then Exit Function
            If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then Exit Function
  
            SocketsInitialize = True

Exit Function
ErrorHandler:
   Err.Raise vbObjectError Or Err, "modWinsock", "modWinsock :: Error in SocketsInitialize on line " & Erl() & " triggered by " & Err.Source & "   (" & Err.Number & ")" & vbCrLf & Err.Description
End Function
 
