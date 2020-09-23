Attribute VB_Name = "modWinsock"
'Purpose
' For limiting memory load as much as possible functions are moved to this module.
' If there are multible instances of a class then there will be only one instance of this module loaded.
Option Explicit

'Winsock Initialization and termination
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

'Server side Winsock API functions
Public Declare Function WSABind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByRef namelen As Long) As Long
Public Declare Function WSAListen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function WSAAccept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long

'String functions
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'Socket Functions
Public Declare Function WSAConnect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long
Public Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function WSACloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

'Data transfer functions
Public Declare Function WSARecv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function WSASend Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

'Network byte ordering functions
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'End point information
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByRef namelen As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByRef namelen As Long) As Long

'Hostname resolving functions
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

'Winsock API functions for resolving hostnames and IP's
Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long

'Memory copy and move functions
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

'Window creation and destruction functions
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

'Messaging functions
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

'Memory allocation functions
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

'..
Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129

'Maximum queue length specifiable by listen.
Public Const SOMAXCONN = &H7FFFFFFF

'Windows Socket types
Public Const SOCK_STREAM = 1     'Stream socket

'Address family
Public Const AF_INET = 2          'Internetwork: UDP, TCP, etc.

'Socket Protocol
Public Const IPPROTO_TCP = 6     'tcp

'Data type conversion constants
Public Const OFFSET_4 = 4294967296#
Public Const MAXINT_4 = 2147483647
Public Const OFFSET_2 = 65536
Public Const MAXINT_2 = 32767

'Fixed memory flag for GlobalAlloc
Public Const GMEM_FIXED = &H0

'Winsock error offset
Private Const WSABASEERR = 10000

'Winsock messages that will go to the window handler
Public Enum WSAMessage
    FD_READ = &H1&      'Data is ready to be read from the buffer
    FD_CONNECT = &H10&  'Connection esatblished
    FD_CLOSE = &H20&    'Connection closed
    FD_ACCEPT = &H8&    'Connection request pending
End Enum

'Winsock Data structure
Public Type WSAData
    wVersion       As Integer                       'Version
    wHighVersion   As Integer                       'High Version
    szDescription  As String * WSADESCRIPTION_LEN   'Description
    szSystemStatus As String * WSASYS_STATUS_LEN    'Status of system
    iMaxSockets    As Integer                       'Maximum number of sockets allowed
    iMaxUdpDg      As Integer                       'Maximum UDP datagrams
    lpVendorInfo   As Long                          'Vendor Info
End Type

'HostEnt Structure
Public Type HOSTENT
    hName     As Long       'Host Name
    hAliases  As Long       'Alias
    hAddrType As Integer    'Address Type
    hLength   As Integer    'Length
    hAddrList As Long       'Address List
End Type

'Socket Address structure
Public Type SOCKADDR_IN
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
Public Enum WSABaseErrors
    INADDR_NONE = &HFFFF
    SOCKET_ERROR = -1
    INVALID_SOCKET = -1
End Enum

'Winsock error constants
Public Enum WSAErrorConstants
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



'Purpose:
' Convert an unsigned long to an integer.
Public Function UnsignedToInteger(Value As Long) As Integer

    If Value < 0 Or Value >= OFFSET_2 Then Error 6  'Overflow
    
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If

End Function



'Purpose:
' Convert an integer to an unsigned long.
Public Function IntegerToUnsigned(Value As Integer) As Long


    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
    
End Function



'Purpose:
' Create a string from a pointer
Public Function StringFromPointer(ByVal lngPointer As Long) As String


  Dim strTemp As String
  Dim lRetVal As Long
    
    strTemp = String$(lstrlen(ByVal lngPointer), 0)    'prepare the strTemp buffer
    lRetVal = lstrcpy(ByVal strTemp, ByVal lngPointer) 'copy the string into the strTemp buffer
    If lRetVal Then StringFromPointer = strTemp        'return the string

End Function



'Purpose:
' Return the Hi Word of a long value.
Public Function HiWord(lngValue As Long) As Long

    If (lngValue And &H80000000) = &H80000000 Then
        HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        HiWord = (lngValue And &HFFFF0000) \ &H10000
    End If
    
End Function



'Purpose:
' Get the received data from the socket and return it in the calling string.
Public Function mRecv(ByVal lngSocket As Long, strBuffer As String) As Long

  Const MAX_BUFFER_LENGTH As Long = 8192

  Dim arrBuffer(1 To MAX_BUFFER_LENGTH)   As Byte
  Dim lngBytesReceived                    As Long
  Dim strTempBuffer                       As String
    
    'Call the recv Winsock API function in order to read data from the buffer
    lngBytesReceived = WSARecv(lngSocket, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)

    If lngBytesReceived > 0 Then
        'If we have received some data, convert it to the Unicode
        'string that is suitable for the Visual Basic String data type
        strTempBuffer = StrConv(arrBuffer, vbUnicode)

        'Remove unused bytes
        strBuffer = Left$(strTempBuffer, lngBytesReceived)
    End If
        
    mRecv = lngBytesReceived

End Function



'Purpose:
' Send data to the specified socket.
Public Function mSend(ByVal lngSocket As Long, strData As String) As Long

  Dim arrBuffer()     As Byte

    'Convert the data string to a byte array
    arrBuffer() = StrConv(strData, vbFromUnicode)
    'Call the send Winsock API function in order to send data
    mSend = WSASend(lngSocket, arrBuffer(0), Len(strData), 0&)

End Function



'Purpose:
' Get the IP adress of an endpoint (client or server).
Public Function GetIPEndPointField(ByVal lngSocket As Long, ByVal EndpointField As IPEndPointFields) As Variant

  Dim udtSocketAddress    As SOCKADDR_IN
  Dim lngReturnValue      As Long
  Dim lngPtrToAddress     As Long
  Dim strIPAddress        As String
  Dim lngAddress          As Long

    Select Case EndpointField
        Case LOCAL_HOST, LOCAL_HOST_IP, LOCAL_PORT

            'If the info of a local end-point of the connection is
            'requested, call the getsockname Winsock API function
            lngReturnValue = getsockname(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
        Case REMOTE_HOST, REMOTE_HOST_IP, REMOTE_PORT
            
            'If the info of a remote end-point of the connection is
            'requested, call the getpeername Winsock API function
            lngReturnValue = getpeername(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
    End Select
    
    
    If lngReturnValue = 0 Then
        'If no errors occurred, the getsockname or getpeername function returns 0.

        Select Case EndpointField
            Case LOCAL_PORT, REMOTE_PORT
                'Get the port number from the sin_port field and convert the byte ordering
                GetIPEndPointField = IntegerToUnsigned(ntohs(udtSocketAddress.sin_port))
            
            Case LOCAL_HOST_IP, REMOTE_HOST_IP
  
                'Get pointer to the string that contains the IP address
                lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
                
                'Retrieve that string by the pointer
                GetIPEndPointField = StringFromPointer(lngPtrToAddress)
            Case LOCAL_HOST, REMOTE_HOST

                'The same procedure as for an IP address only using GetHostNameByAddress
                lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
                strIPAddress = StringFromPointer(lngPtrToAddress)
                lngAddress = inet_addr(strIPAddress)
                GetIPEndPointField = GetHostNameByAddress(lngAddress)

        End Select
    'An error occured
    Else
        GetIPEndPointField = SOCKET_ERROR
    End If
    
End Function



'Purpose:
' Get the hostname of an endpoint (client or server).
Private Function GetHostNameByAddress(lngInetAdr As Long) As String

  Dim lngPtrHostEnt As Long
  Dim udtHostEnt    As HOSTENT
  Dim strHostName   As String
  
    'Get the pointer to the HOSTENT structure
    lngPtrHostEnt = gethostbyaddr(lngInetAdr, 4, AF_INET)
    
    'Copy data into the HOSTENT structure
    RtlMoveMemory udtHostEnt, ByVal lngPtrHostEnt, LenB(udtHostEnt)
    
    'Prepare the buffer to receive a string
    strHostName = String(256, 0)
    
    'Copy the host name into the strHostName variable
    RtlMoveMemory ByVal strHostName, ByVal udtHostEnt.hName, 256
    
    'Cut received string by first chr(0) character
    GetHostNameByAddress = Left(strHostName, InStr(1, strHostName, Chr(0)) - 1)

End Function



'Purpose:
' Get the error description of a socket error.
Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String

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
    
End Function


