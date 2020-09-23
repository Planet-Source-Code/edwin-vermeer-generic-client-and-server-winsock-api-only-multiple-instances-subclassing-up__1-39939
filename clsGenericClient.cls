VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenericClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose:
' The GenericClient object can be used to create any sort of client application.
' This object will hide the difficulties in the Winsock API calls for your.<BR><BR>
' You will probably write your own protocol handler (like SMTP, POP or ...) in the OnDataArive event.<BR>
Option Explicit

'Some variables that are used to store the state of this object.
Private m_intConnectionPort As Integer    'Just remember the port that we connected to
Private m_lngConnectionSocket As Long     'Just remember the socket where the connection is set up for
Private m_objConnection As clsConnection  'When we are connected to the server this object will be created for handeling this connection.

'We need a memory block for resolving the hostname. these variables will be used for handeling those.
Private lngMemoryHandle    As Long      'Handle of the allocated memory block object
Private lngMemoryPointer   As Long      'Address of the memory block

'We are going to subclass a generated window and add 2 new message numbers for the communication
Private objSuperClass As cSuperClass               'This is the subclass handler
Implements iSuperClass                  'This is the interface to the subclass handler
Private lngWinsockMessage As Long          'Winsock resolve host message
Private lngResolveHostMessage As Long   'General Winsock message
Private lngWindowHandle As Long            'Window handle

'Purpose: This event will be called when the connection is closed by the server or the client itself.
Public Event OnClose()
'Purpose: This event will be called when the connect method executed successfully.
Public Event OnConnect()
'Purpose: This event will be called when data arived. After this you have to call the .Connection.Recv method to get the data.
Public Event OnDataArrive()
'Purpose: This event will be called whenever there is a problem.
Public Event OnError(lngRetCode As Long, strDescription As String)



'Purpose:
' The moment an instance of this object is created we can set up a new socket for communication.
Private Sub Class_Initialize()
Dim WSAInfo As WSAData
    
   'Create a blank, invisible window
   lngWindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

   'Create 2 unique message numbers for our messages
   lngWinsockMessage = RegisterWindowMessage(App.EXEName & ".WinsockMessage")
   lngResolveHostMessage = RegisterWindowMessage(App.EXEName & ".ResolveHostMessage")
    
   'Start up the subclassing for the new blank invisible window. This will only be done for the 2 new messages
   Set objSuperClass = New cSuperClass        'Create a cSuperClass instance
   With objSuperClass
     .AddMsg lngWinsockMessage
     .AddMsg lngResolveHostMessage
     .Subclass lngWindowHandle, Me
   End With
   'Later on when we start the connection we will tell the created winsock to
   'post it's messages to the above subclassed window.
     
   'Start version 1.1 of the Microsoft Winsock API
   WSAStartup &H101, WSAInfo
    
   'Since we only need one and the chance is big we are going to connect, just create the connection object now
   Set m_objConnection = New clsConnection
    
End Sub



'Purpose:
' When this object is closed we just make sure that we close the socket, clean up the connection and stop the subclassing.
Private Sub Class_Terminate()
    
   'Cleanup Winsock
   WSACleanup
        
   'clean up the connection object
   Set m_objConnection = Nothing
    
   'Stop subclassing
   Set objSuperClass = Nothing  'Destroy the subclasser
   
   'Also remove the window
   DestroyWindow lngWindowHandle
   
End Sub



'Purpose:
' This method will do the actual connect to the server.
' Before sending data make sure that te onConnect event was fired (and onDisconnect not).
' This method will return before the actuall connection was set up.
Public Function Connect(ByVal strRemoteHost As String, ByVal intRemotePort As Integer) As Long
Dim lngSocket As Long

    'Create a new socket
    lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    m_objConnection.Socket = lngSocket
    
    'Let the socket send the messages to the subclassed window handle (see class_initialize)
    WSAAsyncSelect lngSocket, lngWindowHandle, lngWinsockMessage, FD_CONNECT Or FD_READ Or FD_CLOSE Or FD_ACCEPT
    Connect = lngSocket
    
    'Resolve the host name and if that was successful then give the socket a connect command.
    ResolveHost strRemoteHost
    
    'Save the connection parameters
    m_intConnectionPort = intRemotePort
    m_lngConnectionSocket = lngSocket
    
End Function



'Purpose:
' We can access the connection object from code with this property.
' This will let you use a sintax like GenericClient.Connection.Send
Public Property Get Connection() As clsConnection
  Set Connection = m_objConnection
End Property



'Purpuse:
' This is where the winsock messages are handeled.
Private Sub iSuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
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
      Case FD_CONNECT     'The connection is established
         RaiseEvent OnConnect
         lReturn = 0
      Case FD_READ        'We are ready to read from the winsock buffer
         RaiseEvent OnDataArrive
         lReturn = 0
      Case FD_CLOSE       'The other end point wishes to close the connection
         RaiseEvent OnClose
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

End Sub



'Purpose:
' When you subclass you can receive messages before or after the original windows handler was called.
Private Sub iSuperClass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'We only use the _After
End Sub



'Purpose:
'Connect to the host that was provided. If it was a name instead of an IP then it will be resolved first.
Private Sub ResolveHost(strHostName As String)
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

End Sub



'Purpuse:
' Give a connect message to the winsock.
Private Sub ResolveHostMessage(lngAddress As Long)
Dim udtSocketAddress As SOCKADDR_IN

   ' We will only connect if we have a valid IP adress (converted to a long)
   If lngAddress <> INADDR_NONE Then
      With udtSocketAddress
         .sin_addr = lngAddress
         .sin_port = htons(UnsignedToInteger(CLng(m_intConnectionPort)))
         .sin_family = AF_INET
      End With
      'Call the connect API function
      Call WSAConnect(m_lngConnectionSocket, udtSocketAddress, LenB(udtSocketAddress))
   End If
    
End Sub





