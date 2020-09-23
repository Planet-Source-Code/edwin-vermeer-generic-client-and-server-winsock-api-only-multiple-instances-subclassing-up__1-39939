VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenericServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose:
' The GenericServer object can be used to create any sort of server application.
' This object will hide the difficulties in the Winsock API calls for your.<BR><BR>
' You will probably write your own protocol handler (like SMTP, POP or ...) in the OnDataArive event.<BR>
Option Explicit

'Some variables that are used to store the state of this object.
Private m_lngSocket As Long               'Just remember the socket where we listen on
Private m_objConnection As New Collection 'When a client is connected to this server then a new connection object will be added to this collection for handeling this connection.
Attribute m_objConnection.VB_VarHelpID = -1
Private m_intConnectionPort As Integer    'Just remember the port where we listen on

' We are going to subclass a generated window and add 2 new message numbers for the communication
Private objSuperClass As cSuperClass      'This is the subclass handler
Implements iSuperClass                    'This is the interface to the subclass handler
Private WinsockMessage As Long            'Winsock resolve host message
Private WindowHandle As Long              'Window handle

'Purpose: This event will be called when the connection is closed by the server or the client itself.
Public Event OnClose(lngSocket As Long)
'Purpose: This event will be called when the connect method executed successfully.
Public Event OnConnectRequest(lngSocket As Long)
'Purpose: This event will be called when data arived. After this you have to call the .Connection.Recv method to get the data.
Public Event OnDataArrive(lngSocket As Long)
'Purpose: This event will be called whenever there is a problem.
Public Event OnError(lngRetCode As Long, strDescription As String)




'Purpose:
' The moment an instance of this object is created we can set up a new socket for communication.
Private Sub Class_Initialize()
Dim WSAInfo As WSAData
    
   'Create a blank, invisible window
   WindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

   'Create a unique message number for our messages
   WinsockMessage = RegisterWindowMessage(App.EXEName & ".WinsockMessage")
    
   'Start up the subclassing
   Set objSuperClass = New cSuperClass        'Create a cSuperClass instance
   With objSuperClass
     .AddMsg WinsockMessage
     .Subclass WindowHandle, Me
   End With
     
   'Start version 1.1 of the Microsoft Winsock API
   WSAStartup &H101, WSAInfo
    
End Sub




'Purpose:
' When this object is closed we just make sure that we close the socket, clean up all the connection and stop the subclassing.
Private Sub Class_Terminate()
    
   'Close all clients that are connected
   CloseAll
    
   'Cleanup Winsock
   WSACleanup
        
   'Stop subclassing
   Set objSuperClass = Nothing  'Destroy the subclasser
    
   'Also remove the window
   DestroyWindow WindowHandle
       
End Sub



'Purpose:
' This method will set the socket in listening mode.
' The moment a client requests a connection the event OnConnectRequest will be trigered.
' You must then call the .Accept method to accept this connection.
Public Function Listen(lngLocalPort As Long) As Long
Dim udtSocketAddress As SOCKADDR_IN
Dim lngReturnValue   As Long
Dim lngAddress       As Long

    'Create a new socket
    m_lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    'Let the socket send the messages to the subclassed window handle
    WSAAsyncSelect m_lngSocket, WindowHandle, WinsockMessage, FD_CONNECT Or FD_READ Or FD_CLOSE Or FD_ACCEPT
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
    Call WSABind(m_lngSocket, udtSocketAddress, LenB(udtSocketAddress))

    'Call the listen function
    If WSAListen(m_lngSocket, SOMAXCONN) <> 0 Then
       RaiseEvent OnError(vbObjectError + 1, "Could not start listening on port " & lngLocalPort)
    End If

End Function



'Purpose:
' You have to call this method to accept a client that wants to connect to this server.
Public Function Accept(ByVal lngSocketHandle As Long) As Long
Dim udtSocketAddress As SOCKADDR_IN

    'Call the accept Winsock API function in order to create a new socket
    Accept = WSAAccept(lngSocketHandle, udtSocketAddress, LenB(udtSocketAddress))
    
    Dim m_C As New clsConnection
    m_C.Socket = Accept
    m_objConnection.Add m_C

End Function



'Purpose:
'Make sure that all clients are disconnected.
Public Function CloseAll()
Dim l As Long
    
    ' Close all connections
    If m_objConnection.Count > 0 Then
       For l = m_objConnection.Count To 1 Step -1
          m_objConnection(l).CloseSocket
          m_objConnection.Remove l
       Next
    End If
    
    ' Close the listening socket
    WSACloseSocket m_lngSocket
    m_lngSocket = 0
End Function



'Purpose:
' This property can be used to access the connection collection.<BR><BR>
' For instance if you want to send data to a specific client then you can call GenericServer.Connection(lngConnection).Send<BR>
' The variable lngConnection can be the socket number or the negative counter ID<BR>
' For instance if you have 1 connection with socket number 812 then Connection(812) = Connection(-1)<BR>
Public Property Get Connection(lngConnection As Long) As clsConnection
   If lngConnection < 0 Then
      Set Connection = m_objConnection(-lngConnection)
   Else
      Dim o As clsConnection
      For Each o In m_objConnection
         If o.Socket = lngConnection Then
            Set Connection = o
            Exit For
         End If
      Next
   End If
End Property



'Purpose:
' This will give you the socket number where the server is listening on.
Public Property Get ListenSocket() As Long
  ListenSocket = m_lngSocket
End Property



'Purpose:
' This property will return the number of connected clients.
Public Property Get ConnectionCount() As Long
  ConnectionCount = m_objConnection.Count
End Property





'Purpuse:
' This is where the winsock messages are handeled.
Private Sub iSuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
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
   
   'Was it a winsock message?
   If uMsg = WinsockMessage Then
      'What winsock message
      Select Case lParam
      Case FD_ACCEPT      'There is a connection request pending
         RaiseEvent OnConnectRequest(wParam)
         lReturn = 0
      Case FD_READ        'We are ready to read from the winsock buffer
         RaiseEvent OnDataArrive(wParam)
         lReturn = 0
      Case FD_CLOSE       'The other end point wishes to close the connection
         RaiseEvent OnClose(wParam)
         lReturn = 0
      End Select
   End If

End Sub



'Purpose:
' When you subclass you can receive messages before or after the original windows handler was called.
Private Sub iSuperClass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'We only use the _After
End Sub






