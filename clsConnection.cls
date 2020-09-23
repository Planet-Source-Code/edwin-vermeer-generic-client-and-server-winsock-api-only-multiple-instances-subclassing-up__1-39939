VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Purpose:
' This object is created for a client when a connection is made or for a server for every connection that was made.
' You can get connection specific information or receive and send data.
Option Explicit

Dim m_lngSocket As Long  'The socket number where we are connected to



'Purpose:
' You can get the socket number of this connection.
Public Property Get Socket() As Long
  Socket = m_lngSocket
End Property

' It can only be set by other code within this dll.
Friend Property Let Socket(lngSocket As Long)
  m_lngSocket = lngSocket
End Property



'Purpse:
' You can get the name of the client computer.
Public Function GetLocalHost() As String
Dim lngSocket As Long
   lngSocket = m_lngSocket
   GetLocalHost = GetIPEndPointField(lngSocket, LOCAL_HOST)
End Function



'Purpse:
' You can get the IP adress of the client computer.
Public Function GetLocalIP() As String
Dim lngSocket As Long
   lngSocket = m_lngSocket
   GetLocalIP = GetIPEndPointField(lngSocket, LOCAL_HOST_IP)
End Function



'Purpse:
' You can get the connection port on the client computer.
Public Function GetLocalPort() As Long
Dim lngSocket As Long
   lngSocket = m_lngSocket
   GetLocalPort = GetIPEndPointField(lngSocket, LOCAL_PORT)
End Function




'Purpse:
' You can get the name of the remote computer.
Public Function GetRemoteHost() As String
Dim lngSocket As Long
   lngSocket = m_lngSocket
   GetRemoteHost = GetIPEndPointField(lngSocket, REMOTE_HOST)
End Function



'Purpse:
' You can get the IP adress of the remote computer.
Public Function GetRemoteIP() As String
Dim lngSocket As Long
   lngSocket = m_lngSocket
   GetRemoteIP = GetIPEndPointField(lngSocket, REMOTE_HOST_IP)
End Function



'Purpse:
' You can get the port number of the remote computer.
Public Function GetRemotePort() As Long
Dim lngSocket As Long
   lngSocket = m_lngSocket
   GetRemotePort = GetIPEndPointField(lngSocket, REMOTE_PORT)
End Function



'Purpose:
' Send data to the remote computer.
Public Function Send(strData As String) As Long
Dim lngSocket As Long
   lngSocket = m_lngSocket
    mSend lngSocket, strData
End Function



'Purpose:
' After receiving an OnDataArive event we can get the data with this.
Public Function Recv(strBuffer As String) As Long
Dim lngSocket As Long
   lngSocket = m_lngSocket
   mRecv lngSocket, strBuffer
End Function



'Purpose:
' Close the connection.
Public Sub CloseSocket()
    WSACloseSocket m_lngSocket
    m_lngSocket = 0
End Sub
