VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Selected Client"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   11535
      Begin VB.CommandButton cmdSendAll 
         Caption         =   "Send all"
         Height          =   315
         Left            =   7680
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDataRecv 
         Height          =   1575
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1440
         Width           =   8535
      End
      Begin VB.TextBox txtSendData 
         Height          =   1080
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   7455
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   315
         Left            =   7680
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblRemotePort 
         Caption         =   "Remote Port:"
         Height          =   330
         Left            =   8760
         TabIndex        =   11
         Top             =   2460
         Width           =   2640
      End
      Begin VB.Label lblRemoteIP 
         Caption         =   "Remote IP:"
         Height          =   330
         Left            =   8760
         TabIndex        =   10
         Top             =   2040
         Width           =   2640
      End
      Begin VB.Label lblRemoteHost 
         Caption         =   "Remote Host"
         Height          =   330
         Left            =   8760
         TabIndex        =   9
         Top             =   1620
         Width           =   2640
      End
      Begin VB.Label lblLocalPort 
         Caption         =   "Local Port:"
         Height          =   330
         Left            =   8760
         TabIndex        =   8
         Top             =   1200
         Width           =   2640
      End
      Begin VB.Label lblLocalIP 
         Caption         =   "Local IP:"
         Height          =   330
         Left            =   8760
         TabIndex        =   7
         Top             =   780
         Width           =   2640
      End
      Begin VB.Label lblLocalHost 
         Caption         =   "Local Host:"
         Height          =   330
         Left            =   8760
         TabIndex        =   6
         Top             =   360
         Width           =   2640
      End
   End
   Begin VB.CommandButton cmdDisconnectAll 
      Caption         =   "Disconnect all and quit"
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   2655
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2235
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   3942
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483641
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:
' This form will be used to handle an instance of a server.
' Created by Edwin Vermeer
' Website http://siteskinner.com
'
'Credits:
' The (super) SubClass code is from Paul Canton [Paul_Caton@hotmail.com]
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37102&lngWId=1
' Most of the winsock stuff is based on the code from 'Coding Genius'
' see http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=39858&lngWId=1
Option Explicit

'We will handle a server.
Private WithEvents cServer  As clsGenericServer    'Server class
Attribute cServer.VB_VarHelpID = -1

Dim lngSocket As Long  'The socket were we listen on



'Purpose:
' Create a new instance of the server object.
Private Sub Form_Load()
    Set cServer = New clsGenericServer
    ' Put in the listview headers
    ListView1.ColumnHeaders.Add 1, , "Socket Handle"
    ListView1.ColumnHeaders.Add 2, , "Remote Host"
    ListView1.ColumnHeaders.Add 3, , "Remote IP"
    ListView1.ColumnHeaders.Add 4, , "Remote Port"
    ListView1.ColumnHeaders.Add 5, , "Start time"
    ListView1.ColumnHeaders.Add 6, , "Data in"
    ListView1.ColumnHeaders.Add 7, , "Data out"
    ListView1.ColumnHeaders.Add 8, , "Last communication"
End Sub



'Purpose:
' Make sure that all clients are disconnected and unload the server object.
Private Sub Form_Unload(Cancel As Integer)
'Unload the client class - This MUST be done
   cServer.CloseAll
   Set cServer = Nothing
End Sub



'Purpose:
' Disconnect the active (click on one in the list) client.
Private Sub cmdDisconnect_Click()
   ' You have to specify which connection to close
   If lngSocket = 0 Then
     MsgBox "First you have to select a connection!", vbCritical, "Close connection"
   Else
     'Close the socket
      cServer.Connection(lngSocket).CloseSocket
      'Clear data of active connection
      ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
      ClearData lngSocket
   End If
End Sub



'Purpose:
' Send data to the active (click on one in the list) client.
Private Sub cmdSend_Click()
Dim lngLoop As Long
   
   ' You have to specify which connection you want to use
   If lngSocket = 0 Then
     MsgBox "First you have to select a connection!", vbCritical, "Sending data"
     Exit Sub
   End If
   
   'Send the data
   cServer.Connection(lngSocket).Send txtSendData

   ' Go to the coresponding listview item and update it
   For lngLoop = 1 To ListView1.ListItems.Count
      If CLng(ListView1.ListItems(lngLoop)) = lngSocket Then
         ListView1.ListItems(lngLoop).SubItems(6) = ListView1.ListItems(lngLoop).SubItems(6) + Len(txtSendData)
         ListView1.ListItems(lngLoop).SubItems(7) = Now
         Exit For
      End If
   Next

End Sub



'Purpose:
' Send data to the all connected clients.
Private Sub cmdSendAll_Click()
Dim lngLoop As Long
   
   'Go through all connections
   For lngLoop = 1 To ListView1.ListItems.Count
      'Send the data
      cServer.Connection(CLng(ListView1.ListItems(lngLoop))).Send txtSendData
      'Update the listview
      ListView1.ListItems(lngLoop).SubItems(6) = ListView1.ListItems(lngLoop).SubItems(6) + Len(txtSendData)
      ListView1.ListItems(lngLoop).SubItems(7) = Now
   Next

End Sub



'Purpose:
' Just stop.
Private Sub cmdDisconnectAll_Click()
   Unload Me
End Sub



'Purpose:
' Set the server in listening mode on the specified port.
Public Sub Listen(strPort As String)
   cServer.Listen CInt(strPort)
   Me.Caption = "Server listening at port " & CInt(strPort)
End Sub


'Purpose:
' The connection where you click on will be the active connection.
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
   lngSocket = CLng(Item.Text)
    'Get end point information
   Frame1.Caption = "Selected Client " & cServer.Connection(lngSocket).GetRemoteHost & " (" & cServer.Connection(lngSocket).GetRemoteIP & ") on port " & cServer.Connection(lngSocket).GetRemotePort
   lblLocalHost.Caption = "Local Host: " & cServer.Connection(lngSocket).GetLocalHost
   lblLocalIP.Caption = "Local IP: " & cServer.Connection(lngSocket).GetLocalIP
   lblLocalPort.Caption = "Local Port: " & cServer.Connection(lngSocket).GetLocalPort
   lblRemoteHost.Caption = "Remote Host: " & cServer.Connection(lngSocket).GetRemoteHost
   lblRemoteIP.Caption = "Remote IP: " & cServer.Connection(lngSocket).GetRemoteIP
   lblRemotePort.Caption = "Remote Port: " & cServer.Connection(lngSocket).GetRemotePort
   txtDataRecv = ""
End Sub



'Purpose:
' Whatever was set up as the active connection can not be active anymore.
Private Sub ClearData(lngSocketX As Long)
Dim lngLoop As Long
   'Remove it from the list
   For lngLoop = 1 To ListView1.ListItems.Count
      If CLng(ListView1.ListItems(lngLoop)) = lngSocketX Then
         ListView1.ListItems.Remove lngLoop
         Exit For
      End If
   Next
   'Clear the Selected client info
   If lngSocket = lngSocketX Then
      lngSocket = 0
      Frame1.Caption = "Selected Client"
      lblLocalHost.Caption = "Local Host: "
      lblLocalIP.Caption = "Local IP: "
      lblLocalPort.Caption = "Local Port: "
      lblRemoteHost.Caption = "Remote Host: "
      lblRemoteIP.Caption = "Remote IP: "
      lblRemotePort.Caption = "Remote Port: "
      txtDataRecv = ""
   End If
End Sub






'----------------------------------------------------------
' The Server events
'----------------------------------------------------------

'Purpose:
' A client was closed
Private Sub cServer_OnClose(lngSocketX As Long)
   ClearData lngSocketX
End Sub



'Purpose:
' A client wants to connect. Accept it.
Private Sub cServer_OnConnectRequest(lngSocket As Long)
Dim lngNewSocket As Long
    
    'Accept the connection and store the new socket handle
    lngNewSocket = cServer.Accept(lngSocket)
        
    'We use the listbox to hold the info about the new client
    Dim ListHeader    As ListItem
    Set ListHeader = ListView1.ListItems.Add(, , lngNewSocket)
    ListHeader.SubItems(1) = cServer.Connection(lngNewSocket).GetRemoteHost
    ListHeader.SubItems(2) = cServer.Connection(lngNewSocket).GetRemoteIP
    ListHeader.SubItems(3) = cServer.Connection(lngNewSocket).GetRemotePort
    ListHeader.SubItems(4) = Now
    ListHeader.SubItems(5) = 0
    ListHeader.SubItems(6) = 0
    ListHeader.SubItems(7) = Now
    
    'Get end point information
    Me.Caption = "Server " & cServer.Connection(lngNewSocket).GetLocalHost & " (" & cServer.Connection(lngNewSocket).GetLocalIP & ") is listening at port " & cServer.Connection(lngNewSocket).GetLocalPort
    
End Sub


'Purpose:
' This event will be triggered when data has arived.
' This is the location where you will write your server side protocol handler.
' In this case we just log the data and update the statistics.
Private Sub cServer_OnDataArrive(lngSocketX As Long)
Dim strData As String
Dim lngLoop As Long
    
   'Recieve data on the server socket
   cServer.Connection(lngSocketX).Recv strData
    
   ' Go to the coresponding listview item
   For lngLoop = 1 To ListView1.ListItems.Count
      If CLng(ListView1.ListItems(lngLoop)) = lngSocketX Then
         ListView1.ListItems(lngLoop).SubItems(5) = ListView1.ListItems(lngLoop).SubItems(5) + Len(strData)
         ListView1.ListItems(lngLoop).SubItems(7) = Now
         Exit For
      End If
   Next
    
    ' Only show the data if it's the active/selected client
    If lngSocket = lngSocketX Then
       'Log it
       If Len(strData) > 0 Then
          txtDataRecv.Text = txtDataRecv.Text & strData & vbCrLf
          txtDataRecv.SelStart = Len(txtDataRecv.Text)
       End If
    End If
    
End Sub


'Purpose:
' This event is called whenever there was an error.
Private Sub cServer_OnError(lngRetCode As Long, strDescription As String)
    txtDataRecv.Text = txtDataRecv & "*** Error: " & strDescription & vbCrLf
    txtDataRecv.SelStart = Len(txtDataRecv.Text)
End Sub





