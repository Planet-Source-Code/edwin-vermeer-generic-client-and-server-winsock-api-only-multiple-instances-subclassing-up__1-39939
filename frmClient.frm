VERSION 5.00
Begin VB.Form frmClient 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   435
      Left            =   6360
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtSendData 
      Height          =   1080
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   6255
   End
   Begin VB.TextBox txtDataRecv 
      Height          =   1575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1200
      Width           =   7335
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   435
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLocalHost 
      Caption         =   "Local Host:"
      Height          =   330
      Left            =   7560
      TabIndex        =   8
      Top             =   120
      Width           =   2640
   End
   Begin VB.Label lblLocalIP 
      Caption         =   "Local IP:"
      Height          =   330
      Left            =   7560
      TabIndex        =   7
      Top             =   540
      Width           =   2640
   End
   Begin VB.Label lblLocalPort 
      Caption         =   "Local Port:"
      Height          =   330
      Left            =   7560
      TabIndex        =   6
      Top             =   960
      Width           =   2640
   End
   Begin VB.Label lblRemoteHost 
      Caption         =   "Remote Host"
      Height          =   330
      Left            =   7560
      TabIndex        =   5
      Top             =   1380
      Width           =   2640
   End
   Begin VB.Label lblRemoteIP 
      Caption         =   "Remote IP:"
      Height          =   330
      Left            =   7560
      TabIndex        =   4
      Top             =   1800
      Width           =   2640
   End
   Begin VB.Label lblRemotePort 
      Caption         =   "Remote Port:"
      Height          =   330
      Left            =   7560
      TabIndex        =   3
      Top             =   2220
      Width           =   2640
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:
' This form will be used to handle an instance of a client.
' Created by Edwin Vermeer
' Website http://siteskinner.com
'
'Credits:
' The (super) SubClass code is from Paul Canton [Paul_Caton@hotmail.com]
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37102&lngWId=1
' Most of the winsock stuff is based on the code from 'Coding Genius'
' see http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=39858&lngWId=1

Option Explicit

'We will handle a client
Private WithEvents cClient  As clsGenericClient
Attribute cClient.VB_VarHelpID = -1



'Purpose:
' Create a new instance of the client
Private Sub Form_Load()
    Set cClient = New clsGenericClient
End Sub


'Purpose:
' Unload the client class - This MUST be done
Private Sub Form_Unload(Cancel As Integer)
    cClient.Connection.CloseSocket
    Set cClient = Nothing
End Sub


'Purpose:
' Start the connection to the specified server
Public Sub Connect(strHost As String, strPort As String)
   cClient.Connect strHost, CInt(strPort)
End Sub


'Purpose:
'Send the data
Private Sub cmdSend_Click()
   cClient.Connection.Send txtSendData
End Sub


'Purpose:
'Just stop
Private Sub cmdDisconnect_Click()
   Unload Me
End Sub







'----------------------------------------------------------
' The Client events
'----------------------------------------------------------


'Purpose:
' This event is called by the client object when the connection was closed by the server.
' There is no connection anymore so we can quit.
Private Sub cClient_OnClose()
    Unload Me
End Sub



'Purpose:
' This event is called when the connection to the server was successfull.
Private Sub cClient_OnConnect()
'We are connected. Just show som information about the connection
    
    'Add log
    txtDataRecv.Text = txtDataRecv.Text & "*** Connected ***" & vbCrLf
    txtDataRecv.SelStart = Len(txtDataRecv.Text)
    
    'Show the end point information
    Me.Caption = "Client connected to " & cClient.Connection.GetRemoteHost & "(" & cClient.Connection.GetRemoteIP & ")" & " on port " & cClient.Connection.GetRemotePort
    lblLocalHost.Caption = "Local Host: " & cClient.Connection.GetLocalHost
    lblLocalIP.Caption = "Local IP: " & cClient.Connection.GetLocalIP
    lblLocalPort.Caption = "Local Port: " & cClient.Connection.GetLocalPort
    lblRemoteHost.Caption = "Remote Host: " & cClient.Connection.GetRemoteHost
    lblRemoteIP.Caption = "Remote IP: " & cClient.Connection.GetRemoteIP
    lblRemotePort.Caption = "Remote Port: " & cClient.Connection.GetRemotePort
    
End Sub



'Purpose:
' this event will be called when data has arrived from the server.
' Here is where your client side protocol handler should be.
' For this simple demo whe just log the data.
Private Sub cClient_OnDataArrive()

  Dim strData As String
    
    'Recieve the data
    cClient.Connection.Recv strData
    
    'Log it
    If Len(strData) > 0 Then
        txtDataRecv.Text = txtDataRecv.Text & strData & vbCrLf
        txtDataRecv.SelStart = Len(txtDataRecv.Text)
    End If
    
End Sub



'Purpose:
'There was an error in the client class
Private Sub cClient_OnError(lngRetCode As Long, strDescription As String)
    'Log it
    txtDataRecv.Text = txtDataRecv & "*** Error: " & strDescription
    txtDataRecv.SelStart = Len(txtDataRecv.Text)
End Sub





