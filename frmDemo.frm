VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client Server Demo"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Vote !"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server"
      Height          =   1620
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   3210
      Begin VB.TextBox txtLocalPort 
         Height          =   285
         Left            =   1260
         TabIndex        =   8
         Text            =   "8080"
         Top             =   420
         Width           =   1590
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "New Server"
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Local Port"
         Height          =   330
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   1275
      End
   End
   Begin VB.Frame frmClient 
      Caption         =   "Client"
      Height          =   2025
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3210
      Begin VB.CommandButton cmdConnect 
         Caption         =   "New Client"
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtHostName 
         Height          =   285
         Left            =   1260
         TabIndex        =   2
         Text            =   "localhost"
         Top             =   420
         Width           =   1590
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Text            =   "8080"
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Host Name"
         Height          =   330
         Left            =   210
         TabIndex        =   5
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Remote Port"
         Height          =   330
         Left            =   210
         TabIndex        =   4
         Top             =   840
         Width           =   1275
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "If you like this code the plese vote now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   13
      Top             =   2280
      Width           =   8295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":00AA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   12
      Top             =   1680
      Width           =   8295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":0157
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   11
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":0251
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   10
      Top             =   840
      Width           =   8295
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:
' This demo will show how to create multiple instances of a generic client and server.
' Created by Edwin Vermeer
' Website http://siteskinner.com
'
'Credits:
' The (super) SubClass code is from Paul Canton [Paul_Caton@hotmail.com]
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37102&lngWId=1
' Most of the winsock stuff is based on the code from 'Coding Genius'
' see http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=39858&lngWId=1

Option Explicit

' The functions used for opening a URL
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long



'Purpose:
' Open a new Client form and start up the connection
Private Sub cmdConnect_Click()
Dim Client As New frmClient

  Client.Show
  Client.Connect txtHostName, txtPort

End Sub



'Purpose:
' Open a new Server form and start up the connection
Private Sub cmdListen_Click()
Dim Server As New frmServer

  Server.Show
  Server.Listen txtLocalPort
  
End Sub



'Purpose:
' This will redirect to the correct planet source code page.
Private Sub Command1_Click()
   RunThisURL "http://siteskinner.com/psc.asp"
End Sub


'Purpose:
' This will open any URL with the default browser.
Private Sub RunThisURL(strURL As String)
Dim strFileName    As String
Dim strDummy       As String
Dim strBrowserExec As String * 255
Dim lngRetVal      As Long
Dim intFileNumber  As Integer

  ' Create a temporary HTM file
  strBrowserExec = Space(255)
  strFileName = "~TempBrowserCheck.HTM"
  intFileNumber = FreeFile
  Open strFileName For Output As #intFileNumber
    Write #intFileNumber, "<HTML> <\HTML>"
  Close #intFileNumber

  ' Find the default browser.
  lngRetVal = FindExecutable(strFileName, strDummy, strBrowserExec)
  strBrowserExec = Trim$(strBrowserExec)

  ' If an application is found, launch it!
  If lngRetVal <= 32 Or IsEmpty(strBrowserExec) Then
    MsgBox "Could not find your Browser", vbExclamation, "Browser Not Found"
  Else
    lngRetVal = ShellExecute(App.hInstance, "open", strBrowserExec, strURL, strDummy, 1)
    If lngRetVal <= 32 Then
      MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
    End If
  End If
 
  ' remove the temporary file
  Kill strFileName

End Sub


