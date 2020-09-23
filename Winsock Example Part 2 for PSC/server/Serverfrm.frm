VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ServerFrm 
   Caption         =   "Server"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sock1 
      Index           =   0
      Left            =   4320
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton bntExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton bntListen 
      Caption         =   "Start Listening"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Tag             =   "Connect"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "123"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLog 
      Height          =   5535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Channel Log"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Winsock Example by VirusFree - http://www.phoenixbit.com"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   -240
      TabIndex        =   5
      Top             =   6360
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Listen on Port :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "ServerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SocketCounter As Long

'this is to get you to the tutorial page
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub bntListen_Click()
On Error Resume Next
'close and unload all previous sockets
For n = 1 To SocketCounter
    sock1(n).Close
    Unload sock1(n)
Next

On Error GoTo t

'sock1(0) is the name of our Winsock ActiveX Control

sock1(0).Close 'we close it in case it listening before


'txtPort is the textbox holding the Port number
sock1(0).LocalPort = txtPort  'set the port we want to listen to
                              '( the client will connect on this port too)
                            
                            
sock1(0).Listen                'Start Listening


txtLog = "Listening on Port " & txtPort

Exit Sub
t:
MsgBox "Error : " & Err.Description, vbCritical
End Sub

Private Sub bntExit_Click()
End
End Sub

Private Sub Form_Load()
MsgBox "This is the example code from the tutorial." & vbCrLf & "You can find the full tutorial at www.phoenixbit.com"
'shell execute will open the url of the tutorial
ShellExecute Me.hwnd, "Open", "http://www.phoenixbit.com", 0, App.Path, 1
End Sub

Private Sub sock1_Close(Index As Integer)
'handles the closing of the connection

sock1(Index).Close  'close connection

Unload sock1(Index) 'unload control

txtLog = txtLog & "Client" & Index & " -> *** Disconnected" & vbCrLf

End Sub

Private Sub sock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'txtLog is the textbox used as our log.

'this event is triggered when a client try to connect on our host
'we must accept the request for the connection to be completed,
'but we will create a new control and assign it to that, so
'sock1(0) will still be listening for connection but
'sock1(SocketCounter) , our new sock , will handle the current
'request and the general connection with the client

'increase counter
SocketCounter = SocketCounter + 1

'this will create a new control with index equal to SocketCounter
Load sock1(SocketCounter)

'with this we accept the connection and we are now connected to
'the client and we can start sending/receiving data
sock1(SocketCounter).Accept requestID

'add to the log
txtLog = "Client Connected. IP : " & sock1(0).RemoteHostIP & " , Client Nick : Client" & sockcounter & vbCrLf

'tell our client his assigned nickname
sock1(SocketCounter).SendData "Your Nick is ""Client" & SocketCounter & """"

End Sub

Private Sub sock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

Dim dat As String     'where to put the data

sock1(Index).GetData dat, vbString   'writes the new data in our string dat ( string format )

'add the new message to our chat buffer
txtLog = txtLog & "Client" & Index & " : " & dat & vbCrLf

'now the client says something, wich arrived at the server...
'the server must now redistibute this message to all other connected
'clients...
On Error Resume Next    'Error Handler
For n = 1 To SocketCounter
    If Not n = Index Then   'we don't want to send the msg back to the sender :)
        If sock1(n).State = sckConnected Then   'if socket is connected
            sock1(n).SendData "Client" & Index & " : " & dat
        End If
    End If
Next

End Sub

Private Sub sock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this event is to handle any kind of errors
'happend while using winsock

'Number gives you the number code of that specific error
'Description gives you string with a simple explanation about the error

'append the error message in the chat buffer
txtLog = txtLog & "*** Error ( Client" & Index & ") : " & Description & vbCrLf

'and now we need to close the connection
sock1_Close Index

'you could also use sock1(Index).close function but i
'prefer to call it within the sock1_Close functions that
'handles the connection closing in general


End Sub


