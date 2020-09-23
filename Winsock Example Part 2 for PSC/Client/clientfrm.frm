VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ClientFrm 
   Caption         =   "Client"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sock1 
      Left            =   2640
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton bntSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton bntExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton bntConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Tag             =   "Connect"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   6120
      Width           =   3735
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "123"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtLog 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Winsock Example by VirusFree - http://www.phoenixbit.com"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   -240
      TabIndex        =   9
      Top             =   6600
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Remote Host Port :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Remote Host IP :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "ClientFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bntConnect_Click()
On Error GoTo t

'sock1 is the name of our Winsock ActiveX Control

sock1.Close 'we close it in case it was trying to connect

'txtIP is the textbox holding the host IP
'txtIP can contain both hostnames ( like www.google.com ) or IPs ( like 127.0.0.1 )
sock1.RemoteHost = txtIP    'set the remote host to the ip we wrote
                            'in the txtIP textbox

'txtPort is the textbox holding the Port number
sock1.RemotePort = txtPort  'set the port we want to connect to
                            '( the server must be listening on this port too)
                            
                            
sock1.Connect               'try to connect


Exit Sub
t:
MsgBox "Error : " & Err.Description, vbCritical
End Sub

Private Sub bntExit_Click()
MsgBox "Winsock example by VirusFree - http://www.phoenixbit.com"
End
End Sub

Private Sub bntSend_Click()
On Error GoTo t
'we want to send the contents of txtSend textbox

sock1.SendData txtSend  'trasmits the string to host


'we have send the data to the server by we
'also need to add them to our Chat Buffer
'so we can se what we wrote
txtLog = txtLog & "Client : " & txtSend & vbCrLf

'and then we clear the txtSend textbox so the
'user can write the next message
txtSend = ""

'error handling
'( for example , we will get an error if try to send
'  any data without being connected )
Exit Sub
t:
MsgBox "Error : " & Err.Description
sock1_Close   'close the connection
End Sub

Private Sub sock1_Close()
'handles the closing of the connection

sock1.Close  'close connection

txtLog = txtLog & "*** Disconnected" & vbCrLf

End Sub

Private Sub sock1_Connect()
'txtLog is the textbox used as our
'chat buffer.

'sock1.RemoteHost returns the hostname( or ip ) of the host
'sock1.RemoteHostIP returns the IP of the host

txtLog = "Connected to " & sock1.RemoteHostIP & vbCrLf

End Sub

Private Sub sock1_DataArrival(ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

Dim dat As String     'where to put the data

sock1.GetData dat, vbString   'writes the new data in our string dat ( string format )

'add the new message to our chat buffer
txtLog = txtLog & dat & vbCrLf

End Sub

Private Sub sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this event is to handle any kind of errors
'happend while using winsock

'Number gives you the number code of that specific error
'Description gives you string with a simple explanation about the error

'append the error message in the chat buffer
txtLog = txtLog & "*** Error : " & Description & vbCrLf

'and now we need to close the connection
sock1_Close

'you could also use sock1.close function but I
'prefer to call it within the Sock1_Close functions that
'handles the connection closing in general

End Sub
