VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Server"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Stop Listening"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "1234"
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Listen"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Open the port number from the number in Text1
Winsock1.LocalPort = Text1.Text
'Tell winsock to listen for any connections
Winsock1.Listen
'Disable button 1
Command1.Enabled = False
'Enable button 2
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
'Close the Connection
Winsock1.Close
'Enable button  1
Command1.Enabled = True
'Disable button 2
Command2.Enabled = False
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
  'Close the socket from the listening state
  Winsock1.Close
  'Accepts the request Connect
  Winsock1.Accept requestID
  'Say were connected
  MsgBox "Connected to " & Winsock1.RemoteHostIP
End Sub
