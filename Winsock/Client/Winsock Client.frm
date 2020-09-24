VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Client"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Status 
      BackColor       =   &H80000004&
      Height          =   1125
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1080
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   3
      Text            =   "1234"
      Top             =   720
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   1
      Text            =   "255.255.255.255"
      Top             =   240
      Width           =   4575
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Port to connect to:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP address:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Set the Remote Hosts IP Address
Winsock1.RemoteHost = Text1.Text
'Set the Remote Hosts Port Number
Winsock1.RemotePort = Text2.Text
'Tell winsock to connect
Winsock1.Connect
'Disable button 1
Command1.Enabled = False
'Enable button 2
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
'Close the connection
Winsock1.Close
'Enable button  1
Command1.Enabled = True
'Disable button 2
Command2.Enabled = False
End Sub

Private Sub Form_Load()
'Set the text in Text1 to your IP Address
Text1.Text = Winsock1.LocalIP
End Sub
Private Sub Winsock1_Connect()
  'When we are connected tell the user
  MsgBox "You are Connected", vbInformation, "Success!"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  'If an error occurs tell the user
  MsgBox "Error has occurred", vbCritical, "Connect Error"
  'Add the error in the Sattus Box
  Status.Text = Status.Text & Description & " - Error number: " & Number & vbCrLf
  'Enable button  1
  Command1.Enabled = True
  'Disable button 2
  Command2.Enabled = False
End Sub
