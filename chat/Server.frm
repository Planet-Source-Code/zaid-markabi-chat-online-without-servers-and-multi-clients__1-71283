VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmServer 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  www.YazanMarkabi.Jeeran.com"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1440
      Top             =   600
   End
   Begin MSWinsockLib.Winsock tcpAccept 
      Index           =   0
      Left            =   1440
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpListen 
      Left            =   1440
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "written by : Zaid Markabi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Image Image3 
      Height          =   2160
      Left            =   1800
      Picture         =   "Server.frx":0000
      Top             =   120
      Width           =   3600
   End
   Begin VB.Image Image2 
      Height          =   1365
      Left            =   0
      Picture         =   "Server.frx":19542
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   1365
      Left            =   120
      Picture         =   "Server.frx":1F7A0
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo 5
    'start listening for connections
    tcpListen.LocalPort = listenPortNum
    tcpListen.Listen
    
    'create the slots for the accept sockets
    If Not InitAccept Then
        MsgBox "Error creating sockets.", vbCritical, "Socket Error!"
        frmClient.Show
        Unload Me
    End If
    
    'Tell User that the log is started
    Log ("Application Started")
    Log ("Awaiting Connections")
5:
    frmClient.Show
End Sub

Private Sub tcpAccept_Close(Index As Integer)
On Error Resume Next
    'close connection
    tcpAccept(Index).Close
    'set the port to open
    bSocketUsed(Index) = False
End Sub

Private Sub tcpAccept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
    Dim text As String
    Dim i As Integer
    
    'get data
    tcpAccept(Index).GetData text
    
    'check for disconnects
    If text = ("disband") Then
        tcpAccept_Close (Index)
        Exit Sub
    End If
    
    'send it out
    For i = 1 To maxCon
        If bSocketUsed(i) Then
            tcpAccept(i).SendData (text)
            DoEvents
        End If
    Next i
End Sub

Private Sub tcpListen_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
    Dim portNum As Integer
    
    'find out if there's a free socket
    portNum = GetFreeSocket
    
    If portNum = 0 Then
        'tell the client that the server is full
        tcpAccept(0).Accept (requestID)
        DoEvents
        tcpAccept(0).SendData "Sorry, maximum number of connections reached!"
        DoEvents
        'close the client's connection
        tcpAccept(0).Close
    Else
        'accept the client using the accepting winsock
        tcpAccept(portNum).Accept (requestID)
        DoEvents
        'send welcome message
        tcpAccept(portNum).SendData ("Connection accepted, welcome." & vbNewLine)
        'set port to used
        bSocketUsed(portNum) = True
    End If
End Sub

Private Sub Timer1_Timer()
Me.Hide
Timer1.Enabled = False
End Sub
