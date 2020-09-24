VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmClient 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "http://WwW.YazanMarkabi.Jeeran.Com/"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   13530
   Icon            =   "Client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   13530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer OnlineCounterI 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   8760
   End
   Begin VB.ListBox OnlineCounter 
      Height          =   1620
      ItemData        =   "Client.frx":6852
      Left            =   11160
      List            =   "Client.frx":6859
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ClientsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   8025
      ItemData        =   "Client.frx":6860
      Left            =   10680
      List            =   "Client.frx":6867
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   8520
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8040
      Top             =   8760
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   8520
      Width           =   11895
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12120
      TabIndex        =   0
      Top             =   8520
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtText 
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   11880
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Client.frx":6877
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chat Online With :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   120
      Picture         =   "Client.frx":68F9
      Top             =   120
      Width           =   10500
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sName, sOldName As String

Private Sub cmdSend_Click()
On Error Resume Next
    'send the message to the server
If ClientsList.ListIndex = 0 Then
    tcpClient.SendData (sName & ": " & txtMsg.text & vbNewLine)
Else
    tcpClient.SendData ("<TO<32>" & ClientsList.List(ClientsList.ListIndex) & "<ZaidMarkabi><32>" & sName & ": " & txtMsg.text & vbNewLine)
    'print out a message
    rtText.SelStart = Len(rtText.text)
    rtText.SelText = sName & ": " & txtMsg.text & vbNewLine
    rtText.SelColor = RGB(0, 255, 0)
End If
    txtMsg.text = ""
    txtMsg.SetFocus
End Sub

Private Sub OnlineCounterI_Timer()
On Error Resume Next

Dim i As Integer
For i = 1 To OnlineCounter.ListCount - 1
OnlineCounter.List(i) = OnlineCounter.List(i) - 1
Next
tcpClient.SendData ("LOG<32>" & sName & Space(50 - Len(sName)))

For i = 1 To OnlineCounter.ListCount - 1
If OnlineCounter.List(i) = "-1" Then
ClientsList.RemoveItem i
OnlineCounter.RemoveItem i
End If
Next
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

Timer1.Enabled = False
mnuPref_Click
mnuCnct_Click
ClientsList.ListIndex = 0
OnlineCounterI.Enabled = True
End Sub

Private Sub Status(text As String)
On Error Resume Next

    'print out a status line
    rtText.SelColor = vbRed
    rtText.SelBold = True
    rtText.SelFontSize = 12
    rtText.SelItalic = True
    rtText.SelStart = Len(rtText.text)
    rtText.SelText = text & vbNewLine
    rtText.SelItalic = False
    rtText.SelColor = RGB(0, 255, 0)
    rtText.SelText = rtText.SelText & vbNewLine
End Sub

Private Sub WriteMsg(text As String)
On Error Resume Next

Dim i As Integer
Dim x() As String

Select Case Left(text, 7)
Case Is = "LOG<32>"
  For i = 0 To ClientsList.ListCount - 1
   If ClientsList.List(i) = Trim(Mid(text, 8, 50)) Then
   OnlineCounter.List(i) = "3"
   GoTo 5
   End If
  Next
   ClientsList.AddItem Trim(Mid(text, 8, 50))
   OnlineCounter.AddItem "3"
5:
Case Is = "<TO<32>"
x() = Split(text, "<ZaidMarkabi><32>")
If Right(x(0), Len(x(0)) - 7) = sName Then
text = x(1)
    'print out a message
    rtText.SelStart = Len(rtText.text)
    rtText.SelText = text & vbNewLine
    rtText.SelColor = RGB(0, 255, 0)
End If

Case Else
    'print out a message
    rtText.SelStart = Len(rtText.text)
    rtText.SelText = text & vbNewLine
    rtText.SelColor = RGB(0, 255, 0)
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuCnct_Click()
On Error Resume Next

    Dim IP As String
    
    'make sure you're not already connected
    If tcpClient.State <> sckClosed Then
        MsgBox "Already connected.", vbOKOnly, "Connection Error"
        Exit Sub
    End If
    
    IP = "127.0.0.1"
    
    'check for valid IP
    If IP = "" Then
        Exit Sub
    End If
    
    'Connect to server
    Status ("Connecting to server...")
    Do Until tcpClient.State <> sckClosed
        DoEvents
        tcpClient.Close
        tcpClient.Connect IP, 1337
    Loop
    
End Sub

Private Sub mnuPref_Click()
On Error Resume Next
    'save old name
    sOldName = sName
    
    'Get prefs
    frmPrefs.Visible = True
    frmPrefs.txtName.SetFocus
    Do
        DoEvents
    Loop Until Not frmPrefs.Visible
    
    'check if name is too long
    If Len(sName) > 32 Then
        sName = Left(sName, 32)
    End If
    
    'check if name has changed
    If (sName <> sOldName) And (tcpClient.State <> sckClosed) Then
        tcpClient.SendData (sOldName & " is now called " & sName & vbNewLine)
    End If
    
End Sub

Private Sub tcpClient_Close()
On Error Resume Next
    'notify client of disconnect
    tcpClient.Close
    DoEvents
    Status ("Disconnected")
    cmdSend.Enabled = False
MsgBox "There are need to restart the application !"
Shell App.Path + "\" + App.EXEName + ".EXE", vbNormalFocus
End
End Sub

Private Sub tcpClient_Connect()
On Error Resume Next
    'notify client of connect
    Status ("Connected")
    DoEvents
    cmdSend.Enabled = True
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    Dim text As String
    tcpClient.GetData text
    WriteMsg (text)
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If (KeyAscii = 13) And (cmdSend.Enabled) Then
        cmdSend_Click
    End If
End Sub

