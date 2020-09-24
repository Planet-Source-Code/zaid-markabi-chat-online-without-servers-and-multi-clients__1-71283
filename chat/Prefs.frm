VERSION 5.00
Begin VB.Form frmPrefs 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Enter Your Name"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "My Name"
      Top             =   360
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Public Sub SetOnTop(ByVal hwnd As Long, ByVal bSetOnTop As Boolean)
On Error Resume Next
Dim lR As Long
If bSetOnTop Then
   lR = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
Else
   lR = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End If
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
If Not txtName.text = "" Then
    frmClient.sName = txtName.text
    frmPrefs.Visible = False
    frmClient.Caption = frmClient.Caption + txtName.text + ".php"
Open "Data.txt" For Output As #1
Write #1, txtName.text
Close #1
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
   SetOnTop Me.hwnd, True

Dim x As String
Open "Data.txt" For Input As #1
Input #1, x
txtName.text = x
Close #1
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        cmdOK_Click
    End If
End Sub

