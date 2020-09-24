Attribute VB_Name = "Server"
'All of the general server functions go here.
Option Explicit

Public Const maxCon = 50
Public Const listenPortNum = 1337

Public bSocketUsed(maxCon) As Boolean


'Create the sockets for accepting the connections
Public Function InitAccept() As Boolean
On Error Resume Next
    Dim i As Integer
    
    For i = 1 To maxCon
        'load the sockets
        Load frmServer.tcpAccept(i)
        'set as unused
        bSocketUsed(i) = False
    Next i
    
    'all is well, nothing is ruined
    InitAccept = True
    Exit Function
    
err:
    'flagrant system error
    InitAccept = False
End Function

'find unused socket
Public Function GetFreeSocket() As Integer
On Error Resume Next
    Dim i As Integer
    
    'Find unused socket
    For i = 1 To maxCon
        If Not bSocketUsed(i) Then
            GetFreeSocket = i
            Exit Function
        End If
    Next i
    
    'no socket found
    GetFreeSocket = 0
End Function

