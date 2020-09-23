Attribute VB_Name = "xferInterface"
Public Function sendFile_01(fileName As String, filePath As String, hostIP As String, hostPort As Double, localUserName As String)

Dim a As New frmSend

a.nameOfFile = fileName
a.pathToFile = filePath
a.hostIP = hostIP
a.hostPort = hostPort
a.userName = localUserName

a.Show

End Function

Public Function receiveFile_01(hostIP As String, hostPort As Double)

Dim a As New frmReceive

a.hostIP = hostIP
a.hostPort = hostPort

a.Show

End Function
