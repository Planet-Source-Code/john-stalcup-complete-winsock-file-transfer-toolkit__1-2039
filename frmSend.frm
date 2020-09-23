VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmSend 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sending File"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   435
      Left            =   300
      TabIndex        =   1
      Top             =   240
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   767
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1800
      TabIndex        =   0
      Top             =   1140
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock sckSystem 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock sckSend 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label compLabel 
      Caption         =   "Waiting For Other Side To Accept Transfer..."
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   780
      Width           =   4395
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'file data
Public sizeOfFile As Double
Public sizeOfFileSent As Double
Public nameOfFile As String
Public pathToFile As String
Public userName As String

'specify what host to connect to
Public hostIP As String
Public hostPort As Double

'privates
Private quitNow As Boolean

Private Sub cancel_Click()

sckSystem.Close
sckSystem.Bind
sckSystem.SendData CANCEL_TRANSFER
Unload Me
quitNow = True

End Sub

Private Sub Form_Load()

'this defaults port to connect on to 43597 incase it is not set from outside of this form
If hostPort = 0 Then
    hostPort = 43597
End If

Me.Caption = "Sending " & nameOfFile

'find the file size
sizeOfFile = FileLen(pathToFile)

'prepare progress bar
ProgressBar1.Max = sizeOfFile
ProgressBar1.Min = 0
ProgressBar1.value = ProgressBar1.Min
ProgressBar1.Visible = True

'bind sck controls
sckSystem.Close
sckSystem.RemoteHost = hostIP
sckSystem.LocalPort = hostPort ' Port to monitor
sckSystem.RemotePort = hostPort ' Port to connect to.
sckSystem.Bind

'this one is tcp/ip
sckSend.RemoteHost = hostIP
sckSend.RemotePort = hostPort + 1 ' Port to connect to.

'send initialization information
sckSystem.SendData FILE_NAME & nameOfFile
sckSystem.SendData FILE_SIZE & sizeOfFile
sckSystem.SendData USER_NAME & userName

End Sub

Private Sub sckSystem_DataArrival(ByVal bytesTotal As Long)

Dim temp As String
sckSystem.GetData temp, vbString

Dim command As String
command = Mid(temp, 1, 1)

Select Case command
    Case CANCEL_TRANSFER
        stopSending
    Case ACCEPT_TRANSFER
        DoEvents
        sckSend.Connect
        Do Until sckSend.State = sckConnected ' Wait until connected
            DoEvents
        Loop

        SendFile pathToFile

'        sckSystem.SendData END_TRANSFER
        MsgBox "Transfer Complete"
        Unload Me
        
End Select

End Sub

Private Sub stopSending()
    
    quitNow = True
    MsgBox "User has canceled the file transfer.", vbOKOnly, "File Transfer Canceled"
    Unload Me

End Sub

'*******************************************************************
' Credit:       Dan Evans <devans@jrl.com> (with a few mods my me, John Stalcup 6/5/99)
' Function:     SendFile()
' Purpose:      Send a file via network
' Parameters:   Full path and file name of data to send
' Returns:      True on success, False on error
' Notes:        The socket should already be established
'*******************************************************************
Public Function SendFile(fileName As String) As Boolean
        Dim hIn, fileLength, ret
        Dim temp As String
        Dim blockSize As Long
        blockSize = 2048                                '// Set your read buffer size here

On Error GoTo ErrorHandler

        hIn = FreeFile
        Open fileName For Binary Access Read As hIn
        fileLength = LOF(hIn)
        
        Do Until EOF(hIn)
                ' Adjust blocksize at end so we don't read too much data
                If fileLength - Loc(hIn) <= blockSize Then
                        blockSize = fileLength - Loc(hIn) + 1
                End If
                temp = Space$(blockSize)        '// Allocate the read buffer
                Get hIn, , temp                 '// Read a block of data
                ret = DoEvents()                '// Check for cancel button event etc.
                If quitNow Then Exit Function
                sckSend.SendData temp           '// Off it goes
                
                'update progress bar
                sizeOfFileSent = sizeOfFileSent + blockSize
                On Error GoTo endIt             '//
                ProgressBar1.value = sizeOfFileSent
                compLabel.Caption = sizeOfFileSent & " of " & sizeOfFile & " sent. " & Int(sizeOfFileSent / sizeOfFile * 100) & "%"
        Loop

        sckSend.Close   'this severes the data connection, causing the client to save/end the file

        Close hIn
        SendFile = True
        Exit Function

ErrorHandler:                                           '// Always close the file handle
        Close hIn
        SendFile = False
endIt:
End Function
