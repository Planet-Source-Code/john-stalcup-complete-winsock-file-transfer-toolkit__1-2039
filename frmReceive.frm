VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmReceive 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Receiving File"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
   Begin MSWinsockLib.Winsock sckReceive 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label compLabel 
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   780
      Width           =   4395
   End
End
Attribute VB_Name = "frmReceive"
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
Private fileNum As Double

Private Sub cancel_Click()

sckSystem.SendData CANCEL_TRANSFER
Unload Me

End Sub

Private Sub Form_Activate()
'
End Sub

Private Sub Form_Initialize()
'
End Sub

Private Sub Form_Load()

'this defaults port to connect on to 43597 incase it is not set from outside of this form
If hostPort = 0 Then
    hostPort = 43597
End If

'prepare progress bar
ProgressBar1.Min = 0
ProgressBar1.value = ProgressBar1.Min
ProgressBar1.Visible = True

'bind system & send controls together
'this one is udp
sckSystem.Close
sckSystem.RemoteHost = hostIP
sckSystem.LocalPort = hostPort ' Port to monitor
sckSystem.RemotePort = hostPort ' Port to connect to.
sckSystem.Bind

'this one is a tcp/ip control
sckReceive.Close
sckReceive.LocalPort = hostPort + 1 ' Port to monitor
sckReceive.Listen

End Sub

Private Sub sckReceive_Close()

    Close fileNum
    MsgBox "Transfer of " & nameOfFile & " completed successfully."
    Unload Me

End Sub

Private Sub sckReceive_ConnectionRequest(ByVal requestID As Long)

    ' Check if the control's State is closed. If not,
    ' close the connection before accepting the new
    ' connection.
    If sckReceive.State <> sckClosed Then sckReceive.Close
    ' Accept the request with the requestID
    ' parameter.
    sckReceive.Accept requestID

End Sub

Private Sub sckReceive_DataArrival(ByVal bytesTotal As Long)

On Error GoTo ErrorHandler

    Dim temp As String
    sckReceive.GetData temp
    Put #fileNum, , temp
    fileLength = LOF(fileNum)
        
    'update progress bar
    sizeOfFileSent = sizeOfFileSent + bytesTotal
    On Error GoTo endIt
    ProgressBar1.value = sizeOfFileSent
    compLabel.Caption = sizeOfFileSent & " of " & sizeOfFile & " sent. " & Int(sizeOfFileSent / sizeOfFile * 100) & "%"
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occured while saving " & CommonDialog1.FileTitle & ". File Transfer being canceled.", vbOKOnly, "IO Error"
    cancel_Click
endIt:
End Sub

Private Sub sckSystem_DataArrival(ByVal bytesTotal As Long)

Dim temp As String
sckSystem.GetData temp, vbString

Dim command As String, value As String
command = Mid(temp, 1, 1)
value = Mid(temp, 2, Len(temp) - 1)

Select Case command
    Case FILE_SIZE
        sizeOfFile = value
        'prepare progress bar
        ProgressBar1.Max = sizeOfFile
        queryAcceptDload
    Case USER_NAME
        userName = value
        queryAcceptDload
    Case FILE_NAME
        nameOfFile = value
        Me.Caption = "Receiving " & nameOfFile
        queryAcceptDload
    Case CANCEL_TRANSFER
        stopSending
'    Case END_TRANSFER
'        Close fileNum
'        MsgBox "Transfer of " & nameOfFile & " completed successfully."
'        Unload Me
End Select

End Sub

Private Sub stopSending()

    MsgBox "User has canceled the file transfer.", vbOKOnly, "File Transfer Canceled"
    Unload Me

End Sub

Private Sub queryAcceptDload()
CommonDialog1.CancelError = True
On Error GoTo endIt

    If sizeOfFile <> 0 And nameOfFile <> "" And userName <> "" Then
    
        Dim temp
        temp = MsgBox("Would you like to accept " & nameOfFile & " (" & sizeOfFile & " bytes) from " & userName & "?", vbYesNo, "Transfer " & nameOfFile & "?")
        If temp = vbYes Then
            CommonDialog1.ShowSave
            'open the file
            fileNum = FreeFile
            Open CommonDialog1.fileName For Binary Access Write As fileNum
            'tell other end to begin transfer
            sckSystem.SendData ACCEPT_TRANSFER
                        
        Else
            cancel_Click
        End If
        
    End If

Exit Sub
endIt:
cancel_Click
End Sub
