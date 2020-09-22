VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSendChat 
      Caption         =   "Send Chat"
      Default         =   -1  'True
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   3855
   End
   Begin VB.TextBox txtChat 
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   4695
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select File"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send File"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Timer tmrKBps 
      Interval        =   1000
      Left            =   4800
      Top             =   4680
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   1440
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5160
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblKBps 
      Alignment       =   2  'Center
      Caption         =   "KBps:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   4455
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFriend As String 'holds servers name
Dim strMyName As String 'holds your name
Dim strFileName As String 'holds the name of the file u are receiving
Dim blnSend As Boolean 'if false, it wont send; if true it will
Dim strSize As String 'holds the size of the file
Dim strSoFar As String 'a var for calculating the KBps
Dim strBlock As String 'holds the data you are going to send
Dim strLOF As String 'holds the lenght of the file

Private Sub cmdSendChat_Click()
    If Trim(txtSend.Text) = "" Then Exit Sub 'prevents someone trying to send nothing
    Winsock.SendData "Chat" & txtSend.Text 'sends the text to the chat
    txtChat.SelStart = Len(txtChat) 'put focus on the chat at the end so it is entered in the right place
    txtChat.SelText = strMyName & ":" & vbTab & txtSend.Text & vbCrLf 'puts the text in the chat
    txtSend.Text = "" 'clears the textbox u type in
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock.Close 'closes winsock so program can end
    End 'closes program
End Sub

Private Sub tmrKBps_Timer()
On Error Resume Next 'prevents error
    lblKBps.Caption = "Transfering at: " & Format(strSoFar / 1000, "###0.0") & " / KBps" 'calculates the KBps
    strSoFar = 0 'resets it so it can be calculated again
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
    If Winsock.State <> sckClosed Then Winsock.Close 'closes winsock and allows it to accept a connection
    Winsock.Accept requestID 'allows new connection
    DoEvents 'hahaha
    Winsock.SendData "Nick" & frmConnect.txtName.Text 'sends ur name to client
    DoEvents 'keeps everything running smooth
    strMyName = frmConnect.txtName 'saves ur name into memory
    Me.Show 'shows frmserver
    Unload frmConnect 'hmm what does that do?
End Sub

Private Sub cmdSelect_Click()
    CD.Flags = cdlOFNFileMustExist 'wont let someone open a file not there!
    CD.Filter = "All Files (*.*)|*.*" 'filters all files to be shown!!!!!!!
    CD.ShowOpen 'shows the open dialog box
    txtFile.Text = CD.FileName 'puts the path of the file into the txtfile
    
End Sub

Private Sub cmdSend_Click()
On Error Resume Next 'prevents error
    If txtFile.Text = "" Then Exit Sub 'wont let someone send something not there
    strFileName = "" 'resets the filename
    blnSend = False 'sets it to false
    strSize = "" 'resets the size

    Open txtFile.Text For Binary As #1 'opens the file u want to send so it can be read
    strLOF = LOF(1) 'saves length of file into memory
    Winsock.SendData "Name" & CD.FileTitle & ":" & strLOF 'sends server the name of the file and its length
    DoEvents 'needed  =D

    Do While blnSend = False 'keeps looping until the server accepts the request
        DoEvents 'lets u do stuff while in loop
    Loop 'goes back to Do WHile statement
    
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next 'prevents error
Dim strData As String 'holds data for select case
Dim strData2 As String 'holds data
    Call Winsock.GetData(strData, vbString) 'gets the data sent by the client
    strData2 = Mid(strData, 5) 'gets data
    strData = Left(strData, 4) 'gets data for select case
    Select Case strData 'goes to the right case depending on strData
        Case "File" 'a file transfer is in progress
            Put 1, , strData2 'puts data into file
            PBar.Value = PBar.Value + bytesTotal 'shows how much is done so far
            strSoFar = strSoFar + bytesTotal 'calculates KBps
            Winsock.SendData "OKOK keep sending!" 'tells them ur done with the data and u want some more!
            DoEvents ' =D
        Case "Name" 'client has sent u the filename and is ready to begin transfer
            Dim intX As Integer 'holds position if :
            intX = InStr(1, strData2, ":", vbTextCompare) 'gets position of :
            strSize = Mid(strData2, intX + 1) 'holds the filesize
            PBar.Max = strSize 'sets up the progressbar
            strData = Mid(strData2, 1, intX - 1) 'holds filename
            strFileName = strData 'puts filename into memory
            Dim strResponse As String 'holds either a vbYEs or vbNo
            strResponse = MsgBox(strFriend & " wants to send you [" & strFileName & "].  Do you wish to receive this file?", vbYesNo, "File Exchange Requested") '<=- easy to understand
            If strResponse = vbYes Then 'if they said yes
                Dim strType As String 'holds the type of file
                strType = Right(strFileName, 3) 'gets the type of file
                CD.FileName = strFileName 'sets the filename into the commondialog box
                CD.Filter = "File Type (*." & strType & ")|*." & strType 'sets the filter to the filetype
                CD.Flags = cdlOFNOverwritePrompt 'asks u if u want to overwrite file
                CD.ShowSave 'shows the save commondialog box
                Open CD.FileName For Binary As #1 'opens a file with the name and path u want
                Winsock.SendData "OKOK i want the file" 'tell client u want the damn file
                DoEvents 'yes another doevents
                Me.Enabled = False 'disables to form to PREVENT ERROR!!!!!!!!!!
            ElseIf strResponse = vbNo Then 'if they say no
                Winsock.SendData "Nope dont want it!" 'tell em u dont want their crap!
                DoEvents 'hmmm
            End If 'ok enough of that madness
        Case "Stop" 'the file exchange has ended
            Close #1 'closes the file
            'resets the progressbar
            PBar.Value = 0
            PBar.Max = 1
            '=====================
            Me.Enabled = True 'reenables the form!
        Case "Nick" 'client has sent u their name
            strFriend = strData2 'saves their name into memory
        Case "Nope" 'tells u that they declined ur request to give em a file
            MsgBox strFriend & " declined your file transfer request.", vbInformation, "File Transfer Canceled!" '<=- easy to get again
            Close #1 'closes the file
            'stops the loops that was waiting for the boolean value to be true
            Do
            DoEvents
            Loop
            '==========================
        Case "OKOK" 'tells u they want more of the file
            blnSend = True 'tells u to keep sending
            Me.Enabled = False 'keeps form disabled
            PBar.Max = strLOF 'sets progressbar max to filesize
            If Not EOF(1) Then 'does this if not the end of the file
                If strLOF - Loc(1) < 2040 Then 'if you are at the last chunk of data
                    strBlock = Space$(strLOF - Loc(1)) 'sets the block size to the size of the data (cause its less!)
                    Get 1, , strBlock 'gets data
                    Winsock.SendData "File" & strBlock 'sends data
                    DoEvents ' =/
                    PBar.Value = PBar.Value + Len(strBlock) 'sets progressbar
                    strSoFar = strSoFar + (strLOF - Loc(1)) 'sets KBps
                    Winsock.SendData "Stop the maddness!" 'tells client THE TRANSFER IS ENDED!
                    Close #1 'closes file
                    'resets the progressbar
                    PBar.Max = 1
                    PBar.Value = 0
                    '====================
                    Me.Enabled = True 'reenables the form
                Else 'if not the last chunk
                    strBlock = Space$(2040) 'sets block up to receive only 2040 bytes of data
                End If
                strSoFar = strSoFar + 2040 'calculates KBps
                Get 1, , strBlock 'gets data
                Winsock.SendData "File" & strBlock 'sends data
                DoEvents 'hmmmm  once again
                PBar.Value = PBar.Value + Len(strBlock) 'sets progressbar
            End If
        Case "Chat" 'if they are talking to ya
            txtChat.SelStart = Len(txtChat) 'sets cursor position in chatroom
            txtChat.SelText = strFriend & ":" & vbTab & strData2 & vbCrLf 'puts the chat into the room
    End Select
End Sub
