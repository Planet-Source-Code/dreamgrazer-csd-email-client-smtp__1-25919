VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSMTP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSMTP.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtToAdd 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   22
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox txtFromAdd 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   21
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CheckBox chkPOPAuth 
      BackColor       =   &H00800000&
      Caption         =   "Check1"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1680
      Top             =   5400
   End
   Begin VB.CheckBox chkVerify 
      BackColor       =   &H00800000&
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   5400
      Width           =   255
   End
   Begin MSWinsockLib.Winsock sckSMTP 
      Left            =   1560
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMsg 
      BackColor       =   &H00FFC0C0&
      Height          =   2805
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2520
      Width           =   5535
   End
   Begin VB.TextBox txtBCC 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtCC 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtSubject 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox txtTo 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label cmdAttachment 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Attachments"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Recepient name:"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sender name"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "POP Auth"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label cmdSetup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Address"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Message :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CC :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BCC :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label cmdClear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label cmdCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image imgClose 
      Height          =   210
      Left            =   5520
      Picture         =   "frmSMTP.frx":02E2
      Top             =   0
      Width           =   225
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "csD EmailClient"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label cmdSend 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   5520
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   -1200
      Picture         =   "frmSMTP.frx":05C4
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   6975
   End
End
Attribute VB_Name = "frmSMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevCommand As String
Dim IsSMTPConnected As Boolean
Dim SMTPAdd As String
Dim ToAdd As String
Dim CCAdd As String
Dim BCCAdd As String
Dim FromAdd As String
Dim Subject As String
Dim Msg As String
Dim SenderDomain As String
Dim intWait As Integer
Dim IsVerified As Boolean
Dim IsAvailable As Boolean
Dim IsShouldVerify As Boolean
Dim IsCanceled As Boolean
Dim POPAuthorised As Boolean
Dim POPError As Boolean
Dim SignedOut As Boolean



Private Sub chkVerify_Click()

 If chkVerify.Value = 1 Then
  IsShouldVerify = True
 ElseIf chkVerify.Value = 0 Then
  IsShouldVerify = False
 End If
 
End Sub

Private Sub cmdAttachment_Click()
 frmEmailAttachment.Show 0
End Sub

Private Sub cmdCancel_Click()
 IsCanceled = True
End Sub

Private Sub cmdReset_Click()
 'Reset all form fields
txtTo.Text = ""
txtSubject.Text = ""
txtCC.Text = ""
txtBCC.Text = ""


End Sub



Private Sub cmdClear_Click()
 'Clear all textbox
 txtTo.Text = ""
 txtSubject.Text = ""
 txtMsg.Text = ""
 txtCC.Text = ""
 txtBCC.Text = ""
 txtToAdd.Text = ""
 txtFromAdd.Text = ""
 
End Sub

Private Sub cmdSend_Click()

'Disable the button, so the user will not press it during email transfer
 cmdSend.Enabled = False

ToAdd = txtTo.Text



'Validate form

If EmailUsername = "" Or EmailPass = "" Or SMTPAddress = "" Or POPAddress = "" Then
 csMsgbox "User Information not complete", "User Information", "CSOKONLY"
 cmdSend.Enabled = True
 Exit Sub
End If


If Trim$(txtTo.Text) = "" Or Trim$(txtSubject.Text) = "" Then
 csMsgbox "Form fields are not completed.", "Fields not completed", "CSOKONLY"
 cmdSend.Enabled = True
 Exit Sub
End If



If chkPOPAuth.Value = 1 Then
'If need POP authorisation,

 sckSMTP.Connect POPAddress, 110
 PrevCommand = "POPConnect"
 
 'Wait for authorisation
 Do Until POPAuthorised = True Or POPError = True Or IsCanceled = True
  DoEvents
 Loop
 
 If POPError Then
  'If error occured in authorisation.Inform and exit
  csMsgbox "Error in authorisation", "Error", "CSOKONLY"
  cmdSend.Enabled = True
  Exit Sub
 End If
 
 'CLose connection
 sckSMTP.Close
 
End If
'Connect to the SMTP server

sckSMTP.Connect SMTPAddress, 25

'Start the timer
tmrTimeout.Enabled = True

End Sub

Private Sub cmdSetup_Click()
 frmEmailSetup.Show 0
End Sub

Private Sub Form_Load()
 
 IsShouldVerify = False
 IsCanceled = False
 
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If Button = 1 Then
  DragForm Me
 End If

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 'CLose if needed
 If sckSMTP.State <> sckClosed Then sckSMTP.Close
 
 'If we sign in the POP earlier
  If chkPOPAuth.Value = 1 And POPAuthorised = True Then
  
   sckSMTP.Connect POPAddress, 110
   PrevCommand = "SignOut"
   
   Do Until SignedOut = True Or POPError = True
    DoEvents
   Loop
   'CLose it
   sckSMTP.Close
  
  End If
  
 
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'Add form dragging capability.
 If Button = 1 Then
  DragForm Me
 End If

End Sub

Private Sub imgClose_Click()
 
 'Put the clearance code here ///////
 Unload Me
 
End Sub

Private Sub sckSMTP_Close()
'CLose it
If sckSMTP.State <> sckClosed Then sckSMTP.Close

End Sub

Private Sub sckSMTP_Connect()
 lblStatus.Caption = "Connected to server"
End Sub

Private Sub sckSMTP_DataArrival(ByVal bytesTotal As Long)

Dim DatRec As String
Dim strBuffer As String
Dim Filenum As Integer
Dim I As Integer

Dim TempPath As String

sckSMTP.GetData DatRec

'For intercepting SMTP data

Select Case Val(Left$(DatRec, 3))

Case 220

 If Not IsSMTPConnected Then
  IsSMTPConnected = True
  tmrTimeout.Enabled = False
  'Notify address
  smtpSendData "HELO <" & SMTPDomain & ">" & vbCrLf 'smtp.
  
 
 End If

Case 250

 Select Case PrevCommand
 
 Case "HELO"
 
 
 '--------Start of Verify Codes--------
 
 'Verify the e-mail address.( If IsShouldVerify flag turned on )

  If IsShouldVerify Then
  
   smtpSendData "VRFY " & Left$(ToAdd, InStr(ToAdd, InStr(ToAdd, "@") - 1))
   
   lblStatus.Caption = "Verifying Recepient Address"
 
   Do Until intWait >= 10 Or IsVerified = True   'Wait for verification or timeout
     
     If IsCanceled Then  'User Canceled
     
      IsCanceled = False
      lblStatus.Caption = "Action canceled."
      cmdSend.Enabled = True
      
      Exit Sub
     End If
     
     intWait = intWait + 1
     
     DoEvents
   Loop
 
   If IsAvailable Then
   
    lblStatus.Caption = "E-mail address verified.Starting communication with server.."
    
   ElseIf Not IsAvailable Then
    'Inform user
    csMsgbox "The recepient not available or server busy.Please try again", "Recepient Address Verification Failure", "CSYES"
    'Enable the Send button.
    cmdSend.Enabled = True
    lblStatus.Caption = "Ready"
    Exit Sub
   End If
  End If
  
  '-----End of Verify Code-------
  
  'Send the sender address..
  smtpSendData "MAIL FROM: <" & EmailUsername & "@" & SMTPDomain & ">" & vbCrLf ' & EmailUsername & Mid$(SMTPAddress, 6) & ">" & vbCrLf
  lblStatus.Caption = "Sending Mailer Address"
  
 Case "MAIL"
  smtpSendData "RCPT TO: <" & Trim$(ToAdd) & ">" & vbCrLf
  lblStatus.Caption = "Sending Recepient Address"
  
 Case "RCPT"
  lblStatus.Caption = "Preparing to send the message"
  smtpSendData "DATA" & vbCrLf
  
 Case "VRFY"
  
  IsVerified = True
  
 Case "DATE"
  'Email sent.
  
  lblStatus.Caption = "E-mail sent successfully"
   
  'CLose the connection
  sckSMTP.Close
  IsSMTPConnected = False
  'Sign Out from Server
  If chkPOPAuth.Value = 1 Then
  
   sckSMTP.Connect POPAddress, 110
   PrevCommand = "SignOut"
   
   Do Until SignedOut = True Or POPError = True
    DoEvents
   Loop
   
   If POPError Then
    lblStatus.Caption = "Failed to signed out"
   Else
   lblStatus.Caption = "Signed Out Successfully"
   End If
   
   'CLose it
   sckSMTP.Close
  
  End If
  
  cmdSend.Enabled = True
  
 
 End Select
 
Case 251
 
 Select Case PrevCommand
 
 Case "RCPT"
  smtpSendData "DATA" & vbCrLf
  
  lblStatus.Caption = "Email forwarded to :" & K
 
 End Select
 
Case 354

 Select Case PrevCommand
 
  Case "DATA"
 
  lblStatus.Caption = "Sending Message Body"
   'Server ready for message.Compose the message.
   Msg = "DATE: " & Format(Now, "dd mmm yy ttttt") & vbCrLf & "FROM: " & txtFromAdd.Text & vbCrLf & "TO: " & txtToAdd.Text & vbCrLf & "SUBJECT: " & txtSubject.Text & vbCrLf & vbCrLf & txtMsg.Text
  
'====================Email Attachments=================

'Attach the UUencoded files...If there are any..
If AttachedFiles Then

AttachedFiles = False
TempPath = App.Path
If Right$(TempPath, 1) <> "\" Then TempPath = TempPath & "\"

'Kill any stupid files out there...
If Dir(TempPath & "Temp.dat") <> "" Then Kill TempPath & "Temp.dat"

Filenum = FreeFile
For I = 1 To UBound(FileAttached)
 UUEncode FileAttached(I), TempPath & "Temp.dat", True
Next

Open TempPath & "Temp.dat" For Append As #Filenum
 Print #Filenum, "."
Close #Filenum

'Send the header first
 smtpSendData Msg

lblStatus.Caption = "Sending attached files"

 
 Open TempPath & "Temp.dat" For Binary As Filenum

 Do Until EOF(Filenum)
  strBuffer = Space$(8192)
  Get #Filenum, , strBuffer
  
  sckSMTP.SendData strBuffer
 Loop
 
 
 
 'Mark the end of message
 sckSMTP.SendData vbCrLf & "." & vbCrLf
 
 Close Filenum
 
 lblStatus.Caption = "Attachments transfer complete"
 
  
'======================================================

Else
 Msg = Msg & vbCrLf & "."
 smtpSendData Msg
End If


   
  
 End Select
 
 
'Errors
Case Is >= 400
 
 MsgBox Mid$(DatRec, 4), vbInformation, "Error in Email Transaction"
  
 'Reenable Send Button
 cmdSend.Enabled = True
 lblStatus.Caption = "Ready"
 'Close socket
 sckSMTP.Close
 
 If POPAuthorised Then POPAuthorised = False
 IsSMTPConnected = False
 
End Select






'Intercepting POP messages
Select Case Left$(DatRec, 3)

Case "+OK"
 Select Case PrevCommand
  Case "POPConnect"
   sckSMTP.SendData "USER " & EmailUsername & vbCrLf
   PrevCommand = "Username"
   lblStatus.Caption = "Sending Username"
  Case "Username"
   sckSMTP.SendData "PASS " & EmailPass & vbCrLf
   PrevCommand = "Pass"
   lblStatus.Caption = "Sending Password"
  Case "Pass"
   POPAuthorised = True
   lblStatus.Caption = "Authorised"
  Case "SignOut"
   sckSMTP.SendData "QUIT"
   POPAuthorised = False
   SignedOut = True
 End Select
Case "-ERR"
 'Error from POP server
 csMsgbox "Error " & Mid$(DatRec, 5) & " at POP Server", "Error from server", "CSOKONLY"

End Select

End Sub

Private Sub smtpSendData(strMessage As String)

PrevCommand = Left$(Trim$(strMessage), 4)

 sckSMTP.SendData strMessage



End Sub




Private Sub sckSMTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'AN error occured.Display the error
 csMsgbox "An error occured [" & Number & "] : " & Description, "Error", "VBYES"
 'Close the connection
 sckSMTP.Close
 'Mark unconnected
 IsSMTPConnected = False
 'If User pressed Send button...reenable it
 If Not cmdSend.Enabled Then cmdSend.Enabled = True
 
End Sub

Private Sub tmrTimeout_Timer()

'Timeourt occured.Display the error
lblStatus.Caption = "Timeout Error.Try Again"
cmdSend.Enabled = True
'close the connection
sckSMTP.Close
IsSMTPConnected = False

'disable th etimer
tmrTimeout.Enabled = False
End Sub
