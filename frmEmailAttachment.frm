VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEmailAttachment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmbFile 
      Left            =   3000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstAttach 
      BackColor       =   &H00FFC0C0&
      Height          =   1035
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   3855
   End
   Begin VB.ListBox lstPath 
      Height          =   1035
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox txtFilename 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label cmdBrowse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
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
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin VB.Label cmdOK 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OK"
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
      TabIndex        =   8
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Attached Files :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filename :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label cmdReset 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reset"
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
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label cmdRemoveFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remove File"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image imgClose 
      Height          =   210
      Left            =   3960
      Picture         =   "frmEmailAttachment.frx":0000
      Top             =   0
      Width           =   225
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Attachments"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label cmdAddFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add File"
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
      Left            =   480
      TabIndex        =   0
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image imgMinimize 
      Height          =   210
      Left            =   3720
      Picture         =   "frmEmailAttachment.frx":0352
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   0
      Picture         =   "frmEmailAttachment.frx":0677
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmEmailAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumOfFilesAttached As Integer

Private Sub cmdAddFile_Click()

Dim strFileTitle As String
Dim strFilePath As String
Dim strTemp As String
Dim nPos As Integer
Dim NoMoreSlash As String


strFilePath = Trim$(txtFilename.Text) 'Remove spaces

'Check the availabality
If Dir(strFilePath) = "" Then
 csMsgbox "File can not be found.Please check the path", "File not found", "CSOKONLY"
 Exit Sub
End If

strTemp = strFilePath



'Remove Path
Do
nPos = InStr(1, strTemp, "\")

If nPos Then
 strTemp = Mid$(strTemp, nPos + 1)
Else
 Exit Do
End If

DoEvents
Loop

 
strFileTitle = strTemp


'Add entry to the list.
NumOfFilesAttached = NumOfFilesAttached + 1
lstPath.AddItem txtFilename.Text

lstAttach.AddItem strFileTitle & " (" & FileLen(strFilePath) & " bytes )"


End Sub

Private Sub cmdBrowse_Click()

'Show up the dialog
cmbFile.Filter = "All Files ( *.* )|*.*"
cmbFile.ShowOpen

'Check if user press cancel.

If cmbFile.CancelError Then
 Exit Sub
End If

'Assign the path to the textbox
txtFilename.Text = cmbFile.FileName


End Sub

Private Sub cmdOK_Click()
Dim I As Integer

'Check if there is any file to be attached,otherwise unload the form

If NumOfFilesAttached = 0 Then
 AttachedFiles = False
 Me.Hide
End If

'Add the files path to array.

'ReDim Preserve FileAttached(NumOfFilesAttached)

For I = 1 To NumOfFilesAttached
 FileAttached(I) = lstPath.List(I)
Next I

MsgBox FileAttached(1)
'Unload the form.

AttachedFiles = True


frmSMTP.SetFocus
Me.Hide

End Sub

Private Sub cmdRemoveFile_Click()

lstPath.RemoveItem lstAttach.ListIndex
lstAttach.RemoveItem lstAttach.ListIndex

NumOfFilesAttached = NumOfFilesAttached - 1

End Sub

Private Sub cmdReset_Click()

'Clear listboxes.
lstPath.Clear
lstAttach.Clear

NumOfFilesAttached = 0

End Sub

Private Sub Form_Load()

lstPath.AddItem "Empty"

End Sub
