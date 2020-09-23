VERSION 5.00
Begin VB.Form frmMsgbox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMsgBox.frx":0000
   ScaleHeight     =   1905
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label cmdYes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Yes"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label cmdNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
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
      Left            =   3480
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgClose 
      Height          =   210
      Left            =   6000
      Picture         =   "frmMsgBox.frx":02E2
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblTopic 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "csD Information"
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
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   5655
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   -1200
      Picture         =   "frmMsgBox.frx":05C4
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   7455
   End
End
Attribute VB_Name = "frmMsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
 csMRet = "CSCANCEL"
 Unload Me
End Sub

Private Sub cmdNo_Click()
 csMRet = "CSNO"
 Unload Me
End Sub

Private Sub cmdOK_Click()
 csMRet = "CSOK"
 
 Unload Me
End Sub

Private Sub cmdYes_Click()
 csMRet = "CSYES"
 Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 DragForm Me
End If
End Sub


Private Sub imgClose_Click()
 Unload Me
End Sub

Private Sub lblMsg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 DragForm Me
End If
End Sub
