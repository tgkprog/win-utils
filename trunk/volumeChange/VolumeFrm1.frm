VERSION 5.00
Begin VB.Form Frm1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Volume Control With Timer - sel2in.com"
   ClientHeight    =   1515
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1515
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   240
   End
   Begin VB.TextBox txtVolNow 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Text            =   "0"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtVolLater 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Text            =   "80"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdRrefresh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Refresh"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSet 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Set"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtWaitSec 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Text            =   "30"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "volume to set after wait"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "wait seconds"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "volume set to now 0-100"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuMnd 
      Caption         =   "Actions"
      Begin VB.Menu mnuRefh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Set"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuHelpTopDmy 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "Website"
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "Support and feedback Email "
      End
      Begin VB.Menu mnuFeedbackGet 
         Caption         =   "One click feedback"
      End
   End
End
Attribute VB_Name = "Frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()

End Sub

Private Sub cmdRrefresh_Click()
Dim aa
MsgBox "Not yet implemented"
'waveOutGetVolume 1, aa
'cmdRrefresh.Caption = "Refresh Vol now is " & GetVolume(1)
End Sub

Public Sub cmdSet_Click()
' VolumeLevel is the level value in percentage (0 = min, 100 = max)
' Returns True if successful
'Function SetVolume(VolumeLevel As Long) As Boolean
On Local Error GoTo errh
SetVolume CInt(Val(txtVolNow))
Dim ii As Long
ii = Val(txtWaitSec) * 1000
If ii < 1000 Then ii = 1001
Timer1.Interval = ii
cmdSet.Caption = "Seq   t"
Exit Sub
errh:
MsgBox "Could not set " & Err.Number & " " & Err.Description, vbExclamation
End Sub

Private Sub mnuAbout_Click()
MsgBox "Copyright 2009 Tushar Kapila http://sel2in.com "
End Sub

Private Sub mnuEmail_Click()
Dim iFileNo As Integer
iFileNo = FreeFile
Open App.Path & "\sel2in_feedback.bat" For Output As #iFileNo
Print #iFileNo, "start mailto:tgkprog@gmail.com?subject=volume-control&v=" & getVer & ""
Close #iFileNo
Shell App.Path & "\sel2in-feedback.bat"
Sleep 2000
Open App.Path & "\sel2in_feedback.bat" For Output As #iFileNo
Close #iFileNo
End Sub

Private Sub mnuFeedbackGet_Click()
On Local Error GoTo errh
Dim ss, iFileNo
ss = InputBox("One line feedback on product :", "Contents will be posted to our website using your browser, no sign in")
iFileNo = FreeFile
Open App.Path & "\sel2in-web-site-go.bat" For Output As #iFileNo
ss = Replace(ss, " ", "+")
ss = Replace(ss, Chr(10), "{")
ss = Replace(ss, Chr(13), "|}")
ss = Replace(ss, vbTab, "--")
Print #iFileNo, "start http://sel2in.com/prjs/php/own/feedbackGet.php?c=volume-control+l=vb6&+=" & getVer & "+c=" & ss
Close #iFileNo
Shell App.Path & "\sel2in-web-site-go.bat"
Exit Sub
errh:
End Sub

Private Sub mnuHelp_Click()
MsgBox "Allows you to set the volume and set the volume to a new value after the seconds you enter. Nice to skip ads or put speakers on mute for a while/ change volume after a while." _
 & vbNewLine & "Set the volume you want the system to be set to currently on first text box, seconds you want it wait in the second text box and finally the new volume to be set (from 0 to 100 maximum) in the 3rd. Then press the set button" _
 & vbNewLine & "Command line : param 1 : volume to set now; param 2: seconds to wait; param 3: volume after wait" _
 & vbNewLine & "See batch file ""RunVolumeChangeSilent.bat""for a sample" _
  , vbInformation, "Change volume with timer"
  
End Sub

Private Sub mnuRefh_Click()
cmdRrefresh_Click
End Sub

Private Sub mnuSet_Click()
cmdSet_Click
End Sub

Private Sub mnuWebsite_Click()

Dim iFileNo As Integer
iFileNo = FreeFile
Open App.Path & "\sel2in-web-site-go.bat" For Output As #iFileNo
Print #iFileNo, "start http://sel2in.com?app=volume-control&l=vb6&v=" & getVer & ""
Close #iFileNo
Shell App.Path & "\sel2in-web-site-go.bat"

End Sub

Private Sub Timer1_Timer()
SetVolume CInt(Val(txtVolLater))
Timer1.Interval = 0
cmdSet.Caption = "Set, done"
If bUnat Then End
End Sub

Public Function getVer() As String
getVer = App.Major & "." & App.Minor & "." & App.Revision
End Function
