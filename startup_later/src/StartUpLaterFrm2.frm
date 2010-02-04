VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form StartUpLtrF 
   Caption         =   "Start programs a few minutes after Windows starts"
   ClientHeight    =   5355
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14085
   Icon            =   "StartUpLaterFrm2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel2 
      Caption         =   "Run"
      Height          =   495
      Left            =   7680
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove selected"
      Height          =   855
      Left            =   12840
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   855
      Left            =   12840
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtParams 
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   840
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   3
      AllowUserResizing=   3
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtWaitSub 
      Height          =   375
      Left            =   10320
      TabIndex        =   8
      Text            =   "0"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   11160
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   ".."
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtCmd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox txtWait1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Text            =   "30"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Time  after each program milli sec"
      Height          =   735
      Left            =   10320
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbls 
      Caption         =   "Finally press Add"
      Height          =   240
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lbls 
      Caption         =   "Optionally: * add parameters * add additional wait time before starting this program"
      Height          =   240
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Choose a program or paste its path. "
      Height          =   240
      Index           =   0
      Left            =   2520
      TabIndex        =   5
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Initial Wait Seconds"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuMnd 
      Caption         =   "Actions"
      Begin VB.Menu mnuChoseFile 
         Caption         =   "Choose File"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuRun 
         Caption         =   "Run"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuCancelRun 
         Caption         =   "Cancel run"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save settings"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuShortCutsUpDmy 
      Caption         =   "Options"
      Begin VB.Menu mnuCloseAfterRun 
         Caption         =   "Close after run"
         Checked         =   -1  'True
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuShortcutProg 
         Caption         =   "Shortcut with global hotkey"
      End
      Begin VB.Menu mnuShortAppProg 
         Caption         =   "App shortcut in Programs\apps"
      End
      Begin VB.Menu mnuShortAppDesk 
         Caption         =   "App short cut on desktop"
      End
   End
   Begin VB.Menu mnuHelpTopDmy 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
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
      Begin VB.Menu mnuHelpCopyClip 
         Caption         =   "Copy help to clipboard"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "StartUpLtrF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFile
Dim iState
Dim iTmr


'Private Sub cmdBrowse_Click() 'On Local Error GoTo errH 'Err.Clear ' If Err.Number <> 0 Then 'errH:'    MsgBox "Er " & Err.Number & " " & Err.Description & vbNewLine & errLoc 'End If 'End Sub


''''''''''''''''''
Dim errLoc As String
Private Sub fn1()
On Local Error GoTo errH


Err.Clear
If Err.Number <> 0 Then
errH:
    MsgBox "Er " & Err.Number & " " & Err.Description & vbNewLine & errLoc
End If
errhIg:
End Sub
''''''''''''''''''

Sub refreshGrid()
On Local Error GoTo errH
grd.Rows = 1
grd.Rows = 2

Dim tx As TextStream
Dim s As String
Set tx = fso.OpenTextFile(sFile, ForReading)
Do While Not tx.AtEndOfStream
    s = Trim(tx.ReadLine)
    If Not s = "" Then
        grd.Rows = grd.Rows + 1
        grd.TextMatrix(grd.Rows - 1, 1) = s
    End If
Loop
tx.Close
grd.TextMatrix(0, 0) = "File"
grd.TextMatrix(0, 1) = sFile
Err.Clear
If Err.Number <> 0 Then
errH:
    MsgBox "Er " & Err.Number & " " & Err.Description & vbNewLine & errLoc
End If
errhIg:

End Sub

Private Sub cmdAdd_Click()
MsgBox "Not implemented, can only specify initial wait time, per program wait time and a text file with programs to launch - one program with arguments per line"
End Sub

Private Sub cmdBrowse_Click()
On Local Error GoTo errH
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
txtCmd = CommonDialog1.FileName
On Local Error GoTo errhIg
Err.Clear
If Err.Number <> 0 Then
errH:
    MsgBox "Er " & Err.Number & " " & Err.Description & vbNewLine & errLoc
End If
errhIg:
End Sub

Private Sub cmdCancel2_Click()
If iState < 0 Then
    mnuRun_Click
Else
    mnuCancelRun_Click
End If
End Sub

Private Sub Command1_Click()
cmdAdd_Click
End Sub

Private Sub Command2_Click()
cmdAdd_Click
End Sub

Private Sub Form_Load()
Me.CommonDialog1.InitDir = App.Path
Debug.Print ""
loadSet
grd.ColWidth(1) = 3720
doStop
refreshGrid

End Sub

Private Sub mnuCancelRun_Click()
doStop
End Sub

Private Sub mnuChoseFile_Click()
On Local Error Resume Next
If sFile <> "" Then
    CommonDialog1.InitDir = fso.GetParentFolderName(sFile)
End If
On Local Error GoTo errH
Me.CommonDialog1.CancelError = True
Me.CommonDialog1.ShowOpen
sFile = Me.CommonDialog1.FileName
svSet
refreshGrid
errH:
End Sub

Sub svSet()
SaveSetting App.Title, "prefs", "f", sFile
SaveSetting App.Title, "prefs", "t1", txtWait1
SaveSetting App.Title, "prefs", "t2", txtWaitSub
SaveSetting App.Title, "prefs", "mnuCloseAfterRun", mnuCloseAfterRun.Checked
End Sub

Sub loadSet()
sFile = GetSetting(App.Title, "prefs", "f", sFile)
txtWait1 = GetSetting(App.Title, "prefs", "t1", txtWait1)
txtWaitSub = GetSetting(App.Title, "prefs", "t2", txtWaitSub)
mnuCloseAfterRun.Checked = GetSetting(App.Title, "prefs", "mnuCloseAfterRun", mnuCloseAfterRun.Checked)
End Sub




Private Sub mnuAbout_Click()
MsgBox getAboutText, vbInformation, APP_CAP

End Sub

Private Sub mnuCloseAfterRun_Click()
mnuCloseAfterRun.Checked = Not mnuCloseAfterRun.Checked
End Sub

Private Sub mnuEmail_Click()
Dim iFileNo As Integer
iFileNo = FreeFile
Open App.Path & "\sel2in_feedback.bat" For Output As #iFileNo
Print #iFileNo, "start mailto:tgkprog@gmail.com?subject=start-up-later&v=" & getVer & ""
Close #iFileNo
Shell App.Path & "\sel2in-feedback.bat"
Sleep 2000
Open App.Path & "\sel2in_feedback.bat" For Output As #iFileNo
Close #iFileNo
End Sub

Private Sub mnuFeedbackGet_Click()
On Local Error GoTo errH
Dim ss, iFileNo
ss = InputBox("One line feedback on product :", "Contents will be posted to our website using your browser, no sign in")
iFileNo = FreeFile
Open App.Path & "\sel2in-web-site-go.bat" For Output As #iFileNo
ss = Replace(ss, " ", "+")
ss = Replace(ss, Chr(10), "{")
ss = Replace(ss, Chr(13), "|}")
ss = Replace(ss, vbTab, "--")
Print #iFileNo, "start http://sel2in.com/prjs/php/own/feedbackGet.php?c=start-up-later+l=vb6+" & getVer & "+c=" & ss
Close #iFileNo
Shell App.Path & "\sel2in-web-site-go.bat"
Exit Sub
errH:
End Sub

Public Sub mnuHelp_Click()
MsgBox getHelpText, vbInformation, APP_CAP
End Sub

Private Sub mnuHelpCopyClip_Click()
Clipboard.Clear
Clipboard.SetText getHelpText
End Sub





Public Sub mnuRun_Click()
On Local Error GoTo errH
mnuRun.Enabled = False
mnuCancelRun.Enabled = True
Timer1.Enabled = True
iState = 0
If Val(txtWait1.Text) < 0 Then txtWait1.Text = 0
If Val(txtWaitSub.Text) < 1 Then txtWaitSub.Text = 1
Timer1.Interval = 1000
iTmr = IIf((txtWait1.Text) > 0, txtWait1, 1)
cmdCancel2.Caption = "Cancel (" & iTmr & " sec wait)"
grd.TextMatrix(1, 1) = "wait " & txtWait1
grd.TextMatrix(1, 0) = "  >"
errH:
End Sub

Private Sub mnuSave_Click()
svSet
End Sub

Private Sub mnuShortAppDesk_Click()
Dim ww As WshShell
Dim sht As WshShortcut

Set ww = New WshShell

Dim s
s = ww.SpecialFolders("Desktop")
mnuShortcutMake s
End Sub

Private Sub mnuShortAppProg_Click()
Dim ww As WshShell
Dim sht As WshShortcut

Set ww = New WshShell

Dim s
s = ww.SpecialFolders("programs")
mnuShortcutMake s & "\apps"
End Sub
Private Sub mnuShortcutMake(s)

On Local Error Resume Next

Dim fso As FileSystemObject

Dim tx As TextStream
Dim ww As WshShell
Dim sht As WshShortcut

Set ww = New WshShell

Set fso = New FileSystemObject
If Not fso.FolderExists(s) Then
    Call fso.CreateFolder(s)
End If
Set sht = ww.CreateShortcut(s & "\Auto Change Volume with Timer.lnk")
'App.Path & "\" & App.EXEName & ".exe"
'sht.= "Vol Auto " & Left(txtHotKey, 1) & " " & Me.txtVolNow & " " & Me.txtWaitSec & " " & Me.txtVolLater
sht.TargetPath = App.Path & "\" & App.EXEName & ".exe"

'sht.Arguments = Me.txtVolNow & " " & Me.txtWaitSec & " " & Me.txtVolLater & " "
sht.Description = "Change volume automatically like to level 90 of 100  after 30 seconds. initially volume can be set to 0 (or -1 means not changed) - Tushar Kapila http://sel2in.com/"
'sht.Hotkey = "ALT+SHIFT+" & Left(txtHotKey, 1)
sht.Save

End Sub

Private Sub mnuShortcutProg_Click()
Dim i, s, j
i = MsgBox("This will place a short cut in windows Start/ programs with a global shortcut you specify to start this program without UI with current settings - good way to stop volume when a internet radio station goes to ads!", vbInformation Or vbYesNo, "Ready for short cut old will be over written unless you change name")
If i = vbYes Then
    Me.Width = 6165
    Me.Height = 3360
Else
    Me.Width = 5595
    Me.Height = 2325
End If
End Sub

Private Sub mnuWebsite_Click()

Dim iFileNo As Integer
iFileNo = FreeFile
Open App.Path & "\sel2in-web-site-go.bat" For Output As #iFileNo
Print #iFileNo, "start http://sel2in.com?app=start-up-laterl&l=vb6&v=" & getVer & ""
Close #iFileNo
Shell App.Path & "\sel2in-web-site-go.bat"

End Sub



Public Function getVer() As String
getVer = App.Major & "_" & App.Minor & "_" & App.Revision
End Function

Public Function getHelpText() As String
getHelpText = "Start up programs after windows starts - with a delay in seconds." _
 & vbNewLine & "" _
 & vbNewLine & "can only specify initial wait time, per program wait time and a text file with programs to launch - one program with arguments per line" _
 & vbNewLine & getAboutText
End Function

Public Function getAboutText() As String
getAboutText = "Startup Later. Copyright 2009 Tushar Kapila http://sel2in.com version " & App.Major & "." & App.Minor & "." & App.Revision
End Function

Private Sub Timer1_Timer()

On Local Error GoTo errH



'row 0 and 1 not used, row 2 till rows-1
If grd.Rows < 2 Or iState >= (grd.Rows - 1) Or iState < 0 Then

    doStop
    Exit Sub
End If
If iState = 0 And iTmr > 0 Then
    iTmr = iTmr - 1
    cmdCancel2.Caption = "Cancel (" & iTmr & " sec wait)"
    Exit Sub
ElseIf iState = 0 Then
    iState = 2
    grd.TextMatrix(iState - 1, 0) = ""
    grd.TextMatrix(iState - 1, 1) = ""
End If
If Val(txtWaitSub) < 1 Then txtWaitSub = 1
Timer1.Interval = txtWaitSub
cmdCancel2.Caption = "Cancel (on item :" & (iState - 1) & ")"
grd.TextMatrix(iState, 0) = "  > "
On Local Error GoTo dNxt
Shell Trim(grd.TextMatrix(iState, 1))

On Local Error GoTo errH
iState = iState + 1
Err.Clear
If Err.Number <> 0 Then
errH:
    MsgBox "Er " & Err.Number & " " & Err.Description & vbNewLine & errLoc
End If
errhIg:

Exit Sub
dNxt:
Resume Next
End Sub


Sub doStop()
grd.TextMatrix(1, 1) = ""
iState = -1
Timer1.Interval = 0
Timer1.Enabled = False
mnuRun.Enabled = True
mnuCancelRun.Enabled = False
cmdCancel2.Caption = "Run"
End Sub
