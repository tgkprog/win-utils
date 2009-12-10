VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Settings and help for file list names"
   ClientHeight    =   5130
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   4920
      TabIndex        =   10
      Top             =   2160
      Width           =   3495
      Begin VB.OptionButton optFileNameOnly 
         BackColor       =   &H00FF8080&
         Caption         =   "File name only"
         Height          =   540
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton optFullPath 
         BackColor       =   &H00FF8080&
         Caption         =   "Full file path and name"
         Height          =   540
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame frmOptFileClip 
      BackColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   3855
      Begin VB.TextBox txtFileName 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Text            =   "dir.txt"
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optClip 
         BackColor       =   &H00FF8080&
         Caption         =   "Copy to clipboard"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   3375
      End
      Begin VB.OptionButton optFile 
         BackColor       =   &H00FF8080&
         Caption         =   "Make a file called"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdCopySendTo 
      Caption         =   "Copy overwrite to send to folder"
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdSavClos 
      Caption         =   "S&ave && close"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "&Reload settings"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
   Begin VB.Menu mnuDumyFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ahange()
''Usage: pass files as arguments, whose full path will be written to a file called dir.txt in the path of the first file
'Place this exe in your send to folder for easy usage.
'Your send to folder is "C:\Documents and Settings\tushar\SendTo\ "
'*** Press Yes if you want me to copy my self to that folder.
'- http://code.google.com/p/win-utils/

End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdCopySendTo_Click()
On Local Error Resume Next
Dim sy, syt, a, s, i


a = Environ("USERPROFILE") & "\SendTo\" & App.EXEName & ".exe"
syt = vbOKOnly

s = "Place this exe in your send to folder for easy usage. " _
& vbNewLine & "Your send to folder is """ & Environ("USERPROFILE") & "\SendTo\ """ _
& sy _
& vbNewLine & "- http://code.google.com/p/win-utils/" & vbNewLine & "Tushar Kapila http://sel2in.com Copyright 2009"

If (a = App.Path & "\" & App.EXEName & ".exe") Then
    MsgBox "Already there " & a, vbInformation
Else
    Call fso.CopyFile(App.Path & "\" & App.EXEName & ".exe", a, True)

End If
End Sub

Private Sub cmdReload_Click()
On Local Error Resume Next
optClip.Value = GetSetting("sel2in", App.EXEName, "action", "files") = "clip"
optFile.Value = GetSetting("sel2in", App.EXEName, "action", "files") = "files"
txtFileName = GetSetting("sel2in", App.EXEName, "filename", txtFileName)
optFileNameOnly.Value = GetSetting("sel2in", App.EXEName, "fullFile", IIf(optFileNameOnly.Value, "no", "yes")) = "no"
optFullPath.Value = GetSetting("sel2in", App.EXEName, "fullFile", IIf(optFullPath.Value, "yes", "no")) = "yes"
End Sub

Private Sub cmdSavClos_Click()
On Local Error Resume Next
cmdSave_Click
End
End Sub

Private Sub cmdSave_Click()
On Local Error Resume Next
Call SaveSetting("sel2in", App.EXEName, "action", IIf(optClip.Value, "clip", "files"))
Call SaveSetting("sel2in", App.EXEName, "filename", txtFileName)
Call SaveSetting("sel2in", App.EXEName, "fullFile", IIf(optFileNameOnly.Value, "no", "yes"))
End Sub

Private Sub Form_Load()
On Local Error Resume Next
optFullPath.Value = True
optFile.Value = True
Text1.Text = Replace(Text1.Text, "C:\Documents and Settings\tushar\SendTo\", Environ("USERPROFILE") & "\SendTo\")
Text1.Text = Replace(Text1.Text, "{{[SendTo]}}", Environ("USERPROFILE") & "\SendTo\")
Text1.Text = Text1.Text & vbNewLine & "Version " & App.Major & "." & App.Minor & "." & App.Revision
cmdReload_Click
End Sub

Private Sub mnuQuit_Click()
End
End Sub



Private Sub mnuReload_Click()
cmdReload_Click
End Sub


Private Sub mnuSave_Click()
cmdSave_Click
End Sub

