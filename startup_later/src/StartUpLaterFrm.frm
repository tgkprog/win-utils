VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Start programs a few minutes after Windows starts"
   ClientHeight    =   5355
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14085
   Icon            =   "StartUpLaterFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Remove selected"
      Height          =   735
      Left            =   12840
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   735
      Left            =   12840
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtParams 
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   840
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   3
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
      Left            =   12120
      TabIndex        =   8
      Text            =   "0"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   10080
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
      Text            =   "200"
      Top             =   120
      Width           =   855
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
      Caption         =   "Wait Seconds"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuDmyFile 
      Caption         =   "File"
      Begin VB.Menu mnuChoseFile 
         Caption         =   "Chose File"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'Private Sub cmdBrowse_Click() 'On Local Error GoTo errH 'Err.Clear ' If Err.Number <> 0 Then 'errH:'    MsgBox "Er " & Err.Number & " " & Err.Description & vbNewLine & errLoc 'End If 'End Sub


''''''''''''''''''
Dim errLoc As String
Private Sub fn1()
On Local Error GoTo errh


Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Er " & Err.Number & " " & Err.Description & vbNewLine & errLoc
End If
errhIg:
End Sub
''''''''''''''''''

Private Sub cmdBrowse_Click()
On Local Error GoTo errh
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
txtCmd = CommonDialog1.FileName
On Local Error GoTo errhIg
Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Er " & Err.Number & " " & Err.Description & vbNewLine & errLoc
End If
errhIg:
End Sub

Private Sub Form_Load()
Me.CommonDialog1.InitDir = App.Path
Debug.Print ""
End Sub

Private Sub mnuChoseFile_Click()
On Local Error GoTo errh
Me.CommonDialog1.CancelError = True
Me.CommonDialog1.ShowOpen

errh:
End Sub
