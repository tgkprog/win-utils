VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label l3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

      Dim hmixer As Long          ' mixer handle
      Dim volCtrl As MIXERCONTROL ' waveout volume control
      Dim micCtrl As MIXERCONTROL ' microphone volume control
      Dim rc As Long              ' return code
      Dim ok As Boolean           ' boolean return code
      Dim vol As Long             ' volume

      Private Sub Form_Load()
      ' Open the mixer with deviceID 0.
         rc = mixerOpen(hmixer, 0, 0, 0, 0)
         If ((MMSYSERR_NOERROR <> rc)) Then
             MsgBox "Couldn't open the mixer."
             Exit Sub
             End If

         ' Get the waveout volume control
         ok = GetVolumeControl(hmixer, _
                              MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                              MIXERCONTROL_CONTROLTYPE_VOLUME, _
                              volCtrl)
         If (ok = True) Then
             ' If the function successfully gets the volume control,
             ' the maximum and minimum values are specified by
             ' lMaximum and lMinimum
             Label1.Caption = volCtrl.lMinimum _
                              & " to " _
                              & volCtrl.lMaximum
                              l3 = volCtrl.cMultipleItems
             End If

         ' Get the microphone volume control
         ok = GetVolumeControl(hmixer, _
                              MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, _
                              MIXERCONTROL_CONTROLTYPE_VOLUME, _
                              micCtrl)
         If (ok = True) Then
             Label2.Caption = micCtrl.lMinimum _
                              & " to " _
                              & micCtrl.lMaximum
             End If
      End Sub

      Private Sub Command1_Click()
         vol = CLng(Text1.Text)
         SetVolumeControl hmixer, volCtrl, vol
      End Sub

      Private Sub Command2_Click()
         vol = CLng(Text2.Text)
         SetVolumeControl hmixer, micCtrl, vol
      End Sub

