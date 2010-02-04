Attribute VB_Name = "Module1"
Option Explicit
Public fso As FileSystemObject
Public appAuto As Boolean
Public Const APP_CAP = "Start up later - sel2in.com"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub main()
appAuto = False

Set fso = New FileSystemObject

Load StartUpLtrF


If Command$ = "" Or Left(Command$, 3) = "--s" Or Left(Command$, 2) = "a " Then
    StartUpLtrF.Show
ElseIf Command$ = "-h" Or Command$ = "-help" Then
    StartUpLtrF.Show
    StartUpLtrF.mnuHelp_Click
    
ElseIf Command$ = "-r" Or Command$ = "/run" Then
    StartUpLtrF.Show
    StartUpLtrF.mnuRun_Click
End If
End Sub
