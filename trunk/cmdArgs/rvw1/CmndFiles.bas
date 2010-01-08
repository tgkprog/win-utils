Attribute VB_Name = "CmdFiles"
''''test cases
'''all files
' "D:\usrs\tushar\docs\resume\sel2in-recruit\09\2\Copy 3 CV Jude C.doc" "D:\usrs\tushar\docs\resume\sel2in-recruit\09\2\CV Jude C.doc" D:\usrs\tushar\docs\resume\sel2in-recruit\09\2\LavanyaGururaj-cv.doc D:\usrs\tushar\docs\rme\sel2init\09\2\
''' a folder
'"D:\usrs\tushar\docs\resume\sel2in-recruit\09\2\
Option Explicit
Public fso As FileSystemObject
Dim ff As File
Dim fOut As TextStream
Public Const App_CAP = "Path to Text file or Clipboard"
Sub Main()
On Local Error GoTo errh
Dim frm As Form1
Set frm = New Form1
Dim s, args
Set fso = New FileSystemObject

If Command() = "/set" Or Command() = "" Then
    frm.Show
    Exit Sub
End If

args = GetCommandLine

Dim i, max
Dim fld As Folder
Dim sParentFolder As String
If fso.FileExists(args(1)) Then
    Set ff = fso.GetFile(args(1))
    sParentFolder = ff.ParentFolder
ElseIf fso.FolderExists(args(1)) Then
    Set fld = fso.GetFolder(args(1))
    sParentFolder = fld.Path
Else
    frm.Show
    MsgBox "Invalid file/ folder in parameter 1, please see usage", vbInformation, App_CAP
    Exit Sub
End If

max = UBound(args)
If frm.optFile Then
    If InStr(1, frm.txtFileName, ":") > 0 Then
        Set fOut = fso.OpenTextFile(frm.txtFileName, ForWriting, True)
    Else
        Set fOut = fso.OpenTextFile(sParentFolder & "\" & frm.txtFileName, ForWriting, True)
    End If
End If
#If dbg = 1 Then
    Dim tx As TextStream
    Set tx = fso.OpenTextFile("d:\tmp\g", ForWriting, True)
    tx.WriteLine Command()
    tx.Close
#End If
s = ""
For i = 1 To max
    If frm.optFile Then
        fOut.WriteLine args(i)
    Else
        s = s & args(i) & vbNewLine
    End If
Next
If frm.optFile Then
    fOut.Close
Else
    Clipboard.Clear
    Clipboard.SetText s
End If
End
Exit Sub
errh:
On Local Error Resume Next
Dim sy, syt, a
a = Environ("USERPROFILE") & "\SendTo\" & App.EXEName & ".exe"
syt = vbOKOnly
If Not fso.FileExists(a) Then
    sy = vbNewLine & "*** Press Yes if you want me to copy my self to that folder."
    syt = vbYesNo
End If
s = "Usage: pass files as arguments, whose full path will be written to a file called """ & frm.txtFileName & """  in the path of the first file " _
& vbNewLine & "Place this exe in your send to folder for easy usage. " _
& vbNewLine & "Your send to folder is """ & Environ("USERPROFILE") & "\SendTo\ """ _
& sy _
& vbNewLine & "- http://code.google.com/p/win-utils/" & vbNewLine & "Tushar Kapila http://sel2in.com Copyright 2009"
#If dbg = 1 Then
    Resume
#End If
If Command() = "/set" Then
    frm.Show
    i = MsgBox(s _
        , vbInformation Or syt)
        If i = vbYes Then
            On Local Error Resume Next
            Call fso.CopyFile(App.Path & "\" & App.EXEName & ".exe", a, True)
    End If
ElseIf Not Command() = "" Then
    i = MsgBox(s _
        & vbNewLine & Err.Number & " " & Err.Description, vbInformation)
        If i = vbYes Then
            On Local Error Resume Next
            Call fso.CopyFile(App.Path & "\" & App.EXEName, a, True)
    End If
End If
End

End Sub
Function GetCommandLine(Optional MaxArgs)
   'Declare variables.
   Dim C, CmdLine, CmdLnLen, InArg, i, NumArgs
   'See if MaxArgs was provided.
   If IsMissing(MaxArgs) Then MaxArgs = 990
   'Make array of the correct size.
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   'Get command line arguments.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
    Debug.Print CmdLine
   'Go thru command line one character
   'at a time.
   Dim lastSep
   Dim bInQuote
   Dim d As String
   d = ""
   For i = 1 To CmdLnLen
      C = Mid(CmdLine, i, 1)
      Debug.Print C
      'Test for space or tab.
      If ((C <> " " And C <> vbTab)) Or (bInQuote Or d = "") Then
         'Neither space nor tab.
         'Test if already in argument.
         
         If Not InArg Then
         'New argument begins.
         'Test for too many arguments.
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
            If C = vbTab Or C = """" Or C = " " Then
                'lastSep = C
            End If
         End If
         'Concatenate character to current argument.
         If C <> """" Then
            ArgArray(NumArgs) = ArgArray(NumArgs) & C
        Else
            bInQuote = Not bInQuote
        End If
      Else
        'Found a space or tab.
        'Set InArg flag to False.
        InArg = False
        bInQuote = False
        If C = vbTab Or C = " " Then
            NumArgs = NumArgs + 1
            lastSep = C
            InArg = True
        End If
        'If C = """" Then bInQuote = True
      End If
      
nxt:
    d = C
   Next i
   'Resize array just enough to hold arguments.
   ReDim Preserve ArgArray(NumArgs)
   'Return Array in Function name.
   GetCommandLine = ArgArray()
End Function


