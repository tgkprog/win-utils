Attribute VB_Name = "main1"
Dim oFrm1 As Frm1
Public bUnat As Boolean
Sub Main()
Dim s
s = Command$
Set oFrm1 = New Frm1
If s = "" Then
    bUnat = False
    oFrm1.Show
Else
    bUnat = True
    parse s
End If
End Sub

Sub parse(s)
Dim aa

aa = Split(s, " ")
oFrm1.txtVolNow = aa(0)
oFrm1.txtWaitSec = aa(1)
oFrm1.txtVolLater = aa(2)
oFrm1.cmdSet_Click
End Sub
