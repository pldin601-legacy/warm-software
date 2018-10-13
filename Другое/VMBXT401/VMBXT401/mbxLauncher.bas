Attribute VB_Name = "SMain"
Sub Main()
On Error Resume Next
Dim a
a = Shell(LowPath(App.path) + "vmbxt401.exe", vbNormalFocus)
If Err Then MsgBox "Can't launch the programm Virtual MegaBox XT 4.01. May be not installed!", vbCritical
End Sub

