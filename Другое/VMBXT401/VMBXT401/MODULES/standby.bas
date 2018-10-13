Attribute VB_Name = "SD"
'' STANDBY MODULE ''

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_SYSCOMMAND = &H112
' Public Const SC_SCREENSAVE = &HF140&
Public Const SC_SCREENSAVE = &HF140&

