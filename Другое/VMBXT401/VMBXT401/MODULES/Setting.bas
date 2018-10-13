Attribute VB_Name = "Setting"
' Global Settings
Global GlbPlayMode As Integer
Global GlbRepeat As Integer

' Title Windows Settings
Global GlbWndScroll As Integer
Global GlbWndFluent As Integer
Global GlbWndColor As OLE_COLOR
Global GlbWndFont As String

' Playlist Saving Settings
Global GlbCHFS As Integer
Global GlbCHPP As Integer
Global GlbCHPM As Integer
Global GlbCHTS As Integer
Global GlbCHCP As Integer

Global GlbSTM As Integer

Global GlbTitle As String
Global GlbTagLess As String
Global GlbPropTime As Long

Global GlbPause As Integer


Global GlbRecentDir As String
Global GlbSeeker As Integer
Global GlbNoTags As Integer


Sub GlbSaveSettings()
 On Error Resume Next
 I = FreeFile
 
 Open LowPath(App.Path) + "vmbxt.ini" For Output As #I
 If Err Then MsgBox "Error Saving Settings! File 'vmbxt.ini' unaccesible.", vbCritical: Close #I: Exit Sub
 
 Print #I, "[Global settings]"
 Print #I, "GLOBAL_PLAY_MODE=" + Format(GlbPlayMode, "0")
 Print #I, "GLOBAL_REPEAT_MODE=" + Format(GlbRepeat, "0")
 Print #I, ""
 Print #I, "[Title window settings]"
 Print #I, "GLOBAL_WINDOW_SCROLL="; Format(GlbWndScroll, "0")
 Print #I, "GLOBAL_WINDOW_FLUENT="; Format(GlbWndFluent, "0")
 Print #I, "GLOBAL_WINDOW_COLOR="; GlbWndColor
 Print #I, "GLOBAL_WINDOW_FONT="; GlbWndFont
 Print #I, ""
 Print #I, "[Playlist saving settings]"
 Print #I, "GLOBAL_FS=" + Format(GlbCHFS, "0")
 Print #I, "GLOBAL_PP=" + Format(GlbCHPP, "0")
 Print #I, "GLOBAL_PM=" + Format(GlbCHPM, "0")
 Print #I, "GLOBAL_TS=" + Format(GlbCHTS, "0")
 Print #I, "GLOBAL_CP=" + Format(GlbCHCP, "0")
 Print #I, ""
 Print #I, "[Other settings]"
 Print #I, "STM=" + Format(GlbSTM, "0")
 Print #I, "TITLE_FORMAT=" + GlbTitle
 Print #I, "TAGLESS_FORMAT=" + GlbTagLess
 Print #I, "PROPORTIONAL_TIME=" + Str(GlbPropTime)
 Print #I, "PAUSE=" + Format(GlbPause, "0")
 Print #I, "RECENT_DIR=" + Format(GlbRecentDir, "0")
 Print #I, "SEEKER=" + Format(GlbSeeker, "0")
 Print #I, "NOTAGS=" + Format(GlbNoTags, "0")
 
 Close #I
 
 wndMain.UpdateRecentMenu
 
End Sub


Sub GlbLoadSettings()
 
 GlbPlayMode = Val(GetIniRecord("GLOBAL_PLAY_MODE=", LowPath(App.Path) + "vmbxt.ini"))
 GlbRepeat = Val(GetIniRecord("GLOBAL_REPEAT_MODE=", LowPath(App.Path) + "vmbxt.ini"))
 
 GlbWndScroll = 1 'Val(GetIniRecord("GLOBAL_WINDOW_SCROLL=", LowPath(App.Path) + "vmbxt.ini"))
 GlbWndFluent = 1 ' Val(GetIniRecord("GLOBAL_WINDOW_FLUENT=", LowPath(App.Path) + "vmbxt.ini"))
 GlbWndColor = RGB(244, 244, 0) ' GetIniRecord("GLOBAL_WINDOW_COLOR=", LowPath(App.Path) + "vmbxt.ini")
 GlbWndFont = GetIniRecord("GLOBAL_WINDOW_FONT=", LowPath(App.Path) + "vmbxt.ini")
 
 GlbCHFS = Val(GetIniRecord("GLOBAL_FS=", LowPath(App.Path) + "vmbxt.ini"))
 GlbCHPP = Val(GetIniRecord("GLOBAL_PP=", LowPath(App.Path) + "vmbxt.ini"))
 GlbCHPM = Val(GetIniRecord("GLOBAL_PM=", LowPath(App.Path) + "vmbxt.ini"))
 GlbCHTS = Val(GetIniRecord("GLOBAL_TS=", LowPath(App.Path) + "vmbxt.ini"))
 GlbCHCP = Val(GetIniRecord("GLOBAL_CP=", LowPath(App.Path) + "vmbxt.ini"))
 GlbSTM = Val(GetIniRecord("STM=", LowPath(App.Path) + "vmbxt.ini"))
 GlbTitle = GetIniRecord("TITLE_FORMAT=", LowPath(App.Path) + "vmbxt.ini")
 GlbTagLess = GetIniRecord("TAGLESS_FORMAT=", LowPath(App.Path) + "vmbxt.ini")
 GlbPropTime = Val(GetIniRecord("PROPORTIONAL_TIME=", LowPath(App.Path) + "vmbxt.ini"))
 
 GlbPause = Val(GetIniRecord("PAUSE=", LowPath(App.Path) + "vmbxt.ini"))
 
 GlbRecentDir = GetIniRecord("RECENT_DIR=", LowPath(App.Path) + "vmbxt.ini")
 GlbSeeker = GetIniRecord("SEEKER=", LowPath(App.Path) + "vmbxt.ini")
 GlbNoTags = GetIniRecord("NOTAGS=", LowPath(App.Path) + "vmbxt.ini")
 
 wndMain.UpdateRecentMenu
 
End Sub



