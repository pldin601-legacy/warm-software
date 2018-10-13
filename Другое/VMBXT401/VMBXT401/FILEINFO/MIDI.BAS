Attribute VB_Name = "midi"
Type MIDI_FORMAT
 MIDI_MThd As String * 14
 MIDI_MTrk As String * 38
End Type


Function MIDI_NAME(Filename As String) As String
On Error Resume Next
Dim I As Integer
Dim MIDI_HDR As MIDI_FORMAT

I = FreeFile
Open Filename For Random As #I Len = Len(MIDI_HDR)
Get #I, 1, MIDI_HDR

If Mid$(MIDI_HDR.MIDI_MTrk, 1, 4) <> "MTrk" And Mid$(MIDI_HDR.MIDI_MThd, 1, 4) <> "MThd" Then Close #I: Exit Function

MIDI_NAME = Mid$(MIDI_HDR.MIDI_MTrk, 13, nextzero(Mid$(MIDI_HDR.MIDI_MTrk, 13)))

Close #I

End Function

Function Midi_Size(Filename As String) As String
On Error Resume Next
I = FreeFile
Open Filename For Input As #I
Midi_Size = Format(LOF(I) / 1000, "0") + "K"
Close #I
End Function

Function nextzero(stn As String) As Integer
For X = 1 To Len(stn)
If Asc(Mid$(stn, X, 1)) < 32 Then nextzero = X - 1: Exit Function
Next
End Function

