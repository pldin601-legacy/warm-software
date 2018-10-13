Attribute VB_Name = "mp3"
Sub GetMP3idTAG(Filename As String, idTITLE As String, idARTIST As String, idSUB As String, idYEAR As String, idCOMMENT As String, idGERNE As Integer)
On Error Resume Next
Dim I As Integer
I = FreeFile
Dim idByte As String
' Dim idRate As Integer

Dim TAGData As String, P As Integer

If FileExists(Filename) = False Then Exit Sub
Open Filename For Binary As #I

' frm.iSize.Caption = Format(LOF(I), "### ### ### ##0") + " bytes"

P = 0
TAGData = ""
idByte = String(128, 32)

If LOF(I) < 128 Then Exit Sub

Get #I, LOF(I) - 127, idByte
TAGData = idByte


Close #I

Dim idTOO As String, idTAG As String

idTOO = TAGData
idTAG = Right$(TAGData, 128 - 3)

If Mid(TAGData, 1, 3) <> "TAG" Then: Exit Sub

idTITLE = TrimN(Trim(Mid$(idTAG, 1, 30)))
idARTIST = TrimN(Trim(Mid$(idTAG, 31, 30)))
idSUB = TrimN(Trim(Mid$(idTAG, 61, 30)))
idYEAR = TrimN(Trim(Mid$(idTAG, 91, 4)))
idCOMMENT = TrimN(Trim(Mid$(idTAG, 95, 30)))
idGERNE = Asc(Mid$(idTAG, 125, 1))

End Sub


Sub GetMP3Bitrate(Filename As String, idBit As Integer)
On Error Resume Next
Dim I As Integer
I = FreeFile
Dim idByte As String * 1
' Dim idRate As Integer

If FileExists(Filename) = False Then Exit Sub
Open Filename For Binary As #I
If LOF(I) < 128 Then Exit Sub
Get #I, 3, idByte
Close #I

If idByte > 0 And idByte <= 100 Then idBit = 56
If idByte > 100 And idByte <= 200 Then idBit = 128
If idByte > 200 And idByte <= 255 Then idBit = 256


End Sub


Sub PutMP3idTAG(ByVal Filename As String, ByVal idTITLE As String, ByVal idARTIST As String, ByVal idSUB As String, ByVal idYEAR As String, ByVal idCOMMENT As String, idGERNE As Integer)
On Error Resume Next
Dim I As Integer
I = FreeFile
Dim idByte As String * 1, Path As String
Dim TAGData As String, P As Integer

Open Filename For Random As #I Len = 1

If Err Then MsgBox "Error opening the file!", vbExclamation

P = 0

For X = LOF(I) - 127 To LOF(I)
 Get #I, X, idByte
 TAGData = TAGData + idByte
Next

Path = String(128, 32)
Mid$(Path, 1, 3) = "TAG"
Mid$(Path, 4, 30) = idTITLE
Mid$(Path, 34, 30) = idARTIST
Mid$(Path, 64, 30) = idSUB
Mid$(Path, 94, 4) = idYEAR
Mid$(Path, 98, 30) = idCOMMENT
Mid$(Path, 128, 1) = Chr(idGERNE)

If Mid(TAGData, 1, 3) = "TAG" Then
P = 0
  For X = LOF(I) - 127 To LOF(I)
   P = P + 1
   idByte = Mid(Path, P, 1)
   Put #I, X, idByte
  Next
Else
P = 0
  For X = LOF(I) To LOF(I) + 127
   P = P + 1
   idByte = Mid(Path, P, 1)
   Put #I, X, idByte
  Next
End If


Close #I


End Sub


Function Trim32(InString As String)

InSrting = Trim(InString)

For A = Len(InString) To 1 Step -1
 If Asc(Mid(InString, A, 1)) > 32 Then InString = Mid(InString, 1, A): Exit For
Next

For A = 1 To Len(InString) Step 1
 If Asc(Mid(InString, A, 1)) < 32 Then Trim32 = Mid(InString, 1, A - 1): Exit Function
Next

Trim32 = InString

End Function


Function TrimN(InString As String) As String
 Dim A, B, C
 
 For A = 1 To Len(InString)
  For B = 1 To 31
   C = InStr(1, InString, Chr(B))
   Do While C > 0
     C = InStr(1, InString, Chr(B))
     Mid$(InString, C, 1) = Chr(0)
   Loop
  Next
 Next
 
 TrimN = Trim32(InString)
 
End Function


