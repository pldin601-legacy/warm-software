Attribute VB_Name = "mp3"
Type IdBitRate
 Zero As String * 1
 Bits As Integer
End Type

Dim idRate As IdBitRate
Sub GetMP3idTAG(Filename As String, idTITLE As String, idARTIST As String, idSUB As String, idYEAR As String, idCOMMENT As String)
On Error Resume Next
Dim I As Integer
I = FreeFile
Dim idByte As String
' Dim idRate As Integer

Dim TAGData As String, P As Integer

If FileExists(Filename) = False Then Exit Sub
Open Filename For Binary As #I

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

idTITLE = Trim(Mid$(idTAG, 1, 30))
idARTIST = Trim(Mid$(idTAG, 31, 30))
idSUB = Trim(Mid$(idTAG, 61, 30))
idYEAR = Trim(Mid$(idTAG, 91, 4))
idCOMMENT = Trim(Mid$(idTAG, 95, 30))

End Sub

Sub PutMP3idTAG(ByVal Filename As String, ByVal idTITLE As String, ByVal idARTIST As String, ByVal idSUB As String, ByVal idYEAR As String, ByVal idCOMMENT As String)
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

Path = String(127, 32)
Mid$(Path, 1, 3) = "TAG"
Mid$(Path, 4, 30) = idTITLE
Mid$(Path, 34, 30) = idARTIST
Mid$(Path, 64, 30) = idSUB
Mid$(Path, 94, 4) = idYEAR
Mid$(Path, 98, 30) = idCOMMENT + Chr(0)

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


