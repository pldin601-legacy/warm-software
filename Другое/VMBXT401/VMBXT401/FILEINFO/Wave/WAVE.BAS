Attribute VB_Name = "wave"
Sub GetWaveFormat(FileName As String, Channels As Integer, Sample As Integer)

On Error Resume Next
Dim I As Integer
I = FreeFile
Dim idByte As String * 1
' Dim idRate As Integer

Dim TAGData As String, P As Integer

Open FileName For Random As #I Len = 1

P = 0
TAGData = ""

If LOF(I) < 128 Then Exit Sub
For X = LOF(I) - 127 To LOF(I)
 P = P + 1
 Get #I, X, idByte
 TAGData = TAGData + idByte
Next


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


