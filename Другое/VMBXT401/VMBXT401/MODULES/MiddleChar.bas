Attribute VB_Name = "MiddleChar"
Global Const OUT_MAX = 1
Global Const OUT_MIN = 2
Global Const OUT_MID = 0

Function GetMiddleChar(CharSet As String, outputType As Integer)

Dim Char As String, CharCount As Integer, Chared As Integer
Dim Calc As Currency

CharCount = Len(CharSet)

If outputType = 0 Then
Calc = 0
 For Chared = 1 To CharCount
 Calc = Calc + Asc(Mid(CharSet, Chared, 1))
 Next
 GetMiddleChar = Chr(Fix(Calc / CharCount))
End If

If outputType = 1 Then
Calc = 0
 For Chared = 1 To CharCount
  If Calc < Asc(Mid(CharSet, Chared, 1)) Then Calc = Asc(Mid(CharSet, Chared, 1))
 Next
 GetMiddleChar = Chr(Calc)
End If

If outputType = 2 Then
Calc = 255
 For Chared = 1 To CharCount
  If Calc > Asc(Mid(CharSet, Chared, 1)) Then Calc = Asc(Mid(CharSet, Chared, 1))
 Next
 GetMiddleChar = Chr(Calc)
End If

End Function

