VERSION 5.00
Begin VB.Form lstRecent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Вибрані списки"
   ClientHeight    =   3135
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "lstRecent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstLists 
      Height          =   2400
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Відмінити"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton btnDelRec 
      Caption         =   "Витерти"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "lstRecent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Sub UpdateListBox()
  On Error Resume Next
  
  Dim xLit As String
  lstLists.Clear
  
  xLit = Dir(LowPath(GlbRecentDir) + "r_*.mxt", vbNormal)
  If xLit = "" Then Exit Sub
  lstLists.AddItem Mid(xLit, 3, Len(xLit) - 4 - 2)
  
  Do While xLit > ""
    xLit = Dir
    If xLit = "" Then Exit Do
    lstLists.AddItem Mid(xLit, 3, Len(xLit) - 4 - 2)
  Loop

  If lstLists.ListCount > 0 Then lstLists.ListIndex = 0
  
End Sub


Private Sub btnDelRec_Click()
On Error Resume Next
Dim Quest As Integer

Quest = MsgBox("Ви дійсно бажаєте СТЕРТИ цей файл?", vbExclamation + vbYesNo, "Delete Recent Playlist")

If Quest = 6 Then
  Kill LowPath(GlbRecentDir) + "r_" + lstLists.List(lstLists.ListIndex) + ".mxt"
  If Err Then MsgBox "Видалення не було завершено успішно! Файл не буде видалений!", vbCritical, "Error!"
End If

UpdateListBox

End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Show
UpdateListBox
lstLists_Click
End Sub


Private Sub lstLists_Click()
If lstLists.ListIndex < 0 Then
  OKButton.Enabled = False
  btnDelRec.Enabled = False
Else
  OKButton.Enabled = True
  btnDelRec.Enabled = True
End If
End Sub

Private Sub OKButton_Click()

wndMain.LoadList LowPath(GlbRecentDir) + "r_" + lstLists.List(lstLists.ListIndex) + ".mxt", wndMain.lstMain, wndMain
Unload Me

End Sub


