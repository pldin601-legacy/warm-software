VERSION 5.00
Begin VB.Form frmNav 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Media Navigator"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Re&load"
      Height          =   375
      Left            =   4830
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5970
      TabIndex        =   4
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   4935
      Left            =   4470
      TabIndex        =   1
      Top             =   510
      Width           =   2595
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   30
      TabIndex        =   0
      Top             =   510
      Width           =   4365
   End
   Begin VB.Label Label2 
      Caption         =   "Recent playlists:"
      Height          =   225
      Left            =   4500
      TabIndex        =   3
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Files in player list:"
      Height          =   195
      Left            =   30
      TabIndex        =   2
      Top             =   240
      Width           =   1725
   End
End
Attribute VB_Name = "frmNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
 
 Dim G, H, J
 
 List2.Clear
 List1.Clear
 
 Dim xLit As String
 
 H = 0
 xLit = Dir(LowPath(GlbRecentDir) + "r_*.mxt", vbNormal)
 If xLit = "" Then Exit Sub
  
 List2.List(0) = Mid(xLit, 3, Len(xLit) - 4 - 2)
  
 Do While xLit > ""
    H = H + 1
    xLit = Dir
    List2.List(H) = Mid(xLit, 3, Len(xLit) - 4 - 2)
 Loop

For N = 0 To wndMain.lstSec.ListCount - 1
    List1.List(N) = wndMain.lstSec.List(N)
Next N

End Sub

Private Sub List1_DblClick()
BackIndex = List1.ListIndex
wndMain.Command_Open: wndMain.Command_Play
End Sub

Private Sub List2_Click()
If FileExists(LowPath(GlbRecentDir) + "r_" + List2.List(List2.ListIndex) + ".mxt") = True Then
 wndMain.LoadList LowPath(GlbRecentDir) + "r_" + List2.List(List2.ListIndex) + ".mxt", wndMain.lstMain, wndMain
End If
Command2_Click
End Sub

