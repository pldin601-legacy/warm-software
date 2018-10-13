VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form HTM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Playlist Генератор"
   ClientHeight    =   3015
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6015
   Icon            =   "Html.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   5280
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save HTML as"
      Filter          =   "HTML Files (*.htm;*.html) | *.htm;*.html"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Опції"
      Height          =   1155
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   5775
      Begin VB.CheckBox zTest 
         Caption         =   "Відкрити"
         Height          =   195
         Left            =   2880
         TabIndex        =   10
         Top             =   720
         Width           =   2475
      End
      Begin VB.CheckBox zRefs 
         Caption         =   "Екстра"
         Height          =   255
         Left            =   540
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox zTitle 
         Height          =   285
         Left            =   540
         TabIndex        =   7
         Text            =   "PLAYLIST"
         Top             =   300
         Width           =   5115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Титл:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ім'я файлу"
      Height          =   915
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   4455
      Begin VB.TextBox zFilename 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "default.htm"
         Top             =   360
         Width           =   3675
      End
      Begin VB.CommandButton cmSaveAs 
         Caption         =   "<...>"
         Height          =   315
         Left            =   3780
         TabIndex        =   4
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Відмінити"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Генератор"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "HTML GENERATOR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   5775
   End
End
Attribute VB_Name = "HTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cmSaveAs_Click()

On Error Resume Next

CDialog.ShowSave
If Not Err Then zFilename.Text = CDialog.Filename

End Sub

Sub Generate()
On Error Resume Next
Dim I, NN As Integer
I = FreeFile

Open zFilename.Text For Output As #I

' Generating Default Code
Print #I, "<HTML>"
Print #I, "<HEAD>"
Print #I, "<TITLE>"; zTitle.Text; "</TITLE>"
Print #I, "</HEAD>"
Print #I, "<BODY BGCOLOR="; Chr(34); "#000000"; Chr(34); ">"
Print #I, "<B><I><FONT FACE="; Chr(34); "Arial"; Chr(34); " SIZE=7 COLOR="; Chr(34); "#ff0000"; Chr(34); "><P>V</FONT><FONT FACE="; Chr(34); "Arial"; Chr(34); " SIZE=7 COLOR="; Chr(34); "#ffff00"; Chr(34); ">irtual </FONT><FONT FACE="; Chr(34); "Arial"; Chr(34); " SIZE=7 COLOR="; Chr(34); "#ff0000"; Chr(34); ">M</FONT><FONT FACE="; Chr(34); "Arial"; Chr(34); " SIZE=7 COLOR="; Chr(34); "#ffff00"; Chr(34); ">egaBox </FONT><FONT FACE="; Chr(34); "Arial"; Chr(34); " SIZE=7 COLOR="; Chr(34); "#ff0000"; Chr(34); ">P</FONT><FONT FACE="; Chr(34); "Arial"; Chr(34); " SIZE=7 COLOR="; Chr(34); "#ffff00"; Chr(34); ">laylist</P>"
Print #I, "</B></I></FONT><FONT COLOR="; Chr(34); "#000000"; Chr(34); "><P ALIGN="; Chr(34); "CENTER"; Chr(34); "><HR></P>"
Print #I, "</FONT><B><I><FONT FACE="; Chr(34); "Arial"; Chr(34); " COLOR="; Chr(34); "#ffff00"; Chr(34); ">"
Print #I, "<P>"; zTitle.Text; "</P>"
Print #I, "</B></I></FONT><FONT COLOR="; Chr(34); "#000000"; Chr(34); "><P ALIGN="; Chr(34); "CENTER"; Chr(34); "><HR></P>"
Print #I, ""
Print #I, "</FONT><B><I><FONT FACE="; Chr(34); "Arial"; Chr(34); " COLOR="; Chr(34); "#ffff00"; Chr(34); ">"

Dim A
If zRefs.Value = 0 Then
 For A = 0 To wndMain.lstSec.ListCount - 1
  Print #I, "<LI>"; A + 1; ". "; wndMain.lstSec.List(A); "</LI>"
 Next A
Else
 For A = 0 To wndMain.lstSec.ListCount - 1
  Print #I, "<LI><A HREF="; Chr(34); wndMain.lstMain.List(A); Chr(34); ">"; A + 1; ". "; GetMp3Song(wndMain.lstMain.List(A), NN); "</A></LI>"
 Next A
End If


Print #I, "</B></I></FONT><FONT COLOR="; Chr(34); "#000000"; Chr(34); "><P ALIGN="; Chr(34); "CENTER"; Chr(34); "><HR></P>"
Print #I, "</FONT><B><I><FONT FACE="; Chr(34); "Arial"; Chr(34); " COLOR="; Chr(34); "#ffff00"; Chr(34); "><P>RENNSoft Multimedia (C) 2001-2003</P></B></I></FONT></BODY>"
Print #I, "</HTML>"

Close #I

If Err Then MsgBox "Working Error! Filename error!", vbCritical: Exit Sub
If Not Err And zTest.Value = 1 Then A = Shell("explorer.exe " + zFilename.Text, vbMaximizedFocus)
If Not Err Then Me.Hide

End Sub

Private Sub OKButton_Click()

If zFilename.Text = "" Then MsgBox "Filename not specified?", vbExclamation, "Filename error": Exit Sub

If FileExists(zFilename.Text) = True Then
 Dim RV As Integer
 RV = MsgBox("File allready exists, overwrite?", vbExclamation + vbYesNo, "File exists")
 If RV = 6 Then Generate
 If RV = 7 Then cmSaveAs_Click: Exit Sub
 Exit Sub
Else
 Generate
End If

End Sub


