VERSION 5.00
Begin VB.Form WinSeek 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Пошук мадіа"
   ClientHeight    =   2835
   ClientLeft      =   1920
   ClientTop       =   1890
   ClientWidth     =   6945
   ControlBox      =   0   'False
   ForeColor       =   &H00000080&
   Icon            =   "seek.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Специальный поиск (Under construction)"
      Enabled         =   0   'False
      Height          =   1995
      Left            =   120
      TabIndex        =   13
      Top             =   2940
      Width           =   6675
      Begin VB.CheckBox chkTag 
         Caption         =   "Find MP3 ->"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtSize 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "0"
         Top             =   540
         Width           =   1335
      End
      Begin VB.Frame scrTAG 
         Caption         =   "Search MP3s for:"
         Height          =   1575
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   4935
         Begin VB.ComboBox lstGerne 
            BackColor       =   &H00008000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2820
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox chkGerne 
            Caption         =   "Gerne:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2820
            TabIndex        =   25
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtComment 
            BackColor       =   &H00C00000&
            Enabled         =   0   'False
            ForeColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2820
            MaxLength       =   30
            TabIndex        =   24
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox chkComment 
            Caption         =   "Comment:"
            Height          =   195
            Left            =   2820
            TabIndex        =   23
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox txtAlbum 
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   345
            Left            =   900
            MaxLength       =   30
            TabIndex        =   22
            Top             =   1080
            Width           =   1755
         End
         Begin VB.CheckBox chkAlbum 
            Caption         =   "Album"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtArtist 
            BackColor       =   &H0000FFFF&
            Enabled         =   0   'False
            Height          =   345
            Left            =   900
            MaxLength       =   30
            TabIndex        =   20
            Top             =   660
            Width           =   1755
         End
         Begin VB.CheckBox chkArtist 
            Caption         =   "Artist"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   675
         End
         Begin VB.TextBox txtTitle 
            BackColor       =   &H000000FF&
            Enabled         =   0   'False
            Height          =   345
            Left            =   900
            MaxLength       =   30
            TabIndex        =   18
            Top             =   240
            Width           =   1755
         End
         Begin VB.CheckBox chkTitle 
            Caption         =   "Title"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "Размер"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.CommandButton kbtOK 
      Caption         =   "Вибрати"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5580
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Пошук"
      Default         =   -1  'True
      Height          =   360
      Left            =   5580
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Вихід"
      Height          =   345
      Left            =   5580
      TabIndex        =   1
      Top             =   540
      Width           =   1200
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2895
      ScaleWidth      =   4515
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   4515
      Begin VB.ListBox lstFoundFiles 
         Height          =   2205
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   11
         Top             =   480
         Width           =   4275
      End
      Begin VB.Label lblCount 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblfound 
         Caption         =   "&Files Found:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   0
      Width           =   4515
      Begin VB.DriveListBox drvList 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   2355
      End
      Begin VB.DirListBox dirList 
         Height          =   1665
         Left            =   2040
         TabIndex        =   6
         Top             =   1020
         Width           =   2355
      End
      Begin VB.FileListBox filList 
         Height          =   2235
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtSearchSpec 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Text            =   "*.*"
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label lblCriteria 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search &Criteria:"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "WinSeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SearchFlag As Integer   ' Used as flag for cancel and other operations.

Private Sub chkAlbum_Click()
txtAlbum.Enabled = chkAlbum.Value

End Sub

Private Sub chkArtist_Click()
txtArtist.Enabled = chkArtist.Value
End Sub

Private Sub chkComment_Click()
txtComment.Enabled = chkComment.Value
End Sub

Private Sub chkGerne_Click()
lstGerne.Enabled = chkGerne.Value
End Sub

Private Sub chkSize_Click()
txtSize.Enabled = chkSize.Value
End Sub

Private Sub chkTag_Click()
scrTAG.Enabled = chkTag.Value
End Sub

Private Sub chkTitle_Click()
txtTitle.Enabled = chkTitle.Value
End Sub

Private Sub cmdExit_Click()
    If cmdExit.Caption = "E&xit" Then
        Unload Me
    Else                    ' If user chose Cancel, just end Search.
        SearchFlag = False
    End If
End Sub

Private Sub cmdSearch_Click()
' Initialize for search, then perform recursive search.
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  ' Check what the user did last.
    If cmdSearch.Caption = "&Reset" Then  ' If just a reset, initialize and exit.
        ResetSearch
        txtSearchSpec.SetFocus
        Exit Sub
    End If

    ' Update dirList.Path if it is different from the currently
    ' selected directory, otherwise perform the search.
    If dirList.Path <> dirList.List(dirList.ListIndex) Then
        dirList.Path = dirList.List(dirList.ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If

    ' Continue with the search.
    Picture2.Move 0, 0
    Picture1.Visible = False
    Picture2.Visible = True

    cmdExit.Caption = "Cancel"

    filList.Pattern = txtSearchSpec.Text
    FirstPath = dirList.Path
    DirCount = dirList.ListCount

    ' Start recursive direcory search.
    NumFiles = 0                       ' Reset found files indicator.
    result = DirDiver(FirstPath, DirCount, "")
    filList.Path = dirList.Path
    cmdSearch.Caption = "&Reset"
    cmdSearch.SetFocus
    cmdExit.Caption = "E&xit"
    If Me.lstFoundFiles.ListCount > 0 Then kbtOK.Enabled = True
End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim RetVal As Integer
    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiver = False            ' Set to True if there is an error.
    RetVal = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = dirList.ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = dirList.Path                      ' Save old path for next recursion.
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            ' Get to the node bottom.
            dirList.Path = dirList.List(DirsToPeek - 1)
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = dirList.Path                  ' If at root level, leave as is...
        Else
            ThePath = dirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For ind = 0 To filList.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath + filList.List(ind)
            
            
            lstFoundFiles.AddItem entry
            lstFoundFiles.Selected(lstFoundFiles.ListCount - 1) = True
            lblCount.Caption = Str(Val(lblCount.Caption) + 1)
            
            
        Next ind
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        dirList.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        MsgBox Error
        End
    End If
End Function

Private Sub DirList_Change()
    ' Update the file list box to synchronize with the directory list box.
    filList.Path = dirList.Path
End Sub

Private Sub DirList_LostFocus()
    dirList.Path = dirList.List(dirList.ListIndex)
End Sub

Private Sub DrvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

Private Sub Form_Load()
    Me.Show
    Picture2.Move 0, 0
    Picture2.Width = WinSeek.ScaleWidth
    Picture2.BackColor = WinSeek.BackColor
    lblCount.BackColor = WinSeek.BackColor
    lblCriteria.BackColor = WinSeek.BackColor
    lblfound.BackColor = WinSeek.BackColor
    Picture1.Move 0, 0
    Picture1.Width = WinSeek.ScaleWidth
End Sub


Private Sub ResetSearch()
    ' Reinitialize before starting a new search.
    lstFoundFiles.Clear
    lblCount.Caption = 0
    SearchFlag = False                  ' Flag indicating search in progress.
    Picture2.Visible = False
    cmdSearch.Caption = "&Search"
    cmdExit.Caption = "E&xit"
    Picture1.Visible = True
    dirList.Path = CurDir: drvList.Drive = dirList.Path ' Reset the path.
    kbtOK.Enabled = False
End Sub

Private Sub kbtOK_Click()
On Error Resume Next
Dim X, Y, NN As Integer
  
Y = wndMain.lstSec.ListIndex
For X = 0 To Me.lstFoundFiles.ListCount - 1
 If lstFoundFiles.Selected(X) = True Then
  wndMain.lstMain.AddItem lstFoundFiles.List(X)
  wndMain.lstSec.AddItem GetMp3Song(lstFoundFiles.List(X), NN)
  wndMain.lstTimes.AddItem Str(NN)
 End If
Next
wndMain.lstSec.ListIndex = Y

Unload Me

End Sub

Private Sub txtSearchSpec_Change()
On Error Resume Next
    ' Update file list box if user changes pattern.
    filList.Pattern = txtSearchSpec.Text
End Sub

Private Sub txtSearchSpec_GotFocus()
    txtSearchSpec.SelStart = 0          ' Highlight the current entry.
    txtSearchSpec.SelLength = Len(txtSearchSpec.Text)
End Sub

