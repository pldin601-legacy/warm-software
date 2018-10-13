VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form dlgCopyFile 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Копіювати виділені файли..."
   ClientHeight    =   5625
   ClientLeft      =   2850
   ClientTop       =   3765
   ClientWidth     =   6930
   Icon            =   "dlgCopyFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer updss 
      Interval        =   250
      Left            =   6360
      Top             =   4380
   End
   Begin VB.CheckBox chkDelete 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Витерати оригінали"
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Left            =   3720
      TabIndex        =   11
      Top             =   4380
      Width           =   3075
   End
   Begin VB.Frame lblCopyState 
      BackColor       =   &H00C0C0C0&
      Height          =   1155
      Left            =   120
      TabIndex        =   9
      Top             =   3780
      Width           =   3495
      Begin ComctlLib.ProgressBar Progress 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar DiskSpace 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1500
         TabIndex        =   14
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label fds 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Вільно на гвинті:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Width           =   1260
      End
   End
   Begin VB.CommandButton CDButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Створити папочку"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5100
      Width           =   3615
   End
   Begin VB.CheckBox chkRePath 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Інформувати плеєр про зміну місця проживання файлів"
      CausesValidation=   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3840
      Width           =   3075
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3255
      Left            =   3660
      TabIndex        =   4
      Top             =   60
      Width           =   3135
      Begin VB.ListBox lstFiles 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   2985
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   5
         Top             =   180
         Width           =   2895
      End
   End
   Begin VB.DriveListBox lstDriveCopy 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   3435
   End
   Begin VB.DirListBox lstCopyPath 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3435
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Ні, я передумав"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5100
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Почати"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5100
      Width           =   1455
   End
   Begin VB.Label lblPath 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3420
      Width           =   6675
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   "Выделить все"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelectDown 
         Caption         =   "Выделить вниз"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSelectUp 
         Caption         =   "Выделить вверх"
         Shortcut        =   ^U
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnSelect 
         Caption         =   "Отменить выделение"
         Shortcut        =   ^B
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReverse 
         Caption         =   "Обратить выделение"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "dlgCopyFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function CopyFile(ByVal SourcePath As String, ByVal DestinationPath As String, ByVal SFileName As String, ByVal DFileName As String) As Boolean
    Dim Index As Currency
    Dim FileLength As Long
    Dim LeftOver As Long
    Dim FileData As String
    Dim NumBlocks As Long
    Dim X As Integer

    Screen.MousePointer = 11
    If Right$(SourcePath$, 1) <> "\" Then
        SourcePath$ = SourcePath$ + "\"            'Add ending \ symbols to path variables
    End If
    If Right$(DestinationPath$, 1) <> "\" Then
        DestinationPath$ = DestinationPath$ + "\"   'Add ending \ symbols to path variables
    End If
    
    'Update status dialog info
    '

    If Not FileExists(SourcePath$ + SFileName$) Then
          GoTo ErrorCopy
    End If
    
    On Local Error GoTo ErrorCopy


    'Copy the file
    '
    Progress.Value = Progress.Min
    Const BlockSize = 32768
    
    Open SourcePath$ + SFileName$ For Binary Access Read As #1
    
    If LOF(1) > GotFreeDiskSpace(Left(lstCopyPath.Path, 3)) Then
      MsgBox "Помилка: " + vbCrLf + "- неможливо скопіювати файл" + vbCrLf + vbCrLf + "Причина:" + vbCrLf + "- на вашому гвинті замало місця для продовження процесу копіювання виділених файлів" + vbCrLf + vbCrLf + "Що робити:" + vbCrLf + "1. Почистіть корзину" + vbCrLf + "2. Запустіть старий добрий Провідник і почистіть гвинт, стерши непотрібні файли", vbCritical
      GoTo ErrorCopy
    End If
    
    Open DestinationPath$ + DFileName$ For Output As #2
    Close #2
    Open DestinationPath$ + DFileName$ For Binary As #2
    
    FileLength = LOF(1)
    NumBlocks = FileLength \ BlockSize
    LeftOver = FileLength Mod BlockSize
    FileData = String$(LeftOver, 32)
    
    Get #1, , FileData
    Put #2, , FileData
    
    FileData = String$(BlockSize / 2, 32)
    
    Screen.MousePointer = 11
    
    For Index = 1 To (NumBlocks * 2)
        Get #1, , FileData
        Put #2, , FileData
        Progress.Max = NumBlocks
        Progress.Value = Index / 2
        DoEvents
        If Progress.Tag = "Cancel" Then Exit Function
    Next Index
    
    Progress.Value = Progress.Max
    Screen.MousePointer = 0
    FileData = ""    ' Free up String Allocation
    Close #1, #2
    
    

SkipCopy:
    CopyFile = True

ExitCopyFile:
    Screen.MousePointer = 0
    Exit Function
    
ErrorCopy:
    CopyFile = False
    Close
    Screen.MousePointer = 0
    Exit Function

End Function

Function FileExists(Path$) As Integer
    Dim X As Integer

    X = FreeFile

    On Error Resume Next
    Open Path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X

End Function

Private Sub CancelButton_Click()
If Progress.Tag <> "Copying" Then Unload Me
If Progress.Tag = "Copying" Then Progress.Tag = "Cancel"
End Sub

Private Sub CDButton_Click()
On Error Resume Next
Dim CPath As String

CPath = InputBox("Необхідно ввести ім'я для новоствореної папочки:", "Створення папочки", "Папочка")
If CPath = "" Then Exit Sub
MkDir LowPath(lstCopyPath.Path) + CPath
If Err Then MsgBox "Помилка:" + vbCrLf + "- При створенні папочки виникли незначні ускладнення" + vbCrLf + vbCrLf + "Причина:" + vbCrLf + "- Все, що завгодно" + vbCrLf + vbCrLf + "Що робити:" + vbCrLf + "- Самі викручуйтесь. Хто зна, що могло статися. Чи то така папочка вже існуе, чи то диск вже забитий, чи то в імені були допущені якісь неприйнятні символи, чи то ще якась фігня. Не знаю." + vbCrLf + vbCrLf + "Windows повідомляє:" + vbCrLf + "- " + Err.Description
lstCopyPath.Refresh
lstDriveCopy.Refresh
End Sub

Private Sub chkRePath_Click()
chkDelete.Enabled = chkRePath.Value
End Sub

Private Sub Form_Load()
On Error Resume Next: Dim X
lstFiles.Clear

For X = 0 To wndMain.lstMain.ListCount - 1
 Me.lstFiles.AddItem wndMain.lstMain.List(X)
Next

For X = 0 To wndMain.lstMain.ListCount - 1
 Me.lstFiles.Selected(X) = wndMain.lstMain.Selected(X)
Next

lblPath.Caption = LowPath(lstCopyPath.Path)


End Sub


Private Sub Form_Unload(Cancel As Integer)
Progress.Tag = "Cancel"
End Sub


Private Sub lstCopyPath_Change()
lblPath.Caption = LowPath(lstCopyPath.Path)
End Sub

Private Sub lstDriveCopy_Change()
On Error Resume Next
lstCopyPath.Path = lstDriveCopy.Drive
End Sub


Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuMenu
End Sub


Private Sub mnuReverse_Click()
Dim I, B
For I = 0 To lstFiles.ListCount - 1
 lstFiles.Selected(I) = CBool((1 - (lstFiles.Selected(I) + 1)) - 1)
Next
End Sub

Private Sub mnuSelect_Click()
Dim I, B
For I = 0 To lstFiles.ListCount - 1
 lstFiles.Selected(I) = True
Next
End Sub

Private Sub mnuSelectDown_Click()
Dim I, B
B = lstFiles.ListIndex
For I = B To lstFiles.ListCount - 1
 lstFiles.Selected(I) = True
Next

End Sub

Private Sub mnuSelectUp_Click()
Dim I, B
B = lstFiles.ListIndex
For I = 0 To B
 lstFiles.Selected(I) = True
Next
End Sub


Private Sub mnuUnSelect_Click()
Dim I, B
For I = 0 To lstFiles.ListCount - 1
 lstFiles.Selected(I) = False
Next

End Sub


Private Sub OKButton_Click()
On Error Resume Next
If Progress.Tag = "Copying" Then Exit Sub
Progress.Tag = "Copying"

Dim A, B, X As Boolean, Mp3Name
For A = 0 To lstFiles.ListCount - 1
 lblCopyState.Caption = "Копіюю " & (A + 1) & " з " _
 & lstFiles.ListCount
 DoEvents
  
 
 If lstFiles.Selected(A) = True Then
 
   Mp3Name = FileHead(lstFiles.List(A))
   
   lblPath.Caption = LowPath(lstCopyPath.Path) + _
   Mp3Name: X = CopyFile(PathHead(lstFiles.List(A)), _
   lstCopyPath.Path, FileHead(lstFiles.List(A)), Mp3Name)
   
   If Progress.Tag = "Cancel" Then Exit For
   
   If X = True Then
     If chkRePath.Value Then
        wndMain.lstMain.List(A) = LowPath(lstCopyPath.Path) + Mp3Name
        If chkDelete.Value Then
          Kill lstFiles.List(A)
        End If
     End If
   Else
     MsgBox "Помилки знайдені!", vbCritical: Exit Sub
   End If
   
   If Err Then
     B = MsgBox("Windows повідомляє:" + vbCrLf + "- " + Err.Description + " продовжити процес?", vbCritical + vbYesNo)
     Err = 0
     If B = 7 Then Progress.Tag = "": Exit Sub
   End If
   
 End If

Next

Unload dlgCopyFile

End Sub


Private Sub updss_Timer()
DiskSpace.Max = 100
DiskSpace.Value = 100 / GotTotalDiskSpace(Left(lstCopyPath.Path, 3)) * GotFreeDiskSpace(Left(lstCopyPath.Path, 3))
Label1.Caption = Format(GotFreeDiskSpace(Left(lstCopyPath.Path, 3)), "### ### ### ##0") + " bytes"
End Sub


