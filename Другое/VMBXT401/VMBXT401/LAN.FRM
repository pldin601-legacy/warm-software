VERSION 5.00
Begin VB.Form LangForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LANGUAGE"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   ClipControls    =   0   'False
   Icon            =   "lan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "SELECT"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2700
      Width           =   3555
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2235
      Left            =   0
      Pattern         =   "*.vln"
      TabIndex        =   1
      Top             =   420
      Width           =   3555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SELECT LANGUAGE:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "LangForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim X, a
X = FreeFile
If File1.Filename > "" Then

  Open LowPath(App.Path) + "language.sel" For Output As #X
  Print #X, "LANG_FILE:"; File1.Filename;
  Close #X
  
End If

Unload Me

End Sub


Private Sub File1_DblClick()
Command1_Click
End Sub


Private Sub Form_Load()
File1.Path = App.Path
End Sub


