VERSION 5.00
Begin VB.Form frmDel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Редагування улюблених композицій"
   ClientHeight    =   3600
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmDel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check2 
      Caption         =   "Відновлення композиції в базі"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Value           =   1  'Checked
      Width           =   3195
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Видалення композиції з бази"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   3195
   End
   Begin VB.ListBox lstBase 
      Height          =   2760
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   120
      Width           =   4395
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Відмінити"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Sub LoadM()

Dim G, H, J, L, I, K
 
On Error Resume Next

I = FreeFile

Open (LowPath(App.Path) + "favourites.vmb") For Random As #I Len = Len(RecFileHead)

J = LOF(I) / Len(RecFileHead)

For K = 1 To J
  Get #I, K, RecFileHead
  lstBase.List(K - 1) = Trim(RecFileHead.rfDescription)
  lstBase.Selected(K - 1) = RecFileHead.rfUsed
Next K

Close I

End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub


Private Sub Form_Load()
LoadM
End Sub

Private Sub OKButton_Click()

Dim J, I, K
 
On Error Resume Next

I = FreeFile

Open (LowPath(App.Path) + "favourites.vmb") For Random As #I Len = Len(RecFileHead)

J = LOF(I) / Len(RecFileHead)

For K = 1 To J
  Get #I, K, RecFileHead
  RecFileHead.rfUsed = lstBase.Selected(K - 1)
  Put #I, K, RecFileHead
Next K

Close I

Unload Me

End Sub


