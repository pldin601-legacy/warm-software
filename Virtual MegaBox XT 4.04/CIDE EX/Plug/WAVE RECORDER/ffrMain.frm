VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ffrREC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RENNSoft VMBX TOOLS: WAVE Audio Recorder"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "ffrMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "File Status"
      Height          =   1155
      Left            =   3780
      TabIndex        =   3
      Top             =   1920
      Width           =   3315
      Begin VB.Label outDisk 
         BackStyle       =   0  'Transparent
         Caption         =   "1 024 Kb"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   1140
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label outSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Opened"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   1140
         TabIndex        =   16
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label outPos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   405
      End
      Begin VB.Label outLen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   405
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   2640
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Cap4 
         AutoSize        =   -1  'True
         Caption         =   "Disk Free:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   720
      End
      Begin VB.Label cap3 
         AutoSize        =   -1  'True
         Caption         =   "File Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Cap2 
         AutoSize        =   -1  'True
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Cap1 
         AutoSize        =   -1  'True
         Caption         =   "Length:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Player Control"
      Height          =   1155
      Left            =   60
      TabIndex        =   1
      Top             =   1920
      Width           =   3615
      Begin MCI.MMControl MMControl1 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   1085
         _Version        =   393216
         UpdateInterval  =   250
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   "WAVEAudio"
         FileName        =   ""
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Recording Panel"
      Height          =   1755
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7035
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   5100
         TabIndex        =   19
         Top             =   780
         Width           =   1815
         Begin VB.OptionButton opOver 
            Caption         =   "O&verwrite"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   240
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton opInsert 
            Caption         =   "&Insert"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   540
            Width           =   735
         End
      End
      Begin VB.CommandButton buProperty 
         Caption         =   "&Format..."
         Height          =   375
         Left            =   5460
         TabIndex        =   18
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton buExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   4500
         TabIndex        =   11
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton buSaveAs 
         Caption         =   "Save &as..."
         Height          =   375
         Left            =   3540
         TabIndex        =   10
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton buSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   2580
         TabIndex        =   9
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton buOpen 
         Caption         =   "&Open..."
         Height          =   375
         Left            =   1620
         TabIndex        =   8
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton buNew 
         Caption         =   "&New..."
         Height          =   375
         Left            =   660
         TabIndex        =   7
         Top             =   360
         Width           =   915
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   630
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   1111
         _Version        =   327682
         TickStyle       =   2
         TickFrequency   =   30
      End
   End
End
Attribute VB_Name = "ffrREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = Me.Caption + " " + GetVersion
End Sub
