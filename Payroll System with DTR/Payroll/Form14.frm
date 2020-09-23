VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2490
      TabIndex        =   3
      Top             =   1950
      Width           =   1395
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3930
      TabIndex        =   2
      Top             =   1950
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Security"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1230
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   5370
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   345
         Left            =   660
         TabIndex        =   5
         Top             =   750
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   609
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Label1"
         BuddyDispid     =   196615
         OrigLeft        =   720
         OrigTop         =   570
         OrigRight       =   960
         OrigBottom      =   915
         Max             =   1000
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "minute(s)"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   990
         TabIndex        =   7
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock paroll system when i did not move the mouse within"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   210
         TabIndex        =   6
         Top             =   210
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   4
         Top             =   750
         Width           =   465
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Application Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "Form14.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
sTimeSet = Val(Label1.Caption) * 60

SaveSetting App.Title, "SetVal", "TimeVal", sTimeSet
SaveSetting App.Title, "LabVal", "LastLabVal", Label1.Caption

Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = GetSetting(App.Title, "LabVal", "LastLabVal", 1)
End Sub
