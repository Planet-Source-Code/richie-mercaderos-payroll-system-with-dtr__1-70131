VERSION 5.00
Begin VB.Form Frm_backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Records"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_back 
      Caption         =   "Current backup"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   6855
      Begin VB.CommandButton cmdCancel 
         Height          =   735
         Left            =   5400
         MouseIcon       =   "Frm_backup.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Frm_backup.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cancel"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   735
         Left            =   4080
         MouseIcon       =   "Frm_backup.frx":06D2
         MousePointer    =   99  'Custom
         Picture         =   "Frm_backup.frx":0824
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Backup"
         Top             =   960
         Width           =   1095
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   240
         Pattern         =   "*.mdb"
         TabIndex        =   11
         Top             =   870
         Width           =   255
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   480
         TabIndex        =   1
         ToolTipText     =   "Select folder"
         Top             =   870
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   0
         ToolTipText     =   "Select diskdrive"
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "See the details of last backup"
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
         Left            =   3750
         TabIndex        =   14
         Top             =   2100
         Width           =   3045
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Backup"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label lbl_Status 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   3840
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame fra_last 
      Caption         =   "Last backup"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   3930
      Width           =   6855
      Begin VB.Label lbl_path 
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup path"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   735
         Left            =   1080
         TabIndex        =   10
         Top             =   1140
         Width           =   5655
      End
      Begin VB.Label lbl_lasttime 
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup time"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label lbl_lastdate 
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup date"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lbl_apath 
         BackStyle       =   0  'Transparent
         Caption         =   "Path   :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label lbl_time 
         BackStyle       =   0  'Transparent
         Caption         =   "Time   :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl_Date 
         BackStyle       =   0  'Transparent
         Caption         =   "Date   :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Make back-up for your reference before deleting fine informations at regular intervals and secure your informations."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   840
      TabIndex        =   17
      Top             =   120
      Width           =   6135
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "Frm_backup.frx":0DDD
      Top             =   270
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   120
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "Frm_backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'add ref microsoft script library for file system object
Dim Fsys As New FileSystemObject
Dim bckupFile As File

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdsave_click()
On Error Resume Next
    cmdSave.Enabled = False
    
    Form1.StatusBar1.Panels(1) = "Please Wait, Backup in Progress..."
    Dim destination As String
    Dim Source As String
    Dim currDate, currTime, sNewPath As String
    currDate = Format$(Now, "dd, mmm, yyyy")
    currTime = Format$(Now, "hh:mm:ss AM/PM")
    
    sNewPath = "NewMasterFile(" & Format(Date, "ddmmmyy") & " " & Format(Time, "hh-mm-ss am/pm") & ").mdb"
    
    destination = File1.Path & "\" & sNewPath
    Source = App.Path & "\MasterFile.mdb"
   
    Set bckupFile = Fsys.GetFile(finalpath)
    bckupFile.Attributes = Compressed
    Fsys.CopyFile Source, destination, True
    SaveSetting App.Title, "Settings", "BackupPath", destination
    SaveSetting App.Title, "Settings", "BackupDate", currDate
    SaveSetting App.Title, "Settings", "BackupTime", currTime
    lbl_Status.Caption = "Backup sucessfull"
    cmdSave.Enabled = True
    MsgBox "All data BackUp Succcessfully on disk", vbInformation, "Backup"
    Form1.StatusBar1.Panels(1).Text = ""
    Unload Me
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
    Dim lastPath As String
    Dim lastDate As String
    Dim lastTime As String
    File1.Visible = False

      If (View = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (View = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If
    'Read Registry for previous settings stored
    lastPath = GetSetting(App.Title, "Settings", "BackupPath")
    lastDate = GetSetting(App.Title, "Settings", "BackupDate")
    lastTime = GetSetting(App.Title, "Settings", "BackupTime")
    
  lbl_Status.Caption = "Select path and press Backup."

    If lastPath = "" Then
        lbl_path.Caption = "No Backup made previously"
        lbl_lastdate.Caption = " "
        lbl_lasttime.Caption = " "
    Else
        lbl_path.Caption = lastPath
        lbl_lastdate.Caption = lastDate
        lbl_lasttime.Caption = lastTime
    End If
End Sub

