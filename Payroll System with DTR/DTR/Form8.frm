VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time and Date Settings"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Date"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   105
      TabIndex        =   15
      Top             =   0
      Width           =   2745
      Begin VB.ComboBox Combo7 
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
         Left            =   945
         TabIndex        =   21
         Top             =   1260
         Width           =   1065
      End
      Begin VB.ComboBox Combo6 
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
         Left            =   945
         TabIndex        =   19
         Top             =   840
         Width           =   1065
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   945
         TabIndex        =   17
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Year:"
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
         Left            =   210
         TabIndex        =   20
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label Label8 
         Caption         =   "Day:"
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
         Left            =   315
         TabIndex        =   18
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label7 
         Caption         =   "Month:"
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
         Left            =   105
         TabIndex        =   16
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   4440
      Top             =   1320
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   3960
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3480
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   1320
   End
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   3000
      TabIndex        =   9
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Time"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "Form8.frx":0000
         Left            =   2040
         List            =   "Form8.frx":00B8
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "Form8.frx":01AC
         Left            =   120
         List            =   "Form8.frx":01D4
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox Combo5 
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
         ItemData        =   "Form8.frx":0208
         Left            =   1080
         List            =   "Form8.frx":02C0
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "Form8.frx":03B4
         Left            =   2880
         List            =   "Form8.frx":03BE
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Sec"
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
         Left            =   2040
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Hour"
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
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Min"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Label Label6 
      Caption         =   "To set the System Time, click the Set button first."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   14
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_changes As New ADODB.Recordset

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Date = Combo3.Text & " " & Combo6.Text & ", " & Combo7.Text
Time = Combo4.Text & ":" & Combo5.Text & ":" & Combo2.Text & " " & Combo1.Text

rs_changes.AddNew
    rs_changes.Fields("Username") = Form9.sUser
    rs_changes.Fields("OldDate") = Form9.sDate
    rs_changes.Fields("DateChanges") = FormatDateTime(Date, vbShortDate)
    rs_changes.Fields("OldTime") = Form9.sTime
    rs_changes.Fields("TimeChanges") = FormatDateTime(Time, vbLongTime)
    rs_changes.Fields("Description") = Form9.sDescription
rs_changes.Update

MsgBox "Date and Time Settings has been changed.", vbInformation, "System"
Unload Form9
Unload Me
End Sub

Private Sub Command2_Click()
Unload Form9
Unload Me
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Frame2.Enabled = True
Frame3.Enabled = True
Combo4.SetFocus
Command3.Enabled = False
End Sub

Private Sub Form_Activate()
Command1.SetFocus
End Sub

Private Sub Form_Load()
Call set_rec_getData(rs_changes, cn, "Select * From tblDateTimeChanges")

Combo3.Text = Format(Date, "mmmm")
Combo6.Text = Format(Now, "dd")
Combo7.Text = Format(Date, "yyyy")
Combo1.Text = Format(Now, "AM/PM")
Combo5.Text = Format(Minute(Now), "00")
Combo4.Text = Mid(Format(Now, "hh:mm:ss AM/PM"), 1, 2)
Call LoadMonth
Call LoadDays
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rs_changes = Nothing
End Sub

Private Sub Timer1_Timer()
Combo4.Text = Mid(Format(Now, "hh:mm:ss AM/PM"), 1, 2)
End Sub

Private Sub Timer2_Timer()
Combo2.Text = Format(Second(Now), "00")
End Sub

Private Sub Timer3_Timer()
Combo5.Text = Format(Minute(Now), "00")
End Sub

Private Sub Timer4_Timer()
Combo1.Text = Format(Now, "AM/PM")
End Sub

Private Sub LoadMonth()
With Combo3
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With
End Sub

Private Sub LoadDays()
Dim i As Integer

For i = 1 To 31
    Combo6.AddItem Format(i, "00")
Next i
End Sub
