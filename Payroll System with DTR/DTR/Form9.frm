VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Specify Description"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   765
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Type a specific description here."
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sUser As String
Public sDate As String
Public sTime As String
Public sDescription As String

Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "Please specify a description of this action.", vbExclamation, "System Require"
    Text1.SetFocus
    Exit Sub
End If
sDescription = Text1.Text
Me.Hide
Form8.Show vbModal
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
sDate = FormatDateTime(Now, vbShortDate)
sTime = FormatDateTime(Now, vbLongTime)

Label3.Caption = "Username: " & sUser
Label1.Caption = "Time: " & FormatDateTime(Now, vbLongTime)
Label2.Caption = "Date: " & FormatDateTime(Now, vbShortDate)
End Sub
