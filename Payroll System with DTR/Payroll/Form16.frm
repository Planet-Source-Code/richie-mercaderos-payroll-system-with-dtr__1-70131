VERSION 5.00
Begin VB.Form Form16 
   BorderStyle     =   0  'None
   Caption         =   "Form16"
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   LinkTopic       =   "Form16"
   ScaleHeight     =   780
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   6045
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   210
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Enter password:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   1875
      End
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If sCurrUser = "" Then
        If Text1.Text = "a" Then
            Unload Me
        Else
            MsgBox "Incorrect password.", vbCritical
            Text1.SetFocus
            Exit Sub
        End If
    Else
        If Text1.Text = sCurrUser Then
            Unload Me
        Else
            MsgBox "Incorrect password.", vbCritical
            Text1.SetFocus
            Exit Sub
        End If
    End If
End If
End Sub
