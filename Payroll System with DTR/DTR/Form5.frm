VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   ScaleHeight     =   990
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   ">"
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autorized Personnel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Enter Code:"
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
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim searchlog As New ADODB.Recordset
Public useofthis As Boolean

Dim sType As String
Dim tUser As String

Private Sub Command1_Click()
On Error Resume Next
searchlog.Requery

searchlog.Find "Password='" & Text1.Text & "'"

sType = searchlog.Fields("UserType")
tUser = searchlog.Fields("Username")

If searchlog.EOF Then
    MsgBox "Your not authorized to do this action.", vbExclamation, "Administrator"
    Exit Sub
Else
    Unload Me
    If useofthis = True Then
        If sType <> "Admin" Then
            MsgBox "Sorry, administrator only.", vbExclamation, "System Required"
        Else
            Form10.Show vbModal
        End If
    Else
        Form9.sUser = tUser
        Form9.Show vbModal
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Call set_rec_getData(searchlog, cn, "Select * From tblUsers")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set searchlog = Nothing
End Sub
