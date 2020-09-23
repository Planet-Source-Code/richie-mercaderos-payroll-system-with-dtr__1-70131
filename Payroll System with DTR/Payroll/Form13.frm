VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User's Option"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   LinkTopic       =   "Form13"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
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
      Left            =   3240
      TabIndex        =   24
      Top             =   3060
      Width           =   3165
   End
   Begin VB.Frame Frame6 
      Height          =   15
      Left            =   150
      TabIndex        =   22
      Top             =   3480
      Width           =   6255
   End
   Begin VB.TextBox Text4 
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
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   21
      Top             =   1650
      Width           =   3375
   End
   Begin VB.TextBox Text3 
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
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   20
      Top             =   1170
      Width           =   3375
   End
   Begin VB.TextBox Text2 
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
      Left            =   1380
      TabIndex        =   19
      Top             =   750
      Width           =   3375
   End
   Begin VB.OptionButton Option2 
      Caption         =   "User"
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
      Left            =   1920
      TabIndex        =   18
      Top             =   2580
      Width           =   2145
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Administrator"
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
      Left            =   1920
      TabIndex        =   17
      Top             =   2190
      Width           =   2145
   End
   Begin VB.Frame Frame5 
      Height          =   15
      Left            =   120
      TabIndex        =   15
      Top             =   2100
      Width           =   4635
   End
   Begin VB.Frame Frame4 
      Height          =   15
      Left            =   120
      TabIndex        =   13
      Top             =   1590
      Width           =   4635
   End
   Begin VB.TextBox Text1 
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
      Left            =   1380
      TabIndex        =   10
      Top             =   210
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      Height          =   15
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   4635
   End
   Begin VB.Frame Frame2 
      Height          =   15
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear Fields"
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
      Left            =   4800
      TabIndex        =   5
      Top             =   1890
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete User"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   1230
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Changes"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   150
      TabIndex        =   2
      Top             =   3000
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add This User"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   210
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   150
      TabIndex        =   0
      Top             =   3570
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "i32x32"
      SmallIcons      =   "i32x32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fullname"
         Object.Width           =   9702
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Username"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "User Type"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   90
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":19B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":268E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":3368
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":4042
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":4D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":59F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":66D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":73AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":8084
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":8D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":9A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":A712
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":B3EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":C0C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":CDA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form13.frx":DA7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "Enter a keword to search:"
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
      Left            =   180
      TabIndex        =   23
      Top             =   3090
      Width           =   3045
   End
   Begin VB.Label Label5 
      Caption         =   "User Type:"
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
      Left            =   150
      TabIndex        =   16
      Top             =   2220
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Verify:"
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
      Left            =   150
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
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
      Left            =   150
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Username:"
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
      Left            =   150
      TabIndex        =   11
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Full Name:"
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
      Left            =   150
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sAddDuplicateUser As String
Dim sEdDuplicateUser As String

Private Sub Command1_Click()
On Error Resume Next

If Form1.Label1.Caption = "User" Then MsgBox "System user is not allow to add a new user.", vbExclamation, "Unable to Add": Exit Sub

If isempty(Text1) = True Then Exit Sub
If isempty(Text2) = True Then Exit Sub
If isempty(Text3) = True Then Exit Sub

If Text4.Text <> Text3.Text Then MsgBox "Verify password correctly", vbExclamation, "Verify": Text4.SetFocus: Exit Sub
If Option1.Value = False And Option2.Value = False Then MsgBox "Select a user type.", vbExclamation, "User Type": Exit Sub

sAddDuplicateUser = sUsers.Fields(0).Value

If sAddDuplicateUser <> Text2.Text Then
    If if_exist("tblUsers", "Username", Text2) = True Then Exit Sub
End If

With sUsers
    .AddNew
    .Fields(0) = Text2.Text
    .Fields(1) = Text1.Text
    .Fields(2) = Text3.Text
    If Option1.Value = True Then
        .Fields(3) = "Administrator"
    Else
        .Fields(3) = "User"
    End If
    .Update
End With

MsgBox "Adding of new user has been successfull.", vbInformation, "Save Complete"
Call Command4_Click

ListView1.ListItems.Clear
Call LoadUsers
End Sub

Private Sub Command2_Click()
On Error Resume Next

If isempty(Text1) = True Then Exit Sub
If isempty(Text2) = True Then Exit Sub
If isempty(Text3) = True Then Exit Sub

If Text4.Text <> Text3.Text Then MsgBox "Verify password correctly", vbExclamation, "Verify": Text4.SetFocus: Exit Sub
If Option1.Value = False And Option2.Value = False Then MsgBox "Select a user type.", vbExclamation, "User Type": Exit Sub

sEdDuplicateUser = sUsers.Fields(0).Value

If sEdDuplicateUser <> Text2.Text Then
    If if_exist("tblUsers", "Username", Text2) = True Then Exit Sub
End If

With sUsers
    .Fields(0) = Text2.Text
    .Fields(1) = Text1.Text
    .Fields(2) = Text3.Text
    If Option1.Value = True Then
        .Fields(3) = "Administrator"
    Else
        .Fields(3) = "User"
    End If
    .Update
End With

MsgBox "Changes in the user has been successfull.", vbInformation, "Save Complete"

Dim pos As Integer

pos = sUsers.AbsolutePosition
ListView1.ListItems.Clear
Call LoadUsers

ListView1.ListItems.Item(pos).EnsureVisible
ListView1.ListItems.Item(pos).Selected = True

pos = 0
End Sub

Private Sub Command3_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No item on the list.", vbExclamation, "Delete User": Exit Sub
If ListView1.ListItems.Count = 1 Then
    If Form1.StatusBar1.Panels(3).Text <> "" Then
        MsgBox "Cannot delete user in use.", vbExclamation, "Delete User": Exit Sub
    End If
End If
If ListView1.SelectedItem.SubItems(2) = Form1.StatusBar1.Panels(3).Text Then MsgBox "Cannot delete user in use.", vbExclamation, "Delete User": Exit Sub


If Form1.Label1.Caption = "User" Then MsgBox "System user is not allow to delete any of the user in the list.", vbExclamation, "Delete User": Exit Sub

Dim ans As Integer

ans = MsgBox("Do you want to delete the selected user?", vbExclamation + vbOKCancel, "Delete User")

If ans = vbOK Then


    sUsers.Delete
    
    Call Command4_Click
    ListView1.ListItems.Clear
    Call LoadUsers
End If
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Option1.Value = False
Option2.Value = False

Text1.SetFocus
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
If sUsers.RecordCount < 1 Then Option2.Enabled = False
Command2.Enabled = False

Call LoadUsers
End Sub

Public Sub LoadUsers()
Dim x As ListItem

sUsers.Requery

While Not sUsers.EOF
    If sUsers.Fields(3) = "Administrator" Then
        Set x = ListView1.ListItems.Add(, , sUsers.AbsolutePosition, 1, 3)
    Else
        Set x = ListView1.ListItems.Add(, , sUsers.AbsolutePosition, 1, 1)
    End If
            x.SubItems(1) = sUsers.Fields(1)
            x.SubItems(2) = sUsers.Fields(0)
            x.SubItems(3) = sUsers.Fields(3)
        
    sUsers.MoveNext
Wend
End Sub

Private Sub ListView1_DblClick()
If ListView1.ListItems.Count < 1 Then Exit Sub

If ListView1.SelectedItem.SubItems(2) = Form1.StatusBar1.Panels(3).Text Then MsgBox "Cannot edtit user in use.", vbExclamation, "Edit": Exit Sub
If Form1.Label1.Caption = "User" Then MsgBox "System user is not allow to make any change in the selected user.", vbExclamation, "Edit": Exit Sub

Text1.Text = sUsers.Fields(1)
Text2.Text = sUsers.Fields(0)
Text3.Text = sUsers.Fields(2)
Text4.Text = sUsers.Fields(2)
If sUsers.Fields(3) = "Administrator" Then
    Option1.Value = True
Else
    Option2.Value = True
End If
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Not sUsers.RecordCount < 1 Then sUsers.AbsolutePosition = ListView1.SelectedItem
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub

Private Sub Text1_LostFocus()
Text1.Text = cSentenceCase(Text1.Text)
End Sub

Private Sub Text5_Change()
On Error Resume Next
    
    'check length
    If Len(Text5.Text) < 1 Then
        With Text5
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        Exit Sub
    End If
    
    FindLVItem ListView1, Trim(Text5.Text), , , , True ', , tmpMultiSelect, tmpInverseSelection
End Sub
