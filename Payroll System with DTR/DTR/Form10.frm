VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Employee DTR"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   6120
   End
   Begin VB.Frame Frame4 
      Caption         =   "Date Specification"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   8295
      Begin VB.ComboBox Text2 
         Height          =   315
         ItemData        =   "Form10.frx":0000
         Left            =   5880
         List            =   "Form10.frx":0013
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Text4 
         Height          =   315
         ItemData        =   "Form10.frx":002B
         Left            =   4440
         List            =   "Form10.frx":0035
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7440
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Month:"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin MSForms.ComboBox ComboBox2 
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         VariousPropertyBits=   209733659
         DisplayStyle    =   3
         Size            =   "3201;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label5 
         Caption         =   "From:"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "To:"
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Year:"
         Height          =   255
         Left            =   6960
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
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
      Left            =   6840
      TabIndex        =   10
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
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
      Left            =   5040
      TabIndex        =   9
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "General "
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8295
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   2760
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View DTR"
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
         Left            =   6720
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   6495
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
            Height          =   285
            Left            =   2760
            TabIndex        =   6
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Enter text to search:"
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
            Top             =   240
            Width           =   2535
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee No."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Surname"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Firstname"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Middlename"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Sex"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Designation"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   1200
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form10.frx":0040
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Transferring. . ."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   8295
      Begin MSComctlLib.ListView ListView1 
         Height          =   1545
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2725
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   15269887
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New TUR"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   549
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "TimeInAM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "TimeOutAM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Hours & Mins (Undertime)"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "TimeInPM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "TimeOutPM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Hours & Mins (Undertime)"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Total Undertime (Hours & Mins)"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   360
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form10.frx":0C12
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_emp As New ADODB.Recordset

Public Sub Months()
ComboBox2.Text = Format(Date, "mmmm")
With ComboBox2
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

Private Sub Command1_Click()

If Text4.Text = "" Then
    MsgBox "Specify a range.", vbExclamation, "System Required"
    Text4.SetFocus
    Exit Sub
End If

If Text2.Text = "" Then
    MsgBox "Specify a range.", vbExclamation, "System Required"
    Text2.SetFocus
    Exit Sub
End If

If Val(Text2.Text) < Val(Text4.Text) Then
    MsgBox "Invalid range.", vbExclamation, "System Required"
    Exit Sub
    Text4.SetFocus
End If

If Val(Text4.Text) > 31 Then
    MsgBox "Invalid input value.", vbExclamation, "System Required"
    Text4.SetFocus
    Exit Sub
End If

If Val(Text2.Text) > 31 Then
    MsgBox "Invalid input value.", vbExclamation, "System Required"
    Text2.SetFocus
    Exit Sub
End If

ListView1.ListItems.Clear
Timer1.Enabled = True
Label2.Visible = True
End Sub

Private Sub Command3_Click()
Form6.add_state = True
Form6.Show vbModal
End Sub

Private Sub Command4_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No record in the list. Please check it!", vbExclamation, "Daily Time Record": Exit Sub
Form6.add_state = False
Form6.Show vbModal
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call centerForm(Me, Screen.Height, Screen.Width)

Call set_rec_getData(rs_emp, cn, "Select * From tblEmployeeInfo")
'Call set_rec_getData(time_rec, cn, "Select * From qryComputed")

Call Months
Call LoadEmployeeInfo

Text3.Text = Year(Now)
ComboBox2.Text = Format(Date, "mmmm")
End Sub

Public Sub LoadEmployeeInfo()
Dim X As ListItem

While Not rs_emp.EOF
    
    Set X = ListView2.ListItems.Add(, , rs_emp.AbsolutePosition, 1, 1)
        X.SubItems(1) = rs_emp.Fields("EmployeeNo")
        X.SubItems(2) = rs_emp.Fields("Surname")
        X.SubItems(3) = rs_emp.Fields("Firstname")
        X.SubItems(4) = rs_emp.Fields("Middlename")
        X.SubItems(5) = rs_emp.Fields("Sex")
        X.SubItems(6) = rs_emp.Fields("Designation")
        
    rs_emp.MoveNext

Wend
End Sub

Public Sub LoadRecord()
On Error Resume Next
Dim X As ListItem

While Not time_rec.EOF
    Set X = ListView1.ListItems.Add(, , time_rec.AbsolutePosition, 1, 1)

    X.SubItems(1) = time_rec.Fields("CDate")
    X.SubItems(2) = time_rec.Fields("AMIn")
    X.SubItems(3) = time_rec.Fields("AMOut")
    X.SubItems(4) = time_rec.Fields("ATHour") & " hr(s) " & time_rec.Fields("ATMin") & " min(s)"
    X.SubItems(5) = time_rec.Fields("PMIn")
    X.SubItems(6) = time_rec.Fields("PMOut")
    X.SubItems(7) = time_rec.Fields("PTHour") & " hr(s) " & time_rec.Fields("PTMin") & " min(s)"
    X.SubItems(8) = time_rec.Fields("TTinHour") & " hr(s) " & time_rec.Fields("TTinMin") & " min(s)"
    
    time_rec.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rs_emp = Nothing
'Set load_rec = Nothing
'Set time_rec = Nothing
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Not time_rec.RecordCount < 1 Then time_rec.AbsolutePosition = ListView1.SelectedItem
'MsgBox time_rec.AbsolutePosition
End Sub

Private Sub Text1_Change()
Dim X As ListItem

Set X = ListView2.FindItem(Trim(Text1.Text), lvwSubItem + lvwText, lvwPartial, lvwPartial)
If Not X Is Nothing Then
    X.EnsureVisible
    X.Selected = True
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
pb.Value = pb.Value + 5

If pb.Value > 99 Then

    time_rec.Requery
    
    time_rec.Filter = adFilterNone
    time_rec.Filter = "EmployeeNo='" & ListView2.SelectedItem.SubItems(1) & "' and Month='" & ComboBox2.Text & "' and Day>=" & Text4.Text & " and Day<=" & Text2.Text & " and Year='" & Text3.Text & "'"
    
    Call LoadRecord

    Label2.Visible = False
    Timer1.Enabled = False
    pb.Value = 0

End If
End Sub
