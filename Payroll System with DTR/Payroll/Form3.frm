VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
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
      IMEMode         =   3  'DISABLE
      Left            =   4110
      PasswordChar    =   "*"
      TabIndex        =   24
      Top             =   4770
      Width           =   2850
   End
   Begin VB.Frame Frame3 
      Height          =   15
      Left            =   90
      TabIndex        =   23
      Top             =   4650
      Width           =   6945
   End
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
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   4110
      PasswordChar    =   "*"
      TabIndex        =   21
      Top             =   4170
      Width           =   2850
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
      Left            =   2355
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   3720
      Width           =   4590
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   420
      Top             =   2730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   15
      Left            =   105
      TabIndex        =   16
      Top             =   210
      Width           =   6945
   End
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   105
      TabIndex        =   15
      Top             =   5250
      Width           =   6945
   End
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
      Left            =   4095
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   2760
      Width           =   2850
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
      ItemData        =   "Form3.frx":0000
      Left            =   4095
      List            =   "Form3.frx":000A
      TabIndex        =   13
      Top             =   2280
      Width           =   2850
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
      Height          =   360
      Left            =   4095
      TabIndex        =   6
      Top             =   525
      Width           =   2850
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
      Height          =   360
      Left            =   4095
      TabIndex        =   5
      Top             =   1050
      Width           =   2850
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
      Height          =   360
      Left            =   4095
      TabIndex        =   4
      Top             =   1470
      Width           =   2850
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
      Height          =   360
      Left            =   4095
      TabIndex        =   3
      Top             =   1890
      Width           =   2850
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Look for..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   2
      ToolTipText     =   "Click to look for an appropriate photo."
      Top             =   1995
      Width           =   1485
   End
   Begin VB.Label Label11 
      Caption         =   "Verify Code:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2430
      TabIndex        =   25
      Top             =   4830
      Width           =   1590
   End
   Begin VB.Label Label10 
      Caption         =   "Code:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2430
      TabIndex        =   22
      Top             =   4245
      Width           =   1590
   End
   Begin VB.Label Label9 
      Caption         =   "Project:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1290
      TabIndex        =   20
      Top             =   3780
      Width           =   1020
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4080
      TabIndex        =   18
      Top             =   3240
      Width           =   2865
   End
   Begin VB.Label Label7 
      Caption         =   "Basic Salary:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2400
      TabIndex        =   17
      Top             =   3300
      Width           =   1590
   End
   Begin VB.Label Label6 
      Caption         =   "Designation:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2415
      TabIndex        =   12
      Top             =   2835
      Width           =   1590
   End
   Begin VB.Label Label5 
      Caption         =   "Sex:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2415
      TabIndex        =   11
      Top             =   2340
      Width           =   1590
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1380
      Left            =   210
      Stretch         =   -1  'True
      Top             =   525
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Employee No.:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2415
      TabIndex        =   10
      Top             =   600
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Surname:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2415
      TabIndex        =   9
      Top             =   1140
      Width           =   1590
   End
   Begin VB.Label Label3 
      Caption         =   "Firstname:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2415
      TabIndex        =   8
      Top             =   1560
      Width           =   1590
   End
   Begin VB.Label Label4 
      Caption         =   "Middlename:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2415
      TabIndex        =   7
      Top             =   1980
      Width           =   1590
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   540
      Left            =   3600
      TabIndex        =   1
      Top             =   5490
      Width           =   1650
      Caption         =   "Cancel"
      Size            =   "2910;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   540
      Left            =   5355
      TabIndex        =   0
      Top             =   5490
      Width           =   1650
      Caption         =   "Save"
      Size            =   "2910;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sDesignation    As New ADODB.Recordset
'--------------------------------------
Dim sProject        As New ADODB.Recordset
'--------------------------------------

Public add_state    As Boolean
Public sDuplicate   As String

Private Sub Combo2_Click()
sDesignation.MoveFirst
sDesignation.Find "Designation='" & Combo2.Text & "'"

Label8.Caption = FormatNumber(sDesignation.Fields("SalaryperDay"), 2)

If Combo2.Text = "Admin Aide I (JO)" Then
sProject.MoveFirst
sProject.Find "Designation='" & Combo2.Text & "'"

Combo3.Clear
While Not sProject.EOF
    Combo3.AddItem sProject.Fields("Project")
    sProject.MoveNext
Wend
Else
    Combo3.Clear
End If

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command3_Click()
cd.Filter = "Picuter File(*.jpg;*.bmp)|*.jpg;*.bmp"
cd.ShowOpen

Image1.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub CommandButton1_Click()
If isempty(Text1) = True Then Exit Sub
If isempty(Text2) = True Then Exit Sub
If isempty(Text3) = True Then Exit Sub
If isempty(Text4) = True Then Exit Sub
If isempty(Combo1) = True Then Exit Sub
If isempty(Combo2) = True Then Exit Sub

If Combo2.Text = "Admin Aide I (JO)" And Combo3.Text = "" Then MsgBox "Select project title.", vbExclamation, "Data Manager": Combo3.SetFocus: Exit Sub

If Text5.Text <> Text6.Text Then MsgBox "Verify code correctly.", vbExclamation, "Data Manager": Text6.SetFocus: Exit Sub

If sDuplicate <> Text1.Text Then
    If if_exist("tblEmployeeInfo", "EmployeeNo", Text1) = True Then Exit Sub
End If

With rs
    If add_state = True Then .AddNew
        .Fields("EmployeeNo") = Text1.Text
        .Fields("Surname") = Text2.Text
        .Fields("Firstname") = Text3.Text
        .Fields("Middlename") = Text4.Text
        .Fields("Sex") = Combo1.Text
        .Fields("Designation") = Combo2.Text
        .Fields("Project") = Combo3.Text
        .Fields("Name") = Text2.Text & ", " & Text3.Text & " " & Text4.Text
        .Fields("SalaryBasic") = Label8.Caption
        .Fields("Picture") = cd.FileName
        .Fields("Code") = Text5.Text
        .Update
End With

If add_state = True Then
    MsgBox "Adding of new employee has been successfull.", vbInformation, "Save Complete"
    Dim rep As Integer
    rep = MsgBox("Do you want to add another employee?", vbQuestion + vbYesNo, "Record Manager")
    If rep = vbYes Then
            
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Combo1.Text = ""
        Combo2.Text = ""
        Label8.Caption = ""
        Text5.Text = ""
        Text6.Text = ""
        Image1.Picture = LoadPicture(App.Path & "\Pictures\Photo.jpg")
        cd.FileName = App.Path & "\Pictures\Photo.jpg"
        
        Text1.SetFocus
        
        Form2.ListView1.ListItems.Clear
        Form2.LoadEmployeeInfo
    Else
        Form2.ListView1.ListItems.Clear
        Form2.LoadEmployeeInfo
        Unload Me
    End If
    rep = 0
Else
    MsgBox "Changes in the employee record has been successfully saved.", vbInformation, "Save Complete"
    Dim pos As Long
    
    pos = rs.AbsolutePosition
    
    Form2.ListView1.ListItems.Clear
    Form2.LoadEmployeeInfo
    
    Form2.ListView1.ListItems.Item(pos).EnsureVisible
    Form2.ListView1.ListItems.Item(pos).Selected = True
    
    pos = 0
    Unload Me
End If
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Call set_rec_getData(sDesignation, cn, "Select * From tblDesignation")
Call set_rec_getData(sProject, cn, "Select * From tblProject")

While Not sDesignation.EOF
    Combo2.AddItem sDesignation.Fields("Designation")
    sDesignation.MoveNext
Wend

Call LoadProject

If add_state = True Then
    Me.Caption = "Add Employee Time Record"
    Image1.Picture = LoadPicture(App.Path & "\Pictures\Photo.jpg")
    cd.FileName = App.Path & "\Pictures\Photo.jpg"
Else
    Me.Caption = "Edit Employee Time Record"
    With rs
        
        Text1.Text = .Fields("EmployeeNo")
        Text2.Text = .Fields("Surname")
        Text3.Text = .Fields("Firstname")
        Text4.Text = .Fields("Middlename")
        Text5.Text = .Fields("Code")
        Combo1.Text = .Fields("Sex")
        Combo2.Text = .Fields("Designation")
        Combo3.Text = .Fields("Project")
        Label8.Caption = FormatNumber(.Fields("SalaryBasic"), 2)
        Image1.Picture = LoadPicture(.Fields("Picture"))
        cd.FileName = .Fields("Picture")
    End With
        Text6.Text = Text5.Text
End If
sDuplicate = rs.Fields("EmployeeNo").Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set sDesignation = Nothing
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text1_LostFocus()
Text1.Text = cSentenceCase(Text1.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text2_LostFocus()
Text2.Text = cSentenceCase(Text2.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus
End Sub

Private Sub Text3_LostFocus()
Text3.Text = cSentenceCase(Text3.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Combo1.SetFocus
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text4_LostFocus()
Text4.Text = cSentenceCase(Text4.Text)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text6.SetFocus
End Sub

Public Sub LoadProject()
While Not sProject.EOF
    Combo3.AddItem sProject.Fields("Project")
    sProject.MoveNext
Wend
End Sub
