VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Employee Information"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9885
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lists of Employees"
      TabPicture(0)   =   "Form2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ImageList1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ImageList2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   945
         Top             =   1890
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
               Picture         =   "Form2.frx":001C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   15
         Left            =   240
         TabIndex        =   6
         Top             =   5460
         Width           =   7095
      End
      Begin VB.Frame Frame1 
         Height          =   15
         Left            =   240
         TabIndex        =   5
         Top             =   4920
         Width           =   9135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
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
         Left            =   7560
         TabIndex        =   4
         Top             =   5040
         Width           =   1815
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
         Left            =   2880
         TabIndex        =   3
         Top             =   5040
         Width           =   4455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4335
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7646
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Full Names"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Sex"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Designation"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
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
         Left            =   240
         TabIndex        =   2
         Top             =   5040
         Width           =   2535
      End
   End
   Begin MSForms.CommandButton CommandButton4 
      Height          =   540
      Left            =   5985
      TabIndex        =   10
      Top             =   6090
      Width           =   2010
      Caption         =   "Print List (F4)"
      Size            =   "3545;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton3 
      Height          =   540
      Left            =   4095
      TabIndex        =   9
      Top             =   6090
      Width           =   1800
      Caption         =   "Edit (F3)"
      Size            =   "3175;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   540
      Left            =   2205
      TabIndex        =   8
      Top             =   6090
      Width           =   1800
      Caption         =   "Add New (F2)"
      Size            =   "3175;952"
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
      Left            =   8295
      TabIndex        =   7
      Top             =   6090
      Width           =   1485
      Caption         =   "Close"
      Size            =   "2619;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim obj As PageSet.PrinterControl

Public Sub LoadEmployeeInfo()
Dim x As ListItem

rs.Requery

While Not rs.EOF
    Set x = ListView1.ListItems.Add(, , rs.AbsolutePosition, 1, 1)
        x.SubItems(1) = rs.Fields("EmployeeNo")
        x.SubItems(2) = rs.Fields("Name")
        x.SubItems(3) = rs.Fields("Sex")
        x.SubItems(4) = rs.Fields("Designation")
        
    rs.MoveNext
Wend
End Sub

Private Sub Command1_Click()
On Error Resume Next
    
    'check length
    If Len(Text1.Text) < 1 Then
        With Text1
            .SelStart = 0
            .SelLength = Len(Text1.Text)
        End With
        Exit Sub
    End If
    
    FindLVItem ListView1, Trim(Text1.Text), , , , True ', , tmpMultiSelect, tmpInverseSelection
End Sub

Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
Form3.add_state = True
Form3.Show vbModal
End Sub

Private Sub CommandButton3_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No record in the list. Please check it!", vbExclamation, "Record List": Exit Sub
Form3.add_state = False
Form3.Show vbModal
End Sub

Private Sub CommandButton4_Click()
On Error GoTo errorhandler:
      
      Set obj = New PrinterControl
      obj.ChngOrientationPortrait
      Set DataReport1.DataSource = rs
      DataReport1.Show vbModal

      Exit Sub
      
errorhandler:
       MsgBox Err.Description
       obj.ReSetOrientation
End Sub

Private Sub Form_Load()
On Error Resume Next
Call LoadEmployeeInfo
ListView1.ListItems.Item(1).Selected = True
rs.AbsolutePosition = ListView1.SelectedItem
End Sub

Private Sub ListView1_DblClick()
Call CommandButton3_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Not rs.RecordCount < 1 Then rs.AbsolutePosition = ListView1.SelectedItem
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2: Call CommandButton2_Click
    Case vbKeyF3: Call CommandButton3_Click
    Case vbKeyF4: Call CommandButton4_Click
    Case vbKeyEscape: Call CommandButton1_Click
End Select
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2: Call CommandButton2_Click
    Case vbKeyF3: Call CommandButton3_Click
    Case vbKeyF4: Call CommandButton4_Click
    Case vbKeyEscape: Call CommandButton1_Click
End Select
End Sub

Private Sub Text1_Change()
Call Command1_Click
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2: Call CommandButton2_Click
    Case vbKeyF3: Call CommandButton3_Click
    Case vbKeyF4: Call CommandButton4_Click
    Case vbKeyEscape: Call CommandButton1_Click
End Select
End Sub
