VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Loans, Taxes and Deductions - Regulars"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10620
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7350
      TabIndex        =   17
      Top             =   390
      Width           =   1125
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
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Top             =   4920
      Width           =   5415
   End
   Begin VB.Frame Frame4 
      Height          =   15
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   10335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   2520
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
            Picture         =   "Form10.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   15
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   10335
   End
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   10335
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
      Left            =   3450
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Height          =   15
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   10335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3795
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   6694
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Form10.frx":0452
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Basic"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Withholding Tax"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "GSIS Prem"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "GSIS SL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "GSIS PL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "UOLI Prem"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "UOLI Loan"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "CEAP"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ELA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "CAL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "SOS"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "PhilHealth Prem"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Pag-ibig Prem"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Pag-ibig MPL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "GOCCs LBP"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "TCRF EMPC"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "RBTC"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label2 
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
      Left            =   6540
      TabIndex        =   16
      Top             =   390
      Width           =   675
   End
   Begin MSForms.CommandButton CommandButton7 
      Height          =   540
      Left            =   5730
      TabIndex        =   15
      Top             =   5760
      Width           =   2925
      Caption         =   "Generate Payroll Report"
      Size            =   "5159;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton6 
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   4920
      Width           =   1935
      Caption         =   "Find"
      Size            =   "3413;661"
      FontName        =   "Courier New"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
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
      TabIndex        =   12
      Top             =   4920
      Width           =   2775
   End
   Begin MSForms.CommandButton CommandButton5 
      Height          =   540
      Left            =   3960
      TabIndex        =   10
      Top             =   5760
      Width           =   1665
      Caption         =   "Remove Range"
      Size            =   "2937;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton4 
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   360
      Width           =   1575
      Caption         =   "View"
      Size            =   "2778;661"
      FontName        =   "Courier New"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label17 
      Caption         =   "Date Range:"
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
      Left            =   1890
      TabIndex        =   7
      Top             =   390
      Width           =   1455
   End
   Begin MSForms.CommandButton CommandButton3 
      Height          =   540
      Left            =   8760
      TabIndex        =   3
      Top             =   5760
      Width           =   1695
      Caption         =   "Close"
      Size            =   "2990;952"
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
      Left            =   1950
      TabIndex        =   2
      Top             =   5760
      Width           =   1935
      Caption         =   "Add Deductions"
      Size            =   "3413;952"
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
      Left            =   180
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
      Caption         =   "Create New"
      Size            =   "2990;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim obj As PageSet.PrinterControl

Dim sEmploeeInList As New ADODB.Recordset
Dim sLoanRecorded As New ADODB.Recordset

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub CommandButton1_Click()
Dim a As String

If Text2.Text = "" Then MsgBox "Type a specific year.", vbExclamation, "Data Manager": Text2.SetFocus: Exit Sub
If Combo1.Text = "" Then MsgBox "Type a specific range.", vbExclamation, "Data Manager": Combo1.SetFocus: Exit Sub
If Len(Text2.Text) < 4 Then MsgBox "Invalid year.", vbExclamation, "Data Manager": Text1.SetFocus: Exit Sub

a = Combo1.Text & " " & Text2.Text

sLoanRecorded.Requery
sLoanRecorded.Find "CDate='" & a & "'"

If sLoanRecorded.EOF Then
    Call LoadEmployee
    Call CommandButton4_Click
Else
    MsgBox "Date range already created.", vbInformation, "Data Manager"
    Call CommandButton4_Click
End If
End Sub

Private Sub CommandButton2_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No record in the list please check it!", vbExclamation, "Data List": Exit Sub
Form11.add_state = False
Form11.Show vbModal
End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub CommandButton4_Click()
On Error Resume Next
Dim a

If Text2.Text = "" Then MsgBox "Type a specific year.", vbExclamation, "Data Manager": Exit Sub
If Combo1.Text = "" Then MsgBox "Type a specific range.", vbExclamation, "Data Manager": Exit Sub

a = Combo1.Text & " " & Text2.Text

sLoans.Filter = adFilterNone
sLoans.Filter = "CDate='" & a & "'"

ListView1.ListItems.Clear

Call LoadEmployeeLoans
End Sub

Private Sub CommandButton5_Click()
If sLoans.RecordCount < 1 Then MsgBox "The range you selected is empty.", vbExclamation, "Delete Range": Exit Sub


Dim ans As Integer

ans = MsgBox("Do you want to delete the selected range?", vbExclamation + vbOKCancel, "Delete Range")

If ans = vbOK Then

    sLoans.Requery
    sLoans.Filter = adFilterNone
    sLoans.Filter = "CDate='" & Combo1.Text & " " & Text2.Text & "'"
    
    While Not sLoans.EOF
        sLoans.Delete
        sLoans.MoveNext
    Wend
    
    Call LoadEmployeeLoans
    MsgBox "The selected ranged has been deleted.", vbInformation, "Delete Complete"
    ListView1.ListItems.Clear
End If

End Sub

Private Sub CommandButton7_Click()
On Error GoTo errorhandler:

If ListView1.ListItems.Count < 1 Then MsgBox "No recod in the list please check it!", vbExclamation, "List View": Exit Sub
PeriodNow = ""

sLoans.Requery
sLoans.Filter = "CDate='" & Combo1.Text & " " & Text2.Text & "'"

PeriodNow = "Period of Service (Inclusive Date) " & Combo1.Text & " " & Text2.Text
        
        Set obj = New PrinterControl
        obj.ChngOrientationLandscape

        Set DataReport2.DataSource = sLoans
        Set DataReport3.DataSource = sLoans

        DataReport2.Show vbModal
        DataReport3.Show vbModal
        
        Exit Sub

errorhandler:
       MsgBox Err.Description
       obj.ReSetOrientation
 
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim a As String

'If Format(Date, "mmmm") = "February" Then
'    Combo1.Text = Format(Date, "mmmm") & " 1-28,"

'Text2.Text = Year(Now)

'a = Combo1.Text & " " & Text2.Text

Call set_rec_getData(sEmploeeInList, cn, "Select * From qryEmployeeInfo where Designation<>'Admin Aide I (JO)' and Designation<>'Admin Aide I (WAGES)'")
Call set_rec_getData(sLoans, cn, "Select * From qryLoans")
Call set_rec_getData(sLoanRecorded, cn, "Select * From qryLoans")


Call LoadRanges
Call TraceDate
Text2.Text = Year(Now)
'sLoans.Requery
'sLoans.Find "CDate='" & a & "'"

'If sLoans.EOF Then
'    Call LoadEmployee
'    Call LoadEmployeeLoans
'Else
'    Call LoadEmployeeLoans
'End If
ListView1.ListItems.Item(1).EnsureVisible
ListView1.ListItems.Item(1).Selected = True
End Sub

Public Sub LoadEmployeeLoans()
On Error Resume Next
Dim x As ListItem

sLoans.Requery

ListView1.ListItems.Clear

While Not sLoans.EOF
    Set x = ListView1.ListItems.Add(, , sLoans.AbsolutePosition, 1, 1)
        x.SubItems(1) = sLoans.Fields("CDate")
        x.SubItems(2) = sLoans.Fields("Name")
        x.SubItems(3) = sLoans.Fields("Basic")
        x.SubItems(4) = sLoans.Fields("WHoldingTax")
        x.SubItems(5) = sLoans.Fields("GSISPrem")
        x.SubItems(6) = sLoans.Fields("GSISSL")
        x.SubItems(7) = sLoans.Fields("GSISPL")
        x.SubItems(8) = sLoans.Fields("UOLIPrem")
        x.SubItems(9) = sLoans.Fields("UOLILoan")
        x.SubItems(10) = sLoans.Fields("CEAP")
        x.SubItems(11) = sLoans.Fields("ELA")
        x.SubItems(12) = sLoans.Fields("CAL")
        x.SubItems(13) = sLoans.Fields("SOS")
        x.SubItems(14) = sLoans.Fields("PhilHealthPrem")
        x.SubItems(15) = sLoans.Fields("PAGIBIGPrem")
        x.SubItems(16) = sLoans.Fields("PAGIBIGMPL")
        x.SubItems(17) = sLoans.Fields("GOCCsLBP")
        x.SubItems(18) = sLoans.Fields("TCRFEMPC")
        x.SubItems(19) = sLoans.Fields("RBTC")
    sLoans.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set sLoans = Nothing
Set sEmploeeInList = Nothing
Set sLoanRecorded = Nothing
End Sub

Private Sub ListView1_DblClick()
Call CommandButton2_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Not sLoans.RecordCount < 1 Then sLoans.AbsolutePosition = ListView1.SelectedItem
End Sub

Public Sub LoadEmployee()
Dim x As ListItem
Dim i As Integer

i = 0

sEmploeeInList.Requery
ListView1.ListItems.Clear

While Not sEmploeeInList.EOF
    i = i + 1
    sLoans.AddNew
        sLoans.Fields("Count") = i
        sLoans.Fields("Basic1") = sEmploeeInList.Fields("SalaryBasic")
        sLoans.Fields("Basic") = sEmploeeInList.Fields("SalaryBasic")
        sLoans.Fields("Half") = FormatNumber(Val(sLoans.Fields("Basic1")) / 2, 2)
        sLoans.Fields("CDate") = Combo1.Text & " " & Text2.Text
        sLoans.Fields("Name") = sEmploeeInList.Fields("Name")
        sLoans.Fields("NetAmount") = sEmploeeInList.Fields("SalaryBasic")
        sLoans.Fields("EmployeeNo") = sEmploeeInList.Fields("EmployeeNo")
    sLoans.Update
    
    sEmploeeInList.MoveNext
Wend
End Sub

Private Sub Text1_Change()
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

Public Sub LoadRanges()
With Combo1
    .AddItem "January 1-31,"
    .AddItem "February 1-28,"
    .AddItem "March 1-31,"
    .AddItem "April 1-30,"
    .AddItem "May 1-31,"
    .AddItem "June 1-30,"
    .AddItem "July 1-31,"
    .AddItem "August 1-31,"
    .AddItem "September 1-30,"
    .AddItem "October 1-31,"
    .AddItem "November 1-30,"
    .AddItem "December 1-31,"
End With
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Public Sub TraceDate()
Select Case (Format(Date, "m"))
    Case 1: Combo1.Text = "January 1-31,"
    Case 2: Combo1.Text = "February 1-28,"
    Case 3: Combo1.Text = "March 1-31,"
    Case 4: Combo1.Text = "April 1-30,"
    Case 5: Combo1.Text = "May 1-31,"
    Case 6: Combo1.Text = "June 1-30,"
    Case 7: Combo1.Text = "July 1-31,"
    Case 8: Combo1.Text = "August 1-31,"
    Case 9: Combo1.Text = "September 1-30,"
    Case 10: Combo1.Text = "October 1-31,"
    Case 11: Combo1.Text = "November 1-30,"
    Case 12: Combo1.Text = "December 1-31,"
End Select
End Sub

