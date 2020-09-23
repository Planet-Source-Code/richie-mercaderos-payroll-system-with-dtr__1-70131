VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form18 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Payroll - Non-Regular (Wages)"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   LinkTopic       =   "Form18"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "Form18.frx":0000
      Left            =   840
      List            =   "Form18.frx":0002
      TabIndex        =   23
      Top             =   60
      Width           =   2430
   End
   Begin VB.Frame Frame5 
      Height          =   15
      Left            =   120
      TabIndex        =   9
      Top             =   510
      Width           =   10560
   End
   Begin VB.Frame Frame4 
      Height          =   15
      Left            =   120
      TabIndex        =   8
      Top             =   1110
      Width           =   10560
   End
   Begin VB.Frame Frame3 
      Height          =   15
      Left            =   90
      TabIndex        =   7
      Top             =   5970
      Width           =   10620
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
      Height          =   330
      Left            =   2295
      TabIndex        =   6
      Top             =   5280
      Width           =   6555
   End
   Begin VB.Frame Frame2 
      Height          =   15
      Left            =   90
      TabIndex        =   5
      Top             =   5760
      Width           =   8835
   End
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   90
      TabIndex        =   4
      Top             =   5130
      Width           =   8835
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
      ItemData        =   "Form18.frx":0004
      Left            =   3030
      List            =   "Form18.frx":0006
      TabIndex        =   3
      Top             =   600
      Width           =   2430
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
      Left            =   8340
      TabIndex        =   2
      Top             =   600
      Width           =   945
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
      ItemData        =   "Form18.frx":0008
      Left            =   5580
      List            =   "Form18.frx":0012
      TabIndex        =   1
      Top             =   600
      Width           =   810
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
      ItemData        =   "Form18.frx":001D
      Left            =   6810
      List            =   "Form18.frx":0030
      TabIndex        =   0
      Top             =   600
      Width           =   810
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3795
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Double click to add deductions"
      Top             =   1230
      Width           =   10605
      _ExtentX        =   18706
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
      MouseIcon       =   "Form18.frx":0048
      NumItems        =   22
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Names"
         Object.Width           =   7057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Designation"
         Object.Width           =   4676
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Rate per Day"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "No. of Days Worked"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Gross Salary"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "PERA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ADCOM"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Life & Retirement"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Medicare"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "W/Holding Tax"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "GSIS Salary Loan"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "GSIS Policy Loan"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "ELA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "CAL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "SOS"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Consol Loan"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Pag-ibig Premium"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "TCRF EMPC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "RBTC Loan"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "UnderTime"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Amount Received"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2220
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
            Picture         =   "Form18.frx":0362
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Term:"
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
      Left            =   150
      TabIndex        =   22
      Top             =   120
      Width           =   645
   End
   Begin MSForms.CommandButton CommandButton4 
      Height          =   540
      Left            =   9135
      TabIndex        =   21
      Top             =   6150
      Width           =   1590
      Caption         =   "Close"
      Size            =   "2805;952"
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
      Left            =   6330
      TabIndex        =   20
      Top             =   6150
      Width           =   2640
      Caption         =   "Generate Payroll"
      Size            =   "4657;952"
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
      Left            =   9120
      TabIndex        =   19
      Top             =   5190
      Width           =   1590
      Caption         =   "Find"
      Size            =   "2805;952"
      FontName        =   "Courier New"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label4 
      Caption         =   "Enter text here:"
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
      Left            =   195
      TabIndex        =   18
      Top             =   5310
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Date Ranged:"
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
      Left            =   1350
      TabIndex        =   17
      Top             =   660
      Width           =   1485
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   540
      Left            =   120
      TabIndex        =   16
      Top             =   6150
      Width           =   1740
      Caption         =   "Create New"
      Size            =   "3069;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton5 
      Height          =   540
      Left            =   1920
      TabIndex        =   15
      Top             =   6150
      Width           =   2040
      Caption         =   "Add Deductions"
      Size            =   "3598;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
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
      Height          =   225
      Left            =   7710
      TabIndex        =   14
      Top             =   660
      Width           =   555
   End
   Begin MSForms.CommandButton CommandButton6 
      Height          =   480
      Left            =   9420
      TabIndex        =   13
      Top             =   600
      Width           =   1215
      Caption         =   "View"
      Size            =   "2143;847"
      FontName        =   "Courier New"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label5 
      Caption         =   "to"
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
      Left            =   6450
      TabIndex        =   12
      Top             =   660
      Width           =   255
   End
   Begin MSForms.CommandButton CommandButton7 
      Height          =   540
      Left            =   4050
      TabIndex        =   11
      Top             =   6150
      Width           =   2040
      Caption         =   "Remove Range"
      Size            =   "3598;952"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim obj As PageSet.PrinterControl
Dim obj1 As PageSet.PrinterControl

Dim sProject            As New ADODB.Recordset
Dim sNumDayWork         As New ADODB.Recordset
Dim sCreated            As New ADODB.Recordset
Dim sAllowance          As New ADODB.Recordset
Dim sLifeRet            As New ADODB.Recordset
Dim sMedicare           As New ADODB.Recordset

Dim sTotalNumberDays    As Double
Dim sTotalDeduction     As Double

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_Change()
Call Combo3_Click
End Sub

Private Sub Combo3_Click()
If Combo3.Text = 1 Then
    Combo4.Text = 15
    Combo2.Text = "1st Term"
Else
    Combo4.Text = 31
    Combo2.Text = "2nd Term"
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub CommandButton2_Click()
Dim a As String

If Len(Text1.Text) < 4 Then MsgBox "Invalid year.", vbExclamation, "Data Manager": Text1.SetFocus: Exit Sub

a = Combo1.Text & " " & Combo3.Text & "-" & Combo4.Text & ", " & Text1.Text

sCreated.Requery
sCreated.Find "CDate='" & a & "'"
    
    sCreated.Find "Term='" & Combo2.Text & "'"

        If sCreated.EOF Then
            Call LoadEmployeeData
            Call CommandButton6_Click
        Else
            MsgBox "Date range already created.", vbInformation, "Data Manager"
            ListView1.ListItems.Clear
            Call CommandButton6_Click
        End If
End Sub

Private Sub CommandButton3_Click()
On Error GoTo errorhandler:

CasualPay.Requery
CasualPay.Filter = adFilterNone
CasualPay.Filter = "Term='" & Combo2.Text & "' and CDate='" & Combo1.Text & " " & Combo3.Text & "-" & Combo4.Text & ", " & Text1.Text & "'"

sPeriodUse = Combo1.Text & " " & Combo3.Text & "-" & Combo4.Text & ", " & Text1.Text

If Combo2.Text = "1st Term" Then
    Set DataReport7.DataSource = CasualPay
    DataReport7.Show vbModal
Else
    Set DataReport8.DataSource = CasualPay
    DataReport8.Show vbModal
End If

Exit Sub

errorhandler:
       MsgBox Err.Description, vbExclamation
       'obj.ReSetOrientation
       'obj1.ReSetOrientation
End Sub

Private Sub CommandButton4_Click()
Unload Me
End Sub

Private Sub CommandButton5_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No record in the list.", vbExclamation, "Data Manager": Exit Sub

If Combo2.Text = "1st Term" Then Form19.Text10.Enabled = False
Form19.Show vbModal
End Sub

Private Sub CommandButton6_Click()
If Text1.Text = "" Then MsgBox "Enter a specific year.", vbExclamation, "Data Manager": Text1.SetFocus: Exit Sub
If Combo1.Text = "" Then MsgBox "Enter a specific Month.", vbExclamation, "Data Manager": Combo1.SetFocus: Exit Sub
If Combo2.Text = "" Then MsgBox "Enter a specific term.", vbExclamation, "Data Manager": Combo2.SetFocus: Exit Sub
If Combo3.Text = "" Then MsgBox "Enter a specific Date.", vbExclamation, "Data Manager": Combo3.SetFocus: Exit Sub
If Combo4.Text = "" Then MsgBox "Enter a specific Date.", vbExclamation, "Data Manager": Combo4.SetFocus: Exit Sub

CasualPay.Requery
CasualPay.Filter = adFilterNone
CasualPay.Filter = "Term='" & Combo2.Text & "' and CDate='" & Combo1.Text & " " & Combo3.Text & "-" & Combo4.Text & ", " & Text1.Text & "'"

ListView1.ListItems.Clear
Call LoadCasualPay
End Sub

Private Sub CommandButton7_Click()

If CasualPay.RecordCount < 1 Then MsgBox "The range you selected is empty.", vbExclamation, "Delete Range": Exit Sub


Dim ans As Integer

ans = MsgBox("Do you want to delete the selected range?", vbExclamation + vbOKCancel, "Delete Range")

If ans = vbOK Then

    CasualPay.Requery
    CasualPay.Filter = adFilterNone
    CasualPay.Filter = "Term='" & Combo2.Text & "' and CDate='" & Combo1.Text & " " & Combo3.Text & "-" & Combo4.Text & ", " & Text1.Text & "'"
    
    While Not CasualPay.EOF
        CasualPay.Delete
        CasualPay.MoveNext
    Wend
    
    Call LoadCasualPay
    MsgBox "The selected ranged has been deleted.", vbInformation, "Delete Complete"
    ListView1.ListItems.Clear
End If
End Sub

Private Sub Form_Load()
Call set_rec_getData(sGeneratePayrollJO, cn, "Select * From qryEmployeeInfo where Designation='Admin Aide I (WAGES)'")
Call set_rec_getData(CasualPay, cn, "Select * From tblLoanWages")
Call set_rec_getData(sNumDayWork, cn, "Select * From qryComputed")
Call set_rec_getData(sCreated, cn, "Select * From tblLoanWages")
Call set_rec_getData(sLifeRet, cn, "Select * From tblLifeRetirement")
Call set_rec_getData(sAllowance, cn, "Select * From tblAllowances")
Call set_rec_getData(sMedicare, cn, "Select * From tblMedicare")

Call LoadMonth
Combo1.Text = Format(Date, "mmmm")
Text1.Text = Year(Now)
End Sub

Public Sub LoadEmployeeData()
Dim i As Integer

i = 0
sAllowance.MoveFirst
sLifeRet.MoveFirst
sMedicare.MoveFirst

'sGeneratePayrollJO.Requery

While Not sGeneratePayrollJO.EOF
    i = i + 1
    
    With CasualPay
        .AddNew
        .Fields("Count") = i
        .Fields("CDate") = Combo1.Text & " " & Combo3.Text & "-" & Combo4.Text & ", " & Text1.Text
        
        .Fields("Name") = sGeneratePayrollJO.Fields("Name")
        
        
        sNumDayWork.Requery
        sNumDayWork.Filter = adFilterNone

        sNumDayWork.Filter = "Name='" & sGeneratePayrollJO.Fields("Name") & "' and Day>=" & Combo3.Text & " and Day<=" & Combo4.Text & " and Month='" & Combo1.Text & "' and Year='" & Text1.Text & "'"
        
        sTotalNumberDays = 0
        sTotalDeduction = 0
        
        While Not sNumDayWork.EOF
            sTotalNumberDays = sTotalNumberDays + Val(sNumDayWork.Fields("NumDayWork"))
            sTotalDeduction = sTotalDeduction + Val(sNumDayWork.Fields("Deduction"))
            sNumDayWork.MoveNext
        Wend
        
            If sTotalNumberDays < 11 Then
                .Fields("PERA") = FormatNumber(sTotalNumberDays * 22.72, 2)
                .Fields("ADCOM") = FormatNumber(sTotalNumberDays * 68.18, 2)
                .Fields("LifeRet") = FormatNumber(sLifeRet.Fields("LifeRetirement"), 2)
                .Fields("AmountRecived") = FormatNumber(((sTotalNumberDays * Val(sGeneratePayrollJO.Fields("SalaryBasic"))) + (sTotalNumberDays * 22.72 + sTotalNumberDays * 68.18)) - (Val(sLifeRet.Fields("LifeRetirement")) + sTotalDeduction + Val(sMedicare.Fields("Medicare"))), 2)
            Else
                .Fields("PERA") = sAllowance.Fields("PERA")
                .Fields("ADCOM") = sAllowance.Fields("ADCOM")
                .Fields("LifeRet") = FormatNumber(sLifeRet.Fields("LifeRetirement"), 2)
                .Fields("AmountRecived") = FormatNumber((sTotalNumberDays * Val(sGeneratePayrollJO.Fields("SalaryBasic"))) + Val(sAllowance.Fields("PERA")) + Val(sAllowance.Fields("ADCOM")) - (Val(sLifeRet.Fields("LifeRetirement")) + sTotalDeduction + Val(sMedicare.Fields("Medicare"))), 2)
            End If
        
        .Fields("Medicare") = sMedicare.Fields("Medicare")
        .Fields("NumDayWork") = sTotalNumberDays
        .Fields("RateperDay") = sGeneratePayrollJO.Fields("SalaryBasic")
        .Fields("GrossSalary") = FormatNumber(sTotalNumberDays * Val(sGeneratePayrollJO.Fields("SalaryBasic")), 2)
        .Fields("GrossSalary1") = sTotalNumberDays * Val(sGeneratePayrollJO.Fields("SalaryBasic"))
        .Fields("UnderTime") = FormatNumber(sTotalDeduction, 2)
        .Fields("Term") = Combo2.Text
        .Fields("Designation") = "Admin Aide"
        .Update
    End With
    sGeneratePayrollJO.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set sGeneratePayrollJO = Nothing
Set sProject = Nothing
Set sNumDayWork = Nothing
Set CasualPay = Nothing
Set sLifeRet = Nothing
Set sCreated = Nothing
Set sAllowance = Nothing
Set sMedicare = Nothing
End Sub

Public Sub LoadCasualPay()
On Error Resume Next
Dim x As ListItem

CasualPay.Requery

While Not CasualPay.EOF
    Set x = ListView1.ListItems.Add(, , CasualPay.AbsolutePosition, 1, 1)
        x.SubItems(1) = CasualPay.Fields("Name")
        x.SubItems(2) = CasualPay.Fields("Designation")
        x.SubItems(3) = CasualPay.Fields("RatePerDay")
        x.SubItems(4) = CasualPay.Fields("NumDayWork")
        x.SubItems(5) = CasualPay.Fields("GrossSalary")
        x.SubItems(6) = CasualPay.Fields("PERA")
        x.SubItems(7) = CasualPay.Fields("ADCOM")
        x.SubItems(8) = CasualPay.Fields("LifeRet")
        x.SubItems(9) = CasualPay.Fields("Medicare")
        x.SubItems(10) = CasualPay.Fields("WHoldingTax")
        x.SubItems(11) = CasualPay.Fields("GSISSalaryLoan")
        x.SubItems(12) = CasualPay.Fields("GSISPolicyLoan")
        x.SubItems(13) = CasualPay.Fields("ELA")
        x.SubItems(14) = CasualPay.Fields("CAL")
        x.SubItems(15) = CasualPay.Fields("SOS")
        x.SubItems(16) = CasualPay.Fields("ConsolLoan")
        x.SubItems(17) = CasualPay.Fields("PAGIBIGPrem")
        x.SubItems(18) = CasualPay.Fields("TCRFEMPC")
        x.SubItems(19) = CasualPay.Fields("RBTC")
        x.SubItems(20) = CasualPay.Fields("UnderTime")
        x.SubItems(21) = CasualPay.Fields("AmountRecived")
    CasualPay.MoveNext
Wend
End Sub

Public Sub CountNumberOfDays()
sTotalNumberDays = 0

sNumDayWork.Requery
sNumDayWork.Filter = adFilterNone

sNumDayWork.Filter = "Name"
End Sub

Public Sub LoadMonth()
With Combo1
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

Private Sub ListView1_DblClick()
Call CommandButton5_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Not CasualPay.RecordCount < 1 Then CasualPay.AbsolutePosition = ListView1.SelectedItem
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub Text2_Change()
On Error Resume Next
    
    'check length
    If Len(Text2.Text) < 1 Then
        With Text2
            .SelStart = 0
            .SelLength = Len(Text2.Text)
        End With
        Exit Sub
    End If
    
    FindLVItem ListView1, Trim(Text2.Text), , , , True ', , tmpMultiSelect, tmpInverseSelection
End Sub
