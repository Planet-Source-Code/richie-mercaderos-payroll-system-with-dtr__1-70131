VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Payroll System version 1.0"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7125
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   330
      Top             =   2550
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1770
      Top             =   1260
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   6690
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   767
      SimpleText      =   "Richie S. Mercaderos - Origin"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "Form1.frx":11E87
            Text            =   "Username:"
            TextSave        =   "Username:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            TextSave        =   "2/22/2008"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "i32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Administrator's Option"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modify Emploee Information"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modify Emploee Rate"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Back-up Settings"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Calculator"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help Topics"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   270
      Top             =   960
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
            Picture         =   "Form1.frx":12A59
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13733
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1440D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":150E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15DC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16A9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17775
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1844F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19129
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19E03
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AADD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B7B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C491
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D16B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DE45
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EB1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F7F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":204D3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   285
      Left            =   8370
      TabIndex        =   2
      Top             =   6420
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuModifyEmployee 
         Caption         =   "Modify Employee Information"
         Shortcut        =   ^M
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminOption 
         Caption         =   "Administrator's Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLogHistory 
         Caption         =   "Log History"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSwitch 
         Caption         =   "Switch User"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuRate 
      Caption         =   "&Rate"
      Begin VB.Menu mnuModEmpRate 
         Caption         =   "Modify Employee Rate"
         Shortcut        =   ^R
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedRate 
         Caption         =   "Medicare Rate Non-Regular (Wages)"
      End
      Begin VB.Menu mnuAllowances 
         Caption         =   "Allowances for Non-Regular (Wages)"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLifeRet 
         Caption         =   "Life and Retirement Rate"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuPayroll 
      Caption         =   "&Payroll"
      Begin VB.Menu mnuCreatePayroll 
         Caption         =   "Create Payroll"
         Begin VB.Menu mnuNoneReg 
            Caption         =   "Non-Regular"
            Begin VB.Menu mnuWages 
               Caption         =   "Wages"
            End
            Begin VB.Menu sep15 
               Caption         =   "-"
            End
            Begin VB.Menu mnuJob 
               Caption         =   "Job Order"
            End
         End
         Begin VB.Menu sep5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuModifyLoanTaxes 
            Caption         =   "Regular"
            Shortcut        =   ^P
         End
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "&Tools"
      Begin VB.Menu mnuLock 
         Caption         =   "Lock Application Setting"
      End
      Begin VB.Menu sep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Back-up Setting"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
sTimeSet = GetSetting(App.Title, "SetVal", "TimeVal", 1)
End Sub

Private Sub Form_Load()
Call set_conn_getData(cn, App.Path & "\MasterFile.mdb", True, "xxx")
Call set_rec_getData(rs, cn, "Select * From qryEmployeeInfo")
Call set_rec_getData(sUsers, cn, "Select * From tblUsers")
Call set_rec_getData(sUserLogs, cn, "Select * From tblUserLogs")

sTimeSet = GetSetting(App.Title, "SetVal", "TimeVal", 1)

If Not sUsers.RecordCount < 1 Then
mnuSwitch.Enabled = True
Me.Show
Form9.Show vbModal
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Dim sResponse As Integer

sResponse = MsgBox("This action is to terminate the application." & vbNewLine & vbNewLine & "Are you sure you want to terminate the application?", vbExclamation + vbOKCancel, "Warning")

If sResponse = vbOK Then
    sUserLogs.Fields(3) = Time
    sUserLogs.Update
    End
Else
    Cancel = True
End If
End Sub

Private Sub mnuAbout_Click()
Form20.Show vbModal
End Sub

Private Sub mnuAdminOption_Click()
Form13.Show vbModal
End Sub

Private Sub mnuAllowances_Click()
If Label1.Caption = "User" Then MsgBox "System user is not allow to use this action.", vbExclamation, "Unauthorized Person": Exit Sub
Form7.Show vbModal
End Sub

Private Sub mnuBackup_Click()
Frm_backup.Show vbModal
End Sub

Private Sub mnuCalc_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuJO_Click()
Form12.Show vbModal
End Sub

Private Sub mnuJob_Click()
If Label1.Caption = "User" Then MsgBox "System user is not allow to use this action.", vbExclamation, "Unauthorized Person": Exit Sub
Form12.Show vbModal
End Sub

Private Sub mnuLifeRet_Click()
If Label1.Caption = "User" Then MsgBox "System user is not allow to use this action.", vbExclamation, "Unauthorized Person": Exit Sub
Form8.Show vbModal
End Sub

Private Sub mnuLock_Click()
Form14.Show vbModal
End Sub

Private Sub mnuLogHistory_Click()
Form17.Show vbModal
End Sub

Private Sub mnuMedRate_Click()
If Label1.Caption = "User" Then MsgBox "System user is not allow to use this action.", vbExclamation, "Unauthorized Person": Exit Sub
Form6.Show vbModal
End Sub

Private Sub mnuModEmpRate_Click()
If Label1.Caption = "User" Then MsgBox "System user is not allow to use this action.", vbExclamation, "Unauthorized Person": Exit Sub
Form4.Show vbModal
End Sub

Private Sub mnuModifyEmployee_Click()
Form2.Show vbModal
End Sub

Private Sub mnuModifyLoanTaxes_Click()
If Form1.Label1.Caption = "User" Then MsgBox "System user is not allow to use this action.", vbExclamation, "Unauthorized Person": Exit Sub
Form10.Show vbModal
End Sub

Private Sub mnuReg_Click()
Form9.Show vbModal
End Sub

Private Sub mnuWages_Click()
If Label1.Caption = "User" Then MsgBox "System user is not allow to use this action.", vbExclamation, "Unauthorized Person": Exit Sub
Form18.Show vbModal
End Sub

Private Sub mnuSwitch_Click()
On Error Resume Next
Dim ans As Integer

ans = MsgBox("Switch User?", vbExclamation, "Payoll System")

If ans = vbOK Then

    sUserLogs.Fields(3) = Time
    sUserLogs.Update
    StatusBar1.Panels(3).Text = ""
    Form9.Show vbModal
End If
End Sub

Private Sub Timer1_Timer()
Static ic As Integer
Static op As POINTAPI
    
Dim p As POINTAPI
    
GetCursorPos p
If (p.x < (op.x + 5) And p.x > op.x - 5) And (p.Y < (op.Y + 5) And p.Y > op.Y - 5) Then
    ic = ic + 1
Else
    ic = 0
End If
    
op.x = p.x
op.Y = p.Y

If ic > sTimeSet Then
    ic = 0
    Form16.Show vbModal
End If
End Sub

Private Sub Timer2_Timer()
StatusBar1.Panels(4).Text = Format(Time, "hh:mm:ss am/pm")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case (Button.Index)
    Case 2: Call mnuAdminOption_Click
    Case 4: Call mnuModifyEmployee_Click
    Case 5: Call mnuModEmpRate_Click
    Case 7: Call mnuBackup_Click
    Case 8: Call mnuCalc_Click
End Select
End Sub
