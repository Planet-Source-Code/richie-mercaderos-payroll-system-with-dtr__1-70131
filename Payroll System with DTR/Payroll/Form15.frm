VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   1710
      TabIndex        =   31
      Top             =   2370
      Width           =   2415
   End
   Begin VB.Frame Frame6 
      Height          =   45
      Left            =   240
      TabIndex        =   29
      Top             =   4680
      Width           =   5655
   End
   Begin VB.Frame Frame5 
      Height          =   30
      Left            =   240
      TabIndex        =   22
      Top             =   1410
      Width           =   5655
   End
   Begin VB.Frame Frame4 
      Height          =   30
      Left            =   240
      TabIndex        =   16
      Top             =   5250
      Width           =   5655
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   5655
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Unsubmitt"
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
      Left            =   2340
      TabIndex        =   10
      Top             =   3720
      Width           =   2205
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Submitted"
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
      Left            =   2340
      TabIndex        =   9
      Top             =   3390
      Width           =   2205
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   210
      TabIndex        =   7
      Top             =   3300
      Width           =   5655
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
      Height          =   375
      Left            =   1710
      TabIndex        =   6
      Top             =   1920
      Width           =   2415
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
      Height          =   375
      Left            =   1710
      TabIndex        =   5
      Top             =   1470
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   5655
   End
   Begin VB.Line Line4 
      X1              =   4350
      X2              =   4200
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Line Line3 
      X1              =   4380
      X2              =   4500
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Line Line2 
      X1              =   4350
      X2              =   4350
      Y1              =   1650
      Y2              =   3060
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4350
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   1680
      TabIndex        =   30
      Top             =   2880
      Width           =   2445
   End
   Begin VB.Label Label30 
      Caption         =   "X = Grooss Salary"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3990
      TabIndex        =   28
      Top             =   900
      Width           =   1875
   End
   Begin VB.Label Label28 
      Caption         =   "Deductions"
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
      Left            =   4590
      TabIndex        =   27
      Top             =   2220
      Width           =   1365
   End
   Begin VB.Label Label26 
      BackColor       =   &H80000018&
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
      Left            =   1950
      TabIndex        =   26
      Top             =   990
      Width           =   1965
   End
   Begin VB.Label Label25 
      Caption         =   "No. of Days:"
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
      Left            =   270
      TabIndex        =   25
      Top             =   1020
      Width           =   1605
   End
   Begin VB.Label Label24 
      BackColor       =   &H80000018&
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
      Left            =   1950
      TabIndex        =   24
      Top             =   600
      Width           =   1965
   End
   Begin VB.Label Label23 
      Caption         =   "Salary/Day:"
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
      Left            =   270
      TabIndex        =   23
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label Label20 
      Caption         =   "Undertime:"
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
      Left            =   240
      TabIndex        =   21
      Top             =   2910
      Width           =   1365
   End
   Begin VB.Label Label14 
      Caption         =   "COOP:"
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
      Left            =   240
      TabIndex        =   20
      Top             =   2430
      Width           =   1365
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000018&
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
      Left            =   210
      TabIndex        =   19
      Top             =   8370
      Visible         =   0   'False
      Width           =   1665
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   585
      Left            =   2430
      TabIndex        =   18
      Top             =   5490
      Width           =   1635
      Caption         =   "Cancel"
      Size            =   "2884;1032"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   585
      Left            =   4170
      TabIndex        =   17
      Top             =   5490
      Width           =   1635
      Caption         =   "Save"
      Size            =   "2884;1032"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label9 
      Caption         =   "Net Amount:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   4860
      Width           =   1665
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   2760
      TabIndex        =   14
      Top             =   4830
      Width           =   3045
   End
   Begin VB.Label Label7 
      Caption         =   "Gross Salary:"
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
      Left            =   180
      TabIndex        =   13
      Top             =   4230
      Width           =   1965
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   2730
      TabIndex        =   12
      Top             =   4200
      Width           =   3045
   End
   Begin VB.Label Label5 
      Caption         =   "DTR Submittion:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   3390
      Width           =   1965
   End
   Begin VB.Label Label4 
      Caption         =   "RBTC Loans:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1950
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   "SOS:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1470
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000018&
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
      Left            =   1710
      TabIndex        =   1
      Top             =   60
      Width           =   4125
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   90
      Width           =   1365
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public add_state As Boolean

Private Sub CommandButton1_Click()
'Label8.Caption = Val(Label21.Caption) - Val(Label11.Caption)

With sCasualPay
    If add_state = True Then .AddNew
    .Fields("SOS") = Text1.Text
    .Fields("CAL") = Text2.Text
    .Fields("COOP") = Text3.Text
    If Option1.Value = True Then
        .Fields("SubmittedDTR") = True
    Else
        .Fields("SubmittedDTR") = False
    End If
    
    .Fields("AmountRecived") = Label8.Caption
    .Update
    
End With

If add_state = True Then
    MsgBox "Adding of employee loans has been successfull.", vbInformation, "Save Complete"
    Dim rep As Integer
    rep = MsgBox("Do you want to add another employee loans?", vbQuestion + vbYesNo, "Record Manager")
    If rep = vbYes Then
            
        Text1.Text = ""
        Text2.Text = ""
        
        Text1.SetFocus
        Form12.ListView1.ListItems.Clear
        Form12.LoadCasualPay
    Else
        Form12.ListView1.ListItems.Clear
        Form12.LoadCasualPay
        Unload Me
    End If
    rep = 0
Else
    MsgBox "Changes in the data has been successfully saved.", vbInformation, "Save Complete"
    Dim pos As Long
    
    pos = sCasualPay.AbsolutePosition
    
    Form12.ListView1.ListItems.Clear
    Form12.LoadCasualPay
    
    Form12.ListView1.ListItems.Item(pos).EnsureVisible
    Form12.ListView1.ListItems.Item(pos).Selected = True
    
    pos = 0
    Unload Me
End If
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next

If add_state = True Then
    Me.Caption = "Add New Loans, Taxes and Deductions - Non-Regulars"
Else
    Me.Caption = "Add Loans, Taxes and Deductions - Regulars"
    With sCasualPay
        Text1.Text = .Fields("SOS")
        Text2.Text = .Fields("CAL")
        Text3.Text = .Fields("COOP")
        
        Label2.Caption = .Fields("Name")
        Label6.Caption = .Fields("GrossSalary1")
        Label19.Caption = .Fields("UnderTime")
        Label12.Caption = .Fields("GrossSalary1")
        Label24.Caption = .Fields("SalaryperDay")
        Label26.Caption = .Fields("NumDayWork")
        Label8.Caption = .Fields("AmountRecived")
    End With
End If

End Sub

Private Sub Text1_Change()
Call Compute
End Sub

Private Sub Text1_GotFocus()
With Text1
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_Change()
Call Compute
End Sub

Private Sub Text2_GotFocus()
With Text2
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub Text3_Change()
Call Compute
End Sub

Private Sub Text3_GotFocus()
With Text3
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Public Sub Compute()
Dim a, b, c, d, g, h

a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Label19.Caption)

g = Val(Label24.Caption)
h = Val(Label26.Caption)

Label6.Caption = g * h
Label8.Caption = (g * h) - (a + b + c + d)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
End Sub
