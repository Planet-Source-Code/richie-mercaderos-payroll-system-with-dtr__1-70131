VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form19 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Loans, Taxes and Deductions"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Caption         =   "Legend"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   7530
      TabIndex        =   46
      Top             =   3270
      Width           =   2895
      Begin VB.Label Label28 
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
         Height          =   225
         Left            =   570
         TabIndex        =   50
         Top             =   810
         Width           =   225
      End
      Begin VB.Label Label33 
         Caption         =   "Additionals"
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
         Left            =   840
         TabIndex        =   49
         Top             =   810
         Width           =   1905
      End
      Begin VB.Label Label34 
         BackColor       =   &H80000014&
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
         Height          =   225
         Left            =   570
         TabIndex        =   48
         Top             =   480
         Width           =   225
      End
      Begin VB.Label Label35 
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
         Height          =   255
         Left            =   840
         TabIndex        =   47
         Top             =   480
         Width           =   1905
      End
   End
   Begin VB.Frame Frame6 
      Height          =   30
      Left            =   150
      TabIndex        =   44
      Top             =   2670
      Width           =   4995
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   180
      TabIndex        =   43
      Top             =   1830
      Width           =   4965
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   5460
      TabIndex        =   38
      Top             =   1560
      Width           =   5025
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
      Left            =   2220
      TabIndex        =   12
      Top             =   765
      Width           =   2895
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
      Height          =   285
      Left            =   2220
      TabIndex        =   11
      Top             =   1125
      Width           =   2895
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
      Height          =   285
      Left            =   2220
      TabIndex        =   10
      Top             =   1485
      Width           =   2895
   End
   Begin VB.TextBox Text7 
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
      Left            =   2220
      TabIndex        =   9
      Top             =   3165
      Width           =   2895
   End
   Begin VB.TextBox Text8 
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
      Left            =   2220
      TabIndex        =   8
      Top             =   3525
      Width           =   2895
   End
   Begin VB.TextBox Text9 
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
      Left            =   2220
      TabIndex        =   7
      Top             =   3885
      Width           =   2895
   End
   Begin VB.TextBox Text10 
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
      Left            =   2220
      TabIndex        =   6
      Top             =   4245
      Width           =   2895
   End
   Begin VB.TextBox Text11 
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
      Left            =   2220
      TabIndex        =   5
      Top             =   4605
      Width           =   2895
   End
   Begin VB.TextBox Text12 
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
      Left            =   2220
      TabIndex        =   4
      Top             =   4965
      Width           =   2895
   End
   Begin VB.TextBox Text13 
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
      Left            =   7530
      TabIndex        =   3
      Top             =   780
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Height          =   15
      Left            =   180
      TabIndex        =   2
      Top             =   570
      Width           =   10335
   End
   Begin VB.Frame Frame5 
      Height          =   45
      Left            =   90
      TabIndex        =   1
      Top             =   5460
      Width           =   10425
   End
   Begin VB.Frame Frame7 
      Height          =   15
      Left            =   60
      TabIndex        =   0
      Top             =   6210
      Width           =   10485
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   345
      Left            =   7530
      TabIndex        =   52
      Top             =   4740
      Width           =   2925
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFC0&
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
      Height          =   375
      Left            =   1590
      TabIndex        =   51
      Top             =   90
      Width           =   3555
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   7560
      TabIndex        =   45
      Top             =   1800
      Width           =   2835
   End
   Begin VB.Label Label31 
      Caption         =   "Medicare:"
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
      Left            =   5490
      TabIndex        =   42
      Top             =   1140
      Width           =   1905
   End
   Begin VB.Label Label32 
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
      Height          =   315
      Left            =   2220
      TabIndex        =   41
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label30 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   7530
      TabIndex        =   40
      Top             =   1110
      Width           =   2895
   End
   Begin VB.Label Label25 
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
      Height          =   315
      Left            =   2220
      TabIndex        =   39
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   7560
      TabIndex        =   37
      Top             =   2550
      Width           =   2835
   End
   Begin VB.Label Label14 
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
      Height          =   255
      Left            =   5460
      TabIndex        =   36
      Top             =   2580
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000014&
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
      Height          =   315
      Left            =   2220
      TabIndex        =   35
      Top             =   2790
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "No. of days work:"
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
      Left            =   5460
      TabIndex        =   34
      Top             =   1830
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Withholding Tax:"
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
      Left            =   180
      TabIndex        =   33
      Top             =   765
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "GSIS Salary Loan:"
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
      Left            =   180
      TabIndex        =   32
      Top             =   1125
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "GSIS Policy Loan:"
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
      Left            =   180
      TabIndex        =   31
      Top             =   1485
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "PERA:"
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
      Left            =   180
      TabIndex        =   30
      Top             =   1950
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "ADCOM:"
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
      Left            =   180
      TabIndex        =   29
      Top             =   2310
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Life and Ret.:"
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
      Left            =   180
      TabIndex        =   28
      Top             =   2820
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "ELA:"
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
      Left            =   180
      TabIndex        =   27
      Top             =   3180
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "CAL:"
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
      Left            =   180
      TabIndex        =   26
      Top             =   3540
      Width           =   2055
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   180
      TabIndex        =   25
      Top             =   3900
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Consol Loans:"
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
      Left            =   180
      TabIndex        =   24
      Top             =   4230
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "PAG-IBIG Prem:"
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
      Left            =   180
      TabIndex        =   23
      Top             =   4590
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "TCRF EMPC:"
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
      Left            =   180
      TabIndex        =   22
      Top             =   4980
      Width           =   2055
   End
   Begin VB.Label Label16 
      Caption         =   "RBTC:"
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
      Left            =   5490
      TabIndex        =   21
      Top             =   780
      Width           =   2055
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   8940
      TabIndex        =   20
      Top             =   6420
      Width           =   1575
      Caption         =   "Save"
      Size            =   "2778;873"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   495
      Left            =   7260
      TabIndex        =   19
      Top             =   6420
      Width           =   1575
      Caption         =   "Cancel"
      Size            =   "2778;873"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label20 
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
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "Rate per Day:"
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
      Left            =   5460
      TabIndex        =   17
      Top             =   2190
      Width           =   1575
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFC0&
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
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   2160
      Width           =   2835
   End
   Begin VB.Label Label27 
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
      Height          =   255
      Left            =   5430
      TabIndex        =   15
      Top             =   5700
      Width           =   2055
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   7635
      TabIndex        =   14
      Top             =   5640
      Width           =   2910
   End
   Begin VB.Label Label42 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   7620
      TabIndex        =   13
      Top             =   5700
      Visible         =   0   'False
      Width           =   2910
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
With CasualPay
    .Fields("WHoldingTax") = Text1.Text
    .Fields("GSISSalaryLoan") = Text2.Text
    .Fields("GSISPolicyLoan") = Text3.Text
    .Fields("ELA") = Text7.Text
    .Fields("CAL") = Text8.Text
    .Fields("SOS") = Text9.Text
    .Fields("ConsolLoan") = Text10.Text
    .Fields("PAGIBIGPrem") = Text11.Text
    .Fields("TCRFEMPC") = Text12.Text
    .Fields("RBTC") = Text13.Text
    .Fields("AmountRecived") = Label29.Caption
    .Update
End With

MsgBox "Adding of loans, taxes and deductions has been successful.", vbInformation, "Add Complete"
Dim pos As Integer

pos = CasualPay.AbsolutePosition

Form18.ListView1.ListItems.Clear
Form18.LoadCasualPay

Form18.ListView1.ListItems.Item(pos).EnsureVisible
Form18.ListView1.ListItems.Item(pos).Selected = True

pos = 0
Unload Me
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next

With CasualPay
    Label22.Caption = .Fields("Name")
    Text1.Text = .Fields("WHoldingTax")
    Text2.Text = .Fields("GSISSalaryLoan")
    Text3.Text = .Fields("GSISPolicyLoan")
    Label25.Caption = .Fields("PERA")
    Label32.Caption = .Fields("ADCOM")
    Label13.Caption = .Fields("LifeRet")
    Text7.Text = .Fields("ELA")
    Text8.Text = .Fields("CAL")
    Text9.Text = .Fields("SOS")
    Text10.Text = .Fields("ConsolLoan")
    Text11.Text = .Fields("PAGIBIGPrem")
    Text12.Text = .Fields("TCRFEMPC")
    Text13.Text = .Fields("RBTC")
    Label26.Caption = .Fields("NumDayWork")
    Label24.Caption = .Fields("RateperDay")
    Label17.Caption = .Fields("GrossSalary")
    Label18.Caption = .Fields("GrossSalary1")
    Label30.Caption = .Fields("Medicare")
    Label29.Caption = .Fields("AmountRecived")
End With
End Sub

Private Sub Text1_Change()
Call CalculateDeductions
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text10_Change()
Call CalculateDeductions
End Sub

Private Sub Text11_Change()
Call CalculateDeductions
End Sub

Private Sub Text12_Change()
Call CalculateDeductions
End Sub

Private Sub Text13_Change()
Call CalculateDeductions
End Sub

Private Sub Text2_Change()
Call CalculateDeductions
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_Change()
Call CalculateDeductions
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then Text7.SetFocus
End Sub

Private Sub Text7_Change()
Call CalculateDeductions
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then Text8.SetFocus
End Sub

Private Sub Text8_Change()
Call CalculateDeductions
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then Text9.SetFocus
End Sub

Private Sub Text9_Change()
Call CalculateDeductions
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then
    If Text10.Enabled = False Then
        Text11.SetFocus
    Else
        Text10.SetFocus
    End If
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then Text11.SetFocus
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then Text12.SetFocus
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then Text13.SetFocus
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 59 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46) Then KeyAscii = 0
If KeyAscii = 13 Then CommandButton1.SetFocus
End Sub

Public Sub CalculateDeductions()
Dim TotalDeduction As Double
TotalDeduction = Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Label13.Caption) + Val(Label30.Caption)
Label29.Caption = (Val(Label18.Caption) + Val(Label25.Caption) + Val(Label32.Caption)) - TotalDeduction
End Sub
