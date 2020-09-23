VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form9 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   6375
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3000
      TabIndex        =   5
      Top             =   3000
      Width           =   3285
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   3450
      Width           =   6285
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2265
      Left            =   3030
      TabIndex        =   2
      Top             =   660
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   3995
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      TextBackground  =   -1  'True
      _Version        =   393217
      Icons           =   "i32x32"
      SmallIcons      =   "i32x32"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   0
      Top             =   510
      Width           =   6285
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   300
      Top             =   1950
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
            Picture         =   "Form9.frx":B6CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":C3A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":D07E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":DD58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":EA32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":F70C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":103E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":110C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":11D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":12A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":1374E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":14428
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":15102
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":15DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":16AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":17790
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":1846A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form9.frx":19144
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   540
      TabIndex        =   8
      Top             =   660
      Width           =   2385
   End
   Begin MSForms.CommandButton Command2 
      Height          =   525
      Left            =   3360
      TabIndex        =   7
      Top             =   3600
      Width           =   1395
      Caption         =   "Cancel"
      Size            =   "2461;926"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton Command1 
      Height          =   525
      Left            =   4860
      TabIndex        =   6
      Top             =   3600
      Width           =   1395
      Caption         =   "Enter"
      Size            =   "2461;926"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1830
      TabIndex        =   4
      Top             =   3030
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   3825
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If sLoginUser(sUsers, "Username", ListView1.SelectedItem.Text) = True Then
    If Text1.Text = sUsers.Fields(2) Then
        sCurrUser = sUsers.Fields(2)
        sUserLogs.AddNew
        sUserLogs.Fields(0) = Date
        sUserLogs.Fields(1) = ListView1.SelectedItem.Text
        sUserLogs.Fields(2) = Time
        sUserLogs.Update
        
        Form1.Label1.Caption = sUsers.Fields(3)
        Form1.StatusBar1.Panels(3) = ListView1.SelectedItem.Text
        Unload Me
    Else
        MsgBox "Incorrect password.", vbCritical, "Password Error"
        Exit Sub
    End If
Else
    MsgBox "Invalid username.", vbCritical, "Username Error"
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Call LoadUsers
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub


Public Sub LoadUsers()
While Not sUsers.EOF
    If sUsers.Fields(3) = "Administrator" Then
        ListView1.ListItems.Add , , sUsers.Fields(0), 1, 3
    Else
        ListView1.ListItems.Add , , sUsers.Fields(0), 1, 1
    End If
    
    sUsers.MoveNext
Wend
End Sub
