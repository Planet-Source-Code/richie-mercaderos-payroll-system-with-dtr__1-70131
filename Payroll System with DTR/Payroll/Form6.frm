VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   105
      TabIndex        =   2
      Top             =   735
      Width           =   3900
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
      Height          =   330
      Left            =   1575
      TabIndex        =   1
      Top             =   210
      Width           =   2430
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   435
      Left            =   1365
      TabIndex        =   4
      Top             =   945
      Width           =   1275
      Caption         =   "Cancel"
      Size            =   "2249;767"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   435
      Left            =   2730
      TabIndex        =   3
      Top             =   945
      Width           =   1275
      Caption         =   "Save"
      Size            =   "2249;767"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
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
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
If sMedRate.RecordCount < 1 Then
    If isempty(Text1) = True Then Exit Sub
    If Not IsNumeric(Text1.Text) Then MsgBox "Please type a numeric value.", vbExclamation, "Record Manager": Exit Sub
    sMedRate.AddNew
    sMedRate.Fields("Medicare") = Text1.Text
    sMedRate.Update
    
    MsgBox "Adding of medicare rate has been succesful.", vbInformation, "Record Manager"
    Unload Me
Else
    If isempty(Text1) = True Then Exit Sub
    If Not IsNumeric(Text1.Text) Then MsgBox "Please type a numeric value.", vbExclamation, "Record Manager": Exit Sub
    sMedRate.Fields("Medicare") = Text1.Text
    sMedRate.Update
    
    MsgBox "Changes in medicare rate has been succesful.", vbInformation, "Record Manager"
    Unload Me
End If

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call set_rec_getData(sMedRate, cn, "Select * From tblMedicare")

If Not sMedRate.RecordCount < 1 Then Text1.Text = sMedRate.Fields("Medicare")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set sMedRate = Nothing
End Sub
