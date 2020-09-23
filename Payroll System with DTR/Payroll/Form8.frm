VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Life and Retirement Rate"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3900
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   210
      TabIndex        =   2
      Top             =   735
      Width           =   3480
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
      Left            =   1680
      TabIndex        =   1
      Top             =   210
      Width           =   2010
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   435
      Left            =   1050
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
      Left            =   2415
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
      Caption         =   "Enter Here:"
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
      Width           =   1380
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
If sLifeRetRate.RecordCount < 1 Then
    If isempty(Text1) = True Then Exit Sub
    If Not IsNumeric(Text1.Text) Then MsgBox "Please type a numeric value.", vbExclamation, "Record Manager": Exit Sub
    sLifeRetRate.AddNew
    sLifeRetRate.Fields("LifeRetirement") = Text1.Text
    sLifeRetRate.Update
    
    MsgBox "Adding of life and retirement rate has been succesful.", vbInformation, "Record Manager"
    Unload Me
Else
    If isempty(Text1) = True Then Exit Sub
    If Not IsNumeric(Text1.Text) Then MsgBox "Please type a numeric value.", vbExclamation, "Record Manager": Exit Sub
    sLifeRetRate.Fields("LifeRetirement") = Text1.Text
    sLifeRetRate.Update
    
    MsgBox "Changes in life and retirement rate has been succesful.", vbInformation, "Record Manager"
    Unload Me
End If

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call set_rec_getData(sLifeRetRate, cn, "Select * From tblLifeRetirement")

If Not sLifeRetRate.RecordCount < 1 Then Text1.Text = sLifeRetRate.Fields("LifeRetirement")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set sLifeRetRate = Nothing
End Sub

