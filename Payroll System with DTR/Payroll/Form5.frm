VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   210
      TabIndex        =   4
      Top             =   1260
      Width           =   4845
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
      Left            =   2835
      TabIndex        =   3
      Top             =   735
      Width           =   2220
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
      Left            =   1785
      TabIndex        =   1
      Top             =   210
      Width           =   3270
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   435
      Left            =   2415
      TabIndex        =   6
      Top             =   1470
      Width           =   1275
      Caption         =   "Cancel "
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
      Left            =   3780
      TabIndex        =   5
      Top             =   1470
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
   Begin VB.Label Label2 
      Caption         =   "Salary Per Day:   Php"
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
      TabIndex        =   2
      Top             =   735
      Width           =   2535
   End
   Begin VB.Label Label1 
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
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1485
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sInfo                   As New ADODB.Recordset

Public add_state            As Boolean
Public sDupDesignation      As String

Private Sub CommandButton1_Click()
If isempty(Text1) Then Exit Sub
If isempty(Text2) Then Exit Sub

If Not IsNumeric(Text2.Text) Then MsgBox "Please type a numeric value.", vbExclamation, "Record Manager": Exit Sub

If sDupDesignation <> Text1.Text Then
    If if_exist("tblDesignation", "Designation", Text1) = True Then Exit Sub
End If

With sRate
    If add_state = True Then .AddNew
        .Fields("Designation") = Text1.Text
        .Fields("SalaryperDay") = Text2.Text
        .Update
End With

If add_state = True Then
    MsgBox "Adding of new designation has been successfull.", vbInformation, "Save Complete"
    Dim rep As Integer
    rep = MsgBox("Do you want to add another employee?", vbQuestion + vbYesNo, "Record Manager")
    If rep = vbYes Then
            
        Text1.Text = ""
        Text2.Text = ""
    
        Text1.SetFocus
        
        Form4.ListView1.ListItems.Clear
        Form4.LoadEmployeeRate
    Else
        Form4.ListView1.ListItems.Clear
        Form4.LoadEmployeeRate
        Unload Me
    End If
    rep = 0
Else
    MsgBox "Changes in the designation record has been successfully saved.", vbInformation, "Save Complete"
    Dim pos As Long
    
    pos = sRate.AbsolutePosition
    
    Form4.ListView1.ListItems.Clear
    Form4.LoadEmployeeRate
    
    Form4.ListView1.ListItems.Item(pos).EnsureVisible
    Form4.ListView1.ListItems.Item(pos).Selected = True
    
    sInfo.Requery
    sInfo.Filter = adFilterNone
    
    sInfo.Filter = "Designation='" & Text1.Text & "'"
        
        While Not sInfo.EOF
            sInfo.Fields("SalaryBasic") = Text2.Text
            sInfo.Update
            sInfo.MoveNext
        Wend
        
    pos = 0
    Unload Me
End If
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Call set_rec_getData(sInfo, cn, "Select * From qryEmployeeInfo")

If add_state = True Then
    Me.Caption = "Add New Designation"
Else
    With sRate
        Text1.Text = .Fields("Designation")
        Text2.Text = .Fields("SalaryperDay")
    End With
End If
sDupDesignation = sRate.Fields("Designation").Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set sInfo = Nothing
End Sub
