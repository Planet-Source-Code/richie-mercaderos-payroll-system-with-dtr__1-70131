VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   735
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2640
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   4657
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1500
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Designation"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Salary "
         Object.Width           =   5292
      EndProperty
   End
   Begin MSForms.CommandButton CommandButton3 
      Height          =   435
      Left            =   2835
      TabIndex        =   3
      Top             =   2940
      Width           =   1275
      Caption         =   "Add New"
      Size            =   "2249;767"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   435
      Left            =   4200
      TabIndex        =   2
      Top             =   2940
      Width           =   1275
      Caption         =   "Edit"
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
      Left            =   5775
      TabIndex        =   1
      Top             =   2940
      Width           =   1170
      Caption         =   "Close"
      Size            =   "2064;767"
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No record in the list. Please check it!", vbExclamation, "Record List": Exit Sub
Form5.add_state = False
Form5.Show vbModal
End Sub

Private Sub CommandButton3_Click()
Form5.add_state = True
Form5.Show vbModal
End Sub

Private Sub Form_Load()
Call set_rec_getData(sRate, cn, "Select * From qryDesignation")

Call LoadEmployeeRate

ListView1.ListItems.Item(1).Selected = True
sRate.AbsolutePosition = ListView1.SelectedItem
End Sub

Public Sub LoadEmployeeRate()
Dim x As ListItem
Dim a As String

sRate.Requery

While Not sRate.EOF
    
    a = sRate.Fields("Designation").Value
    
    Set x = ListView1.ListItems.Add(, , sRate.AbsolutePosition, 1, 1)
        
        If UCase(a) = UCase("Admin Aide I (JO)") Or UCase(a) = UCase("Admin Aide I (WAGES)") Then
            x.SubItems(1) = sRate.Fields("Designation")
            x.SubItems(2) = "Php" & FormatNumber(sRate.Fields("SalaryperDay"), 2) & " per Day"
        Else
            x.SubItems(1) = sRate.Fields("Designation")
            x.SubItems(2) = "Php" & FormatNumber(sRate.Fields("SalaryperDay"), 2)
        End If
    sRate.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set sRate = Nothing
End Sub

Private Sub ListView1_DblClick()
Call CommandButton2_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Not sRate.RecordCount < 1 Then sRate.AbsolutePosition = ListView1.SelectedItem
End Sub
