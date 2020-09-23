VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Form11.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ImageList1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
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
         Left            =   6840
         TabIndex        =   3
         Top             =   4320
         Width           =   1335
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
         Left            =   2160
         TabIndex        =   2
         Top             =   4320
         Width           =   4455
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1200
         Top             =   1080
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
               Picture         =   "Form11.frx":001C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3585
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   6324
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   15269887
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New TUR"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "TimeInAM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "TimeOutAM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Hours & Mins (Undertime)"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "TimeInPM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "TimeOutPM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Hours & Mins (Undertime)"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Total Undertime (Hours & Mins)"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Lastname:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   4320
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rrs As New ADODB.Recordset

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Today's Record: " & FormatDateTime(Date, vbLongDate)
Call set_rec_getData(rrs, cn, "Select * From qryComputed")

rrs.Requery

rrs.Filter = adFilterNone
rrs.Filter = "CDate='" & Format(Now, "mm/dd/yyyy") & "'"

Call LoadRecord
End Sub

Public Sub LoadRecord()
On Error Resume Next
Dim X As ListItem

While Not rrs.EOF
    Set X = ListView1.ListItems.Add(, , rrs.AbsolutePosition, 1, 1)
    
        X.SubItems(1) = rrs.Fields("Name")
        X.SubItems(2) = rrs.Fields("CDate")
        X.SubItems(3) = rrs.Fields("AMIn")
        X.SubItems(4) = rrs.Fields("AMOut")
        X.SubItems(5) = rrs.Fields("ATHour") & " hr(s) " & rrs.Fields("ATMin") & " min(s)"
        X.SubItems(6) = rrs.Fields("PMIn")
        X.SubItems(7) = rrs.Fields("PMOut")
        X.SubItems(8) = rrs.Fields("PTHour") & " hr(s) " & rrs.Fields("PTMin") & " min(s)"
        X.SubItems(9) = rrs.Fields("TTinHour") & " hr(s) " & rrs.Fields("TTinMin") & " min(s)"
    
    rrs.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rrs = Nothing
End Sub

Private Sub Command1_Click()
    Dim tmpMultiSelect As Boolean
    Dim tmpInverseSelection As Boolean
    

    On Error Resume Next
    
    If Len(Text1.Text) < 1 Then
        With Text1
            .SelStart = 0
            .SelLength = Len(Text1.Text)
        End With
        Exit Sub
    End If
    
    'execute find
    
    FindLVItem ListView1, Trim(Text1.Text)

End Sub

Private Sub Text1_Change()
Call Command1_Click
End Sub
