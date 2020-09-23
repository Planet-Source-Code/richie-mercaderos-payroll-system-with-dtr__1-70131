VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Time and Date Changes"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2880
      Top             =   2640
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   7455
      Begin MSComctlLib.ListView ListView2 
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2778
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Old Date and Time"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "New Time and Date"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description of Time and Date Changing"
            Object.Width           =   10583
         EndProperty
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   7409
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Username"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":08DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4260
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Autorized Persons"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Log Details"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Loading. . ."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_user As New ADODB.Recordset
Dim rs_log As New ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True
Label1.Visible = True
End Sub

Private Sub Form_Load()
Call set_rec_getData(rs_user, cn, "Select * From tblAuthorizedPerson")
Call set_rec_getData(rs_log, cn, "Select * From tblDateTimeChanges")

Call LoadAuthorizedPerson
End Sub


Public Sub LoadAuthorizedPerson()
Dim X As ListItem

While Not rs_user.EOF
    
    Set X = ListView1.ListItems.Add(, , rs_user.AbsolutePosition, 1, 1)
        X.SubItems(1) = rs_user.Fields("FullName")
        X.SubItems(2) = rs_user.Fields("Username")
        
    rs_user.MoveNext
Wend
End Sub

Public Sub LoadAuthorizedPersonLog()
Dim X As ListItem

ListView2.ListItems.Clear

While Not rs_log.EOF
    
    Set X = ListView2.ListItems.Add(, , rs_log.AbsolutePosition, 1, 2)
        X.SubItems(1) = rs_log.Fields("OldDate") & " " & rs_log.Fields("OldTime")
        X.SubItems(2) = rs_log.Fields("DateChanges") & " " & rs_log.Fields("TimeChanges")
        X.SubItems(3) = rs_log.Fields("Description")
    rs_log.MoveNext
Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rs_user = Nothing
Set rs_log = Nothing
End Sub

Private Sub Timer1_Timer()
pb.Value = pb.Value + 5

Form7.Height = 3615

If pb.Value > 99 Then

    rs_log.Requery

    rs_log.Filter = adFilterNone
    rs_log.Filter = "Username='" & ListView1.SelectedItem.SubItems(2) & "'"

    Call LoadAuthorizedPersonLog
    
    If ListView2.ListItems.Count > 0 Then
        Form7.Height = Form7.Height + Frame1.Height + 200
    End If
    

pb.Value = 0
Timer1.Enabled = False
Label1.Visible = False
End If
End Sub
