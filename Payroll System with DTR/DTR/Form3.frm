VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print DTR"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   4710
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   15
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   4455
   End
   Begin VB.Frame Frame4 
      Height          =   15
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   4455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Print"
      TabPicture(0)   =   "Form3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Can&cel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "O&k"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4215
         Begin VB.Frame Frame3 
            Height          =   15
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   3975
         End
         Begin VB.Frame Frame2 
            Height          =   15
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   3975
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   960
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "Form3.frx":001C
            Left            =   2760
            List            =   "Form3.frx":0032
            TabIndex        =   5
            Top             =   840
            Width           =   855
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form3.frx":004D
            Left            =   1200
            List            =   "Form3.frx":0057
            TabIndex        =   3
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Year:"
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
            Left            =   2520
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Month:"
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
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "To:"
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
            Left            =   2280
            TabIndex        =   4
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "From:"
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
            Left            =   480
            TabIndex        =   2
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Year:"
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
         Left            =   2640
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
   End
   Begin MSForms.ComboBox Combo5 
      Height          =   315
      Left            =   1710
      TabIndex        =   18
      Top             =   210
      Width           =   2895
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "5106;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
      Caption         =   "Employee No.:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsinfo As New ADODB.Recordset

Private Sub Combo1_KeyPress(KeyAscii As Integer)
'If Not KeyAscii = 49 And Not KeyAscii = 50 And Not KeyAscii = 8 Then KeyAscii = 0
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
'If Not KeyAscii = 49 And Not KeyAscii = 50 And Not KeyAscii = 8 Then KeyAscii = 0
KeyAscii = 0
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
'If KeyAscii < 65 And Not KeyAscii = 8 Then KeyAscii = 0
KeyAscii = 0
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim i As Integer

If Combo5.Text = "" Then MsgBox "Printing Error." & vbNewLine & vbNewLine & "Please select an Employee Number.", vbExclamation, "Administrator": Combo5.SetFocus: Exit Sub

If Year(Now) < Val(Combo4.Text) Then MsgBox "Printing Error." & vbNewLine & vbNewLine & "Invalid Year.", vbExclamation, "Administrator": Combo4.SetFocus: Exit Sub

If Combo1.Text = "" Then MsgBox "Printing Error." & vbNewLine & vbNewLine & "Type a specific day from 'From List'.", vbExclamation, "Administrator": Combo1.SetFocus: Exit Sub
If Combo2.Text = "" Then MsgBox "Printing Error." & vbNewLine & vbNewLine & "Type a specific day from 'To List'.", vbExclamation, "Administrator": Combo2.SetFocus: Exit Sub

If Combo1.Text = 1 And Combo2.Text = 1 Then MsgBox "Printing Error." & vbNewLine & vbNewLine & "Type a specific day from 'To List' atleast 15.", vbExclamation, "Administrator": Exit Sub
If Combo1.Text = 16 And Combo2.Text = 1 Then MsgBox "Printing Error." & vbNewLine & vbNewLine & "Type a specific day from 'To List' atleast 30.", vbExclamation, "Administrator": Exit Sub
If Combo1.Text = 16 And Combo2.Text = 15 Then MsgBox "Printing Error." & vbNewLine & vbNewLine & "Type a specific day from 'To List' atleast 30.", vbExclamation, "Administrator": Exit Sub

rsinfo.Requery

rsinfo.Find "EmployeeNo='" & Combo5.Text & "'"

    pview.fname = rsinfo.Fields("Surname") & ", " & rsinfo.Fields("Firstname") & " " & rsinfo.Fields("Middlename")


rs.Requery

rs.Filter = adFilterNone
rs.Filter = "Name='" & Combo5.Text & "'And Month='" & Combo3.Text & "' And Year=" & Combo4.Text & " And Day>=" & Combo1.Text & " And Day<=" & Combo2.Text

total = 0
'rs.MoveFirst

For i = 0 To 30

    tymrec(i) = i + 1
    
    tymrecAMIN(i) = ""
    tymrecPMIN(i) = ""
    tymrecAMOUT(i) = ""
    tymrecPMOUT(i) = ""
    tymrecTotalTardiness(i) = ""
    tymrecTotalTardinessH(i) = ""
    
    If rs.Fields("Day") = tymrec(i) Then
        tymrecAMIN(i) = rs.Fields("AMIn")
        tymrecPMIN(i) = rs.Fields("PMIn")
        tymrecAMOUT(i) = rs.Fields("AMOut")
        tymrecPMOUT(i) = rs.Fields("PMOut")
        tymrecTotalTardinessH(i) = rs.Fields("TTinHour")
        tymrecTotalTardiness(i) = rs.Fields("TTinMin")
        rs.MoveNext
    End If
    
    total = total + Val(tymrecTotalTardiness(i))
    totalh = totalh + Val(tymrecTotalTardinessH(i))
    
Next i

pview.pmonth = Combo3.Text & " " & Combo1.Text & " - " & Combo2.Text & ", " & Combo4.Text

Set DataReport1.DataSource = rs

DataReport1.Show vbModal
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub LoadMonth()
With Combo3
    .AddItem "January"
    .AddItem "February"
    .AddItem "March"
    .AddItem "April"
    .AddItem "May"
    .AddItem "June"
    .AddItem "July"
    .AddItem "August"
    .AddItem "September"
    .AddItem "October"
    .AddItem "November"
    .AddItem "December"
End With
End Sub

Private Sub Form_Load()
Call set_rec_getData(rs, cn, "Select * From qryComputed")
Call set_rec_getData(rsinfo, cn, "Select* From tblEmployeeInfo ")

LoadEmployeeNo
LoadMonth
LoadYear

Combo3.Text = Format(Date, "mmmm")
Combo4.Text = Year(Now)
End Sub

Private Sub LoadYear()
Dim i As Integer

For i = 2007 To 2050
    Combo4.AddItem i
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rs = Nothing
Set rsinfo = Nothing
End Sub

Private Sub LoadEmployeeNo()
Combo5.Clear

Dim rsemp As New ADODB.Recordset

Call set_rec_getData(rsemp, cn, "Select * From tblEmployeeInfo")

rsemp.Requery

While Not rsemp.EOF
    Combo5.AddItem rsemp.Fields("Name")
    rsemp.MoveNext
Wend
End Sub
