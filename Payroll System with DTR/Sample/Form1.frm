VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin lvButton.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Change"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Height          =   15
      Left            =   240
      TabIndex        =   20
      Top             =   1320
      Width           =   6975
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   19
      Text            =   "(Empty)"
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6360
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4920
      TabIndex        =   13
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Count"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   3360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6240
      Top             =   3240
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2990
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EmpNo"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Total Count"
         Object.Width           =   4939
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1215
      Left            =   315
      TabIndex        =   4
      Top             =   1440
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2143
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EmpNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   9701
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   0
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
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   7095
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      ToolTipText     =   "Click to search."
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Form1.frx":0452
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   840
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Year:"
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "To:"
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "From:"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin MSForms.ComboBox ComboBox2 
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   840
      Width           =   1815
      VariousPropertyBits=   209733659
      DisplayStyle    =   3
      Size            =   "3201;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      Caption         =   "Month:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Period:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   615
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4471;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Caption         =   "Search Employee No."
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call set_conn_getData(cn, App.Path & "\MainFiles.mdb", True, "xxx")
Call set_rec_getData(rs, cn, "Select * From tblInfo")
Call set_rec_getData(rst, cn, "Select * From tblRate")

Call LoadListview1(rs, ListView1)
Call Months
Text3.Text = Year(Now)
End Sub

Private Sub Label9_Change()
Dim xchange As New ADODB.Recordset

Call set_rec_getData(xchange, cn, "Select * From tblRate where Count=2")

If Label9.Caption = 12.01 Then
    
    While Not xchange.EOF
        
        xchange.Fields("count") = 0.5
        xchange.Update
    
        xchange.MoveNext
    
    Wend

End If
        
End Sub

Private Sub lvButtons_H2_Click()
Label3.Caption = ComboBox2.Text & " " & Text1.Text & " - " & Text2.Text & ", " & Text3.Text
ListView2.ListItems.Clear
Timer2.Enabled = True
Label8.Caption = "Generating..."
Text4.Visible = False
End Sub

Public Sub Months()
ComboBox2.Text = Format(Date, "mmmm")
With ComboBox2
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

Private Sub Timer2_Timer()
On Error GoTo emptylist

pb.Value = pb.Value + 1

If pb.Value > 99 Then

Dim x As ListItem
Dim totCount As Double
Dim i

i = 1

While i < ListView1.ListItems.Count + 1

ListView1.ListItems.Item(i).EnsureVisible
ListView1.ListItems.Item(i).Selected = True

rst.Requery

rst.Filter = adFilterNone
rst.Filter = "No='" & ListView1.SelectedItem.SubItems(1) & "' And Month='" & ComboBox2.Text & "' And Day>=" & Text1.Text & " And Day<=" & Text2.Text & " And Year=" & Text3.Text

Set x = ListView2.ListItems.Add(, , "", 1, 1)
        x.SubItems(1) = rst.Fields("No")

    Do Until rst.EOF
        totCount = totCount + rst.Fields("Count")
        rst.MoveNext
    Loop
        
        x.SubItems(2) = totCount

        totCount = 0
        
i = i + 1

Wend

pb.Value = 0
Timer2.Enabled = False
Label8.Caption = ""
End If

emptylist:
    If Err.Number Then ListView2.ListItems.Clear: pb.Value = 0: Timer2.Enabled = False: Text4.Visible = True: Label8.Caption = ""
End Sub

Private Sub Timer3_Timer()
Label9.Caption = Format(Time, "hh.mm")
End Sub
