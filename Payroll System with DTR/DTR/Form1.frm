VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Time Record (CEO)"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   8655
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "_"
      Height          =   375
      Left            =   7680
      TabIndex        =   37
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   8160
      Picture         =   "Form1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   495
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   7320
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Max             =   60
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O P T I O N S"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   25
      ToolTipText     =   "Click to view ""Options"""
      Top             =   7200
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "DTR"
      TabPicture(0)   =   "Form1.frx":16CC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(5)=   "Timer1"
      Tab(0).Control(6)=   "txtLog"
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(8)=   "Frame5"
      Tab(0).Control(9)=   "optIn"
      Tab(0).Control(10)=   "optOut"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Splash Screen"
      TabPicture(1)   =   "Form1.frx":16E8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label12(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label12(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label13"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Timer4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Timer6"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.Timer Timer6 
         Interval        =   100
         Left            =   1320
         Top             =   1440
      End
      Begin VB.Timer Timer4 
         Interval        =   2
         Left            =   5640
         Top             =   1560
      End
      Begin VB.OptionButton optOut 
         Caption         =   "Out"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69000
         TabIndex        =   22
         Top             =   3240
         Width           =   735
      End
      Begin VB.OptionButton optIn 
         Caption         =   "In "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70080
         TabIndex        =   21
         Top             =   3240
         Width           =   735
      End
      Begin VB.Frame Frame5 
         Caption         =   "Official Time"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74640
         TabIndex        =   20
         Top             =   2400
         Width           =   2415
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Rockwell"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Rockwell"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -67920
         TabIndex        =   19
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtLog 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -71760
         TabIndex        =   18
         Text            =   "Your Employee No."
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   -74160
         Top             =   1320
      End
      Begin VB.Frame Frame3 
         Caption         =   "Daily Record"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3345
         Left            =   -74760
         TabIndex        =   16
         Top             =   3480
         Width           =   8295
         Begin MSComctlLib.ListView ListView1 
            Height          =   2985
            Left            =   120
            TabIndex        =   17
            Top             =   210
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   5265
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
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   549
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Date"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "TimeInAM"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "TimeOutAM"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Hours & Mins (Undertime)"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "TimeInPM"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "TimeOutPM"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Hours & Mins (Undertime)"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Total Undertime (Hours & Mins)"
               Object.Width           =   6174
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   360
            Top             =   2760
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
                  Picture         =   "Form1.frx":1704
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Employee Information"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74640
         TabIndex        =   6
         Top             =   360
         Width           =   8055
         Begin VB.Timer Timer3 
            Enabled         =   0   'False
            Interval        =   2
            Left            =   960
            Top             =   1440
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   900
            Left            =   960
            Top             =   960
         End
         Begin VB.Frame Frame2 
            Height          =   15
            Left            =   2160
            TabIndex        =   7
            Top             =   1440
            Width           =   5775
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Employee Number:"
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
            Left            =   3960
            TabIndex        =   15
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblEmployeeNo 
            Caption         =   "Employee Number"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   14
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "MI"
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
            Left            =   7440
            TabIndex        =   13
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Firstname"
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
            Left            =   4800
            TabIndex        =   12
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Surname"
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
            Left            =   2160
            TabIndex        =   11
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label MI 
            Alignment       =   2  'Center
            Caption         =   "MI"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7440
            TabIndex        =   10
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Firstname 
            Alignment       =   2  'Center
            Caption         =   "Firstname"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   9
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Surname 
            Alignment       =   2  'Center
            Caption         =   "Surname"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   8
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1575
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Type Your Employee Number Here"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -72120
         TabIndex        =   3
         Top             =   2400
         Width           =   3975
         Begin VB.Frame Frame7 
            Height          =   15
            Left            =   360
            TabIndex        =   4
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "Log Status:"
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
            Left            =   360
            TabIndex        =   5
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   -68040
         TabIndex        =   2
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
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
         Left            =   4920
         TabIndex        =   39
         Top             =   6600
         Width           =   3495
      End
      Begin VB.Label Label13 
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
         TabIndex        =   38
         Top             =   6600
         Width           =   5415
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RICHIE S. MERCADEROS"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   36
         Top             =   5760
         Width           =   3015
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GERIVE L. MARATA"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   35
         Top             =   5520
         Width           =   3015
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RONELL S. DOROTAYO"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   34
         Top             =   5280
         Width           =   3015
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "JELYN L. MARATA"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   33
         Top             =   5040
         Width           =   3015
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Created by:"
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
         Height          =   255
         Left            =   4800
         TabIndex        =   32
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   5175
         Left            =   240
         Picture         =   "Form1.frx":1B56
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   8175
      End
      Begin VB.Label Label1 
         Caption         =   "Log Status:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71640
         TabIndex        =   23
         Top             =   3360
         Width           =   1695
      End
   End
   Begin VB.Label Label10 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   7080
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
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
      Left            =   210
      TabIndex        =   0
      Top             =   7680
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Dim db As DAO.Database
'-------------------------------
Dim rs_emp As DAO.Recordset
Dim rs_rec As DAO.Recordset

Dim dtes As String
Dim countlog As Single

Public sCount As Integer

Dim srateperh As Double
Dim srateperm As Double

Private Sub cmdClear_Click()
txtLog.SetFocus
Call BackNormalView
End Sub

Private Sub Command1_Click()
On Error Resume Next
If txtLog.Text = "Your Employee No." Or txtLog.Text = "" Then MsgBox "Please type your Employee Number.", vbInformation, "Administrator": txtLog.SetFocus: Exit Sub
If optIn.Value = False And optOut.Value = False Then MsgBox "Please select a log status below.", vbInformation, "Administrator":  Exit Sub

Set rs_emp = db.OpenRecordset("Select * From tblEmployeeInfo where EmployeeNo='" & txtLog.Text & "'")

With rs_emp
    .Requery
    .FindFirst "EmployeeNo='" & txtLog.Text & "'"
    
    If .NoMatch = False Then
        
        lblEmployeeNo.Caption = .Fields("EmployeeNo")
        Surname.Caption = .Fields("Surname")
        Firstname.Caption = .Fields("Firstname")
        MI.Caption = Mid(.Fields("Middlename"), 1, 1)
        
        Image1.Picture = LoadPicture(.Fields("Picture"))
                   
        Form4.Show vbModal
        
        Else
            MsgBox "Unregistered Employee Number.", vbExclamation, "Administrator"
            Exit Sub
    End If
End With
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PopupMenu Form2.File
End Sub

Private Sub Command3_Click()
MsgBox "DTR still running.", vbInformation
Timer3.Enabled = True
End Sub

Private Sub Command4_Click()
Me.WindowState = 1
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\MasterFile.mdb", False, False, ";pwd=xxx")

Call set_conn_getData(cn, App.Path & "\MasterFile.mdb", True, "xxx")

Call set_rec_getData(time_rec, cn, "Select * From qryComputed")

If Val(lblTime.Caption) >= 1 Or Val(lblTime.Caption) <= 10 Then optIn.Value = True

Me.Show

trns = 100
SetTrans Me, trns

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static bRunning As Boolean
  Dim lMsg As Long
   'vbModal
  lMsg = X / Screen.TwipsPerPixelX
  If Not (bRunning) Then 'avoid cascades
    bRunning = True
    Select Case lMsg
      Case WM_LBUTTONDBLCLK:
      Case WM_LBUTTONDOWN:
      Case WM_LBUTTONUP:
      Case WM_RBUTTONDBLCLK:
      Case WM_RBUTTONDOWN:
      Case WM_RBUTTONUP: PopupMenu Form2.mnuProperty
            
   End Select
   bRunning = False
   End If
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i

For i = 0 To 3
    Label12(i).ForeColor = vbWhite
Next i
End Sub

Private Sub Label12_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0: Label12(Index).ForeColor = vbBlue: Label12(1).ForeColor = vbWhite: Label12(2).ForeColor = vbWhite: Label12(3).ForeColor = vbWhite
    Case 1: Label12(Index).ForeColor = vbBlue: Label12(0).ForeColor = vbWhite: Label12(2).ForeColor = vbWhite: Label12(3).ForeColor = vbWhite
    Case 2: Label12(Index).ForeColor = vbBlue: Label12(0).ForeColor = vbWhite: Label12(1).ForeColor = vbWhite: Label12(3).ForeColor = vbWhite
    Case 3: Label12(Index).ForeColor = vbBlue: Label12(0).ForeColor = vbWhite: Label12(1).ForeColor = vbWhite: Label12(2).ForeColor = vbWhite
End Select
End Sub

Private Sub Label14_Change()
Dim sChange As DAO.Recordset

Set sChange = db.OpenRecordset("Select * From qryComputed where CDate='" & Format(Date, "mm/dd/yyyy") & "' and TimeOutPm<TimeInPm")

If Label14.Caption = "11:59:59 pm" Then
While Not sChange.EOF
    sChange.Edit
        sChange.Fields("PTHour") = 4
        sChange.Fields("PTMin") = 0
        sChange.Fields("TTinHour") = Val(sChange.Fields("ATHour")) + 4
        sChange.Fields("TTinMin") = Val(sChange.Fields("ATMin")) + 0
        'rs_rec.Fields("Deduction") = Val(rs_rec.Fields("TTinHour"))*
        sChange.Fields("LastStatus") = "OUT"
        sChange.Fields("Count") = 4
    sChange.Update
    sChange.MoveNext
Wend
End If

End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTime.ToolTipText = Format(lblTime.Caption, "hh:mm:ss AM/PM")
End Sub

Private Sub Timer1_Timer()
lblTime.Caption = Format(Time, "hh.mm")
Label9.Caption = Format(Time, "hh:mm:ss am/pm")
Label7.Caption = Format(Time, "hh:mm am/pm")
Frame3.Caption = Format(Date, "mmmm dd, yyyy")
Label3.Caption = Format(Date, "mmmm dd, yyyy")
End Sub

Private Sub Timer2_Timer()
On Error Resume Next

sCount = sCount - 1

Label10.Caption = "You have " & sCount & " seconds to review the record."

pb.Value = pb.Value - 1

If sCount = 0 Then sCount = 0: Call cmdClear_Click: pb.Visible = False
End Sub

Private Sub Timer3_Timer()
trns = trns - 5

If trns = 0 Then
    Timer3.Enabled = False
    trns = 100
    Me.Hide
    SysTrayData = SysTrayInit(Me.Caption, Me, Me.Icon)
End If

SetTrans Me, trns

End Sub


Private Sub Timer5_Timer()
Dim i

For i = 0 To 3
    Label12(i).Enabled = True
Next i
End Sub

Private Sub Timer6_Timer()
Label13.Caption = "Today: (" & Format(Date, "dddd") & ") " & Format(Date, "mmmm dd, yyyy")
Label14.Caption = Format(Time, "hh:mm:ss am/pm")
End Sub

Private Sub txtLog_Change()
Surname.Caption = "Surname"
Firstname.Caption = "Firstname"
MI.Caption = "MI"
lblEmployeeNo.Caption = "Employee Number"
Image1.Picture = LoadPicture("")
ListView1.ListItems.Clear
Timer2.Enabled = False
Label10.Caption = ""
pb.Visible = False
optIn.Value = False
optOut.Value = False
End Sub

Private Sub txtLog_GotFocus()
With txtLog
    .SelStart = 0
    .SelLength = Len(txtLog.Text)
End With
End Sub


Public Sub LoadRecord()
On Error Resume Next
Dim X As ListItem
Dim load_rec As DAO.Recordset

dtes = Format(Date, "mmmm")

Set load_rec = db.OpenRecordset("Select * From qryComputed where EmployeeNo='" & txtLog.Text & "'and Month ='" & dtes & "'and Year='" & Year(Now) & "'")

While Not load_rec.EOF
    Set X = ListView1.ListItems.Add(, , "", 1, 1)

    X.SubItems(1) = load_rec.Fields("CDate")
    X.SubItems(2) = load_rec.Fields("AMIn")
    X.SubItems(3) = load_rec.Fields("AMOut")
    X.SubItems(4) = load_rec.Fields("ATHour") & " hr(s) " & load_rec.Fields("ATMin") & " min(s)"
    X.SubItems(5) = load_rec.Fields("PMIn")
    X.SubItems(6) = load_rec.Fields("PMOut")
    X.SubItems(7) = load_rec.Fields("PTHour") & " hr(s) " & load_rec.Fields("PTMin") & " min(s)"
    X.SubItems(8) = load_rec.Fields("TTinHour") & " hr(s) " & load_rec.Fields("TTinMin") & " min(s)"
    
    load_rec.MoveNext
Wend
End Sub

Sub onFocus()
On Error Resume Next
Dim X As ListItem

Set X = ListView1.FindItem(Format(Date, "mm/dd/yyyy"), lvwSubItem + lvwText, lvwPartial, lvwPartial)

If Not X Is Nothing Then
    X.EnsureVisible
    X.Selected = True
End If
End Sub

Public Sub BackNormalView()
Surname.Caption = "Surname"
Firstname.Caption = "Firstname"
MI.Caption = "MI"
lblEmployeeNo.Caption = "Employee Number"
Image1.Picture = LoadPicture("")
txtLog.Text = "Your Employee No."
ListView1.ListItems.Clear
Timer2.Enabled = False
Label10.Caption = ""
End Sub

Private Sub txtLog_LostFocus()
If txtLog.Text = "" Then txtLog.Text = "Your Employee No."
End Sub

Public Sub FieldChange()
Set rs_rec = db.OpenRecordset("Select * From qryComputed where CDate='" & (Format(Date, "mm/dd/yyyy") - 1) & "' and TimeOutPm<TimeInPm")

While Not rs_rec.EOF
    rs_rec.Edit
        rs_rec.Fields("PTHour") = 4
        rs_rec.Fields("PTMin") = 0
        rs_rec.Fields("TTinHour") = Val(rs_rec.Fields("ATHour")) + 4
        rs_rec.Fields("TTinMin") = Val(rs_rec.Fields("ATMin")) + 0
        'rs_rec.Fields("Deduction") = Val(rs_rec.Fields("TTinHour"))*
        rs_rec.Fields("LastStatus") = "OUT"
        rs_rec.Fields("Count") = 4
    rs_rec.Update
    rs_rec.MoveNext
Wend

End Sub
