VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8565
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6120
      TabIndex        =   64
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6120
      TabIndex        =   63
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   62
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   61
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   60
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   7920
      TabIndex        =   59
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   3600
      TabIndex        =   58
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6960
      TabIndex        =   18
      Top             =   6825
      Width           =   1530
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5280
      TabIndex        =   17
      Top             =   6825
      Width           =   1530
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      HideSelection   =   0   'False
      MaxLength       =   8
      Format          =   "hh:mm AM/PM"
      Mask            =   "##:## am"
      PromptChar      =   "_"
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   420
   End
   Begin VB.Frame Frame3 
      Height          =   15
      Left            =   120
      TabIndex        =   9
      Top             =   630
      Width           =   8265
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20643843
      CurrentDate     =   39433
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
      Left            =   1680
      TabIndex        =   0
      Text            =   "0"
      Top             =   1080
      Width           =   2295
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      AutoTab         =   -1  'True
      HideSelection   =   0   'False
      MaxLength       =   8
      Format          =   "hh:mm AM/PM"
      Mask            =   "##:## am"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text3 
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
      Left            =   6120
      TabIndex        =   1
      Text            =   "0"
      Top             =   1080
      Width           =   2295
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      HideSelection   =   0   'False
      MaxLength       =   8
      Format          =   "hh:mm AM/PM"
      Mask            =   "##:## am"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      AutoTab         =   -1  'True
      HideSelection   =   0   'False
      MaxLength       =   8
      Format          =   "hh:mm AM/PM"
      Mask            =   "##:## am"
      PromptChar      =   "_"
   End
   Begin VB.Label Label49 
      Caption         =   "Empolyee Number:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   66
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Label48 
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
      Left            =   2640
      TabIndex        =   65
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   2340
      TabIndex        =   57
      Top             =   2625
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   2340
      TabIndex        =   56
      Top             =   2940
      Width           =   1140
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   2340
      TabIndex        =   55
      Top             =   3675
      Width           =   1140
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   2340
      TabIndex        =   54
      Top             =   3990
      Width           =   1140
   End
   Begin VB.Label Label9 
      Caption         =   "AM Late:"
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
      Left            =   795
      TabIndex        =   53
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "AM Undertime:"
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
      Left            =   765
      TabIndex        =   52
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "PM Late:"
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
      Left            =   4860
      TabIndex        =   51
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "PM Undertime:"
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
      Left            =   4860
      TabIndex        =   50
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Tardiness AM"
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
      Left            =   1395
      TabIndex        =   49
      Top             =   2100
      Width           =   1485
   End
   Begin VB.Line Line19 
      X1              =   2970
      X2              =   3495
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Line Line2 
      X1              =   795
      X2              =   1395
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label Label16 
      Caption         =   "Hours:"
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
      Left            =   1455
      TabIndex        =   48
      Top             =   2640
      Width           =   690
   End
   Begin VB.Label Label17 
      Caption         =   "Minutes:"
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
      Left            =   1185
      TabIndex        =   47
      Top             =   2940
      Width           =   1005
   End
   Begin VB.Label Label18 
      Caption         =   "Hours:"
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
      Left            =   1455
      TabIndex        =   46
      Top             =   3675
      Width           =   690
   End
   Begin VB.Label Label19 
      Caption         =   "Minutes:"
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
      Left            =   1185
      TabIndex        =   45
      Top             =   3990
      Width           =   1005
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   5490
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Line Line4 
      X1              =   7170
      X2              =   7755
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Tardiness PM"
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
      Left            =   5595
      TabIndex        =   44
      Top             =   2100
      Width           =   1485
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   6600
      TabIndex        =   43
      Top             =   2640
      Width           =   1140
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   6600
      TabIndex        =   42
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Label Label23 
      Caption         =   "Hours:"
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
      Left            =   5775
      TabIndex        =   41
      Top             =   2625
      Width           =   690
   End
   Begin VB.Label Label24 
      Caption         =   "Minutes:"
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
      Left            =   5490
      TabIndex        =   40
      Top             =   2940
      Width           =   1005
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   6645
      TabIndex        =   39
      Top             =   3720
      Width           =   1140
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   6645
      TabIndex        =   38
      Top             =   4080
      Width           =   1140
   End
   Begin VB.Label Label27 
      Caption         =   "Hours:"
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
      Left            =   5775
      TabIndex        =   37
      Top             =   3675
      Width           =   690
   End
   Begin VB.Label Label28 
      Caption         =   "Minutes:"
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
      Left            =   5490
      TabIndex        =   36
      Top             =   3990
      Width           =   1005
   End
   Begin VB.Line Line5 
      X1              =   3510
      X2              =   5175
      Y1              =   5775
      Y2              =   5775
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "Total in All"
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
      Left            =   1830
      TabIndex        =   35
      Top             =   5670
      Width           =   1695
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   2775
      TabIndex        =   34
      Top             =   6000
      Width           =   1140
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   2760
      TabIndex        =   33
      Top             =   6360
      Width           =   1140
   End
   Begin VB.Label Label33 
      Caption         =   "Hours:"
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
      Left            =   1935
      TabIndex        =   32
      Top             =   5985
      Width           =   690
   End
   Begin VB.Label Label34 
      Caption         =   "Minutes:"
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
      Left            =   1620
      TabIndex        =   31
      Top             =   6300
      Width           =   1005
   End
   Begin VB.Label Label37 
      Caption         =   "Total:"
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
      Left            =   765
      TabIndex        =   30
      Top             =   4515
      Width           =   1215
   End
   Begin VB.Label Label38 
      Caption         =   "Minutes:"
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
      Left            =   1185
      TabIndex        =   29
      Top             =   5145
      Width           =   1005
   End
   Begin VB.Label Label39 
      Caption         =   "Hours:"
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
      Left            =   1455
      TabIndex        =   28
      Top             =   4830
      Width           =   690
   End
   Begin VB.Label Label40 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      TabIndex        =   27
      Top             =   4800
      Width           =   1140
   End
   Begin VB.Label Label41 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   2295
      TabIndex        =   26
      Top             =   5160
      Width           =   1140
   End
   Begin VB.Line Line13 
      X1              =   765
      X2              =   3495
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line14 
      X1              =   240
      X2              =   4860
      Y1              =   5565
      Y2              =   5565
   End
   Begin VB.Label Label42 
      Caption         =   "Total:"
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
      Left            =   4860
      TabIndex        =   25
      Top             =   4515
      Width           =   1215
   End
   Begin VB.Label Label43 
      Caption         =   "Minutes:"
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
      Left            =   5490
      TabIndex        =   24
      Top             =   5145
      Width           =   1005
   End
   Begin VB.Label Label44 
      Caption         =   "Hours:"
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
      Left            =   5775
      TabIndex        =   23
      Top             =   4830
      Width           =   690
   End
   Begin VB.Label Label45 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   6645
      TabIndex        =   22
      Top             =   4800
      Width           =   1140
   End
   Begin VB.Label Label46 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   6645
      TabIndex        =   21
      Top             =   5160
      Width           =   1140
   End
   Begin VB.Line Line15 
      X1              =   4860
      X2              =   7800
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line16 
      X1              =   4860
      X2              =   8535
      Y1              =   5565
      Y2              =   5565
   End
   Begin VB.Line Line6 
      X1              =   1830
      X2              =   1200
      Y1              =   5775
      Y2              =   5775
   End
   Begin VB.Line Line17 
      X1              =   240
      X2              =   8535
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line18 
      X1              =   4335
      X2              =   4335
      Y1              =   840
      Y2              =   5565
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "Deductions"
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
      Left            =   5205
      TabIndex        =   20
      Top             =   5655
      Width           =   1695
   End
   Begin VB.Line Line20 
      X1              =   6885
      X2              =   7485
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label47 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   5535
      TabIndex        =   19
      Top             =   6120
      Width           =   1140
   End
   Begin VB.Line Line21 
      X1              =   3975
      X2              =   4215
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line22 
      X1              =   3975
      X2              =   4215
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line23 
      X1              =   4215
      X2              =   4215
      Y1              =   6120
      Y2              =   6480
   End
   Begin VB.Line Line24 
      X1              =   4215
      X2              =   5415
      Y1              =   6240
      Y2              =   6240
   End
   Begin MSForms.ComboBox Text5 
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   120
      Width           =   3015
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "5318;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line12 
      X1              =   4440
      X2              =   8400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Caption         =   "P.M."
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
      Left            =   6270
      TabIndex        =   11
      Top             =   735
      Width           =   540
   End
   Begin VB.Line Line11 
      X1              =   6825
      X2              =   8460
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line10 
      X1              =   4620
      X2              =   6195
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line9 
      X1              =   105
      X2              =   4515
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      Caption         =   "A.M."
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
      Left            =   1755
      TabIndex        =   10
      Top             =   735
      Width           =   540
   End
   Begin VB.Line Line8 
      X1              =   2310
      X2              =   4620
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line7 
      X1              =   105
      X2              =   1680
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Date:"
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
      Left            =   5145
      TabIndex        =   7
      Top             =   210
      Width           =   615
   End
   Begin VB.Label Label13 
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
      Left            =   210
      TabIndex        =   6
      Top             =   210
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Time Out PM:"
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
      Left            =   4545
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Time In PM:"
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
      Left            =   4515
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Time Out AM:"
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
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Time In AM:"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Timein24HAm As Double
Dim Timein24HPm As Double
Dim TimeOut24HAm As Double
Dim TimeOut24HPm As Double

Dim numHour As Double

Dim addentry As New ADODB.Recordset
Dim LoadEmpNo As New ADODB.Recordset
Dim emp_sal As New ADODB.Recordset

Dim rateperh As Double
Dim rateperm As Double

Public add_state As Boolean

Dim empno As String
Dim desig As String

Private Sub Command1_Click()
On Error Resume Next
With time_rec
    If add_state = True Then .AddNew
        .Fields("EmployeeNo") = Label48.Caption
        .Fields("Name") = Text5.Text
        .Fields("CDate") = Format(DTPicker1.Value, "mm/dd/yyyy")
        .Fields("Month") = Format(DTPicker1.Value, "mmmm")
        .Fields("Day") = Format(DTPicker1.Value, "dd")
        .Fields("Year") = Format(DTPicker1.Value, "yyyy")
        .Fields("AMIn") = Text2.Text
        .Fields("AMOut") = Text4.Text
        .Fields("PMIn") = Text6.Text
        .Fields("PMOut") = Text7.Text
        
        .Fields("TimeInAm") = Format(Text2.Text, "hh.mm")
        .Fields("TimeOutAm") = Format(Text4.Text, "hh.mm")
        .Fields("TimeInPm") = Format(Text6.Text, "hh.mm")
        .Fields("TimeOutPm") = Format(Text7.Text, "hh.mm")
        
        .Fields("TTinHour") = Label31.Caption
        .Fields("TTinMin") = Label32.Caption
        
        If Text2.Text = "" And Text4.Text = "" And Text6.Text <> "" And Text7.Text <> "" Then
            .Fields("ATHour") = 4
            .Fields("ATMin") = 0
            .Fields("PTHour") = Label45.Caption
            .Fields("PTMin") = Label46.Caption
            .Fields("NumDayWork") = 0.5
            .Fields("Count") = 4
            
        ElseIf Text6.Text = "" And Text7.Text = "" And Text2.Text <> "" And Text4.Text <> "" Then
            .Fields("ATHour") = Label40.Caption
            .Fields("ATMin") = Label41.Caption
            .Fields("PTHour") = 4
            .Fields("PTMin") = 0
            .Fields("NumDayWork") = 0.5
            .Fields("LastStatus") = "OUT"
            .Fields("Count") = 2
            
        Else
            .Fields("ATHour") = Label40.Caption
            .Fields("ATMin") = Label41.Caption
            .Fields("PTHour") = Label45.Caption
            .Fields("PTMin") = Label46.Caption
            .Fields("NumDayWork") = 1
            .Fields("LastStatus") = "OUT"
            .Fields("Count") = 4
            
        End If
        
        .Fields("Deduction") = Label47.Caption
                         
        .Update
End With

If add_state = True Then
    MsgBox "Adding of employee time record has been successfull.", vbInformation, "Save Complete"
    Dim rep As Integer
    rep = MsgBox("Do you want to add another employee time record?", vbQuestion + vbYesNo, "Daily Time Record")
    If rep = vbYes Then
            
        Text2.Text = ""
        Text4.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        
        Label5.Caption = 0
        Label6.Caption = 0
        Label7.Caption = 0
        Label8.Caption = 0
        Label40.Caption = 0
        Label41.Caption = 0
        Label45.Caption = 0
        Label46.Caption = 0
        Label25.Caption = 0
        Label26.Caption = 0
        Label21.Caption = 0
        Label22.Caption = 0
        Label31.Caption = 0
        Label32.Caption = 0
        Label47.Caption = 0

        Text5.SetFocus
        Text5.Locked = False
        Command1.Enabled = False
        
        Form10.ListView1.ListItems.Clear
        Form10.LoadRecord
    Else
        Form10.ListView1.ListItems.Clear
        Form10.LoadRecord
        Unload Me
    End If
    rep = 0
Else
    MsgBox "Changes in record has been successfully saved.", vbInformation, "Daily Time Record"
    Dim pos As Long
    
    pos = time_rec.AbsolutePosition
    Form10.ListView1.ListItems.Clear
    Form10.LoadRecord
    time_rec.AbsolutePosition = pos
    
    Form10.ListView1.ListItems.Item(pos).EnsureVisible
    Form10.ListView1.ListItems.Item(pos).Selected = True
    
    pos = 0
    Unload Me
End If
      
      
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Call set_rec_getData(addentry, cn, "Select * From tblTimeRecord")
Call set_rec_getData(LoadEmpNo, cn, "Select * From tblEmployeeInfo")
Call set_rec_getData(emp_sal, cn, "Select * From tblDesignation")

If add_state = True Then
    Me.Caption = "Add Employee Time Record"
    DTPicker1.Value = Date
Else
    Me.Caption = "Edit Employee Time Record"
    Text5.Locked = True
    DTPicker1.Enabled = False
    
    With time_rec
        
        Text5.Text = .Fields("Name")
        
        Label48.Caption = .Fields("EmployeeNo")
        DTPicker1.Value = .Fields("CDate")
                
        Text2.Text = .Fields("AMIn")
        Text4.Text = .Fields("AMOut")
        Text6.Text = .Fields("PMIn")
        Text7.Text = .Fields("PMOut")

        Label40.Caption = .Fields("ATHour")
        Label41.Caption = .Fields("ATMin")
        
        Label45.Caption = .Fields("PTHour")
        Label46.Caption = .Fields("PTMin")
        
        Label31.Caption = .Fields("TTinHour")
        Label32.Caption = .Fields("TTinMin")
        
        
    End With
    
    If Not Text2.Text = "" Then Text2.Locked = True
    If Not Text4.Text = "" Then Text4.Locked = True
    If Not Text6.Text = "" Then Text6.Locked = True
    If Not Text7.Text = "" Then Text7.Locked = True

End If

Call LoadEmpNos
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set addentry = Nothing
Set LoadEmpNo = Nothing
'Set time_rec = Nothing
End Sub

Private Sub Command3_Click()
If Label7.Caption = 4 Then
    Label40.Caption = 4
    Label41.Caption = 0
Else
    Label40.Caption = Val(Label5.Caption) + Val(Label7.Caption)
    Label41.Caption = Val(Label6.Caption) + Val(Label8.Caption)
    
    If Val(Label41.Caption) > 59 Then
        Label40.Caption = Val(Label40.Caption) + 1
        Label41.Caption = Val(Label41.Caption) - 60
    End If
End If
End Sub

Private Sub Command4_Click()
If Val(Label25.Caption) = 4 Then
    Label45.Caption = 4
    Label46.Caption = 0
Else
    Label45.Caption = Val(Label21.Caption) + Val(Label25.Caption)
    Label46.Caption = Val(Label22.Caption) + Val(Label26.Caption)
    
    If Val(Label46.Caption) > 59 Then
        Label45.Caption = Val(Label45.Caption) + 1
        Label46.Caption = (Val(Label22.Caption) + Val(Label26.Caption)) - 60
    End If
End If
End Sub

Private Sub Command5_Click()

Call Command3_Click
Call Command4_Click

If Text2.Text = "" And Text4.Text <> "" Then MsgBox "Invalid procedure. Enter time in AM first!", vbExclamation, "Daily Time Record": Exit Sub
If Text6.Text = "" And Text7.Text <> "" Then MsgBox "Invalid procedure. Enter time in PM first!", vbExclamation, "Daily Time Record": Exit Sub

If Text2.Text = "" And Text4.Text = "" Then Label40.Caption = 4: Label41.Caption = 0 ': Exit Sub
If Text6.Text = "" And Text7.Text = "" Then Label45.Caption = 4: Label46.Caption = 0 ': Exit Sub

Label31.Caption = Val(Label40.Caption) + Val(Label45.Caption)
Label32.Caption = Val(Label41.Caption) + Val(Label46.Caption)

If Val(Label32.Caption) > 59 Then
    Label31.Caption = Val(Label31.Caption) + 1
    Label32.Caption = (Val(Label41.Caption) + Val(Label46.Caption)) - 60
End If

Label47.Caption = FormatNumber((Val(Label31.Caption) * rateperh) + (Val(Label32.Caption) * rateperm), 2)

Command1.Enabled = True
Text5.Locked = True
End Sub

Private Sub Label48_Change()
On Error Resume Next

desig = ""

LoadEmpNo.Requery

LoadEmpNo.Find "EmployeeNo='" & Trim(Label48.Caption) & "'"

desig = LoadEmpNo.Fields("Designation")
    
    If desig = "Admin Aide" Then
        
        emp_sal.Requery
        emp_sal.Find "Designation='" & desig & "'"
        
        rateperh = Val(emp_sal.Fields("SalaryperDay")) / 8
        rateperm = rateperh / 60
    
    End If

End Sub

Private Sub Text2_Change()
On Error Resume Next
'------------------------
Dim hours As Integer
Dim min As Integer
'------------------------

Label5.Caption = 0
Label6.Caption = 0

Label31.Caption = 0
Label32.Caption = 0

If Text5.Text = "" Then MsgBox "Type employee name first.", vbExclamation, "System Required": Text5.SetFocus: Exit Sub

Timein24HAm = FormatNumber(Format(Trim(Text2.Text), "hh.mm"), 2)

hours = 0
min = 0

If Timein24HAm > 11.59 Then MsgBox "Invalid time in.", vbExclamation, "System Requiered": MaskEdBox1.SetFocus: Exit Sub

Select Case Timein24HAm
    
    Case 8.01 To 8.59
        
        hours = 0
        min = (Timein24HAm - 8) * 100
        
        Label5.Caption = hours
        Label6.Caption = min
        
    Case 9
        
        hours = 1
        min = 0
        
        Label5.Caption = hours
        Label6.Caption = min
        
    Case 9.01 To 9.59
        
        hours = 1
        min = (Timein24HAm - 9) * 100
        
        Label5.Caption = hours
        Label6.Caption = min
        
    Case 10
        
        hours = 2
        min = 0
        
        Label5.Caption = hours
        Label6.Caption = min
        
    Case 10.01 To 10.59
        
        hours = 2
        min = (Timein24HAm - 10) * 100
        
        Label5.Caption = hours
        Label6.Caption = min
        
    Case 11
        
        hours = 3
        min = 0
        
        Label5.Caption = hours
        Label6.Caption = min
        
    Case 11.01 To 11.59
        
        hours = 4
        min = 0
        
        Label5.Caption = hours
        Label6.Caption = min
        
    Case Else
    
        hours = 0
        min = 0
        
        Label5.Caption = hours
        Label6.Caption = min
        
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_Change()
On Error Resume Next
'-------------------------
Dim hours As Integer
Dim min As Integer
'-------------------------

Label7.Caption = 0
Label8.Caption = 0

Label31.Caption = 0
Label32.Caption = 0

If Text5.Text = "" Then MsgBox "Type employee name first.", vbExclamation, "System Required": Text5.SetFocus: Exit Sub

TimeOut24HAm = FormatNumber(Format(Trim(Text4.Text), "hh.mm"), 2)

If TimeOut24HAm > 12.59 Then MsgBox "Invalid time out.", vbExclamation, "System Requiered": MaskEdBox2.SetFocus: Exit Sub

hours = 0
min = 0

Select Case TimeOut24HAm
    
    Case 8.01 To 8.59
            
        If TimeOut24HAm = Timein24HAm Then hours = 4: min = 0
        If TimeOut24HAm < Timein24HAm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        hours = 4
        min = 0
        
        Label7.Caption = hours
        Label8.Caption = min
        
    Case 9
        
        If TimeOut24HAm = Timein24HAm Then hours = 4: min = 0
        If TimeOut24HAm < Timein24HAm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HAm - Timein24HAm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 3
            min = 0
        End If
        
        Label7.Caption = hours
        Label8.Caption = min
        
    Case 9.01 To 9.59
            
        If TimeOut24HAm = Timein24HAm Then hours = 4: min = 0
        If TimeOut24HAm < Timein24HAm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HAm - Timein24HAm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 2
            min = (9.6 - TimeOut24HAm) * 100
        End If
        
        Label7.Caption = hours
        Label8.Caption = min
            
    Case 10
        
        If TimeOut24HAm = Timein24HAm Then hours = 4: min = 0
        If TimeOut24HAm < Timein24HAm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HAm - Timein24HAm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 2
            min = 0
        End If
        
        Label7.Caption = hours
        Label8.Caption = min
        
    Case 10.01 To 10.59
            
        If TimeOut24HAm = Timein24HAm Then hours = 4: min = 0
        If TimeOut24HAm < Timein24HAm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HAm - Timein24HAm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 1
            min = (10.6 - TimeOut24HAm) * 100
        End If
        
        Label7.Caption = hours
        Label8.Caption = min
    
    Case 11
        
        If TimeOut24HAm = Timein24HAm Then hours = 4: min = 0
        If TimeOut24HAm < Timein24HAm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HAm - Timein24HAm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 1
            min = 0
        End If
        
        Label7.Caption = hours
        Label8.Caption = min
        
    Case 11.01 To 11.59
            
        If TimeOut24HAm = Timein24HAm Then hours = 4: min = 0
        If TimeOut24HAm < Timein24HAm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HAm - Timein24HAm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 0
            min = (11.6 - TimeOut24HAm) * 100
        End If
        
        Label7.Caption = hours
        Label8.Caption = min
End Select
    
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text6.SetFocus
End Sub

Private Sub Text5_GotFocus()
With Text5
    .SelStart = 0
    .SelLength = Len(Text5.Text)
End With
End Sub

Private Sub Text6_Change()
On Error Resume Next
'------------------------
Dim hours As Integer
Dim min As Integer
'------------------------

Label21.Caption = 0
Label22.Caption = 0

Label31.Caption = 0
Label32.Caption = 0

If Text5.Text = "" Then MsgBox "Type employee name first.", vbExclamation, "System Required": Text5.SetFocus: Exit Sub

Timein24HPm = FormatNumber(Format(Trim(Text6.Text), "hh.mm"), 2)

If Timein24HPm > 16.59 Then MsgBox "Invalid time in.", vbExclamation, "System Requiered": MaskEdBox3.SetFocus: Exit Sub

hours = 0
min = 0

Select Case Timein24HPm
    
    Case 13.01 To 13.59
        
        hours = 0
        min = (Timein24HPm - 13) * 100
        
        Label21.Caption = hours
        Label22.Caption = min
        
    Case 14
        
        hours = 1
        min = 0
        
        Label21.Caption = hours
        Label22.Caption = min
    
    Case 14.01 To 14.59
        
        hours = 1
        min = (Timein24HPm - 14) * 100
        
        Label21.Caption = hours
        Label22.Caption = min
    
    Case 15
        
        hours = 2
        min = 0
        
        Label21.Caption = hours
        Label22.Caption = min
        
    Case 15.01 To 15.59
        
        hours = 2
        min = (Timein24HAm - 15) * 100
        
        Label21.Caption = hours
        Label22.Caption = min
    
    Case 16
        
        hours = 3
        min = 0
        
        Label21.Caption = hours
        Label22.Caption = min
        
    Case 16.01 To 11.69
        
        hours = 4
        min = 0
        
        Label21.Caption = hours
        Label22.Caption = min
        
    Case Else
    
        hours = 0
        min = 0
        
        Label21.Caption = hours
        Label22.Caption = min
        
End Select
End Sub

Private Sub Text7_Change()
On Error Resume Next
'-------------------------
Dim hours As Integer
Dim min As Integer
'-------------------------

Label25.Caption = 0
Label26.Caption = 0

Label31.Caption = 0
Label32.Caption = 0

If Text5.Text = "" Then MsgBox "Type employee name first.", vbExclamation, "System Required": Text5.SetFocus: Exit Sub

TimeOut24HPm = FormatNumber(Format(Trim(Text7.Text), "hh.mm"), 2)

If TimeOut24HPm > 20.59 Then MsgBox "Invalid time out.", vbExclamation, "System Requiered": MaskEdBox1.SetFocus: Exit Sub

hours = 0
min = 0

Select Case TimeOut24HPm
    
    Case 13.01 To 13.59
            
        If TimeOut24HPm = Timein24HPm Then hours = 4: min = 0
        If TimeOut24HPm < Timein24HPm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        hours = 4
        min = 0
        
        Label25.Caption = hours
        Label26.Caption = min
        
    Case 14
        
        If TimeOut24HPm = Timein24HPm Then hours = 4: min = 0
        If TimeOut24HPm < Timein24HPm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HPm - Timein24HPm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 3
            min = 0
        End If
        
        Label25.Caption = hours
        Label26.Caption = min
        
    Case 14.01 To 14.59
            
        If TimeOut24HPm = Timein24HPm Then hours = 4: min = 0
        If TimeOut24HPm < Timein24HPm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HPm - Timein24HPm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 2
            min = (14.6 - TimeOut24HPm) * 100
        End If
        
        Label25.Caption = hours
        Label26.Caption = min
            
    Case 15
        
        If TimeOut24HPm = Timein24HPm Then hours = 4: min = 0
        If TimeOut24HPm < Timein24HPm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HPm - Timein24HPm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 2
            min = 0
        End If
        
        Label25.Caption = hours
        Label26.Caption = min
        
    Case 15.01 To 15.59
            
        If TimeOut24HPm = Timein24HPm Then hours = 4: min = 0
        If TimeOut24HPm < Timein24HPm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HPm - Timein24HPm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 1
            min = (15.6 - TimeOut24HPm) * 100
        End If
        
        Label25.Caption = hours
        Label26.Caption = min
    
    Case 16
        
        If TimeOut24HPm = Timein24HPm Then hours = 4: min = 0
        If TimeOut24HPm < Timein24HPm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HPm - Timein24HPm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 1
            min = 0
        End If
        
        Label25.Caption = hours
        Label26.Caption = min
        
    Case 16.01 To 16.59
            
        If TimeOut24HPm = Timein24HPm Then hours = 4: min = 0
        If TimeOut24HPm < Timein24HPm Then MsgBox "Invalid time input.", vbExclamation, "System Required": MaskEdBox2.SetFocus
        
        numHour = TimeOut24HPm - Timein24HPm
        
        If numHour < 1 Then
            hours = 4
            min = 0
        Else
            hours = 0
            min = (16.6 - TimeOut24HPm) * 100
        End If
        
        Label25.Caption = hours
        Label26.Caption = min
    
    Case Else
        
        Label25.Caption = 0
        Label26.Caption = 0

End Select
    
End Sub

Public Sub LoadEmpNos()
LoadEmpNo.Requery

Text5.Clear

While Not LoadEmpNo.EOF
    Text5.AddItem LoadEmpNo.Fields("Name")
    LoadEmpNo.MoveNext
Wend
End Sub

Private Sub Text5_Change()
On Error Resume Next

LoadEmpNo.Requery

LoadEmpNo.Find "Name='" & Trim(Text5.Text) & "'"

Label48.Caption = LoadEmpNo.Fields("EmployeeNo")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command5.SetFocus
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text7.SetFocus
End Sub

