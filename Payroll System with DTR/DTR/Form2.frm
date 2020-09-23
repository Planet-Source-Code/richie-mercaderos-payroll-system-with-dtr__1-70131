VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2550
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mnuSearch 
         Caption         =   "Search for..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print DTR Report"
         Begin VB.Menu mnuRegular 
            Caption         =   "Regular"
         End
         Begin VB.Menu mnuNRegular 
            Caption         =   "Non-Regular"
         End
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Time and Date Settings"
         Begin VB.Menu mnuSetTimeDate 
            Caption         =   "Set Time and Date"
         End
         Begin VB.Menu sep6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewDetails 
            Caption         =   "View List of Log Details"
         End
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddlog 
         Caption         =   "Modify Employee DTR"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu Calculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutDTR 
         Caption         =   "About"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnucancel 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuProperty 
      Caption         =   "Properties"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open DTR"
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About DTR"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rrs As New ADODB.Recordset

Private Sub Calculator_Click()
On Error GoTo calcpath

Shell "calc.exe", vbNormalFocus

calcpath:
    If Err.Number Then MsgBox "Your system does not have a calculator.", vbInformation, "Administrator"
End Sub

Private Sub Form_Load()
Call set_rec_getData(rrs, cn, "Select * From qryComputed")
End Sub

Private Sub mnuAbout_Click()
Form1.SSTab1.Tab = 1
Form1.Show
End Sub

Private Sub mnuAboutDTR_Click()
Form1.SSTab1.Tab = 1
End Sub

Private Sub mnuAddlog_Click()
Form5.useofthis = True
Form5.Show vbModal
End Sub

Private Sub mnucancel_Click()
MsgBox "DTR still running.", vbInformation
Form1.Timer3.Enabled = True
End Sub

Private Sub mnuLegend_Click()
Form7.Show vbModal
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuNRegular_Click()
Form3.Show vbModal
End Sub

Private Sub mnuOpen_Click()
Form1.SSTab1.Tab = 0
Form1.Show
End Sub

Private Sub mnuRegular_Click()
Set DataReport2.DataSource = rrs
DataReport2.Show vbModal
End Sub

Private Sub mnuSearch_Click()
Form11.Show vbModal
End Sub

Private Sub mnuSetTimeDate_Click()
Form5.useofthis = False
Form5.Show vbModal
End Sub

Private Sub mnuViewDetails_Click()
Form7.Show vbModal
End Sub
