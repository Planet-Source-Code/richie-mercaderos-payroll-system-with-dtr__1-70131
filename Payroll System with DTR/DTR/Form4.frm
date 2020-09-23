VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6270
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
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
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   ">"
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Enter Code:"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As DAO.Database

Dim ed_emp_timerec_out As DAO.Recordset
Dim emp_timerec As DAO.Recordset
Dim emp_no As DAO.Recordset

Dim emp_sal As DAO.Recordset
'-------------------------------------
Dim pangtrap As DAO.Recordset
'-------------------------------------
Dim checkit As DAO.Recordset
'-------------------------------------
Dim res As Integer
                                    
Dim tardy As Double
Dim minwork As Double

Dim hours As Integer
Dim min As Integer
Dim late As Double

Dim fulname As String
Dim desig As String

Dim pmin As Integer
Dim apmin As Integer

Dim log As Integer

Dim totalmin As Integer

Private Sub Command1_Click()
On Error GoTo checkmas

emp_no.Requery
emp_no.FindFirst "Code='" & Trim(Text1.Text) & "' and EmployeeNo='" & Form1.txtLog.Text & "'"

    If emp_no.NoMatch = False Then
    
        desig = emp_no.Fields("Designation")
        
        If desig = "Admin Aide I (Wages)" Or desig = "Admin Aide I (JO)" Then
            
            emp_sal.Requery
            emp_sal.FindFirst "Designation='" & desig & "'"
        
                If emp_sal.NoMatch = False Then
                    rateperh = Val(emp_sal.Fields("SalaryperDay")) / 8
                    rateperm = rateperh / 60
                End If
        End If
            
            Select Case Val(Form1.lblTime.Caption)
                
                Case 1 To 8: Call Category_8
                Case 8.01 To 8.59: Call Category_8K
                Case 9: Call Category_9
                Case 9.01 To 9.59: Call Category_9k
                Case 10: Call Category_10
                Case 10.01 To 10.59: Call Category_10k
                Case 11: Call Category_11
                Case 11.01 To 11.59: Call Category_11k
                Case 12: Call Category_12
                Case 12.01 To 12.59: Call Category_12k
                Case 13: Call Category_13
                Case 13.01 To 13.59: Call Category_13k
                Case 14: Call Category_14
                Case 14.01 To 14.59: Call Category_14k
                Case 15: Call Category_15
                Case 15.01 To 15.59: Call Category_15k
                Case 16: Call Category_16
                Case 16.01 To 16.59: Call Category_16k
                Case 17 To 21: Call Category_17
                
            End Select
            
        Unload Me
        Form1.pb.Visible = True
        Form1.pb.Value = 60
        Form1.sCount = 60
        Form1.Timer2.Enabled = True
    
    Else
        MsgBox "Incorrect code.", vbExclamation, "Administrator"
        Exit Sub
    End If

checkmas:
    If Err.Number = 380 Then MsgBox "Log out without log in is not valid.", vbExclamation, "Administrator"

End Sub

Private Sub Command2_Click()
Form1.BackNormalView
Unload Me
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\MasterFile.mdb", False, False, ";pwd=xxx")

fulname = Form1.Surname.Caption & ", " & Form1.Firstname.Caption & " " & Mid(Form1.MI.Caption, 1, 1)

Set emp_timerec = db.OpenRecordset("Select * From tblTimeRecord")

Set emp_sal = db.OpenRecordset("Select * From tblDesignation")

Set pangtrap = db.OpenRecordset("Select * From tblTimeRecord where EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and Name='" & fulname & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=2")

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Set ed_emp_timerec_out = db.OpenRecordset("Select * From tblTimeRecord where EmployeeNo='" & Trim(Form1.txtLog.Text) & "' And Name='" & fulname & "' And CDate='" & Format(Date, "mm/dd/yyyy") & "' And Month='" & Format(Date, "mmmm") & "' and Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & " and LastStatus='IN' Or LastStatus='OUT' And Count>=0 Or Count<=4 and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and TimeInPm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutPm<=" & Trim(Form1.lblTime.Caption))
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Set emp_no = db.OpenRecordset("Select * From tblEmployeeInfo")

End Sub

Private Sub Text1_GotFocus()
With Text1
    .SelStart = 0
    .SelLength = Len(Text1.Text)
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub


Public Sub Category_8()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "" ' and LastStatus='IN' and Count=1"
                                
        If emp_timerec.NoMatch = True Then
                                
            emp_timerec.AddNew
                emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                emp_timerec.Fields("Name") = fulname
                emp_timerec.Fields("Month") = Format(Date, "mmmm")
                emp_timerec.Fields("Day") = Format(Date, "dd")
                emp_timerec.Fields("Year") = Format(Date, "yyyy")
                emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                emp_timerec.Fields("AMIn") = Form1.Label7.Caption
                emp_timerec.Fields("TimeInAm") = Trim(Form1.lblTime.Caption)
                emp_timerec.Fields("LastStatus") = "IN"
                emp_timerec.Fields("NumDayWork") = 0
                emp_timerec.Fields("Count") = 1
            emp_timerec.Update
                            
        Else
            MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
        End If
                                               
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                          
End If
                    
End Sub

Public Sub Category_8K()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = FormatNumber((Val(Form1.lblTime.Caption) - 8) * 100)
                            
    hours = 0
    min = late

        emp_timerec.Requery
                            
        emp_timerec.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "" ' and LastStatus='IN' and Count=1 "
                                
            If emp_timerec.NoMatch = True Then
                                
                emp_timerec.AddNew
                    emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                    emp_timerec.Fields("Name") = fulname
                    emp_timerec.Fields("Month") = Format(Date, "mmmm")
                    emp_timerec.Fields("Day") = Format(Date, "dd")
                    emp_timerec.Fields("Year") = Format(Date, "yyyy")
                    emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                    emp_timerec.Fields("AMIn") = Form1.Label7.Caption
                    emp_timerec.Fields("TimeInAm") = Trim(Form1.lblTime.Caption)
                    emp_timerec.Fields("ATHour") = hours
                    emp_timerec.Fields("ATMin") = min
                                        
                    emp_timerec.Fields("TTinHour") = hours
                    emp_timerec.Fields("TTinMin") = min
                                        
                    emp_timerec.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                        
                    emp_timerec.Fields("LastStatus") = "IN"
                    emp_timerec.Fields("NumDayWork") = 0
                    emp_timerec.Fields("Count") = 1
                emp_timerec.Update
                            
            Else
                MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
            End If
                                               
        Form1.ListView1.ListItems.Clear
        Form1.LoadRecord
        Form1.onFocus
                            
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'"
                           
    If ed_emp_timerec_out.NoMatch = True Then Exit Sub
                           
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<TimeInAm"
                            
        If ed_emp_timerec_out.NoMatch = False Then
            MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
            Exit Sub
        Else
                                    
            pangtrap.Requery
                                    
            pangtrap.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=2"
                                    
                If pangtrap.NoMatch = True Then
                                     
                    hours = 4
                    min = 0
                                                
                    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<TimeInAM and LastStatus='IN' and Count=1"
                                                
                    ed_emp_timerec_out.Edit
                        ed_emp_timerec_out.Fields("AMOut") = Form1.Label7.Caption
                        ed_emp_timerec_out.Fields("TimeOutAm") = Trim(Form1.lblTime.Caption)
                        ed_emp_timerec_out.Fields("ATHour") = hours
                        ed_emp_timerec_out.Fields("ATMin") = min
                                        
                        ed_emp_timerec_out.Fields("TTinHour") = hours
                        ed_emp_timerec_out.Fields("TTinMin") = min
                                                    
                        ed_emp_timerec_out.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                                    
                        ed_emp_timerec_out.Fields("NumDayWork") = 0
                        ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                        ed_emp_timerec_out.Fields("Count") = 2
                    ed_emp_timerec_out.Update
                        
                Else
                    
                    MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time out.", vbExclamation, "Administrator"
                
                End If
        End If
                        
        Form1.ListView1.ListItems.Clear
        Form1.LoadRecord
        Form1.onFocus
End If
End Sub


Public Sub Category_9()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    hours = 1
    min = 0

    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "" ' and LastStatus='IN' and Count=1"
                                
    If emp_timerec.NoMatch = True Then
                                
        emp_timerec.AddNew
            emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
            emp_timerec.Fields("Name") = fulname
            emp_timerec.Fields("Month") = Format(Date, "mmmm")
            emp_timerec.Fields("Day") = Format(Date, "dd")
            emp_timerec.Fields("Year") = Format(Date, "yyyy")
            emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
            emp_timerec.Fields("AMIn") = Form1.Label7.Caption
            emp_timerec.Fields("TimeInAm") = Trim(Form1.lblTime.Caption)
            emp_timerec.Fields("ATHour") = hours
            emp_timerec.Fields("ATMin") = min
                                        
            emp_timerec.Fields("TTinHour") = hours
            emp_timerec.Fields("TTinMin") = min
                                        
            emp_timerec.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                        
            emp_timerec.Fields("LastStatus") = "IN"
            emp_timerec.Fields("NumDayWork") = 0
            emp_timerec.Fields("Count") = 1
        emp_timerec.Update
                            
    Else
        
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                
    End If
                                               
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                            
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    ed_emp_timerec_out.Requery
    
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'"
                           
    If ed_emp_timerec_out.NoMatch = True Then Exit Sub
    
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<TimeInAm"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                    
        pangtrap.Requery
                                    
        pangtrap.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=2"
                                    
        If pangtrap.NoMatch = True Then
                                                
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<TimeInAM and LastStatus='IN' and Count=1"
                                                
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("AMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutAm") = Form1.lblTime.Caption
                                                    
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutAm")) - Val(ed_emp_timerec_out.Fields("TimeInAm"))

                If minwork < 1 Then
                                                            
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("ATHour") = hours
                    ed_emp_timerec_out.Fields("ATMin") = min
                                        
                    ed_emp_timerec_out.Fields("TTinHour") = hours
                    ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                    ed_emp_timerec_out.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                                                                                                        
                Else
                                                        
                    hours = 3
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("ATHour") = hours
                    ed_emp_timerec_out.Fields("ATMin") = min
                                        
                    ed_emp_timerec_out.Fields("TTinHour") = hours
                    ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                    ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                    ed_emp_timerec_out.Fields("NumDayWork") = 0.5
                                                            
                End If
                                        
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 2
            ed_emp_timerec_out.Update

        Else
            MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time out.", vbExclamation, "Administrator"
        End If
                          
    End If
    
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
End If
End Sub


Public Sub Category_9k()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = FormatNumber((Val(Form1.lblTime.Caption) - 9) * 100)
                            
    hours = 1
    min = late

    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "" ' and LastStatus='IN' and Count=1"
                                
    If emp_timerec.NoMatch = True Then
                                
        emp_timerec.AddNew
            emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
            emp_timerec.Fields("Name") = fulname
            emp_timerec.Fields("Month") = Format(Date, "mmmm")
            emp_timerec.Fields("Day") = Format(Date, "dd")
            emp_timerec.Fields("Year") = Format(Date, "yyyy")
            emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
            emp_timerec.Fields("AMIn") = Form1.Label7.Caption
            emp_timerec.Fields("TimeInAm") = Trim(Form1.lblTime.Caption)
            emp_timerec.Fields("ATHour") = hours
            emp_timerec.Fields("ATMin") = min
                                        
            emp_timerec.Fields("TTinHour") = hours
            emp_timerec.Fields("TTinMin") = min
                                        
            emp_timerec.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                        
            emp_timerec.Fields("LastStatus") = "IN"
            emp_timerec.Fields("NumDayWork") = 0
            emp_timerec.Fields("Count") = 1
        emp_timerec.Update
                            
    Else
    
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                
    End If
                                               
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                            
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber(11.6 - (Val(Form1.lblTime.Caption)), 2)
                            
    hours = 2
    min = (late - Int(late)) * 100
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'"
                           
    If ed_emp_timerec_out.NoMatch = True Then Exit Sub
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<TimeInAm"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                    
        pangtrap.Requery
                                    
        pangtrap.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=2"
                                    
        If pangtrap.NoMatch = True Then
                                    
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<TimeInAM and LastStatus='IN' and Count=1"
            
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("AMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutAm") = Form1.lblTime.Caption
                                                    
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutAm")) - Val(ed_emp_timerec_out.Fields("TimeInAm"))

                If minwork < 1 Then
                                                            
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("ATHour") = hours
                    ed_emp_timerec_out.Fields("ATMin") = min
                                        
                    ed_emp_timerec_out.Fields("TTinHour") = hours
                    ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                    ed_emp_timerec_out.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                                        
                Else
                                                            
                    pmin = 0
                    pmin = Val(ed_emp_timerec_out.Fields("ATMin")) + min
                                                            
                    If pmin > 59 Then
                                                                
                        hours = hours + 1
                        min = pmin - 60
                                                            
                        ed_emp_timerec_out.Fields("ATHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                        ed_emp_timerec_out.Fields("ATMin") = min
                                                            
                        ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                        ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                                
                    Else
                                                            
                        ed_emp_timerec_out.Fields("ATHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                        ed_emp_timerec_out.Fields("ATMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + min
                                                            
                        ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                        ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("TTinMin")) + min
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                            
                    End If
                                                            
                        ed_emp_timerec_out.Fields("NumDayWork") = 0.5
                                                            
                End If
                                        
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 2
            ed_emp_timerec_out.Update

        Else
            MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time out.", vbExclamation, "Administrator"
        End If
    End If
                        
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
End If
End Sub

Public Sub Category_10()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    hours = 2
    min = 0
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "" ' and LastStatus='IN' and Count=1"
                                
    If emp_timerec.NoMatch = True Then
                                
        emp_timerec.AddNew
            emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
            emp_timerec.Fields("Name") = fulname
            emp_timerec.Fields("Month") = Format(Date, "mmmm")
            emp_timerec.Fields("Day") = Format(Date, "dd")
            emp_timerec.Fields("Year") = Format(Date, "yyyy")
            emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
            emp_timerec.Fields("AMIn") = Form1.Label7.Caption
            emp_timerec.Fields("TimeInAm") = Trim(Form1.lblTime.Caption)
            emp_timerec.Fields("ATHour") = hours
            emp_timerec.Fields("ATMin") = min
                                        
            emp_timerec.Fields("TTinHour") = Val(emp_timerec.Fields("TTinHour")) + hours
            emp_timerec.Fields("TTinMin") = Val(emp_timerec.Fields("TTinMin")) + min
                                        
            emp_timerec.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                        
            emp_timerec.Fields("LastStatus") = "IN"
            emp_timerec.Fields("NumDayWork") = 0
            emp_timerec.Fields("Count") = 1
        emp_timerec.Update
                            
    Else
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
    End If
                                               
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                            
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber(11.6 - (Val(Form1.lblTime.Caption)), 2)
                        
    min = (late - Int(late)) * 100
                            
    If min = 60 Then
        min = 0
        hours = Int(late) + 1
    Else
        hours = Int(late)
    End If
                            
                            'MsgBox
    ed_emp_timerec_out.Requery
    
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'"
                           
    If ed_emp_timerec_out.NoMatch = True Then Exit Sub
    
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<TimeInAm"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                    
        pangtrap.Requery
                                    
        pangtrap.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=2"
                                    
        If pangtrap.NoMatch = True Then
                                
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<TimeInAM and LastStatus='IN' and Count=1"
                                                
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("AMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutAm") = Form1.lblTime.Caption
                                                    
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutAm")) - Val(ed_emp_timerec_out.Fields("TimeInAm"))

                If minwork < 1 Then
                                                            
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("ATHour") = hours
                    ed_emp_timerec_out.Fields("ATMin") = min
                                        
                    ed_emp_timerec_out.Fields("TTinHour") = hours
                    ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                    ed_emp_timerec_out.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                                        
                Else
                                                        
                    ed_emp_timerec_out.Fields("ATHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                    ed_emp_timerec_out.Fields("ATMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + min
                                        
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("TTinMin")) + min
                                        
                    ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                    ed_emp_timerec_out.Fields("NumDayWork") = 0.5
                                                            
                End If
      
                                        
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 2
            ed_emp_timerec_out.Update

        Else
            
            MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time out.", vbExclamation, "Administrator"
        
        End If
                            
    End If
                        
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
End If
End Sub

Public Sub Category_10k()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber((Val(Form1.lblTime.Caption) - 10) * 100, 2)
                            
    hours = 2
    min = late
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "" ' and LastStatus='IN' and Count=1"
                                
    If emp_timerec.NoMatch = True Then
                                
        emp_timerec.AddNew
            emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
            emp_timerec.Fields("Name") = fulname
            emp_timerec.Fields("Month") = Format(Date, "mmmm")
            emp_timerec.Fields("Day") = Format(Date, "dd")
            emp_timerec.Fields("Year") = Format(Date, "yyyy")
            emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
            emp_timerec.Fields("AMIn") = Form1.Label7.Caption
            emp_timerec.Fields("TimeInAm") = Trim(Form1.lblTime.Caption)
            emp_timerec.Fields("ATHour") = hours
            emp_timerec.Fields("ATMin") = min
                                        
            emp_timerec.Fields("TTinHour") = hours
            emp_timerec.Fields("TTinMin") = min
                                        
            emp_timerec.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                
            emp_timerec.Fields("LastStatus") = "IN"
            emp_timerec.Fields("NumDayWork") = 0
            emp_timerec.Fields("Count") = 1
        emp_timerec.Update
                            
    Else
        
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                
    End If
                                               
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                            
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber(11.6 - (Val(Form1.lblTime.Caption)), 2)
                            
    hours = Int(late)
    min = (late - Int(late)) * 100
                            
    ed_emp_timerec_out.Requery
                     
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'"
                           
    If ed_emp_timerec_out.NoMatch = True Then Exit Sub
    
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<TimeInAm"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                    
        pangtrap.Requery
                                    
        pangtrap.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=2"
                                    
        If pangtrap.NoMatch = True Then
                                
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<TimeInAM and LastStatus='IN' and Count=1"
                                                                                                
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("AMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutAm") = Form1.lblTime.Caption
                                                    
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutAm")) - Val(ed_emp_timerec_out.Fields("TimeInAm"))

                If minwork < 1 Then
                                                            
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("ATHour") = hours
                    ed_emp_timerec_out.Fields("ATMin") = min
                                        
                    ed_emp_timerec_out.Fields("TTinHour") = hours
                    ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                    ed_emp_timerec_out.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                                        
                Else
                                                        
                    pmin = 0
                    pmin = Val(ed_emp_timerec_out.Fields("ATMin")) + min
                                                            
                    If pmin > 59 Then
                                                                
                        hours = hours + 1
                        min = pmin - 60
                                                            
                        ed_emp_timerec_out.Fields("ATHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                        ed_emp_timerec_out.Fields("ATMin") = min
                                                            
                        ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                        ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                                
                    Else
                                                            
                        ed_emp_timerec_out.Fields("ATHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                        ed_emp_timerec_out.Fields("ATMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + min
                                                            
                        ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                        ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("TTinMin")) + min
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                            
                    End If
                                                            
                    ed_emp_timerec_out.Fields("NumDayWork") = 0.5
                    
                End If

                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 2
            ed_emp_timerec_out.Update

        Else
            
            MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time out.", vbExclamation, "Administrator"
                                        
        End If
        
    End If
                        
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
End If
End Sub

Public Sub Category_11()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    hours = 3
    min = 0
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=1"
                                
    If emp_timerec.NoMatch = True Then
                                
        emp_timerec.AddNew
            emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
            emp_timerec.Fields("Name") = fulname
            emp_timerec.Fields("Month") = Format(Date, "mmmm")
            emp_timerec.Fields("Day") = Format(Date, "dd")
            emp_timerec.Fields("Year") = Format(Date, "yyyy")
            emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
            emp_timerec.Fields("AMIn") = Form1.Label7.Caption
            emp_timerec.Fields("TimeInAm") = Trim(Form1.lblTime.Caption)
            emp_timerec.Fields("ATHour") = hours
            emp_timerec.Fields("ATMin") = min
                                        
            emp_timerec.Fields("TTinHour") = Val(emp_timerec.Fields("TTinHour")) + hours
            emp_timerec.Fields("TTinMin") = Val(emp_timerec.Fields("TTinMin")) + min
                                        
            emp_timerec.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                        
            emp_timerec.Fields("LastStatus") = "IN"
            emp_timerec.Fields("NumDayWork") = 0
            emp_timerec.Fields("Count") = 1
        emp_timerec.Update
                            
    Else
        
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                
    End If
                                               
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                            
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber(11.6 - (Val(Form1.lblTime.Caption)), 2)
                            
    min = (late - Int(late)) * 100
                            
    If min = 60 Then
        min = 0
        hours = Int(late) + 1
    Else
        hours = Int(late)
    End If
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'"
                           
    If ed_emp_timerec_out.NoMatch = True Then Exit Sub
                          
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<TimeInAm"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                    
        pangtrap.Requery
                                    
        pangtrap.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Form1.Frame3.Caption, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=2"
                                    
        If pangtrap.NoMatch = True Then
                                
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<TimeInAM and LastStatus='IN' and Count=1"
                                                
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("AMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutAm") = Form1.lblTime.Caption
                                                    
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutAm")) - Val(ed_emp_timerec_out.Fields("TimeInAm"))

                If minwork < 1 Then
                                                            
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("ATHour") = hours
                    ed_emp_timerec_out.Fields("ATMin") = min
                                        
                    ed_emp_timerec_out.Fields("TTinHour") = hours
                    ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                    ed_emp_timerec_out.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                                        
                Else
                                                        
                    ed_emp_timerec_out.Fields("ATHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                    ed_emp_timerec_out.Fields("ATMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + min
                                        
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("TTinMin")) + min
                                        
                    ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                    ed_emp_timerec_out.Fields("NumDayWork") = 0.5
                                                            
                End If

                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 2
            ed_emp_timerec_out.Update

        Else
            
            MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time out.", vbExclamation, "Administrator"
                                        
        End If
    
    End If
                        
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
End If
End Sub

Public Sub Category_11k()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    hours = 4
    min = 0
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "" ' and LastStatus='IN' and Count=1"
                                
    If emp_timerec.NoMatch = True Then
                                
        emp_timerec.AddNew
            emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
            emp_timerec.Fields("Name") = fulname
            emp_timerec.Fields("Month") = Format(Date, "mmmm")
            emp_timerec.Fields("Day") = Format(Date, "dd")
            emp_timerec.Fields("Year") = Format(Date, "yyyy")
            emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
            emp_timerec.Fields("AMIn") = Form1.Label7.Caption
            emp_timerec.Fields("TimeInAm") = Trim(Form1.lblTime.Caption)
            emp_timerec.Fields("AMOut") = Form1.Label7.Caption
            emp_timerec.Fields("TimeOutAm") = Trim(Form1.lblTime.Caption)
                                        
            emp_timerec.Fields("ATHour") = hours
            emp_timerec.Fields("ATMin") = min
                                        
            emp_timerec.Fields("TTinHour") = hours
            emp_timerec.Fields("TTinMin") = min
                                        
            emp_timerec.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                        
            emp_timerec.Fields("LastStatus") = "OUT"
            emp_timerec.Fields("NumDayWork") = 0
            emp_timerec.Fields("Count") = 2
        emp_timerec.Update
                            
    Else
        
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                
    End If
                                               
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                            
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber(11.6 - (Val(Form1.lblTime.Caption)), 2)
                            
    hours = 0
    min = (late - Int(late)) * 100
                                                      
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'"
                           
    If ed_emp_timerec_out.NoMatch = True Then Exit Sub
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<TimeInAm"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                    
        pangtrap.Requery
                                    
        pangtrap.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=2"
                                    
        If pangtrap.NoMatch = True Then
                                
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<TimeInAM and LastStatus='IN' and Count=1"
                                                
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("AMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutAm") = Form1.lblTime.Caption
                                                    
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutAm")) - Val(ed_emp_timerec_out.Fields("TimeInAm"))

                If minwork < 1 Then
                                                            
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("ATHour") = hours
                    ed_emp_timerec_out.Fields("ATMin") = min
                                        
                    ed_emp_timerec_out.Fields("TTinHour") = hours
                    ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                    ed_emp_timerec_out.Fields("Deduction") = (hours * rateperh) + (min * rateperm)
                                                        
                Else
                                                        
                    pmin = 0
                    pmin = Val(ed_emp_timerec_out.Fields("ATMin")) + min
                                                            
                    If pmin > 59 Then
                                                                
                        hours = hours + 1
                        min = pmin - 60
                                                            
                        ed_emp_timerec_out.Fields("ATHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                        ed_emp_timerec_out.Fields("ATMin") = min
                                                            
                        ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                        ed_emp_timerec_out.Fields("TTinMin") = min
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                                
                    Else
                                                            
                        ed_emp_timerec_out.Fields("ATHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                        ed_emp_timerec_out.Fields("ATMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + min
                                                            
                        ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                        ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("TTinMin")) + min
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                            
                    End If
                                                           
                       ed_emp_timerec_out.Fields("NumDayWork") = 0.5
                       
                End If
                                        
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 2
            ed_emp_timerec_out.Update
            
        Else
            
            MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time out.", vbExclamation, "Administrator"
                                        
        End If
        
    End If
                        
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
End If
End Sub

Public Sub Category_12()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                        
    MsgBox Form1.lblTime.Caption & " is not allow. Please wait for a few minutes for time in pm.", vbExclamation, "Administration"
                        
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                                                                                    
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'"
                           
    If ed_emp_timerec_out.NoMatch = True Then Exit Sub
                        
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<TimeInAm"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                    
        pangtrap.Requery
                                    
        pangtrap.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=2"
                                    
        If pangtrap.NoMatch = True Then
                                        
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<TimeInAM and LastStatus='IN' and Count=1"
                                        
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("AMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutAm") = Form1.lblTime.Caption
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("NumDayWork") = 0.5
                ed_emp_timerec_out.Fields("Count") = 2
            ed_emp_timerec_out.Update
        Else
                
            MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time out.", vbExclamation, "Administrator"
                                    
        End If
    
    End If
                        
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
End If
End Sub

Public Sub Category_12k()
 If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Form1.txtLog.Text & "'And CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "' And Month='" & Format(Date, "mmmm") & "'And Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & "And LastStatus='IN' And Count=3 and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " And TimeInPm<=" & Trim(Form1.lblTime.Caption)
                            
    If emp_timerec.NoMatch = True Then
                                
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='OUT' and Count=2"
                                
        If ed_emp_timerec_out.NoMatch = True Then
                                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<TimeInAm and LastStatus='IN' and Count=1"
                                            
            If ed_emp_timerec_out.NoMatch = True Then
                                            
                emp_timerec.AddNew
                    emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                    emp_timerec.Fields("Name") = fulname
                    emp_timerec.Fields("Month") = Format(Date, "mmmm")
                    emp_timerec.Fields("Day") = Format(Date, "dd")
                    emp_timerec.Fields("Year") = Format(Date, "yyyy")
                    emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                    emp_timerec.Fields("PMIn") = Form1.Label7.Caption
                    emp_timerec.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    emp_timerec.Fields("ATHour") = 4
                    emp_timerec.Fields("ATMin") = 0
                    emp_timerec.Fields("TTinHour") = 4
                    emp_timerec.Fields("TTinMin") = 0
                    emp_timerec.Fields("Deduction") = 93.17
                    emp_timerec.Fields("LastStatus") = "IN"
                    emp_timerec.Fields("NumDayWork") = 0
                    emp_timerec.Fields("Count") = 3
                emp_timerec.Update
                                            
            Else
                                            
                ed_emp_timerec_out.Edit
                    ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                    ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    ed_emp_timerec_out.Fields("ATHour") = 4
                    ed_emp_timerec_out.Fields("ATMin") = 0
                    ed_emp_timerec_out.Fields("TTinHour") = 4
                    ed_emp_timerec_out.Fields("TTinMin") = 0
                    ed_emp_timerec_out.Fields("Deduction") = 93.17
                    ed_emp_timerec_out.Fields("LastStatus") = "IN"
                    ed_emp_timerec_out.Fields("NumDayWork") = 0
                    ed_emp_timerec_out.Fields("Count") = 3
                ed_emp_timerec_out.Update
                                            
            End If
        
        Else
                                                                            
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                ed_emp_timerec_out.Fields("LastStatus") = "IN"
                ed_emp_timerec_out.Fields("Count") = 3
            ed_emp_timerec_out.Update
                                    
        End If
                                    
    Else
                
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                    
    End If
                                                            
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and Name='" & fulname & "' and CDate='" & Format(Date, "mm/dd/yyyy") & "'"
                           
    If ed_emp_timerec_out.NoMatch = True Then Exit Sub
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and TimeInPm=" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                                                  
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=1"
                                       
        If ed_emp_timerec_out.NoMatch = False Then
                                         
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<TimeInAM and LastStatus='IN' and Count=1"
                                         
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("AMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutAm") = Trim(Form1.lblTime.Caption)
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("NumDayWork") = Val(ed_emp_timerec_out.Fields("NumDayWork")) + 0.5
                ed_emp_timerec_out.Fields("Count") = 2
            ed_emp_timerec_out.Update
        
        End If
                                    
    End If
                                
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus

End If
End Sub

Public Sub Category_13()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Form1.txtLog.Text & "'And CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "' And Month='" & Format(Date, "mmmm") & "'And Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & "And LastStatus='IN' And Count=3 and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " And TimeInPm<=" & Trim(Form1.lblTime.Caption)
                            
    If emp_timerec.NoMatch = True Then
                                
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='OUT' and Count=2"
                                
        If ed_emp_timerec_out.NoMatch = True Then
                                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='IN' and Count=1"
                                            
            If ed_emp_timerec_out.NoMatch = True Then
                                            
                emp_timerec.AddNew
                    emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                    emp_timerec.Fields("Name") = fulname
                    emp_timerec.Fields("Month") = Format(Date, "mmmm")
                    emp_timerec.Fields("Day") = Format(Date, "dd")
                    emp_timerec.Fields("Year") = Format(Date, "yyyy")
                    emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                    emp_timerec.Fields("PMIn") = Form1.Label7.Caption
                    emp_timerec.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    emp_timerec.Fields("ATHour") = 4
                    emp_timerec.Fields("ATMin") = 0
                    emp_timerec.Fields("TTinHour") = 4
                    emp_timerec.Fields("TTinMin") = 0
                    emp_timerec.Fields("Deduction") = 93.17
                    emp_timerec.Fields("LastStatus") = "IN"
                    emp_timerec.Fields("Count") = 3
                emp_timerec.Update
                                            
            Else
                                            
                ed_emp_timerec_out.Edit
                    ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                    ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    ed_emp_timerec_out.Fields("ATHour") = 4
                    ed_emp_timerec_out.Fields("ATMin") = 0
                    ed_emp_timerec_out.Fields("TTinHour") = 4
                    ed_emp_timerec_out.Fields("TTinMin") = 0
                    ed_emp_timerec_out.Fields("Deduction") = 93.17
                    ed_emp_timerec_out.Fields("LastStatus") = "IN"
                    ed_emp_timerec_out.Fields("Count") = 3
                ed_emp_timerec_out.Update
                                            
                End If
                                            
        Else
                                                                            
                                        'ed_emp_timerec_out.Requery
                                                                                       
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                ed_emp_timerec_out.Fields("LastStatus") = "IN"
                ed_emp_timerec_out.Fields("Count") = 3
            ed_emp_timerec_out.Update
                                            
        End If
                                    
    Else
        
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                    
    End If
                                                        
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                
End If
End Sub

Public Sub Category_13k()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber((Val(Form1.lblTime.Caption) - 13) * 100)
                            
    hours = 0
    min = late
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Form1.txtLog.Text & "'And CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "' And Month='" & Format(Date, "mmmm") & "'And Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & "And LastStatus='IN' And Count=3 and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " And TimeInPm<=" & Trim(Form1.lblTime.Caption)
                            
    If emp_timerec.NoMatch = True Then
                                
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='OUT' and Count=2"
                                
        If ed_emp_timerec_out.NoMatch = True Then
                                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='IN' and Count=1"
                                            
            If ed_emp_timerec_out.NoMatch = True Then
                                            
                emp_timerec.AddNew
                    emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                    emp_timerec.Fields("Name") = fulname
                    emp_timerec.Fields("Month") = Format(Date, "mmmm")
                    emp_timerec.Fields("Day") = Format(Date, "dd")
                    emp_timerec.Fields("Year") = Format(Date, "yyyy")
                    emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                    emp_timerec.Fields("PMIn") = Form1.Label7.Caption
                    emp_timerec.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    emp_timerec.Fields("ATHour") = 4
                    emp_timerec.Fields("ATMin") = 0
                    emp_timerec.Fields("PTHour") = hours
                    emp_timerec.Fields("PTMin") = min
                    emp_timerec.Fields("TTinHour") = 4 + hours
                    emp_timerec.Fields("TTinMin") = 0 + min
                    emp_timerec.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    emp_timerec.Fields("LastStatus") = "IN"
                    emp_timerec.Fields("NumDayWork") = 0
                    emp_timerec.Fields("Count") = 3
                emp_timerec.Update
                                            
            Else
                                            
                ed_emp_timerec_out.Edit
                    ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                    ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    ed_emp_timerec_out.Fields("ATHour") = 4
                    ed_emp_timerec_out.Fields("ATMin") = 0
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = 4 + hours
                    ed_emp_timerec_out.Fields("TTinMin") = min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    ed_emp_timerec_out.Fields("LastStatus") = "IN"
                    ed_emp_timerec_out.Fields("NumDayWork") = 0
                    ed_emp_timerec_out.Fields("Count") = 3
                ed_emp_timerec_out.Update
                                            
            End If
                                            
        Else
            
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                ed_emp_timerec_out.Fields("PTHour") = hours
                ed_emp_timerec_out.Fields("PTMin") = min
            
                apmin = 0
                apmin = Val(ed_emp_timerec_out.Fields("ATMin")) + min
            
                If apmin > 59 Then
                    hours = hours + 1
                    apmin = apmin - 60
                End If

                ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                ed_emp_timerec_out.Fields("TTinMin") = apmin
                ed_emp_timerec_out.Fields("Deduction") = ((Val(ed_emp_timerec_out.Fields("ATHour")) + hours) * rateperh) + (apmin * rateperm)
                ed_emp_timerec_out.Fields("LastStatus") = "IN"
                ed_emp_timerec_out.Fields("Count") = 3
            ed_emp_timerec_out.Update
                                            
        End If
                                    
    Else
                    
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                    
    End If
                                                            
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                
                    
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    hours = 4
    min = 0
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and TimeInPm=" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                                                  
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                                       
        If ed_emp_timerec_out.NoMatch = False Then
                                         
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInPm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutPm<TimeInPM and LastStatus='IN' and Count=3"
                                         
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutPm") = Trim(Form1.lblTime.Caption)
                ed_emp_timerec_out.Fields("PTHour") = hours
                ed_emp_timerec_out.Fields("PTMin") = min
                ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("ATMin"))
                ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 4
            ed_emp_timerec_out.Update
                                    
        End If
                                    
    End If
                                
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                       
End If
End Sub

Public Sub Category_14()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    hours = 1
    min = 0
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Form1.txtLog.Text & "'And CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "' And Month='" & Format(Date, "mmmm") & "'And Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & "And LastStatus='IN' And Count=3 and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " And TimeInPm<=" & Trim(Form1.lblTime.Caption)
                            
    If emp_timerec.NoMatch = True Then
                                
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='OUT' and Count=2"
                                
            If ed_emp_timerec_out.NoMatch = True Then
                                            
                ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='IN' and Count=1"
                                            
                    If ed_emp_timerec_out.NoMatch = True Then
                                            
                        emp_timerec.AddNew
                            emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                            emp_timerec.Fields("Name") = fulname
                            emp_timerec.Fields("Month") = Format(Date, "mmmm")
                            emp_timerec.Fields("Day") = Format(Date, "dd")
                            emp_timerec.Fields("Year") = Format(Date, "yyyy")
                            emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                            emp_timerec.Fields("PMIn") = Form1.Label7.Caption
                            emp_timerec.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                            emp_timerec.Fields("ATHour") = 4
                            emp_timerec.Fields("ATMin") = 0
                            emp_timerec.Fields("PTHour") = hours
                            emp_timerec.Fields("PTMin") = min
                            emp_timerec.Fields("TTinHour") = 4 + hours
                            emp_timerec.Fields("TTinMin") = 0 + min
                            emp_timerec.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                            emp_timerec.Fields("LastStatus") = "IN"
                            emp_timerec.Fields("NumDayWork") = 0
                            emp_timerec.Fields("Count") = 3
                        emp_timerec.Update
                                            
                    Else
                                            
                        ed_emp_timerec_out.Edit
                            ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                            ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                            ed_emp_timerec_out.Fields("ATHour") = 4
                            ed_emp_timerec_out.Fields("ATMin") = 0
                            ed_emp_timerec_out.Fields("PTHour") = hours
                            ed_emp_timerec_out.Fields("PTMin") = min
                            ed_emp_timerec_out.Fields("TTinHour") = 4 + hours
                            ed_emp_timerec_out.Fields("TTinMin") = 0 + min
                            ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                            ed_emp_timerec_out.Fields("LastStatus") = "IN"
                            ed_emp_timerec_out.Fields("NumDayWork") = 0
                            ed_emp_timerec_out.Fields("Count") = 3
                        ed_emp_timerec_out.Update
                                            
                    End If
                                            
                Else
                                                                            
                                        'ed_emp_timerec_out.Requery
                                                                                        
                    ed_emp_timerec_out.Edit
                        ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                        ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                        ed_emp_timerec_out.Fields("PTHour") = hours
                        ed_emp_timerec_out.Fields("PTMin") = min
                        ed_emp_timerec_out.Fields("TTinHour") = 4 + hours
                        ed_emp_timerec_out.Fields("TTinMin") = 0 + min
                        ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                        ed_emp_timerec_out.Fields("LastStatus") = "IN"
                        ed_emp_timerec_out.Fields("Count") = 3
                    ed_emp_timerec_out.Update
                                            
                End If
                                    
            Else
                
                MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
            
            End If
                                         
        Form1.ListView1.ListItems.Clear
        Form1.LoadRecord
        Form1.onFocus
                                
                    
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    hours = 4
    min = 0
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and TimeInPm=" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                                                  
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                                       
        If ed_emp_timerec_out.NoMatch = False Then
                                         
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInPm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutPm<TimeInPM and LastStatus='IN' and Count=3"
                                         
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutPm") = Trim(Form1.lblTime.Caption)
                                                        
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutPm")) - Val(ed_emp_timerec_out.Fields("TimeInPm"))
                                                    
                If minwork < 1 Then
                                                        
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                    ed_emp_timerec_out.Fields("TTinMin") = min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                                                        
                Else
                    
                    hours = 3
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("TTinMin")) + min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                                                        
                End If
                                                        
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("NumDayWork") = Val(ed_emp_timerec_out.Fields("NumDayWork")) + 0.5
                ed_emp_timerec_out.Fields("Count") = 4
            ed_emp_timerec_out.Update
        End If
                                    
    End If
                                
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                       
End If
End Sub

Public Sub Category_14k()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber((Val(Form1.lblTime.Caption) - 14) * 100)
                            
    hours = 1
    min = late
                            
    emp_timerec.Requery
                
    emp_timerec.FindFirst "EmployeeNo='" & Form1.txtLog.Text & "'And CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "' And Month='" & Format(Date, "mmmm") & "'And Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & "And LastStatus='IN' And Count=3 and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " And TimeInPm<=" & Trim(Form1.lblTime.Caption)
                            
    If emp_timerec.NoMatch = True Then
                                
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='OUT' and Count=2"
                                
        If ed_emp_timerec_out.NoMatch = True Then
                                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='IN' and Count=1"
                                            
            If ed_emp_timerec_out.NoMatch = True Then
                                            
                emp_timerec.AddNew
                    emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                    emp_timerec.Fields("Name") = fulname
                    emp_timerec.Fields("Month") = Format(Date, "mmmm")
                    emp_timerec.Fields("Day") = Format(Date, "dd")
                    emp_timerec.Fields("Year") = Format(Date, "yyyy")
                    emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                    emp_timerec.Fields("PMIn") = Form1.Label7.Caption
                    emp_timerec.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    emp_timerec.Fields("ATHour") = 4
                    emp_timerec.Fields("ATMin") = 0
                    emp_timerec.Fields("PTHour") = hours
                    emp_timerec.Fields("PTMin") = min
                    emp_timerec.Fields("TTinHour") = 4 + hours
                    emp_timerec.Fields("TTinMin") = 0 + min
                    emp_timerec.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    emp_timerec.Fields("LastStatus") = "IN"
                    emp_timerec.Fields("NumDayWork") = 0
                    emp_timerec.Fields("Count") = 3
                emp_timerec.Update
                                            
            Else
                                            
                ed_emp_timerec_out.Edit
                    ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                    ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    ed_emp_timerec_out.Fields("ATHour") = 4
                    ed_emp_timerec_out.Fields("ATMin") = 0
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = 4 + hours
                    ed_emp_timerec_out.Fields("TTinMin") = 0 + min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    ed_emp_timerec_out.Fields("LastStatus") = "IN"
                    ed_emp_timerec_out.Fields("NumDayWork") = 0
                    ed_emp_timerec_out.Fields("Count") = 3
                ed_emp_timerec_out.Update
                                            
            End If
                                            
        Else
                                                                            
                                        'ed_emp_timerec_out.Requery
                                                                                        
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                ed_emp_timerec_out.Fields("PTHour") = hours
                ed_emp_timerec_out.Fields("PTMin") = min
                
                apmin = 0
                apmin = Val(ed_emp_timerec_out.Fields("ATMin")) + min
            
                If apmin > 59 Then
                    hours = hours + 1
                    apmin = apmin - 60
                End If

                ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                ed_emp_timerec_out.Fields("TTinMin") = apmin
                ed_emp_timerec_out.Fields("Deduction") = ((Val(ed_emp_timerec_out.Fields("ATHour")) + hours) * rateperh) + (apmin * rateperm)
                ed_emp_timerec_out.Fields("LastStatus") = "IN"
                ed_emp_timerec_out.Fields("Count") = 3
            ed_emp_timerec_out.Update
                                            
        End If
                                    
    Else
        
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                    
    End If
                                              
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                
                    
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber(16.6 - (Val(Form1.lblTime.Caption)), 2)
                            
    hours = Int(late)
    min = (late - Int(late)) * 100
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and TimeInPm=" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                                                  
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                                       
        If ed_emp_timerec_out.NoMatch = False Then
                                         
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInPm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutPm<TimeInPM and LastStatus='IN' and Count=3"
                                         
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutPm") = Trim(Form1.lblTime.Caption)
                                                        
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutPm")) - Val(ed_emp_timerec_out.Fields("TimeInPm"))

                If minwork < 1 Then
                                                        
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + Val(ed_emp_timerec_out.Fields("PTHour"))
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                                                        
                Else
                
                    pmin = 0
                    pmin = Val(ed_emp_timerec_out.Fields("PTMin")) + min
                                                            
                    If pmin > 59 Then
                                                                
                        hours = hours + 1
                        min = pmin - 60
                                                            
                        ed_emp_timerec_out.Fields("PTHour") = Val(ed_emp_timerec_out.Fields("PTHour")) + hours
                        ed_emp_timerec_out.Fields("PTMin") = min
                                                            
                        totalmin = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        
                        If totalmin > 59 Then
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                            ed_emp_timerec_out.Fields("TTinMin") = totalmin - 60
                        Else
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                            ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        End If
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                                
                    Else
                                                            
                        ed_emp_timerec_out.Fields("PTHour") = Val(ed_emp_timerec_out.Fields("PTHour")) + hours
                        ed_emp_timerec_out.Fields("PTMin") = Val(ed_emp_timerec_out.Fields("PTMin")) + min
                                                            
                        totalmin = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        
                        If totalmin > 59 Then
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + 3
                            ed_emp_timerec_out.Fields("TTinMin") = totalmin - 60
                        Else
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                            ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        End If
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                            
                    End If
                          
                    ed_emp_timerec_out.Fields("NumDayWork") = Val(ed_emp_timerec_out.Fields("NumDayWork")) + 0.5
                                                        
                End If

                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 4
            ed_emp_timerec_out.Update
                                            
        End If
                                    
    End If
                                
        Form1.ListView1.ListItems.Clear
        Form1.LoadRecord
        Form1.onFocus
                                       
End If
End Sub

Public Sub Category_15()

If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    hours = 2
    min = 0
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Form1.txtLog.Text & "'And CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "' And Month='" & Format(Date, "mmmm") & "'And Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & "And LastStatus='IN' And Count=3 and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " And TimeInPm<=" & Trim(Form1.lblTime.Caption)
                            
    If emp_timerec.NoMatch = True Then
                                
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='OUT' and Count=2"
                                
        If ed_emp_timerec_out.NoMatch = True Then
                                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='IN' and Count=1"
                                            
            If ed_emp_timerec_out.NoMatch = True Then
                                            
                emp_timerec.AddNew
                    emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                    emp_timerec.Fields("Name") = fulname
                    emp_timerec.Fields("Month") = Format(Date, "mmmm")
                    emp_timerec.Fields("Day") = Format(Date, "dd")
                    emp_timerec.Fields("Year") = Format(Date, "yyyy")
                    emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                    emp_timerec.Fields("PMIn") = Form1.Label7.Caption
                    emp_timerec.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    emp_timerec.Fields("ATHour") = 4
                    emp_timerec.Fields("ATMin") = 0
                    emp_timerec.Fields("PTHour") = hours
                    emp_timerec.Fields("PTMin") = min
                    emp_timerec.Fields("TTinHour") = 4 + hours
                    emp_timerec.Fields("TTinMin") = 0 + min
                    emp_timerec.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    emp_timerec.Fields("LastStatus") = "IN"
                    emp_timerec.Fields("NumDayWork") = 0
                    emp_timerec.Fields("Count") = 3
                emp_timerec.Update
                                            
            Else
                                            
                ed_emp_timerec_out.Edit
                    ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                    ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    ed_emp_timerec_out.Fields("ATHour") = 4
                    ed_emp_timerec_out.Fields("ATMin") = 0
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = 4 + hours
                    ed_emp_timerec_out.Fields("TTinMin") = 0 + min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    ed_emp_timerec_out.Fields("LastStatus") = "IN"
                    ed_emp_timerec_out.Fields("NumDayWork") = 0
                    ed_emp_timerec_out.Fields("Count") = 3
                ed_emp_timerec_out.Update
                                            
            End If
                                            
        Else
                                                                            
                                        'ed_emp_timerec_out.Requery
                                                                                        
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                ed_emp_timerec_out.Fields("PTHour") = hours
                ed_emp_timerec_out.Fields("PTMin") = min
                
                apmin = 0
                apmin = Val(ed_emp_timerec_out.Fields("ATMin")) + min
            
                If apmin > 59 Then
                    hours = hours + 1
                    apmin = apmin - 60
                End If

                ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                ed_emp_timerec_out.Fields("TTinMin") = apmin
                ed_emp_timerec_out.Fields("Deduction") = ((Val(ed_emp_timerec_out.Fields("ATHour")) + hours) * rateperh) + (apmin * rateperm)
                ed_emp_timerec_out.Fields("LastStatus") = "IN"
                ed_emp_timerec_out.Fields("Count") = 3
            ed_emp_timerec_out.Update
                                            
            End If
                                    
        Else
            
            MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                    
        End If
                                                          
        Form1.ListView1.ListItems.Clear
        Form1.LoadRecord
        Form1.onFocus
                                
                    
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and TimeInPm=" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                                                  
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                                       
        If ed_emp_timerec_out.NoMatch = False Then
                                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInPm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutPm<TimeInPM and LastStatus='IN' and Count=3"
                                         
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutPm") = Trim(Form1.lblTime.Caption)
                                                        
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutPm")) - Val(ed_emp_timerec_out.Fields("TimeInPm"))

                If minwork < 1 Then
                                                        
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + Val(ed_emp_timerec_out.Fields("PTHour"))
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                                                        
                Else
                                                        
                    hours = 2
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("PTHour") = Val(ed_emp_timerec_out.Fields("PTHour")) + hours
                    ed_emp_timerec_out.Fields("PTMin") = Val(ed_emp_timerec_out.Fields("PTMin")) + min
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("TTinMin")) + min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    ed_emp_timerec_out.Fields("NumDayWork") = Val(ed_emp_timerec_out.Fields("NumDayWork")) + 0.5
                
                End If
                                                        
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 4
            ed_emp_timerec_out.Update
                                          
        End If
                                    
    End If
                                
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                       
End If
End Sub

Public Sub Category_15k()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber((Val(Form1.lblTime.Caption) - 15) * 100)
                            
    hours = 2
    min = late
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Form1.txtLog.Text & "'And CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "' And Month='" & Format(Date, "mmmm") & "'And Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & "And LastStatus='IN' And Count=3 and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " And TimeInPm<=" & Trim(Form1.lblTime.Caption)
                            
    If emp_timerec.NoMatch = True Then
                                
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='OUT' and Count=2"
                                
        If ed_emp_timerec_out.NoMatch = True Then
                                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='IN' and Count=1"
                                            
            If ed_emp_timerec_out.NoMatch = True Then
                                            
                emp_timerec.AddNew
                    emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                    emp_timerec.Fields("Name") = fulname
                    emp_timerec.Fields("Month") = Format(Date, "mmmm")
                    emp_timerec.Fields("Day") = Format(Date, "dd")
                    emp_timerec.Fields("Year") = Format(Date, "yyyy")
                    emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                    emp_timerec.Fields("PMIn") = Form1.Label7.Caption
                    emp_timerec.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    emp_timerec.Fields("ATHour") = 4
                    emp_timerec.Fields("ATMin") = 0
                    emp_timerec.Fields("PTHour") = hours
                    emp_timerec.Fields("PTMin") = min
                    emp_timerec.Fields("TTinHour") = 4 + hours
                    emp_timerec.Fields("TTinMin") = 0 + min
                    emp_timerec.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    emp_timerec.Fields("LastStatus") = "IN"
                    emp_timerec.Fields("NumDayWork") = 0
                    emp_timerec.Fields("Count") = 3
                emp_timerec.Update
                                            
            Else
                                            
                ed_emp_timerec_out.Edit
                    ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                    ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    ed_emp_timerec_out.Fields("ATHour") = 4
                    ed_emp_timerec_out.Fields("ATMin") = 0
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = 4 + hours
                    ed_emp_timerec_out.Fields("TTinMin") = 0 + min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    ed_emp_timerec_out.Fields("LastStatus") = "IN"
                    ed_emp_timerec_out.Fields("NumDayWork") = 0
                    ed_emp_timerec_out.Fields("Count") = 3
                ed_emp_timerec_out.Update
                                            
            End If
                                            
        Else
                                                                            
                                        'ed_emp_timerec_out.Requery
                                                                                        
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                emp_timerec.Fields("PTHour") = hours
                emp_timerec.Fields("PTMin") = min
                
                apmin = 0
                apmin = Val(ed_emp_timerec_out.Fields("ATMin")) + min
            
                If apmin > 59 Then
                    hours = hours + 1
                    apmin = apmin - 60
                End If

                ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                ed_emp_timerec_out.Fields("TTinMin") = apmin
                ed_emp_timerec_out.Fields("Deduction") = ((Val(ed_emp_timerec_out.Fields("ATHour")) + hours) * rateperh) + (apmin * rateperm)
                ed_emp_timerec_out.Fields("LastStatus") = "IN"
                ed_emp_timerec_out.Fields("Count") = 3
            ed_emp_timerec_out.Update
                                            
        End If
                                    
    Else
        
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                    
    End If
                                                                  
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                
                    
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber(16.6 - (Val(Form1.lblTime.Caption)), 2)
                            
    hours = Int(late)
    min = (late - Int(late)) * 100
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and TimeInPm=" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                                                  
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                                       
        If ed_emp_timerec_out.NoMatch = False Then
                                         
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInPm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutPm<TimeInPM and LastStatus='IN' and Count=3"
                                         
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutPm") = Trim(Form1.lblTime.Caption)
                                                        
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutPm")) - Val(ed_emp_timerec_out.Fields("TimeInPm"))

                If minwork < 1 Then
                                                        
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + Val(ed_emp_timerec_out.Fields("PTHour"))
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                                                        
                Else
                    
                    pmin = 0
                    pmin = Val(ed_emp_timerec_out.Fields("PTMin")) + min
                                                            
                    If pmin > 59 Then
                                                                
                        hours = hours + 1
                        min = pmin - 60
                                                            
                        ed_emp_timerec_out.Fields("PTHour") = Val(ed_emp_timerec_out.Fields("PTHour")) + hours
                        ed_emp_timerec_out.Fields("PTMin") = min
                                                           
                        totalmin = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        
                        If totalmin > 59 Then
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                            ed_emp_timerec_out.Fields("TTinMin") = totalmin - 60
                        Else
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                            ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        End If
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                                
                    Else
                                                            
                        ed_emp_timerec_out.Fields("PTHour") = Val(ed_emp_timerec_out.Fields("PTHour")) + hours
                        ed_emp_timerec_out.Fields("PTMin") = Val(ed_emp_timerec_out.Fields("PTMin")) + min
                                                         
                        totalmin = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                                                           
                        If totalmin > 59 Then
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + 2
                            ed_emp_timerec_out.Fields("TTinMin") = totalmin - 60
                        Else
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                            ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        End If
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                            
                    End If
                          
                    ed_emp_timerec_out.Fields("NumDayWork") = Val(ed_emp_timerec_out.Fields("NumDayWork")) + 0.5
                                                        
                End If
                                                        
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 4
            ed_emp_timerec_out.Update
                                            
        End If
                                    
    End If
                                
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                       
End If
End Sub

Public Sub Category_16()
If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
    hours = 3
    min = 0
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Form1.txtLog.Text & "'And CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "' And Month='" & Format(Date, "mmmm") & "'And Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & "And LastStatus='IN' And Count=3 and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " And TimeInPm<=" & Trim(Form1.lblTime.Caption)
                            
    If emp_timerec.NoMatch = True Then
                                
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='OUT' and Count=2"
                                
        If ed_emp_timerec_out.NoMatch = True Then
                                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='IN' and Count=1"
                                            
            If ed_emp_timerec_out.NoMatch = True Then
                                            
                emp_timerec.AddNew
                    emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                    emp_timerec.Fields("Name") = fulname
                    emp_timerec.Fields("Month") = Format(Date, "mmmm")
                    emp_timerec.Fields("Day") = Format(Date, "dd")
                    emp_timerec.Fields("Year") = Format(Date, "yyyy")
                    emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                    emp_timerec.Fields("PMIn") = Form1.Label7.Caption
                    emp_timerec.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    emp_timerec.Fields("ATHour") = 4
                    emp_timerec.Fields("ATMin") = 0
                    emp_timerec.Fields("PTHour") = hours
                    emp_timerec.Fields("PTMin") = min
                    emp_timerec.Fields("TTinHour") = 4 + hours
                    emp_timerec.Fields("TTinMin") = 0 + min
                    emp_timerec.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    emp_timerec.Fields("LastStatus") = "IN"
                    emp_timerec.Fields("NumDayWork") = 0
                    emp_timerec.Fields("Count") = 3
                emp_timerec.Update
                                            
            Else
                                            
                ed_emp_timerec_out.Edit
                    ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                    ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    ed_emp_timerec_out.Fields("ATHour") = 4
                    ed_emp_timerec_out.Fields("ATMin") = 0
                    emp_timerec.Fields("PTHour") = hours
                    emp_timerec.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = 4 + hours
                    ed_emp_timerec_out.Fields("TTinMin") = 0 + min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    ed_emp_timerec_out.Fields("LastStatus") = "IN"
                    ed_emp_timerec_out.Fields("NumDayWork") = 0
                    ed_emp_timerec_out.Fields("Count") = 3
                ed_emp_timerec_out.Update
                                            
            End If
                                            
        Else
                                                                            
                                        'ed_emp_timerec_out.Requery
                                                                                        
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                ed_emp_timerec_out.Fields("PTHour") = hours
                ed_emp_timerec_out.Fields("PTMin") = min
                ed_emp_timerec_out.Fields("TTinHour") = 4 + hours
                ed_emp_timerec_out.Fields("TTinMin") = 0 + min
                ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                ed_emp_timerec_out.Fields("LastStatus") = "IN"
                ed_emp_timerec_out.Fields("Count") = 3
            ed_emp_timerec_out.Update
                                            
        End If
                                    
    Else
        
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                    
    End If
                      
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                
                    
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and TimeInPm=" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                                                  
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                                       
        If ed_emp_timerec_out.NoMatch = False Then
                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInPm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutPm<TimeInPM and LastStatus='IN' and Count=3"
                                         
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutPm") = Trim(Form1.lblTime.Caption)
                                                        
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutPm")) - Val(ed_emp_timerec_out.Fields("TimeInPm"))

                If minwork < 1 Then
                                                        
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("PTHour")) + Val(ed_emp_timerec_out.Fields("ATHour"))
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("PTMin")) + Val(ed_emp_timerec_out.Fields("ATMin"))
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                                                        
                Else
                    
                    hours = 1
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("PTHour") = Val(ed_emp_timerec_out.Fields("PTHour")) + hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    ed_emp_timerec_out.Fields("NumDayWork") = Val(ed_emp_timerec_out.Fields("NumDayWork")) + 0.5
                
                End If
                                                        
                    ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                    ed_emp_timerec_out.Fields("Count") = 4
                ed_emp_timerec_out.Update
                                    
            End If
                                    
        End If
                                
        Form1.ListView1.ListItems.Clear
        Form1.LoadRecord
        Form1.onFocus
                                       
End If
End Sub

Public Sub Category_16k()

If Form1.optIn.Value = True And Val(Form1.lblTime.Caption) Then
                            
                            
    hours = 4
    min = 0
                            
    emp_timerec.Requery
                            
    emp_timerec.FindFirst "EmployeeNo='" & Form1.txtLog.Text & "'And CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "' And Month='" & Format(Date, "mmmm") & "'And Year='" & Format(Date, "yyyy") & "'And Day=" & Format(Date, "dd") & "and TimeInAm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & " And TimeInPm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutPm<=" & Trim(Form1.lblTime.Caption)
                            
    If emp_timerec.NoMatch = True Then
                                
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='OUT' and Count=2"
                                
        If ed_emp_timerec_out.NoMatch = True Then
                                            
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInAm<=" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<=" & Trim(Form1.lblTime.Caption) & "and LastStatus='IN' and Count=1"
                                            
            If ed_emp_timerec_out.NoMatch = True Then
                                            
                emp_timerec.AddNew
                    emp_timerec.Fields("EmployeeNo") = Form1.txtLog.Text
                    emp_timerec.Fields("Name") = fulname
                    emp_timerec.Fields("Month") = Format(Date, "mmmm")
                    emp_timerec.Fields("Day") = Format(Date, "dd")
                    emp_timerec.Fields("Year") = Format(Date, "yyyy")
                    emp_timerec.Fields("CDate") = Trim(Format(Date, "mm/dd/yyyy"))
                    emp_timerec.Fields("PMIn") = Form1.Label7.Caption
                    emp_timerec.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    emp_timerec.Fields("ATHour") = 4
                    emp_timerec.Fields("ATMin") = 0
                    emp_timerec.Fields("PTHour") = hours
                    emp_timerec.Fields("PTMin") = min
                    emp_timerec.Fields("TTinHour") = 4 + hours
                    emp_timerec.Fields("TTinMin") = 0 + min
                    emp_timerec.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    emp_timerec.Fields("LastStatus") = "OUT"
                    emp_timerec.Fields("NumDayWork") = 0
                    emp_timerec.Fields("Count") = 4
                emp_timerec.Update
                                            
            Else
                                            
                ed_emp_timerec_out.Edit
                    ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                    ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                    ed_emp_timerec_out.Fields("ATHour") = 4
                    ed_emp_timerec_out.Fields("ATMin") = 0
                    emp_timerec.Fields("PTHour") = hours
                    emp_timerec.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = 4 + hours
                    ed_emp_timerec_out.Fields("TTinMin") = 0 + min
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                    ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                    ed_emp_timerec_out.Fields("NumDayWork") = 0
                    ed_emp_timerec_out.Fields("Count") = 4
                ed_emp_timerec_out.Update
                                            
            End If
                                            
        Else
                                                                            
                                        'ed_emp_timerec_out.Requery
                                                                                        
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMIn") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeInPm") = Trim(Form1.lblTime.Caption)
                emp_timerec.Fields("PTHour") = hours
                emp_timerec.Fields("PTMin") = min
                
                apmin = 0
                apmin = Val(ed_emp_timerec_out.Fields("ATMin")) + min
            
                If apmin > 59 Then
                    hours = hours + 1
                    apmin = apmin - 60
                End If

                ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("ATHour")) + hours
                ed_emp_timerec_out.Fields("TTinMin") = apmin
                ed_emp_timerec_out.Fields("Deduction") = ((Val(ed_emp_timerec_out.Fields("ATHour")) + hours) * rateperh) + (apmin * rateperm)
                ed_emp_timerec_out.Fields("LastStatus") = "IN"
                ed_emp_timerec_out.Fields("Count") = 3
            ed_emp_timerec_out.Update
                                            
        End If
                                    
    Else
                                        
        MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time in.", vbExclamation, "Administrator"
                                    
    End If
                                               
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                
                    
ElseIf Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                            
    late = 0
    late = FormatNumber(16.6 - (Val(Form1.lblTime.Caption)), 2)
                            
    hours = Int(late)
    min = (late - Int(late)) * 100
                            
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and TimeInPm=" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                            
    If ed_emp_timerec_out.NoMatch = False Then
        MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
        Exit Sub
    Else
                                                                  
        ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "' and CDate='" & Trim(Format(Now, "mm/dd/yyyy")) & "' and TimeInAm<" & Trim(Form1.lblTime.Caption) & " and TimeOutAm<" & Trim(Form1.lblTime.Caption) & " and LastStatus='IN' and Count=3"
                                       
        If ed_emp_timerec_out.NoMatch = False Then
                                         
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInPm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutPm<TimeInPM and LastStatus='IN' and Count=3"
                                         
            ed_emp_timerec_out.Edit
                ed_emp_timerec_out.Fields("PMOut") = Form1.Label7.Caption
                ed_emp_timerec_out.Fields("TimeOutPm") = Trim(Form1.lblTime.Caption)
                                                        
                minwork = Val(ed_emp_timerec_out.Fields("TimeOutPm")) - Val(ed_emp_timerec_out.Fields("TimeInPm"))

                If minwork < 1 Then
                                                        
                    hours = 4
                    min = 0
                                                            
                    ed_emp_timerec_out.Fields("PTHour") = hours
                    ed_emp_timerec_out.Fields("PTMin") = min
                    ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                    ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                    ed_emp_timerec_out.Fields("Deduction") = 93.17 + (hours * rateperh) + (min * rateperm)
                                                        
                Else
                        
                    pmin = 0
                    pmin = Val(ed_emp_timerec_out.Fields("PTMin")) + min
                                                            
                    If pmin > 59 Then
                                                                
                        hours = hours + 1
                        min = pmin - 60
                                                            
                        ed_emp_timerec_out.Fields("PTHour") = Val(ed_emp_timerec_out.Fields("PTHour")) + hours
                        ed_emp_timerec_out.Fields("PTMin") = min
                                                           
                        totalmin = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        
                        If totalmin > 59 Then
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                            ed_emp_timerec_out.Fields("TTinMin") = totalmin - 60
                        Else
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                            ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        End If
                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                        
                    Else
                                                            
                        ed_emp_timerec_out.Fields("PTHour") = Val(ed_emp_timerec_out.Fields("PTHour")) + hours
                        ed_emp_timerec_out.Fields("PTMin") = Val(ed_emp_timerec_out.Fields("PTMin")) + min
                                                            
                        totalmin = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        
                        totalmin = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        
                        If totalmin > 59 Then
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + 1
                            ed_emp_timerec_out.Fields("TTinMin") = totalmin - 60
                        Else
                            ed_emp_timerec_out.Fields("TTinHour") = Val(ed_emp_timerec_out.Fields("TTinHour")) + hours
                            ed_emp_timerec_out.Fields("TTinMin") = Val(ed_emp_timerec_out.Fields("ATMin")) + Val(ed_emp_timerec_out.Fields("PTMin"))
                        End If
                                        
                        ed_emp_timerec_out.Fields("Deduction") = (Val(ed_emp_timerec_out.Fields("TTinHour")) * rateperh) + (Val(ed_emp_timerec_out.Fields("TTinMin")) * rateperm)
                                                            
                    End If
                        
                    ed_emp_timerec_out.Fields("NumDayWork") = Val(ed_emp_timerec_out.Fields("NumDayWork")) + 0.5
                    
                End If
                                                        
                ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                ed_emp_timerec_out.Fields("Count") = 4
            ed_emp_timerec_out.Update
                                            
        End If
                                    
    End If
                                
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
                                       
End If
End Sub

Public Sub Category_17()
If Form1.optOut.Value = True And Val(Form1.lblTime.Caption) Then
                                                                                    
    ed_emp_timerec_out.Requery
                            
    ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Date & "'and TimeInAm=" & Trim(Form1.lblTime.Caption) & "and LastStatus='IN' and Count=1"
                            
        If ed_emp_timerec_out.NoMatch = False Then
            MsgBox "Invalid log out time. Please clarify your authorized personnel.", vbExclamation, "Administrator"
            Exit Sub
        Else
                                  
            ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "' and LastStatus='OUT' and Count=4"
                                  
            If ed_emp_timerec_out.NoMatch = True Then
            
            pangtrap.Requery
                                    
            pangtrap.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Trim(Format(Date, "mm/dd/yyyy")) & "'and TimeInPm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutPm<=" & Trim(Form1.lblTime.Caption) & " and LastStatus='OUT' and Count=4"
                                    
            If pangtrap.NoMatch = True Then
                                        
                ed_emp_timerec_out.FindFirst "EmployeeNo='" & Trim(Form1.txtLog.Text) & "'and CDate='" & Format(Date, "mm/dd/yyyy") & "'and TimeInPm<=" & Trim(Form1.lblTime.Caption) & "and TimeOutPm<TimeInPm and LastStatus='IN' and Count=3"
                                        
                ed_emp_timerec_out.Edit
                    ed_emp_timerec_out.Fields("PMOut") = Form1.Label7.Caption
                    ed_emp_timerec_out.Fields("TimeOutPm") = Form1.lblTime.Caption
                    ed_emp_timerec_out.Fields("LastStatus") = "OUT"
                    ed_emp_timerec_out.Fields("NumDayWork") = Val(ed_emp_timerec_out.Fields("NumDayWork")) + 0.5
                    ed_emp_timerec_out.Fields("Count") = 4
                ed_emp_timerec_out.Update
                
            End If
            Else
                MsgBox "Employee Number '" & Form1.txtLog.Text & "' has been already time out.", vbExclamation, "Administrator"
            
            End If
        End If
End If
            
    Form1.ListView1.ListItems.Clear
    Form1.LoadRecord
    Form1.onFocus
End Sub


