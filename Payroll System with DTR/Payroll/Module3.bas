Attribute VB_Name = "Module3"
Option Explicit

'----------------------------------------------------
Global cn                     As New ADODB.Connection
'----------------------------------------------------


'----------------------------------------------------
Global rs                     As New ADODB.Recordset
'----------------------------------------------------
Global sRate                  As New ADODB.Recordset
'----------------------------------------------------
Global sMedRate               As New ADODB.Recordset
'----------------------------------------------------
Global sAllowanceRate         As New ADODB.Recordset
'----------------------------------------------------
Global sLifeRetRate           As New ADODB.Recordset
'----------------------------------------------------
Global sLoans                 As New ADODB.Recordset
'----------------------------------------------------
Global sGeneratePayroll       As New ADODB.Recordset
'----------------------------------------------------
Global sLoadName              As New ADODB.Recordset
'----------------------------------------------------
Global sGeneratePayrollJO     As New ADODB.Recordset
'----------------------------------------------------
Global sCasualPay             As New ADODB.Recordset
'----------------------------------------------------
Global sUsers                 As New ADODB.Recordset
'----------------------------------------------------
Global sUserLogs              As New ADODB.Recordset
'----------------------------------------------------
Global CasualPay              As New ADODB.Recordset


Global PeriodNow              As String

Global sPeriodUse             As String

Global sProjectName           As String

Public sCurrUser              As String
