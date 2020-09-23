Attribute VB_Name = "Module4"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub ReleaseCapture Lib "user32" ()

Public Declare Function Inp Lib "inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Public Declare Sub Out Lib "inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)


Declare Function Shell_NotifyIconA Lib "shell32" _
(ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_MESSAGE = &H1

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NULL = &H0

Type NOTIFYICONDATA
  cbSize              As Long
  hwnd                As Long
  uID                 As Long
  uFlags              As Long
  uCallbackMessage    As Long
  hIcon               As Long
  szTip               As String * 64
End Type
    
Public Enum SysTrayAction
   modify = 1
   Delete = 2
End Enum
Public SysTrayData As NOTIFYICONDATA

' Sets transparency of forms

Private Declare Function GetWindowLong Lib "user32" _
                         Alias "GetWindowLongA" _
                         (ByVal hwnd As Long, ByVal nIndex As Long) _
                         As Long
Private Declare Function SetWindowLong Lib "user32" _
                         Alias "SetWindowLongA" _
                         (ByVal hwnd As Long, ByVal nIndex As Long, _
                         ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                         (ByVal hwnd As Long, ByVal crey As Byte, _
                         ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Public trns As Integer

Public Sub FormDrag(TheForm As Form)
  ReleaseCapture
  Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
Public Function SysTrayInit(Tip As String, Hook As Object, hIconPic As Long) As NOTIFYICONDATA

With SysTrayInit
  .cbSize = Len(SysTrayInit)
  .hwnd = Hook.hwnd
  .uID = 1&
  .szTip = Tip & Chr(0)
  .uCallbackMessage = WM_MOUSEMOVE
  .hIcon = hIconPic
  .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
End With

 Shell_NotifyIconA NIM_ADD, SysTrayInit

End Function
Public Sub SysTrayModify(objSysData As NOTIFYICONDATA, Action As SysTrayAction, _
                        Optional Tip As String, Optional hIconPic As Long)

With objSysData
   If IsMissing(Tip) Then Tip = .szTip
   If IsMissing(hIconPic) Then hIconPic = .hIcon
End With
   Shell_NotifyIconA Action, objSysData

End Sub

Public Sub SetTrans(frm As Form, percentage As Integer)
  If percentage > 100 Or percentage < 0 Then Exit Sub

  Dim Opacity As Single
  'Set the transparency level
  Opacity = 2.55 * percentage
  Call SetWindowLong(frm.hwnd, GWL_EXSTYLE, _
  GetWindowLong(frm.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
  Call SetLayeredWindowAttributes(frm.hwnd, 0, Opacity, LWA_ALPHA)
End Sub





