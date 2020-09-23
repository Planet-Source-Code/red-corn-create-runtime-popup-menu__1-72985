Attribute VB_Name = "modRuntimeMenu"
Option Explicit
Public Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal Lparam As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long

Public Const MF_SEPARATOR = &H800&
Public Const MF_BYPOSITION = &H400&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const GWL_WNDPROC = (-4)
Public Const WM_COMMAND = &H111
Public Const WM_CLOSE = &H10

'Hold Menu's Handle
Public aryMenuHandle(1 To 2000) As Long
Public TotalMenuHandle As Integer
'Hold Menu's ID & file path
Public TotalMenuItem As Integer
Public aryFilePath(1 To 2000) As String
Public aryMenuID(1 To 2000) As Long
Public Const MENU_ID_BASE As Long = 100


'Hold the address of the old window procedure
Dim LocalPrevWndProc As Long


Public Sub HookPopupMenuProc(PassedForm As Form)
  On Error Resume Next
  LocalPrevWndProc = SetWindowLong(PassedForm.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookPopupMenuProc(PassedForm As Form)
  Dim WorkFlag As Long

  On Error Resume Next
  WorkFlag = SetWindowLong(PassedForm.hwnd, GWL_WNDPROC, LocalPrevWndProc)
End Sub

Public Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal Lparam As Long) As Long
    Select Case Lmsg
    Case WM_COMMAND:
        'Whenever user clicks a menu, WM_COMMAND is sent to the window
        'And the wParam hold the menu's item ID
         If wParam >= MENU_ID_BASE And wParam <= (TotalMenuItem + MENU_ID_BASE - 1) Then
             MsgBox aryFilePath(wParam - MENU_ID_BASE + 1)

         End If
    
    End Select
    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, wParam, Lparam)
End Function


