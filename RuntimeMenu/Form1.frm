VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9600
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.CommandButton cmdShowPopupMenu 
      Caption         =   "Show Popup Menu"
      Height          =   510
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdShowPopupMenu_Click()
    Dim r As RECT
    Dim x As Long
    Dim y As Long
    
    'adjust the position of popup menu here.
    x = Me.Left + cmdShowPopupMenu.Left + cmdShowPopupMenu.Width
    y = Me.Top + cmdShowPopupMenu.Top
    x = x / Screen.TwipsPerPixelX + 10
    y = y / Screen.TwipsPerPixelY + 50
    
    Call CreateMyPopupMenu
    If TotalMenuHandle > 0 Then
        TrackPopupMenu aryMenuHandle(1), 0, x, _
             y, 0, Me.hwnd, r
    End If
End Sub
Private Function CreateMyPopupMenu() As Boolean
    
    'Now create a 2 layers menu
    Dim fso As FileSystemObject
    Dim StartPath As String

    Dim f As File
    Dim fd As Folder

    Dim StartFolder As Folder
    
    Dim menuid As Long

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim retHandle As Long
    
    menuid = MENU_ID_BASE

    Set fso = New FileSystemObject
    
    '***Change the path here you want to show in the popup menu ***
    '***It will take a some time, if there are too many folders and files to show.
    StartPath = App.Path '"C:\"
    If Not fso.FolderExists(StartPath) Then
        Exit Function
    End If
    
    Set StartFolder = fso.GetFolder(StartPath)
    
    If StartFolder.Files.Count = 0 And StartFolder.SubFolders.Count = 0 Then
        Exit Function
    End If

    If TotalMenuHandle > 0 Then
        For i = TotalMenuHandle To 1 Step -1
            If IsMenu(aryMenuHandle(i)) Then
                DestroyMenu aryMenuHandle(i)
            End If
        Next
    End If
    
    TotalMenuHandle = StartFolder.SubFolders.Count + 1
    TotalMenuItem = 0
    retHandle = CreatePopupMenu
    If retHandle = 0 Then
        Exit Function
    End If
    aryMenuHandle(1) = retHandle
    
    i = 1
    j = 0
    For Each fd In StartFolder.SubFolders
        i = i + 1
        retHandle = CreatePopupMenu
        If retHandle = 0 Then
            Exit Function
        End If
        aryMenuHandle(i) = retHandle
        'create a submenu
        Call InsertMenu(aryMenuHandle(1), j, MF_STRING Or MF_BYPOSITION Or MF_POPUP, aryMenuHandle(i), fd.Name)
        j = j + 1
        k = 0
        On Error Resume Next
        For Each f In fd.Files
            If Err.Number = 0 Then
                Call InsertMenu(aryMenuHandle(i), k, MF_STRING Or MF_BYPOSITION, menuid, f.Name)
                TotalMenuItem = TotalMenuItem + 1
                aryMenuID(TotalMenuItem) = menuid
                aryFilePath(TotalMenuItem) = f.Path
                menuid = menuid + 1
                k = k + 1
            End If
        Next
    Next
    If StartFolder.Files.Count > 0 Then
        'create a separator
        Call InsertMenu(aryMenuHandle(1), j, MF_SEPARATOR Or MF_BYPOSITION, menuid, "-")
        TotalMenuItem = TotalMenuItem + 1
        aryMenuID(TotalMenuItem) = menuid
        aryFilePath(TotalMenuItem) = "-"
        menuid = menuid + 1
        j = j + 1
        For Each f In StartFolder.Files
                Call InsertMenu(aryMenuHandle(1), j, MF_STRING Or MF_BYPOSITION, menuid, f.Name)
                TotalMenuItem = TotalMenuItem + 1
                aryMenuID(TotalMenuItem) = menuid
                aryFilePath(TotalMenuItem) = f.Path
                menuid = menuid + 1
                j = j + 1
        Next
    End If
    
End Function

Private Sub Form_Load()
    TotalMenuHandle = 0
    TotalMenuItem = 0
    Call HookPopupMenuProc(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    If TotalMenuHandle > 0 Then
        For i = TotalMenuHandle To 1 Step -1
            If IsMenu(aryMenuHandle(i)) Then
                DestroyMenu aryMenuHandle(i)
            End If
        Next
    End If
    On Error Resume Next
    'On Error GoTo ErrHandler
    Call UnHookPopupMenuProc(Me)
End Sub
