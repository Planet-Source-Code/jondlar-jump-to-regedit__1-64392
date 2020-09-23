Attribute VB_Name = "modRegEditJump"
'########################################################
'Author: Sudesh Katariya
'PSC ID: GreenEye2oo4
'
'Updated: 21st Feb 2006
'
'This code is VB port of original code in VC by Mark
'Russinovich(www.sysinternals.com). I tried to find it's VB
'port but couldnt find. So here my contribution to PSC.
'########################################################


'=========================================================
'Win32 API Declaration
'=========================================================
'ShellExecute API
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long
Private Const SW_NORMAL As Long = 1
Private Const SW_SHOW As Long = 5
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &H40
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type


'Window API
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long


'Message API
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_CHAR As Long = &H102
Private Const VK_LEFT As Long = &H25
Private Const VK_RIGHT As Long = &H27
Private Const VK_HOME As Long = &H24

'Process API
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function WaitForInputIdle Lib "user32.dll" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Const SYNCHRONIZE As Long = &H100000
Private Const INFINITE As Long = &HFFFFFFFF
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

'Sleep
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


'==========================================================
Public Function RegEditJump(ByVal strRegKey As String, Optional ByVal strValueName As String)
'Function RegEditJump
'
'Input values:
'strRegKey As String : Registry key to open
'strValueName As String : Registry Value name to select[optional]
'==========================================================
    Dim hndRegEdit As Long 'Regedit Window's Handle
    Dim hndTreeView As Long 'TreeView in Regedit's Handle
    Dim hndListView As Long 'ListView Handle
    Dim hndProcess As Long 'RegEdit Process
    Dim SHInfo As SHELLEXECUTEINFO
    Dim j As Integer
    Dim strTmp As String
    Dim lngVK As Long
    Dim ret As Long
    
    'First we check if RegEdit is running or not?
    'If not we start it.
    'To see if it is running, try to find Window with class
    'RegEdit_RegEdit
    '-----------Open RegEdit-------------------------------
    hndRegEdit = FindWindow("RegEdit_RegEdit" & vbNullChar, vbNullString)
    If hndRegEdit = 0 Then 'Not running
        With SHInfo
            .cbSize = Len(SHInfo)
            .lpVerb = "open" & vbNullChar
            .lpFile = "regedit.exe" & vbNullChar
            .fMask = SEE_MASK_NOCLOSEPROCESS
            .nShow = SW_NORMAL
        End With
        ShellExecuteEx SHInfo 'ShellExecute RegEdit
    End If
    
    'Sleep for about one second, so RegEdit can start
    'Because we will check for its Window again
    Sleep (750)
    
    'Check for RegEdit window
    hndRegEdit = FindWindow("RegEdit_RegEdit" & vbNullChar, vbNullString)
    
    'If still not running then show message that we cannot run it
    'And exit :(
    If hndRegEdit = 0 Then
        MsgBox "Unable to launch RegEdit.", vbCritical, "Error"
        Exit Function
    End If
    
    
    'If now RegEdit is running then put its window in foreground
    ret = ShowWindow(hndRegEdit, SW_SHOW)
    ret = SetForegroundWindow(hndRegEdit)
    
    
    'Now we will get handle of TreeView in RegEdit's main Window
    'This we do by finding a child window with class SysTreeView32
    '------------Get TreeView------------------------------
    hndTreeView = FindWindowEx(hndRegEdit, 0, "SysTreeView32" & vbNullChar, vbNullString)
    
    'Set focus to treeview
    ret = SetForegroundWindow(hndTreeView)
    ret = SetFocus(hndTreeView)
    
    'We need Process Handle of TreeView so we can make it wait
    'after inputing KeyStrokes in it.
    
    'Get ProcessID of TreeView
    ret = GetWindowThreadProcessId(hndTreeView, hndProcess)
    'Get Process Handle of TreeView
    ret = OpenProcess(PROCESS_ALL_ACCESS, False, hndProcess)
    
    
    'Now we close TreeView by sending 'Left' keys to it.
    'This we do as a precaution, in case treeview is already open
    '------------Close TreeView----------------------------
    'Just send 30 'Left' Keys to TreeView
    For j = 1 To 30
        ret = SendMessage(hndTreeView, WM_KEYDOWN, VK_LEFT, 0)
    Next
    ret = WaitForInputIdle(hndProcess, INFINITE) 'Wait for input process to finish
    
    'Open MyComputer Tree(Root Branch)
    ret = SendMessage(hndTreeView, WM_KEYDOWN, VK_RIGHT, 0)
    ret = WaitForInputIdle(hndProcess, INFINITE) 'Wait for input process to finish
    
    'Now we will open Registry Path by inputing key strokes in TreeView
    'We send equivalent Virtual Key codes  for all chars
    'And a 'Right' Key stroke for '\', to open a node
    '------------Open Path---------------------------------
    For j = 1 To Len(strRegKey)
        strTmp = Mid(strRegKey, j, 1) 'Get a char
        'If '\' then send 'Right'
        If strTmp = "\" Then
            ret = SendMessage(hndTreeView, WM_KEYDOWN, VK_RIGHT, 0)
            ret = WaitForInputIdle(hndProcess, INFINITE)
        Else 'Send VirtualKey Code
            lngVK = Asc(UCase(strTmp))
            ret = SendMessage(hndTreeView, WM_CHAR, lngVK, 0)
        End If
    Next
    
    'Close TreeView Process Handle
    ret = CloseHandle(hndProcess)
    
    
    'If strValueName is specified. Select ValueName in ListView
    '-----------------Select Registry Value------------------
    If strValueName <> "" Then
        'First Get Handle of ListView in RegEdit
        hndListView = FindWindowEx(hndRegEdit, 0, "SysListView32" & vbNullChar, vbNullString)
        
        'Set ListView in focus
        ret = SetForegroundWindow(hndListView)
        ret = SetFocus(hndListView)
        
        'Give time to adjust
        Sleep (1500)
        
        'Select first item in ListView
        ret = SendMessage(hndListView, WM_KEYDOWN, VK_HOME, 0)
        
        'Select Value
        For j = 1 To Len(strValueName)
            lngVK = Asc(UCase(strValueName))
            ret = SendMessage(hndListView, WM_CHAR, lngVK, 0)
        Next
    End If
    
    'Finally again set focus to Main RegEdit window
    ret = SetForegroundWindow(hndRegEdit)
    ret = SetFocus(hndRegEdit)
End Function
