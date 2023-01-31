Attribute VB_Name = "modRelatives"
' This module is dedicated to finding relative threads of a hWnd
' For Example. Visual Basic.
'              If you minimize to tray Visual Basic, it leaves a Task bar entry
'              which is because it has many thread windows which are not controlled by each other
'              Before, my app would just hide the visible thread.
'              This module should fix that. (upto 99 threads supported)
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWND As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassNameAPI Lib "user32" Alias "GetClassNameA" (ByVal hWND As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function EnumThreadWindows Lib "user32.dll" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

Private WinCount As Long
Private Type ProcessType
    ThreadID As Long
    ProcessID As Long
    hWND As Long
    hInstance As Long
    ClassName As String
End Type

Private hWnds(99) As Long

Private Const GWL_HINSTANCE = -6    ' For GetWindowLong(..)

Sub GetRelatives(hWND As Long, vbForm As Form)
        Dim CP As ProcessType   'Current Process
        
        ' Incase I need the other stuff,. I added the code and 'd it out
        With CP
            .hWND = hWND
        '    .hInstance = GetInstance(hwnd)
        '    ' Get Process ID and Thread ID with 1 call :)
            .ThreadID = GetWindowThreadProcessId(hWND, .ProcessID)
        '    .ClassName = GetClassName(hwnd)
        End With
    
        For i = 0 To 99
            hWnds(i) = 0
        Next
        WinCount = 0
        
        'Enum Thread Windows
        EnumThreadWindows CP.ThreadID, AddressOf EnumThreadWndProc, 0&
        For i = 1 To 99
            If hWnds(i) > 0 Then
                vbForm.SetbHwndsVals i, hWnds(i)
                vbForm.SeTBSizeStylesVals i, GetWindowLong(hWnds(i), GWL_STYLE)
            End If
        Next
        
        End Sub

Private Function AddNullChar(TempStr As String) As String
        
        If Right$(TempStr, Len(vbNullChar)) <> vbNullChar Then TempStr = TempStr & vbNullChar
        AddNullChar = TempStr
        End Function

Private Function GetClassName(hWND As Long) As String
        Dim TempStr As String:        Dim TempLng As Long:        Const MaxLen = 255
        
        TempStr = Space(MaxLen)
        TempLng = GetClassNameAPI(hWND, TempStr, MaxLen)
        GetClassName = Left$(TempStr, TempLng)
        End Function

Private Function GetInstance(hWND As Long) As Long
        
        GetInstance = GetWindowLong(hWND, GWL_HINSTANCE)
        End Function

Public Function EnumThreadWndProc(ByVal hWND As Long, ByVal lParam As Long) As Long
        'Static winnum As Integer  ' counter keeps track of how many windows have been enumerated
        
        Dim G As Long
        G = GetWindowLong(hWND, GWL_STYLE)
        
        If (G And WS_CHILDWINDOW) <> WS_CHILDWINDOW Then
            ' not a child window (controls)
            If GetParent(hWND) = 0 Then
                'no parent window (forms)
                    WinCount = WinCount + 1  ' one more window enumerated....
'                    Debug.Print "ThreadWnd #" & winnum & "=" & Hex(hWND)
                    If WinCount <= 99 Then hWnds(WinCount) = hWND
            End If
        End If
        
        
        EnumThreadWndProc = 1  ' return value of 1 means continue enumeration
        End Function
