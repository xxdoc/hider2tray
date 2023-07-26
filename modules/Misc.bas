Attribute VB_Name = "Misc"
Public Declare Function GetWindowRect Lib "user32" (ByVal hWND As Long, lpRect As RECT) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWND As Long, ByVal nIndex As Long) As Long

'Global Const WS_MAXIMIZEBOX = &H10000
'Global Const WS_MINIMIZEBOX = &H20000
'Global Const WS_SYSMENU = &H80000
Public Const WM_GETICON = &H7F
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wparam As Long, lParam As Long) As Long

Public Declare Function SetFocus Lib "user32" (ByVal hWND As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWND As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWND As Long, lpdwProcessId As Long) As Long

Public Type FormArray
    vbForm As Form
    inUse As Boolean
    hWND As Long
    ThreadID As Long
    SentAwayTime As String
End Type

Global Formz(1 To 999) As FormArray
Function IsValid(hWND As Long) As Boolean
    ' Denies windows that are not supported
    ' First, Check to see if it is already minimized
    '
'******************************
'Exclude already Minimized Windows (Check ThreadID and hWnd)
'******************************
Dim PID As Long
Dim TID As Long
   DoEvents
 TID = GetWindowThreadProcessId(hWND, PID)
    For i = 1 To 999
        With Formz(i)
            If .inUse = True Then
                If .ThreadID = TID Then IsValid = False: Exit Function
                If .hWND = hWND Then IsValid = False: Exit Function
            End If
        End With
    Next
    

'******************************
'Exclude ICQ    (built in tray features)
'******************************
    Dim WC As String
    WC = Trim$(GetWindowCaption(hWND))
    
    If Val(WC) > 0 Then
        ' The window is a number. I'm guessing ICQ
        IsValid = False
        Exit Function
    End If
    
'******************************
'Window Passed.
IsValid = True
'******************************

End Function

Function GetNextFreeFormz() As Integer
    For i = 1 To 999
        If Formz(i).inUse = False Then
            GetNextFreeFormz = i
            Exit Function
        End If
    Next
    
    GetNextFreeFormz = 0
End Function
Function GetWindowCaption(hWND As Long)
    On Error Resume Next
    Dim X As String
    X = Space(255)
    GetWindowText hWND, X, 255
    
    GetWindowCaption = Left$(X, InStr(X, vbNullChar))
End Function
