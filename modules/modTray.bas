Attribute VB_Name = "modTray"
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWND As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA_5) As Boolean
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWND As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWND As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWND As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DllGetVersion Lib "shell32" (ByRef pdvi As DLLVERSIONINFO) As Long

' Used for cbSize if old DLL installed
Public Type NOTIFYICONDATA
        cbSize As Long
        hWND As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type NOTIFYICONDATA_5
  cbSize As Long
  hWND As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128  ' 64 for non v5 DLLs. will send a null terminated string.
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutOrVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type


Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

        
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'Public Const SW_SHOWMINIMIZED = 2
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
Public Const SW_MINIMIZE = 6
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNORMAL = 1

'Private Const WS_VISIBLE = &H10000000
Private Const GWL_STYLE = -16
'Private Const WS_MINIMIZE = &H20000000
Const WS_BORDER = &H800000
Const WS_CAPTION = &HC00000
Const WS_CHILD = &H40000000
Const WS_CHILDWINDOW = &H40000000
Const WS_CLIPCHILDREN = &H2000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_DISABLED = &H8000000
Const WS_DLGFRAME = &H400000
Const WS_GROUP = &H20000
Const WS_HSCROLL = &H100000
Const WS_ICONIC = &H20000000
Const WS_MAXIMIZE = &H1000000
Const WS_MAXIMIZEBOX = &H10000
Const WS_MINIMIZE = &H20000000
Const WS_MINIMIZEBOX = &H20000
Const WS_OVERLAPPED = &H0
Const WS_OVERLAPPEDWINDOW = &HCF0000
Const WS_POPUP = &H80000000
Const WS_POPUPWINDOW = &H80880000
Const WS_SIZEBOX = &H40000
Const WS_SYSMENU = &H80000
Const WS_TABSTOP = &H10000
Const WS_THICKFRAME = &H40000
Const WS_TILED = &H0
Const WS_TILEDWINDOW = &HCF0000
Const WS_VISIBLE = &H10000000
Const WS_VSCROLL = &H200000
Private Const WM_LBUTTONDBLCLICK = &H203
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONUP = &H205
Private Const WM_MOUSEMOVE = &H200

Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4
Const NIF_STATE = &H8
Const NIF_INFO = &H10
Const NIS_HIDDEN = &H1
Const NIS_SHAREDICON = &H2
Const NOTIFYICON_VERSION = &H1
Const NIIF_WARNING = &H30
Const NIIF_ERROR = &H10
Const NIIF_INFO = &H40


Sub SendToTray(vbForm As Form, hWND As Long, TrayCaption As String, HideHWND As Boolean, HideForm As Boolean, IconH As Long)
  Dim TC    As String
  Dim xTray As NOTIFYICONDATA_5
  
  On Error Resume Next
  ' Trim the Caption
           
  If IsV5Compat = True And Len(TrayCaption) > 120 Then
    TC = Left$(TrayCaption, 90) & ".." & Right$(TrayCaption, 20)
  ElseIf IsV5Compat = False And Len(TrayCaption) > 63 Then
    TC = Left$(TrayCaption, 40) & ".." & Right$(TrayCaption, 20)
  Else
    TC = TrayCaption
  End If
           
  TC = TC & vbNullChar
           
  ' Now, Build the Struct up
  With xTray
    .cbSize = GetLen()                              ' Length of Struct. (See GetLen())
    .dwInfoFlags = NIIF_INFO                        ' v5 DLLs, Info Icon
    .dwState = NIS_SHAREDICON                       ' v5 DLLs, Icon is shared
    .dwStateMask = 0&                               ' v% DLLs, n/a
    .hIcon = IconH                                  ' Handle to icon
    .hWND = vbForm.hWND                             ' Handle of Owner Window
    .szInfo = TrayCaption & vbNullChar              ' Balloon Text
                                                    ' Balloon Caption (in bold)
    .szInfoTitle = "Para restaurar este Programa, click aqui:" & vbNullChar
    .szTip = TC                                     ' normal Tool Tip Caption
    .uCallbackMessage = WM_MOUSEMOVE                ' Callback methods
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE    ' flags
    If IsV5Compat = True Then .uFlags = .uFlags Or NIF_INFO
                
    .uID = vbNull                                   ' unique ID.. n/a
    .uTimeoutOrVersion = 1000&                      ' v% DLLs, timeout value for balloon..  (in MS.. will default to min value)
  End With
        
  If HideForm = True Then
'   App.TaskVisible = False
    vbForm.Hide
  End If
            
            
  ' Save the Style settings for when resetting it.
  Dim xStyle As Long
        
  xStyle = GetWindowLong(hWND, GWL_STYLE)
  vbForm.Tag = Trim$(Str$(xStyle))
  
  If HideHWND = True Then
    ' First, Minimize it if not already
    If (xStyle And WS_MINIMIZE) <> WS_MINIMIZE Then
      xStyle = (xStyle Xor WS_MINIMIZE)
                    
      'SetWindowLong hWnd, GWL_STYLE, xStyle
      'DoEvents    ' let it minimize
      ShowWindow hWND, SW_MINIMIZE
      DoEvents
    End If
            
    If (xStyle And WS_VISIBLE) = WS_VISIBLE Then
      ' Is visible, take it out
      'xStyle = (xStyle Xor WS_VISIBLE)
      'SetWindowLong hWnd, GWL_STYLE, xStyle
      ShowWindow hWND, SW_HIDE
    End If
  End If
        
  ' Cool, now add the Tray icon
  Shell_NotifyIcon NIM_ADD, xTray
End Sub
        
Sub UpdateTray(vbForm As Form, TrayCaption As String)
  Dim xTray As NOTIFYICONDATA_5
  
  On Error Resume Next
  
  With xTray
    .cbSize = GetLen()
    .hWND = vbForm.hWND
    .uID = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = vbForm.Icon    ' handle of iCon
    .szTip = TrayCaption & vbNullChar
  End With
  
  Shell_NotifyIcon NIM_MODIFY, xTray

End Sub

Sub KillTray_And_RestoreHwnd(vbForm As Form, hWND As Long, ShowForm As Boolean, SetHWNDFocus As Boolean)
    On Error Resume Next
    Dim xTray As NOTIFYICONDATA_5
        With xTray
            .cbSize = GetLen()
            .hWND = vbForm.hWND
            .uID = vbNull
        End With
        
        ' Remove the Tray Icon
        Shell_NotifyIcon NIM_DELETE, xTray
        
        ' Restore the Window and set focus
        
        Dim xStyle As Long
        xStyle = Val(Trim$(vbForm.Tag))
        
        If (xStyle And WS_VISIBLE) = WS_VISIBLE Then
            ShowWindow hWND, SW_SHOW
        End If
        
        If (xStyle And WS_MINIMIZE) = WS_MINIMIZE Then
            ShowWindow hWND, SW_SHOWMINIMIZED
        ElseIf (xStyle And WS_MAXIMIZE) = WS_MAXIMIZE Then
            ShowWindow hWND, SW_SHOWMAXIMIZED
        Else
            ShowWindow hWND, SW_SHOWNORMAL
        End If
        
        If ShowForm = True Then
            vbForm.Show
        End If
        
        If SetHWNDFocus = True Then
            SetForegroundWindow hWND
            SetFocus hWND
        End If
        
        
End Sub

Public Function IsV5Compat(Optional OverRide As Boolean = False) As Boolean
  'Purpose: Get Version info of Shell32.dll (and other main DLLs..all same version#)
  Dim X As DLLVERSIONINFO
  On Error GoTo Sair:
  X.cbSize = Len(X)
  DllGetVersion X
  
  IsV5Compat = False
    
  ' Check Options for Balloon Value.. if disabled, return no v5 dlls.
  
  
  If Val(Trim$(GetSetting("MIN2TRAY", "Options", "UseBalloon", "1"))) > 0 Then
        If X.dwMajorVersion >= 5 Then IsV5Compat = True 'only if dll is v5 =)
  End If
      
   If OverRide = True And X.dwMajorVersion >= 5 Then IsV5Compat = True
   Exit Function
Sair:
'   MsgBox Err.Description, vbCritical + vbMsgBoxHelpButton, "Erro...: " & Err.Source & ".Tray_lsV5Compat", Err.HelpFile, Err.HelpContext
End Function

Function GetLen() As Long
    If IsV5Compat Then
        Dim X As modTray.NOTIFYICONDATA_5
        GetLen = Len(X)
    Else
        Dim Y As modTray.NOTIFYICONDATA
        GetLen = Len(Y)
    End If
End Function


Function HideHWND(hWND As Long) As Long
On Error Resume Next
    HideHWND = 0
    If hWND <= 0 Then Exit Function
    
    Dim G As Long
    
    G = GetWindowLong(hWND, GWL_STYLE)
    
    If (G And WS_VISIBLE) = WS_VISIBLE Then
        ShowWindow hWND, SW_HIDE
    End If
    
    HideHWND = G
    End Function

Sub ShowHWND(hWND As Long, Style As Long)
On Error Resume Next
    If hWND <= 0 Then Exit Sub
    
    Dim G As Long
    
    G = GetWindowLong(hWND, GWL_STYLE)
    
    If (Style And WS_VISIBLE) = WS_VISIBLE And (G And WS_VISIBLE) <> WS_VISIBLE Then
        ShowWindow hWND, SW_SHOW
    End If
    End Sub
