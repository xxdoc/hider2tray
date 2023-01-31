Attribute VB_Name = "modICON"
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWND As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWND As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Private Declare Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASSEX) As Long

Private Type WNDCLASSEX     ' Same as WNDCLASS but has a few advanced values
    cbSize As Long
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long               ' Handle to large icon (Alt-Tab icon)
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long             ' Handle to Small icon (Top Left Icon/Taskbar Icon)
    End Type

Private Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long               ' Handle to icon (only 1 size)
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    End Type

Private Const GWL_HINSTANCE = -6    ' For GetWindowLong(..)
Private Const GCL_HICON = -14       ' For GetClassLong(..)
  
Public Function GetIconHandle(hWND As Long) As Long
' OK, This function is confusing
' Many windows have different ways of handling Icons.
'---------------------------
'1. All VB apps use SendMessage(..WM_GETICON..) to get the Icon (not only VB apps)
'2. Other Programs like GetClassInfoEx(..)                      (most non-SendMessage Apps)
'3. Others Like GetClassInfo(..)                                (Very Rare)
'4. And the rest like GetClassLong(GCL_HICON)                   (The rest.)
'----------------------------
' Any program that doesn't work with these 4 methods have issues.
'
' All apps I have tried work fine with these 4 methods.. one or the other.
'

  
'*************************************
'Method: SendMessage (Small Icon)
'*************************************
  Dim hIcon As Long
  'frmMain.Text1.Text = ""
  ' First, Try for the small icon. This would be nice.
  hIcon = SendMessage(hWND, WM_GETICON, CLng(0), CLng(0))
  
  If hIcon > 0 Then GetIconHandle = hIcon: Exit Function  ' found it
  ' Nope, keep trying
    
   
'*************************************
'Method: SendMessage (Large Icon)
'*************************************
   ' Hmm.. No small Icon, Try LARGE icon.
   hIcon = SendMessage(hWND, WM_GETICON, CLng(1), CLng(0))
   
  If hIcon > 0 Then GetIconHandle = hIcon: Exit Function  ' found it
  ' Nope, keep trying
    
    
'*************************************
'Method: GetClassInfoEx (Small or Large with Small Pref.)
'*************************************
    
    Dim ClassName As String
    Dim WCX As WNDCLASSEX
    Dim hInstance As Long
    
    ' First, get the Instance of the Class via GetWindowLong
    hInstance = GetWindowLong(hWND, GWL_HINSTANCE)
    
    ' Now set the Size Value of WndClassEx
    WCX.cbSize = Len(WCX)
    
    ' Set The ClassName variable to 255 spaces (max len of the class name)
    ClassName = Space(255)
    
    Dim X As Long   ' temp variable
    ' Get the Classname of hWnd and put into ClassName (max 255 chars)
    X = GetClassName(hWND, ClassName, 255)
    
    ' Now Trim the Classname and add a NullChar to the end (reqd. for GetClassInfoEx)
    ClassName = Left$(ClassName, X) & vbNullChar
    
    ' Now, if GetClassInfoEx(..) Returns 0, their was an error. >0 = No probs
    X = GetClassInfoEx(hInstance, ClassName, WCX)
    If X > 0 Then
        ' Returned True
        ' So we should now have both WCX.hIcon and WCX.hIconSm
        If WCX.hIconSm = 0 Then 'No small icon
            hIcon = WCX.hIcon ' No small icon.. Windows should have given default.. weird
        Else
            hIcon = WCX.hIconSm ' Small Icon is better
        End If
        GetIconHandle = hIcon   ' found it =]
        Exit Function
        
    End If
    
    
'*************************************
'Method: GetClassInfo (Large Icon)
'*************************************
        
        ' Hmm.. ClassInfoEX failed, Try ClassInfo
        Dim WC As WNDCLASS
        X = GetClassInfo(hInstance, ClassName, WC)
        If X > 0 Then
            ' Woohoo.. dunno why but it liked that
            hIcon = WC.hIcon
            GetIconHandle = hIcon: Exit Function    ' Found it
        End If
        
        
'*************************************
'Method: GetClassLong (Large Icon)
'*************************************
            ' Hmm.. One more try
            X = GetClassLong(hWND, GCL_HICON)
            If X > 0 Then
                ' Yay, about time.. annoying windows.. Example: NOTEPAD
                hIcon = X
            Else
                ' This is most prob a Icon-less window.
                 hIcon = 0
            End If

If hIcon < 0 Then hIcon = 0     ' Handles must be > 0
GetIconHandle = hIcon
End Function
