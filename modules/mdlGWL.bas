Attribute VB_Name = "GWL"

Option Explicit
'Styles
Public Declare Function GetParent Lib "user32" (ByVal hWND As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32.dll" (ByVal hWND As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Buttons
'   WS_MAXIMIZEBOX Creates a window that has a maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
'   WS_MINIMIZEBOX Creates a window that has a minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
'   WS_SIZEBOX Creates a window that has a sizing border. Same as the WS_THICKFRAME style.
'   WS_SYSMENU Creates a window that has a window menu on its title bar. The WS_CAPTION style must also be specified.

'Boarders
'   WS_BORDER Creates a window that has a thin-line border.
'   WS_DLGFRAME Creates a window that has a border of a style typically used with dialog boxes. A window with this style cannot have a title bar.
'   WS_THICKFRAME Creates a window that has a sizing border. Same as the WS_SIZEBOX style.
'   WS_EX_TOOLWINDOW App Window.

'Other
'   WS_CAPTION Creates a window that has a title bar (includes the WS_BORDER style).
'   WS_VISIBLE Creates a window that is initially visible. This style can be turned on and off by using ShowWindow or SetWindowPos.



'Metrics

'Buttons
'***********
    'SM_CXSIZE,
    'SM_CYSIZE
    'Width and height, in pixels, of a button in a window's caption or title bar.

'***********

    'SM_CXSMSIZE
    'SM_CYSMSIZE
    'Dimensions, in pixels, of small caption buttons.

'***********
'Boarders
'***********

    'SM_CXBORDER,
    'SM_CYBORDER
    'Width and height, in pixels, of a window border. This is equivalent to the
    'SM_CXEDGE value for windows with the 3-D look.

'***********

    'SM_CXDLGFRAME,
    'SM_CYDLGFRAME
    'Same as SM_CXFIXEDFRAME and SM_CYFIXEDFRAME.
    
    'SM_CXFIXEDFRAME,
    'SM_CYFIXEDFRAME
    'Thickness, in pixels, of the frame around the perimeter of a window that
    'has a caption but is not sizable. SM_CXFIXEDFRAME is the width of the
    'horizontal border and SM_CYFIXEDFRAME is the height of the vertical border.

'***********

    'SM_CXEDGE,
    'SM_CYEDGE
    'Dimensions, in pixels, of a 3-D border. These are the 3-D counterparts
    'of SM_CXBORDER and SM_CYBORDER.

'***********
    'SM_CXFRAME,
    'SM_CYFRAME
    'Sam as SM_CXSIZEFRAME and SM_CYSIZEFRAME.

    'SM_CXSIZEFRAME,
    'SM_CYSIZEFRAME
    'Thickness, in pixels, of the sizing border around the perimeter of a window
    'that can be resized. SM_CXSIZEFRAME is the width of the horizontal border,
    'and SM_CYSIZEFRAME is the height of the vertical border.
'***********

    
'Other
    'SM_CXMENUSIZE,
    'SM_CYMENUSIZE
    'Dimensions, in pixels, of menu bar buttons, such as the child window close
    'button used in the multiple document interface.

'***********

    'SM_CXSMICON,
    'SM_CYSMICON
    'Recommended dimensions, in pixels, of a small icon. Small icons typically
    'appear in window captions and in small icon view.

'***********
    'SM_CYMENU
    'Height, in pixels, of a single-line menu bar.


'Caption
'***********
    'SM_CYCAPTION
    'Height, in pixels, of a normal caption area.
'***********

    'SM_CYSMCAPTION
    'Height, in pixels, of a small caption.
    
'***********



' So, to get the Y location of the Button:
' Find Caption Size - Big or Small
' Find Border Size add it to 1/2 the Caption Size
' Then Take Away 1/2 the Size of a Button for the Caption Size
' and bam. Y location

'#define WS_OVERLAPPEDWINDOW (WS_OVERLAPPED     | \
 '                            WS_CAPTION        | \
  '                           WS_SYSMENU        | \
   '                          WS_THICKFRAME     | \
    '                         WS_MINIMIZEBOX    | \
     '                        WS_MAXIMIZEBOX)
'Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const WS_BORDER = &H800000
Const WS_CAPTION = &HC00000
Const WS_CHILD = &H40000000
Global Const WS_CHILDWINDOW = &H40000000
Const WS_CLIPCHILDREN = &H2000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_DISABLED = &H8000000
Const WS_DLGFRAME = &H400000
Const WS_GROUP = &H20000
Const WS_HSCROLL = &H100000
Const WS_ICONIC = &H20000000
Global Const WS_MAXIMIZE = &H1000000
Const WS_MAXIMIZEBOX = &H10000
Global Const WS_MINIMIZE = &H20000000
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
Global Const WS_VISIBLE = &H10000000
Const WS_VSCROLL = &H200000

Const SM_CYCAPTION = 4
Const SM_CYSMCAPTION = 51

Const SM_CYEDGE = 46
Const SM_CXEDGE = 45

Const SM_CYBORDER = 6
Const SM_CXBORDER = 5

Const SM_CYSMSIZE = 53
Const SM_CXSMSIZE = 52

Const SM_CYSIZEFRAME = 33
Const SM_CXSIZEFRAME = 32

Const SM_CYSIZE = 31
Const SM_CXSIZE = 30


Const WS_EX_WINDOWEDGE = &H100  '       0x00000100L
Const WS_EX_TOOLWINDOW = &H80
Const WS_EX_STATICEDGE = &H20000
Const WS_EX_CLIENTEDGE = &H200
Const WS_EX_CONTEXTHELP = &H400
Public Const WS_EX_TOPMOST = &H8

Public Const GWL_STYLE = -16
Public Const GWL_EXSTYLE = -20

'Const WS_TILEDWINDOW = &HCF0000
'Const WS_OVERLAPPEDWINDOW = &HCF0000
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function MoveWindow Lib "user32.dll" (ByVal hWND As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
        End Type


Sub GetPos(hWND As Long, vbForm As Form)
' This routine positions the form and resizes it to correct size

    Dim xStyle As Long
    Dim ExStyle As Long
    
    xStyle = GetWindowLong(hWND, GWL_STYLE)
    ExStyle = GetWindowLong(hWND, GWL_EXSTYLE)
    
   
' Oops.. First check to see if the window is visible
'******************************
'If Not (xStyle And WS_VISIBLE) Then
'    vbForm.Visible = False: Exit Sub
'End If
   
   
'Get The Boarder Size
'*******************************
    Dim BHeight As Long
    
    BHeight = GetBoarderSize(xStyle, ExStyle)
    
    If BHeight = 0 Then vbForm.Visible = False: Exit Sub
'*******************************

    
    
'Get Caption Size + Button Width/Height
'*******************************
    Dim bCaptionSize As Long
    Dim ButWidth As Long
    Dim ButHeight As Long
    
    bCaptionSize = GetCaptionSize(xStyle, ExStyle, ButWidth, ButHeight)
    
    If bCaptionSize = 0 Then vbForm.Visible = False:  Exit Sub
 '*******************************

    
    
'Find the X location of the button
'*******************************
    Dim BLeft As Long
    BLeft = GetLeftPos(hWND, ButWidth)
    If BLeft = 0 Then vbForm.Visible = False: Exit Sub
        
'*******************************
 

'Find the Y location of the button
'*******************************
    Dim BTop As Long
    Dim hWndRECT As RECT
    GetWindowRect hWND, hWndRECT
    BTop = hWndRECT.Top + (BHeight + ((bCaptionSize - ButHeight) / 2))
    
    
'Move the form, and make it ontop
'*******************************
    Dim NewRECT As RECT
    With NewRECT
        .Left = BLeft
        .Top = BTop
        .Right = .Left + ButWidth - 1
        .Bottom = .Top + ButHeight - 1
        
        ' Make it ontop and Resize it at the same time =)
'        SetWindowPos vbForm.hwnd, -1, .Left, .Top, .Right, .Bottom, 0
        vbForm.Visible = True
        MoveWindow vbForm.hWND, .Left, .Top, .Right - .Left, .Bottom - .Top, 1
        End With
        
        BuildButton vbForm
    
    ' Cool, now we have the border height, the caption height and the Button RECT
    ' So, lets build us a button :~)
    
    'y = (BHeight * 15) + (((bCaptionSize * 15) - (ButHeight * 15)) / 2) + 15
    'Dim xRECT As RECT
    'GetWindowRect hwnd, xRECT
 '   GetY = (15 * xRECT.Top) + Y
    
        Dim W As Long
        W = frmOnTop.Width / 15
        Dim H As Long
        H = frmOnTop.Height / 15
        
End Sub
    
'End Function
Sub BuildButton(A As Form, Optional MD As Boolean = False)
    'GetY A.hWnd
        Dim W As Long
        Dim H As Long
A.Cls

If MD = False Then
    ' Build normal
    With A
        ' Width + Height are in stupid VB measurements.. put to pixels
        A.ScaleMode = 3 ' pixel
        W = (.Width / 15) - 1       ' Work from 0 so loose 1 =)
        H = (.Height / 15) - 1
        
        ' Draw Black Back
        A.Line (0, 0)-(W, H), RGB(64, 64, 64), BF
                
        ' Draw dark Grey Outline
        A.Line (1, 1)-(W - 1, H - 1), RGB(128, 128, 128), BF
        
        ' Draw Grey Middle
        A.Line (1, 1)-(W - 2, H - 2), RGB(212, 208, 200), BF
        
        ' Draw White lines top + left
        
        A.Line (0, 0)-(W, 0), RGB(255, 255, 255)
        A.Line (0, 0)-(0, H), RGB(255, 255, 255)
    End With
        
Else
    ' Build with Button down
    With A
        ' Width + Height are in stupid VB measurements.. put to pixels
        A.ScaleMode = 3 ' pixel
        W = (.Width / 15) - 1       ' Work from 0 so loose 1 =)
        H = (.Height / 15) - 1
        
        ' Draw Black Back
        A.Line (0, 0)-(W, H), RGB(64, 64, 64), BF
                
        ' Draw dark Grey Outline
        A.Line (1, 1)-(W - 1, H - 1), RGB(128, 128, 128), BF
        
        ' Draw Grey Middle
        A.Line (2, 2)-(W - 1, H - 1), RGB(212, 208, 200), BF
        
        ' Draw White lines top + left
        
        A.Line (W, H)-(W, 0), RGB(255, 255, 255)
        A.Line (W, H)-(0, H), RGB(255, 255, 255)
    End With
End If
    
Const col = vbBlack
Const X = 3

Dim M As Long
Dim MM As Long
MM = H - X
M = W / 2
If W > 13 Then
    A.DrawWidth = 2
Else
    A.DrawWidth = 1
End If


' Draw small square bottom right of window 4x4
A.Line (W - X, H - X)-(W - X - 1, H - X - 1), vbBlack, BF

    
End Sub


Function GetBoarderSize(xStyle As Long, ExStyle As Long) As Long

    
    If (xStyle And WS_THICKFRAME) = WS_THICKFRAME And (ExStyle And WS_EX_TOOLWINDOW) <> WS_EX_TOOLWINDOW Then
            ' Re-Sizeable Window
            GetBoarderSize = GetSystemMetrics(SM_CYSIZEFRAME)
    
    ElseIf (ExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then
            ' Normal Window
            GetBoarderSize = GetSystemMetrics(SM_CYEDGE) + 1
    
    ElseIf (xStyle And WS_BORDER) = WS_BORDER Then
            ' Single Boarder, Will fail next routine in 99% of cases
            GetBoarderSize = GetSystemMetrics(SM_CYBORDER)
    
    Else
            ' No Boarder, Exit Function
            GetBoarderSize = 0
            Exit Function
    End If


End Function

Function GetCaptionSize(xStyle As Long, ExStyle As Long, ByRef ButWidth As Long, ByRef ButHeight As Long) As Long
    ' Valid Options:
    '  Small Caption    (Tool Windows Etc)  WS_EX_TOOLWINDOW
    '  Large Caption    (Normal Windows)    WS_CAPTION or WS_OVERLAPPEDWINDOW
    '  No Caption
    
    If (ExStyle And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW Then
            ' Tool Bar Window
            ' Get Height of Caption
            GetCaptionSize = GetSystemMetrics(SM_CYSMCAPTION)
            ButHeight = GetSystemMetrics(SM_CYSMSIZE) - 3
            ButWidth = GetSystemMetrics(SM_CXSMSIZE) - 1
            
    ElseIf (xStyle And WS_CAPTION) = WS_CAPTION Or (xStyle And WS_OVERLAPPEDWINDOW) = WS_OVERLAPPEDWINDOW Or (xStyle And WS_TILEDWINDOW) = WS_TILEDWINDOW Then
            'Normal Caption
            GetCaptionSize = GetSystemMetrics(SM_CYCAPTION)
            ButHeight = GetSystemMetrics(SM_CYSIZE) - 3
            ButWidth = GetSystemMetrics(SM_CXSIZE) - 1
            
    Else
            'No Caption, Abort
            GetCaptionSize = 0
            Exit Function
    End If

End Function

Function GetLeftPos(hWND As Long, ButWidth As Long)
    ' This gets the Windows Long Style and checks for boxes already visible.
    Dim xStyle As Long          ' Style
    Dim ExStyle As Long         ' Style EX
    Dim xRECT As RECT           ' Windows X,Y
    Dim BoarderSize As Long     ' Right boarder
    Dim X As Long               ' Temp X for ret value
    
    xStyle = GetWindowLong(hWND, GWL_STYLE)
    ExStyle = GetWindowLong(hWND, GWL_EXSTYLE)
    GetWindowRect hWND, xRECT
    
    ' Cool.. now first, work out the Right most side.
    
    If (xStyle And WS_THICKFRAME) = WS_THICKFRAME Then
            ' Re-Sizeable Window
            BoarderSize = GetSystemMetrics(SM_CXSIZEFRAME)
    
    ElseIf (ExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then
            ' Normal Window
            BoarderSize = GetSystemMetrics(SM_CXEDGE)
    
    ElseIf (xStyle And WS_BORDER) = WS_BORDER Then
            ' Single Boarder, Will fail next routine in 99% of cases
            BoarderSize = GetSystemMetrics(SM_CXBORDER)
    
    Else
            ' No Boarder, Exit Function
            GetLeftPos = 0
            Exit Function
    End If
    
    
    ' OK, so now we have the boarder size.
    X = BoarderSize - 2     ' 2 Pixels left is the first one.
    
    X = xRECT.Right - X     ' Now we should have X = right side of First button
    
    If (xStyle And WS_SYSMENU) = WS_SYSMENU Then
            ' X is there
            X = X - ButWidth - 2        ' X has 2 pixels on each side
    Else
            ' NO SYS MENU!!! Return ZERO
            ' If a form does not have a system menu, they do not want a min to tray button!
            ' IE GAMES, Taskbars.. They have borders but no buttons.
            GetLeftPos = 0
            Exit Function
    End If
    
    If (xStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX Or (xStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then
            ' Either MAX/RESIZE or MIN button is there and enabled.
            ' (or both.. but can't have 1 without the other.. 1 is just enabled)
            X = X - (ButWidth * 2)
    ElseIf (ExStyle And WS_EX_CONTEXTHELP) = WS_EX_CONTEXTHELP Then
            ' CANNOT HAVE MAX/MIN AND ? AT SAME TIME :)
            ' Same as Max/Min box but only one of them
            X = X - ButWidth
    End If
    
    ' Cool, that is all of them. Now take away 2 pixels for the gap
    X = X - 4
    
    ' Then take away another Width for our button
    X = X - ButWidth
    
    GetLeftPos = X  ' simple as that
            
End Function

Function IsOnTop(hWND As Long) As Boolean
    Dim X As Long
    X = GetWindowLong(hWND, GWL_EXSTYLE)
    If (X And WS_EX_TOPMOST) = WS_EX_TOPMOST Then IsOnTop = True Else IsOnTop = False
End Function


Sub hWndontop(hWND As Long, OnTop As Boolean)
' Sets Z of window to foreground (topmost)
    On Error Resume Next
   Dim Flags As Long
   Const SWP_NOMOVE = &H2
   Const SWP_NOSIZE = &H1
   Flags = SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos hWND, IIf(OnTop = True, -1, -2), 0, 0, 0, 0, Flags
    
End Sub

