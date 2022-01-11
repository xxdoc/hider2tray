VERSION 5.00
Begin VB.Form frmOnTop 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2520
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3375
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOnTop.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   915
      Top             =   1785
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   1215
   End
End
Attribute VB_Name = "frmOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastHWND As Long
Private NewRECT As RECT
Private oldRECT As RECT
Private newHWND As Long
Private MMTF As Boolean

Private Sub Form_Load()
    Me.Move 0, 0, 0, 0
    Me.Refresh
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'frmMain.Caption = "MouseDown"
  modGWL.BuildButton Me, True
  MMTF = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 0 Then
    'frmMain.Caption = "MouseMove"
    With Me
      If (X < 0) Or (X > (.Width / 15)) Or (Y < 0) Or (Y > (.Height / 15)) Then
        ' outside
        If MMTF = True Then
          modGWL.BuildButton Me, False
          MMTF = False
        End If
      Else
        'inside
        If MMTF = False Then
          modGWL.BuildButton Me, True
          MMTF = True
        End If
      End If
    End With
    hWndontop Me.hWND, True
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With Me

        If Button = vbRightButton Then
            If (X < 0) Or (X > (.Width / 15)) Or (Y < 0) Or (Y > (.Height / 15)) Then
                ' outside
            Else
                'inside
                PopupMenu frmMain.mnuArqRestore
            End If
        
            modGWL.GetPos LastHWND, Me
            Exit Sub
        End If

        modGWL.GetPos LastHWND, Me
        modGWL.hWndontop Me.hWND, True
    
        If (X < 0) Or (X > (.Width / 15)) Or (Y < 0) Or (Y > (.Height / 15)) Then
            ' Outside
            modMisc.SetFocus LastHWND
        Else
            ' Inside
            CreateNewWindow LastHWND
            Me.Visible = False
            'SetParent Me.hWnd, LastHWND
            'Misc.SetFocus LastHWND
 
            LastHWND = 0
            newHWND = 0
        End If
    End With
End Sub

Sub CreateNewWindow(bHWND As Long)
    Dim PID As Long
    Dim APos As Integer
    
    APos = modMisc.GetNextFreeFormz
    Formz(APos).inUse = True
    Set Formz(APos).vbForm = New frmTray
    Formz(APos).hWND = bHWND
    Formz(APos).SentAwayTime = Format$(Time$, "h:mm.ss AMPM") & " na " & modDate.GetDateStr
    Formz(APos).ThreadID = modMisc.GetWindowThreadProcessId(bHWND, PID)
            
    With Formz(APos).vbForm
        modRelatives.GetRelatives bHWND, Formz(APos).vbForm
        .xHWND = bHWND
        .APos = APos
        .xIconH = modICON.GetIconHandle(bHWND)
        If .xIconH = 0 Then .xIconH = frmMain.Icon.Handle
        
        .Caption = GetWindowCaption(bHWND)
        .SendToTray
    End With

    ' Check for Problems
    Dim W As Long
    W = GetWindowLong(LastHWND, GWL_STYLE)
 
    'If (W And WS_VISIBLE) = WS_VISIBLE Then
    '       MsgBox "This application is not supported", vbInformation, "Cannot hide window"
    '       Tray.KillTray_And_RestoreHwnd Formz(APos).vbForm, bHWND, False, True
    '       Unload Formz(APos).vbForm
    '       Formz(APos).inUse = False
    '       Misc.SetFocus bHWND
    'End If

    frmMain.ReBuildWindowList  ' add to Menu
End Sub

Private Sub Timer1_Timer()
    ' OK, Bug Fix time
    ' The problem:
    ' -----------------------------------------
    ' when you click on the button, the ForeGroundWindow is set to
    ' the hWnd of this form. However if you close the app, sometimes your
    ' app will gain focus without wanting it. We must detect this then hide
    ' the form until another window is generated
       
    ' Find the hWnd of the Current ForeGround
     newHWND = GetForegroundWindow
    
    
    If newHWND = Me.hWND Then
        ' Yes, the focus is on this window.
        Dim GW As Long
        GW = GetWindowLong(LastHWND, GWL_STYLE)
        If (GW And WS_VISIBLE) = WS_VISIBLE And (GW And WS_MINIMIZE) <> WS_MINIMIZE Then
            ' The Window is Visible. And NOT minimized
            ' The user is clicking on the button
            'Misc.SetFocus LastHWND
            Exit Sub
        Else
            ' The Window is NOT visible or is minimized. Hide this window
            ' Most likley reason, user closed app.
            LastHWND = 0
            Me.Visible = False
            Exit Sub
        End If
    End If
        
    modMisc.GetWindowRect newHWND, NewRECT
    
    If (oldRECT.Right <> NewRECT.Right) Or (oldRECT.Top <> NewRECT.Top) Or (newHWND <> LastHWND) Then
        
        LastHWND = newHWND
        GetWindowRect newHWND, oldRECT
        
        If GetParent(newHWND) <> 0 Then
            ' Window has a parent.
            ' Do not allow this window to be sent to system tray
            Me.Visible = False
            Exit Sub
        End If
        
        If modMisc.IsValid(newHWND) = False Then
            Me.Visible = False
            Exit Sub
        End If
        
        modGWL.BuildButton Me
        
        Dim W As Long
        W = GetWindowLong(newHWND, GWL_STYLE)
        
        If (W And WS_VISIBLE) <> WS_VISIBLE Then
            ' Not Visible
            Me.Visible = False
        Else
            Me.Visible = True
        End If
        
        modGWL.GetPos newHWND, Me
        
        If Me.Visible = True Then modGWL.hWndontop Me.hWND, True
        
        modMisc.SetFocus newHWND
    End If
End Sub

Private Sub Timer2_Timer()
    Dim X As Long
    X = GetWindowLong(Me.hWND, GWL_EXSTYLE)
    If (X And WS_EX_TOPMOST) <> WS_EX_TOPMOST Then
       If Me.Visible = True Then modGWL.hWndontop Me.hWND, True
       modMisc.SetFocus LastHWND
    End If
End Sub



