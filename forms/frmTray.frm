VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$TRAY FORM$"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2535
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1155
      Top             =   315
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnuRestoreWindow 
         Caption         =   "&Restaurar"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestoreALL 
         Caption         =   "Restaurar &Todas"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public xIconH As Long
Public xHWND As Long
Private bHWnds(1 To 99) As Long
Private bStyle(1 To 99) As Long
Public APos As Integer

Public Sub SetbHwndsVals(bValNum, bValue As Long)
    bHWnds(bValNum) = bValue
End Sub

Public Sub SeTBSizeStylesVals(bValNum, bValue As Long)
    bStyle(bValNum) = bValue
End Sub
    
Private Sub Form_Load()
    Me.Visible = False
End Sub

Sub RemME()
    Dim i As Integer
    
    For i = 1 To 99
        If bHWnds(i) > 0 Then
            modTray.ShowHWND bHWnds(i), bStyle(i)
'            Debug.Print "Removed hwnd: " & Hex(bHWnds(I)) & " from tray"
        End If
         
        bHWnds(i) = 0
        bStyle(i) = 0
        DoEvents
    Next
       
    modTray.KillTray_And_RestoreHwnd Me, xHWND, False, True
    modMisc.SetFocus xHWND
  
    With Formz(APos)
        .inUse = False
        .hWND = 0
        .SentAwayTime = ""
        .ThreadID = 0
        Set .vbForm = Nothing
    End With
    
    frmMain.ReBuildWindowList
    frmOnTop.LastHWND = 0
 
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.Visible = False Then
        Static lngMsg As Long
        On Error Resume Next
        lngMsg = X / Screen.TwipsPerPixelX
        
        Select Case lngMsg
        Case WM_LBUTTONUP
'            Me.WindowState = Normal
'            Me.Visible = True
'            SysTray.FormOnTop Me, True
            RemME
        
        Case WM_RBUTTONUP
            PopupMenu mnuOptions
            ' Popup the &File menu (invisible normally)
        End Select
    End If
    
End Sub

Public Function SendToTray()
    Dim i As Integer
    
    'Me.Caption = Me.Caption ' & "(" & APos & ")"
'    frmMain.Caption = "SendToTray"
    modTray.SendToTray Me, xHWND, Me.Caption, True, True, xIconH
    For i = 1 To 99
        If bHWnds(i) > 0 Then
            bStyle(i) = modTray.HideHWND(bHWnds(i))
'            Debug.Print "Sent hwnd: " & Hex(bHWnds(I)) & " to tray"
        End If
        DoEvents
    Next
    
End Function

Private Sub mnuAbout_Click()
    On Error Resume Next
    frmSplash.Show
End Sub

Private Sub mnuRestoreALL_Click()
    frmMain.ReleaseAll
End Sub

Private Sub mnuRestoreWindow_Click()
    RemME
End Sub

Private Sub Timer1_Timer()
  ' Check to see if window is still minimized
  Dim WL As Long
  WL = GetWindowLong(Me.xHWND, GWL_STYLE)
  
  If WL = 0 Then
    ' Window Handle gone.
    ' Kill Tray and unload form
    'set values to zero so no window is effected by mistake
    Me.Tag = 0      ' loose WindowStyle
    Me.xHWND = 0    ' set Handle to zero
    RemME
  End If
  
  If (WL And WS_VISIBLE) = WS_VISIBLE Then
    ' Window is visible... damnit
    RemME
  End If
    
End Sub
