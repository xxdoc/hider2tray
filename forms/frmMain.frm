VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Window Minimizer v1.1 - 2001"
   ClientHeight    =   1845
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgSys 
      Left            =   1035
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   750
      Left            =   2010
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMain.frx":058A
      Top             =   240
      Visible         =   0   'False
      Width           =   2070
   End
   Begin MSComctlLib.ListView lstJanelas 
      Height          =   1635
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   2884
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "imgSys"
      SmallIcons      =   "imgSys"
      ColHdrIcons     =   "imgView"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Menu menuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Opções"
         Index           =   1
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Restaurar Janelas"
         Index           =   2
         Begin VB.Menu mnuArqRestore 
            Caption         =   "Janelas Miinimizadas:"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuArqRestore 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuArqResJanelas 
            Caption         =   "(nenhuma)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Sair"
         Index           =   4
      End
   End
   Begin VB.Menu menuAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuAjduda 
         Caption         =   "&Sobre"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lstJanelas.BackColor = RGB(191, 175, 127)
    '  Me.BackColor = RGB(191, 175, 52)
    
    frmOnTop.Show
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    'Constroe o listview - com os sistemas
    With lstJanelas
      .Top = 45
      .Left = 45
      .Height = (Me.ScaleHeight - 90) '1400)
      .Width = (Me.ScaleWidth - 90)
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
 ReleaseAll
 Unload frmOnTop
 Unload frmTray
 End
End Sub

Sub ReleaseAll()
    Dim OC As String
    Dim i As Integer
    
    OC = Me.Caption
    Me.Caption = "Restaurar todas as janelas.."
    
    For i = 1 To 999 'frmTray.Count + 1
        If Formz(i).inUse = True Then
            Formz(i).vbForm.RemME
            DoEvents
        End If
    Next
    Me.Caption = OC
End Sub

Public Sub ReBuildWindowList()
   Dim i As Integer
   Dim CurPos As Integer
   
   ' max of 999.. never happen but ah well
   For i = 1 To mnuArqResJanelas.Count - 1
      Unload mnuArqResJanelas(i)
   Next
   
   CurPos = mnuArqResJanelas.Count
   
   lstJanelas.ListItems.Clear
   For i = 1 To 999
'      On Error Resume Next
      If Formz(i).inUse = True Then
          Load mnuArqResJanelas(CurPos)
          With mnuArqResJanelas(CurPos)
             .Caption = Formz(i).vbForm.Caption
             .Enabled = True
             .Tag = Trim$(Str$(i))
             
    '         'Incluindo Icone na lista
    '          Call SavePicture(Formz(i).vbForm.xIconH, i & ".ico")
    
              imgSys.ListImages.Add , , Formz(i).vbForm.Icon  ' LoadPicture(i & ".ico")
    '          'Apagar o arquivo
              'Kill i & ".ico"
              lstJanelas.ListItems.Add(, "KEY:" & CurPos, Formz(i).vbForm.Caption, imgSys.ListImages.Count, imgSys.ListImages.Count).Tag = .Tag
          
          End With
          CurPos = CurPos + 1
      End If
      On Error GoTo 0
   Next

'   On Error Resume Next
   If CurPos = 0 Then
      Load mnuArqResJanelas(CurPos)
      With mnuArqResJanelas(CurPos)
         .Caption = "(nenhuma)"
         .Enabled = False
         .Tag = "0"
         
         lstJanelas.ListItems.Clear
      End With
   End If
   On Error GoTo 0
End Sub

Private Sub lstJanelas_ItemClick(ByVal Item As MSComctlLib.ListItem)
  mnuArqResJanelas_Click CInt(Mid(Item.Key, 5))
End Sub

Private Sub mnuAjduda_Click(Index As Integer)
   On Error Resume Next
   frmSplash.Show
End Sub

Private Sub mnuArqResJanelas_Click(Index As Integer)
   Dim Msg As String
   Dim WS As String
   Dim X As Long
   
   With Formz(Val(Trim$(mnuArqResJanelas(Index).Tag)))
      Msg = Msg & .vbForm.Caption & vbCrLf
      Msg = Msg & "----------------------------------------------" & vbCrLf & vbCrLf
      Msg = Msg & "Guardado às: " & .SentAwayTime & vbCrLf
      'Msg = Msg & "ThreadID=" & .ThreadID & " hWND=" & .hwnd & vbCrLf & vbCrLf
      X = Val(Trim$(.vbForm.Tag))
      If (X And WS_MAXIMIZE) = WS_MAXIMIZE Then
         WS = "Maximizado"
      ElseIf (X And WS_MINIMIZE) = WS_MINIMIZE Then
         WS = "Minimizado"
      Else
         WS = "Normal (não maximiza)"
      End If
      
      Msg = Msg & "Estado da janela: " & WS & vbCrLf & vbCrLf
      Msg = Msg & "Restaurar esta janela?"
      
      If MsgBox(Msg, vbYesNo Or vbQuestion, "Restaurar Aplicativo?") = vbYes Then
         .vbForm.RemME
      End If
   End With
End Sub

Private Sub mnuArquivo_Click(Index As Integer)
   Select Case Index
      Case 1   'Opcoes
         frmOptions.Show
      Case 2   'Restore
      Case 4   'Sair
         Unload Me
   End Select
End Sub
