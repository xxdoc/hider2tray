VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00171717&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3480
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00171717&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3300
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   5130
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   45
         Top             =   0
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2820
         Width           =   1305
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00212121&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2445
         Left            =   45
         ScaleHeight     =   2415
         ScaleWidth      =   5010
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   255
         Width           =   5040
         Begin VB.Label Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Scrollling Text :~)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   930
            Left            =   45
            TabIndex        =   3
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   4905
            WordWrap        =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  Call FillText
  Me.Text1.AutoSize = True
  Me.Text1.Top = Me.Picture1.Height
End Sub

Private Sub ScrollText(sText As String)
    Me.Text1.Caption = Me.Text1.Caption & sText & vbCrLf
End Sub

Private Sub FillText()
  Me.Text1.Caption = ""
  Call ScrollText("Hider2Tray Tool")
  Call ScrollText("")
  Call ScrollText("")
  Call ScrollText("")
  Call ScrollText("Ferramenta que cria opção de ESCONDER janelas na bandeja do sistema. " & vbCrLf & "Versão " & App.Major & "." & App.Minor & "." & App.Revision)
  Call ScrollText("")
  Call ScrollText("")
  Call ScrollText("Este ferramenta foi projetada para lhe permitir esconder qualquer janela na bandeja de sistema e restaurar com facilidade.")
  Call ScrollText("")
   
  Call ScrollText("By: Heliomar P.Marques (c) 2001 " & vbCrLf & "E-Mail: heliomarpm@hotmail.com")
  Call ScrollText("ICQ-UIN: 42989242 ")
  Call ScrollText("")
  Call ScrollText("")
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then Me.Text1.Top = Me.Text1.Top - 50
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then Me.Text1.Top = Me.Text1.Top - 50
End Sub

Public Sub Timer1_Timer()
    With Me.Text1
       .Top = .Top - (Me.Picture1.Height / 150)
       If .Top + .Height <= 0 Then .Top = Me.Picture1.Height
    End With
End Sub
