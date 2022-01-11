VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções (Viram mais a sim que precisar)"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5460
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   360
      Left            =   4065
      TabIndex        =   2
      Top             =   1005
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   5265
      Begin VB.CheckBox chkUseBalloon 
         Caption         =   "Usar legendas de balão quando enviar aplicativo para a bandeja"
         Height          =   285
         Left            =   210
         TabIndex        =   1
         Top             =   375
         Width           =   4950
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUseBalloon_Click()
    SaveSetting "MIN2TRAY", "Options", "UseBalloon", CStr(Me.chkUseBalloon.Value)
    End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If modTray.IsV5Compat(True) = False Then
        Me.chkUseBalloon.Enabled = False
        Me.chkUseBalloon.Value = 0
    Else
        Me.chkUseBalloon.Enabled = True
        Me.chkUseBalloon.Value = CStr(GetSetting("MIN2TRAY", "Options", "UseBalloon", "1"))
    End If
 
End Sub
