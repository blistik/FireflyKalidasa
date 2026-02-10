VERSION 5.00
Begin VB.Form frmSkillSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Showdown"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSkillSel.frx":0000
   ScaleHeight     =   1500
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   2
      Left            =   3060
      MaskColor       =   &H00000000&
      Picture         =   "frmSkillSel.frx":1513E6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   510
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   1
      Left            =   1680
      MaskColor       =   &H00000000&
      Picture         =   "frmSkillSel.frx":151EDA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   510
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   0
      Left            =   300
      MaskColor       =   &H00000000&
      Picture         =   "frmSkillSel.frx":1528C0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   510
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Label lblSupply 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Choose a Skill to use in the Showdown"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   3
      Top             =   60
      Width           =   3825
   End
End
Attribute VB_Name = "frmSkillSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public skill
Private fmode

Private bOnTopState As Boolean

Public Property Let AlwaysOnTop(bState As Boolean)
    Dim lFlag As Long
    On Error Resume Next
    If bState = True Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    bOnTopState = bState
    SetWindowPos Me.hwnd, lFlag, 0&, 0&, 0&, 0&, (SWP_NOSIZE Or SWP_NOMOVE)
End Property
Public Property Get AlwaysOnTop() As Boolean
    AlwaysOnTop = bOnTopState
End Property

Private Sub cmd_Click(Index As Integer)
   skill = Index + 1
   If fmode = 2 Then actionSeq = ASBountySkillSel
   playsnd 8
   Me.hide
End Sub

Private Sub Form_Load()
   If fmode > 0 Then playsnd 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub
'mode 0 = showdowns (default), mode 1 = boarding test, 2 - defender
Public Sub setMode(ByVal mode)
   fmode = mode
   If mode = 1 Then
      cmd(0).Enabled = False
      lblSupply.Caption = "Choose a Skill to use for the Boarding Test"
   End If
End Sub
