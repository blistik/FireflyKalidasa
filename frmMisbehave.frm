VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmMisbehave 
   BackColor       =   &H00004080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "doin' some Misbehavin'"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMisbehave.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "use Ace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "The Operative's Sword"
      Top             =   570
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "option 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   930
      Width           =   5025
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "option 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2700
      Width           =   5025
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   885
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5025
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUnseen 
      BackStyle       =   0  'Transparent
      Caption         =   "unseen"
      ForeColor       =   &H00004040&
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   4260
      Width           =   1815
   End
   Begin VB.Label lblDetail 
      BackColor       =   &H00000040&
      Caption         =   "Misbehave 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1125
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   1350
      Width           =   5025
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDetail 
      BackColor       =   &H00000040&
      Caption         =   "Misbehave 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1125
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   3120
      Width           =   5025
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Alpha 
      Height          =   645
      Left            =   150
      Top             =   285
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1138
      Effects         =   "frmMisbehave.frx":BB00
   End
End
Attribute VB_Name = "frmMisbehave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MBCardID, MBOption

Private Sub cmd_Click(Index As Integer)
Dim x
   playsnd 8
   Select Case Index
      Case 0
         MBOption = 1
         actionSeq = ASnavEnd
      Case 1
         MBOption = 2
         actionSeq = ASnavEnd
      Case 2
         doDiscardGear player.ID, hasGearCard(player.ID, 33)
         MBOption = 3
         actionSeq = ASnavEnd
   End Select
   
   For x = 0 To 2
      cmd(x).Enabled = False
   Next x
   Me.Hide
End Sub


Private Sub Form_Load()
   If hasGear(player.ID, 33) Then
      cmd(2).Visible = True
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If

End Sub
