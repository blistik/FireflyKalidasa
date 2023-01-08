VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmMisbehave 
   BackColor       =   &H00CBE1ED&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "doin' some Misbehavin'"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   5985
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "The Operative's Sword"
      Top             =   7620
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Curse your sudden but inevitable Betrayal"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1530
      Width           =   3135
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "These are not the Crooks you're looking for"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4140
      Width           =   3135
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl SkillImg 
      Height          =   390
      Index           =   1
      Left            =   4500
      Top             =   4845
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   688
      Trans           =   100663295
      Effects         =   "frmMisbehave.frx":0000
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl SkillImg 
      Height          =   390
      Index           =   0
      Left            =   4500
      Top             =   2235
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   688
      Trans           =   100663295
      Effects         =   "frmMisbehave.frx":0018
   End
   Begin VB.Label lblKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EXPLOSIVES"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Index           =   1
      Left            =   1600
      TabIndex        =   9
      Top             =   6090
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label lblKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EXPLOSIVES"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Index           =   0
      Left            =   1600
      TabIndex        =   8
      Top             =   3510
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label lblAce 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EXPLOSIVES: Proceed"
      BeginProperty Font 
         Name            =   "FORQUE"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   555
      Left            =   1320
      TabIndex        =   7
      Top             =   6960
      Visible         =   0   'False
      Width           =   4395
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaInv 
      Height          =   645
      Left            =   270
      Top             =   7100
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1138
      Effects         =   "frmMisbehave.frx":0030
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Let's go to the Crappy Town where I,m a Hero"
      BeginProperty Font 
         Name            =   "FORQUE"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1095
      Left            =   600
      TabIndex        =   4
      Top             =   300
      Width           =   4425
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUnseen 
      BackStyle       =   0  'Transparent
      Caption         =   "unseen"
      ForeColor       =   &H00004040&
      Height          =   225
      Left            =   2600
      TabIndex        =   6
      Top             =   7830
      Width           =   1215
   End
   Begin VB.Label lblDetail 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Fight: 1-5 Kill All Crew, Warrant Issued. 6-9 Kill 2 Crew, Warrant Issued.  10+ Attempt Botched"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1425
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   2480
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDetail 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "If Mercs fight total is higher than the rest of the Crew, incl Gear, discard all Mercs. Proceed"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1395
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   5090
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Alpha 
      Height          =   645
      Left            =   5100
      Top             =   360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1138
      Effects         =   "frmMisbehave.frx":0048
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
