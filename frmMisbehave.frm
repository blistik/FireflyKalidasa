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
      Caption         =   "$200 to Redraw"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "FORQUE"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Dalin: redraw another Misbhave for $200"
      Top             =   7650
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   ": Proceed"
      BeginProperty Font 
         Name            =   "FORQUE"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ace in the hole"
      Top             =   6990
      Visible         =   0   'False
      Width           =   4425
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
      Top             =   6990
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
      Left            =   4890
      TabIndex        =   6
      Top             =   7890
      Width           =   1065
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
Public MBCardID, MBOption, hasAce As Boolean, dalin As Boolean

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
         MBOption = 3
         If cmd(2).Tag = "discardsword" Then
            doDiscardGear player.ID, hasGearCard(player.ID, 33)
            PutMsg player.PlayName & " uses the Operative's Sword as an Ace in the Hole", player.ID, Logic!Gamecntr
         ElseIf cmd(2).Tag <> "" Then
            discardGearKeyword player.ID, cmd(2).Tag
            PutMsg player.PlayName & " misbhavin' with " & lblName & " had an Ace in the Hole with " & cmd(2).Caption, player.ID, Logic!Gamecntr
         Else
            PutMsg player.PlayName & " misbhavin' with " & lblName & " had an Ace in the Hole with " & cmd(2).Caption, player.ID, Logic!Gamecntr
         End If
         actionSeq = ASnavEnd
      Case 3 'dalin
         MBCardID = 0
         MBOption = 0
         PutMsg player.PlayName & "'s Dalin deals $200 to step over " & lblName & " to try another.", player.ID, Logic!Gamecntr
         frmAction.setPay getMoney(player.ID, -200)
         dalin = False
   End Select
   
   For x = 0 To 3
      cmd(x).Enabled = False
   Next x
   Me.hide
End Sub

Private Sub Form_Load()
   hasAce = False
    Set Me.Picture = LoadPicture(App.Path & "\pictures\MisbehaveTemplate.bmp")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If

End Sub

Public Sub setAce(ByVal title As String, Optional ByVal discardGear As String = "", Optional ByVal discardsword As Boolean = False)
   cmd(2).Caption = title & ": Proceed"
   cmd(2).Visible = True
   If discardsword Then
      cmd(2).Tag = "discardsword"
   ElseIf discardGear <> "" Then
      cmd(2).Tag = discardGear
   End If
   hasAce = True
End Sub

Public Sub setDalin(ByVal d As Boolean)
   dalin = d
   cmd(3).Visible = d
   If getMoney(player.ID) >= 200 Then cmd(3).Enabled = d

End Sub

Public Sub setAcelbl(ByVal title As String)

   lblAce.Caption = title & ": Proceed"
   lblAce.Visible = True

End Sub

