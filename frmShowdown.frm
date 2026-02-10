VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmShowdown 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Showdown"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmShowdown.frx":0000
   ScaleHeight     =   3615
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Force Reroll"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   510
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "force your Rival to re-roll"
      Top             =   3150
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Reroll"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "use The Guardian to re-roll"
      Top             =   3150
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   6200
      TabIndex        =   3
      Top             =   1530
      Width           =   5655
      Begin VB.ListBox lstGear 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FF00&
         Height          =   1410
         Index           =   1
         Left            =   0
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   0
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Accept"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "add any single use Gear that may help, then ACCEPT the outcome"
      Top             =   3150
      Width           =   1035
   End
   Begin VB.Timer Timing 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   10470
      Top             =   3390
   End
   Begin VB.ListBox lstGear 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FF00&
      Height          =   1410
      Index           =   0
      Left            =   140
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1530
      Width           =   5655
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgCrew 
      Height          =   650
      Left            =   5640
      Top             =   0
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1138
      Attr            =   516
      FixedCx         =   47
      FixedCy         =   45
      Effects         =   "frmShowdown.frx":91F8
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discardable Gear cannot be withdrawn once selected "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   2940
      Width           =   5355
   End
   Begin VB.Label lblTotal2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "= 20"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   675
      Left            =   10560
      TabIndex        =   10
      Top             =   660
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblTotal1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "= 20"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   675
      Left            =   4490
      TabIndex        =   9
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label lblRight 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   675
      Left            =   7200
      TabIndex        =   8
      Top             =   660
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblLeft 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   675
      Left            =   1110
      TabIndex        =   7
      Top             =   660
      Width           =   825
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   345
      Left            =   6210
      TabIndex        =   6
      Top             =   90
      Width           =   5655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   345
      Left            =   180
      TabIndex        =   5
      Top             =   60
      Width           =   5655
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl picDice 
      Height          =   915
      Index           =   3
      Left            =   8250
      Top             =   540
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      Effects         =   "frmShowdown.frx":9210
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl picDice 
      Height          =   915
      Index           =   4
      Left            =   9390
      Top             =   540
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      Effects         =   "frmShowdown.frx":9228
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl picDice 
      Height          =   675
      Index           =   5
      Left            =   6450
      Top             =   670
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      Effects         =   "frmShowdown.frx":9240
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl picDice 
      Height          =   915
      Index           =   0
      Left            =   2130
      Top             =   540
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      Effects         =   "frmShowdown.frx":9258
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl picDice 
      Height          =   915
      Index           =   1
      Left            =   3240
      Top             =   540
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      Effects         =   "frmShowdown.frx":9270
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl picDice 
      Height          =   675
      Index           =   2
      Left            =   370
      Top             =   670
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      Effects         =   "frmShowdown.frx":9288
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   6210
      TabIndex        =   1
      Top             =   3090
      Width           =   5625
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2895
      Left            =   5925
      Top             =   630
      Width           =   135
   End
End
Attribute VB_Name = "frmShowdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OpponentID As Integer, isHost As Boolean
Public Skilltype, skill, Dice
Public ASkillType, ASkill, ADice, winShowdown As Boolean
Private rerollused As Boolean, forcererollused As Boolean

Private Sub cmd_Click(Index As Integer)
   Select Case Index
   Case 0
      If isHost Then
         DB.Execute "UPDATE GameSeq SET HostAccept = 1"
         'Logic!HostAccept = 1
      Else
         DB.Execute "UPDATE GameSeq SET ClientAccept = 1"
         'Logic!ClientAccept = 1
      End If
      'Logic.Update
      Logic.Requery
      cmd(0).Enabled = False
      cmd(1).Enabled = False
   Case 1 're-roll
      Dice = RollDice(6, True)
      rerollused = True
      DB.Execute "Update ShowdownScores set Dice = " & Dice & " WHERE PlayerID = " & player.ID
      cmd(1).Enabled = False
      PutMsg player.PlayName & " uses The Guardian's skill to re-roll and gets a " & Dice, player.ID, Logic!Gamecntr
      refreshPage
   Case 2 'force re-roll
      DB.Execute "Update ShowdownScores set forcereroll = 1 WHERE PlayerID = " & player.ID
      If isHost Then 'reset opponent in case it was cleared after acceptance
         DB.Execute "UPDATE GameSeq SET Trader = " & CStr(OpponentID)
         'Logic!trader = OpponentID
         'Logic.Update
         Logic.Requery
      End If
      forcererollused = True
      cmd(2).Enabled = False
   
   End Select
   playsnd 8
End Sub

Private Sub Form_Load()
   DB.Execute "Insert into ShowdownScores (PlayerID,SkillType,Skill,Dice) Values (" & player.ID & "," & Skilltype & "," & skill & "," & Dice & ")"
   refreshPage
   listGear player.ID, Skilltype
   listGear OpponentID, ASkillType
   Label3 = PlayCode(player.ID).PlayName & IIf(isHost, " - attacker", " - defender")
   Label4 = PlayCode(OpponentID).PlayName & IIf(isHost, " - defender", " - attacker")
   If hasCrew(player.ID, 90) Then cmd(1).Visible = True
   If hasCrew(player.ID, 91) Then cmd(2).Visible = True
   Timing.Enabled = True
   playsnd 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub lstGear_ItemCheck(Index As Integer, Item As Integer)
   If Not cmd(0).Enabled Then Exit Sub
   If Index = 0 Then
      If lstGear(Index).ListCount = 0 Then Exit Sub
      If lstGear(Index).selected(Item) Then
         DB.Execute "Insert into ShowdownGear (PlayerID, CardID) values (" & player.ID & ", " & lstGear(Index).ItemData(Item) & ")"
         skill = skill + getGearAttrib(lstGear(Index).ItemData(Item), cstrSkill(Skilltype))
      Else
         lstGear(Index).selected(Item) = True
         MessBox "Single use Gear once committed to the Showdown cannot be withdrawn", "Item committed", "Ooops", "", 0, getGearAttrib(lstGear(Index).ItemData(Item), "GearID")
         'DB.Execute "Delete from ShowdownGear where CardID = " & lstGear(Index).ItemData(Item)
         'skill = skill - getGearAttrib(lstGear(Index).ItemData(Item), cstrSkill(Skilltype))
      End If
      DB.Execute "UPDATE ShowdownScores set Skill = " & skill & " WHERE PlayerID = " & player.ID
      refreshPage
   End If
End Sub

Private Sub Timing_Timer()
Dim msg As String
   refreshPage
   listGear OpponentID, ASkillType
   If checkDone Then 'finalise, disgard gear used and return
      Timing.Enabled = False
      discardGearUsed
      If isHost Then 'attacker
         If skill + Dice > ASkill + ADice Then
            msg = " wins"
            winShowdown = True
         Else
            msg = " has lost"
         End If
      Else 'defender
         If skill + Dice >= ASkill + ADice Then
            msg = " wins"
            winShowdown = True
         Else
            msg = " has lost"
         End If
      End If
      PutMsg player.PlayName & msg & " the Showdown", player.ID, Logic!Gamecntr, True, getLeader()
      
      Me.hide
   End If
End Sub

Private Function checkDone() As Boolean
   Logic.Requery
   If Logic!HostAccept = 1 And Logic!ClientAccept = 1 Then checkDone = True
End Function

Private Sub refreshPage()
Dim SQL As String
Dim rst As New ADODB.Recordset

   SQL = "SELECT * FROM ShowdownScores WHERE PlayerID = " & OpponentID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If rst.EOF Then
      Label2 = "awaiting opponent to respond"
   Else
      'Label2 = "Skill: " & cstrSkill(rst!Skilltype) & " - " & rst!skill & "   Dice: " & rst!Dice & "   Total: " & rst!skill + rst!Dice
      lblRight.Visible = True
      lblTotal2.Visible = True
      lblRight = CStr(rst!skill)
      lblTotal2 = "= " & CStr(rst!skill + rst!Dice)
      ASkillType = rst!Skilltype
      ASkill = rst!skill
      ADice = rst!Dice
      doPics OpponentID, ASkillType, ASkill, ADice
      Logic.Requery
      If (isHost And Logic!HostAccept = 0) Or (Not isHost And Logic!ClientAccept = 0) Then
         cmd(0).Enabled = True
         cmd(1).Enabled = cmd(1).Visible And Not rerollused
         cmd(2).Enabled = cmd(2).Visible And Not forcererollused
      End If
      If (isHost And Logic!ClientAccept = 1) Or (Not isHost And Logic!HostAccept = 1) Then
         If Label2 <> "your opponent has locked in their score" Then playsnd 13
         Label2 = "your opponent has locked in their score"
      Else
         If Label2 <> "your opponent has initiated their score" Then playsnd 13
         Label2 = "your opponent has initiated their score"
      End If
      If rst!forcereroll = 1 Then
         Dice = RollDice(6, True)
         PutMsg PlayCode(OpponentID).PlayName & " uses Chari to force you into a re-roll, and you got a " & Dice, player.ID, Logic!Gamecntr, True, 91, 0, 0, 0, 0, Dice
         cmd(0).Enabled = True
         cmd(1).Enabled = cmd(1).Visible And Not rerollused
         cmd(2).Enabled = cmd(2).Visible And Not forcererollused
         
         DB.Execute "Update ShowdownScores set Dice = " & Dice & " WHERE PlayerID = " & player.ID
         DB.Execute "Update ShowdownScores set forcereroll = 0 WHERE PlayerID = " & OpponentID
         If isHost Then
            DB.Execute "UPDATE GameSeq SET HostAccept = 0"
            'Logic!HostAccept = 0
         Else
            DB.Execute "UPDATE GameSeq SET ClientAccept = 0"
            'Logic!ClientAccept = 0
         End If
         'Logic.Update
         Logic.Requery
      End If
   End If
   rst.Close
   
   lblLeft = CStr(skill)
   lblTotal1 = "= " & CStr(skill + Dice)
   doPics player.ID, Skilltype, skill, Dice
   

End Sub

Private Sub listGear(ByVal playerID As Integer, ByVal Skilltype As Integer)
Dim SQL As String, x
Dim rst As New ADODB.Recordset

   If Skilltype = 0 Then Exit Sub

   x = IIf(playerID = player.ID, 0, 1)
   lstGear(x).Clear

   SQL = "SELECT SupplyDeck.CardID, ShowdownGear.PlayerID, Gear.* "
   SQL = SQL & "FROM ShowdownGear RIGHT JOIN (PlayerSupplies INNER JOIN (Gear INNER JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID) ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ShowdownGear.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND Gear.Discard=1 AND PlayerSupplies.CrewID>0 AND PlayerSupplies.OffJob=0 AND Gear." & cstrSkill(Skilltype) & " >0"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      lstGear(x).AddItem Nz(rst!GearName) & " " & Nz(rst!GearDescr)
      lstGear(x).ItemData(lstGear(x).NewIndex) = rst!CardID
      If Nz(rst!playerID) = playerID Then
         lstGear(x).selected(lstGear(x).NewIndex) = True
      End If
      rst.MoveNext
   Wend
   rst.Close

End Sub

Private Sub doPics(ByVal playerID As Integer, ByVal Skilltype, ByVal skill, ByVal Dice)
Dim x, Y, z
   If playerID = player.ID Then
      x = 0: Y = 1: z = 2
   Else
      x = 3: Y = 4: z = 5
   End If

   If Dice > 0 And Dice < 7 Then
      picDice(x).Visible = True
      picDice(x).Picture = LoadPictureGDIplus(App.Path & "\pictures\D" & Dice & ".bmp") ' LoadPicture(App.Path & "\pictures\D" & dice & ".bmp")

      picDice(x).TransparentColor = 0
      picDice(x).TransparentColorMode = lvicUseTransparentColor
      picDice(Y).Visible = False

   ElseIf Dice > 6 Then
      picDice(x).Visible = True
      picDice(x).Picture = LoadPictureGDIplus(App.Path & "\pictures\D6.bmp") 'LoadPicture(App.Path & "\pictures\D6.jpg")
      picDice(x).TransparentColor = 0
      picDice(x).TransparentColorMode = lvicUseTransparentColor
      picDice(Y).Visible = True
      picDice(Y).Picture = LoadPictureGDIplus(App.Path & "\pictures\D" & (Dice - 6) & ".bmp")  'LoadPicture(App.Path & "\pictures\D" & (dice - 6) & ".bmp")
      picDice(Y).TransparentColor = 0
      picDice(Y).TransparentColorMode = lvicUseTransparentColor
   End If
   If Skilltype > 0 Then
      picDice(z).Visible = True
      picDice(z).Picture = LoadPictureGDIplus(App.Path & "\pictures\" & picSkill(Skilltype) & ".bmp") 'LoadPicture(App.Path & "\pictures\D6.jpg")
      picDice(z).TransparentColor = 0
   End If
End Sub

Private Sub discardGearUsed()
Dim x
   With lstGear(0)
      For x = 0 To .ListCount - 1
         If .selected(x) Then 'discard it
            doDiscardGear player.ID, .ItemData(x)
            PutMsg player.PlayName & " discards " & getGearAttrib(.ItemData(x), "GearName"), player.ID, Logic!Gamecntr
         End If
      Next x
   End With
End Sub
