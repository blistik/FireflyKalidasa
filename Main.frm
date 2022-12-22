VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Main 
   BackColor       =   &H80000006&
   Caption         =   "Firefly - The PC Game"
   ClientHeight    =   12375
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20070
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "Main.frx":030A
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timing 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1830
      Top             =   1800
   End
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   12120
      Width           =   20070
      _ExtentX        =   35401
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   34880
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   1110
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4093C
            Key             =   "ship"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":40D90
            Key             =   "start"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":411E4
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":41638
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":41CB2
            Key             =   "chat"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":42104
            Key             =   "grap"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":42556
            Key             =   "crew"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":431A8
            Key             =   "job"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":435FA
            Key             =   "deal"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":43A4C
            Key             =   "join"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":43E9E
            Key             =   "cash"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":441B8
            Key             =   "help"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":444D2
            Key             =   "grapz"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":447EC
            Key             =   "hat"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":44B06
            Key             =   "crewz"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":44E20
            Key             =   "serenity"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4513A
            Key             =   "upgrd"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4558C
            Key             =   "log"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":459DE
            Key             =   "graph"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20070
      _ExtentX        =   35401
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "Images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "start"
            Object.ToolTipText     =   "Host"
            ImageKey        =   "start"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "join"
            Object.ToolTipText     =   "Join"
            ImageKey        =   "join"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "exit"
            Object.ToolTipText     =   "End Game"
            ImageKey        =   "exit"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "chat"
            Object.ToolTipText     =   "Chat"
            ImageKey        =   "chat"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "graph"
            Object.ToolTipText     =   "Game Info"
            ImageKey        =   "graph"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "log"
            Object.ToolTipText     =   "Game Log"
            ImageKey        =   "log"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "crew"
            Object.ToolTipText     =   "Crew Browser"
            ImageKey        =   "crewz"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "gear"
            Object.ToolTipText     =   "Gear Browser"
            ImageKey        =   "hat"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "upgrd"
            Object.ToolTipText     =   "Ship Upgrades Browser"
            ImageKey        =   "upgrd"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ship"
            Object.ToolTipText     =   "Ship Browser"
            ImageKey        =   "serenity"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "allships"
                  Text            =   "All"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "myship"
                  Text            =   "Me"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "job"
            Object.ToolTipText     =   "Job Browser"
            ImageKey        =   "job"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "alljobs"
                  Text            =   "All"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "myjobs"
                  Text            =   "Me"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deal"
            Object.ToolTipText     =   "Deal Browser"
            ImageKey        =   "deal"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "alldeals"
                  Text            =   "All"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "localdeals"
                  Text            =   "Local Deals"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "buy"
            Object.ToolTipText     =   "Buy Browser"
            ImageKey        =   "cash"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "allbuys"
                  Text            =   "All"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "localbuys"
                  Text            =   "Local Buys"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Game Rules"
            ImageKey        =   "help"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "firefly"
                  Text            =   "Firefly Rulebook"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bluesun"
                  Text            =   "Blue Sun Rulebook"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kalidasa"
                  Text            =   "Firefly Kalidasa Rulebook"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pcguide"
                  Text            =   "Firefly for PC Guide"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "jobs"
                  Text            =   "Job View/Edit"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "about"
                  Text            =   "About"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin XDOCKFLOATLibCtl.DockFrame DockFrame1 
      Left            =   570
      Top             =   1770
      _cx             =   688
      _cy             =   688
      DragAreaStyle   =   0
      PICTCNT         =   0
      MENUCNT         =   0
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Verse As Board
Attribute Verse.VB_VarHelpID = -1
Private usedStitchSkill As Boolean
Public frmJob As frmJobs, frmShip As frmShips, frmDeal As frmDeals, frmBuy As frmSupply, frmStat As frmStats

Private Sub MDIForm_Load()
Dim x
   
   PlayCode(1).Color = "Orange"
   PlayCode(2).Color = "Blue"
   PlayCode(3).Color = "Yellow"
   PlayCode(4).Color = "Green"
   pickStartSector = -1
   actionSeq = ASidle

   DockFrame1.LoadStates "Firefly"
   
   Set Verse = New Board
   
   initToolbar False


   If Not Logon Then End

   Logic.Open "GameSeq", DB, adOpenDynamic, adLockPessimistic ', adLockOptimistic
   x = GetSeq
   Timing.Enabled = True
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   DockFrame1.SaveStates "Firefly"
   If Me.Visible Then
      If MessBox("Are you sure you want to close the game?", "Closing Game", "Yes", "No") = 0 Then
         If DB.State = adStateOpen Then DB.Close
         End
      Else
         Cancel = True
      End If
   End If

End Sub

'THE MAIN ENGINE of the GAME
' Game States E - Idle/End, H - Host screen, 1-4 players go. S - setup Game, R - run Game, T-Trade
' W - Reaver to any Rim or Border sector, X-Move a Reaver 1 sector, Y=Move the Cruiser 1 sector, Z- move the Cruiser adjacent player, V-move Corvette Adjacent player
' actionSeq States = ASidle , ASselect --- >>> , ASend, -> ASidle, <repeat>
Private Sub Timing_Timer()
Dim status As Variant, errh, thisPlayer As Integer
Dim SectorID, ContactID As Integer, SupplyID As Integer, x
Dim maxConsider
On Error GoTo err_handler

   SectorID = getPlayerSector(player.ID)
   ContactID = Nz(varDLookup("ContactID", "Contact", "SectorID=" & SectorID), 0)
   SupplyID = Nz(varDLookup("SupplyID", "Supply", "SectorID=" & SectorID), 0)

   status = GetSeqX(thisPlayer)
   'aminmate the current player
   If status = "R" And player.ID > 0 Then animatePlayer thisPlayer

   If status <> "H" And status <> "E" And status <> "L" And pickStartSector > -1 Then
      RefreshBoard
   End If
   If status = "E" Then 'currently in End Game
      PutMsg "Waiting to Host or Join a Game"
   ElseIf status = "S" And thisPlayer = player.ID And pickStartSector = 0 Then  'your go to pick starting sector on MAP
      Verse.Caption = "the 'Verse - " & varDLookup("StoryTitle", "Story", "StoryID = " & Logic!StoryID)
      NumOfReavers = varDLookup("NoOfReavers", "Story", "StoryID = " & Logic!StoryID)
      'set game ships
      For x = 5 To 6 + NumOfReavers
         MoveShip x, varDLookup("StartSectorID", "Players", "PlayerID=" & CStr(x))
      Next x
      PutMsg player.PlayName & " selecting Start Sector", player.ID, Logic!Gamecntr
      
      If useHavens(Logic!StoryID) Then
         MessBox "Click on the Planet Sector to be your Haven", "Pick your Haven", "Will do", "", getLeader()
      Else
         MessBox "Click on the Sector you want to start in", "Place your Ship", "Will do", "", getLeader()
      End If
      
      pickStartSector = 1
      
   ElseIf status = "S" And thisPlayer = player.ID And pickStartSector = 2 Then  'setup
         'MsgBox "Stage: " & Logic!Seq & " -  PlayerID: " & Logic!player & " - Counter: " & Logic!GameCntr, vbInformation, "Sector Picked"
      PutMsg player.PlayName & "'s on the Map", player.ID, Logic!Gamecntr
      
      'deal start drive core, and Jobs
      dealDriveAndJobs player.ID
      
      'starting point selected, pass to next person, or kick the main Running Game cycle off
      setNextPlayerREV player.ID, "R"
      Logic.Requery
      If Logic!Seq = "R" Then
         PutMsg PlayCode(Logic!player).PlayName & "'s Turn", Logic!player, Logic!Gamecntr
      End If
   
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASidle Then   'MAIN Cycle - init your go
      playsnd 6
      If (getCutterSector(SectorID) > 0 Or getCruiserCorvette(SectorID) > 0) And CruiserCutter <> SectorID Then
         If checkWhisperX1(SectorID) Then
            actionSeq = ASNavEvade ' and get away
            Exit Sub
         End If
      End If

      actionSeq = ASselect 'in limbo awaiting user to select
      
      showActions
      
   ElseIf status = "T" And thisPlayer <> player.ID And actionSeq = ASidle And Logic!trader = player.ID Then
      doSlaveTrade Logic!player
      
   ElseIf status = "U" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Corvette to any planetary sector
      x = setPlayer(player.ID, "", 0, True)
      MessBox "Move the Operative's Corvette to any Planetary Sector", "Place the Corvette", "OK"
      'kick it off
      actionSeq = ASNavCorvPlanetary
      
   ElseIf status = "W" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Reaver Cycle from another Player's Nav move
      MessBox "Move a Reaver to any Rim or Border sector", "Place a Reaver", "OK"
      'kick it off
      actionSeq = ASNavReavBorder
      
   ElseIf status = "W" And thisPlayer = player.ID And actionSeq = ASNavReavEnd Then    'fullburn Cycle
       actionSeq = ASidle
      'turn finished, push to next player (for SP thats you)
      PutMsg player.PlayName & " 'baited' the Reaver Cutter", thisPlayer, Logic!Gamecntr
      'change back
      thisPlayer = setPlayer(player.ID, "R", 0)
      If thisPlayer <> player.ID Then
         PutMsg PlayCode(thisPlayer).PlayName & "'s Turn", thisPlayer, Logic!Gamecntr
      End If
      
   ElseIf status = "X" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Reaver Cycle from another Player's Nav move
      MessBox "Move a Reaver 1 sector", "Move a Reaver", "OK"
      'kick it off
      actionSeq = ASNavReav
      
   ElseIf status = "X" And thisPlayer = player.ID And actionSeq = ASNavReavEnd Then    'fullburn Cycle
       actionSeq = ASidle
       PutMsg player.PlayName & " 'summonded' the Reaver Cutter", thisPlayer, Logic!Gamecntr
      'turn finished, push to next player (for SP thats you)
      thisPlayer = setPlayer(player.ID, "R", 0)
      If thisPlayer <> player.ID Then
         PutMsg PlayCode(thisPlayer).PlayName & "'s Turn", thisPlayer, Logic!Gamecntr
      End If
      
   ElseIf status = "Z" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Cruiser Cycle from another Player's Nav move
      x = setPlayer(player.ID, "", 0, True)
      MessBox "Move the Alliance Cruiser adjacent to " & PlayCode(x).PlayName, "Move the Alliance Cruiser", "OK"
      'kick it off
      actionSeq = ASNavCrusAdjacent
         
   ElseIf status = "V" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Corvette Cycle from another Player's Nav move
      x = setPlayer(player.ID, "", 0, True)
      MessBox "Move the Operative's Corvette adjacent to " & PlayCode(x).PlayName, "Move the Operative's Corvette", "OK"
      'kick it off
      actionSeq = ASNavCorvAdjacent
         
   ElseIf status = "Y" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Cruiser Cycle from another Player's Nav move
      MessBox "Move the Alliance Cruiser 1 sector", "Move the Alliance Cruiser", "OK"
      'kick it off
      actionSeq = ASNavCrus
      
   ElseIf (status = "U" Or status = "V" Or status = "Y" Or status = "Z") And thisPlayer = player.ID And actionSeq = ASNavCrusEnd Then      'fullburn Cycle
       actionSeq = ASidle
      'turn finished, push to next player (for SP thats you)
      thisPlayer = setPlayer(player.ID, "R", 0)
      PutMsg player.PlayName & " 'directed' the " & IIf(status = "U" Or status = "V", "Operative's Corvette", "Alliance Cruiser"), thisPlayer, Logic!Gamecntr
  
      If thisPlayer <> player.ID Then
         PutMsg PlayCode(thisPlayer).PlayName & "'s Turn", thisPlayer, Logic!Gamecntr
      End If
   
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASMoseyEnd Then   'Mosey Cycle - your go
      PutMsg player.PlayName & " moseyed to sector " & SectorID, player.ID, Logic!Gamecntr
      resolveToken SectorID
      checkFlacGun SectorID
      actionSeq = ASselect 'in limbo awaiting user to select
      showActions   'throw it back to the action window to resolve end of mosey and offer other actions
   
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASFullburnEnd Then   'fullburn Cycle - your go
      PutMsg player.PlayName & " fullburned to sector " & SectorID, player.ID, Logic!Gamecntr
      If resolveToken(SectorID) = 6 And isOutlaw(player.ID) Then 'no Nav card when Corvette arrives
         frmAction.fullburndone = True
         actionSeq = ASselect 'in limbo awaiting user to select
         showActions   'throw it back to the action window
      Else
         checkFlacGun SectorID
         actionSeq = ASNav 'pick a Nav card
         showNav
      End If
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASnavEnd Then   'fullburn Cycle
      'deal with the Nav option chosen
      If frmNav.NavOption > 0 Then
         doNav frmNav.NavCardID, frmNav.NavOption
         If hasShipUpgrade(player.ID, 20) And TheBigBlack >= 0 Then 'Emissions Recycler
            checkBigBlack frmNav.NavCardID
         End If
      End If
      frmNav.FDPane1.PaneVisible = False
      'avoid special move actions like EVADE
      If actionSeq = ASnavEnd Then 'has not been modified by special moves
        'reset OffJob status
         clearOffJob player.ID
         actionSeq = ASselect 'in limbo awaiting user to select
         showActions   'throw it back to the action window
      End If
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASNavEvadeEnd Then   'fullburn Cycle
      resolveToken SectorID
      checkFlacGun SectorID
      actionSeq = ASselect 'in limbo awaiting user to select
      showActions   'throw it back to the action window
      
   ElseIf status = "R" And thisPlayer = player.ID And (actionSeq = ASNavReavEnd Or actionSeq = ASNavCrusEnd) Then   'fullburn Cycle
      actionSeq = ASselect 'in limbo awaiting user to select
      showActions   'throw it back to the action window
            
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASDeal Then   'Deal Cycle - your go
      If showDeals(False, "locals") = 0 And getUnseenDeck("Contact", ContactID) = 0 And Not HigginsDealPerk Then
         actionSeq = ASselect
      Else
         actionSeq = ASDealSelDiscard
         frmDeal.Timer1.Enabled = False
      End If
      showActions
   
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASDealDrew Then   'Deal Cycle - your go
      'save selected card as Seq = 6
      x = frmDeal.setSelected("UN", CONSIDERED)
      maxConsider = MAXJOBCARDDRAW + getGearFeature(player.ID, "MaxJobs")
      If isSolid(player.ID, 4) And ContactID = 4 Then
         maxConsider = 4
      End If
      'and draw cards up to 3
      If x < maxConsider Then
         DrawDeck "Contact", IIf(HigginsDealPerk, 8, ContactID), maxConsider - x, CONSIDERED
      End If
      actionSeq = ASDealSelect
      showDeals False, "localdeal" 'only show those considered (6)
      showActions
      
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASDealEnd Then   'Deal Cycle - your go
      'save selected (Seq=6 + selected) to players Jobs, unselected back to 5
      x = doDeal(player.ID)
      
      PutMsg player.PlayName & " dealt and accepted " & IIf(x = 0, "no", CStr(x)) & " deals from " & varDLookup("ContactName", "Contact", "ContactID=" & IIf(HigginsDealPerk, 8, ContactID)), player.ID, Logic!Gamecntr
      
      'do any Sell Cargo/Contra Dealing now----------------
      If ContactID = 6 Then  'lord Harrow
         If doBuyCargo(player.ID, Val(frmAction.txtCargo)) > 0 Then
            PutMsg player.PlayName & " bought " & frmAction.txtCargo & " Cargo from " & varDLookup("ContactName", "Contact", "ContactID=" & ContactID), player.ID, Logic!Gamecntr
         End If
         frmAction.txtCargo = "0"
         
      ElseIf ContactID = 9 Then  'FANTY MINGO
         If doBuyContra(player.ID, Val(frmAction.txtContra)) Then
            PutMsg player.PlayName & " bought " & frmAction.txtContra & " Contraband from " & varDLookup("ContactName", "Contact", "ContactID=" & ContactID), player.ID, Logic!Gamecntr
         End If
         frmAction.txtContra = "0"
         
      ElseIf doSellCargoContra(player.ID, ContactID, Val(frmAction.txtCargo), Val(frmAction.txtContra)) > 0 Then
         PutMsg player.PlayName & " sold " & frmAction.txtCargo & " Cargo and " & frmAction.txtContra & " Contraband to " & varDLookup("ContactName", "Contact", "ContactID=" & ContactID), player.ID, Logic!Gamecntr
         frmAction.txtCargo = "0"
         frmAction.txtContra = "0"
         
      End If
      'Deal with Harken to source Fuel (not a Buy action
      If frmAction.txtFuel.Enabled And doBuyFuelParts(player.ID, Val(frmAction.txtFuel), 0, True) <= getMoney(player.ID) And ContactID = 5 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= (Val(frmAction.txtFuel) / 2) Then
            If doBuyFuelParts(player.ID, Val(frmAction.txtFuel), 0) Then
               PutMsg player.PlayName & " bought " & frmAction.txtFuel & " Fuel from Harken", player.ID, Logic!Gamecntr
            End If
         Else
         MessBox "Not enough Cargo Space for the Fuel order", "Cargo Space", "Ooops", "", getLeader()
         End If
         frmAction.txtFuel = "0"
         frmAction.txtParts = "0"
      End If

      
      'clear all Warrants?
      If frmAction.chkWarrant.Value = 1 Then
         If varDLookup("Pay", "Players", "PlayerID=" & player.ID) >= 1000 Then
            DB.Execute "UPDATE Players SET Warrants = 0, Pay = Pay - 1000 WHERE PlayerID=" & player.ID
            PutMsg player.PlayName & " had Badger clear all Warrants", player.ID, Logic!Gamecntr
            frmAction.chkWarrant.Value = 0
         Else
            MessBox "Not enough money left to pay Badger to clear all Warrants", "Warrants", "Ooops", "", getLeader()
         End If
      End If
      
      'load pasengers & Fugitives at Amnons
      If frmAction.txtPass.Visible Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= Val(frmAction.txtPass) + Val(frmAction.txtFug) Then
            DB.Execute "UPDATE Players SET Passenger = Passenger + " & CStr(Val(frmAction.txtPass)) & ", Fugitive = Fugitive + " & CStr(Val(frmAction.txtFug)) & " WHERE PlayerID = " & player.ID
            PutMsg player.PlayName & " loaded " & CStr(Val(frmAction.txtPass)) & " Passengers and " & CStr(Val(frmAction.txtFug)) & " Fugitives", player.ID, Logic!Gamecntr
         End If
         frmAction.txtPass = "0"
         frmAction.txtFug = "0"
      End If
      
      drawLine 1, -1
      actionSeq = ASselect 'in limbo awaiting user to select
      showDeals False, "local"
      If Not (frmJob Is Nothing) Then frmJob.RefreshJobs
      frmDeal.RefreshDeals
      'frmDeal.Timer1.Enabled = True
      showActions

      
    ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASBuy Then   'Buy Cycle - your go
      showBuys False, "local"
      actionSeq = ASBuySelDiscard
      'frmBuy.Timer1.Enabled = False
      showActions
   
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASBuyDrew Then   'Buy Cycle - your go
      'save selected card as Seq = 6
      x = frmBuy.setSelected("UN", CONSIDERED)
      'and draw cards up to 3
      If x < MAXJOBCARDDRAW Then
         DrawDeck "Supply", SupplyID, MAXJOBCARDDRAW - x, CONSIDERED
      End If
      actionSeq = ASBuySelect
      showBuys False, "localbuy" 'only show those considered (6)
      showActions
      
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASBuyShore Then   'Buy Cycle - your go
      x = doShoreLeave(player.ID)
      actionSeq = ASselect 'in limbo awaiting user to select
      
      If getPerkAttributeCrew(player.ID, "FreeShoreLeave") > 0 Then
         PutMsg player.PlayName & " had the Barkeep shout the Crew some free Shore Leave", player.ID, Logic!Gamecntr, True, 71
      ElseIf hasShipUpgrade(player.ID, 19) Then
         PutMsg player.PlayName & " treated the Crew with a shiny Board Game for $" & CStr(x), player.ID, Logic!Gamecntr, True, 0, 0, 19
      Else
         PutMsg player.PlayName & " went on Shore Leave at " & varDLookup("PlanetName", "Planet", "SectorID=" & SectorID) & " for $" & CStr(Abs(x)), player.ID, Logic!Gamecntr
      End If
            
      showActions
      
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASBuyHaven Then   'Buy Cycle - your go
    
      'buy fuel & parts now
      If frmAction.txtFuel.Enabled Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= (Val(frmAction.txtFuel)) / 2 Then
            If doBuyFuelParts(player.ID, Val(frmAction.txtFuel), 0, False, 4) = 0 Then
               PutMsg player.PlayName & " loaded " & frmAction.txtFuel & " Fuel at the Haven, up to 4 for free!", player.ID, Logic!Gamecntr
            End If
         Else
            MessBox "Not enough Cargo Space for the Fuel/Parts order", "Fuel/Parts order", "Ooops", "", getLeader()
         End If
         frmAction.txtFuel = "0"
         frmAction.txtParts = "0"
      End If
      x = 0
      If frmAction.chkShore.Value = 1 Then
         x = doShoreLeave(player.ID, False, True)
      End If
      
      actionSeq = ASselect 'in limbo awaiting user to select
      If x = -1 Then
         PutMsg player.PlayName & " took some free Shore Leave at the Haven", player.ID, Logic!Gamecntr, True, getLeader()
      End If
      showActions
      
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASBuyEnd Then   'Buy Cycle - your go
      'save selected (Seq=6 + selected) to players Jobs, unselected back to 5
      x = doBuy(player.ID)
      PutMsg player.PlayName & " accepted and bought " & IIf(x = 0, "no", CStr(x)) & " buys from " & varDLookup("SupplyName", "Supply", "SupplyID=" & SupplyID), player.ID, Logic!Gamecntr
      
      'buy fuel & parts now
      If frmAction.txtFuel.Enabled Then
         If doBuyFuelParts(player.ID, Val(frmAction.txtFuel), Val(frmAction.txtParts), True) <= getMoney(player.ID) Then
            If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= (Val(frmAction.txtFuel) + Val(frmAction.txtParts)) / 2 Then
               If doBuyFuelParts(player.ID, Val(frmAction.txtFuel), Val(frmAction.txtParts)) Then
                  PutMsg player.PlayName & " bought " & frmAction.txtFuel & " Fuel and " & frmAction.txtParts & " Parts", player.ID, Logic!Gamecntr
               End If
            Else
               MessBox "Not enough Cargo Space for the Fuel/Parts order", "Fuel/Parts order", "Ooops", "", getLeader()
            End If
         Else
            MessBox "Not enough money left to pay for the Fuel or Parts", "Fuel/Parts order", "Ooops", "", getLeader()
         End If
         frmAction.txtFuel = "0"
         frmAction.txtParts = "0"
      End If
      
      actionSeq = ASselect 'in limbo awaiting user to select
      showBuys False, "local"
      'frmBuy.Timer1.Enabled = True
      If frmShip Is Nothing Then Set frmShip = New frmShips
      frmShip.RefreshShips
      frmBuy.RefreshBuys
      showActions
      
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASWork Then   'Work Cycle - your go
      
      If GetCombo(frmAction.cbo) = 0 Then  'make work
         If hasCrew(player.ID, 73) Then  'Busker adds 100
            getMoney player.ID, 300
            PutMsg player.PlayName & " made Extra Work with Busker at " & varDLookup("PlanetName", "Planet", "SectorID=" & SectorID), player.ID, Logic!Gamecntr
         Else
            getMoney player.ID, 200
            PutMsg player.PlayName & " made Work at " & varDLookup("PlanetName", "Planet", "SectorID=" & SectorID), player.ID, Logic!Gamecntr
         End If
         
         If hasCrew(player.ID, 78) And (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID)) >= 1 Then ' Holder- When you Make-Work, you may also take a Fugitive
            If MessBox("Holder has lined up a Fugitive for us" & vbNewLine & "Do you want to take them on board?", "Load 1 Fugitive?", "Yes", "No", 78) = 0 Then
               DB.Execute "UPDATE Players SET Fugitive = Fugitive + 1 WHERE PlayerID =" & player.ID
            End If
         End If
         
         actionSeq = ASselect 'in limbo awaiting user to select
      Else
         If doWork(player.ID, GetCombo(frmAction.cbo)) = 0 Then ' normal exit
            actionSeq = ASselect 'in limbo awaiting user to select
         End If
         'reset OffJob status
         clearOffJob player.ID
      End If
      
      showActions
      
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASRemoveDisgr Then
      removeSelDisgruntled player.ID
      PutMsg player.PlayName & " removed Disgruntle from a Crew", thisPlayer, Logic!Gamecntr
      actionSeq = ASselect 'in limbo awaiting user to select
      showActions
   
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASResolveAlertEnd Then
      actionSeq = ASselect 'in limbo awaiting user to select
      showActions
   
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASEnd Then 'Finish up your turn
      frmAction.FDPane1.PaneVisible = False
      wormHoleOpen = False
      drawLine 2, -1
      'Check if WON!
      If CheckWon(player.ID) Then
         If MessBox("Do you want to end the game for all players?", "End Game?", "Yes", "No", getLeader()) = 0 Then
         'If MsgBox("Do you want to end the game for all players?", vbQuestion + vbYesNo, "End Game?") = vbYes Then
            PutMsg PlayCode(thisPlayer).PlayName & " has ENDED the Game", thisPlayer, Logic!Gamecntr
            Logic.Update "Seq", "E"
            Exit Sub
         End If
      End If
      
      'turn finished, push to next player (for SP thats you)
      thisPlayer = setNextPlayer(player.ID)
      If thisPlayer <> player.ID Then
         PutMsg PlayCode(thisPlayer).PlayName & "'s Turn", thisPlayer, Logic!Gamecntr
      End If
      actionSeq = ASidle
      If Not (frmShip Is Nothing) Then frmShip.RefreshShips
   End If
  
   Exit Sub
  
err_handler:
  errh = MsgBox(Err.Description, vbCritical + vbAbortRetryIgnore, "Error in Main Cycle")
  Select Case errh
  Case vbRetry
    Resume
  Case vbAbort
    'exit
  Case vbIgnore
    Resume Next
  End Select
  
   
   
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim ChatTxt As String, x
  playsnd 13
  Select Case Button.Key
  Case "start"  'host
    
    Select Case Logic!Seq
     Case "E"
       SoloGame = False
       player.ID = 0
       player.Color = ""
       player.PlayName = ""
       Logic.Update "Seq", "H"
       Logic.Update "GameCntr", 0
       ClearBoard
       'Starter.cbo.Enabled = True
       Starter.isHost = True
       Starter.Show 1
              
       If player.ID = 0 Then 'no active player
         Logic.Update "Seq", "E"
         ClearBoard
       Else
         
         pickStartSector = 0
         actionSeq = ASidle
         initBoard
         'Verse.Timer1.Enabled = True
         showEvents
         initToolbar True
         Toolbar1.Buttons("exit").Enabled = True
       End If
       
     Case "H"
        If MsgBox("Game already being hosted, do you need to reset it?", vbYesNo + vbCritical, "Game in Host mode") = vbYes Then
            Logic.Update "Seq", "E"
        End If

     Case Else
        x = MsgBox("Game in progress. If you want to re-join, use JOIN button." & vbNewLine & "otherwise press OK to reset the Game", vbExclamation + vbOKCancel, "Game in Progress")
        Select Case x
        Case vbOK
            Logic.Update "Seq", "E"
            MsgBox "Game has been reset, press Host to start", vbInformation
        End Select
     End Select
       
  Case "join"
     
     Select Case Logic!Seq
     Case "H", "E"
         player.ID = 0
         player.Color = ""
         player.PlayName = ""
         Starter.isHost = False
         Starter.Show 1
         
         If player.ID > 0 Then
            initBoard
            'Verse.Timer1.Enabled = True
            pickStartSector = 0
            actionSeq = ASidle
            showEvents
            initToolbar True, False
         End If

     Case Else
        If MessBox("Do you want to rejoin in " & IIf(getPlayerCount() > 1, "Multiplayer", "Single Player") & " mode?" & vbNewLine & vbNewLine & "Press OK to join the Game", "Game in Progress", "OK", "Cancel", 1) = 0 Then
        'x = MsgBox("Game in progress. Do you want to rejoin it in " & IIf(getPlayerCount() > 1, "MP", "SP") & " mode?" & vbNewLine & "press OK to join the Game", vbExclamation + vbOKCancel, "Game in Progress")
        'Select Case x
        'Case vbOK
            player.ID = getNewPlayer()
            player.PlayName = Nz(varDLookup("Name", "Players", "PlayerID =" & player.ID))
            SoloGame = (getPlayerCount(True) = 1)
            pickStartSector = 2  'flag the selection is done
            actionSeq = ASidle
            initBoard
            'Verse.Timer1.Enabled = True
            showEvents
            initToolbar True, False
        End If
     End Select

  Case "exit"  'END the Game
    ' Confirm Exit msgbox ?
    If MessBox("Are you sure you want to leave this game?", "Closing Game?", "Yes", "No") = 1 Then
       Exit Sub
    End If

    killAllForms
    
    If Logic!player = player.ID Then setNextPlayer player.ID
    DB.Execute "Update Players Set Name = NULL WHERE PlayerID = " & player.ID
    If Nz(varDLookup("PlayerID", "Players", "Name IS NOT NULL"), 0) = 0 Then
       Logic.Update "Seq", "E"
       Logic.Update "GameCntr", 0
    End If

    player.ID = 0
    player.Color = ""
    player.PlayName = ""
    pickStartSector = -1
    actionSeq = ASidle
    initToolbar False
    
  Case "chat"
    ChatTxt = InputBox("Enter your message", "Chat")
    If ChatTxt <> "" Then
      PutMsg player.PlayName & " : " & ChatTxt, 0
    End If
    
  Case "graph"
   If frmStat Is Nothing Then
       Set frmStat = New frmStats
       frmStat.FDPane1.PaneVisible = False
    End If
    If frmStat.FDPane1.PaneVisible = False Then
      showStats
    Else
       frmStat.FDPane1.PaneVisible = False
       frmStat.Timer1.Enabled = False
    End If
    DockFrame1.SaveStates "Firefly"
    
  Case "log"

    Events.FDPane1.PaneVisible = Not Events.FDPane1.PaneVisible
    DockFrame1.SaveStates "Firefly"
        
  Case "crew"
    Dim frmCrew As New frmCrewSel
    frmCrew.crewFilter = " Order By CrewName"
    frmCrew.AlwaysOnTop = True
    frmCrew.Show
    Set frmCrew = Nothing
  
    
  Case "gear"
    Dim frmGear As New frmGearView
    frmGear.gearFilter = " Order By GearName"
    frmGear.AlwaysOnTop = True
    frmGear.Show
    Set frmGear = Nothing
    
  Case "upgrd"
    Dim frmUpGrd As New frmShipUpgrdView
    frmUpGrd.gearFilter = " Order By UpgradeName"
    frmUpGrd.AlwaysOnTop = True
    frmUpGrd.Show
    Set frmUpGrd = Nothing
  
    
  Case "job"
    If frmJob Is Nothing Then
       Set frmJob = New frmJobs
       frmJob.FDPane1.PaneVisible = False
    End If
    With frmJob
    If .FDPane1.PaneVisible Then
        .FDPane1.PaneVisible = False
    Else
        If .jobFilter = "" Then
            .Caption = player.PlayName & "'s Jobs"
            .jobFilter = ""
        End If
        .FDPane1.InitDockHW = 120
        .FDPane1.InitDockStyle = DockToTop
        .FDPane1.PaneVisible = True
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        .Show
        .RefreshJobs
    End If
    End With

    
  Case "deal"
     If actionSeq < ASDeal Or actionSeq > ASDealEnd Then
       showDeals True
     End If
    
  Case "buy"
    If actionSeq < ASBuy Or actionSeq > ASBuyEnd Then
      showBuys True
    End If
    
  Case "ship"
    If frmShip Is Nothing Then
       Set frmShip = New frmShips
       frmShip.FDPane1.PaneVisible = False
    End If
    With frmShip
      If .FDPane1.PaneVisible Then
          .FDPane1.PaneVisible = False
      Else
           If .shipFilter = "" Then
              .Caption = player.PlayName & "'s Ship"
              .shipFilter = "me"
           End If
          .FDPane1.InitDockHW = 400
          .FDPane1.InitDockStyle = DockToTop
          .FDPane1.PaneVisible = True
          .FDPane1.PinState = Pinned
          .FDPane1.SetLayoutReference Nothing
          .Show
          .RefreshShips
      End If
    End With
  
  Case "help"
    x = ShellExecute(x, "OPEN", App.Path & "\FireflyForPC.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
  
  End Select
End Sub

Private Sub initBoard()
Dim rst As New ADODB.Recordset
Dim coords, c() As String, x
   NumOfReavers = varDLookup("NoOfReavers", "Story", "StoryID = " & Logic!StoryID)
   If Verse Is Nothing Then
      Set Verse = New Board
   End If
   With Verse
      
      .Picture1.Picture = LoadPicture(App.Path & "\Pictures\" & Logic!BoardPicture)
      .Height = Logic!BHeight
      .Width = Logic!BWidth
   
      For x = 1 To 4
         Load .imgHaven(x)
         Set .imgHaven(x).Container = .Picture1
         .imgHaven(x).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Haven" & x & ".bmp")
         .imgHaven(x).TransparentColor = &HFFFFFF
         .imgHaven(x).TransparentColorMode = lvicUseTransparentColor
      Next x
   
      rst.Open "SELECT * FROM Board WHERE SectorID > 0 ORDER BY SectorID", DB, adOpenDynamic, adLockOptimistic
      While Not rst.EOF
         Load .HotSpot(rst!SectorID)
         Set .HotSpot(rst!SectorID).Container = .Picture1
         .HotSpot(rst!SectorID).top = rst!STop
         .HotSpot(rst!SectorID).Left = rst!SLeft
         .HotSpot(rst!SectorID).Height = rst!SHeight
         .HotSpot(rst!SectorID).Width = rst!SWidth
         .HotSpot(rst!SectorID).ZOrder
         .HotSpot(rst!SectorID).Visible = True
         coords = rst.Fields("Slot5").Value
         c = Split(coords, ",")
         
         Load .imgAToken(rst!SectorID)
         .imgAToken(rst!SectorID).Left = c(0)
         .imgAToken(rst!SectorID).top = c(1)
         Set .imgAToken(rst!SectorID).Container = .Picture1
         If rst!AToken > 0 Then
            .imgAToken(rst!SectorID).Picture = LoadPictureGDIplus(App.Path & "\Pictures\AToken" & IIf(rst!AToken > 6, 6, rst!AToken) & ".bmp")
            .imgAToken(rst!SectorID).Visible = True
         End If
         .imgAToken(rst!SectorID).TransparentColor = &HFFFFFF
         .imgAToken(rst!SectorID).TransparentColorMode = lvicUseTransparentColor
         
         Load .imgToken(rst!SectorID)
         .imgToken(rst!SectorID).Left = c(0) + 100
         .imgToken(rst!SectorID).top = c(1) + 100
         Set .imgToken(rst!SectorID).Container = .Picture1
         If rst!Token > 0 Then
            .imgToken(rst!SectorID).Picture = LoadPictureGDIplus(App.Path & "\Pictures\RToken" & IIf(rst!Token > 6, 6, rst!Token) & ".bmp")
            .imgToken(rst!SectorID).Visible = True
         End If
         .imgToken(rst!SectorID).TransparentColor = &HFFFFFF
         .imgToken(rst!SectorID).TransparentColorMode = lvicUseTransparentColor
         
         If rst!Haven > 0 Then
            .imgHaven(rst!Haven).Left = c(0)
            .imgHaven(rst!Haven).top = c(1)
            .imgHaven(rst!Haven).Visible = True
         End If
         
         rst.MoveNext
      Wend
      For x = 5 To 6 + NumOfReavers ' .Imag.Count
         .Imag(x).Animate2.StartAnimation
      Next x
      .Caption = "the 'Verse - " & varDLookup("StoryTitle", "Story", "StoryID=" & Logic!StoryID)
      .Show
      
      
   End With
End Sub


Private Sub initToolbar(ByVal start As Boolean, Optional ByVal admin As Boolean = True)
   With Toolbar1
      .Buttons("exit").Enabled = start
      .Buttons("chat").Enabled = start
      .Buttons("graph").Enabled = start
      .Buttons("log").Enabled = start
      .Buttons("crew").Enabled = start
      .Buttons("gear").Enabled = start
      .Buttons("upgrd").Enabled = start
      .Buttons("job").Enabled = start
      .Buttons("ship").Enabled = start
      .Buttons("deal").Enabled = start
      .Buttons("buy").Enabled = start
      .Buttons("start").Enabled = Not start
      .Buttons("join").Enabled = Not start
      .Buttons("help").ButtonMenus("jobs").Enabled = admin
   End With
End Sub

Private Sub killAllForms()

    If Verse.Visible Then
       Verse.Hide
       Unload Verse
       Set Verse = Nothing
    End If

    If frmJob Is Nothing Then
    Else
       Unload frmJob
       Set frmJob = Nothing
    End If
    If frmShip Is Nothing Then
    Else
       Unload frmShip
       Set frmShip = Nothing
    End If
    If frmDeal Is Nothing Then
    Else
       Unload frmDeal
       Set frmDeal = Nothing
    End If
    If frmBuy Is Nothing Then
    Else
       Unload frmBuy
       Set frmBuy = Nothing
    End If
    If frmStat Is Nothing Then
    Else
      Unload frmStat
      Set frmStat = Nothing
    End If
    If Events.FDPane1.PaneVisible Then Unload Events
    If frmAction.FDPane1.PaneVisible Then Unload frmAction



End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim frmJobEdit As frmJobEditor, x

   playsnd 13
   Select Case ButtonMenu.Key
   Case "alljobs"
      If actionSeq < ASDeal Or actionSeq > ASDealEnd Then
         If frmJob Is Nothing Then
            Set frmJob = New frmJobs
            frmJob.FDPane1.PaneVisible = False
         End If
         frmJob.Caption = "All Jobs"
         frmJob.jobFilter = "all"
         frmJob.RefreshJobs
         frmJob.Show
         frmJob.FDPane1.PaneVisible = True
      End If
      
   Case "myjobs"
      If actionSeq < ASDeal Or actionSeq > ASDealEnd Then
         If frmJob Is Nothing Then
            Set frmJob = New frmJobs
            frmJob.FDPane1.PaneVisible = False
         End If
         frmJob.Caption = player.PlayName & "'s Jobs"
         frmJob.jobFilter = ""
         frmJob.RefreshJobs
         frmJob.Show
         frmJob.FDPane1.PaneVisible = True
      End If
      
   Case "alldeals"
     If actionSeq < ASDeal Or actionSeq > ASDealEnd Then
       showDeals True
     End If

   Case "localdeals"
      showDeals False, "local"

   Case "allbuys"
      If actionSeq < ASBuy Or actionSeq > ASBuyEnd Then
         showBuys
      End If

   Case "localbuys"
      If actionSeq < ASBuy Or actionSeq > ASBuyEnd Then
         showBuys False, "local"
      End If
      
   Case "allships"
      If frmShip Is Nothing Then
         Set frmShip = New frmShips
         frmShip.FDPane1.PaneVisible = False
      End If
      frmShip.Caption = "All Ships"
      frmShip.shipFilter = "all"
      frmShip.RefreshShips
      frmShip.Show
      frmShip.FDPane1.PaneVisible = True
   Case "myship"
      If frmShip Is Nothing Then
         Set frmShip = New frmShips
         frmShip.FDPane1.PaneVisible = False
      End If
      frmShip.Caption = player.PlayName & "'s Ship"
      frmShip.shipFilter = "me"
      frmShip.RefreshShips
      frmShip.Show
      frmShip.FDPane1.PaneVisible = True
      
   Case "firefly"
      x = ShellExecute(x, "OPEN", App.Path & "\Firefly_rulebook.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind

   Case "bluesun"
      x = ShellExecute(x, "OPEN", App.Path & "\FireflyBlueSun_rulebook.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
   Case "kalidasa"
      x = ShellExecute(x, "OPEN", App.Path & "\FireflyKalidasa_rulebook.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
   Case "pcguide"
      x = ShellExecute(x, "OPEN", App.Path & "\FireflyForPC.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
   Case "about"
      MessBox "Firefly + Blue Sun/Kalidasa  V" & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & "*Freeware* - use at your own risk" & vbNewLine & "Made by: Vee Bee-er (c)2021-22 BLiSoftware", "About", "Shiny"
  Case "jobs"
    Set frmJobEdit = New frmJobEditor
    frmJobEdit.Show 1
   End Select

End Sub

Private Sub showStats()
    With frmStat
        .FDPane1.InitDockHW = 200
        .FDPane1.InitDockStyle = DockToLeft
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        .FDPane1.PaneVisible = True
        .refreshform
        .Timer1.Enabled = True
    End With
End Sub
Private Sub showEvents()
    With Events
        .FDPane1.InitDockHW = 200
        .FDPane1.InitDockStyle = DockToLeft
        .FDPane1.PaneVisible = True
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        .Timer1.Enabled = True
        '.Show
    End With
End Sub

Public Function showDeals(Optional ByVal toggle As Boolean = False, Optional ByVal filter As String = "all") As Variant

    If frmDeal Is Nothing Then
       Set frmDeal = New frmDeals
       frmDeal.FDPane1.PaneVisible = False
    End If
    With frmDeal
    If .FDPane1.PaneVisible And toggle Then
        .FDPane1.PaneVisible = False
    Else
        .dealFilter = filter
        .FDPane1.InitDockHW = 400
        .FDPane1.InitDockStyle = DockToTop
        .FDPane1.PaneVisible = True
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        '.Show
        showDeals = .RefreshDeals
    End If
    End With
    
End Function


Public Sub showBuys(Optional ByVal toggle As Boolean = False, Optional ByVal filter As String = "all")

    If frmBuy Is Nothing Then
       Set frmBuy = New frmSupply
       frmBuy.FDPane1.PaneVisible = False
    End If
    With frmBuy
    If .FDPane1.PaneVisible And toggle Then
        .FDPane1.PaneVisible = False
    Else
        .buyFilter = filter
        .FDPane1.InitDockHW = 400
        .FDPane1.InitDockStyle = DockToTop
        .FDPane1.PaneVisible = True
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        .RefreshBuys
    End If
    End With
    
End Sub

Private Sub showNav(Optional ByVal CardID As Integer = 0)
Dim SQL, SectorID, reshuffle, Zone
Dim rst As New ADODB.Recordset

   With frmNav
      .FDPane1.InitDockHW = 200
      .FDPane1.InitDockStyle = DockToLeft
      .FDPane1.PinState = Pinned
      .FDPane1.SetLayoutReference Nothing
      
      .NavCardID = 0
      .NavOption = 0
      
      SectorID = Nz(varDLookup("SectorID", "Players", "PlayerID=" & player.ID), 0)
      Zone = varDLookup("Zones", "Board", "SectorID=" & SectorID)
      
      'Read in the next NAV card and display either 1 or 2 options
      
       'OPTION 1 ===================================================================================
      SQL = "SELECT NavDeck.CardID, NavDeck.CardName, NavDeck.Reshuffle, NavDeck.Seq, NavOption.* "
      SQL = SQL & "FROM NavOption INNER JOIN NavDeck ON NavOption.OptionID = NavDeck.Option1ID "
      If CardID <> 0 Then
         SQL = SQL & "Where NavDeck.CardID = " & CardID
      Else
         SQL = SQL & "Where NavDeck.Zones = '" & Zone & "' And NavDeck.Seq > 6 "
         SQL = SQL & "ORDER BY NavDeck.Seq"
      End If
      rst.Open SQL, DB, adOpenDynamic, adLockOptimistic
      If rst.EOF Then  ' this happens when the reshuffle card is in the discard pile at start of game setup
         ShuffleDeck "Nav", True, False, Zone
         PutMsg player.PlayName & " Reshuffling NavDeck " & Zone & " due to end of deck", player.ID, Logic!Gamecntr, True, getLeader()
         rst.Close
         rst.Open SQL, DB, adOpenDynamic, adLockOptimistic
      End If
      If Not rst.EOF Then
         'a LOT of these tests are only applied to the 1st option only
         .cmd(0).Enabled = hasNavReqs(player.ID, rst!CardID, 1)
                  
         If (rst!CardName = "Reaver Cutter!") And getCruiserCorvette(SectorID) = 6 Then 'corvette shoos the Reavers away
            frmNav.NavOption = 0
            actionSeq = ASnavEnd
            .NavCardID = 0
            PutMsg player.PlayName & " is Shielded from a Reaver Cutter attack by the Alliance Corvette", player.ID, Logic!Gamecntr, True, getLeader()
            
         'skip Customs Inspection if solid with Harken
         ElseIf (rst!CardName = "Customs Inspection") And isSolid(player.ID, 5) Then
            frmNav.NavOption = 0
            actionSeq = ASnavEnd
            .NavCardID = 0
            PutMsg player.PlayName & " being Solid with Harken avoided a Customs Inspection", player.ID, Logic!Gamecntr, True, getLeader()
         Else
            .NavCardID = rst!CardID
            .lblName.Caption = rst!CardName
            .cmd(0).Caption = rst!OptionName
            .cmd(0).ToolTipText = rst!OptionName
            .lblDetail(0).Caption = rst!Details
         End If
         reshuffle = rst!reshuffle
         'pull the card out of the deck, assign it to the user for debugging
         rst!Seq = player.ID
         rst.Update
      Else
         MsgBox "Error in NavDeck"
         Exit Sub
      End If
      rst.Close
      
      'OPTION 2 ===================================================================================
      SQL = "SELECT NavOption.* "
      SQL = SQL & "FROM NavOption INNER JOIN NavDeck ON NavOption.OptionID = NavDeck.Option2ID "
      SQL = SQL & "Where NavDeck.CardID = " & .NavCardID
      
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         .cmd(1).Visible = True
         .cmd(1).Enabled = hasNavReqs(player.ID, .NavCardID, 2)
         
         .lblDetail(0).Height = 1125
         .lblDetail(1).Visible = True

         .cmd(1).Caption = rst!OptionName
         .cmd(1).ToolTipText = rst!OptionName
         .lblDetail(1).Caption = rst!Details
      Else 'no option 2
         .cmd(1).Visible = False
         .lblDetail(1).Visible = False
         .lblDetail(0).Height = 2085
      End If
      rst.Close
      
      If reshuffle = 1 Then 'ready for next turn
         ShuffleDeck "Nav", True, False, Zone
         PutMsg player.PlayName & " Reshuffling NavDeck " & Zone & " due to reshuffle card", player.ID, Logic!Gamecntr, True, getLeader()
      End If
      
      .lblUnseen = getZoneDesc(Zone) & "unseen: " & getUnseenNavDeck(Zone)
      
      If .NavCardID <> 0 Then
         .FDPane1.PaneVisible = True
      End If
      
   End With
      
End Sub

Public Sub showActions()
Dim SQL, SectorID, onlyFullburn As Boolean
Dim rst As New ADODB.Recordset, reaverActive As Boolean, moseyrng As Integer
Dim frmJoSel As frmJobSel

   SectorID = getPlayerSector(player.ID)
   SoloGame = isSoloGame() 'as a player may drop out
   
   If ignoreToken <> SectorID Then resolveToken SectorID
   
   'check that the REAVER is or is not here
   If getCutterSector(SectorID) > 0 Then
      checkFlacGun SectorID 'possibly chase it away
   End If
   If getCutterSector(SectorID) > 0 And frmAction.checkNoOfActions = 0 And FullburnMovesDone = 0 And MoseyMovesDone = 0 And CruiserCutter <> SectorID Then
      reaverActive = True
      showNav -1
      CruiserCutter = SectorID
   End If
   
   If getCruiserCorvette(SectorID) = 5 And CruiserCutter <> SectorID Then
      CruiserCutter = SectorID 'set it as faced regardless of outcome
      If isOutlaw(player.ID) And actionSeq <> ASNavEvade Then  'it just arrived so face it
         showNav -2
         frmAction.moseydone = True 'Full Stop!
         frmAction.fullburndone = True
         Exit Sub
      End If
   End If
   If getCruiserCorvette(SectorID) = 6 And CorvetteSeq <> getCorvetteSeq Then
      CorvetteSeq = getCorvetteSeq
      CruiserCutter = SectorID 'set it as faced regardless of outcome
      If isOutlaw(player.ID) And actionSeq <> ASNavEvade Then  'it just arrived so face it
         showNav -3
         frmAction.moseydone = True 'Full Stop!
         frmAction.fullburndone = True
         Exit Sub
      End If
   End If
   
   'check active job limit not exceeded
   While getPlayerJobs(player.ID, "1,2") > MAXACTIVEJOBS + IIf(isSolid(player.ID, 8), 1, 0)
      Set frmJoSel = New frmJobSel
      frmJoSel.jobFilter = " IN (1,2)"
      frmJoSel.Caption = "Too many active Jobs, select one to remove"
      
      frmJoSel.Show 1
      If frmJoSel.CardID > 0 Then
        removeJob player.ID, frmJoSel.CardID
      End If
      Set frmJoSel = Nothing
   Wend
   'check inactive job limit not exceeded
   While getPlayerJobs(player.ID, "0") > MAXINACTIVEJOBS
      Set frmJoSel = New frmJobSel
      frmJoSel.jobFilter = "=0"
      frmJoSel.Caption = "Too many inactive Jobs, select one to remove"
      
      frmJoSel.Show 1
      If frmJoSel.CardID > 0 Then
        removeJob player.ID, frmJoSel.CardID
      End If
      Set frmJoSel = Nothing
   Wend
   
   With frmAction
      'check if action limit reached
      If Not SoloGame And .checkNoOfActions > 1 And actionSeq <> ASNavEvade Then
         .endAction
         Exit Sub
      'check if we are currently in Fullburn/Mosey on the 2nd action to diable other actions
      ElseIf Not SoloGame And .checkNoOfActions = 1 And ((.fullburndone = False And FullburnMovesDone > 0) Or (.moseydone = False And MoseyMovesDone > 0)) Then
         onlyFullburn = True
      End If
      
      .FDPane1.InitDockHW = 200
      .FDPane1.InitDockStyle = DockToLeft
      .FDPane1.PinState = Pinned
      .FDPane1.SetLayoutReference Nothing
      
      .lblGo.Caption = CStr(Logic!Gamecntr - 1)
      
      .lblMoney.Tag = varDLookup("Pay", "Players", "PlayerID=" & player.ID)
      .lblMoney = "$" & .lblMoney.Tag
      If Val(.lblMoney.Tag) < 200 Then
         .lblMoney.ForeColor = 0
         .lblMoney.BackColor = 11468799
      Else
         .lblMoney.ForeColor = 16777215
         .lblMoney.BackColor = 8388736
      End If
      
      
      'load Fuel stats,
      .lblFuelOn.Caption = varDLookup("Fuel", "Players", "PlayerID=" & player.ID)
      Select Case Val(.lblFuelOn.Caption)
      Case 0
         .lblFuelOn.BackColor = 5987327
      Case 1, 2
         .lblFuelOn.BackColor = 9109503
      Case Else
         .lblFuelOn.BackColor = &HCBE1ED
      End Select
      SQL = "SELECT SUM(ShipUpgrade.BurnRange) AS BurnRange, MAX(ShipUpgrade.BurnFuel) AS BurnFuel, MAX(ShipUpgrade.MoseyRange) AS MoseyRange"
      SQL = SQL & " FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID"
      SQL = SQL & " WHERE PlayerSupplies.PlayerID=" & player.ID   'ShipUpgrade.DriveCore=1 AND
      
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         
         '>>>>>  CMD  FULLBURN  <<<<
         .cmd(1).Caption = "Full Burn"
         .lblRange.Caption = 5 + rst!BurnRange + getRangeMod(player.ID, 1) + IIf(.chkRange2.Value = 1, 2, 0) + turnExtraRange - FullburnMovesDone 'ADD WASH's extra Range
         If Val(.lblRange.Caption) = 0 Then .fullburndone = True
         If Not SoloGame And .checkNoOfActions > 1 Then
            .endAction
            Exit Sub
         End If
         .cmd(1).Enabled = (((rst!burnFuel + getExtraBurn(player.ID) + .chkRange2.Value) <= Val(.lblFuelOn.Caption)) Or FullburnMovesDone > 0) And (Not .fullburndone) And (actionSeq = ASselect) And Not reaverActive And hasValidFBMove(player.ID) And Not (HemmorrhagingFuel And FullburnMovesDone > 0 And Val(.lblFuelOn.Caption) = 0)
         
         'single use extended Range
         .chkRange2.Visible = (hasShipUpgrade(player.ID, 17) > 0)
         .chkRange2.Enabled = (hasShipUpgrade(player.ID, 17) > 0 And FullburnMovesDone = 0 And .cmd(1).Enabled And Val(.lblFuelOn.Caption) >= rst!burnFuel + getExtraBurn(player.ID))
         .lblRange2.Visible = (hasShipUpgrade(player.ID, 17) > 0)
         
         .lblFuelRq.Caption = rst!burnFuel + getExtraBurn(player.ID) + .chkRange2.Value
         If getExtraBurn(player.ID) > 0 Then
            .lblFuelRq.BackColor = 9109503
         Else
            .lblFuelRq.BackColor = &HCBE1ED
         End If
         '>>>>>  CMD  MOSEY  <<<<
         .cmd(0).Caption = "Mosey"
         If hasShipUpgrade(player.ID, 7) Then
            moseyrng = 2 + getRangeMod(player.ID, 2)
         Else
            moseyrng = rst!MoseyRange + getRangeMod(player.ID, 2)
         End If
         .lblMosey.Caption = moseyrng - MoseyMovesDone
         If moseyrng = MoseyMovesDone Then .moseydone = True
         If Not SoloGame And .checkNoOfActions > 1 Then
            .endAction
            Exit Sub
         End If
         .cmd(0).Enabled = (moseyrng > MoseyMovesDone) And (actionSeq = ASselect) And (Not .moseydone) And Not reaverActive
         
      End If

      rst.Close
      
      'load Supply in this sector
      .lblSupply.Caption = varDLookup("SupplyName", "Supply", "SectorID=" & SectorID) & ""
      .lblSupply.Tag = varDLookup("SupplyID", "Supply", "SectorID=" & SectorID) & ""
      If .lblSupply.Tag = "" Then
         .lblSupply.BackColor = &HCBE1ED
      Else
         .lblSupply.BackColor = varDLookup("Colour", "Supply", "SectorID=" & SectorID)
      End If
      
      '>>>>>  CMD  BUY  <<<<
      Select Case actionSeq
         Case ASBuy
            'Beep
         Case ASBuySelDiscard
            .cmd(2).FontSize = 7
            .cmd(2).Caption = "Draw Cards"
         Case ASBuyDrew
            'should never happen, moves straight to ASBuySelect
         Case ASBuySelect
            .cmd(2).FontSize = 7
            .cmd(2).Caption = "Close Buy"
         Case Else
            .cmd(2).FontSize = 8
            .cmd(2).Caption = "Buy"
      End Select
      
      'SHORE LEAVE
      .chkShore.Value = 0
      .cmd(2).Enabled = False
      If (Not .buydone) And (Not onlyFullburn) And Not reaverActive Then  ' Buy and Shore leave *may* be active
         
         .chkShore.Enabled = (Nz(varDLookup("SupplyID", "Supply", "SectorID=" & SectorID), 0) > 0 Or hasShipUpgrade(player.ID, 19) Or getHaven(SectorID) = player.ID) And hasDisgruntled(player.ID) And (Abs(doShoreLeave(player.ID, True)) <= getMoney(player.ID) Or getHaven(SectorID) = player.ID)
         
         If (.lblSupply.Caption <> "") And (actionSeq = ASselect Or (actionSeq = ASBuySelDiscard And getUnseenDeck("Supply", Val(.lblSupply.Tag)) > 0) Or actionSeq = ASBuySelect) Then 'we can BUY
            .cmd(2).Enabled = True
         ElseIf getHaven(SectorID) = player.ID And CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0 Then
            .cmd(2).Enabled = True
         ElseIf .chkShore.Enabled And actionSeq = ASselect Then 'shore leave ONLY
            .cmd(2).Enabled = True
            .cmd(2).FontSize = 7
            .cmd(2).Caption = "Shore Leave"
            .chkShore.Value = 1
         End If
      
      Else 'nothing is enabled
         .chkShore.Enabled = False
         .cmd(2).Enabled = False
      End If
      .chkShore.Visible = .chkShore.Enabled
      .Label4.Visible = .chkShore.Enabled
      .Label7.Visible = .chkShore.Enabled
      
      'FUEL & PARTS
      '.txtFuel.Enabled = ((Nz(varDLookup("SupplyID", "Supply", "SectorID=" & SectorID), 0) > 0) And (Not .buydone) And CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0) Or (Nz(varDLookup("ContactID", "Contact", "SectorID=" & SectorID), 0) = 5 And isSolid(player.ID, 5))
      .txtFuel.Enabled = CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0 And (((Nz(varDLookup("SupplyID", "Supply", "SectorID=" & SectorID), 0) > 0) And (Not .buydone)) Or (Nz(varDLookup("ContactID", "Contact", "SectorID=" & SectorID), 0) = 5 And isSolid(player.ID, 5)) Or getHaven(SectorID) = player.ID)
      .txtParts.Enabled = ((Nz(varDLookup("SupplyID", "Supply", "SectorID=" & SectorID), 0) > 0) And (Not .buydone) And CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0)
            
                     
      'load Dealer in this sector
      .lblContact.Tag = varDLookup("ContactID", "Contact", "SectorID=" & SectorID) & ""
      .lblContact.Caption = varDLookup("ContactName", "Contact", "SectorID=" & SectorID) & ""
      
      If (.lblContact.Caption = "" And hasCrew(player.ID, 75)) Or HigginsDealPerk Then
         .lblContact.Caption = "Mag. Higgins"
         .lblContact.Tag = "8"
      End If
      
      If Val(.lblContact.Tag) > 0 Then
         .lblContact.Caption = .lblContact.Caption & IIf(isSolid(player.ID, Val(.lblContact.Tag)), " -(S)", "")
      End If
      
      If .lblContact.Tag = "" Then  'nothing doing here
         .lblContact.BackColor = &HCBE1ED
         .txtCargo.Enabled = False
         .txtContra.Enabled = False
         
      ElseIf .lblContact.Tag = "6" Then 'harrow
         .lblContact.BackColor = varDLookup("Colour", "Contact", "ContactID=" & .lblContact.Tag)
         .txtCargo.Enabled = (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= 1) And isSolid(player.ID, 6) And getMoney(player.ID) >= 300
         .txtCargo.ToolTipText = "buy Cargo for $300ea"
         .txtContra.Enabled = False

      ElseIf .lblContact.Tag = "9" Then 'fanty mingo
         .lblContact.BackColor = varDLookup("Colour", "Contact", "ContactID=" & .lblContact.Tag)
         .txtContra.Enabled = (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= 1) And isSolid(player.ID, 9) And getMoney(player.ID) >= 400
         .txtCargo.Enabled = False
         .txtContra.ToolTipText = "buy Contraband @ $400ea"

      Else  'regular Contact
         .lblContact.BackColor = varDLookup("Colour", "Contact", "ContactID=" & .lblContact.Tag)
         .txtCargo.Enabled = (doSellCargoContra(player.ID, .lblContact.Tag, 1, 0, True) > 0)
         .txtContra.Enabled = (doSellCargoContra(player.ID, .lblContact.Tag, 0, 1, True) > 0)
         .txtCargo.ToolTipText = "sell Cargo to Contact"
         .txtContra.ToolTipText = "sell Contraband to Contact"
      End If
         
      '>>>>>  CMD  DEAL  <<<<
      Select Case actionSeq
         Case ASDeal
            Beep
         Case ASDealSelDiscard
            .cmd(3).FontSize = 7
            .cmd(3).Caption = "Draw Cards"
         Case ASDealDrew
            'should never happen, moves straight to ASDealSelect
         Case ASDealSelect
            .cmd(3).FontSize = 7
            .cmd(3).Caption = "Close Deal"
         Case Else
            .cmd(3).FontSize = 8
            .cmd(3).Caption = "Deal"
      End Select
      
      .cmd(3).Enabled = .lblContact.Caption <> "" And (actionSeq = ASselect Or (actionSeq = ASDealSelDiscard And getUnseenDeck("Contact", Val(.lblContact.Tag)) > 0) Or actionSeq = ASDealSelect) And (Not .dealdone) And (Not onlyFullburn) And Not reaverActive
      
      'Remove Warrants with Badger
      .chkWarrant.Visible = (varDLookup("Warrants", "Players", "PlayerID=" & player.ID) > 0) And (Nz(varDLookup("ContactID", "Contact", "SectorID=" & SectorID), 0) = 2) And isSolid(player.ID, 2) And (Val(.lblMoney.Tag) > 1000)
      .Label10.Visible = .chkWarrant.Visible
      .Label11.Visible = .chkWarrant.Visible
      .cbo.Clear
      
      'Load WORK Combo with Make Work & Jobs in this Sector
      SQL = "SELECT ContactDeck.CardID, Job.JobID AS JOB1, Job.JobDesc AS JOBDES1, Job.SectorID AS SECTOR1, Job_1.JobID AS JOB2, Job_1.JobDesc AS JOBDES2, Job_1.SectorID AS SECTOR2, "
      SQL = SQL & "PlayerJobs.JobStatus, ContactDeck.Immoral, ContactDeck.JobName, Job_2.JobID AS JOB3, Job_2.JobDesc AS JOBDES3, Job_2.SectorID AS SECTOR3 "
      SQL = SQL & "FROM (Job INNER JOIN ((PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) LEFT JOIN "
      SQL = SQL & "Job AS Job_1 ON ContactDeck.Job2ID = Job_1.JobID) ON Job.JobID = ContactDeck.Job1ID) LEFT JOIN Job AS Job_2 ON ContactDeck.Job3ID = Job_2.JobID "
      
      SQL = SQL & "Where PlayerJobs.PlayerID = " & player.ID & " And (Job.SectorID IN (1,2," & SectorID & ") Or Job_1.SectorID IN (1,2," & SectorID & ") Or Job_2.SectorID IN (1,2," & SectorID & "))"
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      .cbo.Clear
      While Not rst.EOF
         If ((rst!sector1 = 1 And getCruiserSector() = SectorID) Or (rst!sector1 = 2 And getCorvetteSector() = SectorID) Or (SectorID = rst!sector1)) And rst!JobStatus = 0 And getPlayerJobs(player.ID, "1,2") < MAXACTIVEJOBS + IIf(isSolid(player.ID, 8), 1, 0) Then ' check requirements met for job
            If hasJobReqs(player.ID, rst!CardID, rst!Job1) Then
               .cbo.AddItem rst!Jobdes1 & " (" & CStr(rst!CardID) & ")"
               .cbo.ItemData(.cbo.NewIndex) = rst!CardID
            End If
            
         ElseIf ((rst!Sector3 = 1 And getCruiserSector() = SectorID) Or (rst!Sector3 = 2 And getCorvetteSector() = SectorID) Or (SectorID = rst!Sector3)) And rst!JobStatus = 1 Then 'Job3 must be in the sector
            If hasJobReqs(player.ID, rst!CardID, rst!Job3) Then
               .cbo.AddItem rst!Jobdes3 & " (" & CStr(rst!CardID) & ")"
               .cbo.ItemData(.cbo.NewIndex) = rst!CardID
            End If
            
         ElseIf ((rst!sector2 = 1 And getCruiserSector() = SectorID) Or (rst!sector2 = 2 And getCorvetteSector() = SectorID) Or (SectorID = rst!sector2)) And (rst!JobStatus = 1 Or rst!JobStatus = 2) Then 'Job2 must be in the sector
            If hasJobReqs(player.ID, rst!CardID, rst!Job2) Then
               .cbo.AddItem rst!Jobdes2 & " (" & CStr(rst!CardID) & ")"
               .cbo.ItemData(.cbo.NewIndex) = rst!CardID
            End If
            
         End If
         rst.MoveNext
      Wend
      rst.Close
      'Make Work if at a Planet
      If Nz(varDLookup("PlanetID", "Planet", "SectorID=" & SectorID), 63) <> 63 And Nz(varDLookup("PlanetID", "Planet", "SectorID=" & SectorID), 64) <> 64 Then 'but not Cruiser/Corvette dummy planetID 63,64
         .cbo.AddItem "Make Work at " & varDLookup("PlanetName", "Planet", "SectorID=" & SectorID)
         .cbo.ItemData(.cbo.NewIndex) = 0
      End If

         
      '>>>>>  CMD  WORK  <<<<
      .cmd(4).Enabled = (.cbo.ListCount > 0) And (actionSeq = ASselect) And (Not .workdone) And (Not onlyFullburn) And Not reaverActive
      'may as well show the first one
      If .cbo.ListCount > 0 Then .cbo.ListIndex = 0
         
      
      '>>>>>  CMD  END TURN  <<<<
      .cmd(5).Enabled = Not (actionSeq = ASDealSelDiscard Or actionSeq = ASDealSelect Or actionSeq = ASBuySelDiscard Or actionSeq = ASBuySelect) And Not reaverActive
      
      '>>>>>> remove Disgruntled <<<<<
      .cmd(6).Enabled = (getPerkAttributeCrew(player.ID, "RemoveDisgruntled") > 0 Or hasGear(player.ID, 27)) And hasDisgruntled(player.ID, True) And (Not .disgruntledone) And Not reaverActive
      .cmd(6).Visible = .cmd(6).Enabled
      
      '>>>>Resolve Alerts <<<<<<<<<<
      .cmd(7).Enabled = (hasAdjacentAlert(player.ID) And hasShipUpgrade(player.ID, 16) > 0 And (Not .fullburndone Or Not .moseydone))
      .cmd(7).Visible = .cmd(7).Enabled
      
      '>>>>>> load Passengers & Fugitives at Amnon's <<<<<
      .lblPassFugi.Visible = (SectorID = 23) And isSolid(player.ID, 1) And .cmd(3).Enabled And (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0.6)
      .txtPass.Visible = .lblPassFugi.Visible
      .txtFug.Visible = .lblPassFugi.Visible
      
      .FDPane1.PaneVisible = True
   End With
   If Not (frmShip Is Nothing) Then frmShip.RefreshShips
   If Not (frmJob Is Nothing) Then frmJob.RefreshJobs
   If Not (frmBuy Is Nothing) Then frmBuy.RefreshBuys
   
   Set rst = Nothing
End Sub

'returns doWork = 0 Normal, 1= Evade
Public Function doWork(ByVal playerID, ByVal CardID) As Integer
Dim rst As New ADODB.Recordset, x, parts As Integer, a() As String, DoubleDown As Integer
Dim SQL, SectorID, ContactID, JobID, finalstate, result As Integer, misbehaveNum, bonus, cargofit As Integer, fugifit As Integer, cargopay As Integer
Dim frmCrew As frmCrewLst, riverskill As Integer, dice As Integer, payment As Integer, KeywordInUse As Boolean
Dim skillcnt, skilldiscards, frmDiscardGr As frmDiscardGear, skillwin, skillint, payCrewTotal As Integer, Wskill As Integer, fruityBar As Integer
Dim frmSalvage As frmSalvaging, frmKillCrw As frmKillCrew

   SectorID = varDLookup("SectorID", "Players", "PlayerID=" & playerID)
   ContactID = varDLookup("ContactID", "ContactDeck", "CardID=" & CardID)
   usedStitchSkill = False
   DoubleDown = 1
   
   SQL = "SELECT ContactDeck.CardID, ContactDeck.Bonus, Job.JobID AS JOB1, Job.SectorID AS SECTOR1, Job_1.JobID AS JOB2, Job_1.SectorID AS SECTOR2, "
   SQL = SQL & "PlayerJobs.JobStatus, ContactDeck.Immoral, ContactDeck.JobName, Job_2.JobID AS JOB3, Job_2.SectorID AS SECTOR3 "
   SQL = SQL & "FROM (Job INNER JOIN ((PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) LEFT JOIN "
   SQL = SQL & "Job AS Job_1 ON ContactDeck.Job2ID = Job_1.JobID) ON Job.JobID = ContactDeck.Job1ID) LEFT JOIN Job AS Job_2 ON ContactDeck.Job3ID = Job_2.JobID "
   SQL = SQL & "WHERE PlayerJobs.PlayerID= " & playerID & " AND ContactDeck.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      If rst!JobStatus = 0 And ((rst!sector1 = 1 And getCruiserSector() = SectorID) Or (rst!sector1 = 2 And getCorvetteSector() = SectorID) Or (SectorID = rst!sector1)) Then ' we're doing Job 1
      
         JobID = rst!Job1
         If IsNull(rst!Job2) Then
            finalstate = JOB_SUCCESS
         Else
            finalstate = 1
         End If
         PutMsg player.PlayName & " Started Job: " & rst!JobName, playerID, Logic!Gamecntr
         
      ElseIf rst!JobStatus = 1 And ((rst!Sector3 = 1 And getCruiserSector() = SectorID) Or (rst!Sector3 = 2 And getCorvetteSector() = SectorID) Or (SectorID = rst!Sector3)) And Not IsNull(rst!Job3) Then ' we're doing Bonus Job
         JobID = rst!Job3
         bonus = rst!bonus
         finalstate = 2
         
      ElseIf (rst!JobStatus = 1 Or rst!JobStatus = 2) And ((rst!sector2 = 1 And getCruiserSector() = SectorID) Or (rst!sector2 = 2 And getCorvetteSector() = SectorID) Or (SectorID = rst!sector2)) And Not IsNull(rst!Job2) Then  ' we're doing Job 2
         JobID = rst!Job2
         finalstate = JOB_SUCCESS
         
      Else
         MsgBox "Job Card " & CardID & " Error for Player " & playerID, vbCritical
         Exit Function
      End If
   Else
      MsgBox "Job Card " & CardID & " Error for Player " & playerID, vbCritical
      Exit Function
   End If
   
   If finalstate = JOB_SUCCESS And rst!Immoral = 1 And hasDisgruntled(playerID) Then
      If MessBox("Warning: Completing this Immoral Job with Disgruntle Crew will result in Crew leaving." & vbNewLine & "Are you sure you want to continue?", "Immoral Job Consequences", "Yes", "No", getLeader()) = 1 Then
         frmAction.workdone = False
         Exit Function
      End If
   End If
   'check for Shepherd  (93 / 47) and keep him off immoral jobs
   If rst!Immoral = 1 And hasCrew(playerID, 47) Then
      DB.Execute "UPDATE PlayerSupplies SET OffJob = 1 WHERE CardID = 93"
      If Not (frmShip Is Nothing) Then frmShip.RefreshShips  'update display
      PutMsg player.PlayName & "'s Work log: Shepherd's having none of this Immoral Job citing '..a special place in Hell!'", playerID, Logic!Gamecntr, True, 47
   End If
   rst.Close
   
   'assume we now have a JobID to carry out in the current Sector
   SQL = "SELECT * FROM Job WHERE JobID = " & JobID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
   
      If rst!misbehave <> 0 Then
         misbehaveNum = rst!misbehave
         'if TwoFry has a Sniper Rifle, then one less misbv
         If hasGearKeyword(playerID, "SNIPERRIFLE", 56) And misbehaveNum > 1 Then
            misbehaveNum = misbehaveNum - 1
            PutMsg player.PlayName & "'s Two-Fry used his DeadEye Sniper skills to good effect and eliminates 1 misbehave", playerID, Logic!Gamecntr, True, 56
         End If
         'go do the number of misbehaves
         result = doMisbehaves(playerID, misbehaveNum, SectorID)
         Select Case result
         Case 1, 4 'proceed
            result = 0 'reset as Win for below tests
         Case 2 'botched
            PutMsg player.PlayName & "'s Work log: Job was Botched!", playerID, Logic!Gamecntr, True, getLeader()
            Exit Function
            
         Case 3 'discard Job and clear solid, with Contact
            'and if Niska - Kill 1 Crew
            'add a Warrant and clear any Solid with Harken (5)
            doJobWarrant playerID, ContactID, CardID
            frmJob.RefreshJobs
            Main.drawLine 0, -1
            Exit Function
            
         Case 5 'double down success
            result = 0 'reset as Win for below tests
            If rst!DoubleDown > 0 Then
               DoubleDown = 2
               'payment = payment + rst!DoubleDown
               PutMsg player.PlayName & " scored the Double-Down bonus", playerID, Logic!Gamecntr
            End If
         
         End Select
      End If
   
      'pickup/drop off Cargo, etc
      If rst!cargo <> 0 Then
         If (rst!cargo * -1) > varDLookup("Cargo", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
            MessBox "Not enough cargo to meet the quota, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
         
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= rst!cargo Then 'we have room
            DB.Execute "UPDATE Players SET Cargo = Cargo + " & rst!cargo & " WHERE PlayerID = " & playerID
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      
      If rst!Contraband <> 0 Then
         If rst!Contraband = 14 Then
            If Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID)) > 0 Then
               cargofit = Val(InputBox("How much Contraband do you want to load onboard?", "Load Contraband", CStr(Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID)))))
            End If
         ElseIf rst!Contraband = -14 Then
            cargofit = Val(InputBox("How much Contraband do you want to deliver?", "Deliver Contraband", varDLookup("Contraband", "Players", "PlayerID=" & playerID)))
            If cargofit > varDLookup("Contraband", "Players", "PlayerID=" & playerID) Then
               cargofit = varDLookup("Contraband", "Players", "PlayerID=" & playerID)
            End If
            cargopay = cargofit * 500
            cargofit = cargofit * -1
         Else
            cargofit = rst!Contraband
            If (cargofit * -1) > varDLookup("Contraband", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
               MessBox "Not enough Contraband to meet the quota, Job botched", "Job Requirements", "Ooops", "", getLeader()
               Exit Function
            End If
         End If
         
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= cargofit Then 'we have room
            DB.Execute "UPDATE Players SET Contraband = Contraband + " & cargofit & " WHERE PlayerID = " & playerID
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job botched.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      
      End If
      
      If rst!Passenger <> 0 Then
         If rst!Passenger = 14 Then
            If Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID)) > 0 Then
               cargofit = Val(InputBox("How many Passengers do you want to take onboard?", "Load Passengers", CStr(Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID)))))
            End If
         ElseIf rst!Passenger = -14 Then
            cargofit = Val(InputBox("How many Passengers do you want to deliver?", "Deliver Passengers", varDLookup("Passenger", "Players", "PlayerID=" & playerID)))
            If cargofit > varDLookup("Passenger", "Players", "PlayerID=" & playerID) Then
               cargofit = varDLookup("Passenger", "Players", "PlayerID=" & playerID)
            End If
            cargopay = cargofit * IIf(rst!Fugitive = -14, 300, 200) 'chk being sold as Fugitives
            cargofit = cargofit * -1
            
         Else
            cargofit = rst!Passenger
            If (cargofit * -1) > varDLookup("Passenger", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
               MessBox "Not enough Passengers to meet the quota, Job botched", "Job Requirements", "Ooops", "", getLeader()
               Exit Function
            End If
         End If
                      
         'pay at the end, passing cargofit
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= cargofit Then 'we have room
            SQL = "UPDATE Players SET Passenger = Passenger + " & cargofit
            SQL = SQL & " WHERE PlayerID = " & playerID
            DB.Execute SQL
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, Job botched", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      If rst!Fugitive <> 0 Then
         If rst!Fugitive = 14 Then
         
            If Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID)) > 0 Then
               fugifit = Val(InputBox("How many Fugitives do you want to take onboard?", "Load Fugitives", CStr(Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID)))))
            End If

         ElseIf rst!Fugitive = -14 Then
            fugifit = Val(InputBox("How many Fugitives do you want to deliver?", "Deliver Fugitives", varDLookup("Fugitive", "Players", "PlayerID=" & playerID)))
            If fugifit > varDLookup("Fugitive", "Players", "PlayerID=" & playerID) Then
               fugifit = varDLookup("Fugitive", "Players", "PlayerID=" & playerID)
            End If
            cargopay = cargopay + fugifit * 300  'may get paid for pasngrs too
            fugifit = fugifit * -1
            
            
         Else
            fugifit = rst!Fugitive
            If (fugifit * -1) > varDLookup("Fugitive", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
               MessBox "Not enough Fugitives to meet the quota, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
               Exit Function
            End If
         End If
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= fugifit Then 'we have room
            DB.Execute "UPDATE Players SET Fugitive = Fugitive + " & fugifit & " WHERE PlayerID = " & playerID
            If fugifit < 0 Then
               If beaDirtySlaver(playerID) Then
                  cargopay = cargopay + Abs(fugifit) * 100
               End If
            End If
            
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      If rst!fuel <> 0 Then
         If (rst!fuel * -1) > varDLookup("Fuel", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
            MessBox "Not enough Fuel to meet the quota, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= (rst!fuel / 2) Then 'we have room
            DB.Execute "UPDATE Players SET Fuel = Fuel + " & rst!fuel & " WHERE PlayerID = " & playerID
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      If rst!parts <> 0 Then
         If (rst!parts * -1) > varDLookup("Parts", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
            MessBox "Not enough Parts to meet the quota, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
         
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= (rst!parts / 2) Then 'we have room
            DB.Execute "UPDATE Players SET Parts = Parts + " & rst!parts & " WHERE PlayerID = " & playerID
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      'TAG and BAG
      If rst!tagnbag > 0 And CargoCapacity(playerID) > CargoSpaceUsed(playerID) Then 'load to your capacity
         If rst!tagnbag = 1 Then
            skillcnt = getSkill(playerID, cstrSkill(2), 0, False) + RollDice(6)
            PutMsg player.PlayName & " Tech Test comes to " & skillcnt & " for the Tag and Bag", playerID, Logic!Gamecntr
            Select Case skillcnt
            Case 1 - 4
               x = 3
            Case 5 - 7
               x = 6
            Case Else
               x = 20
            End Select
         ElseIf rst!tagnbag = 20 Then
            x = 20
            PutMsg player.PlayName & " does a Tag and Bag to grab some goods", playerID, Logic!Gamecntr
         Else
            x = rst!tagnbag
         End If
         If frmSalvage Is Nothing Then
            Set frmSalvage = New frmSalvaging
         End If
         frmSalvage.mode = 2
         frmSalvage.salvageCount = x
         frmSalvage.Show 1
         
      End If
      
      
   End If
   
   rst.Close
   
   If finalstate = 1 Then

      PutMsg player.PlayName & " has completed the first Work Part of " & varDLookup("JobName", "ContactDeck", "CardID=" & CardID) & " at " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & SectorID)), playerID, Logic!Gamecntr, True, getLeader()
      
   ElseIf finalstate = 2 Then 'Bonus Job done
      'Pay Bonus - todo
      bonus = bonus + cargopay
      getMoney player.ID, bonus
      PutMsg player.PlayName & " has completed the $" & bonus & " Bonus Work Part of " & varDLookup("JobName", "ContactDeck", "CardID=" & CardID) & " at " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & SectorID)), playerID, Logic!Gamecntr, True, getLeader()

   ElseIf finalstate = JOB_SUCCESS Then 'job is ending, but do any remaining challenges Negotiate Pay or Cover your Tracks
      
      SQL = "SELECT * FROM ContactDeck WHERE CardID=" & CardID
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         If rst!RemoveDisgruntled <> 0 Then
            doDisgruntled player.ID, 3
         End If
         'KEYWORD CHECKS - ========================
         If rst!WinOptKeyword > 0 Then ' give option to use Keyword to Win -  only applies to HACKINGRIG or Explosives (getting paid half)
               
            'check if the keyword was single use, and discard
            
            If discardGearKeyword(playerID, rst!KeyWords, True) Then
               If MessBox("In this final Work Challenge, do you want to use your discardable " & rst!KeyWords & " instead of the Skill Test?" & IIf(rst!KeyWords = "EXPLOSIVES", vbNewLine & "This would result in Half Pay.", ""), "Final Work Challenge", "Yes", "No", getLeader()) = 0 Then
                  discardGearKeyword playerID, rst!KeyWords
                  KeywordInUse = True
                  If rst!KeyWords = "EXPLOSIVES" Then
                     result = 2 'half pay
                  Else
                     result = 0
                  End If
               End If
            Else
               If MessBox("In this final Work Challenge, do you want to use your " & rst!KeyWords & " instead of the Skill Test?" & IIf(rst!KeyWords = "EXPLOSIVES", vbNewLine & "This would result in Half Pay.", ""), "Final Work Challenge", "Yes", "No", getLeader()) = 0 Then
                  KeywordInUse = True
                  If rst!KeyWords = "EXPLOSIVES" Then
                     result = 2 'half pay
                  Else
                     result = 0
                  End If
               End If
            End If
             
         ElseIf rst!WinOptKeyword = 0 And rst!KeywordBonus = 0 And Not IsNull(rst!KeyWords) And Not (rst!KeywordOrSkill > 0 And hasCrewAttribute(playerID, cstrProfession(rst!RequireProfession))) Then   'check for discard
            a = Split(rst!KeyWords, " ")
            For x = LBound(a) To UBound(a)
               If discardGearKeyword(playerID, a(x), True) Then
                  MessBox "Discarding spent " & a(x), "Job Required Gear Keyword", "OK", "", getLeader()
                  discardGearKeyword playerID, a(x)
               End If
            Next x
            
         End If
         Wskill = rst!skill
         If Wskill > 0 And Not KeywordInUse Then 'we have a skill test
         
            '-----------------------------------------
            'Stitch & Sheydra can change a Fight to a Nego once per Job
            If Wskill = 3 And hasCrew(playerID, 27) And Not usedStitchSkill Then
               If MessBox("Stitch wants to change this Negotiation to a Fight.  Do you want to use those skills instead?", "Negotiate -> Fight", "Yes", "No", 27) = 0 Then
                  Wskill = 1
                  usedStitchSkill = True
                  PutMsg player.PlayName & " uses Stitch's one time Negotiation to Fight Skills", playerID, Logic!Gamecntr, True, 27
               End If
            End If
            If Wskill = 1 And getPerkAttributeCrew(playerID, "ChangeTestType") > 0 And Not usedStitchSkill Then
               If MessBox("Sheydra wants to Negotiate instead of Fight.  Do you want to use those skills instead?", "Fight -> Negotiate", "Yes", "No", 66) = 0 Then
                  Wskill = 3
                  usedStitchSkill = True
                  PutMsg player.PlayName & " uses Sheydra's one time Fight to Negotiation Skills", playerID, Logic!Gamecntr, True, 66
               End If
            End If
            
            'Crazy River Tam (cardID 51/CrewID 32)
            If hasCrew(playerID, 32) Then
               dice = RollDice(6)
               If hasCrew(playerID, 33) Then  'simon adds 2 to her rolls
                  dice = dice + 2
               End If
               Select Case dice
               Case 1, 2 'stay onboard
                  DB.Execute "UPDATE PlayerSupplies SET OffJob = 1 WHERE CardID = 51"
                  If Not (frmShip Is Nothing) Then frmShip.RefreshShips  'update display
                  PutMsg player.PlayName & "'s River Tam cowers onboard and won't be workin' anymore today", playerID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, dice
               Case 3 'fight
                  If Wskill = 1 Then
                     riverskill = 2
                  End If
               Case 4 'Tech
                  If Wskill = 2 Then
                     riverskill = 2
                  End If
               Case 5 'negot
                  If Wskill = 3 Then
                     riverskill = 2
                  End If
               Case Else 'any skill
                     riverskill = 2
               End Select
               If riverskill = 2 Then
                  PutMsg player.PlayName & "'s River Tam" & IIf(hasCrew(playerID, 33), ", encouraged by Simon,", "") & " channels the " & cstrSkill(Wskill) & " skill + 2", playerID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, dice
               ElseIf dice > 2 Then
                  PutMsg player.PlayName & "'s River Tam ain't workin' this time", playerID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, dice
               End If
            End If
            
            fruityBar = hasGearCard(playerID, 24)
            If fruityBar > 0 Then 'we got one or more
               If MessBox("Do you wish to Eat the Fruity Bar and add 1 to the Test Roll?", "Extra Bite", "Yes", "No", 0, 24) = 0 Then
                  doDiscardGear playerID, fruityBar
                  fruityBar = 1
               Else
                  fruityBar = 0
               End If
            End If
            
             x = hasGearCrew(playerID, 28) 'Mal's Brown Coat
            If x > 0 And varDLookup("Disgruntled", "Crew", "CrewID=" & x) > 0 And varDLookup("Fight", "Crew", "CrewID=" & x) > 0 And Wskill = 3 Then
               fruityBar = fruityBar + varDLookup("Fight", "Crew", "CrewID=" & x)
               PutMsg player.PlayName & "'s Disgruntled Crew wearing the Brown Coat adds their Fight skills to the Negotiation", playerID, Logic!Gamecntr, True, 0, 28
            End If
            
            If Wskill = 1 Then
               removeDigruntled playerID, Wskill
            End If
            
            '<<<<<<<<<<<<<< ROLL THE DICE >>>>>>>>>>>>>>>>>>>>>>>>>
            dice = RollDice(6, IIf(Wskill = 3 And hasCrew(playerID, 55), False, True))
            
            If Wskill = 1 And hasGear(player.ID, 47) Then ' Zoe's Mare's Leg Rifle -When making a Fight Test, roll two dice and use the highest.
               x = RollDice(6, IIf(rst!skill = 3 And hasCrew(player.ID, 55), False, True))
               If x > dice Then
                  PutMsg player.PlayName & " had rolled a " & CStr(dice) & " so using Zoe's Mare's Leg Rifle rerolled a " & CStr(x), player.ID, Logic!Gamecntr, True, 0, 47, 0, 0, 0, x
                  dice = x
               End If
            End If
            
            If dice = 1 Then  'reroll ones?
               If hasGear(player.ID, 35) And Wskill = 1 Then 'Inara's Bow
                  x = hasGearCrew(player.ID, 35)
                  If x > 0 Then
                     If hasCrewAttribute(player.ID, "Companion", 0, x) Then
                        Do While dice = 1
                           dice = RollDice(6, True)
                        Loop
                        PutMsg player.PlayName & "'s Companion uses Inara's Bow to reRoll a 1 and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, 0, 35, 0, 0, 0, dice
                     End If
                  End If
               End If
            End If
            
            'Inara & Kaylee can reroll a negotiate test
            If Wskill = 3 And getPerkAttributeCrew(playerID, "RerollNegotiate") And dice < 6 Then
               If MessBox("You rolled a " & dice & vbNewLine & "Your Negotiation Skills allow you a second chance, do you want to take that chance?", "Re-Roll option", "Yes", "No", getLeader(), 0, 0, dice) = 0 Then
                  dice = RollDice(6, IIf(Wskill = 3 And hasCrew(playerID, 55), False, True))
                  PutMsg player.PlayName & " uses extra Negotiation Skills to reRoll and got a " & CStr(dice), playerID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
               End If
            End If
            'Zoe's skill
            If Wskill = 1 And getPerkAttributeCrew(playerID, "RerollFight") And dice < 6 Then
               If MessBox("You rolled a " & dice & vbNewLine & "Your Fight Skills allow you a second chance, do you want to take that extra chance?", "Re-Roll option", "Yes", "No", getLeader(), 0, 0, dice) = 0 Then
                  dice = RollDice(6, IIf(Wskill = 3 And hasCrew(playerID, 55), False, True))
                  PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(dice), playerID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
               End If
            End If
            
            If Wskill = 1 And hasGear(player.ID, 45) And dice < 6 Then 'yolanda's pistol - Discard to re-roll a Fight Test.
               If MessBox("You rolled a " & dice & vbNewLine & "Yolanda's pistol allows you a second chance, do you want to Discard the Pistol to take that extra chance?", "Re-Roll option", "Yes", "No", 0, 45, 0, dice) = 0 Then
                  doDiscardGear player.ID, hasGearCard(player.ID, 45)
                  dice = RollDice(6, IIf(Wskill = 3 And hasCrew(playerID, 55), False, True))
                  PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(dice), playerID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
               End If
            End If
            
            If Wskill = 1 And hasGear(player.ID, 48) And dice < 6 Then 'Extra Ammo Clip - Discard to re-roll a Fight Test.
               If MessBox("You rolled a " & dice & vbNewLine & "Extra Ammo Clips allow you a second chance, do you want to Discard the Clips to take that extra chance?", "Re-Roll option", "Yes", "No", 0, 48, 0, dice) = 0 Then
               'If MsgBox("You rolled a " & dice & vbNewLine & "Extra Ammo Clips allow you a second chance, do you want to Discard the Clips to take that extra chance?", vbYesNo + vbQuestion, "Re-Roll option") = vbYes Then
                  doDiscardGear player.ID, hasGearCard(player.ID, 48)
                  dice = RollDice(6, IIf(Wskill = 3 And hasCrew(player.ID, 55), False, True))
                  PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
               End If
            End If
            
            '----------------------------------------- see if we need to use the discardable skills & keywords...
            
           
            skillwin = rst!Win
            skillint = rst!Intermediate
   
            'get our skill totals, no Kosherized rules in play for Jobs, only MB
            skillcnt = getSkill(playerID, cstrSkill(Wskill), 0, True) + dice + riverskill + fruityBar
            skilldiscards = getSkillDiscards(playerID, cstrSkill(Wskill))
            

               
            '-----------------------------------------
            If skillcnt < skillwin And skillcnt + skilldiscards >= skillwin Then 'we're in trouble 'we could use some help
               If MessBox("With the help of " & skillwin - skillcnt & " skill points, we can succeed" & vbNewLine & "Do you want to use a discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
                  'show a list of gear to pick from up to or exceeding the value skillwin - skillcnt
                  Set frmDiscardGr = New frmDiscardGear
                  frmDiscardGr.nbrSelect = skillwin - skillcnt
                  frmDiscardGr.skill = cstrSkill(Wskill)
                  frmDiscardGr.Caption = "Select single use Gear to provide at least " & CStr(frmDiscardGr.nbrSelect) & " skill points"
                  frmDiscardGr.Show 1
                  'then add selected skill points to skillcnt, discard gear, and go on...
                  skillcnt = skillcnt + frmDiscardGr.nbrSelected
               End If
            ElseIf skillcnt < skillint And skillcnt + skilldiscards >= skillint Then 'we're in trouble 'we could use some help
               If MessBox("With the help of " & skillint - skillcnt & " skill points, we can make the intermediate outcome" & vbNewLine & "Do you want to use discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
                  'show a list of gear to pick from up to or exceeding the value skillint - skillcnt
                  Set frmDiscardGr = New frmDiscardGear
                  frmDiscardGr.nbrSelect = skillint - skillcnt
                  frmDiscardGr.skill = cstrSkill(Wskill)
                  frmDiscardGr.Caption = "Select single use Gear to provide at least " & CStr(frmDiscardGr.nbrSelect) & " skill points"
                  frmDiscardGr.Show 1
                  'then add selected skill points to skillcnt, discard gear, and go on...
                  skillcnt = skillcnt + frmDiscardGr.nbrSelected
               End If
            End If
            '-----------------------------------------
            
            If hasGear(playerID, 32) And Wskill = 1 And skillcnt < rst!Win Then '  use Simon's Sonic Stun Baton??
               If MessBox("The Fights not going so well with a skill score of " & skillcnt & vbNewLine & "Simon's Sonic Stun Baton might turn things around, wanna try another Thrillin' Heroics Roll and Discard the Baton?", "Stun Baton to the Fight", "Yes", "No", 0, 32) = 0 Then
                  dice = RollDice(6) + 6
                  skillcnt = getSkill(playerID, cstrSkill(Wskill), 0, True) + dice + riverskill + fruityBar
                  doDiscardGear playerID, hasGearCard(playerID, 32)
               End If
            End If
            
            If skillcnt >= rst!Win Then
               result = 0
            ElseIf skillcnt >= rst!Intermediate And rst!Intermediate > 0 Then
               result = 1
            Else 'you lose :(
               result = 3
            End If
            PutMsg player.PlayName & "'s Work log: Rolls a " & dice & " with added " & cstrSkill(Wskill) & " skill points to a total of " & skillcnt & " to " & IIf(result = 0, "succeed :^)", IIf(result = 1, "part win", "lose :^(")), playerID, Logic!Gamecntr, True, getLeader(), getLeader(), 0, 0, 0, dice
            
         End If  'end of the skill Tests

         Select Case result
         Case 0 'win results
            Select Case rst!WinResult
            Case 1 'passngr  - now done above using job's passngr count -14
            
            Case 2 'fugi - now done above using job's fugi count -14
               
            Case 3 ' 1 cargo per Crew on Job
               skillcnt = Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID))
               'limit to what we can fit, or the number of crew on job
               If skillcnt >= getCrewCount(playerID, True) Then skillcnt = getCrewCount(playerID, True)
               If skillcnt <> 0 Then
                  DB.Execute "UPDATE Players Set Cargo = Cargo + " & skillcnt & " WHERE PlayerID = " & playerID
                  PutMsg player.PlayName & IIf(skillcnt > 0, " scored ", " lost ") & skillcnt & " Cargo", playerID, Logic!Gamecntr
               End If
               
            Case 4 'move Cruiser to sector and EVADE - work done
               MoveShip 5, SectorID
               If getHaven(SectorID) > 0 Then
                  PutMsg player.PlayName & "'s Nav log: refuge found at this Haven, the Alliance Cruiser sails on by", player.ID, Logic!Gamecntr, True, 0, 0, 1
                  moveAutoAI 5
               Else
                  PutMsg player.PlayName & " needs to EVADE!", playerID, Logic!Gamecntr, True
                  actionSeq = ASNavEvade
                  doWork = 1 ' Evade
               End If
               
               
            Case 5 ' 1 contraband per Crew on Job
               skillcnt = Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID))
               'limit to what we can fit, or the number of crew on job
               If skillcnt >= getCrewCount(playerID, True) Then skillcnt = getCrewCount(playerID, True)
               If skillcnt <> 0 Then
                  DB.Execute "UPDATE Players Set Contraband = Contraband + " & skillcnt & " WHERE PlayerID = " & playerID
                  PutMsg player.PlayName & IIf(skillcnt > 0, " scored ", " lost ") & skillcnt & " Contraband", playerID, Logic!Gamecntr
               End If
               
            Case 6 'kill a merc
               Set frmKillCrw = New frmKillCrew
         
               frmKillCrw.nbrSelect = 1
               frmKillCrw.extrafilter = " AND Crew.Merc = 1"
               frmKillCrw.Show 1
               Set frmKillCrw = Nothing
               
            Case 7 ' EVADE - work done
               PutMsg player.PlayName & " needs to EVADE!", playerID, Logic!Gamecntr, True
               actionSeq = ASNavEvade
               doWork = 1 ' Evade
               
            Case Is > 99 'paid extra
               payment = payment + rst!WinResult
            
            Case Is < 0 'make payment

               If getMoney(playerID) >= Abs(rst!WinResult) Then
                  payment = payment + rst!WinResult
               Else
                  PutMsg player.PlayName & " doesn't have enough Money to pay the $" & CStr(Abs(rst!WinResult)) & " fee, Job botched.", playerID, Logic!Gamecntr, True, getLeader()
                  Exit Function 'fail
               End If

            End Select
            
            
         Case 1 'inter
            'no change
            If rst!IntermediateResult < 99 Then
               If getMoney(playerID) >= (rst!IntermediateResult * -1) Then
                  payment = payment + rst!IntermediateResult
               Else
                  PutMsg player.PlayName & " doesn't have enough Money to pay the $" & CStr(Abs(rst!IntermediateResult)) & " fee, Job botched.", playerID, Logic!Gamecntr, True, getLeader()
                  Exit Function 'fail
               End If
            End If
   
         Case 2 'half pay
            'handled below
            
         Case 3 ' lose results -0 = continue
            Select Case rst!FailResult
            Case 1  'lose rep only
               If rst!FailLoseRep > 0 Then
                  If Not discardRoberta(playerID) Then
                     DB.Execute "UPDATE Players SET Solid" & rst!FailLoseRep & "=0 WHERE PlayerID =" & playerID
                     PutMsg player.PlayName & " loses any Rep with " & varDLookup("ContactName", "Contact", "ContactID=" & rst!FailLoseRep), playerID, Logic!Gamecntr, True, 0, 0, 0, rst!FailLoseRep
                  End If
               End If
               
            Case 2 'warrant issued - attempt botched
               doJobWarrant playerID, ContactID, CardID
               Main.drawLine 0, -1
               Exit Function
                              
            Case 3 ' pay 1000 attempt botched
               If getMoney(playerID) >= 1000 Then
                  getMoney playerID, -1000
                  PutMsg player.PlayName & " botches the final negotiation and the job, and loses $1000", playerID, Logic!Gamecntr, True, getLeader()
                  Exit Function 'fail
               Else 'take it all
                  getMoney playerID, -1 * getMoney(playerID)
                  PutMsg player.PlayName & " botches the final negotiation and the job, and loses all their money", playerID, Logic!Gamecntr, True, getLeader()
                  Exit Function 'fail
               End If
               
             Case 4 'attempt botched
               PutMsg player.PlayName & " botches the final negotiation and the job", playerID, Logic!Gamecntr, True, getLeader()
               Exit Function
               
            Case Is < -99
               If getMoney(playerID) + payment < Abs(rst!FailResult) Then
                  payment = getMoney(playerID) * -1
               Else
                  payment = payment + rst!FailResult
               End If
            End Select
            
            'ignore for Niska + Warrant, as this already applies
            If rst!FailKillCrew > 0 And Not (ContactID = 3 And rst!FailResult = 2) Then
               doKillCrews playerID, rst!FailKillCrew, True
            End If
   
         End Select
      
         '-----------------------------------------
         'if complete - finish up
         If rst!Immoral = 1 Then
            doDisgruntled player.ID, 1
            PutMsg player.PlayName & " does an immoral Job and any Moral Crew will be disgruntled", playerID, Logic!Gamecntr
         End If
   
         If ContactID = 0 Then  ' a Goal job
            payCrewTotal = 0
         ElseIf getCrewCount(playerID) > 1 Then 'someone to pay?
            'show a list of Crew to choose who gets paid, then return deduct amt
            Set frmCrew = New frmCrewLst
            frmCrew.noMoralDisgruntle = (rst!GoodDeeds = 1)
            frmCrew.Label1 = "Job Pay: " & "$" & rst!pay & "  " & IIf(rst!BonusPart > 0, " +" & rst!BonusPart & " part: ", "") & IIf(rst!bonus > 0, " +$" & rst!bonus & ":", "") & _
               IIf(rst!KeywordBonus = 1, rst!KeyWords, "") & IIf(rst!ProfessionID = 0, "", " " & cstrProfession(rst!ProfessionID)) & IIf(rst!BonusPerSkill > 0, " /" & cstrSkill(rst!BonusPerSkill), "") & _
               IIf(rst!Job3ID > 0, "Bonus Job", "") & IIf(payment > 0, "  plus $" & CStr(payment), "") & IIf(payment < 0, "  minus $" & CStr(Abs(payment)), "")

            frmCrew.Show 1
            payCrewTotal = frmCrew.payTotal
         End If
         'Get Paid, with any Bonus, less deductions & Go Solid with you Contact
         SQL = "UPDATE Players SET "
         bonus = getJobBonus(playerID, CardID, parts)
         If parts > 0 Then
            SQL = SQL & "Parts= Parts + " & CStr(parts) & ", "
         End If
         'final pay with Leader Perk Bonus, and on the job profession bonus added, less crew hire
         'if result = 1 Then 'half pay
         bonus = (rst!pay * IIf(result = 2, 0.5, 1) * DoubleDown) + getJobCrewBonus(playerID, rst!JobTypeID) + getJobCrewBonus(playerID, rst!JobType2D) + bonus + payment + cargopay - payCrewTotal
         SQL = SQL & "Pay = Pay + " & bonus
         If ContactID > 0 Then
            SQL = SQL & ", Solid" & ContactID & "=1 "  'setting SOLID with the Contact
         End If
         SQL = SQL & " Where playerID = " & playerID
         DB.Execute SQL
         
         If hasGear(playerID, 31) And (rst!JobTypeID = 1 Or rst!JobType2D = 1) Then 'MF-813 Flying Mule After completing a Crime Job, Load 6 Goods, minus 1 per Crew Working the Job.
            x = getCrewCount(playerID, True)
            If x < 6 Then
               If frmSalvage Is Nothing Then
                  Set frmSalvage = New frmSalvaging
               End If
               frmSalvage.mode = 2
               frmSalvage.salvageCount = (6 - x)
               frmSalvage.Show 1
               PutMsg player.PlayName & " uses the MF-813 Flying Mule to grab some goods (" & CStr(6 - x) & ")", playerID, Logic!Gamecntr, 0, 31
            End If
         End If
         
         'do this last as Cargo may have changed.
         If hasShipUpgrade(playerID, 21) And (rst!JobTypeID = 1 Or rst!JobType2D = 1) Then
            PutMsg player.PlayName & "'s Hydraulic Docking Clamps can grab Salvage for this Crime Job", playerID, Logic!Gamecntr
            doSalvage player.ID
         ElseIf (rst!JobTypeID = 6 Or rst!JobType2D = 6) Then ' SalvageOps (+ Hydraulic Docking Clamps/Crime)
            doSalvage player.ID
         End If
         
         PutMsg player.PlayName & " completed the Job: " & varDLookup("JobName", "ContactDeck", "CardID=" & CardID) & " for $" & Abs(bonus) & IIf(bonus > 0, " profit", " loss") & IIf(ContactID = 0, "", " and is solid with " & varDLookup("ContactName", "Contact", "ContactID=" & rst!ContactID)), playerID, Logic!Gamecntr, True, 0, 0, 0, rst!ContactID
      
      End If 'close off record
      rst.Close
      
   End If 'end of finalstate results check 1/2/3

   'update the status of the job
   DB.Execute "UPDATE PlayerJobs SET JobStatus =" & finalstate & " WHERE PlayerID = " & playerID & " AND CardID = " & CardID
   frmJob.RefreshJobs
   Main.drawLine 0, -1

   'if we got this far, we're good!
   'doWork = 0 only set this to a value above to change the exit actionSeq behavior.  1=EVADE
   
   Set rst = Nothing
End Function

'opt 1,2, 3-ace in hole
Private Function getMisbehave(opt, cnt, total, suit) As Integer
Dim SQL, reshuffle
Dim rst As New ADODB.Recordset
Dim frmMB As New frmMisbehave

   With frmMB

      .MBCardID = 0
      .MBOption = 0

      'Read in the next NAV card and display either 1 or 2 options
      SQL = "SELECT MisbehaveDeck.CardID, MisbehaveDeck.CardName, MisbehaveDeck.Reshuffle, MisbehaveDeck.Seq, MisbehaveDeck.Suit, "
      SQL = SQL & "MisbehaveDeck.Keyword AS Keywords , MisbehaveDeck.CrewID, MisbehaveDeck.GearID, MisbehaveDeck.ProfessionID AS ProfesionID, MisOption.* "
      SQL = SQL & "FROM MisOption INNER JOIN MisbehaveDeck ON MisOption.OptionID = MisbehaveDeck.Option1ID "
      SQL = SQL & "Where MisbehaveDeck.Seq > 5 "
      SQL = SQL & "ORDER BY MisbehaveDeck.Seq"
       
      rst.Open SQL, DB, adOpenDynamic, adLockOptimistic
      If Not rst.EOF Then
         suit = rst!suit
         'pull the card out of the deck
         rst!Seq = 5
         rst.Update
         reshuffle = rst!reshuffle
         If reshuffle = 1 Then 'ready for next turn
            PutMsg player.PlayName & " Reshuffling MisbehaveDeck due to " & rst!CardName, player.ID, Logic!Gamecntr, True, getLeader()
            ShuffleDeck "Misbehave"
         End If
         
         'check for an ACE in the HOLE
         If Nz(rst!KeyWords, "") <> "" Then
            If hasKeyword(player.ID, rst!KeyWords) Then
               'check if the keyword was single use, and discard
               If discardGearKeyword(player.ID, rst!KeyWords, True) Then
                  If MessBox("Do you want to use your discardable " & rst!KeyWords & " for Ace-in-the-Hole to proceed?", "Ace in the Hole", "Yes", "No", getLeader()) = 0 Then
                     discardGearKeyword player.ID, rst!KeyWords
                     getMisbehave = rst!CardID
                     opt = 3
                     PutMsg player.PlayName & " misbhavin' with " & rst!CardName & " had an Ace in the Hole with " & rst!KeyWords, player.ID, Logic!Gamecntr, True, getLeader()
                     Exit Function
                  End If
               Else 'just your regular multi-use Keyword
                  getMisbehave = rst!CardID
                  opt = 3
                   PutMsg player.PlayName & " misbhavin' with " & rst!CardName & " had an Ace in the Hole with " & rst!KeyWords, player.ID, Logic!Gamecntr, True, getLeader()
                  Exit Function
               End If
            End If
         End If
         If rst!CrewID <> 0 Then
            If hasCrew(player.ID, rst!CrewID) Then
               getMisbehave = rst!CardID
               opt = 3
               PutMsg player.PlayName & " misbhavin' with " & rst!CardName & " had an Ace in the Hole with " & getCrewName(0, rst!CrewID), player.ID, Logic!Gamecntr, True, rst!CrewID
               Exit Function
            End If
         End If
         If rst!GearID <> 0 Then
             If hasGear(player.ID, rst!GearID) Then
               getMisbehave = rst!CardID
               opt = 3
               PutMsg player.PlayName & " misbhavin' with " & rst!CardName & " had an Ace in the Hole with a " & getGearName(0, rst!GearID), player.ID, Logic!Gamecntr, True, 0, rst!GearID
               Exit Function
            End If
         End If
         If rst!ProfesionID <> 0 Then
             If hasCrewAttribute(player.ID, cstrProfession(rst!ProfessionID)) Then
               getMisbehave = rst!CardID
               opt = 3
               PutMsg player.PlayName & " misbhavin' with " & rst!CardName & " had an Ace in the Hole with a " & cstrProfession(rst!ProfessionID), player.ID, Logic!Gamecntr, True, getLeader()
               Exit Function
            End If
         End If
         
         'other, soldier on..
         .MBCardID = rst!CardID
         .cmd(0).Enabled = True
         
         If Nz(rst!Keyword, "") <> "" Then
            If Not hasKeyword(player.ID, rst!Keyword) And rst!WinOptKeyword = 0 And rst!KeywordOrSkill = 0 Then
               .cmd(0).Enabled = False
            End If
         End If
         If rst!Disgruntled = -1 And hasDisgruntled(player.ID) Then
            .cmd(0).Enabled = False
         End If
         'if this option requires Cargo, check we have enough to honour it
         If rst!cargo < 0 And varDLookup("Cargo", "Players", "PlayerID=" & player.ID) < Abs(rst!cargo) Then
            .cmd(0).Enabled = False
         End If
               
         .lblName.Caption = rst!CardName
         .cmd(0).Caption = rst!OptionName
         .cmd(0).ToolTipText = rst!OptionName
         .lblDetail(0).Caption = rst!Details

      Else
         PutMsg player.PlayName & " Reshuffling MisbehaveDeck due to end of deck", player.ID, Logic!Gamecntr, True, getLeader()
         ShuffleDeck "Misbehave"
         Exit Function
      End If
      rst.Close
      
      '-- Option 2--------------------------------------------------------------
      SQL = "SELECT MisOption.* "
      SQL = SQL & "FROM MisOption INNER JOIN MisbehaveDeck ON MisOption.OptionID = MisbehaveDeck.Option2ID "
      SQL = SQL & "Where MisbehaveDeck.CardID = " & .MBCardID
      
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         .cmd(1).Visible = True
         .cmd(1).Enabled = True
         If Nz(rst!Keyword, "") <> "" Then
            If Not hasKeyword(player.ID, rst!Keyword) And rst!WinOptKeyword = 0 And rst!KeywordOrSkill = 0 Then
               .cmd(1).Enabled = False
            End If
         End If
         .lblDetail(0).Height = 1125
         .lblDetail(1).Visible = True

         .cmd(1).Caption = rst!OptionName
         .cmd(1).ToolTipText = rst!OptionName
         .lblDetail(1).Caption = rst!Details
      Else 'no option 2
         .cmd(1).Visible = False
         .lblDetail(1).Visible = False
         .lblDetail(0).Height = 2985
      End If
      rst.Close
      .Caption = "have fun Misbehavin' " & cnt & " of " & total
      .lblUnseen = "unseen: " & getUnseenMBDeck()
      .Alpha.Picture = LoadPictureGDIplus(App.Path & "\Pictures\suit" & suit & ".bmp")
      .Alpha.Visible = True
      .Alpha.TransparentColor = &HFFFFFF
      .Alpha.TransparentColorMode = lvicUseTransparentColor
      .Show 1
      
      getMisbehave = .MBCardID
      opt = .MBOption
      
   End With
      
End Function

'returns the Result Flag passed from doMisbehave: 1=proceed, 2=botched, 3=warrant, 4=load 1 contra per crew wit no gear
Public Function doMisbehaves(ByVal playerID, ByVal cnt As Integer, ByVal SectorID) As Integer
Dim x, CardID As Integer, opt, actualcnt As Integer, suit, c(1 To 4) As Integer
   actualcnt = 0
   For x = 1 To cnt
      CardID = 0
      While CardID = 0 'allow for reshuffle
         CardID = getMisbehave(opt, x, cnt, suit)
      Wend
      c(suit) = c(suit) + 1
      If opt = 3 Then
         actualcnt = actualcnt + 1
         'skip - ace in the hole
      Else
         doMisbehaves = doMisbehave(playerID, CardID, opt)
         Select Case doMisbehaves
         Case 1, 4 ' proceed
            actualcnt = actualcnt + 1
            'stamp the card so it can be counted for goals
            DB.Execute "UPDATE MisbehaveDeck SET Seq =" & playerID & " WHERE CardID =" & CardID
            frmAction.lblMis = CStr(countMisbehaves(playerID))
         Case 2 'botched

            Exit Function
         
         Case 3 'warrant issued-done, job discarded
            Exit Function

         End Select
      End If
      'refresh as stuff may have changed
      If Not (frmShip Is Nothing) Then frmShip.RefreshShips
   Next x
   'do double down check
   If doMisbehaves = 1 Then 'normal success (not 4)
      For x = 1 To 4
         If c(x) > 1 Then
            doMisbehaves = 5
            Exit For
         End If
      Next x
   End If
   
   'inform of the success
   If IsNull(varDLookup("PlanetName", "Planet", "SectorID=" & SectorID)) Then
      If getCruiserSector() = SectorID Then
         PutMsg player.PlayName & " Misbehaved successfully " & actualcnt & " times at the Alliance Cruiser", playerID, Logic!Gamecntr, True, getLeader()
      ElseIf getCorvetteSector() = SectorID Then
         PutMsg player.PlayName & " Misbehaved successfully " & actualcnt & " times at the Operative's Corvette", playerID, Logic!Gamecntr, True, getLeader()
      End If
   Else
      PutMsg player.PlayName & " Misbehaved successfully " & actualcnt & " times at " & varDLookup("PlanetName", "Planet", "SectorID=" & SectorID), playerID, Logic!Gamecntr, True, getLeader()
   End If

End Function

' CardID is from MisbehaveDeck, opt is which option selected returns: 1=proceed, 2=botched, 3=warrant, 4=load 1 contra per crew wit no gear
Public Function doMisbehave(ByVal playerID, ByVal CardID, ByVal opt) As Integer
Dim SQL, skillcnt, skillwin, skillint, skilldiscards, x, bribe As Integer, riverskill As Integer
Dim dice As Integer, Wskill As Integer, fruityBar As Integer, KeywordSkill As Integer, result '0=win,1-inter,2=fail
Dim rst As New ADODB.Recordset, frmDiscardGr As frmDiscardGear

   If opt = 0 Then
      MsgBox "option error", vbCritical, "Misbehave"
      Exit Function
   End If

   'grab the Nav Option chosen
   SQL = "SELECT MisOption.* "
   SQL = SQL & "FROM MisOption INNER JOIN MisbehaveDeck ON MisOption.OptionID = MisbehaveDeck.Option" & opt & "ID "
   SQL = SQL & "Where MisbehaveDeck.CardID = " & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
   
      PutMsg player.PlayName & "'s Misbehavin': " & rst!Details, player.ID, Logic!Gamecntr
      Events.getNewEvents
      Wskill = rst!skill
      'let the tests begin ... :O  WIN, INTER OR FAIL ?
      If rst!Win = 0 Then 'no test, just do Win outcomes
         result = 0
      ElseIf rst!ProfessionID = 8 And hasCrewAttribute(playerID, cstrProfession(rst!ProfessionID)) Then
         result = 0
      ElseIf rst!KeywordOrSkill > 0 And hasKeyword(playerID, rst!Keyword & "") Then
         result = 0
         
      ElseIf Wskill > 0 Then 'we have a skill test
         '-----------------------------------------
         'Stitch & Sheydra can change a Fight to a Nego once per Job
         If Wskill = 3 And hasCrew(playerID, 27) And Not usedStitchSkill Then
            If MessBox("Stitch wants to Fight instead of Negotiation.  Do you want to use those skills instead?", "Negotiate -> Fight", "Yes", "No", 27) = 0 Then
               Wskill = 1
               usedStitchSkill = True
               PutMsg player.PlayName & " uses Stitch's one time Negotiation to Fight Skills", player.ID, Logic!Gamecntr, True, 27
            End If
         End If
         If Wskill = 1 And getPerkAttributeCrew(playerID, "ChangeTestType") = 1 And Not usedStitchSkill Then
            If MessBox("Sheydra wants to Negotiate instead of Fight.  Do you want to use those skills instead?", "Fight -> Negotiate", "Yes", "No", 66) = 0 Then
               Wskill = 3
               usedStitchSkill = True
               PutMsg player.PlayName & " uses Sheydra's one time Fight to Negotiation Skills", player.ID, Logic!Gamecntr, True, 66
            End If
         End If
      
         'if card accepts a bribe, ask for $100 a point
         If rst!bribe = 1 Or hasPerkAttributeValue(player.ID, "Bribe", Wskill) Then
            Do
               bribe = Val(InputBox("They accept Bribes, $100 per skill point" & vbNewLine & vbNewLine & "Enter the number of POINTS you would bribe with..", "Money Talks", "0"))
               If bribe > 20 Then
                  MessBox "Seems a bit much don't ya think? Try that again..", "Too much!", "Ooops", "", getLeader()
               ElseIf bribe * 100 <= getMoney(playerID) Then 'can pay
                  getMoney playerID, (bribe * 100 * -1)
                  Exit Do
               Else
                  MessBox "Why you low-down thief, whatcha tryin' to pull?  Try again!", "Insufficient dough!", "Gorram it", "", getLeader()
               End If
            Loop
         End If
         
         'Crazy River Tam (cardID 51/CrewID 32)
         If hasCrew(playerID, 32) Then
            dice = RollDice(6)
            If hasCrew(playerID, 33) Then  'simon adds 2 to her rolls
               dice = dice + 2
            End If
            Select Case dice
            Case 1, 2 'stay onboard
               DB.Execute "UPDATE PlayerSupplies SET OffJob = 1 WHERE CardID = 51"
               If Not (frmShip Is Nothing) Then frmShip.RefreshShips  'update display
               PutMsg player.PlayName & "'s River Tam cowers onboard and won't be misbehavin' any further on this job", player.ID, Logic!Gamecntr, True, 32
            Case 3 'fight
               If Wskill = 1 Then
                  riverskill = 2
               End If
            Case 4 'Tech
               If Wskill = 2 Then
                  riverskill = 2
               End If
            Case 5 'negot
               If Wskill = 3 Then
                  riverskill = 2
               End If
            Case Else 'any skill
                  riverskill = 2
            End Select
            If riverskill = 2 Then
               'If hasCrew(player.ID, 33) Then riverskill = 4
               PutMsg player.PlayName & "'s River Tam" & IIf(hasCrew(playerID, 33), ", encouraged by Simon,", "") & " channels the " & cstrSkill(Wskill) & " skill + 2", player.ID, Logic!Gamecntr, True, 32
            ElseIf dice > 2 Then
               PutMsg player.PlayName & "'s River Tam ain't misbehavin' this time", player.ID, Logic!Gamecntr, True, 32
            End If
         End If
         
         fruityBar = hasGearCard(player.ID, 24)
         If fruityBar > 0 Then 'we got one or more
            If MessBox("Do you wish to Eat the Fruity Bar and add 1 to the Test Roll?", "Extra Bite", "Yes", "No", 0, 24) = 0 Then
               doDiscardGear player.ID, fruityBar
               fruityBar = 1
            Else
               fruityBar = 0
            End If
         End If
         
         x = hasGearCrew(player.ID, 28) 'Mal's Brown Coat
         If x > 0 And varDLookup("Disgruntled", "Crew", "CrewID=" & x) > 0 And varDLookup("Fight", "Crew", "CrewID=" & x) > 0 And Wskill = 3 Then
            fruityBar = fruityBar + varDLookup("Fight", "Crew", "CrewID=" & x)
            PutMsg player.PlayName & "'s Disgruntled Crew wearing the Brown Coat adds their Fight skills to the Negotiation", player.ID, Logic!Gamecntr, True, 0, 28
         End If
         
         If Wskill = 1 Then
            removeDigruntled player.ID, Wskill ' Mal's Frontier Model B -Before each Fight Test, remove Disgruntled from the Owner.
         End If
         
         '<<<<<<<<<< ROLL THE DICE >>>>>>>>>>>>>>>>>
         dice = RollDice(6, IIf(Wskill = 3 And hasCrew(player.ID, 55), False, True)) 'Bester -On negotiate test, +6 "Thillin' Heroics" bonus dice does not apply
         
         
         If Wskill = 1 And hasGear(player.ID, 47) Then ' Zoe's Mare's Leg Rifle -When making a Fight Test, roll two dice and use the highest.
            x = RollDice(6, True)
            If x > dice Then
               PutMsg player.PlayName & " had rolled a " & CStr(dice) & " so using Zoe's Mare's Leg Rifle rerolled a " & CStr(x), player.ID, Logic!Gamecntr, True, 0, 47, 0, 0, 0, dice
               dice = x
            End If
         End If
         
         If dice = 1 Then  'reroll ones?
            If hasGear(player.ID, 6) Then ' has Jaynes Cunning Hat
               Do While dice = 1
                  dice = RollDice(6, IIf(Wskill = 3 And hasCrew(player.ID, 55), False, True))
               Loop
               PutMsg player.PlayName & " uses Jaynes Cunning Hat to reRoll a 1 and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, 0, 6, 0, 0, 0, dice
               
            ElseIf hasGear(player.ID, 35) And Wskill = 1 Then 'Inara's Bow
               x = hasGearCrew(player.ID, 35)
               If x > 0 Then
                  If hasCrewAttribute(playerID, "Companion", 0, x) Then
                     Do While dice = 1
                        dice = RollDice(6, True)
                     Loop
                     PutMsg player.PlayName & " uses Inara's Bow to reRoll a 1 and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, 0, 35, 0, 0, 0, dice
                  End If
               End If
            End If
         End If
         
         'Inara & Kaylee can reroll a negotiate test
         If Wskill = 3 And getPerkAttributeCrew(player.ID, "RerollNegotiate") And dice < 6 Then
            If MessBox("You rolled a " & dice & vbNewLine & "Your Negotiation Skills allow you a second chance, do you want to take that chance?", "Re-Roll option", "Yes", "No", getLeader(), 0, 0, dice) = 0 Then
               dice = RollDice(6, IIf(Wskill = 3 And hasCrew(player.ID, 55), False, True))
               PutMsg player.PlayName & " uses extra Negotiation Skills to reRoll and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
            End If
         End If
         'Zoe's skill
         If Wskill = 1 And getPerkAttributeCrew(player.ID, "RerollFight") And dice < 6 Then
            If MessBox("You rolled a " & dice & vbNewLine & "Your Fight Skills allow you a second chance, do you want to take that chance?", "Re-Roll option", "Yes", "No", getLeader(), 0, 0, dice) = 0 Then
               dice = RollDice(6, IIf(Wskill = 3 And hasCrew(player.ID, 55), False, True))
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, getLeader()
            End If
         End If
         
         If Wskill = 1 And hasGear(player.ID, 45) And dice < 6 Then 'yolanda's pistol - Discard to re-roll a Fight Test.
            If MessBox("You rolled a " & dice & vbNewLine & "Yolanda's pistol allows you a second chance, do you want to Discard the Pistol to take that extra chance?", "Re-Roll option", "Yes", "No", 0, 45, 0, dice) = 0 Then
               doDiscardGear player.ID, hasGearCard(player.ID, 45)
               dice = RollDice(6, IIf(Wskill = 3 And hasCrew(playerID, 55), False, True))
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(dice), playerID, Logic!Gamecntr, True, getLeader()
            End If
         End If
         
         If Wskill = 1 And hasGear(player.ID, 48) And dice < 6 Then 'Extra Ammo Clip - Discard to re-roll a Fight Test.
            If MessBox("You rolled a " & dice & vbNewLine & "Extra Ammo Clips allow you a second chance, do you want to Discard the Clips to take that extra chance?", "Re-Roll option", "Yes", "No", 0, 48, 0, dice) = 0 Then
               doDiscardGear player.ID, hasGearCard(player.ID, 48)
               dice = RollDice(6, IIf(Wskill = 3 And hasCrew(playerID, 55), False, True))
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(dice), playerID, Logic!Gamecntr, True, getLeader()
            End If
         End If
         
         '----------------------------------------- see if we need to use the discardable skills & keywords...
         
         If rst!WinOptKeyword > 0 And hasKeyword(player.ID, rst!Keyword & "") Then 'keyword reduces win minimum - also could be discardable ???
            KeywordSkill = rst!WinOptKeyword
         End If
         
         skillwin = rst!Win
         
         skillint = rst!Intermediate

         'get our skill totals, exclude gear from Kosherized rules
         skillcnt = getSkill(player.ID, cstrSkill(Wskill), 0, True, (rst!kosher = 1)) + dice + bribe + riverskill + fruityBar + KeywordSkill
         skilldiscards = getSkillDiscards(player.ID, cstrSkill(Wskill), (rst!kosher = 1))

         
         If skillcnt < skillwin And skillcnt + skilldiscards >= skillwin Then  'we're in trouble 'we could use some help
            If MessBox("With the help of " & skillwin - skillcnt & " skill points, we can succeed" & vbNewLine & "Do you want to use a discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
               'show a list of gear to pick from up to or exceeding the value skillwin - skillcnt
               Set frmDiscardGr = New frmDiscardGear
               frmDiscardGr.kosher = (rst!kosher = 1)
               frmDiscardGr.nbrSelect = skillwin - skillcnt
               frmDiscardGr.skill = cstrSkill(Wskill)
               frmDiscardGr.Caption = "Select single use Gear to provide at least " & CStr(frmDiscardGr.nbrSelect) & " skill points"
               frmDiscardGr.Show 1
               'then add selected skill points to skillcnt, discard gear, and go on...
               skillcnt = skillcnt + frmDiscardGr.nbrSelected
            End If
                  
         ElseIf skillcnt < skillint And skillcnt + skilldiscards >= skillint Then  'we're in trouble 'we could use some help
            If MessBox("With the help of " & skillint - skillcnt & " skill points, we can make the intermediate outcome" & vbNewLine & "Do you want to use discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
               'show a list of gear to pick from up to or exceeding the value skillint - skillcnt
               Set frmDiscardGr = New frmDiscardGear
               frmDiscardGr.kosher = (rst!kosher = 1)
               frmDiscardGr.nbrSelect = skillint - skillcnt
               frmDiscardGr.skill = cstrSkill(Wskill)
               frmDiscardGr.Caption = "Select single use Gear to provide at least " & CStr(frmDiscardGr.nbrSelect) & " skill points"
               frmDiscardGr.Show 1
               'then add selected skill points to skillcnt, discard gear, and go on...
               skillcnt = skillcnt + frmDiscardGr.nbrSelected
            End If
         
         End If
         
         
         If hasGear(player.ID, 32) And Wskill = 1 And skillcnt < rst!Win Then
            If MessBox("The Fights not going so well with a skill score of " & skillcnt & vbNewLine & "Simon's Sonic Stun Baton might turn things around, wanna try another Thrillin' Heroics Roll and Discard the Baton?", "Stun Baton to the Fight", "Yes", "No", 0, 32) = 0 Then
               skillcnt = RollDice(6) + 6
               doDiscardGear player.ID, hasGearCard(player.ID, 32)
               PutMsg player.PlayName & " used Simon's Sonic Stun Baton to try and turn the Fight around ", player.ID, Logic!Gamecntr
            End If
         End If
         '-----------------------------------------
         
         If skillcnt >= rst!WinOptKeyword And rst!WinOptKeyword > 0 And hasKeyword(player.ID, rst!Keyword & "") Then
            If skillcnt < rst!Win Then 'needed the keyword to win
               'check if the keyword was single use, and discard
               discardGearKeyword player.ID, rst!Keyword
            End If
            result = 0
         ElseIf skillcnt >= rst!Win Then
            result = 0
         ElseIf skillcnt >= rst!Intermediate And rst!Intermediate > 0 Then
            result = 1
         Else 'you lose
            result = 2
         End If
         PutMsg player.PlayName & "'s MB log: Rolls a " & dice & " with added " & cstrSkill(Wskill) & " skill points to a total of " & skillcnt & " to " & IIf(result = 0, "succeed :^)", IIf(result = 1, "partially succeed :^|", "lose :^(")), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
         
      End If  'end of the initial Tests
         
      Select Case result
      Case 0 ' winners are grinners :D
         doMisbehave = rst!WinResult
         
         If rst!WinCash > 0 Then
            DB.Execute "UPDATE Players Set Pay = Pay + " & rst!WinCash & " WHERE PlayerID = " & player.ID
         ElseIf rst!WinCash < 0 Then
            dice = rst!WinCash
            If getMoney(playerID) <= Abs(dice) Then
               dice = getMoney(playerID) * -1
               PutMsg player.PlayName & "'s MB log: Funds depleted!", player.ID, Logic!Gamecntr, True, getLeader()
            End If
            
            DB.Execute "UPDATE Players Set Pay = Pay + " & dice & " WHERE PlayerID = " & player.ID
            
         End If

         If rst!WinKillCrew <> 0 Then
            doKillCrews player.ID, rst!WinKillCrew
         End If
         
      Case 1 'intermediate outcomes  :|
         doMisbehave = rst!InterResult
    
         If rst!InterKillCrew <> 0 Then
            doKillCrews player.ID, rst!InterKillCrew
         End If
         
      Case 2 'loser outcomes :(
         doMisbehave = rst!FailResult

         If rst!FailKillCrew = 99 Then ':((
            doKillAllCrew player.ID
         ElseIf rst!FailKillCrew <> 0 Then
            doKillCrews player.ID, rst!FailKillCrew
         End If
         
         If rst!Disgruntled = 4 Then 'discard all Mercs
            If doMercDiscard(player.ID) Then
               PutMsg player.PlayName & " had the Mercs mutiny and leave", player.ID, Logic!Gamecntr, True, getLeader()
            End If
         End If
            
      End Select
       
      'DO the tests that run whatever the above outcome -----------------------------------------
      
      'check if the keyword was single use, and discard
      If Nz(rst!Keyword, "") <> "" Then
         If discardGearKeyword(player.ID, Nz(rst!Keyword), True) Then
            discardGearKeyword player.ID, rst!Keyword
            PutMsg player.PlayName & " used up the " & rst!Keyword, player.ID, Logic!Gamecntr
         End If
      End If
      
      If rst!cargo <> 0 Then ' could be -neg
         skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
         If skillcnt > rst!cargo Then skillcnt = rst!cargo
         If skillcnt <> 0 Then
            DB.Execute "UPDATE Players Set Cargo = Cargo + " & skillcnt & " WHERE PlayerID = " & player.ID
            PutMsg player.PlayName & IIf(skillcnt > 0, " scored ", " lost ") & skillcnt & " Cargo", player.ID, Logic!Gamecntr
         End If
      End If
      
      If rst!Contraband <> 0 Then ' could be -neg
         If doMisbehave = 4 Then ' one per Crew with no gear (doMisbehave=4)
            x = getCrewWithNoGear(player.ID)
         Else
            x = rst!Contraband
         End If
         skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
         If skillcnt > x Then skillcnt = x
         If skillcnt <> 0 Then
            DB.Execute "UPDATE Players Set Contraband = Contraband + " & skillcnt & " WHERE PlayerID = " & player.ID
            PutMsg player.PlayName & IIf(skillcnt > 0, " scored ", " lost ") & skillcnt & " Contraband", player.ID, Logic!Gamecntr
         End If
      End If
      
      '1-Moral only, 2-All Crew, 3-remove from Moral Crew,
      '4=Discard all Mercs, 5=Discard Mercs if fight is greater than crew
      
      If (rst!Disgruntled > 0 And rst!Disgruntled < 4 And (result = 2 Or rst!skill = 0)) Or (rst!Disgruntled = 6 And result = 0) Or (rst!Disgruntled = 3 And result = 0) Then 'apply disgruntled changes if lose or no test
         doDisgruntled player.ID, rst!Disgruntled
         result = 0
      End If
      
      If rst!Disgruntled = 5 Then 'discard all Mercs if Merc Fight higher than others
         If getSkill(playerID, cstrSkill(1), 1) > getSkill(playerID, cstrSkill(1), 2) Then  'discard all Mercs
            If doMercDiscard(player.ID) Then
               PutMsg player.PlayName & " had the Mercs outgun the crew and leave", player.ID, Logic!Gamecntr, True, getLeader()
            End If
         End If
      End If
                  
   Else
      MsgBox "Error: Nav Card " & CardID & " Option " & opt & " not found!", vbCritical
   End If

Set rst = Nothing

End Function

'this is where we apply the 1000 rules and outcomes of the nav option :(
'these apply to FULLBURN only. to Full Stop, set fullburndone = True
Public Function doNav(ByVal CardID, ByVal opt) As Boolean
Dim SQL, SectorID, skillcnt, skillwin, skillint, skilldiscards, x, y, z, bribe As Integer
Dim dice As Integer, riverskill As Integer, fruityBar As Integer, result '0=win,1-inter,2=fail
Dim rst As New ADODB.Recordset
Dim frmShUp As frmShipUpgd, frmDiscardGr As frmDiscardGear, frmBart As frmBarter
Dim frmSalvage As frmSalvaging, frmCrewList As frmCrewLst, frmSeize As frmSeized

   'grab the Nav Option chosen
   SQL = "SELECT NavOption.* "
   SQL = SQL & "FROM NavOption INNER JOIN NavDeck ON NavOption.OptionID = NavDeck.Option" & opt & "ID "
   SQL = SQL & "Where NavDeck.CardID = " & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      PutMsg player.PlayName & "'s Nav log: " & rst!Details, player.ID, Logic!Gamecntr
      'let the tests begin ... :O  WIN, INTER OR FAIL ?
      
      'has breakdown insurance ?
      If rst!Breakdown = 1 And (hasShipUpgrade(player.ID, 3) Or hasShipUpgrade(player.ID, 7)) Then
         result = 0
         PutMsg player.PlayName & "'s Ship is breakdown proof!", player.ID, Logic!Gamecntr, True, getLeader()
         Exit Function
         
      ElseIf rst!WinProfession > 0 And Not hasCrewAttribute(player.ID, cstrProfession(rst!WinProfession)) Then
         result = 2
         
      ElseIf rst!skill = 0 Then 'no test, just do Win outcomes
         result = 0
     
      ElseIf rst!skill > 0 Then 'we have a skill test
      
         'if card accepts a bribe, ask for $100 a point
         If rst!bribe = 1 Or (getPerkAttributeCrew(player.ID, "Bribe") > 0 And rst!skill = 3) Then
            Do
               bribe = Val(InputBox("They accept Bribes on this Job, $100 per skill point" & vbNewLine & vbNewLine & "Enter the number of points you would bribe with..", "Money Talks", "0"))
               If bribe > 20 Then
                  MessBox "Seems a bit much don't ya think? Try that again..", "Too much!", "Ooops", "", getLeader()
               ElseIf bribe * 100 <= getMoney(player.ID) Then 'can pay
                  getMoney player.ID, (bribe * 100 * -1)
                  Exit Do
               Else
                  MessBox "Why you lousy thief, whataya tryin' to pull?  Try again will ya!", "Insufficient dough!", "Sorry", "", getLeader()
               End If
            Loop

         End If
         
         'Crazy River Tam (cardID 51/CrewID 32)
         If hasCrew(player.ID, 32) Then
            dice = RollDice(6)
            If hasCrew(player.ID, 33) Then 'simon adds 2 to her rolls
               dice = dice + 2
            End If
            Select Case dice
            Case 1, 2 'stay onboard
               DB.Execute "UPDATE PlayerSupplies SET OffJob = 1 WHERE CardID = 51"
               If Not (frmShip Is Nothing) Then frmShip.RefreshShips  'update display
               PutMsg player.PlayName & "'s River Tam cowers onboard and won't be playin'", player.ID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, dice
            Case 3 'fight
               If rst!skill = 1 Then
                  riverskill = 2
               End If
            Case 4 'Tech
               If rst!skill = 2 Then
                  riverskill = 2
               End If
            Case 5 'negot
               If rst!skill = 3 Then
                  riverskill = 2
               End If
            Case Else 'any skill
                  riverskill = 2
            End Select
            If riverskill = 2 Then
               PutMsg player.PlayName & "'s River Tam" & IIf(hasCrew(player.ID, 33), ", encouraged by Simon,", "") & " channels the " & cstrSkill(rst!skill) & " skill + 2", player.ID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, dice
            ElseIf dice > 2 Then
               PutMsg player.PlayName & "'s River Tam ain't gettin' involved this time", player.ID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, dice
            End If
         End If
         
         fruityBar = hasGearCard(player.ID, 24)
         If fruityBar > 0 Then 'we got one or more
            If MessBox("Do you wish to Eat the Fruity Bar and add 1 to the Test Roll?", "Extra Bite", "Yes", "No", 0, 24) = 0 Then
               doDiscardGear player.ID, fruityBar
               fruityBar = 1
            Else
               fruityBar = 0
            End If
         End If

         x = hasGearCrew(player.ID, 28) 'Mal's Brown Coat
         If x > 0 And varDLookup("Disgruntled", "Crew", "CrewID=" & x) > 0 And varDLookup("Fight", "Crew", "CrewID=" & x) > 0 And rst!skill = 3 Then
            fruityBar = fruityBar + varDLookup("Fight", "Crew", "CrewID=" & x)
            PutMsg player.PlayName & "'s Disgruntled Crew wearing the Brown Coat adds their Fight skills to the Negotiation", player.ID, Logic!Gamecntr, True, 0, 28
         End If
            
         If rst!skill = 1 Then
            removeDigruntled player.ID, rst!skill
         End If
            
         'Roll the flippin Dice already!!!
         dice = RollDice(6, IIf(rst!skill = 3 And hasCrew(player.ID, 55), False, True))
         
         If rst!skill = 1 And hasGear(player.ID, 47) Then ' Zoe's Mare's Leg Rifle -When making a Fight Test, roll two dice and use the highest.
            x = RollDice(6, IIf(rst!skill = 3 And hasCrew(player.ID, 55), False, True))
            If x > dice Then
               PutMsg player.PlayName & " had rolled a " & CStr(dice) & " so using Zoe's Mare's Leg Rifle rerolled a " & CStr(x), player.ID, Logic!Gamecntr, True, 0, 47, 0, 0, 0, x
               dice = x
            End If
         End If
         
         If dice = 1 Then  'reroll ones?
            If hasGear(player.ID, 35) And rst!skill = 1 Then 'Inara's Bow
               x = hasGearCrew(player.ID, 35)
               If x > 0 Then
                  If hasCrewAttribute(player.ID, "Companion", 0, x) Then
                     Do While dice = 1
                        dice = RollDice(6, True)
                     Loop
                     PutMsg player.PlayName & " uses Inara's Bow to reRoll a 1 and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, 0, 35, 0, 0, 0, dice
                  End If
               End If
            End If
         End If
         
         'Inara & Kaylee can reroll a negotiate test
         If rst!skill = 3 And getPerkAttributeCrew(player.ID, "RerollNegotiate") And dice < 6 Then
            If MessBox("You rolled a " & dice & vbNewLine & "Your Negotiation Skills allow you a second chance, do you want to take that chance?", "Re-Roll option", "Yes", "No", getLeader()) = 0 Then
               dice = RollDice(6, IIf(rst!skill = 3 And hasCrew(player.ID, 55), False, True))
               PutMsg player.PlayName & " uses extra Negotiation Skills to reRoll and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
            End If
         End If
         'Zoe's skill
         If rst!skill = 1 And getPerkAttributeCrew(player.ID, "RerollFight") And dice < 6 Then
            If MessBox("You rolled a " & dice & vbNewLine & "Your Fight Skills allow you a second chance, do you want to take that extra chance?", "Re-Roll option", "Yes", "No", getLeader()) = 0 Then
               dice = RollDice(6, IIf(rst!skill = 3 And hasCrew(player.ID, 55), False, True))
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
            End If
         End If
         
         If rst!skill = 1 And hasGear(player.ID, 45) And dice < 6 Then 'yolanda's pistol - Discard to re-roll a Fight Test.
            If MessBox("You rolled a " & dice & vbNewLine & "Yolanda's pistol allows you a second chance, do you want to Discard the Pistol to take that extra chance?", "Re-Roll option", "Yes", "No", 0, 45) = 0 Then
               doDiscardGear player.ID, hasGearCard(player.ID, 45)
               dice = RollDice(6, IIf(rst!skill = 3 And hasCrew(player.ID, 55), False, True))
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
            End If
         End If
         
         If rst!skill = 1 And hasGear(player.ID, 48) And dice < 6 Then 'Extra Ammo Clip - Discard to re-roll a Fight Test.
            If MessBox("You rolled a " & dice & vbNewLine & "Extra Ammo Clips allow you a second chance, do you want to Discard the Clips to take that extra chance?", "Re-Roll option", "Yes", "No", 0, 48) = 0 Then
               doDiscardGear player.ID, hasGearCard(player.ID, 48)
               dice = RollDice(6, IIf(rst!skill = 3 And hasCrew(player.ID, 55), False, True))
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(dice), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
            End If
         End If
         
         'skillcnt = getSkill(player.ID, cstrSkill(rst!skill)) + dice
         '----------------------------------------- see if we need to use the discardable skills & keywords...
         
         skillwin = rst!Win
         skillint = rst!Intermediate

         'get our skill totals
         skillcnt = getSkill(player.ID, cstrSkill(rst!skill), 0, True) + dice + bribe + riverskill + fruityBar
         skilldiscards = getSkillDiscards(player.ID, cstrSkill(rst!skill))
         
         '-----------------------------------------
         If skillcnt < skillwin And skillcnt + skilldiscards >= skillwin Then 'we're in trouble 'we could use some help
            If MessBox("Rolled a " & CStr(dice) & vbNewLine & "With the help of " & skillwin - skillcnt & " single use skill points, we can succeed." & vbNewLine & "Do you want to use discardable Gear items for this?", "Skill Test Trouble", "Yes", "No", getLeader()) = 0 Then
               'show a list of gear to pick from up to or exceeding the value skillwin - skillcnt
               Set frmDiscardGr = New frmDiscardGear
               frmDiscardGr.nbrSelect = skillwin - skillcnt
               frmDiscardGr.skill = cstrSkill(rst!skill)
               frmDiscardGr.Show 1
               'then add selected skill points to skillcnt, discard gear, and go on...
               skillcnt = skillcnt + frmDiscardGr.nbrSelected
            End If
                  
         ElseIf skillint > 0 And skillcnt < skillint And skillcnt + skilldiscards >= skillint Then 'we're in trouble 'we could use some help
            If MessBox("Rolled a " & CStr(dice) & vbNewLine & "With the help of " & skillint - skillcnt & " single use skill points, we can make the intermediate outcome." & vbNewLine & "Do you want to use discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
               'show a list of gear to pick from up to or exceeding the value skillint - skillcnt
               Set frmDiscardGr = New frmDiscardGear
               frmDiscardGr.nbrSelect = skillint - skillcnt
               frmDiscardGr.skill = cstrSkill(rst!skill)
               frmDiscardGr.Show 1
               'then add selected skill points to skillcnt, discard gear, and go on...
               skillcnt = skillcnt + frmDiscardGr.nbrSelected
            End If
         
         End If
         
         If hasGear(player.ID, 32) And rst!skill = 1 And skillcnt < rst!Win Then
            If MessBox("The Fights not going so well with a skill score of " & skillcnt & vbNewLine & "Simon's Sonic Stun Baton might turn things around, wanna try another Thrillin' Heroics Roll and Discard the Baton?", "Stun Baton to the Fight", "Yes", "No", 0, 32) = 0 Then
               skillcnt = RollDice(6) + 6
               doDiscardGear player.ID, hasGearCard(player.ID, 32)
            End If
         End If
         '-----------------------------------------
         
         If skillcnt >= rst!Win Then
            result = 0
         ElseIf skillcnt >= rst!Intermediate And rst!Intermediate > 0 Then
            result = 1
         Else 'you lose
            result = 2
         End If
         PutMsg player.PlayName & "'s Nav log: Rolls a " & dice & " with added " & cstrSkill(rst!skill) & " skill points to a total of " & skillcnt & " to " & IIf(result = 0, "succeed :^)", IIf(result = 1, "partially succeed :^|", "lose :^(")), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
         
      End If
         
      Select Case result
      Case 0 ' winners are grinners :D
         If rst!WinKeepFlying = 0 Then  'full stop
            frmAction.fullburndone = True
            frmAction.moseydone = True
         End If
         If rst!WinCash <> 0 Then
            DB.Execute "UPDATE Players Set Pay = Pay + " & rst!WinCash & " WHERE PlayerID = " & player.ID
         End If
         If rst!WinCargo <> 0 Then ' could be -neg
            skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            If skillcnt > rst!WinCargo Then skillcnt = rst!WinCargo
            If skillcnt <> 0 Then
               DB.Execute "UPDATE Players Set Cargo = Cargo + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!WinPassenger <> 0 Then ' could be -neg
            skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            If skillcnt < rst!WinPassenger And rst!WinPassenger = 2 Then 'cannot fit eryone
               PutMsg player.PlayName & " couldn't fit all Passengers, Moral Crew are not going to be happy!", player.ID, Logic!Gamecntr, True, getLeader()
               doDisgruntled player.ID, 1
            ElseIf skillcnt > rst!WinPassenger Then
               skillcnt = rst!WinPassenger
            End If
            If skillcnt <> 0 Then
               DB.Execute "UPDATE Players Set Passenger = Passenger + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!WinFugitive <> 0 Then ' could be -neg
            skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            If skillcnt < rst!WinFugitive And rst!WinFugitive = 4 Then 'cannot fit eryone
               PutMsg player.PlayName & " couldn't fit all Fugitives, Moral Crew are not going to be happy!", player.ID, Logic!Gamecntr, True, getLeader()
               doDisgruntled player.ID, 1
            ElseIf skillcnt > rst!WinFugitive Then
               skillcnt = rst!WinFugitive
            End If
            If skillcnt <> 0 Then
               DB.Execute "UPDATE Players Set Fugitive = Fugitive + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!WinFuel < 0 Then ' -neg
            skillcnt = rst!WinFuel
            If varDLookup("Fuel", "Players", "PlayerID=" & player.ID) >= Abs(skillcnt) Then  'check we're not going -ve
               DB.Execute "UPDATE Players Set Fuel = Fuel + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         ElseIf rst!WinFuel = 14 Then ' all you can load
            skillcnt = (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID)) * 2
            Do
               x = Val(InputBox("Select up to " & skillcnt & " Fuel to salvage", "Salvage Fuel"))
               If x <= skillcnt And x > -1 Then
                  skillcnt = x
                  Exit Do
               Else
                  MessBox "Invalid Fuel quantity", "Fuel Requirements", "Ooops", "", getLeader()
               End If
            Loop
            DB.Execute "UPDATE Players Set Fuel = Fuel + " & skillcnt & " WHERE PlayerID = " & player.ID
         ElseIf rst!WinFuel > 0 Then ' small +ve
            skillcnt = (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID)) * 2
            If skillcnt > rst!WinFuel Then skillcnt = rst!WinFuel
            DB.Execute "UPDATE Players Set Fuel = Fuel + " & skillcnt & " WHERE PlayerID = " & player.ID
         End If
         
         If rst!WinParts = -99 Then 'sell up to 3 parts for $500ea
            Do
               y = varDLookup("Parts", "Players", "PlayerID=" & player.ID)
               If y = 0 Then Exit Do
               x = Val(InputBox("How many Parts (you have " & y & ") would you like to sell for $500ea?", "Sell Parts", "0"))
               If x > y Then
                  MessBox "Invalid Parts quantity", "Parts Requirements", "Ooops", "", getLeader()
               Else
                  If x > 0 Then
                     DB.Execute "UPDATE Players SET Parts = Parts - " & x & ", Pay = Pay + " & CStr(x * 500) & " WHERE PlayerID=" & player.ID
                  End If
                  Exit Do
               End If
            Loop
            
         ElseIf rst!WinParts <> 0 Then ' could be -neg  . skillcnt re-used to count parts here
            skillcnt = (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID)) * 2
            If skillcnt > rst!WinParts Or rst!WinParts < 0 Then skillcnt = rst!WinParts
            If skillcnt * -1 > varDLookup("Parts", "Players", "PlayerID=" & player.ID) Then 'stop going neg
               skillcnt = Val(varDLookup("Parts", "Players", "PlayerID=" & player.ID)) * -1
               If rst!Breakdown = 1 Then 'no parts, fullstop as breakdown proof is tested at start /|\
                  frmAction.fullburndone = True
                  frmAction.moseydone = True
               End If
            End If
            If skillcnt <> 0 Then DB.Execute "UPDATE Players Set Parts = Parts + " & skillcnt & " WHERE PlayerID = " & player.ID
         End If
         If rst!WinContraband <> 0 Then ' could be -neg
            skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            If skillcnt > rst!WinContraband Then skillcnt = rst!WinContraband
            If skillcnt <> 0 Then
               DB.Execute "UPDATE Players Set Contraband = Contraband + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!WinShipUpgrade <> 0 Then
            
            'present list of discarded upgrades to choose one for free
            Set frmShUp = New frmShipUpgd
            If getShipUpgrades(player.ID) < 3 Then
               frmShUp.discardMode = rst!WinShipUpgrade
            Else 'DriveCores only, no spare slots
               frmShUp.discardMode = 5
            End If
            frmShUp.Show 1
            
         End If
         
         If rst!WinGoods > 0 Then
            Set frmSalvage = New frmSalvaging
            frmSalvage.mode = 1
            frmSalvage.salvageCount = rst!WinGoods
            frmSalvage.Show 1
         ElseIf rst!WinGoods = -99 Then
            DB.Execute "UPDATE Players Set Fuel = 0, Parts = 0, Cargo = 0, Contraband = 0 WHERE PlayerID = " & player.ID
            PutMsg player.PlayName & " lost all Goods overboard", player.ID, Logic!Gamecntr
         End If
         
         If rst!WinKillCrew <> 0 Then
            x = doKillCrews(player.ID, rst!WinKillCrew)
            If rst!OptionName = "If we're very lucky" And hasShipUpgrade(player.ID, 18) > 0 And x > 0 Then
               doDiscardGear player.ID, hasShipUpgrade(player.ID, 18)
               PutMsg player.PlayName & " lost the Reaver-Flage upgrade in the Reaver skuffle", player.ID, Logic!Gamecntr
            End If

         End If
         
         'UNIQUE WIN OPTIONS-----------------------------------
         If rst!WinFunction > 0 Then  'here lies all the new weird functions
            Select Case rst!WinFunction
               Case 1 ' Add 1 to the Range of this Fly Action for each Moral Crew on board
                  turnExtraRange = countCrewAttribute(player.ID, "Moral")
                  frmAction.lblRange.Caption = CStr(Val(frmAction.lblRange.Caption) + turnExtraRange)
               
               Case 2 'Gambling
                  If getMoney(player.ID) < 1000 Then
                     MessBox "You don't have the Cash to make the bet", "Cashflow Problem", "Ooops", "", getLeader()
                  Else
                     x = RollDice(6)
                     If x > 4 Then
                        getMoney player.ID, 2000
                        PutMsg player.PlayName & " rolls a " & x & " and wins $2000", player.ID, Logic!Gamecntr, True, getLeader()
                     Else
                        getMoney player.ID, -1000
                        PutMsg player.PlayName & " rolls a " & x & " and loses $1000", player.ID, Logic!Gamecntr, True, getLeader()
                     End If
                  End If
               
               Case 3 'Passengers and Fugitives dispute
                  x = RollDice(6, True) 'use Thrillin heroics roll
                  skillcnt = x
                  y = varDLookup("Fugitive", "Players", "PlayerID=" & player.ID)
                  z = varDLookup("Passenger", "Players", "PlayerID=" & player.ID)
                  If x < (y + z) Then  'out they go -auto mode. preference to Passengers go first
                     If x >= y Then
                        x = x - y
                        y = 0
                        If x >= z Then
                           z = 0
                        Else
                           z = z - x
                        End If
                     Else
                        y = y - x
                     End If
                     DB.Execute "UPDATE Players SET Passenger =" & z & ", Fugitive =" & y & " WHERE PlayerID=" & player.ID
                     PutMsg player.PlayName & " rolls a " & skillcnt & " and is left with " & z & " Passengers and " & y & " Fugitives", player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, skillcnt
                  Else
                     PutMsg player.PlayName & " rolls a " & skillcnt & " and retains any Passengers and Fugitives", player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, skillcnt
                  End If
                  
               Case 4   'She's Hemmorrhaging Fuel!
                  HemmorrhagingFuel = True  'set global - picked up by frmAction
                  
               Case 5   'Slingshot Roundhouse
                  If getCrewAttribute(player.ID, cstrProfession(2)) > 0 And getPlanetID(player.ID) > 0 Then
                     turnExtraRange = 3
                     frmAction.lblRange.Caption = CStr(Val(frmAction.lblRange.Caption) + turnExtraRange)
                  End If
                  
               Case 6
                  'Fancy Meetin' You Here, take 1 Crew Card from any discard pile for free
                  If getCrewCount(player.ID) < CrewCapacity(player.ID) Then
                     Set frmCrewList = New frmCrewLst
                     frmCrewList.selectCrew = -1
                     frmCrewList.Caption = "Select 1 Crew from Discards"
                     frmCrewList.Show 1
                  End If
                                    
               Case 7  'Shanghai Surprise! Take 1 Crew from Regina's Discard Pile.
                  If getCrewCount(player.ID) < CrewCapacity(player.ID) Then
                     Set frmCrewList = New frmCrewLst
                     frmCrewList.selectCrew = -1
                     frmCrewList.SupplyID = 5 'Regina
                     frmCrewList.Caption = "Select 1 Regina Crew from Discards"
                     frmCrewList.Show 1
                  End If
                  
            End Select
         End If
         
         If rst!SalvageOp <> 0 Then 'last win function, ignored if lose
            'load any Crew modifiers to add salvage due to Perk (SOCargo, SOContra...)
            'if can fit them of course
            doSalvage player.ID
         End If
         
      Case 1 'intermediate outcomes  :|
         If rst!InterKeepFlying = 0 Then  'full stop
            frmAction.fullburndone = True
            frmAction.moseydone = True
         End If
         If rst!InterGoods > 0 Then
            Set frmSalvage = New frmSalvaging
            frmSalvage.mode = 1
            frmSalvage.salvageCount = rst!InterGoods
            frmSalvage.Show 1
         End If
         If rst!InterCargo <> 0 Then ' could be -neg
            DB.Execute "UPDATE Players Set Cargo = Cargo + " & rst!InterCargo & " WHERE PlayerID = " & player.ID
         End If
         If rst!InterKillCrew <> 0 Then
            x = doKillCrews(player.ID, rst!InterKillCrew)
            If rst!OptionName = "If we're very lucky" And hasShipUpgrade(player.ID, 18) > 0 And x > 0 Then
               doDiscardGear player.ID, hasShipUpgrade(player.ID, 18)
               PutMsg player.PlayName & " lost the Reaver-Flage upgrade in the Reaver skuffle", player.ID, Logic!Gamecntr
            End If

         End If
         
      Case 2 'loser outcomes :(
         If rst!FailKeepFlying = 0 Then  'full stop
            frmAction.fullburndone = True
            frmAction.moseydone = True
         End If
         If rst!FailCargo <> 0 Then ' could be -neg
            If (rst!FailCargo * -1) <= varDLookup("Cargo", "Players", "PlayerID = " & player.ID) Then
               DB.Execute "UPDATE Players Set Cargo = Cargo + " & rst!FailCargo & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!FailFuel <> 0 Then ' could be -neg
            If (rst!FailFuel * -1) <= varDLookup("Fuel", "Players", "PlayerID = " & player.ID) Then
               DB.Execute "UPDATE Players Set Fuel = Fuel + " & rst!FailFuel & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!FailParts <> 0 Then ' could be -neg
            If (rst!FailParts * -1) <= varDLookup("Parts", "Players", "PlayerID = " & player.ID) Then
               DB.Execute "UPDATE Players Set Parts = Parts + " & rst!FailParts & " WHERE PlayerID = " & player.ID
            End If
         End If
         
         If rst!FailGoods = -99 Then 'goods seized not in Stash
            'allow for stash modifiers.  reduce by 4+mods
            If SeizeAllContraCargo(player.ID) Then   'this is a compromise - todo - rework for ALL Goods
               PutMsg player.PlayName & "'s Nav log: lost some Contraband/Cargo not in Stash", player.ID, Logic!Gamecntr, True, getLeader()
            End If
            
         ElseIf rst!FailGoods <> 0 Then ' could be -neg
            Set frmSalvage = New frmSalvaging
            frmSalvage.mode = IIf(rst!FailGoods > 0, 1, 4) 'add/discard
            frmSalvage.salvageCount = Abs(rst!FailGoods)
            frmSalvage.Show 1

         End If
         
         If rst!FailShipUpgrade <> 0 Then
            If getShipUpgrades(player.ID) > 0 Then
               'present list of player upgrades to discard one
               Set frmShUp = New frmShipUpgd
               frmShUp.discardMode = 1
               frmShUp.Show 1
            End If
         End If
         If rst!FailKillCrew <> 0 Then
            x = doKillCrews(player.ID, rst!FailKillCrew)
            If rst!OptionName = "If we're very lucky" And hasShipUpgrade(player.ID, 18) > 0 And x > 0 Then
               doDiscardGear player.ID, hasShipUpgrade(player.ID, 18)
               PutMsg player.PlayName & " lost the Reaver-Flage upgrade in the Reaver skuffle", player.ID, Logic!Gamecntr
            End If
         End If
         
         If rst!FailNestedTest > 0 Then 'go all Inception on its ass
            doNav CardID, rst!FailNestedTest
         End If
            
            
      End Select
       
      'DO the tests that run whatever the above outcome -----------------------------------------
      SectorID = varDLookup("SectorID", "Players", "PlayerID=" & player.ID)
       
      If rst!Disgruntled <> 0 Then 'apply disgruntled changes
         doDisgruntled player.ID, rst!Disgruntled
      End If
      
      Select Case rst!MoveReaver
         Case 1   ' 1 - move 1
            If Logic!AutoAI = 0 Then
               setPlayer player.ID, "X", 1
               If SoloGame Then
                  actionSeq = ASNavReav
               Else
                  actionSeq = ASNavReavEnd
               End If
            
            Else
               moveAutoAI 6 + RollDice(NumOfReavers)
      
            End If
            
         Case 2    '2-you move reaver to any B zone,
            MessBox "Move a Reaver to any Rim or Border sector", "Reavers on the Move", "OK", "", getLeader()
            actionSeq = ASNavReavBorder
            
         Case 3    '3-move to your location  (evade done later)
            If getCutterSector(SectorID) = 0 Then
               MoveShip 6 + RollDice(NumOfReavers), SectorID
            End If
            
         Case 4  'other player move reaver to any B zone,
            If Logic!AutoAI = 0 Then
               setPlayer player.ID, "W", 1
               If SoloGame Then
                  MessBox "Move a Reaver to any Rim or Border sector", "Reavers on the Move", "OK", "", getLeader()
                  actionSeq = ASNavReavBorder
               Else
                  actionSeq = ASNavReavEnd
               End If
            
            Else
               doMoveCutterPlanetary 6 + RollDice(NumOfReavers)
      
            End If
      End Select
      
      Select Case rst!MoveAlliance
         Case 1   ' 1 - move 1
             If Logic!AutoAI = 0 Then
               setPlayer player.ID, "Y", 1
               If SoloGame Then
                  MessBox "Move the Alliance Cruiser one sector", "Cruiser on the Move", "OK", "", getLeader()
                  actionSeq = ASNavCrus
               Else
                  actionSeq = ASNavCrusEnd
               End If
             
            Else
               moveAutoAI 5
      
            End If
            
         Case 2   '2- move to any
            MessBox "Move the Alliance Crusier to any Alliance sector not occupied by a Firefly", "Wild Gosling Chase", "OK", "", getLeader()
            actionSeq = ASNavCrusBorder
            
         Case 3   '3-move to outlaw ship
            If outlawExists(player.ID) Then
               MessBox "Move the Crusier to a sector with a rival Outlaw Ship", "A Legitimate Tip", "OK", "", getLeader()
               actionSeq = ASNavCrusOutlaw
            End If
         Case 4 'alliance pays you a visit
            'for each Wanted Crew: 1-Remove Crew, 2+ Crew safe
            'may use Cry Baby - or other modifiers? eg. Concealed Smuggling Compartments
            If doMoveAlliance(player.ID, SectorID) Then
               CruiserCutter = SectorID 'set it as faced
            Else
               frmAction.fullburndone = False
            End If
         
         Case 5 'move adjacent if failed
            If result = 2 Then
               If Logic!AutoAI = 0 And doMoveAllianceAdjacent(SectorID, True) Then  'there is a valid solution
                  setPlayer player.ID, "Z", 1
                  If SoloGame Then
                    MessBox "Move the Alliance Cruiser adjacent your Ship", "Cruiser on the Move", "OK", "", getLeader()
                    actionSeq = ASNavCrusAdjacent
                  Else
                    actionSeq = ASNavCrusEnd
                  End If
               
               Else
                  doMoveAllianceAdjacent SectorID
                 
               End If
            End If
            
         Case 6 'alert tokens adjacent your posn
            doAddTokensAdjacent SectorID
            RefreshBoard
         Case 7 'corvette contact
            Set frmSeize = New frmSeized
            If frmSeize.RefreshList(True) > 0 Then 'some are not stashed
               frmSeize.RefreshList False
               frmSeize.Show 1
            End If
            If SeizeAllFugi(player.ID) Then
               PutMsg player.PlayName & " lost some Fugitives not in Stash", player.ID, Logic!Gamecntr, True, getLeader()
            End If
         
         Case 8 'discard 1 crew
            Set frmSeize = New frmSeized
            frmSeize.Caption = "Select the Crew Member detained by the Alliance"
            If frmSeize.RefreshDiscardList() > 0 Then 'crew exist
               frmSeize.Show 1
            End If
         
         Case 9 'alert tokens at every Outlaw Ship
            doAddTokensOutlaws
            If isOutlaw(player.ID) Then ignoreToken = SectorID 'so as to not trip on one put here
            RefreshBoard
         Case 10 ' Move Corvette Adjacent player
            If Logic!AutoAI = 0 And doMoveCorvetteAdjacent(SectorID, True) Then
               setPlayer player.ID, "V", 1
               If SoloGame Then
                 MessBox "Move the Operative's Corvette adjacent your Ship", "Corvette on the Move", "OK", "", getLeader()
                 actionSeq = ASNavCorvAdjacent
               Else
                 actionSeq = ASNavCrusEnd
               End If
            
            Else
               doMoveCorvetteAdjacent SectorID
              
            End If
            
         Case 11  'Corvette to an unoccupied Alliance, Border, or Rim Planetary Sector.
            If Logic!AutoAI = 0 Then
               setPlayer player.ID, "U", 1
               If SoloGame Then
                 MessBox "Move the Operative's Corvette to a Planetary Sector", "Corvette on the Move", "OK", "", getLeader()
                 actionSeq = ASNavCorvPlanetary
               Else
                 actionSeq = ASNavCrusEnd
               End If
            
            Else
               doMoveCorvettePlanetary
            End If
            
         Case 12  'move Operative's Corvette 1 or 2 Sectors within Alliance, Border or Rim Space
            y = getCorvetteSector
            moveAutoCorvette2 0, False, y
                  
      End Select
      
      If rst!MovePlayer > 0 Then
         For x = 1 To rst!MovePlayer
            moveAutoAI player.ID, 1, True
            drawLine 0, -2, getPlayerSector(player.ID)
         Next x
      End If
      
      
      If rst!trader <> 0 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < 0.5 Then
            MessBox "You have no spare cargo capacity", "Trading on the go", "Ooops", "", getLeader()
         Else
            ' enable 1&2 Trader modes
            Set frmBart = New frmBarter
            frmBart.trader = rst!trader
            frmBart.Show 1
         End If
      End If
      
      If rst!KillAllPassFugi <> 0 Then
          DB.Execute "UPDATE Players SET Fugitive = 0, Passenger = 0 WHERE PlayerID = " & player.ID
      End If
      
      If rst!SeizeGoods = 1 Then 'Contraband and Fugitives not in your Stash are seized. Full Stop
         'allow for stash modifiers.  reduce by 4+mods
         If SeizeAllContraFugi(player.ID) Then
            PutMsg player.PlayName & "'s Nav log: lost some Contraband/Fugitives not in Stash", player.ID, Logic!Gamecntr, True, getLeader()
         End If
      End If
      
      If rst!SeizeGoods = 2 Then 'Contraband and Cargo not in your Stash are seized. Full Stop
         'allow for stash modifiers.  reduce by 4+mods
         If SeizeAllContraCargo(player.ID) Then
            PutMsg player.PlayName & "'s Nav log: lost some Contraband/Cargo not in Stash", player.ID, Logic!Gamecntr, True, getLeader()
         End If
      End If
      
      If rst!Warrant = 1 Or (result = 2 And rst!Warrant = 2) Or (isOutlaw(player.ID) And rst!Warrant = 3) Then
         If Not warrantDodge(player.ID) Then
            PutMsg player.PlayName & "'s Nav log: a Warrant has been issued" & IIf(isSolid(player.ID, 5), " and you are no longer Solid with Harken", "") & "!", player.ID, Logic!Gamecntr, True, getLeader()
            'add a Warrant and clear any Solid with Harken (5)
            DB.Execute "UPDATE Players SET Warrants = Warrants + 1" & IIf(discardRoberta(player.ID), "", ", Solid5 = 0") & " WHERE PlayerID = " & player.ID
         End If
      End If
      If rst!Warrant = -1 Then
         'clear Warrants
         DB.Execute "UPDATE Players SET Warrants = 0 WHERE PlayerID = " & player.ID
         PutMsg player.PlayName & "'s Nav log: any Warrants have been cleared!", player.ID, Logic!Gamecntr, True, getLeader()
      End If
      
      If rst!Token = 1 Then
         changeToken SectorID, 1
         ignoreToken = SectorID
      ElseIf rst!Token = 2 Then
         changeAToken SectorID, 1
         ignoreToken = SectorID
      End If
      
      If rst!Evade = 1 Or (result = 0 And rst!Evade = 2) Then
         PutMsg player.PlayName & "'s Nav log: EVADE!", player.ID, Logic!Gamecntr, True, getLeader()
         actionSeq = ASNavEvade
      End If
                  
   Else
      MsgBox "Error: Nav Card " & CardID & " Option " & opt & " not found!", vbCritical
   End If

Set rst = Nothing

End Function

'save selected (Seq=6 + selected) to players Jobs, unselected back to 5 DISCARDED
Public Function doDeal(ByVal playerID As Integer) As Integer
Dim Index As Integer
   With frmDeal.sftTree
      
      For Index = 0 To .ListCount - 1
         Select Case .ItemDataString(Index)
         Case "R"  'selected
            doDeal = doDeal + 1
            assignDeal playerID, .ItemData(Index)

         Case "UN" 'place back in discard (5)
            DB.Execute "UPDATE ContactDeck SET Seq =" & CStr(DISCARDED) & " WHERE CardID = " & .ItemData(Index)
            .ItemDataString(Index) = "O"
            Set .ItemPicture(Index) = frmDeal.AssetImages.Overlay("L", "O")
         End Select
      Next Index
   
   End With
End Function

'save selected (Seq=6 + selected) to players Jobs, unselected back to 5 DISCARDED
Public Function doBuy(ByVal playerID As Integer) As Integer
Dim Index As Integer, cost As Integer, imposter As Integer
   With frmBuy.sftTree
      cost = 0
      For Index = 0 To .ListCount - 1
         Select Case .ItemDataString(Index)
         Case "R"  'selected -pay up!!
            If .ItemData(Index) = 28 Then 'If Saffron is hired by anyone, Remove the existing imposter from Play
               If haveCrewAnyone(54) Then
                  doDiscardCrew 100
                  imposter = 54
               ElseIf haveCrewAnyone(41) Then
                  doDiscardCrew 70
                  imposter = 41
               End If
            End If
            
            If .ItemData(Index) = 100 Then 'If Bridgit is hired by anyone, Remove the existing imposter from Play
               If haveCrewAnyone(23) Then
                  doDiscardCrew 28
                  imposter = 23
               ElseIf haveCrewAnyone(41) Then
                  doDiscardCrew 70
                  imposter = 41
               End If
            End If
            
            If .ItemData(Index) = 70 Then 'If Yolonda is hired by anyone, Remove the existing imposter from Play
               If haveCrewAnyone(54) Then
                  doDiscardCrew 100
                  imposter = 54
               ElseIf haveCrewAnyone(23) Then
                  doDiscardCrew 28
                  imposter = 23
               End If
            End If
            
            If .CellItemData(Index, 2) = 1 Then
               'if buying a Drive Core, swap out the existing one
               removeDriveCore player.ID
            End If
            doBuy = doBuy + 1

            cost = cost + .CellItemData(Index, 8)

            DB.Execute "UPDATE SupplyDeck SET Seq =" & playerID & " WHERE CardID = " & .ItemData(Index)
            'add the card to the players deck
            DB.Execute "INSERT INTO PlayerSupplies (PlayerID, CardID) VALUES (" & playerID & ", " & .ItemData(Index) & ")"
         
            If imposter > 0 Then
               PutMsg getCrewName(0, imposter) & " has turned up as " & getCrewName(.ItemData(Index)) & " on " & player.PlayName & "'s Ship", playerID, Logic!Gamecntr, True, getCrewID(.ItemData(Index)), 0, 0, 0, 1
            End If
         
         Case "UN" 'place back in discard (5)
            DB.Execute "UPDATE SupplyDeck SET Seq =" & CStr(DISCARDED) & " WHERE CardID = " & .ItemData(Index)
            .ItemDataString(Index) = "O"
            Set .ItemPicture(Index) = frmBuy.AssetImages.Overlay("L", "O")
         End Select
      Next Index
      If cost > 0 Then
         DB.Execute "UPDATE Players SET Pay=Pay - " & cost & " WHERE PlayerID = " & playerID
      End If
   
   End With
End Function

Private Function getNewPlayer() As Integer
Dim frmplayer As New frmSelPlayer
   frmplayer.Show 1
   getNewPlayer = frmplayer.playerID

End Function

Private Sub checkFlacGun(ByVal SectorID)
Dim x, g

   x = getCutterSector(SectorID)
   If x > 0 Then 'we got company!
      g = hasShipUpgrade(player.ID, 15)
      If g > 0 Then 'Flac Gun
         If MessBox("Reaver within firing Range, do you want to use the single-use Flac Gun to fend it off?", "Reaver Cutter", "Yes", "No", 0, 0, 15) = 0 Then
            doDiscardGear player.ID, g
            moveAutoAI x
            PutMsg player.PlayName & " depleted their Hull-Mounted Flak Gun to fend off a Reaver", player.ID, Logic!Gamecntr
         End If
      End If
   End If
   

End Sub

Private Sub checkBigBlack(ByVal CardID)

   If varDLookup("CardName", "NavDeck", "CardID=" & CardID) = "The Big Black" Then
      TheBigBlack = TheBigBlack + 1
      If TheBigBlack = 2 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0 Then 'got room
            If MessBox("Your Emissions Recycler is working well in the Big Black." & vbNewLine & "Do you want to recover one Fuel?", "Emissions Recycler", "Yes", "No", 0, 0, 20) = 0 Then
               DB.Execute "Update Players Set Fuel = Fuel + 1 WHERE PlayerID =" & player.ID
            End If
         End If
         TheBigBlack = -1 'disable
      End If
   Else
      TheBigBlack = 0 'reset the count
   End If
   
End Sub

Public Function doKillCrews(ByVal playerID, ByVal NoToKill As Integer, Optional ByVal onJobOnly As Boolean = False) As Integer
Dim tracey As Integer, crewCount As Integer, killed As Integer
Dim frmKillCrw As frmKillCrew
      tracey = 0
      'Tracey must be Killed first "KillFirst" perk CardID 4, CrewID 11
      If hasCrew(playerID, 11) Then
         killed = doKillCrew(playerID, 4)
         tracey = 1
      End If
      
      If NoToKill - tracey > 0 Then
         Set frmKillCrw = New frmKillCrew
         crewCount = getCrewCount(playerID, onJobOnly)
         If crewCount >= (NoToKill - tracey) Then 'more or equal crew than to be killed
            crewCount = (NoToKill - tracey)
            
         End If
         If crewCount > 0 Then
            frmKillCrw.nbrSelect = crewCount
            frmKillCrw.Show 1
            killed = frmKillCrw.killed
         End If

      End If
      doKillCrews = killed
      If killed > 0 Then PutMsg player.PlayName & " gets " & CStr(killed) & " Crew killed", playerID, Logic!Gamecntr
End Function

Private Sub doSlaveTrade(ByVal TraderID)
Dim frmTrade As New frmTrader
   frmTrade.TraderID = Logic!player
   frmTrade.lblTitle(1).Caption = PlayCode(TraderID).PlayName & "'s Trade Items"
   frmTrade.Show 1
   If Not (frmShip Is Nothing) Then frmShip.RefreshShips
End Sub


Private Sub Verse_SectClick(ByVal Index As Integer)
Dim Havens As Boolean

   'picking starting sector
   If pickStartSector = 1 Then
      Havens = useHavens(Logic!StoryID)
      If Not CheckClash(player.ID, Index, Havens) Then
         If Havens Then placeHaven player.ID, Index
         MoveShip player.ID, Index
         pickStartSector = 2  'flag the selection is done
      End If
   End If
   
   If actionSeq = ASmosey Then
      If validMove(player.ID, Index, True) Then
         frmAction.cmd(0).Enabled = False
         'get players current posn and check route
         MoveShip player.ID, Index, 7
         MoseyMovesDone = MoseyMovesDone + 1
         drawLine 0, -2, Index
         wormHoleOpen = False
         drawLine 2, -1
         actionSeq = ASMoseyEnd 'throw to main loop
       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASfullburn Then
      If validMove(player.ID, Index, hasShipUpgrade(player.ID, 18)) Then
         frmAction.cmd(1).Enabled = False
         MoveShip player.ID, Index
         FullburnMovesDone = FullburnMovesDone + 1
         If FullburnMovesDone = 1 And Val(frmAction.lblFuelRq.Caption) > 0 Then burnFuel player.ID, Val(frmAction.lblFuelRq.Caption)
         If HemmorrhagingFuel Then burnFuel player.ID, 1
         drawLine 0, -2, Index
         wormHoleOpen = False
         drawLine 2, -1
         actionSeq = ASFullburnEnd 'throw to main loop
       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavEvade Then
      If validMove(player.ID, Index) Then
         'if evading a reaver at the beginning of turn, then don't stop fullburn
         If FullburnMovesDone > 0 Then frmAction.fullburndone = True
         MoveShip player.ID, Index
         drawLine 0, -2, Index
         actionSeq = ASNavEvadeEnd
         CruiserCutter = 0
         CorvetteSeq = 0
         ignoreToken = 0
       Else
         playsnd 9
      End If
   End If
   
   'move reaver one space - manual option
   If actionSeq = ASNavReav Then
      If Logic!player = player.ID Then
         If reaverMove(Index) Then actionSeq = ASNavReavEnd
      End If
   End If
   
   If actionSeq = ASNavReavBorder Then
      If Logic!player = player.ID And (getClearSector(Index) = "B" Or getClearSector(Index) = "R") Then
         MoveShip 6 + RollDice(NumOfReavers), Index
         actionSeq = ASNavReavEnd
       Else
         playsnd 9
      End If

   End If
   
   'move cruiser one space - manual option
   If actionSeq = ASNavCrus Then
      If validMove(5, Index) And Logic!player = player.ID And Not getHaven(Index) > 0 Then
         MoveShip 5, Index
         actionSeq = ASNavCrusEnd
       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCrusBorder Then
      If Logic!player = player.ID And getClearSector(Index) = "A" And getCruiserSector() <> Index And Not getHaven(Index) > 0 Then
         MoveShip 5, Index
         actionSeq = ASNavCrusEnd
      Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCrusOutlaw Then
      If Logic!player = player.ID And outlawExists(player.ID) And getCruiserSector() <> Index And Not getHaven(Index) > 0 Then
         MoveShip 5, Index
         actionSeq = ASNavCrusEnd
      Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCrusAdjacent Then
      If Logic!player = player.ID And getClearSector(Index) = "A" And Not getHaven(Index) > 0 Then
         MoveShip 5, Index
         actionSeq = ASNavCrusEnd
      Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCorvAdjacent Then
      If Logic!player = player.ID And getClearSector(Index) <> "" Then
         MoveShip 6, Index
         actionSeq = ASNavCrusEnd
       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCorvPlanetary Then
      If Logic!player = player.ID And getClearSector(Index) <> "" And Nz(varDLookup("PlanetID", "Planet", "SectorID=" & Index), 0) > 0 Then
         MoveShip 6, Index
         actionSeq = ASNavCrusEnd
       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASResolveAlert Then
      If isAdjacent(player.ID, Index) Then
         resolveToken Index, True
         actionSeq = ASResolveAlertEnd
       Else
         playsnd 9
      End If
   End If
   
   
End Sub

Public Sub drawLine(ByVal mode, ByVal sector1, Optional ByVal sector2, Optional ByVal silent As Boolean = True)
Dim rst As New ADODB.Recordset
Dim SQL, X1, X2, Y1, Y2

   If sector1 = -1 Then
      Verse.LineB(mode).Visible = False
      Exit Sub
   End If
   If sector1 = -2 And Verse.LineB(mode).Visible = False Then Exit Sub
   
   If sector1 = 1 Then
      sector1 = getCruiserSector()
   ElseIf sector1 = 2 Then
      sector1 = getCorvetteSector()
   End If
   
   If sector1 = -2 Then
      X1 = Verse.LineB(mode).X1
      Y1 = Verse.LineB(mode).Y1
   Else
      SQL = "SELECT * "
      SQL = SQL & "FROM Board WHERE SectorID=" & sector1
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         X1 = rst!SLeft + Int(rst!SWidth / 2)
         Y1 = rst!STop + Int(rst!SHeight / 2)
      End If
      rst.Close
   
   End If
   
   SQL = "SELECT * "
   SQL = SQL & "FROM Board WHERE SectorID=" & sector2
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       X2 = rst!SLeft + Int(rst!SWidth / 2)
       Y2 = rst!STop + Int(rst!SHeight / 2)
   End If
   rst.Close
      
   If Verse.LineB(mode).X1 = X1 And Verse.LineB(mode).Y1 = Y1 And Verse.LineB(mode).X2 = X2 And Verse.LineB(mode).Y2 = Y2 And Verse.LineB(mode).Visible = True Then
      Verse.LineB(mode).Visible = False
   Else
      Verse.LineB(mode).X1 = X1
      Verse.LineB(mode).Y1 = Y1
      Verse.LineB(mode).X2 = X2
      Verse.LineB(mode).Y2 = Y2
      Verse.LineB(mode).Visible = True
      'Verse.LineB(mode).ZOrder
      If Not silent Then playsnd 2
   End If
   
   Set rst = Nothing
End Sub

Private Sub animatePlayer(ByVal playerID)
Dim x
   For x = 1 To 4
      If x = playerID Then
         If Verse.Imag(x).Animate2.AnimationState = lvicAniCmdStop Then
            Verse.Imag(x).Animate2.StartAnimation
         End If
      Else
         If Verse.Imag(x).Animate2.AnimationState = lvicAniCmdStart Then
            Verse.Imag(x).Animate2.StopAnimation
            Verse.Imag(x).ImageIndex = 1
         End If
      End If
   Next x
End Sub
