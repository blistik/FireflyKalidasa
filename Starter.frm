VERSION 5.00
Begin VB.Form Starter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Waiting Room"
   ClientHeight    =   4725
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5760
   Icon            =   "Starter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Starter.frx":0442
   ScaleHeight     =   4725
   ScaleWidth      =   5760
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5200
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Edit this Story Card"
      Top             =   3150
      Width           =   375
   End
   Begin VB.CheckBox chkAI 
      BackColor       =   &H00CBE1ED&
      Caption         =   "auto move Crusier, Corvette && Cutters"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2100
      TabIndex        =   14
      Top             =   2190
      Value           =   1  'Checked
      Width           =   3105
   End
   Begin VB.ComboBox cbo 
      BackColor       =   &H00CBE1ED&
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3150
      Width           =   5085
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timing 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5160
      Top             =   2520
   End
   Begin VB.ListBox Lst 
      BackColor       =   &H00CBE1ED&
      Height          =   1230
      Left            =   2070
      TabIndex        =   8
      Top             =   780
      Width           =   3345
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Join"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Choose a Firefly"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Tag             =   "Orange"
      Top             =   600
      Width           =   1755
      Begin VB.OptionButton opt 
         BackColor       =   &H0000C000&
         Caption         =   "Green"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Tag             =   "Green"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H0000FFFF&
         Caption         =   "Yellow"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Tag             =   "Yellow"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00FF0000&
         Caption         =   "Blue"
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Tag             =   "Blue"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000080FF&
         Caption         =   "Orange"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Tag             =   "Orange"
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblStory 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   1065
      Left            =   120
      TabIndex        =   13
      Top             =   3570
      Width           =   5475
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Story"
      Height          =   255
      Left            =   210
      TabIndex        =   12
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Players"
      Height          =   225
      Left            =   2070
      TabIndex        =   11
      Top             =   570
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player Name"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Starter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isHost As Boolean, started As Boolean

Private Sub cbo_Click()
    
   If isHost Then
      Logic.Update "StoryID", GetCombo(cbo)
'      If GetCombo(cbo) > 0 Then 'custom story
'         If doCustomStory = 0 Then
'            LoadCombo cbo, "story", " WHERE ACTIVE = 1 Order by StoryID"
'            cbotrig = True
'            SetCombo cbo, "", 1
'            cbotrig = False
'         Else
'            cbo.List(cbo.ListIndex) = Nz(varDLookup("StoryTitle", "Story", "StoryID = " & GetCombo(cbo)))
'         End If
'      End If
   End If
   lblStory.Caption = Nz(varDLookup("StoryDesc", "Story", "StoryID = " & GetCombo(cbo)))
End Sub

Private Sub chkAI_Click()
   Logic.Update "AutoAI", chkAI.Value
End Sub

Private Sub cmd_Click(Index As Integer)
Dim rst As New ADODB.Recordset, col, cnt, x
Dim frmCrew As frmCrewSel, leader, nextplayer As Integer, noOfCrew As Integer, costLimit As Integer, randCrew As Integer
Dim frmCrewList As frmCrewLst
         
   On Error GoTo err_handler
   playsnd 8
   UpdateLst
   
   Select Case Index
   Case 0 'join
      If cmd(0).Caption = "Pick Leader" Then
         Timing.Enabled = False
         'MsgBox "Pick a LEADER"
         Set frmCrew = New frmCrewSel
         frmCrew.crewFilter = " WHERE Leader = 1 AND NOT EXISTS(select 1 from Players WHERE Leader = CrewID)" & IIf(getExcludeCrew = "", "", " AND CrewID NOT IN (" & getExcludeCrew & ")") & " Order By CrewName"
         frmCrew.Show 1
         leader = GetCombo(frmCrew.cboCrew)
         PutMsg player.PlayName & " has chosen " & frmCrew.cboCrew.Text, player.ID
         cmd(0).Enabled = False
         cmd(0).Caption = "waiting"
         SetupPlayer player.ID, Logic!StoryID
         'drop this leaders Card into the Player's supplies
         DB.Execute "INSERT INTO PlayerSupplies (PlayerID,CardID) VALUES (" & player.ID & ", " & varDLookup("CardID", "SupplyDeck", "CrewID =" & leader) & ")"
         
         'get story requirements
         noOfCrew = varDLookup("StartingCrew", "Story", "StoryID=" & Logic!StoryID)
         costLimit = varDLookup("CrewCostLimit", "Story", "StoryID=" & Logic!StoryID)
         randCrew = varDLookup("RandomCrew", "Story", "StoryID=" & Logic!StoryID)
         
         If noOfCrew > 0 And randCrew = 1 Then
            getRandomCrew noOfCrew, leader
         
         ElseIf noOfCrew > 0 Then
           'show a list of Crew to choose up to $1000 hire fee
            Set frmCrewList = New frmCrewLst
            frmCrewList.crewFilter = getExcludeCrew
            frmCrewList.selectCrew = noOfCrew
            frmCrewList.costLimit = costLimit
            frmCrewList.Caption = "Select up to " & noOfCrew & " Crew up to $" & costLimit
            frmCrewList.Show 1
         End If

         'cycle players, and set "S" to start when done all
         setNextLeader player.ID, leader
         
         Timing.Enabled = True
      Else
         For x = 0 To 3
            If opt(x).Value Then
               col = Left(opt(x).Tag, 1)
               Exit For
            End If
            'Opt(x).Enabled = False
         Next x
         If IsEmpty(col) Or txt.Text = "" Then Exit Sub
         rst.Open "SELECT * FROM Players WHERE Colour = '" & col & "' ORDER BY PlayerID", DB, adOpenDynamic, adLockOptimistic
         If IsNull(rst!Name) Then
            rst.Update "Name", txt.Text
            player.ID = rst!playerID
            player.PlayName = txt.Text
            For x = 0 To 3
               opt(x).Enabled = False
            Next x
            UpdateLst
            cmd(0).Enabled = False
            cmd(0).Caption = "Pick Leader"
            'show start button
            If isHost Then
               If Not cmd(1).Visible Then playsnd 6
               cmd(1).Visible = True
            End If
            
         Else
            For x = 0 To 3
               If col = Left(opt(x).Tag, 1) Then
                  opt(x).Enabled = False
               End If
            Next x
            UpdateLst
            MessBox rst!ship & " is taken by " & rst!Name, "Ship taken", "Ooops", "", 0, 0, 6
            
         End If
      End If
   Case 1 'start
      'lock the story & start button
      Logic.Update "StoryID", GetCombo(cbo)
      cbo.Enabled = False
      cmd(1).Enabled = False
      'shuffle the decks
      PutMsg "Decks are Shuffled"
      ShuffleDeck "Contact", True
      ShuffleDeck "Supply", True
      'exclude Crew per Story
      doExcludes
      DrawDeck "Supply", 1, 3
      DrawDeck "Supply", 2, 3
      DrawDeck "Supply", 3, 3
      DrawDeck "Supply", 4, 3
      DrawDeck "Supply", 5, 3
      DrawDeck "Supply", 6, 3
      DrawDeck "Supply", 7, 3
      ShuffleDeck "Nav", True, (Lst.ListCount > 2) 'Reshuffle Cards at end for 3 or more players
      ShuffleDeck "Misbehave", False, True
      
      'do story specific setup
      If Logic!StoryID = 12 Then 'take Wash out of Deck
         DB.Execute "UPDATE SupplyDeck SET Seq = 0  WHERE CardID = 1"
      End If
      
      'show who has entered game
      cnt = 0
      For x = 1 To 4
         If PlayCode(x).PlayName <> "" Then
            cnt = cnt + 1
            PutMsg PlayCode(x).PlayName & " has entered the game", x
         End If
      Next x
      SoloGame = isSoloGame()
               
      nextplayer = 0
      Randomize Timer
      Do
          x = Int((4 * Rnd)) + 1
          If PlayCode(x).PlayName <> "" Then nextplayer = x
      Loop While nextplayer = 0
         
      Logic!Seq = "L"  'pick leader
      Logic!player = nextplayer
      Logic!AutoAI = chkAI.Value
      Logic.Update
      If nextplayer = player.ID Then 'your go to pick leader
         cmd(0).Enabled = True
      End If
      chkAI.Enabled = False
      'started = True
      'Unload Me
   
   Case 2 'edit story
      If GetCombo(cbo) > 0 Then 'custom story
         Timing.Enabled = False
         If doCustomStory(Not isHost) = 0 Then 'deleted, reset back to 1
            LoadCombo cbo, "story", " WHERE ACTIVE = 1 Order by StoryID"
            SetCombo cbo, "", 1
         Else 'update the display
            cbo.List(cbo.ListIndex) = Nz(varDLookup("StoryTitle", "Story", "StoryID = " & GetCombo(cbo)))
            lblStory.Caption = Nz(varDLookup("StoryDesc", "Story", "StoryID = " & GetCombo(cbo)))
         End If
         Timing.Enabled = True
      End If
   
   End Select
  
  Exit Sub
  
err_handler:
  MsgBox Err.Description, vbCritical, "Error"
  UpdateLst
End Sub

Private Sub Form_Load()
  initForm
  UpdateLst
  Timing.Enabled = True
  player.ID = 0
  player.Color = ""
  player.PlayName = ""
  cmd(0).Enabled = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not started Then
      player.ID = 0
      player.Color = ""
      player.PlayName = ""
   End If
End Sub

Private Sub Timing_Timer()
   UpdateLst
End Sub

Private Sub initForm()
Dim rst As New ADODB.Recordset
   started = False
   rst.Open "SELECT * FROM Players WHERE PlayerID > 0 AND PlayerID < 5 ORDER BY PlayerID", DB, adOpenStatic, adLockReadOnly
   While Not rst.EOF
      opt(rst!playerID - 1).Caption = rst!ship
      rst.MoveNext
   Wend
   cbo.Enabled = isHost
   chkAI.Enabled = isHost
   LoadCombo cbo, "story", " WHERE ACTIVE = 1 Order by StoryID"
   SetCombo cbo, "", 1

End Sub

Private Sub UpdateLst()
Dim rst As New ADODB.Recordset, col, x, playerID As Integer
   Lst.Clear
   x = GetSeqX(playerID)
   If Not isHost Then 'client only processes
      Select Case x
      Case "H"  'in host mode, enable join
         If cmd(0).Caption <> "Pick Leader" Then
            If Not cmd(0).Enabled Then playsnd 6
            cmd(0).Enabled = True
         End If
      Case "E"
         'reset the form
         Lst.AddItem "..waiting for Host .."
         cmd(0).Enabled = False
         cmd(0).Caption = "Join"
         For x = 0 To 3
            opt(x).Enabled = True
         Next x
         Exit Sub
      Case "L"
         'pick leader
      Case Else  'joining game
         started = True
      End Select
   ElseIf x = "S" Then
      started = True
   End If

   If x = "L" And Val(player.ID) = playerID Then 'your go to pick leader
       cmd(0).Enabled = True
   End If

  rst.Open "SELECT Players.*,Crew.CrewName FROM Crew RIGHT JOIN Players ON Crew.CrewID = Players.Leader Where Players.Name Is Not Null And Players.PlayerID < 5 ORDER BY Players.PlayerID", DB, adOpenStatic, adLockReadOnly
  
  While Not rst.EOF
      col = rst!Colour
      Lst.AddItem rst!ship & "  -  " & rst!Name & "  -  " & rst!CrewName

      PlayCode(rst!playerID).PlayName = rst!Name
      'disable selected Ships, other than the one you have selected
      For x = 0 To 3
         If col = Left(opt(x).Tag, 1) And player.ID <> x + 1 Then
            opt(x).Enabled = False
            opt(x).Value = False
            Exit For
         End If
      Next x
      
      rst.MoveNext
  Wend
  
  rst.Close
  Lst.Refresh
  Logic.Requery
  
  If isHost Then
      If Logic!Seq = "H" And player.ID = 0 Then
         If Not cmd(0).Enabled Then playsnd 6
         cmd(0).Enabled = True
            
      End If
  Else
      SetCombo cbo, "", Logic!StoryID
      chkAI.Value = Logic!AutoAI
  End If
  If started Then Unload Me
End Sub

Private Sub doExcludes()
Dim excludes As String
   excludes = getExcludeCrew
   If excludes <> "" Then
      DB.Execute "UPDATE SupplyDeck Set Seq = 0 WHERE  CrewID IN (" & excludes & ")"
   End If
End Sub

Private Sub getRandomCrew(ByVal noOfCrew As Integer, ByVal leader)
Dim rst As New ADODB.Recordset, SQL, CrewID, maxCrewID, crewcnt

   maxCrewID = varDLookup("max(CrewID) AS maxcrew", "Crew", "Leader=0", "maxcrew")
   SQL = "SELECT SupplyDeck.CardID, SupplyDeck.Seq, Crew.* FROM Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID WHERE Crew.Leader=0 AND Seq > 4 AND Crew.CrewID NOT IN (23,54)"
   If leader = 69 Then 'add Atherton check
      SQL = SQL & " AND Crew.Companion = 0"
   End If
   crewcnt = 0
   rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
   While crewcnt < noOfCrew
      rst.Requery
      CrewID = Int((maxCrewID * Rnd)) + 1
      rst.filter = "CrewID =" & CrewID
      If Not rst.EOF Then
          DB.Execute "UPDATE SupplyDeck SET Seq =" & player.ID & " WHERE CardID = " & rst!CardID
          'add the card to the players deck
          DB.Execute "INSERT INTO PlayerSupplies (PlayerID, CardID) VALUES (" & player.ID & ", " & rst!CardID & ")"
         rst.Update "Seq", player.ID
         crewcnt = crewcnt + 1
      End If
   Wend
   rst.Close
End Sub

