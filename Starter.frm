VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form Starter 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Waiting Room"
   ClientHeight    =   4770
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7350
   Icon            =   "Starter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ComboTxt 
      BackColor       =   &H003DCBFF&
      Height          =   315
      Left            =   2340
      TabIndex        =   0
      Top             =   90
      Width           =   1485
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "add a bot player"
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
      Index           =   3
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2490
      Visible         =   0   'False
      Width           =   1905
   End
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Edit this Story Card"
      Top             =   3090
      Width           =   375
   End
   Begin VB.CheckBox chkAI 
      Appearance      =   0  'Flat
      BackColor       =   &H00343644&
      Caption         =   "auto move Crusier, Corvette && Cutters"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   3000
      TabIndex        =   7
      Top             =   2130
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin VB.ComboBox cbo 
      BackColor       =   &H001F2025&
      Enabled         =   0   'False
      ForeColor       =   &H003DCBFF&
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3090
      Width           =   5100
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
      Left            =   5260
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timing 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6060
      Top             =   2460
   End
   Begin VB.ListBox Lst 
      BackColor       =   &H001F2025&
      ForeColor       =   &H003DCBFF&
      Height          =   1230
      Left            =   2970
      TabIndex        =   13
      Top             =   740
      Width           =   3495
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
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H001F2025&
      Caption         =   "Choose a Firefly"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   2175
      Left            =   1020
      TabIndex        =   12
      Tag             =   "Orange"
      Top             =   540
      Width           =   1845
      Begin VB.OptionButton opt 
         BackColor       =   &H0000C000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Tag             =   "Green"
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Tag             =   "Yellow"
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Tag             =   "Blue"
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H000080FF&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Tag             =   "Orange"
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgShip 
         Height          =   375
         Index           =   3
         Left            =   510
         Top             =   1680
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Effects         =   "Starter.frx":0442
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgShip 
         Height          =   375
         Index           =   2
         Left            =   510
         Top             =   1200
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Effects         =   "Starter.frx":045A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgShip 
         Height          =   375
         Index           =   1
         Left            =   510
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Effects         =   "Starter.frx":0472
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgShip 
         Height          =   375
         Index           =   0
         Left            =   510
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Effects         =   "Starter.frx":048A
      End
   End
   Begin VB.Label lblStory 
      BackColor       =   &H001F2025&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H003DCBFF&
      Height          =   1065
      Left            =   1020
      TabIndex        =   16
      Top             =   3510
      Width           =   5475
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Story"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   255
      Left            =   1040
      TabIndex        =   15
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Players"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   225
      Left            =   2970
      TabIndex        =   14
      Top             =   540
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player Name"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   255
      Left            =   1100
      TabIndex        =   11
      Top             =   180
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
      DB.Execute "UPDATE GameSeq set StoryID = " & CStr(GetCombo(cbo))
      cmd(3).Visible = Nz(varDLookup("AIBotEnabled", "Story", "StoryID = " & GetCombo(cbo)), 0)
   End If
   
   lblStory.Caption = Nz(varDLookup("StoryDesc", "Story", "StoryID = " & GetCombo(cbo)))
End Sub

Private Sub chkAI_Click()
   Logic.Update "AutoAI", chkAI.Value
End Sub

Private Sub cmd_Click(Index As Integer)
Dim rst As New ADODB.Recordset, col, cnt, X
Dim frmCrew As frmCrewSel, leader, nextplayer As Integer, noOfCrew As Integer, costLimit As Integer
Dim randCrew As Integer, forceFugi As Integer
Dim frmCrewList As frmCrewLst
Dim supplyInit, contactInit
         
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
         Logic.Requery
         SetupPlayer player.ID, Logic!StoryID
         'drop this leaders Card into the Player's supplies
         DB.Execute "INSERT INTO PlayerSupplies (PlayerID,CardID) VALUES (" & player.ID & ", " & varDLookup("CardID", "SupplyDeck", "CrewID =" & leader) & ")"
         
         'get story requirements
         noOfCrew = varDLookup("StartingCrew", "Story", "StoryID=" & Logic!StoryID)
         costLimit = varDLookup("CrewCostLimit", "Story", "StoryID=" & Logic!StoryID)
         randCrew = varDLookup("RandomCrew", "Story", "StoryID=" & Logic!StoryID)
         forceFugi = varDLookup("Fugitives", "Story", "StoryID=" & Logic!StoryID)
         
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
         
         If forceFugi > 0 Then
            doForceFugitives
         End If
         
         'cycle players, and set "S" to start when done all
         setNextLeader player.ID, leader
         
         Timing.Enabled = True
      Else
         For X = 0 To 3
            If opt(X).Value Then
               col = Left(opt(X).Tag, 1)
               Exit For
            End If
            'Opt(x).Enabled = False
         Next X
         If IsEmpty(col) Or ComboTxt.Text = "" Then
            MessBox "Please enter your player name and select a Ship", "Joining", "Ooops", "", 0, 0, 0, 0, 3
            Exit Sub
         End If
         rst.CursorLocation = adUseClient
         rst.Open "SELECT * FROM Players WHERE Colour = '" & col & "' ORDER BY PlayerID", DB, adOpenStatic, adLockReadOnly
         If IsNull(rst!Name) Then
            'rst.Update "Name", txt.Text
            player.ID = rst!playerID
            player.PlayName = Trim(ComboTxt.Text)
            DB.Execute "UPDATE Players SET Name = '" & SQLFilter(player.PlayName) & "' WHERE PlayerID = " & CStr(player.ID)
            
            For X = 0 To 3
               opt(X).Enabled = False
            Next X
            UpdateLst
            cmd(0).Enabled = False
            cmd(0).Caption = "Pick Leader"
            'show start button
            If isHost Then
               If Not cmd(1).Visible Then playsnd 6
               cmd(1).Visible = True
            End If
            
         Else
            For X = 0 To 3
               If col = Left(opt(X).Tag, 1) Then
                  opt(X).Enabled = False
               End If
            Next X
            UpdateLst
            MessBox rst!ship & " is taken by " & player.PlayName, "Ship taken", "Ooops", "", 0, 0, 6
            
         End If
      End If
   Case 1 'start
      'lock the story & start button
      DB.Execute "UPDATE GameSeq set StoryID = " & GetCombo(cbo)
      Logic.Requery
      'Logic.Update "StoryID", GetCombo(cbo)
      cbo.Enabled = False
      cmd(1).Enabled = False
      cmd(3).Enabled = False
      'shuffle the decks
      
      ShuffleDeck "Contact", True
      contactInit = varDLookup("contactInit", "Story", "StoryID=" & Logic!StoryID)
      If contactInit > 0 Then
         For X = 1 To NO_OF_CONTACTS
            DrawDeck "Contact", X, contactInit
         Next X
      End If
      
      
      ShuffleDeck "Supply", True
      'exclude Crew per Story
      doExcludes
      supplyInit = varDLookup("supplyInit", "Story", "StoryID=" & Logic!StoryID)
      If supplyInit > 0 Then
         For X = 1 To NO_OF_SUPPLY
            DrawDeck "Supply", X, supplyInit
         Next X
      End If
      
      ShuffleDeck "Nav", True, (Lst.ListCount > 2) 'Reshuffle Cards at end for 3 or more players
      
      ShuffleDeck "Misbehave", False, True
      PutMsg "Decks are Shuffled"
      If isBountyEnabled Then
         DrawDeck "Contact", 10, 3
      End If
      
      'do story specific setup// - now done in doExcludes
'      If Logic!StoryID = 12 Then 'take Wash out of Deck
'         DB.Execute "UPDATE SupplyDeck SET Seq = 0  WHERE CardID = 1"
'      End If
      
      'show who has entered game
      cnt = 0
      For X = 1 To 4
         If PlayCode(X).PlayName <> "" Then
            cnt = cnt + 1
            PutMsg PlayCode(X).PlayName & " has entered the game", X
         End If
      Next X
      SoloGame = isSoloGame()
               
      nextplayer = 0
      Randomize Timer
      Do
          X = Int((4 * Rnd)) + 1
          If PlayCode(X).PlayName <> "" Then nextplayer = X
      Loop While nextplayer = 0
       
      'pick leader
      DB.Execute "UPDATE GameSeq SET Seq = 'L',Player = " & CStr(nextplayer) & ", AutoAI = " & CStr(chkAI.Value)
      'Logic!Seq = "L"
      'Logic!player = nextplayer
      'Logic!AutoAI = chkAI.Value
      'Logic.Update
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
   Case 3
      X = ShellExecute(X, "OPEN", App.Path & "\FireflyAIBot.exe ", datab, vbNullString, 1)                '1=normal, 2=min, 3=max, 4=behind
   End Select
  
  Exit Sub
  
err_handler:
  MsgBox Err.Description, vbCritical, "Setup Error"
  UpdateLst
End Sub

Private Sub Form_Load()
Dim lastplayer As String
  initForm
  UpdateLst
  Timing.Enabled = True
  player.ID = 0
  player.Color = ""
  player.PlayName = ""
  cmd(0).Enabled = False
  'cmd(3).Visible = isHost
  Set Me.Picture = LoadPicture(App.Path & "\pictures\waiting.jpg")
  lastplayer = getLastPlayer
  If lastplayer <> vbNullString Then
    ComboTxt.AddItem lastplayer
  End If
  ComboTxt.Text = GetSetting(App.ProductName, "Game", "PlayerName", vbNullString)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not started Then
      player.ID = 0
      player.Color = ""
      player.PlayName = ""
   End If
End Sub

Private Sub imgShip_Click(Index As Integer)
   If opt(Index).Enabled Then opt(Index).Value = True
End Sub

Private Sub Timing_Timer()
   UpdateLst
End Sub

Private Sub initForm()
Dim X As Integer  ' rst As New ADODB.Recordset
   started = False
   For X = 1 To 4
      imgShip(X - 1).Picture = LoadPictureGDIplus(App.Path & "\gui\FireflyShip" & CStr(X) & ".jpg")
   Next X
   'rst.CursorLocation = adUseClient
   'rst.Open "SELECT * FROM Players WHERE PlayerID > 0 AND PlayerID < 5 ORDER BY PlayerID", DB, adOpenStatic, adLockReadOnly
   'While Not rst.EOF
   '   opt(rst!playerID - 1).Caption = rst!ship
   '   rst.MoveNext
   'Wend
   cbo.Enabled = isHost
   chkAI.Enabled = isHost
   LoadCombo cbo, "story", " WHERE ACTIVE = 1 Order by StoryTitle"
   SetCombo cbo, "", 5

End Sub

Private Sub UpdateLst()
Dim rst As New ADODB.Recordset, X, playerID As Integer
   Lst.Clear
   X = GetSeqX(playerID)
   If Not isHost Then 'client only processes
      Select Case X
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
         For X = 0 To 3
            opt(X).Enabled = True
            imgShip(X).GrayScale = lvicNoGrayScale
         Next X
         Exit Sub
      Case "L"
         'pick leader
      Case Else  'joining game
         started = True
      End Select
   ElseIf X = "S" Then
      started = True
   End If

   If X = "L" And Val(player.ID) = playerID Then 'your go to pick leader
       cmd(0).Enabled = True
   End If
  rst.CursorLocation = adUseClient
  rst.Open "SELECT Players.*,Crew.CrewName FROM Crew RIGHT JOIN Players ON Crew.CrewID = Players.Leader Where Players.PlayerID < 5 ORDER BY Players.PlayerID", DB, adOpenStatic, adLockReadOnly ' Players.Name Is Not Null And
  
  While Not rst.EOF
    If IsNull(rst!Name) Then
      If opt(rst!playerID - 1).Enabled = False And player.ID = 0 Then
         opt(rst!playerID - 1).Enabled = True
         opt(rst!playerID - 1).Value = False
         opt(rst!playerID - 1).Visible = True
         imgShip(rst!playerID - 1).GrayScale = lvicNoGrayScale
      End If
      PlayCode(rst!playerID).PlayName = ""
    Else
      Lst.AddItem rst!ship & "  -  " & IIf(rst!AI = 1, "AI Bot", rst!Name) & "  -  " & rst!CrewName
      PlayCode(rst!playerID).PlayName = rst!Name
      
      If player.ID <> rst!playerID Then
         opt(rst!playerID - 1).Enabled = False
         opt(rst!playerID - 1).Value = False
         opt(rst!playerID - 1).Visible = False
         imgShip(rst!playerID - 1).GrayScale = lvicSimpleAverage
      End If
    End If
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
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
   While crewcnt < noOfCrew
      rst.Requery
      CrewID = Int((maxCrewID * Rnd)) + 1
      rst.filter = "CrewID =" & CrewID
      If Not rst.EOF Then
         DB.Execute "UPDATE SupplyDeck SET Seq =" & player.ID & " WHERE CardID = " & rst!CardID
          'add the card to the players deck
         DB.Execute "INSERT INTO PlayerSupplies (PlayerID, CardID) VALUES (" & player.ID & ", " & rst!CardID & ")"
         DB.Execute "UPDATE SupplyDeck SET Seq = " & CStr(player.ID) & " WHERE CardID = " & CStr(rst!CardID)
         'rst.Update "Seq", player.ID
         crewcnt = crewcnt + 1
      End If
   Wend
   rst.Close
End Sub

Private Function getLastPlayer() As String
Dim rst As New ADODB.Recordset, SQL
   SQL = "SELECT Scores.PlayerName FROM Scores LEFT JOIN Crew ON Scores.PlayerName = Crew.CrewName WHERE Crew.CrewName Is Null ORDER BY Scores.StartDate DESC"
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
   If Not rst.EOF Then
      getLastPlayer = rst!PlayerName
   End If
   rst.Close

End Function

