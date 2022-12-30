Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function PlaySound Lib "WINMM.DLL" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SND_ASYNC = &H1

Public Enum actionSeqCntr
   ASidle    'catch the start of turn when playerID matched
   ASselect  'bypass main Loop and await action
   ASmosey   'in mosey mode, disable fullburn, go in limbo
   ASMoseyEnd 'throw to main loop after each mosey move
   ASfullburn
   ASFullburnEnd
   ASNav
   ASnavEnd
   ASNavEvade
   ASNavEvadeEnd
   ASNavReav
   ASNavReavBorder
   ASNavReavEnd
   ASNavCrus
   ASNavCrusBorder
   ASNavCrusOutlaw
   ASNavCrusAdjacent
   ASNavCorvAdjacent
   ASNavCorvPlanetary
   ASNavCrusEnd
   ASDeal
   ASDealSelDiscard
   ASDealDrew
   ASDealSelect
   ASDealEnd
   ASBuy
   ASBuySelDiscard
   ASBuyDrew
   ASBuySelect
   ASBuyShore
   ASBuyHaven
   ASBuyEnd
   ASWork
   ASRemoveDisgr
   ASResolveAlert
   ASResolveAlertEnd
   ASEnd     'end action, selectnext player
End Enum
Public actionSeq As actionSeqCntr, NumOfReavers As Integer

Public MoseyMovesDone As Integer, FullburnMovesDone As Integer

Type Playertype
  ID As Integer
  Color As String
  PlayName As String
End Type

Public DB As ADODB.Connection, Logic As New ADODB.Recordset
Public player As Playertype, SoloGame As Boolean, PlayCode(4) As Playertype
Public HemmorrhagingFuel As Boolean, turnExtraRange As Integer
Public CruiserCutter As Integer     'record the Sector of a Cruiser/Cutter.Corvette Visit to prevenet re-triggering
Public CorvetteSeq As Integer       'extra Corvette detection flag due to 2 moves possible
Public ignoreToken As Integer       'ignore the Token in the current Sector
Public TheBigBlack As Integer       'count TheBigBlack Nav cards for Emissions Recycler
Public HigginsDealPerk As Boolean
Public pickStartSector As Integer
Public wormHoleOpen As Boolean      '133 - 104
Public DataB                        'database for game

'Public Bitpic() As Control
Public Const JOB_SUCCESS As Integer = 3      'final JobStatus value once complete
Public Const MAXJOBCARDDRAW As Integer = 3
Public Const MAXACTIVEJOBS As Integer = 3
Public Const MAXINACTIVEJOBS As Integer = 3
Public Const MAXJOBCARDACCEPT As Integer = 2
Public Const CONSIDERED As Integer = 6       'Seq status for considered cards
Public Const DISCARDED As Integer = 5        'Seq Status for Discarded Cards
Public Const DEF_CREWCAPACITY As Integer = 6
Public Const DEF_CARGOCAPACITY As Integer = 8
Public Const DEF_STASHCAPACITY As Integer = 4
Public Const NO_OF_CONTACTS As Integer = 9

Public Function Logon() As Boolean
On Error Resume Next
  If Command$ = "" Then
     DataB = App.Path & "\FireflyKalidasa.mdb"
  Else
     DataB = Command$
  End If
  Set DB = New ADODB.Connection
  DB.Open "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & DataB & ";Persist Security Info=False"
  If Err Then
     Logon = False
     MsgBox "Unable to open game datasource at " & DataB, vbCritical
  Else
     Logon = True
  End If
  
  
End Function
Public Function GetSeq()
Dim msg
   Logic.Requery
   GetSeq = Logic!Seq
   
   Select Case GetSeq
   Case "H"
      msg = "Waiting for players to join"
   Case "E"
      msg = "Waiting for a new game to be hosted"
   Case "L"
      msg = "Waiting for a Leaders to be chosen"
   Case "S"
      msg = "Waiting for the Game Setup to complete"
   Case "R"
      Select Case Logic!player
      Case "1"
         msg = "Waiting for " & PlayCode(1).PlayName & " [ORANGE] to finish their GO"
      Case "2"
         msg = "Waiting for " & PlayCode(2).PlayName & " [BLUE] to finish their GO"
      Case "3"
         msg = "Waiting for " & PlayCode(3).PlayName & " [YELLOW] to finish their GO"
      Case "4"
         msg = "Waiting for " & PlayCode(4).PlayName & " [GREEN] to finish their GO"
      End Select
   Case Else
      msg = "Wait, there's a logic ERROR!!"
   End Select
      
   PutMsg msg
End Function

Public Function GetSeqX(playerID As Integer)
Dim msg
   Logic.Requery
   GetSeqX = Logic!Seq
   playerID = Logic!player
   
    Select Case GetSeqX
   Case "H"
      msg = "Waiting for players to join"
   Case "E"
      msg = "Waiting for a new game to be hosted"
   Case "L"
      msg = "Waiting for a Leaders to be chosen"
   Case "S"
      msg = "Waiting for the Game Setup to complete"
   Case "R"
      Select Case Logic!player
      Case "1"
         msg = "Waiting for " & PlayCode(1).PlayName & " [ORANGE] to finish their GO"
      Case "2"
         msg = "Waiting for " & PlayCode(2).PlayName & " [BLUE] to finish their GO"
      Case "3"
         msg = "Waiting for " & PlayCode(3).PlayName & " [YELLOW] to finish their GO"
      Case "4"
         msg = "Waiting for " & PlayCode(4).PlayName & " [GREEN] to finish their GO"
      End Select
   Case "T"
      msg = "Waiting for Player Trading to complete"
   
   Case "U", "V"
      msg = "Waiting for the Operative's Corvette to be moved"
   
   Case "W", "X"
      msg = "Waiting for a Reaver to be moved"
   Case "X", "Y", "Z"
      msg = "Waiting for the Alliance Cruiser to be moved"
   
   Case Else
      
   End Select
   Main.Stat.Panels(1).Text = msg
 
End Function

Public Sub ClearBoard()
Dim x
   For x = 1 To 4
      PlayCode(x).PlayName = ""
   Next x
   
   DB.Execute "UPDATE Players SET Name = Null, Seq=0, SectorID = Null, Leader=0, Pay = 0, Warrants=0, Goals = 0, Fuel = 0, Parts = 0, Cargo = 0, Contraband = 0, " & _
              "Fugitive = 0, Passenger = 0, AI = 0, Solid1 = 0, Solid2 = 0, Solid3 = 0, Solid4 = 0, Solid5 = 0, Solid6 = 0, Solid7 = 0, Solid8 = 0, Solid9 = 0"
   'set starting positions for NPC ships - setup now in Timing startup
   'For x = 5 To 6 + NumOfReavers
   '   DB.Execute "UPDATE Players SET SectorID = " & varDLookup("StartSectorID", "Players", "PlayerID=" & CStr(x)) & " WHERE PlayerID= " & CStr(x)
   'Next x
   
   DB.Execute "DELETE * from Events"
   DB.Execute "DELETE * from PlayerSupplies"
   DB.Execute "DELETE * from PlayerJobs"
   DB.Execute "UPDATE Crew Set Disgruntled = 0"
   DB.Execute "UPDATE SupplyDeck Set Seq = 0"
   DB.Execute "UPDATE ContactDeck Set Seq = 0"
   DB.Execute "UPDATE Board Set Token = 0, AToken = 0, Haven = 0"
   DB.Execute "UPDATE Board Set Token = 1 WHERE SectorID IN (120,121,122)"
   CruiserCutter = 0
   CorvetteSeq = 0
   ignoreToken = 0
   
End Sub

Public Sub SetupPlayer(ByVal playerID, ByVal StoryID)
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM Story WHERE StoryID =" & StoryID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      DB.Execute "UPDATE Players SET Pay = " & rst!StartingCash & ", Warrants=0, Fuel = " & rst!StartingFuel & ", Parts = " & rst!StartingParts & " WHERE PlayerID =" & playerID
   End If
   rst.Close

Set rst = Nothing
End Sub

Public Function setNextLeader(ByVal lastplayer, ByVal leader)
Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM Players WHERE NAME IS NOT NULL ORDER BY PlayerID", DB, adOpenDynamic, adLockOptimistic
    rst.Find "PlayerID = " & lastplayer

    If Not rst.EOF Then
       'set leader for outgoing player
       rst!leader = leader
       rst.Update
       'mark the Card as selected
       DB.Execute "UPDATE SupplyDeck SET Seq =" & lastplayer & " WHERE CrewID =" & leader
       'drop this leaders Card into the Player's supplies
       'DB.Execute "INSERT INTO PlayerSupplies (PlayerID,CardID) VALUES (" & lastplayer & ", " & varDLookup("CardID", "SupplyDeck", "CrewID =" & leader) & ")"
       rst.MoveNext
       If rst.EOF Then   'end of this round
         rst.MoveFirst
       End If
       If rst!leader = 0 Then 'not set yet
          setNextLeader = rst!playerID
          Logic.Update "Player", setNextLeader
       Else 'we done here as we're back to the first player
          setNextLeader = 0
          Logic!Seq = "S"    'start game setup in main cycle
          Logic!Gamecntr = 1 'start counter, players will be on 0
          Logic!player = player.ID  'with this player as first
          Logic.Update
       End If
   End If
End Function
Public Function setNextPlayer(ByVal playerID)
Dim rst As New ADODB.Recordset
    Logic.Requery
    
    rst.Open "SELECT * FROM Players WHERE NAME IS NOT NULL ORDER BY PlayerID", DB, adOpenDynamic, adLockOptimistic
    rst.Find "PlayerID = " & playerID

    If Not rst.EOF Then
       'set my cntr to current Game Seq
       rst!Seq = Logic!Gamecntr 'set my go as done
       rst.Update
       rst.MoveNext
       If rst.EOF Then   'end of this round
         rst.MoveFirst
       End If
       setNextPlayer = rst!playerID
       Logic.Update "Player", setNextPlayer
       
       If rst!Seq = Logic!Gamecntr Then  'round over, increment GameCntr
          Logic!Gamecntr = Logic!Gamecntr + 1
          Logic.Update
       End If
       
       
   End If
End Function

Public Function setNextPlayerREV(ByVal playerID, Optional ByVal nextStatus As String = "")
Dim rst As New ADODB.Recordset
    Logic.Requery
    
    rst.Open "SELECT * FROM Players WHERE NAME IS NOT NULL ORDER BY PlayerID DESC", DB, adOpenDynamic, adLockOptimistic
    rst.Find "PlayerID = " & playerID

    If Not rst.EOF Then
       'set my cntr to current Game Seq
       rst!Seq = Logic!Gamecntr 'set my go as done
       rst.Update
       rst.MoveNext
       If rst.EOF Then   'end of this round
         rst.MoveFirst
       End If

       
       If rst!Seq = Logic!Gamecntr Then  'round over, increment GameCntr
          setNextPlayerREV = player.ID
          If nextStatus <> "" Then
             Logic!Seq = nextStatus
          End If
          Logic!player = player.ID
          Logic!Gamecntr = Logic!Gamecntr + 1
          Logic.Update
       Else
          setNextPlayerREV = rst!playerID
          Logic.Update "Player", setNextPlayerREV
       End If
       
       
   End If
End Function
'direction 0=forward,1 = reverse
Public Function setPlayer(ByVal playerID, ByVal nextStatus As String, ByVal direction As Integer, Optional ByVal check As Boolean = False) As Integer
Dim rst As New ADODB.Recordset, SQL
   Logic.Requery
   
   SQL = "SELECT * FROM Players WHERE NAME IS NOT NULL ORDER BY PlayerID "
   If direction = 1 Then
      SQL = SQL & "DESC"
   End If
   rst.Open SQL, DB, adOpenDynamic, adLockOptimistic
   rst.Find "PlayerID = " & playerID
   
   If Not rst.EOF Then
      rst.MoveNext
      If rst.EOF Then   'end of this round
        rst.MoveFirst
      End If
   End If
   
   setPlayer = rst!playerID
   
   rst.Close
   
   If Not check Then  'update it
      Logic!Seq = nextStatus
      Logic!player = setPlayer
      Logic.Update
   End If
   
End Function

Public Sub playsnd(bittype, Optional ByVal sync As Boolean = False)
Dim x As Long
Dim y As Long
Dim z As Long
Dim Path As String
On Error Resume Next

Path = App.Path & "\sounds\"

Select Case bittype
    Case 1
        Path = Path & "Burn"
    Case 2
        Path = Path & "Alert"
    Case 3
        Path = Path & "Reaver"
    Case 4
        Path = Path & "Cruiser"
    Case 5
        Path = Path & "Win"
    Case 6
        Path = Path & "yourgo"

    Case 7
        Path = Path & "mosey"

    Case 8
        Path = Path & "beep"

    Case 9
        Path = Path & "no"

    Case 10
        Path = Path & "msg"

    Case 11
        Path = Path & "gear"

    Case 12
        Path = Path & "reload"

    Case 13
        Path = Path & "clack"

    Case Else
        Path = Path & "RadioChat"
End Select

    Path = Path & ".wav"
    If sync Then
      x = PlaySound(Path, y, z)
    Else
      x = PlaySound(Path, y, SND_ASYNC)
    End If

End Sub

'can the player fullburn without hitting Reavers, looking for 1 free sector
Public Function hasValidFBMove(ByVal playerID) As Boolean
Dim currentSectorID, adjacent, a() As String, x
   
   currentSectorID = Nz(varDLookup("SectorID", "Players", "PlayerID=" & playerID), 0)
   
   adjacent = varDLookup("AdjacentRows", "Board", "SectorID=" & currentSectorID)
   a = Split(adjacent, ",")
   For x = LBound(a) To UBound(a)
      If getCutterSector(Val(a(x))) = 0 Then
         hasValidFBMove = True
         Exit For
      End If
   Next x

End Function

Public Function isAdjacent(ByVal playerID, ByVal SectorID) As Boolean
Dim currentSectorID, adjacent, a() As String, x
   
   currentSectorID = Nz(varDLookup("SectorID", "Players", "PlayerID=" & playerID), 0)
   
   adjacent = varDLookup("AdjacentRows", "Board", "SectorID=" & currentSectorID)
   a = Split(adjacent, ",")
   For x = LBound(a) To UBound(a)
      If SectorID = Val(a(x)) Then
         isAdjacent = True
         Exit For
      End If
   Next x

End Function

Public Function reaverMove(ByVal SectorID) As Boolean
Dim adjacent, a() As String, x, y
   
   If Nz(varDLookup("SectorID", "Players", "PlayerID > 4 AND SectorID=" & SectorID), 0) > 0 Or getZone(SectorID) = "A" Then
       Exit Function ' something is already there or wrong zone
   End If
      
   adjacent = varDLookup("AdjacentRows", "Board", "SectorID=" & SectorID)
   a = Split(adjacent, ",")
   For x = LBound(a) To UBound(a)
      For y = 7 To 6 + NumOfReavers 'see if a Corvette or Reaver is in the adjacent sector to the click
         If Val(a(x)) = varDLookup("SectorID", "Players", "PlayerID=" & y) Then
            MoveShip y, SectorID
            reaverMove = True
            Exit Function
         End If
      Next y
   Next x

End Function

Public Function validMove(ByVal playerID, ByVal SectorID, Optional ByVal mosey As Boolean = False) As Boolean
Dim currentSectorID, adjacent, a() As String, x, reaver
   If Not mosey Then
      For x = 7 To 6 + NumOfReavers
         reaver = varDLookup("SectorID", "Players", "PlayerID=" & x)
         If SectorID = reaver Then
            MessBox "You do not have the necessary Ship Upgrade to Full Burn into Reaver held territory", "Reaver Cutter Ahead!", "Ooops", "", getLeader()
            Exit Function ' no go
         End If
      Next x
   End If
   
   currentSectorID = Nz(varDLookup("SectorID", "Players", "PlayerID=" & playerID), 0)
   If SectorID = currentSectorID Or currentSectorID = 0 Then Exit Function 'same spot
   'NPC Zones must match
   If playerID > 4 And (IIf(getZone(SectorID) = "A", "A", "B") <> IIf(getZone(currentSectorID) = "A", "A", "B")) Then Exit Function
      
   adjacent = varDLookup("AdjacentRows", "Board", "SectorID=" & currentSectorID)
   a = Split(adjacent, ",")
   For x = LBound(a) To UBound(a)
      If SectorID = Val(a(x)) Then
         validMove = True
         Exit For
      End If
   Next x
   
   If wormHoleOpen = True And ((SectorID = 133 And currentSectorID = 104) Or (SectorID = 104 And currentSectorID = 133)) Then
      validMove = True
   End If

End Function

Public Sub displayShip(ByVal playerID, SectorID)
Dim rst As New ADODB.Recordset
Dim coords, slot
Dim c() As String
   
   slot = IIf(playerID > 4, 5, playerID)
   If SectorID = 0 Then  'remove
      Main.Verse.Imag(playerID).Visible = False
   Else
      rst.Open "SELECT * FROM Board WHERE SectorID = " & SectorID, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         coords = rst.Fields("Slot" & slot).Value
         c = Split(coords, ",")
         Main.Verse.Imag(playerID).Visible = True
         Main.Verse.Imag(playerID).Left = c(0)
         Main.Verse.Imag(playerID).top = c(1)
      
      End If
   End If
   
End Sub

Public Sub MoveShip(ByVal playerID, ByVal SectorID, Optional ByVal sound As Integer = 0, Optional ByVal syncsound As Boolean = False, Optional ByVal leaveToken As Boolean = True)
Dim rst As New ADODB.Recordset
Dim coords, slot, lastSectorID As Integer, x, a, b, TimingState As Boolean
Dim c() As String
   
   If SectorID = 0 Then Exit Sub
   TimingState = Main.Timing.Enabled
   Main.Timing.Enabled = False
   lastSectorID = getPlayerSector(playerID)
   slot = IIf(playerID > 4, 5, playerID)
   DB.BeginTrans
   DB.Execute "Update Players Set SectorID = " & SectorID & " WHERE PlayerID = " & playerID
   DB.CommitTrans

   If playerID > 6 And SectorID <> lastSectorID And lastSectorID > 0 And leaveToken Then    'cutter 7-12
      changeToken lastSectorID, 1, False    'leave another token behind
   End If
   
   If Not syncsound Then
      If sound > 0 Then
         playsnd sound, syncsound
      Else
         Select Case playerID
         Case 1 To 4  'fireflys
            playsnd 1, syncsound
         Case 5, 6  'alliance
            playsnd 4, syncsound
         Case Else   'cutters
            playsnd 3, syncsound
         End Select
      End If
   End If
   
   rst.Open "SELECT * FROM Board WHERE SectorID = " & SectorID, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      coords = rst.Fields("Slot" & slot).Value
      c = Split(coords, ",")
      Main.Verse.Imag(playerID).Visible = True
      a = Main.Verse.Imag(playerID).Left
      b = Main.Verse.Imag(playerID).top
      Main.Verse.Imag(playerID).Animate2.StopAnimation
      Main.Verse.Imag(playerID).ImageIndex = 2
      For x = b To Val(c(1)) Step IIf(b > Val(c(1)), -1, 1)
         Main.Verse.Imag(playerID).top = x            'c(1)
         Main.Verse.Imag(playerID).Refresh
         DoEvents 'slow down!
      Next x
      
      If a < Val(c(0)) Then Main.Verse.Imag(playerID).Mirror = lvicMirrorHorizontal
      For x = a To Val(c(0)) Step IIf(a > Val(c(0)), -1, 1)
         Main.Verse.Imag(playerID).Left = x           'c(0)
         Main.Verse.Imag(playerID).Refresh
         DoEvents
      Next x
      Main.Verse.Imag(playerID).Mirror = lvicMirrorNone
      Main.Verse.Imag(playerID).Animate2.StartAnimation
   End If
   rst.Close
   
   If syncsound Then
      If sound > 0 Then
         playsnd sound, syncsound
      Else
         Select Case playerID
         Case 1 To 4
            playsnd 1, syncsound
         Case 5, 6
            playsnd 4, syncsound
         Case Else
            playsnd 3, syncsound
         End Select
      End If
   End If
   
   If playerID = 6 Then 'moving the Corvette, check a Reaver is not here
      x = getCutterSector(SectorID)
      If x > 0 Then 'move this reaver back to Reaver Space
         PutMsg "The Corvette chases a Reaver Cutter off, which hightails it back to Reaver Space", playerID, Logic!Gamecntr
         'place it at Miranda and use the AI move to get it back to the Reaver Space with preference to any Player Ship :O
         DB.Execute "UPDATE Players SET SectorID = 123 WHERE PlayerID = " & x
         moveAutoAI x, 0, False, False
      End If
      'clear any Reaver Tokens
      changeToken SectorID, -1, False
      'update the Seq counter
      DB.Execute "UPDATE Players SET Seq = Seq + 1 WHERE PlayerID = 6"
   End If
      

   If playerID = 5 Then ' Alliance & Harken
      DB.Execute "UPDATE Contact SET SectorID =" & SectorID & " WHERE ContactID = 5"
   End If
   
   Main.Timing.Enabled = TimingState
   
   RefreshBoard
   
   If playerID > 4 And getPlayerSector(player.ID) = SectorID And actionSeq <> ASNavEvade Then
      If checkWhisperX1(SectorID) Then
         actionSeq = ASNavEvade ' and get away
      End If
   End If
   
End Sub

Public Function checkWhisperX1(ByVal SectorID) As Boolean
Dim x, g, dice

   x = getCutterSector(SectorID)
   If x = 0 And getCruiserSector() = SectorID Then 'not a cutter
      x = 5 'but a cruiser
   ElseIf x = 0 And getCorvetteSector() = SectorID Then 'not a cutter
      x = 6 'but a Corvette
   End If
   
   If x > 0 Then 'we got company!
      g = hasShipUpgrade(player.ID, 14)
      If g > 0 Then 'WhisperX1
         If MessBox("The Xunsu Whisper X1 should outrun your contact, do you want to give it a burn?", "Sector Contact", "Yes", "No", 0, 0, 14) = 0 Then
            dice = RollDice(6)
            If dice > 3 Then
               checkWhisperX1 = True
               PutMsg player.PlayName & " fired up the Xunsu Whisper X1, now EVADE", player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
            Else
               PutMsg player.PlayName & " fired up the Xunsu Whisper X1 but she needs more power.  Outrun Failed!", player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, dice
            End If
         End If
      End If
   End If
   

End Function

Public Sub changeToken(ByVal SectorID As Integer, ByVal cnt As Integer, Optional ByVal sound As Boolean = True)
   If cnt < 0 Then
      DB.Execute "UPDATE Board SET Token = " & IIf(SectorID > 119 And SectorID < 123, "1", "0") & " WHERE SectorID = " & SectorID
   Else
      If varDLookup("Token", "Board", "SectorID=" & SectorID) < 6 Then
         If sound Then playsnd 2
         DB.Execute "UPDATE Board Set Token = Token + " & Str(cnt) & " WHERE SectorID = " & SectorID
      End If
   End If

End Sub

Public Function getAToken(ByVal SectorID As Integer) As Integer
   getAToken = varDLookup("AToken", "Board", "SectorID=" & SectorID)
End Function

Public Sub changeAToken(ByVal SectorID As Integer, ByVal cnt As Integer)
   If getAToken(SectorID) + cnt < 0 Then
      DB.Execute "UPDATE Board Set AToken = 0 WHERE SectorID = " & SectorID
   Else
      If Not getHaven(SectorID) > 0 And SectorID <> 120 And SectorID <> 121 And SectorID <> 122 And varDLookup("AToken", "Board", "SectorID=" & SectorID) < 6 Then
         playsnd 2
         DB.Execute "UPDATE Board Set AToken = AToken + " & Str(cnt) & " WHERE SectorID = " & SectorID
      End If
   End If

End Sub

Public Sub placeHaven(ByVal playerID, ByVal SectorID)
    DB.Execute "UPDATE Board Set Haven = " & Str(playerID) & " WHERE SectorID = " & SectorID
End Sub

Public Function useHavens(ByVal StoryID) As Boolean
   
   useHavens = (varDLookup("Havens", "Story", "StoryID=" & StoryID) = "1")

End Function

Public Function getHaven(ByVal SectorID) As Integer
   
   getHaven = varDLookup("Haven", "Board", "SectorID=" & SectorID)

End Function

Public Sub MoveSolid(ByRef Imag As Label, ByVal ContactID)
Dim rst As New ADODB.Recordset
Dim coords, slot, SectorID
Dim c() As String
   
   SectorID = varDLookup("SectorID", "Contact", "ContactID=" & ContactID)
   'avoid clash with Cruiser/Cutter
   If ContactID = 5 Then
      slot = 5
   Else
      slot = 4
   End If
   rst.Open "SELECT * FROM Board WHERE SectorID = " & SectorID, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      coords = rst.Fields("Slot" & slot).Value
      c = Split(coords, ",")
      
      Imag.Left = c(0)
      Imag.top = c(1)
   
   End If
   
End Sub

Public Sub RefreshBoard()
Dim rst As New ADODB.Recordset, Index
Dim coords, c() As String

   With Main.Verse

   'ships
   rst.Open "SELECT * FROM Players", DB, adOpenForwardOnly, adLockReadOnly ' WHERE PlayerID < 5 AND Name is not null
   While Not rst.EOF  'no other players
      If IsNull(rst!SectorID) Then
         .Imag(rst!playerID).Visible = False
      Else
         displayShip rst!playerID, rst!SectorID
      End If
      rst.MoveNext
   Wend
   rst.Close
   
   'solid labels
   For Index = 1 To NO_OF_CONTACTS
      If isSolid(player.ID, Index) Then
         .lblSolid(Index).Visible = True
         MoveSolid .lblSolid(Index), Index
      Else
         .lblSolid(Index).Visible = False
      End If
   Next Index
   
   'tokens
   rst.Open "SELECT * FROM Board WHERE SectorID > 0 ORDER BY SectorID", DB, adOpenDynamic, adLockOptimistic
   While Not rst.EOF
      'coords = rst.Fields("Slot5").Value
      'c = Split(coords, ",")
      '.imgToken(rst!SectorID).Left = c(0) ' rst!SLeft + rst!SWidth / 4
      '.imgToken(rst!SectorID).top = c(1)  ' rst!STop + rst!SHeight / 4
      If rst!Token > 0 Then
         .imgToken(rst!SectorID).Picture = LoadPictureGDIplus(App.Path & "\Pictures\RToken" & IIf(rst!Token > 6, 6, rst!Token) & ".bmp")
         .imgToken(rst!SectorID).Visible = True
         .imgToken(rst!SectorID).TransparentColor = &HFFFFFF
         .imgToken(rst!SectorID).TransparentColorMode = lvicUseTransparentColor
      Else
         .imgToken(rst!SectorID).Visible = False
      End If
      
      If rst!AToken > 0 Then
         .imgAToken(rst!SectorID).Picture = LoadPictureGDIplus(App.Path & "\Pictures\AToken" & IIf(rst!AToken > 6, 6, rst!AToken) & ".bmp")
         .imgAToken(rst!SectorID).Visible = True
         .imgAToken(rst!SectorID).TransparentColor = &HFFFFFF
         .imgAToken(rst!SectorID).TransparentColorMode = lvicUseTransparentColor
      Else
         .imgAToken(rst!SectorID).Visible = False
      End If

      If rst!Haven > 0 Then
         coords = rst.Fields("Slot5").Value
         c = Split(coords, ",")
         .imgHaven(rst!Haven).Left = c(0)
         .imgHaven(rst!Haven).top = c(1)
         .imgHaven(rst!Haven).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Haven" & rst!Haven & ".bmp")
         .imgHaven(rst!Haven).Visible = True
         .imgHaven(rst!Haven).TransparentColor = &HFFFFFF
         .imgHaven(rst!Haven).TransparentColorMode = lvicUseTransparentColor
      End If

      rst.MoveNext
   Wend
   
   End With

End Sub

Public Function getLeader() As Integer
      getLeader = Nz(varDLookup("Leader", "Players", "PlayerID = " & player.ID), 0)
End Function

Public Function getPlayerSector(ByVal playerID) As Integer
      getPlayerSector = Nz(varDLookup("SectorID", "Players", "PlayerID = " & playerID), 0)
End Function

Public Function getCruiserSector() As Integer
      getCruiserSector = Nz(varDLookup("SectorID", "Players", "PlayerID = 5"), 0)
End Function

Public Function getCorvetteSector() As Integer
      getCorvetteSector = Nz(varDLookup("SectorID", "Players", "PlayerID = 6"), 0)
End Function
Public Function getCorvetteSeq() As Integer
   getCorvetteSeq = varDLookup("Seq", "Players", "PlayerID=6")
End Function


'return the playerid of any Alliance Ship in the sector
Public Function getCruiserCorvette(ByVal SectorID) As Integer
      getCruiserCorvette = Nz(varDLookup("PlayerID", "Players", "PlayerID IN (5, 6) AND SectorID=" & SectorID), 0)
End Function

'return the playerid of any reaver in the sector
Public Function getCutterSector(ByVal SectorID) As Integer
      getCutterSector = Nz(varDLookup("PlayerID", "Players", "PlayerID > 6 AND SectorID=" & SectorID), 0)
End Function

Public Function getZone(ByVal SectorID As Integer) As String
  getZone = Nz(varDLookup("Zones", "Board", "SectorID=" & SectorID))
End Function

Public Function getSupplyName(ByVal SupplyID As Integer) As String
  getSupplyName = Nz(varDLookup("SupplyName", "Supply", "SupplyID=" & SupplyID))
End Function

Public Function getClearSector(ByVal SectorID As Integer) As String
  If Nz(varDLookup("SectorID", "Players", "SectorID=" & SectorID), 0) = 0 Then
      getClearSector = getZone(SectorID)
  End If
End Function

'Public Function getOutlawZone(ByVal SectorID As Integer) As String
'Dim playerID As Integer
'   playerID = Nz(varDLookup("SectorID", "Players", "Name IS NOT NULL AND PlayerID <> " & player.ID & " AND SectorID=" & SectorID), 0)
'   If isOutlaw(playerID) Then
'      getOutlawZone = getZone(SectorID)
'   End If
'
'End Function

'return playerID that is already in this player's chosen sector
Public Function CheckClash(ByVal playerID, ByVal SectorID, ByVal Havens As Boolean) As Boolean
Dim rst As New ADODB.Recordset
   
   rst.Open "SELECT * FROM Players WHERE PlayerID < 5 AND PlayerID <> " & playerID & " AND Name is not null AND SectorID = " & SectorID, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then 'other players here
      CheckClash = True
      MessBox PlayCode(rst!playerID).PlayName & "'s Ship is in that Sector already", "Sector Clash", "Ooops", "", getLeader()
      Exit Function
   End If
   rst.Close
   'if Haven, then must be a planet sector with no Contact or Supply
   If Havens Then
      rst.Open "SELECT Planet.SectorID FROM Supply RIGHT JOIN (Contact RIGHT JOIN Planet ON Contact.SectorID = Planet.SectorID) ON Supply.SectorID = Planet.SectorID Where Contact.ContactID Is Null And Supply.SupplyID Is Null And Planet.SectorID = " & SectorID, DB, adOpenForwardOnly, adLockReadOnly
      If rst.EOF Then
         CheckClash = True
         MessBox "Pick a Planet Sector with no Contact, Supply, or Cruiser", "Sector Clash", "Ooops", "", getLeader()
      End If
   End If
   
End Function

'this Controls all the Story Goals, their Jobs and the WIN
Public Function CheckWon(ByVal playerID) As Boolean
Dim rst As New ADODB.Recordset, SQL, frmWin As frmWinner

   SQL = "SELECT * FROM Players WHERE PlayerID=" & playerID
   
   rst.Open SQL, DB, adOpenDynamic, adLockReadOnly
   If Not rst.EOF Then
      CheckWon = doGoalCheck(playerID, Logic!StoryID, rst!Goals, rst!Seq)
   End If

   rst.Close

   If CheckWon = True Then
      playsnd 5
      DB.Execute "INSERT INTO Scores (StoryID,PlayerName,Turns,StartDate,PlayDate) Values (" & CStr(Logic!StoryID) & ",'" & SQLFilter(player.PlayName) & "'," & CStr(Logic!Gamecntr - 1) & ", #" & Format(varDLookup("EventTime", "Events", "Event ='" & player.PlayName & "''s on the Map'"), "MM-DD-YY HH:nn") & "#, #" & Format(Now, "MM-DD-YY HH:nn") & "#)"
      PutMsg PlayCode(playerID).PlayName & " has WON the Game in " & Logic!Gamecntr - 1 & " turns", playerID, Logic!Gamecntr
      Set frmWin = New frmWinner
      frmWin.Show 1
   End If


Set rst = Nothing
End Function

Private Function doGoalCheck(ByVal playerID, ByVal StoryID, ByVal Goal, ByVal Seq) As Boolean
Dim rst As New ADODB.Recordset, goaldone As Boolean, a() As String
Dim SQL, x, cnt As Integer
   If Goal = -1 Then Exit Function
   goaldone = True 'until proven otherwise
   SQL = "SELECT * FROM StoryGoals WHERE StoryID=" & StoryID & " AND Goal = " & CStr(Goal + 1)
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      'SOLID
      If rst!SolidCount > 0 Then
         For x = 1 To NO_OF_CONTACTS
            If isSolid(playerID, x) Then
               cnt = cnt + 1
            End If
         Next x
         If cnt < rst!SolidCount Then
            goaldone = False
         End If
      ElseIf goaldone And Nz(rst!Solid) <> "" Then
         a = Split(rst!Solid, ",")
         For x = LBound(a) To UBound(a)
            If Not isSolid(playerID, a(x)) Then
               goaldone = False
               Exit For
            End If
         Next x
      End If
      
      'CompleteJob
      If goaldone And rst!CompleteJobID > 0 Then
         If Not jobSuccess(playerID, rst!CompleteJobID) Then
            goaldone = False
         End If
      End If
      'money
      If goaldone And rst!Cash > 0 Then
         If getMoney(playerID) < rst!Cash Then
            goaldone = False
         End If
      End If
      If goaldone And rst!fight > 0 Then
         If getSkill(playerID, cstrSkill(1)) < rst!fight Then
            goaldone = False
         End If
      End If
      If goaldone And rst!tech > 0 Then
         If getSkill(playerID, cstrSkill(2)) < rst!tech Then
            goaldone = False
         End If
      End If
      If goaldone And rst!Negotiate > 0 Then
         If getSkill(playerID, cstrSkill(3)) < rst!Negotiate Then
            goaldone = False
         End If
      End If
      'be at a Sector
      If goaldone And rst!SectorID > 0 Then
         If getPlayerSector(playerID) <> rst!SectorID Then
            goaldone = False
         End If
      End If
      'Misbehaves
      If goaldone And rst!Misbehaves > 0 Then
         frmAction.lblMis.Visible = True
         If countMisbehaves(playerID) < rst!Misbehaves Then
            goaldone = False
         End If
      End If
      
      If goaldone And rst!MeetCruiser > 0 Then
         If getCruiserSector() <> getPlayerSector(playerID) Then
            goaldone = False
         End If
      End If
      If goaldone And rst!MeetCorvette > 0 Then
         If getCorvetteSector() <> getPlayerSector(playerID) Then
            goaldone = False
         End If
      End If
      ' END of positive Tests ================================================
      
      'If we still good and we have Win flag, we WIN
      If goaldone And rst!Win > 0 Then
         doGoalCheck = True
      End If
            
      'Negative tests ---- TurnLimit
      If rst!TurnLimit > 0 And Not doGoalCheck Then
         If Seq > rst!TurnLimit Then
            addGoal playerID, -1
            MessBox "You have Failed to meet the Story Goals :( " & vbNewLine & "You may continue on, your call..", "GAME OVER", "Hmmm", "", getLeader()
            goaldone = False
         End If
      End If
      
      'load any Passengers is there is room
      If goaldone And rst!Passenger > 0 Then
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) < rst!Passenger Then
            MessBox "You have no Room for " & rst!Passenger & " Passenger/s", "Passenger/s", "Ooops", "", getLeader()
            goaldone = False
         Else
            DB.Execute "UPDATE Players SET Passenger = Passenger + " & CStr(rst!Passenger) & " WHERE PlayerID = " & player.ID
         End If
      End If
      
      'if we here and goaldone then Goal IS Done
      If goaldone Then
         addGoal playerID, 1
      End If
      
      
      'if we here and goaldone, we good to deliver job
      If goaldone And rst!IssueJobID > 0 Then
         assignDeal playerID, rst!IssueJobID
         If Not (Main.frmJob Is Nothing) Then
            Main.frmJob.RefreshJobs
         End If

      End If
      
     
      ' we good to give new instructions
      If goaldone And Nz(rst!Instructions) <> "" And Not doGoalCheck Then
         PutMsg player.PlayName & ", you have completed Goal " & Goal + 1 & vbNewLine & rst!Instructions, playerID, Logic!Gamecntr, True, getLeader()
      End If
      
      
   End If
   rst.Close
   Set rst = Nothing

End Function

Public Function countMisbehaves(ByVal playerID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Count(CardID) AS cnt FROM MisbehaveDeck WHERE Seq =" & playerID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      countMisbehaves = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
End Function

Private Sub addGoal(ByVal playerID, Optional ByVal change As Integer = 1)

   If change = -1 Then
      DB.Execute "UPDATE Players SET Goals = -1 WHERE PlayerID =" & playerID
   Else
      DB.Execute "UPDATE Players SET Goals = Goals + " & change & " WHERE PlayerID =" & playerID
   End If

End Sub


Public Sub PutMsg(msg, Optional playerID = 0, Optional turn = 0, Optional ByVal force As Boolean = False, Optional ByVal CrewID As Integer = 0, Optional ByVal GearID As Integer = 0, Optional ByVal ShipUpgradeID As Integer = 0, Optional ByVal ContactID As Integer = 0, Optional ByVal refreshShip As Integer = 0, Optional ByVal dice As Integer = 0)
Dim SQL, frmPop As frmPopup
On Error GoTo err_handler

   If Left(msg, 3) <> "Wai" Then 'waiting for game to start
      SQL = "INSERT INTO Events (Eventtime, Event, PlayerID, Turn, RefreshShip"
      SQL = SQL & ") Values (#" & Now & "#, '" & SQLFilter(msg) & "', " & playerID & ", " & turn & ", " & refreshShip
      SQL = SQL & ")"
      DB.Execute SQL
   End If

   If force Then
      Events.getNewEvents
      Set frmPop = New frmPopup
      With frmPop
      .lblMsg = msg
      If CrewID > 0 Then
         .Width = 10320
         .Height = 5040
         .lblMsg.Height = 1600
         .cmd(0).top = 1900
         .pic.Visible = True
         Set .pic.Picture = LoadPicture(App.Path & "\pictures\" & varDLookup("Picture", "Crew", "CrewID=" & CrewID))
      ElseIf GearID > 0 Then
         .Width = 10320
         .Height = 5040
         .lblMsg.Height = 1600
         .cmd(0).top = 1900
         .pic.Visible = True
         Set .pic.Picture = LoadPicture(App.Path & "\pictures\" & varDLookup("Picture", "Gear", "GearID=" & GearID))
      ElseIf ShipUpgradeID > 0 Then
         .Width = 10320
         .Height = 5040
         .lblMsg.Height = 1600
         .cmd(0).top = 1900
         .pic.Visible = True
         Set .pic.Picture = LoadPicture(App.Path & "\pictures\" & varDLookup("Picture", "ShipUpgrade", "ShipUpgradeID=" & ShipUpgradeID))
      ElseIf ContactID > 0 Then
         .Width = 10320
         .Height = 5040
         .lblMsg.Height = 1600
         .cmd(0).top = 1900
         .pic.Visible = True
         Set .pic.Picture = LoadPicture(App.Path & "\pictures\" & varDLookup("Picture", "Contact", "ContactID=" & ContactID))
      End If
      If dice > 0 And dice < 7 Then
         .picDice(0).Visible = True
         .picDice(0).Picture = LoadPictureGDIplus(App.Path & "\pictures\D" & dice & ".bmp") ' LoadPicture(App.Path & "\pictures\D" & dice & ".bmp")

         .picDice(0).TransparentColor = 0
         .picDice(0).TransparentColorMode = lvicUseTransparentColor

      ElseIf dice > 6 Then
         .picDice(0).Visible = True
         .picDice(0).Picture = LoadPictureGDIplus(App.Path & "\pictures\D6.bmp") 'LoadPicture(App.Path & "\pictures\D6.jpg")
         .picDice(0).TransparentColor = 0
         .picDice(0).TransparentColorMode = lvicUseTransparentColor
         .picDice(1).Visible = True
         .picDice(1).Picture = LoadPictureGDIplus(App.Path & "\pictures\D" & (dice - 6) & ".bmp")  'LoadPicture(App.Path & "\pictures\D" & (dice - 6) & ".bmp")
         .picDice(1).TransparentColor = 0
         .picDice(1).TransparentColorMode = lvicUseTransparentColor
      End If
      playsnd 10
      frmPop.Show 1, Main
      End With

   End If

   Main.Stat.Panels(1).Text = msg
   
normal_exit:
   Set frmPop = Nothing
   Exit Sub
   
err_handler:
   MsgBox "PutMsg Error: " & vbCrLf & Err.Description
   Resume normal_exit
   
End Sub

Public Function MessBox(ByVal msg As String, ByVal title As String, ByVal button1 As String, Optional ByVal button2 As String = vbNullString, Optional ByVal CrewID As Integer = 0, Optional ByVal GearID As Integer = 0, Optional ByVal ShipUpgradeID As Integer = 0, Optional ByVal dice As Integer = 0)
Dim frmPop As frmPopup
On Error GoTo err_handler
  
   Set frmPop = New frmPopup
   With frmPop
      .lblMsg = msg
      .Caption = title
      .cmd(0).Caption = button1
      If button2 <> "" Then
         .cmd(1).Visible = True
         .cmd(1).Caption = button2
      End If
      If CrewID > 0 Then
         .Width = 10320
         .Height = 5040
         .lblMsg.Height = 1600
         .cmd(0).top = 1900
         .cmd(1).top = 1900
         .pic.Visible = True
         Set .pic.Picture = LoadPicture(App.Path & "\pictures\" & varDLookup("Picture", "Crew", "CrewID=" & CrewID))
      ElseIf GearID > 0 Then
         .Width = 10320
         .Height = 5040
         .lblMsg.Height = 1600
         .cmd(0).top = 1900
         .cmd(1).top = 1900
         .pic.Visible = True
         Set .pic.Picture = LoadPicture(App.Path & "\pictures\" & varDLookup("Picture", "Gear", "GearID=" & GearID))
      ElseIf ShipUpgradeID > 0 Then
         .Width = 10320
         .Height = 5040
         .lblMsg.Height = 1600
         .cmd(0).top = 1900
         .cmd(1).top = 1900
         .pic.Visible = True
         Set .pic.Picture = LoadPicture(App.Path & "\pictures\" & varDLookup("Picture", "ShipUpgrade", "ShipUpgradeID=" & ShipUpgradeID))
         
      End If
      If dice > 0 And dice < 7 Then
         .picDice(0).Visible = True
         .picDice(0).Picture = LoadPictureGDIplus(App.Path & "\pictures\D" & dice & ".bmp") ' LoadPicture(App.Path & "\pictures\D" & dice & ".bmp")

         .picDice(0).TransparentColor = 0
         .picDice(0).TransparentColorMode = lvicUseTransparentColor

      ElseIf dice > 6 Then
         .picDice(0).Visible = True
         .picDice(0).Picture = LoadPictureGDIplus(App.Path & "\pictures\D6.bmp") 'LoadPicture(App.Path & "\pictures\D6.bmp")
         .picDice(0).TransparentColor = 0
         .picDice(0).TransparentColorMode = lvicUseTransparentColor
         .picDice(1).Visible = True
         .picDice(1).Picture = LoadPictureGDIplus(App.Path & "\pictures\D" & (dice - 6) & ".bmp")  'LoadPicture(App.Path & "\pictures\D" & (dice - 6) & ".bmp")
         .picDice(1).TransparentColor = 0
         .picDice(1).TransparentColorMode = lvicUseTransparentColor
      End If
      playsnd 10
      frmPop.Show 1, Main
      MessBox = .result
   End With

   
normal_exit:
   Exit Function
   
err_handler:
   MsgBox "MessBox Error: " & vbCrLf & Err.Description
   Resume normal_exit

End Function

'Deck Seq 0: unset/removed, 1-4 held by PlayerID, 5 Discard pile, 6 consider, 10+ Deck
Public Sub ShuffleDeck(ByVal Deck As String, Optional ByVal filter As Boolean = False, Optional ByVal reshuffatend As Boolean = False, Optional ByVal Zone As String = "")
Dim rst As New ADODB.Recordset
Dim SQL, y, CardID, cnt, primeKey As String
   'reset the card seq to zero to show not shuffled yet
   Select Case Deck
   Case "Nav"
      primeKey = "CardID"
   Case Else
      primeKey = Deck & "ID"
   End Select
   
   SQL = "UPDATE " & Deck & "Deck SET Seq = 500" & IIf(filter, " WHERE " & primeKey & " > 0", "") 'skip the system cards in the -ves
     
   If Zone <> "" Then
      SQL = SQL & " AND Zones = '" & Zone & "'"
   End If
   
   DB.Execute SQL
   
   
   rst.Open "SELECT MAX(CardID) as CNT FROM " & Deck & "Deck", DB, adOpenForwardOnly, adLockReadOnly
   cnt = rst!cnt + 100
   rst.Close
   
   
   'filter out ID's not allocated to the owner (eg. SupplyDeck has system owned records for Leaders and Upgrades that should not be shuffled)
   rst.Open "SELECT * FROM " & Deck & "Deck WHERE Seq=500" & IIf(filter, " AND " & primeKey & " > 0", "") & IIf(reshuffatend, " AND Reshuffle = 0", "") & IIf(Zone <> "", " AND Zones = '" & Zone & "'", ""), DB, adOpenDynamic, adLockOptimistic
   
   Randomize Timer
   
   Do
      CardID = rst!CardID
      y = Int((cnt * Rnd)) + 10
      rst.MoveFirst
      rst.Find "Seq = " & y
      If rst.EOF Then 'seq value not found
         'go back to the card
         rst.MoveFirst
         rst.Find "CardID = " & CardID
         rst!Seq = y
         rst.Update
         rst.MoveNext
         If rst.EOF Then Exit Do
      Else 'already a seq with that value
         'go back and try again
         rst.Find "CardID = " & CardID
      End If
   Loop
   If reshuffatend Then 'chuck em in the discard pile
      DB.Execute "UPDATE " & Deck & "Deck SET Seq = 5 WHERE Reshuffle = 1"
   End If
   rst.Close
   
   
   
   Set rst = Nothing
End Sub

Public Sub DrawDeck(ByVal Deck As String, ByVal ID As Integer, ByVal draw As Integer, Optional ByVal Seq As Integer = DISCARDED)
Dim rst As New ADODB.Recordset
Dim SQL, cnt
   cnt = 0
   SQL = "SELECT * FROM " & Deck & "Deck WHERE Seq > 6 AND " & Deck & "ID =" & CStr(ID) & " ORDER BY Seq"
   rst.Open SQL, DB, adOpenDynamic, adLockOptimistic
   Do While Not rst.EOF
      cnt = cnt + 1
      rst!Seq = Seq
      rst.Update
      If draw = cnt Then Exit Do
      rst.MoveNext
   Loop
   rst.Close
   
   Set rst = Nothing

End Sub


Public Function getUnseenDeck(ByVal Deck As String, ByVal ID As Integer) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Count(CardID) AS cnt FROM " & Deck & "Deck WHERE Seq > 6 AND " & Deck & "ID=" & ID
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getUnseenDeck = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function getUnseenMBDeck() As Variant
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Count(CardID) AS cnt FROM MisbehaveDeck WHERE Seq > 5"
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getUnseenMBDeck = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function getUnseenNavDeck(ByVal Zone) As Variant
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Count(CardID) AS cnt FROM NavDeck WHERE Seq > 6 AND Zones = '" & Zone & "'"
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getUnseenNavDeck = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function getZoneDesc(ByVal Zone) As String
   Select Case Zone
   Case "A"
      getZoneDesc = "Alliance Nav "
   Case "B"
      getZoneDesc = "Border Nav "
   Case "R"
      getZoneDesc = "Rim Nav "
   End Select
      
End Function

Public Function SQLFilter(ByVal Source)
Dim x, y
'  Looks for single quotes and doubles them ('') to create a literal
   x = 1
   Do
      y = InStr(x, Source, "'")
      If y Then
        Source = Left(Source, y) & Mid(Source, y)
        x = y + 2
      End If
   Loop While y
      
'   Source = Replace(Source, "%", "-")
'   Source = Replace(Source, "#", "-")
'   Source = Replace(Source, "*", "-")
'   Source = Replace(Source, "^", "-")
'   Source = Replace(Source, "$", "-")
'   Source = Replace(Source, "!", "-")
      
   SQLFilter = Source
End Function

Public Function getPlayerColor(mode) As OLE_COLOR
    Select Case mode
     Case 1
        getPlayerColor = &H80FF&          '  &H737CF7
     Case 2
        getPlayerColor = 16677990 ' &HFFC71D  '1DC7FF  FF8080
     Case 3
        getPlayerColor = &H68F8F6 ' F6F868 C0FFFF
     Case 4
        getPlayerColor = &H98F4A3 ' A3F498 FF00&
     Case Else
        getPlayerColor = &H80000005
    End Select
End Function

Public Function LoadCombo(cbo As Control, ByVal mode As String, Optional filter As String = "") As Boolean
Dim SQL As String, last  ', MP
Dim rst As New ADODB.Recordset

On Error GoTo err_handler

   If mode = vbNullString Then Exit Function
   
   'MP = Screen.MousePointer
   'Screen.MousePointer = 11
 
   cbo.Clear
   Select Case mode
   Case "crew"
      SQL = "SELECT Crew.* FROM Crew " & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!CrewName
         cbo.ItemData(cbo.NewIndex) = rst!CrewID
         rst.MoveNext
      Wend
      
   Case "gear"
      SQL = "SELECT SupplyDeck.CardID, Gear.GearName FROM Gear INNER JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID " & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!GearName
         cbo.ItemData(cbo.NewIndex) = rst!CardID
         rst.MoveNext
      Wend
      
   Case "shipupgrd"
      SQL = "SELECT SupplyDeck.CardID, ShipUpgrade.UpgradeName FROM ShipUpgrade INNER JOIN SupplyDeck ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID " & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!UpgradeName
         cbo.ItemData(cbo.NewIndex) = rst!CardID
         rst.MoveNext
      Wend
      
   Case "story"
      SQL = "SELECT * FROM Story" & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!StoryTitle
         cbo.ItemData(cbo.NewIndex) = rst!StoryID
         last = rst!StoryID
         rst.MoveNext
      Wend
      cbo.AddItem "Add New"
      cbo.ItemData(cbo.NewIndex) = (last + 1)
      
   Case "shipupgd"
      SQL = "SELECT ShipUpgrade.UpgradeName, ShipUpgrade.UpgradeDescr, ShipUpgrade.Pay, SupplyDeck.CardID "
      SQL = SQL & "FROM ShipUpgrade INNER JOIN SupplyDeck ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
      SQL = SQL & "WHERE SupplyDeck.Seq " & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!UpgradeName & " - " & rst!UpgradeDescr & " $" & rst!pay
         cbo.ItemData(cbo.NewIndex) = rst!CardID
         rst.MoveNext
      Wend
      
   Case "killcrew"
      SQL = "SELECT Crew.CrewName, Crew.CrewDescr, PlayerSupplies.CardID "
      SQL = SQL & "FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID "
      SQL = SQL & "WHERE PlayerSupplies.OffJob=0 AND PlayerSupplies.PlayerID=" & filter  'Crew.Leader <> 1 AND
       SQL = SQL & " ORDER BY Crew.Leader, Crew.CrewName"
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!CrewName & " - " & rst!CrewDescr
         cbo.ItemData(cbo.NewIndex) = rst!CardID
         rst.MoveNext
      Wend
      
   Case "planet"
      SQL = "SELECT * FROM planet WHERE planetID > 0 " & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!PlanetName
         cbo.ItemData(cbo.NewIndex) = rst!SectorID
         rst.MoveNext
      Wend
      
   Case "contact"
      SQL = "SELECT * FROM Contact " & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!ContactName
         cbo.ItemData(cbo.NewIndex) = rst!ContactID
         rst.MoveNext
      Wend
   
      
   Case "contactdeck"
      SQL = "SELECT * FROM ContactDeck " & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem CStr(rst!CardID) & " - " & rst!JobName & ": " & rst!JobOrder
         cbo.ItemData(cbo.NewIndex) = rst!CardID
         rst.MoveNext
      Wend
      
      
   Case "profession"
      SQL = "SELECT * FROM profession " & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!ProfessionName
         cbo.ItemData(cbo.NewIndex) = rst!ProfessionID
         rst.MoveNext
      Wend
      
   Case "skill"
      For last = 1 To 3
         cbo.AddItem cstrSkill(last)
         cbo.ItemData(cbo.NewIndex) = last
      Next last
      
   Case "jobtype"
      SQL = "SELECT * FROM jobtype " & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!JobTypeDescr
         cbo.ItemData(cbo.NewIndex) = rst!JobTypeID
         rst.MoveNext
      Wend
      
   Case "task"
      SQL = "SELECT * FROM job " & filter
      SQL = SQL & " ORDER BY SectorID, jobdesc"
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst!JobID & ": " & getPlanetSector(rst!JobID) & " - " & rst!JobDesc
         ' & "- Misbehaves: " & rst!Misbehave & ", Cargo: " & rst!cargo & ", Contraband: " & rst!contraband & ", Passengers: " & rst!Passenger & ", Fugitives: " & rst!fugitive & _
         ", Fuel: " & rst!fuel & ", Parts: " & rst!parts & ", TagnBag: " & rst!Tagnbag & ", Dbl Down: " & rst!DoubleDown
         cbo.ItemData(cbo.NewIndex) = rst!JobID
         rst.MoveNext
      Wend
      
   Case Else
      SQL = "SELECT * FROM " & mode & filter
      rst.Open SQL, DB, 0, 1
      While Not rst.EOF
         cbo.AddItem rst.Fields(2).Value
         cbo.ItemData(cbo.NewIndex) = rst.Fields(1).Value
         rst.MoveNext
      Wend
   End Select
     
   LoadCombo = True

normal_exit:
   Set rst = Nothing
   'Screen.MousePointer = MP
   Exit Function
   
err_handler:
   MsgBox "LoadCombo Error: " & vbCrLf & Err.Description
   Resume normal_exit

End Function

Public Function SetCombo(cmbo As Control, ByVal itemTxt As String, ByVal itemVal, Optional RightSide As Boolean = True) As Boolean
Dim x
On Error GoTo err_handler

   'validate input itemdata to screen Null value
   itemVal = Nz(itemVal, 0)

   With cmbo
      For x = 0 To .ListCount - 1
         If itemTxt = vbNullString Then   'set using itemdata value in itemVal
            If .ItemData(x) = Val(itemVal) Then
              .ListIndex = x
              SetCombo = True
              Exit For
            End If
         Else                             'set using an itemdatastring value stored in the list text.
            If (UCase(Trim(Right(.List(x), itemVal))) = UCase(itemTxt) And RightSide) Or _
               (UCase(Trim(Left(.List(x), itemVal))) = UCase(itemTxt) And Not RightSide) Then
              .ListIndex = x
              SetCombo = True
              Exit For
            End If
         End If
      Next x
   
      'if no match found then reset combo
      If Not SetCombo = True Then .ListIndex = -1

   End With
   
normal_exit:
   Exit Function

err_handler:
   MsgBox "SetCombo " & itemTxt & " " & itemVal & vbCrLf & Err.Description
   Resume normal_exit
   
End Function

Public Function GetCombo(cmbo As Control, Optional isItemData As Boolean = True, Optional trimSize = 0, Optional RightSide As Boolean = True) As Variant
   With cmbo
      If .ListIndex = -1 Then
         GetCombo = -1
      Else
         If isItemData Then
            GetCombo = .ItemData(.ListIndex)
         Else
            If RightSide Then
               GetCombo = Trim(Right(.List(.ListIndex), trimSize))
            Else
               GetCombo = Trim(Left(.List(.ListIndex), trimSize))
            End If
         End If
      End If
   End With
End Function

Public Function Nz(ByVal vvarValue As Variant, _
  Optional ByVal vvarValueIfNull As Variant = vbNullString) As Variant

  On Error GoTo errhandler

  'if the supplied field is null, return the alternative value
  Nz = IIf(IsNull(vvarValue), vvarValueIfNull, vvarValue)

  Exit Function
  
errhandler:
  
  'Return null for any errors
  Nz = Null

End Function

Public Function varDLookup(ByVal vstrField As String, ByVal vstrDomain As String, Optional ByVal vstrCriteria As String = vbNullString, Optional ByVal alias As String = vbNullString) As Variant

 Dim rstLookup As ADODB.Recordset

  'The SQL to locate the status code from the schema
  Dim strSQL As String

  On Error GoTo errhandler

  'Assume no record will be found
  varDLookup = Null

  'Prefix the where clause to the criteria if supplied
  If Len(vstrCriteria) > 0 Then vstrCriteria = " WHERE " & vstrCriteria

  'Generate the SQL statement to return the status code
  'for the currently displayed outage
  strSQL = "SELECT " & vstrField & " FROM " & vstrDomain & vstrCriteria

  'Generate a new instance of the recordset object
  Set rstLookup = New ADODB.Recordset

  'Return all the data to the client machine
  rstLookup.CursorLocation = adUseClient
  
  'Open the selected record
  rstLookup.Open strSQL, DB

  'Provided a record was returned, set the return
  'value to the value of the required field
  If Not rstLookup.EOF Then
     varDLookup = rstLookup.Fields(IIf(alias = "", vstrField, alias))
  End If
  
  'Close the recordset and clean up memory
  rstLookup.Close
  Set rstLookup = Nothing
  
normalexit:
  Exit Function
  
errhandler:
  
  'Display the error description for the moment.
  MsgBox "varDLookup:" & strSQL & vbCrLf & Err.Description
  Resume normalexit
  
End Function

Public Sub dealDriveAndJobs(ByVal playerID)
Dim rst As New ADODB.Recordset
Dim startjobs As String, a() As String, x, msg As String

   'std Drive Core IDs 132 - 135
   DB.Execute "INSERT INTO PlayerSupplies (PlayerID,CardID) VALUES (" & playerID & ", " & 131 + playerID & ")"
   DB.Execute "Update SupplyDeck SET Seq = " & playerID & " WHERE CardID = " & 131 + playerID
   
   'get Story Issued Job
   x = Nz(varDLookup("IssueJobID", "StoryGoals", "StoryID=" & Logic!StoryID & " and Goal = 0"), 0)
   If x > 0 Then
      assignDeal playerID, x
   End If
   msg = Nz(varDLookup("Instructions", "StoryGoals", "StoryID=" & Logic!StoryID & " and Goal = 0"))
   If msg <> "" Then
      MessBox msg, "Story - First Goal", "Shiny", "", getLeader()
   End If
   
   'Grab a Job from configured list out of Contact decks
   startjobs = Nz(varDLookup("StartingJobs", "Story", "StoryID=" & Logic!StoryID), "")
   'possible future change to give optional of ALL Contact Jobs.  Use frmDeals to select 3 from the 5 Contacts
   If startjobs = "" Then Exit Sub
   
   rst.Open "SELECT * FROM ContactDeck WHERE ContactID > 0 AND Seq > " & CStr(CONSIDERED) & " ORDER BY ContactID, Seq", DB, adOpenStatic, adLockReadOnly
   
   a = Split(startjobs, ",")
   For x = LBound(a) To UBound(a)
      rst.Find "ContactID = " & a(x)
      If Not rst.EOF Then
         assignDeal playerID, rst!CardID
      End If
   Next x
   
   rst.Close
   Set rst = Nothing
   
End Sub

Public Function CrewCapacity(ByVal playerID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL

   CrewCapacity = DEF_CREWCAPACITY
   SQL = "SELECT SUM(ShipUpgrade.ExtraCrewSpace) AS Cnt FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) "
   SQL = SQL & "ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID WHERE ShipUpgrade.ExtraCrewSpace>0 AND PlayerSupplies.PlayerID=" & playerID
        
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      CrewCapacity = CrewCapacity + Nz(rst!cnt, 0)
      rst.MoveNext
   End If
   
   rst.Close
   Set rst = Nothing
   
End Function

Public Function StashCapacity(ByVal playerID) As Variant
Dim rst As New ADODB.Recordset
Dim SQL

   StashCapacity = DEF_STASHCAPACITY
   SQL = "SELECT SUM(ShipUpgrade.ExtraStashSpace) "
   SQL = SQL & "AS Cnt FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) "
   SQL = SQL & "ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID WHERE PlayerSupplies.PlayerID=" & playerID
        
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      StashCapacity = StashCapacity + Nz(rst!cnt, 0)
   End If
   
   rst.Close
   Set rst = Nothing
   
End Function

'return total Cargo + Stash space incl modifiers
Public Function CargoCapacity(ByVal playerID) As Variant
Dim rst As New ADODB.Recordset
Dim SQL

   CargoCapacity = DEF_CARGOCAPACITY + DEF_STASHCAPACITY
   SQL = "SELECT SUM(ShipUpgrade.ExtraStashSpace) "
   SQL = SQL & " + SUM(ShipUpgrade.ExtraCargoSpace) "
   SQL = SQL & "AS Cnt FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) "
   SQL = SQL & "ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID WHERE PlayerSupplies.PlayerID=" & playerID
        
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      CargoCapacity = CargoCapacity + Nz(rst!cnt, 0)
   End If
   
   rst.Close
   Set rst = Nothing
   
End Function

Public Function CargoSpaceUsed(ByVal playerID) As Variant
Dim rst As New ADODB.Recordset, x
Dim SQL

   SQL = "SELECT * FROM Players WHERE PlayerID=" & playerID
        
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      x = x + rst!fuel / 2
      x = x + rst!parts / 2
      x = x + rst!cargo
      x = x + rst!Passenger
      x = x + rst!Contraband
      x = x + rst!Fugitive
   End If
   CargoSpaceUsed = x
   rst.Close
   Set rst = Nothing
   
End Function

'Should aways be >=1 for the Leader
Public Function getCrewCount(ByVal playerID, Optional ByVal onJobOnly As Boolean = False) As Integer
Dim rst As New ADODB.Recordset
Dim SQL

   getCrewCount = 0
   SQL = "SELECT Count(Crew.CrewID) AS CrewCnt FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck "
   SQL = SQL & " ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID WHERE "
   If onJobOnly Then
      SQL = SQL & "PlayerSupplies.OffJob=0 AND "
   End If
   SQL = SQL & "PlayerSupplies.playerID = " & playerID
    
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getCrewCount = rst!crewcnt
   End If
   rst.Close
   Set rst = Nothing

End Function

Public Function getCrewName(ByVal CardID, Optional ByVal CrewID As Integer = 0) As String
Dim rst As New ADODB.Recordset
Dim SQL

   SQL = "SELECT Crew.CrewName FROM Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID WHERE "
   If CrewID > 0 Then
      SQL = SQL & " SupplyDeck.CrewID = " & CrewID
   Else
      SQL = SQL & " SupplyDeck.CardID = " & CardID
   End If
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getCrewName = rst!CrewName
   End If
   rst.Close
   Set rst = Nothing

End Function

Public Function getGearName(ByVal CardID, Optional ByVal GearID As Integer = 0) As String
Dim rst As New ADODB.Recordset
Dim SQL

   SQL = "SELECT Gear.GearName FROM Gear INNER JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID WHERE "
   If GearID > 0 Then
      SQL = SQL & " SupplyDeck.GearID = " & GearID
   Else
      SQL = SQL & " SupplyDeck.CardID = " & CardID
   End If
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getGearName = rst!GearName
   End If
   rst.Close
   Set rst = Nothing

End Function

Public Function getCrewWithNoGear(ByVal playerID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL

   getCrewWithNoGear = 0
   SQL = "SELECT PlayerSupplies_1.CardID "
   SQL = SQL & "FROM PlayerSupplies AS PlayerSupplies_1 RIGHT JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN "
   SQL = SQL & "SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies_1.CrewID = Crew.CrewID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 And PlayerSupplies.playerID = " & playerID & " AND PlayerSupplies_1.CardID IS NULL"
    
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      getCrewWithNoGear = getCrewWithNoGear + 1
      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing

End Function

Public Function getCrewID(ByVal CardID) As Integer
   getCrewID = Nz(varDLookup("CrewID", "SupplyDeck", "CardID=" & CardID), 0)
End Function

Public Function getPlanetID(ByVal playerID) As Integer
Dim SectorID
   SectorID = varDLookup("SectorID", "Players", "PlayerID=" & playerID)
   getPlanetID = Nz(varDLookup("PlanetID", "Planet", "SectorID=" & SectorID), 0)
End Function

'return extra range value.  mode 1 = fullburn, mode 2 = mosey
Public Function getRangeMod(ByVal playerID, ByVal mode As Integer) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Sum(Perk.RangeMod) AS ExtraRange"
   SQL = SQL & " FROM Perk INNER JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID)"
   SQL = SQL & " ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID"
   SQL = SQL & " WHERE PlayerSupplies.PlayerID=" & playerID & " AND Perk.RangeMod " & IIf(mode = 1, "> 0", "< 0")
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      If mode = 1 And Nz(rst!extrarange, 0) > 0 Then 'fullburns are +ve values
         getRangeMod = rst!extrarange
      ElseIf mode = 2 And Nz(rst!extrarange, 0) < 0 Then 'moseys are -ve values
         getRangeMod = Abs(rst!extrarange)
      End If
   End If
   rst.Close
   Set rst = Nothing
End Function


Public Function getPlayerJobs(ByVal playerID, Optional ByVal JobStatus) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   getPlayerJobs = 0
   SQL = "SELECT Count(PlayerJobs.CardID) AS Cnt FROM PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID WHERE ContactDeck.ContactID > 0 AND PlayerID = " & playerID
   If Not IsMissing(JobStatus) Then
      SQL = SQL & " AND JobStatus IN (" & JobStatus & ")"
   Else 'don't count completed
      SQL = SQL & " AND JobStatus <> " & JOB_SUCCESS
   End If
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getPlayerJobs = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
End Function

' returns cash + or parts -ve based on Crew Perk and JobType
Public Function getJobCrewBonus(ByVal playerID, ByVal JobType) As Integer
Dim rst As New ADODB.Recordset
Dim SQL

   getJobCrewBonus = 0
   If JobType = 0 Then Exit Function
   
   SQL = "SELECT SUM(Perk.Payment) AS Pay "
   SQL = SQL & "FROM Perk INNER JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND Perk.JobTypeID=" & JobType

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getJobCrewBonus = Nz(rst!pay, 0)
   End If
   rst.Close
   Set rst = Nothing
End Function

' returns bonus cash & parts based on Job Card and jobtype
Public Function getJobBonus(ByVal playerID, ByVal CardID, ByRef parts As Integer) As Integer
Dim rst As New ADODB.Recordset
Dim SQL

   getJobBonus = 0
   SQL = "SELECT ContactDeck.* "
   SQL = SQL & "From ContactDeck "
   SQL = SQL & "WHERE ContactDeck.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      If rst!ProfessionID > 0 Then
         If hasCrewAttribute(player.ID, cstrProfession(rst!ProfessionID)) Then
            If rst!BonusPart > 0 Then 'part takes precedence, as bonus$ may apply to other triggers
               parts = rst!BonusPart
            ElseIf rst!bonus > 0 Then
               getJobBonus = rst!bonus
            End If
         End If
      End If
      If rst!KeywordBonus = 1 Then
         If hasKeyword(playerID, rst!KeyWords) Then
            If discardGearKeyword(playerID, rst!KeyWords, True) Then
               If MessBox("You can take another $" & rst!bonus & " on this Job if you use your discardable " & rst!KeyWords & vbNewLine & "Do you want to use it up?", "Discardable Keyword", "Yes", "No", getLeader()) = 0 Then
                  discardGearKeyword playerID, rst!KeyWords
                  getJobBonus = getJobBonus + rst!bonus
               End If
            Else ' we got solid gear
               getJobBonus = getJobBonus + rst!bonus
            End If
         End If
      End If
      
      'harrow solid bonus for smuggling & shipping
      If isSolid(playerID, 6) And (rst!JobTypeID = 2 Or rst!JobTypeID = 3 Or rst!JobType2D = 2 Or rst!JobType2D = 3) Then
         getJobBonus = getJobBonus + 500
         PutMsg player.PlayName & " gets an extra $500 bonus for a Smuggling or Shipping Job due to having a Solid Rep with Lord Harrow", playerID, Logic!Gamecntr, True, 0, 0, 0, 6
      End If
      
       'fanty mingo solid bonus for transport
      If isSolid(playerID, 9) And (rst!JobTypeID = 5 Or rst!JobType2D = 5) Then
         getJobBonus = getJobBonus + 500
         PutMsg player.PlayName & " gets an extra $500 bonus for a Transport Job due to having a Solid Rep with Fanty and Mingo", playerID, Logic!Gamecntr, True, 0, 0, 0, 9
      End If
      
      If rst!BonusPerSkill > 0 Then
         getJobBonus = getJobBonus + (rst!bonus * getSkill(playerID, cstrSkill(rst!BonusPerSkill)))
      End If
      
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Sub burnFuel(ByVal playerID, Optional ByVal qty As Integer = 1)
   DB.Execute "UPDATE Players SET Fuel = Fuel-" & qty & " WHERE PlayerID = " & playerID
   frmAction.lblFuelOn.Caption = CStr(Val(frmAction.lblFuelOn.Caption) - qty)
End Sub

' returns True if Crew (eg.Nandi) can hire Crew for Free
Public Function freeCrew(ByVal playerID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT PlayerSupplies.PlayerID, Perk.FreeCrew "
   SQL = SQL & "FROM Perk INNER JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "Where PlayerSupplies.PlayerID = " & playerID & " And Perk.freeCrew = 1"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      freeCrew = True
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function discounts(ByVal playerID, ByVal perkType As String) As Single
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT PlayerSupplies.PlayerID, Perk." & perkType
   SQL = SQL & " FROM Perk INNER JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "Where PlayerSupplies.PlayerID = " & playerID & " And Perk." & perkType & " > 0"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      discounts = rst.Fields(perkType)
   End If
   rst.Close
   Set rst = Nothing
End Function

'get final balance, and optionally add & subtract money from player
Public Function getMoney(ByVal playerID, Optional ByVal change As Integer = 0) As Long
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Pay from Players "
   SQL = SQL & "Where Players.PlayerID = " & playerID
   rst.Open SQL, DB, adOpenDynamic, adLockOptimistic
   If Not rst.EOF Then
      If change <> 0 Then
         rst!pay = rst!pay + change
         rst.Update
      End If
      getMoney = rst!pay
   End If
   rst.Close
   Set rst = Nothing
End Function

'$100 per Crew, remove Disgruntled, check is for seeing what the cost would be without committing
Public Function doShoreLeave(ByVal playerID, Optional ByVal check As Boolean = False, Optional ByVal free As Boolean = False) As Integer
Dim rst As New ADODB.Recordset, cost As Integer, costpercrew As Integer, hadDis As Boolean
Dim SQL
   cost = 0
   costpercrew = IIf(free, 0, 100)
   
   SQL = "SELECT Crew.CrewID, Crew.Disgruntled "
   SQL = SQL & "FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck "
   SQL = SQL & "ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If Not (getPerkAttributeCrew(playerID, "FreeShoreLeave") > 0 And getPlanetID(playerID) > 0) Then
         cost = cost - costpercrew
      End If
      If rst!Disgruntled > 0 And Not check Then
         DB.Execute "Update Crew SET Disgruntled = 0 WHERE CrewID=" & rst!CrewID
         hadDis = True
      End If
      rst.MoveNext
   Wend
   rst.Close
   If Not check And Not free Then getMoney playerID, cost

   If free And hadDis Then
      doShoreLeave = -1
   Else
      doShoreLeave = cost
   End If
   
   Set rst = Nothing
End Function

Public Function hasDisgruntled(ByVal playerID, Optional ByVal moraleboost As Boolean = False) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Crew.CrewID, Crew.Disgruntled "
   SQL = SQL & "FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck "
   SQL = SQL & "ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID "
   SQL = SQL & "WHERE disgruntled > 0 AND PlayerSupplies.PlayerID=" & playerID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   Do While Not rst.EOF
      If moraleboost Then
         If getPerkAttributeCrew(player.ID, "RemoveDisgruntled") <> rst!CrewID Then
            hasDisgruntled = True
            Exit Do
         End If
      Else
         hasDisgruntled = True
         Exit Do
      End If
      rst.MoveNext
   Loop
   rst.Close
   Set rst = Nothing
End Function

'buy Fuel & Parts and deduct cost, or Check and return cost only
Public Function doBuyFuelParts(ByVal playerID, ByVal fuel As Integer, ByVal parts As Integer, Optional ByVal check As Boolean = False, Optional ByVal free As Integer = 0) As Long
Dim partscost As Integer, fuelcost As Integer, fueltopay
   partscost = 300
   fuelcost = 100
   
   If fuel <= 0 And parts <= 0 Then Exit Function
   
   If fuel > free Then
      fueltopay = fuel - free
   Else
      fueltopay = 0
   End If

   doBuyFuelParts = (fueltopay * fuelcost) + (parts * partscost)
   
   If Not check And (doBuyFuelParts <= getMoney(playerID) Or doBuyFuelParts = 0) Then
      DB.Execute "UPDATE Players SET Fuel = Fuel + " & fuel & ", Parts = Parts + " & parts & ", Pay = Pay - " & doBuyFuelParts & " WHERE PlayerID = " & playerID
   End If

End Function
'buy Cargo and deduct cost, or Check and return cost only
Public Function doBuyCargo(ByVal playerID, ByVal cargo As Integer, Optional ByVal check As Boolean = False) As Long

   If cargo <= 0 Then Exit Function
   
   If CargoCapacity(playerID) - CargoSpaceUsed(playerID) < cargo Then
      MessBox "Not enough cargo space for " & cargo & " Cargo", "Storage Hold", "Ooops", "", 0, 0, 6
      Exit Function
   End If

   doBuyCargo = cargo * 300
   
   If doBuyCargo > getMoney(playerID) Then
      MessBox "You cannot afford " & cargo & " Cargo for $" & doBuyCargo, "Insufficient Funds", "Ooops", "", getLeader()
      doBuyCargo = 0
      Exit Function
   End If
   
   If Not check And doBuyCargo <= getMoney(playerID) Then
      DB.Execute "UPDATE Players SET cargo = cargo + " & cargo & ", Pay = Pay - " & doBuyCargo & " WHERE PlayerID = " & playerID
   End If

End Function

Public Function doBuyContra(ByVal playerID, ByVal contra As Integer, Optional ByVal check As Boolean = False) As Long

   If contra <= 0 Then Exit Function
   If CargoCapacity(playerID) - CargoSpaceUsed(playerID) < contra Then
      MessBox "Not enough cargo space for " & contra & " Contraband", "Storage Hold", "Ooops", "", 0, 0, 6
      MsgBox "Not enough cargo space for " & contra & " Contraband", vbExclamation
      Exit Function
   End If

   doBuyContra = contra * 400
   
   If doBuyContra > getMoney(playerID) Then
      MessBox "You cannot afford " & contra & " Contraband for $" & doBuyContra, "Insufficient Funds", "Ooops", "", getLeader()
      doBuyContra = 0
      Exit Function
   End If
   
   If Not check And doBuyContra <= getMoney(playerID) Then
      DB.Execute "UPDATE Players SET Contraband = Contraband + " & contra & ", Pay = Pay - " & doBuyContra & " WHERE PlayerID = " & playerID
   End If

End Function

Public Function doSellCargoContra(ByVal playerID, ByVal ContactID, ByVal cargo As Integer, ByVal contra As Integer, Optional ByVal check As Boolean = False) As Integer
Dim car, con, perk
Dim rst As New ADODB.Recordset
Dim SQL

   'check if SOLID first
   If Not isSolid(playerID, ContactID) Then Exit Function

   'get prices
   car = varDLookup("Cargo", "Contact", "ContactID=" & ContactID)
   con = varDLookup("Contraband", "Contact", "ContactID=" & ContactID)
   If car = 0 And con = 0 Then Exit Function
   
   'get crew bargaining bonuses
   perk = getPerkAttributeValue(playerID, "GoodsBonus")
   Select Case perk
   Case 1
      con = con + 100
   Case 2
      con = con + 100
      car = car + 100
   End Select
   
   SQL = "SELECT * from Players "
   SQL = SQL & "Where Players.PlayerID = " & playerID
   rst.Open SQL, DB, adOpenDynamic, adLockOptimistic
   If Not rst.EOF Then

      'check we're not exceeding what's in the Hold
      If rst!cargo < cargo Or rst!Contraband < contra Then Exit Function
      'good to sell
      doSellCargoContra = (car * cargo) + (con * contra)
      If Not check Then
         rst!cargo = rst!cargo - cargo
         rst!Contraband = rst!Contraband - contra
         rst!pay = rst!pay + (car * cargo) + (con * contra)
         rst.Update
      End If
   End If
   
   rst.Close
   Set rst = Nothing

End Function

'mode 0 - off, 1 - ON
Public Function doChangeGear(ByVal playerID, ByVal CrewID, ByVal CardID, ByVal mode)
     'check gear doesn't already exist or is Jaynes Hat, Kaylee's Parasol or CrewID = 57 (Grange Bros) and count(CrewIDs) < 2   / Crow (59) cannot carry a Firearm
      'If mode = 1 And (Nz(varDLookup("CardID", "PlayerSupplies", "PlayerID=" & playerID & " AND CrewID = " & crewID & " AND CardID <> 21 AND CardID <> 155  AND CardID <> 157"), 0) = 0 Or CardID = 21 Or CardID = 155 Or CardID = 157 Or (crewID = 57 And getCrewGearCount(57) < 2) Or (crewID = 22 And getCrewGearCount(22) < 3)) And Not (crewID = 59 And gearHasKeyword(CardID, "FIREARM")) Then    'good to go
      
      If mode = 1 Then
         
         If noGearSlot(CardID) Then   ' we are all good, these do not count as a spot
         
         ElseIf CardID = 45 And Not hasCrewAttribute(playerID, "Tech", 0, CrewID) Then 'all good Burgess' laser need 1 Tech
            playsnd 9
            Exit Function
         
         'ElseIf CrewID = 57 And getCrewGearCount(CrewID) < 2 Then 'grange bros
         
         'ElseIf CrewID = 22 And getCrewGearCount(CrewID) < 3 Then 'jayne - up to 3
         
         ElseIf CrewID = 59 And gearHasKeyword(CardID, "FIREARM") Then ' Crow - no go
            playsnd 9
            Exit Function
            
         ElseIf getCrewGearCount(CrewID) >= hasCrewPerkAttributeValue(CrewID, "GearCount") Then 'check if the Crew has max gear
            playsnd 9
            Exit Function
         
         End If
           
         playsnd 11
         DB.Execute "UPDATE PlayerSupplies SET CrewID = " & CrewID & " WHERE CardID = " & CardID
      End If
      
      If mode = 0 Then
         playsnd 11
         DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CardID = " & CardID
      End If

End Function

'return true if this gear card uses no extra slot. eg. Jaynes Hat, Kaylees Parasol..
Public Function noGearSlot(ByVal CardID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   noGearSlot = 0
   SQL = "SELECT SupplyDeck.CardID "
   SQL = SQL & "FROM Gear INNER JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID "
   SQL = SQL & "WHERE SupplyDeck.CardID=" & CardID & " AND Gear.noGearLimit=1"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      noGearSlot = (rst!CardID = CardID)
   End If
   rst.Close
   Set rst = Nothing
End Function

'return the max value set on a Gear feature that is held by a Crew
Public Function getGearFeature(ByVal playerID, ByVal colName As String) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   getGearFeature = 0
   SQL = "SELECT MAX(Gear." & colName & ") As MaxVal "
   SQL = SQL & "FROM Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND Gear." & colName & " > 0 AND PlayerSupplies.CrewID > 0 AND PlayerSupplies.PlayerID = " & playerID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getGearFeature = Nz(rst!MaxVal, 0)
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function RollDice(Optional ByVal top As Integer = 6, Optional ByVal Thrill As Boolean = False) As Integer

   Randomize Timer
   RollDice = Int((top * Rnd) + 1)
   'CrewID 55 is Bester who blocks extra roll
   If RollDice = 6 And Thrill Then
      RollDice = RollDice + Int((top * Rnd) + 1)
   End If

End Function

Public Function cstrProfession(ByVal Profession) As String
   Select Case Profession
   Case 0 ' ""
   Case 12
      cstrProfession = "Pilot Mechanic"  'not really
   Case Else
      cstrProfession = varDLookup("ProfessionName", "Profession", "ProfessionID=" & Profession)
   End Select
      
End Function

Public Function cstrSkill(ByVal skill) As String
   Select Case skill
   Case 1
      cstrSkill = "fight"
   Case 2
      cstrSkill = "tech"
   Case 3
      cstrSkill = "negotiate"
   End Select
      
End Function

Public Function hasCrew(ByVal playerID, ByVal CrewID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   'may need to manage "On Job" status
   SQL = "SELECT PlayerSupplies.PlayerID, SupplyDeck.CrewID "
   SQL = SQL & "FROM PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND PlayerSupplies.PlayerID=" & playerID & " AND SupplyDeck.CrewID=" & CrewID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
         hasCrew = True
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function haveCrewAnyone(ByVal CrewID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   'may need to manage "On Job" status
   SQL = "SELECT PlayerSupplies.PlayerID, SupplyDeck.CrewID "
   SQL = SQL & "FROM PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE SupplyDeck.CrewID=" & CrewID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
         haveCrewAnyone = True
   End If
   rst.Close
   Set rst = Nothing
End Function

'return the first crewID that has the wanted attribute
Public Function getCrewAttribute(ByVal playerID, ByVal Attrib As String, Optional ByVal CardID As Integer = 0, Optional ByVal CrewID As Integer = 0) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   If Attrib = "" Then Exit Function
   
   'may need to manage "On Job" status
   SQL = "SELECT Crew.CrewID, Crew." & Attrib
   SQL = SQL & " FROM (Players INNER JOIN PlayerSupplies ON Players.PlayerID = PlayerSupplies.PlayerID) INNER JOIN (Crew INNER JOIN SupplyDeck "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND Players.PlayerID=" & playerID & " AND Crew." & Attrib & " <> 0"
   If CardID > 0 Then 'if the perk needs to be specific to a crew
      SQL = SQL & " AND PlayerSupplies.CardID=" & CardID
   End If
   If CrewID > 0 Then 'if the perk needs to be specific to a crew
      SQL = SQL & " AND SupplyDeck.CrewID=" & CrewID
   End If
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
         getCrewAttribute = rst!CrewID
   End If
   rst.Close
   Set rst = Nothing
End Function


Public Function countCrewAttribute(ByVal playerID, ByVal Attrib As String, Optional ByVal CardID As Integer = 0, Optional ByVal CrewID As Integer = 0) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   If Attrib = "" Then Exit Function
   
   'may need to manage "On Job" status
   SQL = "SELECT count(Crew.CrewID) AS cnt"
   SQL = SQL & " FROM (Players INNER JOIN PlayerSupplies ON Players.PlayerID = PlayerSupplies.PlayerID) INNER JOIN (Crew INNER JOIN SupplyDeck "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND Players.PlayerID=" & playerID & " AND Crew." & Attrib & " <> 0"
   If CardID > 0 Then 'if the perk needs to be specific to a crew
      SQL = SQL & " AND PlayerSupplies.CardID=" & CardID
   End If
   If CrewID > 0 Then 'if the perk needs to be specific to a crew
      SQL = SQL & " AND SupplyDeck.CrewID=" & CrewID
   End If
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
         countCrewAttribute = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function hasCrewAttribute(ByVal playerID, ByVal Attrib As String, Optional ByVal CardID As Integer = 0, Optional ByVal CrewID As Integer = 0) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   If Attrib = "" Then Exit Function
   If Attrib = "Pilot Mechanic" Then
      If Not hasCrewAttribute(playerID, "Pilot", CardID, CrewID) Then Exit Function
      Attrib = "Mechanic"
   End If
   'may need to manage "On Job" status
   SQL = "SELECT Crew." & Attrib
   SQL = SQL & " FROM (Players INNER JOIN PlayerSupplies ON Players.PlayerID = PlayerSupplies.PlayerID) INNER JOIN (Crew INNER JOIN SupplyDeck "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND Players.PlayerID=" & playerID & " AND Crew." & Attrib & " <> 0"
   If CardID > 0 Then 'if the perk needs to be specific to a crew
      SQL = SQL & " AND PlayerSupplies.CardID=" & CardID
   End If
   If CrewID > 0 Then 'if the perk needs to be specific to a crew
      SQL = SQL & " AND SupplyDeck.CrewID=" & CrewID
   End If
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
         hasCrewAttribute = True
   End If
   
   If LCase(Attrib) = "companion" And Not hasCrewAttribute Then 'check for Inara's Guild Papers
       hasCrewAttribute = hasGearCrew(playerID, 36)
   End If
   rst.Close
   Set rst = Nothing
End Function

' returns the (first) CrewID if that crew has a given Perk column attribute, may be more than one crew that has it tho
Public Function getPerkAttributeCrew(ByVal playerID, ByVal Attrib As String, Optional ByVal CardID As Integer = 0, Optional ByVal CrewID As Integer = 0) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   If Attrib = "" Then Exit Function
   'may need to manage "On Job" status
   SQL = "SELECT SupplyDeck.CrewID, Perk." & Attrib
   SQL = SQL & " FROM Perk INNER JOIN (PlayerSupplies INNER JOIN (Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND PlayerSupplies.PlayerID=" & playerID & " AND Perk." & Attrib & " <> 0"
   If CardID > 0 Then 'if the perk needs to be specific to a crew
      SQL = SQL & " AND PlayerSupplies.CardID=" & CardID
   End If
   If CrewID > 0 Then 'if the perk needs to be specific to a crew
      SQL = SQL & " AND SupplyDeck.CrewID=" & CrewID
   End If
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       getPerkAttributeCrew = rst!CrewID
   End If
   rst.Close
   Set rst = Nothing
End Function
'return the value of the Perk attribute of the first crew with it
Public Function getPerkAttributeValue(ByVal playerID, ByVal Attrib As String) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   If Attrib = "" Then Exit Function
   'may need to manage "On Job" status
   SQL = "SELECT SupplyDeck.CrewID, Perk." & Attrib
   SQL = SQL & " FROM Perk INNER JOIN (PlayerSupplies INNER JOIN (Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND PlayerSupplies.PlayerID=" & playerID & " AND Perk." & Attrib & " <> 0"
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       getPerkAttributeValue = rst.Fields(Attrib)
   End If
   rst.Close
   Set rst = Nothing
End Function
'return True if the player has any crew with a particular value in a given attribute
Public Function hasPerkAttributeValue(ByVal playerID, ByVal Attrib As String, ByVal perkValue As Integer) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   If Attrib = "" Then Exit Function
   'may need to manage "On Job" status
   SQL = "SELECT SupplyDeck.CrewID, Perk." & Attrib
   SQL = SQL & " FROM Perk INNER JOIN (PlayerSupplies INNER JOIN (Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND PlayerSupplies.PlayerID=" & playerID & " AND Perk." & Attrib & " = " & CStr(perkValue)
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       hasPerkAttributeValue = True
   End If
   rst.Close
   Set rst = Nothing
End Function

'grab the Perk attribute value for a specific Crew member
Public Function hasCrewPerkAttributeValue(ByVal CrewID, ByVal Attrib As String) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   If Attrib = "" Then Exit Function
   SQL = "SELECT Perk." & Attrib
   SQL = SQL & " FROM Perk INNER JOIN Crew ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "WHERE CrewID=" & CrewID & " AND Perk." & Attrib & " <> 0"
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       hasCrewPerkAttributeValue = rst.Fields(Attrib)
   End If
   rst.Close
   Set rst = Nothing
End Function

'returns the perk keywork for a cardID (crew)
Public Function hasPerkKeyword(ByVal playerID, ByVal CardID As Integer) As String
Dim rst As New ADODB.Recordset
Dim SQL
   
   'may need to manage "On Job" status
   SQL = "SELECT Perk.Keyword "
   SQL = SQL & "FROM Perk INNER JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND PlayerSupplies.PlayerID=" & playerID & " AND PlayerSupplies.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       hasPerkKeyword = rst!Keyword & ""
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function hasGear(ByVal playerID, ByVal GearID As Integer, Optional ByVal CrewID As Integer = 0) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   'may need to manage "Off Job" status
   SQL = "SELECT PlayerSupplies.PlayerID, PlayerSupplies.CrewID, Gear.GearID "
   SQL = SQL & "FROM Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND PlayerSupplies.PlayerID=" & playerID & " AND Gear.GearID=" & GearID
   If CrewID > 0 Then
      SQL = SQL & " AND PlayerSupplies.CrewID=" & CrewID
   Else
      SQL = SQL & " AND PlayerSupplies.CrewID<>0"
   End If
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       hasGear = True
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function hasGearCard(ByVal playerID, ByVal GearID As Integer, Optional ByVal CrewID As Integer = 0) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   'may need to manage "Off Job" status
   SQL = "SELECT PlayerSupplies.CardID, PlayerSupplies.CrewID, Gear.GearID "
   SQL = SQL & "FROM Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND PlayerSupplies.PlayerID=" & playerID & " AND Gear.GearID=" & GearID
   If CrewID > 0 Then
      SQL = SQL & " AND PlayerSupplies.CrewID=" & CrewID
   Else
      SQL = SQL & " AND PlayerSupplies.CrewID<>0"
   End If
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       hasGearCard = rst!CardID
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function hasGearCrew(ByVal playerID, ByVal GearID As Integer) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   'may need to manage "Off Job" status
   SQL = "SELECT PlayerSupplies.CardID, PlayerSupplies.CrewID, Gear.GearID "
   SQL = SQL & "FROM Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND PlayerSupplies.PlayerID=" & playerID & " AND Gear.GearID=" & GearID

   SQL = SQL & " AND PlayerSupplies.CrewID<>0"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       hasGearCrew = rst!CrewID
   End If
   rst.Close
   Set rst = Nothing
End Function

'return true if a player that has an active crew, or with linked gear, that has a keyword
Public Function hasKeyword(ByVal playerID, ByVal Keyword As String) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   If Keyword = "" Then Exit Function
   'may need to manage "Off Job" status
   SQL = "SELECT PlayerSupplies.PlayerID "
   SQL = SQL & "FROM (((Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) "
   SQL = SQL & "LEFT JOIN PlayerSupplies AS PlayerSupplies_1 ON Crew.CrewID = PlayerSupplies_1.CrewID) LEFT JOIN SupplyDeck AS SupplyDeck_1 ON PlayerSupplies_1.CardID = SupplyDeck_1.CardID) LEFT JOIN Gear ON SupplyDeck_1.GearID = Gear.GearID "
   SQL = SQL & "WHERE (PlayerSupplies.PlayerID=" & playerID & " AND PlayerSupplies.OffJob=0 AND Gear.Keywords Like '%" & Keyword & "%') OR (PlayerSupplies.PlayerID=" & playerID & " AND PlayerSupplies.OffJob=0 AND Crew.Keywords Like '%" & Keyword & "%')"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
         hasKeyword = True
   End If
   rst.Close
   
   If Not hasKeyword And Keyword = "TRANSPORT" Then 'has pilot and a skyhook
      If hasShipUpgrade(playerID, 10) And hasCrewAttribute(playerID, "Pilot") Then
         hasKeyword = True
      End If
   End If
   
   Set rst = Nothing
End Function


'return true for a player's crew's gear attrib <>0
Public Function hasGearAttribute(ByVal playerID, ByVal Attrib As String, Optional ByVal GearID As Integer = 0) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   If Attrib = "" Then Exit Function
   ' "On Job" status is managed from CrewID
   SQL = "SELECT SupplyDeck.CardID, PlayerSupplies.CrewID, Gear.* "
   SQL = SQL & "FROM Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND Gear." & Attrib & " <> 0"
   
   If GearID > 0 Then 'if the Attrib needs to be specific to Gear
      SQL = SQL & " AND Gear.GearID=" & GearID
   End If
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       hasGearAttribute = rst.Fields(Attrib)
   End If
   rst.Close
   Set rst = Nothing
End Function

'return true if a player has any crew's gear that has a keyword
Public Function hasGearKeyword(ByVal playerID, ByVal Keyword As String, Optional ByVal CrewID As Integer = 0) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   If Keyword = "" Then Exit Function
   ' "On Job" status is managed from CrewID
   SQL = "SELECT SupplyDeck.CardID, PlayerSupplies.CrewID, Gear.Keywords "
   SQL = SQL & "FROM Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND Gear.Keywords Like '%" & Keyword & "%'"
   
   If CrewID > 0 Then 'if the perk needs to be specific to a crew
      SQL = SQL & " AND PlayerSupplies.CrewID=" & CrewID
   End If
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       hasGearKeyword = True
   End If
   rst.Close
   Set rst = Nothing
End Function

'return true if a specific gear item carries a keyword
Public Function gearHasKeyword(ByVal CardID, ByVal Keyword As String) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   If Keyword = "" Then Exit Function
   'may need to manage "Off Job" status
   SQL = "SELECT SupplyDeck.CardID "
   SQL = SQL & "FROM Gear INNER JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID "
   SQL = SQL & "Where SupplyDeck.cardID = " & CardID & " And Gear.Keywords Like '%" & Keyword & "%'"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       gearHasKeyword = True
   End If
   rst.Close
   Set rst = Nothing
End Function

'return true if a gear was discarded with needed keyword
Public Function discardGearKeyword(ByVal playerID, ByVal Keyword As String, Optional ByVal checkonly As Boolean = False) As Boolean
Dim rst As New ADODB.Recordset, CardID As Integer, found As Boolean
Dim SQL
   SQL = "SELECT  Crew.CrewID, Gear.Discard, Gear.Keywords as GearWord, Crew.Keywords as CrewWord, SupplyDeck_1.CardID "
   SQL = SQL & "FROM (((Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) "
   SQL = SQL & "LEFT JOIN PlayerSupplies AS PlayerSupplies_1 ON Crew.CrewID = PlayerSupplies_1.CrewID) LEFT JOIN SupplyDeck AS SupplyDeck_1 ON PlayerSupplies_1.CardID = SupplyDeck_1.CardID) LEFT JOIN Gear ON SupplyDeck_1.GearID = Gear.GearID "
   SQL = SQL & "WHERE (PlayerSupplies.PlayerID=" & playerID & " AND PlayerSupplies.OffJob=0 AND Gear.Keywords Like '%" & Keyword & "%') OR (PlayerSupplies.PlayerID=" & playerID & " AND PlayerSupplies.OffJob=0 AND Crew.Keywords Like '%" & Keyword & "%')"
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If InStr(Nz(rst!CrewWord), Keyword) > 0 Then 'we have a solid keyword, no need to discard
         found = True
      ElseIf InStr(Nz(rst!GearWord), Keyword) > 0 And Nz(rst!discard, 0) = 0 Then 'found a solid gear keyword
         found = True
      ElseIf InStr(Nz(rst!GearWord), Keyword) > 0 And Nz(rst!discard, 0) = 1 Then  'has a discard for this keyword
         CardID = rst!CardID
      End If
      rst.MoveNext
   Wend
   rst.Close
   
   If Not found And CardID > 0 Then
      'discardable keyword in use
      discardGearKeyword = True
     If Not checkonly Then doDiscardGear playerID, CardID
   End If
   
   Set rst = Nothing
End Function

Public Function hasShipUpgrade(ByVal playerID, ByVal ShipUpgradeID As Integer) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   'may need to manage "On Job" status
   SQL = "SELECT PlayerSupplies.CardID "
   SQL = SQL & "FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND ShipUpgrade.ShipUpgradeID=" & ShipUpgradeID
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then 'find the first one, as could be multi (eg.Cry Baby)
       hasShipUpgrade = rst!CardID
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function doMercDiscard(ByVal playerID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   
   'may need to manage "On Job" status
   SQL = "SELECT SupplyDeck.CardID, SupplyDeck.CrewID "
   SQL = SQL & "FROM (Players INNER JOIN PlayerSupplies ON Players.PlayerID = PlayerSupplies.PlayerID) INNER JOIN (Crew INNER JOIN SupplyDeck "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND Players.PlayerID=" & playerID
   SQL = SQL & " AND Crew.Merc = 1"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      'update their pile status - 0 removed, 5 -discarded
      DB.Execute "UPDATE SupplyDeck SET Seq =5 WHERE CardID = " & rst!CardID
      'remove any Gear first
      DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & rst!CrewID
      'delete the card to the players deck
      DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID =" & playerID & " AND CardID = " & rst!CardID
      doMercDiscard = True
      rst.MoveNext
   Wend

End Function

'mercs = 1 count then, mercs=2 count e'ryone else
Public Function getSkill(ByVal playerID, ByVal skill As String, Optional ByVal mercs As Integer = 0, Optional ByVal noDiscards As Boolean = False, Optional ByVal kosher As Boolean = False) As Integer
Dim rst As New ADODB.Recordset, rst2 As New ADODB.Recordset
Dim SQL
   getSkill = 0
   'may need to manage "On Job" status
   SQL = "SELECT SupplyDeck.CardID, Crew.* "
   SQL = SQL & "FROM (Players INNER JOIN PlayerSupplies ON Players.PlayerID = PlayerSupplies.PlayerID) INNER JOIN (Crew INNER JOIN SupplyDeck "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND Players.PlayerID=" & playerID
   
   If mercs = 1 Then
      SQL = SQL & " AND Crew.Merc = 1"
   ElseIf mercs = 2 Then
      SQL = SQL & " AND Crew.Merc = 0"
   End If

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      getSkill = getSkill + rst.Fields(skill)
      'check for perk to add skill
      
      '+1 skill When carrying a Keyword (Firearm / Explosives)
      If getPerkAttributeCrew(playerID, skill, rst!CardID) > 0 And hasGearKeyword(playerID, hasPerkKeyword(playerID, rst!CardID), rst!CrewID) Then
         getSkill = getSkill + 1
      End If

      If rst!HillFolk = 1 And mercs = 0 Then
         'check for HillFolk fight bonus
         If countCrewAttribute(playerID, "HillFolk") > 2 And skill = cstrSkill(1) Then getSkill = getSkill + 1
      End If
      'Head Goon
      If countCrewAttribute(playerID, "Merc") > 2 And rst!CrewID = 65 And skill = cstrSkill(3) Then
         getSkill = getSkill + 2
      End If

      'no kosherised rule or its Lund who can have gear counted
      If Not kosher Or rst!CrewID = 60 Then
         'grab skill from gear crew is carrying-----------------------------
         SQL = "SELECT Gear.* "
         SQL = SQL & "FROM (Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID) INNER JOIN Crew ON PlayerSupplies.CrewID = Crew.CrewID "
         SQL = SQL & "WHERE PlayerSupplies.CrewID=" & rst!CrewID
         If noDiscards Then
            SQL = SQL & " AND Gear.Discard=0"
         End If
         
         If mercs = 1 Then
            SQL = SQL & " AND Crew.Merc = 1"
         ElseIf mercs = 2 Then
            SQL = SQL & " AND Crew.Merc = 0"
         End If
         
         rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
         While Not rst2.EOF
            getSkill = getSkill + rst2.Fields(skill)
            rst2.MoveNext
         Wend
         rst2.Close
         '------------------------------------------------------------
      Else 'koshized rules apply, with exceptions
         If hasGear(playerID, 37) And LCase(skill) = "fight" Then
            getSkill = getSkill + 1
         End If
      
      End If
      
      
      rst.MoveNext
   Wend
   rst.Close

   'Foreman + 2 mudders
   If countCrewAttribute(playerID, "Mudder") > 2 And skill = cstrSkill(1) And hasCrew(playerID, 76) Then getSkill = getSkill + 2
   
   Set rst = Nothing
End Function

Public Function getSkillDiscards(ByVal playerID, ByVal skill As String, Optional ByVal kosher As Boolean = False) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   getSkillDiscards = 0
 
   'grab all gear crew is carrying that has to be discared after use
   SQL = "SELECT Gear.* "
   SQL = SQL & "FROM (Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID) INNER JOIN Crew ON PlayerSupplies.CrewID = Crew.CrewID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID
   SQL = SQL & " AND Gear.Discard=1"
   If kosher Then
      SQL = SQL & " AND PlayerSupplies.CrewID = 60" ' Lund
   End If

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      getSkillDiscards = getSkillDiscards + rst.Fields(skill)
      rst.MoveNext
   Wend
   rst.Close
   
   Set rst = Nothing
End Function

Public Function doCustomStory(Optional ByVal readonly As Boolean = False)
Dim frmStory As New frmStories
   With frmStory
      .StoryID = Logic!StoryID
      If readonly Then .readonly
      .Show 1
      doCustomStory = .StoryID
         
   End With
End Function

'NB: for only one crew, supply crewID and use -1 for remove all, 1-add to Moral Only, 2 for add all disgruntled, 3-remove disgruntle Moral Only
Public Function doDisgruntled(ByVal playerID, ByVal mode, Optional ByVal CrewID As Integer = 0) As Integer
Dim rst As New ADODB.Recordset
Dim SQL, leader

   If mode = 6 Then mode = 1 'moral crew only (on win)

   'see if Leader is recieving a 2nd disgruntle -all crew fired!
   leader = varDLookup("Leader", "Players", "PlayerID=" & playerID)

   
   SQL = "SELECT Crew.*, SupplyDeck.CardID "
   SQL = SQL & "FROM (Players INNER JOIN PlayerSupplies ON Players.PlayerID = PlayerSupplies.PlayerID) INNER JOIN (Crew INNER JOIN SupplyDeck "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND Players.PlayerID=" & playerID
   
   If hasCrewAttribute(playerID, "Disgruntled", 0, leader) And (mode = 2 Or (mode = 1 And hasCrewAttribute(playerID, "Moral", 0, leader))) And (CrewID = 0 Or CrewID = leader) Then
      mode = 4 'all fired regardless of original mode
      If getCrewCount(playerID) > 1 Then
         PutMsg player.PlayName & "'s entire Crew has been Fired by a disgruntled Captain!", playerID, Logic!Gamecntr, True, leader
      End If
      
   Else
      If mode = 1 Or mode = 3 Then  'moral only
         SQL = SQL & " AND Moral =1"
      End If
      
      If CrewID <> 0 Then 'just do one crew
         SQL = SQL & " AND SupplyDeck.CrewID=" & CrewID
      End If
      
   End If
      
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF

      Select Case mode
         Case -1, 3 'remove all/moral only disgruntled
            DB.Execute "UPDATE Crew SET Disgruntled = 0 WHERE CrewID = " & rst!CrewID
         Case 1, 2, 4 'moral only / all
            If rst!Moral = 1 Or mode >= 2 Then
               If rst!Disgruntled > 0 Or mode = 4 Then 'if this is the 2nd disgruntled or fired then crew go to discard pile, gear returns to ship
                  'remove Disgruntled
                  DB.Execute "UPDATE Crew SET Disgruntled = 0 WHERE CrewID = " & rst!CrewID
                  If rst!CrewID <> leader Then
                     PutMsg player.PlayName & "'s Crew member " & rst!CrewName & " left the Ship, fully disgruntled", playerID, Logic!Gamecntr
                     'remove any Gear from Crew
                     DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE PlayerID = " & playerID & " AND CrewID = " & rst!CrewID
                     'remove from players hand
                     DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID = " & playerID & " AND CardID = " & rst!CardID
                     'return to discard pile
                     DB.Execute "UPDATE SupplyDeck SET Seq = 5 WHERE CardID =" & rst!CardID
                  End If
               Else 'add a Disgruntled
                  DB.Execute "UPDATE Crew SET Disgruntled = 1 WHERE CrewID = " & rst!CrewID
               End If
            End If
         
      End Select

      rst.MoveNext
   Wend
   rst.Close
   doDisgruntled = mode
      
   Set rst = Nothing
End Function

Public Function removeSelDisgruntled(ByVal playerID) As Boolean
Dim frmCrew As New frmCrewSel, CrewID, filter
   
   If hasGear(playerID, 27) Then 'lovebot
      filter = ""
   Else 'otherwise exclude the first crew that has the remove disgruntled perk
      CrewID = getPerkAttributeCrew(player.ID, "RemoveDisgruntled")
      filter = " AND SupplyDeck.CrewID <> " & CrewID
   End If
 
   frmCrew.crewFilter = " INNER JOIN (PlayerSupplies INNER JOIN  SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID WHERE PlayerSupplies.PlayerID=" & playerID & " AND Crew.Disgruntled=1" & filter
   frmCrew.Caption = "Pick a Crew to remove Disgruntle from.."
   frmCrew.Show 1
   DB.Execute "UPDATE Crew SET Disgruntled = 0 WHERE CrewID =" & GetCombo(frmCrew.cboCrew)
   
   Set frmCrew = Nothing

End Function

'if player has a Crew carrying the GearID, then set their disgruntle to 0
Public Sub removeDigruntled(ByVal playerID, ByVal skill)
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT PlayerSupplies.PlayerID, Crew.CrewID, Crew.CrewName, Gear.GearID, Gear.GearName "
   SQL = SQL & "FROM PlayerSupplies AS PlayerSupplies_1 INNER JOIN (SupplyDeck AS SupplyDeck_1 INNER JOIN ((Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID) INNER JOIN Crew ON PlayerSupplies.CrewID = Crew.CrewID) ON SupplyDeck_1.CrewID = Crew.CrewID) ON PlayerSupplies_1.CardID = SupplyDeck_1.CardID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND Crew.Disgruntled=1 AND Gear.RemoveDisgruntle=" & skill & " AND PlayerSupplies_1.OffJob=0"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      DB.Execute "UPDATE Crew SET Disgruntled=0 WHERE CrewID=" & rst!CrewID
      PutMsg player.PlayName & "'s " & rst!CrewName & " is no longer Disgruntled thanks to " & rst!GearName, playerID, Logic!Gamecntr, True, 0, rst!GearID
      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing

End Sub

Public Function getShipUpgradeID(ByVal CardID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT ShipUpgrade.ShipUpgradeID "
   SQL = SQL & "FROM ShipUpgrade INNER JOIN SupplyDeck ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
   SQL = SQL & "WHERE SupplyDeck.CardID= " & CardID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getShipUpgradeID = rst!ShipUpgradeID
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function getShipUpgrades(ByVal playerID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT count(PlayerSupplies.CardID) as cnt "
   SQL = SQL & "FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND ShipUpgrade.DriveCore=0"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getShipUpgrades = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function isDriveCore(ByVal CardID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT SupplyDeck.CardID, ShipUpgrade.DriveCore "
   SQL = SQL & "FROM ShipUpgrade INNER JOIN SupplyDeck ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
   SQL = SQL & "WHERE ShipUpgrade.DriveCore=1 AND SupplyDeck.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      isDriveCore = True
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Sub removeDriveCore(ByVal playerID)
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT PlayerSupplies.CardID, SupplyDeck.SupplyID "
   SQL = SQL & "FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND ShipUpgrade.DriveCore=1"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      DB.Execute "UPDATE SupplyDeck SET Seq = " & IIf(rst!SupplyID > 0, "5", "0") & " WHERE CardID = " & rst!CardID
      DB.Execute "DELETE FROM PlayerSupplies WHERE CardID = " & rst!CardID
   End If
   rst.Close
   Set rst = Nothing
End Sub

'from Nav option, load extra Salvage as per the Perks
Public Sub doSalvage(ByVal playerID)
Dim rst As New ADODB.Recordset
Dim SQL, u, v As Integer, w As Integer, x As Integer, y As Integer, z As Integer, msg
   SQL = "SELECT Perk.* "
   SQL = SQL & "FROM Perk INNER JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID = " & playerID
 
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If rst!SOCargo > 0 Then
         w = w + rst!SOCargo
      End If
      If rst!SOContraband > 0 Then
         x = x + rst!SOContraband
      End If
      If rst!SOFuel > 0 Then
         y = y + rst!SOFuel
      End If
      If rst!SOPart > 0 Then
         z = z + rst!SOPart
      End If
      
      rst.MoveNext
   Wend
   rst.Close
   If hasShipUpgrade(playerID, 9) > 0 Then
      x = x + 1
      DB.Execute "UPDATE Players SET Pay = Pay + 500 WHERE PlayerID = " & playerID
      msg = msg & "and Chop Shop added $500, "
   End If
   
   v = CargoCapacity(playerID)
   u = CargoSpaceUsed(playerID)
   
   If y > 0 Then
      If (y / 2 + u) <= v Then
         DB.Execute "UPDATE Players SET Fuel = Fuel + " & y & " WHERE PlayerID = " & playerID
         msg = msg & "added " & y & " Fuel, "
         u = (y / 2 + u)
      End If
   End If
   If z > 0 Then
      If (z / 2 + u) <= v Then
         DB.Execute "UPDATE Players SET Parts = Parts + " & z & " WHERE PlayerID = " & playerID
         msg = msg & "added " & z & " Parts, "
         u = (z / 2 + u)
      End If
   End If
   
   If w > 0 Then
      If w + u <= v Then
         DB.Execute "UPDATE Players SET Cargo = Cargo + " & w & " WHERE PlayerID = " & playerID
         msg = msg & "added " & w & " Cargo, "
         u = w + u
      End If
   End If
   If x > 0 Then
      If x + u <= v Then
         DB.Execute "UPDATE Players SET Contraband = Contraband + " & x & " WHERE PlayerID = " & playerID
         msg = msg & "added " & x & " Contraband "
         u = x + u
      End If
   End If
   If Nz(msg) <> "" Then
      PutMsg player.PlayName & "'s Crew " & msg & "to the Salvage Op", playerID, Logic!Gamecntr
   End If
      
   Set rst = Nothing
End Sub

Public Function SeizeAllContraFugi(ByVal playerID) As Boolean
Dim contra, fugi, stash, x

   contra = varDLookup("Contraband", "Players", "PlayerID=" & playerID)
   fugi = varDLookup("Fugitive", "Players", "PlayerID=" & playerID)
   stash = StashCapacity(playerID)
   
   If contra + fugi > stash Then 'we gotta problem
      SeizeAllContraFugi = True
      x = stash - fugi 'give priority to fugi
      If x < 0 Then      'we got more fugi than stash, so all cargo goes, and some fugi
         DB.Execute "UPDATE Players SET Contraband= 0, Fugitive=" & stash & "  WHERE PlayerID =" & playerID
      Else 'can hold all of Fugi and some of contra
         DB.Execute "UPDATE Players SET Contraband= " & x & " WHERE PlayerID =" & playerID
      End If
   End If

End Function

Public Function SeizeAllFugi(ByVal playerID) As Boolean
Dim fugi, stash, x

   fugi = varDLookup("Fugitive", "Players", "PlayerID=" & playerID)
   stash = StashCapacity(playerID)
   
   If fugi > stash Then 'we gotta problem
      SeizeAllFugi = True
      x = stash - fugi 'give priority to fugi
      If x < 0 Then      'we got more fugi than stash
         DB.Execute "UPDATE Players SET Fugitive=" & stash & "  WHERE PlayerID =" & playerID
      End If
   End If

End Function

Public Function SeizeAllContraCargo(ByVal playerID) As Boolean
Dim contra, cargo, stash, x

   contra = varDLookup("Contraband", "Players", "PlayerID=" & playerID)
   cargo = varDLookup("cargo", "Players", "PlayerID=" & playerID)
   stash = StashCapacity(playerID)
   
   If contra + cargo > stash Then 'we gotta problem
      SeizeAllContraCargo = True
      x = stash - contra 'give priority to contra
      If x < 0 Then      'we got more contra than stash, so all cargo goes, and some contra
         DB.Execute "UPDATE Players SET cargo= 0, Contraband=" & stash & "  WHERE PlayerID =" & playerID
      Else 'can hold all of Fugi and some of contra
         DB.Execute "UPDATE Players SET cargo= " & x & " WHERE PlayerID =" & playerID
      End If
   End If

End Function

Public Function doKillAllCrew(ByVal playerID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT SupplyDeck.CardID, Crew.* "
   SQL = SQL & "FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " ORDER BY Crew.Leader" 'leader last

   rst.Open SQL, DB, adOpenStatic, adLockReadOnly
   While Not rst.EOF
      If rst!leader = 1 Then
         doDisgruntled playerID, 2, rst!CrewID
      Else
         doKillCrew playerID, rst!CardID
      End If
      rst.MoveNext
   Wend
   rst.Close

End Function

Public Function doKillCrew(ByVal playerID, ByVal CardID) As Integer
Dim result As Integer, x As Integer, CrewID

   CrewID = varDLookup("CrewID", "SupplyDeck", "CardID=" & CardID)
   
   If CrewID = varDLookup("Leader", "Players", "PlayerID=" & playerID) Then
      doDisgruntled playerID, 2, CrewID
      Exit Function
   End If

   'Simon Tam adds 2 to Dice Roll for Medic Check
   If getPerkAttributeCrew(playerID, "MedicCheck") > 0 Then x = 2
   'Simon Tam's bag adds 1 to Dice Roll for Medic Check
   If hasGear(playerID, 19) Then x = x + 1
   'Medic check to save Crew
   If hasCrewAttribute(playerID, "Medic") And (RollDice(6) + x) > 4 Then
      PutMsg player.PlayName & "'s Crew member " & getCrewName(CardID) & " was saved by our Medic", playerID, Logic!Gamecntr, True, CrewID
      Exit Function
   End If
   
   If hasCrewAttribute(playerID, "Medic") And hasShipUpgrade(playerID, 8) > 0 And (RollDice(6) + x) > 4 Then
      PutMsg player.PlayName & "'s Crew member " & getCrewName(CardID) & " was saved in our Fully Equipped Med Bay", playerID, Logic!Gamecntr, True, CrewID
      Exit Function
   End If
   
   If hasGear(playerID, 49, CrewID) Then
      doDiscardGear player.ID, hasGearCard(player.ID, 49, CrewID)
      PutMsg player.PlayName & "'s Crew member " & getCrewName(CardID) & " was saved by Med Foam, which once used, had to be discarded", playerID, Logic!Gamecntr, True, CrewID
      Exit Function
   End If
   
   If hasGear(playerID, 46, CrewID) Then
      doDiscardGear player.ID, hasGearCard(player.ID, 46, CrewID)
      PutMsg player.PlayName & "'s Crew member " & getCrewName(CardID) & " was saved by Zoe's Flak Jacket, which then had to be discarded", playerID, Logic!Gamecntr, True, CrewID
      Exit Function
   End If
   
   '-----==== DEAD =====------ R.I.P.
   
   'some crew go back to discard pile, "KillDiscard" Perk
   'eg.When Killed, Discard instead of removing from Play
   If getPerkAttributeCrew(playerID, "KillDiscard", CardID) > 0 Then
      result = 5
      PutMsg player.PlayName & "'s Crew member " & getCrewName(CardID) & " left your employment", playerID, Logic!Gamecntr, True, CrewID
   Else
      PutMsg player.PlayName & "'s Crew member " & getCrewName(CardID) & " was Killed.  RIP.", playerID, Logic!Gamecntr, True, CrewID
   End If
   
   'update their pile status - 0 removed, 5 -discarded
   DB.Execute "UPDATE SupplyDeck SET Seq =" & result & " WHERE CardID = " & CardID
   'remove any Gear first
   DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & CrewID
   'delete the card to the players deck
   DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID =" & playerID & " AND CardID = " & CardID
   doKillCrew = 1
   
End Function

Public Function doSeizeCrew(ByVal playerID, ByVal CardID, ByVal wanted) As Integer
Dim result As Integer, CrewID

   CrewID = varDLookup("CrewID", "SupplyDeck", "CardID=" & CardID)
   
   If hasGear(playerID, 20, CrewID) Then
      PutMsg player.PlayName & "'s Crew member " & getCrewName(CardID) & " flashes their fake Alliance Ident.Card", playerID, Logic!Gamecntr, True, CrewID
      Exit Function
   End If

   If RollDice(6) > 1 Then
      If wanted > 1 Then
         If RollDice(6) = 1 Then '2nd roll for Grange Bros
            'busted
         Else
            PutMsg player.PlayName & "'s Crew members " & getCrewName(CardID) & " both managed to narrowly AVOID detection and arrest.", playerID, Logic!Gamecntr
            Exit Function
         End If
      Else
         PutMsg player.PlayName & "'s Crew member " & getCrewName(CardID) & " narrowly managed to AVOID detection and arrest.", playerID, Logic!Gamecntr
         Exit Function
      End If
   End If
   
   PutMsg player.PlayName & "'s Crew member " & getCrewName(CardID) & " was Seized by the Alliance", playerID, Logic!Gamecntr, True, CrewID, 0, 0, 0, 0, 1

   result = 0 ' same as killed
   
   'update their pile status - 0 removed, 5 -discarded
   DB.Execute "UPDATE SupplyDeck SET Seq =" & result & " WHERE CardID = " & CardID
   'remove any Gear first
   DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & CrewID
   'delete the card to the players deck
   DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID =" & playerID & " AND CardID = " & CardID

   
End Function

'update their pile status - 0 removed, 5 -discarded
Public Sub doDiscardCrew(ByVal CardID, Optional ByVal status As Variant = 5)
Dim CrewID

   CrewID = varDLookup("CrewID", "SupplyDeck", "CardID=" & CardID)

   
   DB.Execute "UPDATE SupplyDeck SET Seq = " & status & " WHERE CardID = " & CardID
   'remove any Gear first
   DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & CrewID
   'delete the card to the players deck
   DB.Execute "DELETE FROM PlayerSupplies WHERE CardID = " & CardID

   
End Sub

'use for linked, unlinked Gear and any ShipUpgrade
Public Sub doDiscardGear(ByVal playerID, ByVal CardID)

   ' Grimey - When you discard a Gear Card, roll a dice. If you roll a 6, keep the Gear Card. Otherwise, discard it normally
   If Nz(varDLookup("GearID", "SupplyDeck", "CardID=" & CardID), 0) > 0 And hasCrew(playerID, 77) Then
      If RollDice(6) = 6 Then
         If MessBox("Grimey the 'Errand Boy' can retrieve " & getGearName(CardID) & " for you" & vbNewLine & "Do you want it back?", "Gear return", "Yes", "No", 77) = 0 Then
         'If MsgBox("Grimey the 'Errand Boy' can retrieve " & getGearName(CardID) & " for you" & vbNewLine & "Do you want it back?", vbYesNo + vbQuestion, "Gear return") = vbYes Then
            PutMsg player.PlayName & " gets Grimey to retrieve " & getGearName(CardID), player.ID, Logic!Gamecntr
            Exit Sub
         End If
      Else
         PutMsg player.PlayName & " was unable to get Grimey to retrieve " & getGearName(CardID), player.ID, Logic!Gamecntr
      End If
   End If
   
   playsnd 12
   'update their pile status - 0 removed, 5 -discarded
   DB.Execute "UPDATE SupplyDeck SET Seq =5 WHERE CardID = " & CardID

   'delete the card to the players deck
   DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID =" & playerID & " AND CardID = " & CardID

   
End Sub

' Return True if Cruiser was faced.
Public Function doMoveAlliance(ByVal playerID, ByVal SectorID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL, pay As Integer, CryBaby As Integer, crewcnt As Integer
   
   'Move Crusier to your Sector,
   If getCruiserSector <> SectorID Then
      MoveShip 5, SectorID
   End If
   
   If getHaven(SectorID) > 0 Then
         PutMsg player.PlayName & "'s Nav log: refuge found at this Haven, the Alliance Cruiser sails on by", player.ID, Logic!Gamecntr, True, 0, 0, 1
         moveAutoAI 5
         Exit Function
   End If
   
   'does player have Cry Baby?
   CryBaby = hasShipUpgrade(playerID, 1)
   If CryBaby > 0 Then
      If MessBox("Do you want to deploy (and discard) the Cry Baby to decoy the Alliance Cruisier?", "Alliance Cruiser Alert!", "Deploy", "Nope", 0, 0, 1) = 0 Then
         'discard it and go
         doDiscardGear playerID, CryBaby
         PutMsg player.PlayName & "'s Nav log: Cry Baby Deployed, Cruiser decoyed!", player.ID, Logic!Gamecntr, True, 0, 0, 1
         moveAutoAI 5
         Exit Function
      End If
   End If
      
   pay = 0
   SQL = "SELECT * FROM Players WHERE PlayerID = " & playerID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
    'Pay Fines: $1000 per Warrant, clear Warrants, All Contraband and Fugitives seized.
      If rst!Warrants > 0 Then
         pay = 1000 * rst!Warrants
         If pay > rst!pay Then 'take alls that's left
            pay = rst!pay
         End If
      End If
   End If
   
   rst.Close
   
   DB.Execute "UPDATE Players SET Warrants = 0, Pay = Pay - " & pay & ", Contraband=0, Fugitive = 0 WHERE PlayerID = " & playerID
   
   'Roll for each Wanted Crew: 1-Remove Crew, 2+ Crew safe
   SQL = "SELECT PlayerSupplies.CardID, Crew.* "    ', Crew.* "
   SQL = SQL & "FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND Crew.Wanted>0"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If hasGearCard(playerID, 20, rst!CrewID) > 0 Then
         'skip this Crew that has Alliance Ident Card
         PutMsg player.PlayName & "'s Crew member " & rst!CrewName & " flashes their fake Alliance Ident.Card", playerID, Logic!Gamecntr, True, rst!CrewID
      ElseIf hasShipUpgrade(playerID, 11) And crewcnt < 2 Then
         If crewcnt = 0 Then PutMsg player.PlayName & "'s Nav log: Concealed Smuggling Compartments hides up to 2 Wanted Crew", playerID, Logic!Gamecntr, True, getLeader()
         crewcnt = crewcnt + 1
      Else
         doSeizeCrew playerID, rst!CardID, rst!wanted
      End If
      rst.MoveNext
   Wend
   rst.Close
   doMoveAlliance = True
   Set rst = Nothing
End Function

'move the Cruiser adjacent the sector given
Public Function doMoveAllianceAdjacent(ByVal SectorID, Optional ByVal check As Boolean = False) As Boolean
Dim adjacent, a() As String, x, y
   
   adjacent = varDLookup("AdjacentRows", "Board", "SectorID=" & SectorID)
   a = Split(adjacent, ",")
   
   y = 0
   For x = LBound(a) To UBound(a)
      If getClearSector(Val(a(x))) = "A" Then  'no ship in this spot
         y = 1 'we have at least one possible solution
         If check Then
            doMoveAllianceAdjacent = True
            Exit Function
         End If
         Exit For
      End If
   Next x
   If y = 1 Then
      Do
         x = RollDice(UBound(a) - LBound(a) + 1) - 1
         If x > UBound(a) Then x = UBound(a)
         If getClearSector(Val(a(x))) = "A" Then
            MoveShip 5, Val(a(x))
            doMoveAllianceAdjacent = True
            Exit Do
         End If
      Loop
   End If

   
End Function


'move the Cruiser adjacent the sector given
Public Function doMoveCorvetteAdjacent(ByVal SectorID, Optional ByVal check As Boolean = False) As Boolean
Dim adjacent, a() As String, x, y
   
   adjacent = varDLookup("AdjacentRows", "Board", "SectorID=" & SectorID)
   a = Split(adjacent, ",")
   
   y = 0
   For x = LBound(a) To UBound(a)
      If getClearSector(Val(a(x))) <> "" And (Val(a(x)) < 120 Or Val(a(x)) > 122) Then 'no ship in this spot, and not in reaver zones
         y = 1 'we have at least one possible solution
         If check Then
            doMoveCorvetteAdjacent = True
            Exit Function
         End If
         Exit For
      End If
   Next x
   If y = 1 Then
      Do
         x = RollDice(UBound(a) - LBound(a) + 1) - 1
         If x > UBound(a) Then x = UBound(a)
         If getClearSector(Val(a(x))) <> "" And (Val(a(x)) < 120 Or Val(a(x)) > 122) Then
            MoveShip 6, Val(a(x))
            doMoveCorvetteAdjacent = True
            Exit Do
         End If
      Loop
   End If

   
End Function

'move the Cruiser adjacent the sector given
Public Function doMoveCorvettePlanetary() As Boolean
Dim x
   
      Do
         x = RollDice(152)
         If Nz(varDLookup("PlanetID", "Planet", "SectorID=" & x), 0) > 0 And getClearSector(x) <> "" And x <> 63 And x <> 64 Then
            MoveShip 6, x
            PutMsg "Corvette turns up at " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & x), " the unknown..")
            doMoveCorvettePlanetary = True
            Exit Do
         End If
      Loop

   
End Function

'move the Cruiser adjacent the sector given
Public Function doMoveCutterPlanetary(ByVal ship) As Boolean
Dim x
   
      Do
         x = RollDice(152)
        
         If Nz(varDLookup("PlanetID", "Planet", "SectorID=" & x), 0) > 0 And getClearSector(x) <> "" And getZone(x) <> "A" And x > 2 Then
            MoveShip ship, x
            PutMsg "Cutter turns up at " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & x), " the unknown..")
            doMoveCutterPlanetary = True
            Exit Do
         End If
      Loop

   
End Function

'move the tokens adjacent the sector given
Public Function doAddTokensAdjacent(ByVal SectorID) As Boolean
Dim adjacent, a() As String, x
   
   adjacent = varDLookup("AdjacentRows", "Board", "SectorID=" & SectorID)
   a = Split(adjacent, ",")

   For x = LBound(a) To UBound(a)
      changeAToken Val(a(x)), 1
   Next x
      
End Function

'add an Alliance Alert token at every Outlaw Ship
Public Sub doAddTokensOutlaws()
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM Players Where Name IS NOT NULL"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If isOutlaw(rst!playerID) Then
         changeAToken rst!SectorID, 1
      End If
      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing
End Sub

Public Function isOutlaw(ByVal playerID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   isOutlaw = False
   SQL = "SELECT * FROM Players WHERE PlayerID= " & playerID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      If rst!Warrants > 0 Then isOutlaw = True
      If rst!Fugitive > 0 Then isOutlaw = True
      If rst!Contraband > 0 Then isOutlaw = True
   End If
   rst.Close
   
   If isOutlaw Then Exit Function
   
   SQL = "SELECT SupplyDeck.CrewID "
   SQL = SQL & "FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND Crew.Wanted>0"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If Not hasGear(playerID, 20, rst!CrewID) Then ' has warrant, but may have Harken's Card
         isOutlaw = True
      End If
      rst.MoveNext
   Wend
   rst.Close
   
   
   Set rst = Nothing
End Function

Public Function outlawExists(ByVal playerID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM Players WHERE Name IS NOT NULL AND PlayerID <> " & playerID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   Do While Not rst.EOF
      If isOutlaw(rst!playerID) And getZone(Nz(rst!SectorID, 0)) = "A" Then
         outlawExists = True
         Exit Do
      End If
      rst.MoveNext
   Loop
   rst.Close
   Set rst = Nothing
End Function

Public Function getPlayerCount(Optional ByVal loadnames As Boolean = False) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT playerID, name FROM Players WHERE Name IS NOT NULL"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      getPlayerCount = getPlayerCount + 1
      If loadnames Then PlayCode(rst!playerID).PlayName = rst!Name
      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing
End Function

'discard Job and clear solid, with Contact
Public Function doJobWarrant(ByVal playerID, ByVal ContactID, ByVal CardID)
Dim frmKillCrw As frmKillCrew, SQL

   If warrantDodge(playerID) Then Exit Function
   'and if Niska - Kill 1 Crew
   'add a Warrant and clear any Solid with the Contact & Harken (5)
   If ContactID = 0 Then 'Goal
      SQL = "UPDATE Players SET Warrants = Warrants + 1"
      If Not discardRoberta(playerID) Then
         SQL = SQL & ", Solid5 = 0"
      End If
      SQL = SQL & " WHERE PlayerID = " & playerID
      DB.Execute SQL
      PutMsg player.PlayName & "'s Goal log: a Warrant has been issued, you're an Outlaw Ship!", playerID, Logic!Gamecntr, True, getLeader()
   Else
      If ContactID = 3 Then
         Set frmKillCrw = New frmKillCrew
         frmKillCrw.nbrSelect = 1
         frmKillCrw.Caption = "Niska demands a Crew is slain, pick one!"
         frmKillCrw.Show 1
      End If
      
      
      SQL = "UPDATE Players SET Warrants = Warrants + 1, Solid5 = 0"
      If Not discardRoberta(playerID) Then
         If ContactID <> 5 Then
            SQL = SQL & ", Solid" & ContactID & "=0"
         End If
      End If
      SQL = SQL & " WHERE PlayerID = " & playerID
      DB.Execute SQL

      
      'discard the Job
      DB.Execute "DELETE FROM PlayerJobs WHERE PlayerID = " & playerID & " AND CardID = " & CardID
      DB.Execute "UPDATE ContactDeck SET Seq = 5 WHERE CardID =" & CardID
      PutMsg player.PlayName & "'s Work log: a Warrant has been issued, the Job is forfeited, you've lost any Rep with this Contact. You're an Outlaw Ship!", playerID, Logic!Gamecntr, True, getLeader()
   End If

End Function

Public Function beaDirtySlaver(ByVal playerID) As Boolean
   If hasCrew(playerID, 86) Then
      If hasCrewAttribute(playerID, "Moral") Then
         If MessBox("Wright, the Dirty Slaver, can get an extra $100 per Fugitive. This will upset your Moral Crew." & vbNewLine & "Do you want to take the money anyway?", "Immoral Money", "Yes Way", "No Way", 86) = 0 Then
            doDisgruntled playerID, 1
            beaDirtySlaver = True
            PutMsg player.PlayName & " used Dirt Slaver to get an extra $100 per Fugitive.", playerID, Logic!Gamecntr
         End If
      Else 'of course you'll take the money
         beaDirtySlaver = True
         PutMsg player.PlayName & " used Dirt Slaver to get an extra $100 per Fugitive.", playerID, Logic!Gamecntr
      End If
   End If

End Function

Public Function discardRoberta(ByVal playerID) As Boolean
   If hasCrew(playerID, 79) Then
      If MessBox("Roberta can go and smooth things over so you don't lose Solid." & vbNewLine & "Do you want to discard her to do that?", "Solid on the line", "Discard", "Keep", 79) = 0 Then
         discardRoberta = True
         doDiscardCrew 171
         PutMsg player.PlayName & " used Roberta to avoid losing Solid.", playerID, Logic!Gamecntr
      End If
   End If

End Function

Public Function warrantDodge(ByVal playerID) As Boolean

   If hasCrew(playerID, 74) Then
      If MessBox("Fan Dancer can make this Warrant go away." & vbNewLine & "Do you want to discard her and avoid the Warrant?", "Warrant Avoidance", "Discard", "Keep", 74) = 0 Then
         doDiscardCrew 166
         PutMsg player.PlayName & " used their Fan Dancer to avoid a Warrant.", playerID, Logic!Gamecntr
         warrantDodge = True
      End If
   End If
   
End Function

Public Function hasJobReqs(ByVal playerID, ByVal CardID, ByVal JobID) As Boolean
Dim rst As New ADODB.Recordset, a() As String, x
Dim SQL

   hasJobReqs = True  'until proven otherwise
   SQL = "SELECT ContactDeck.*, Job.*, p.Cargo AS PCargo, p.Contraband AS PContraband, p.Fugitive as PFugitive, p.Passenger as PPassenger, p.Fuel as PFuel, p.Parts as PParts "
   SQL = SQL & " FROM ContactDeck, Job, Players AS p WHERE CardID=" & CardID & " AND JobID = " & JobID & " AND PlayerID = " & playerID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      If rst!fight > getSkill(playerID, "fight") Then
         hasJobReqs = False
      End If
      If rst!tech > getSkill(playerID, "tech") Then
         hasJobReqs = False
      End If
      If rst!Negotiate > getSkill(playerID, "negotiate") Then
         hasJobReqs = False
      End If
      
      'Solid only check
      If rst!Solid > 0 And rst!KeywordOrSolid = 0 Then 'must be solid
         If Not isSolid(playerID, rst!Solid) Then 'see if we have Keywords
           hasJobReqs = False
         End If
      End If
      
      ' Keyword Checks
      If rst!KeywordOrSkill > 0 And hasCrewAttribute(playerID, cstrProfession(rst!RequireProfession)) Then 'we have the profession needed, no need to check keyword
      
      ElseIf Nz(rst!KeyWords, "") <> "" And rst!Solid > 0 And rst!KeywordOrSolid > 0 Then 'needs the keywd if not solid
         If Not isSolid(playerID, rst!Solid) Then 'see if we have Keywords
            a = Split(rst!KeyWords, " ")
            For x = LBound(a) To UBound(a)
               If Not hasKeyword(playerID, a(x)) Then hasJobReqs = False
            Next x
         End If
         
      ElseIf Nz(rst!KeyWords, "") <> "" And rst!WinOptKeyword = 0 And rst!KeywordBonus = 0 Then ' as some Keywords are not a requirement
         a = Split(rst!KeyWords, " ")
         For x = LBound(a) To UBound(a)
            If Not hasKeyword(playerID, a(x)) Then hasJobReqs = False
         Next x
      End If
      
      If rst!RequireProfession > 0 And hasJobReqs And rst!KeywordOrSkill = 0 Then
         'If rst!RequireProfession = 12 Then 'pilot & mech
         '   If Not (hasCrewAttribute(playerID, cstrProfession(2)) And hasCrewAttribute(playerID, cstrProfession(1))) Then hasJobReqs = False
         'Else
            If Not hasCrewAttribute(playerID, cstrProfession(rst!RequireProfession)) Then hasJobReqs = False
         'End If
      End If
         
      'check Job Requirements
      If (rst!Contraband = -14 And rst!PContraband = 0) Or (rst!Passenger = -14 And rst!PPassenger = 0) Or (rst!Fugitive = -14 And rst!PFugitive = 0) Then
         hasJobReqs = False
      ElseIf rst!cargo + rst!PCargo < 0 Or (rst!Contraband + rst!PContraband < 0 And rst!Contraband > -14) Or (rst!Fugitive + rst!PFugitive < 0 And rst!Fugitive > -14) Or (rst!Passenger + rst!PPassenger < 0 And rst!Passenger > -14) Or rst!fuel + rst!PFuel < 0 Or rst!parts + rst!PParts < 0 Then
         hasJobReqs = False
      End If
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function hasNavReqs(ByVal playerID, ByVal CardID, ByVal opt) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL

   hasNavReqs = True
   SQL = "SELECT NavOption.* "
   SQL = SQL & "FROM NavOption INNER JOIN NavDeck ON NavOption.OptionID = NavDeck.Option" & opt & "ID "
   SQL = SQL & "Where NavDeck.CardID = " & CardID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
         If rst!WinFuel < 0 And hasNavReqs Then
            hasNavReqs = (Abs(rst!WinFuel) <= varDLookup("Fuel", "Players", "PlayerID = " & playerID))
         End If
         
         If rst!WinParts <> -99 And rst!WinParts < 0 And rst!Breakdown = 0 And hasNavReqs Then  'breakdown allows action to stop if no parts avail, -99 is optional parts sell
            hasNavReqs = (Abs(rst!WinParts) <= varDLookup("Parts", "Players", "PlayerID = " & playerID))
         End If
         
         If rst!WinCargo < 0 And hasNavReqs Then
            hasNavReqs = (Abs(rst!WinCargo) <= varDLookup("Cargo", "Players", "PlayerID = " & playerID))
         End If

         If rst!WinProfession > 0 And Not hasCrewAttribute(playerID, cstrProfession(rst!WinProfession)) And hasNavReqs Then  'check we meet fail conditions
            If rst!FailFuel < 0 And Abs(rst!FailFuel) > varDLookup("Fuel", "Players", "PlayerID = " & playerID) Then
               hasNavReqs = False
            End If
            If rst!FailParts < 0 And Abs(rst!FailParts) > varDLookup("Parts", "Players", "PlayerID = " & playerID) Then
               hasNavReqs = False
            End If
         End If
         
         If rst!Profession > 0 And hasNavReqs Then
            hasNavReqs = hasCrewAttribute(playerID, cstrProfession(rst!Profession))
         End If
         
         If rst!Planet = -1 And hasNavReqs Then 'requires a planetary sector
            hasNavReqs = (getPlanetID(playerID) > 0)
         End If
         
         If rst!Planet > 0 And hasNavReqs Then 'must be at a specific Planet
            hasNavReqs = (getPlanetID(playerID) = rst!Planet)
         End If
         
         If rst!WinSolid > 0 And hasNavReqs Then 'requirement to be solid with (rst!WinSolid) to enable this option
            hasNavReqs = isSolid(playerID, rst!WinSolid)
         End If

         If rst!MoveAlliance = 3 And hasNavReqs Then  'only if outlaw ship is in zone A
            hasNavReqs = outlawExists(player.ID)
         End If
         If rst!MoralCrew = 1 And hasNavReqs Then
            hasNavReqs = hasCrewAttribute(playerID, "Moral")
         End If
         
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function isSolid(ByVal playerID, ByVal ContactID) As Boolean

   If ContactID = 0 Or playerID = 0 Then Exit Function
   isSolid = (varDLookup("Solid" & ContactID, "Players", "PlayerID=" & playerID) = 1)
   
   If Not isSolid And ContactID = 5 Then 'alliance card gives solid with Harken
      isSolid = hasGear(playerID, 20)
   End If
End Function

Public Sub clearOffJob(ByVal playerID)
   DB.Execute "UPDATE PlayerSupplies SET OffJob = 0 WHERE PlayerID = " & playerID
End Sub


Public Function getCrewGearCount(ByVal CrewID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   
   SQL = "SELECT Count([PlayerSupplies].[CardID]) AS cnt "
   SQL = SQL & "FROM PlayerSupplies INNER JOIN (Gear INNER JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.CrewID=" & CrewID & " AND Gear.noGearLimit=0"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getCrewGearCount = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function getExtraBurn(ByVal playerID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Max(ContactDeck.ExtraFuel) as cnt "
   SQL = SQL & "FROM PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID "
   SQL = SQL & "WHERE PlayerJobs.PlayerID=" & playerID & " AND (PlayerJobs.JobStatus=1 OR PlayerJobs.JobStatus=2) AND ContactDeck.ExtraFuel>0"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       getExtraBurn = Nz(rst!cnt, 0)
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Sub assignDeal(ByVal playerID, ByVal CardID)

      DB.Execute "UPDATE ContactDeck SET Seq =" & playerID & " WHERE CardID = " & CardID
      'add the card to the players deck
      DB.Execute "INSERT INTO PlayerJobs (PlayerID, CardID) VALUES (" & playerID & ", " & CardID & ")"

End Sub

Public Function jobSuccess(ByVal playerID, ByVal CardID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT JobStatus FROM PlayerJobs WHERE PlayerID =" & playerID & " AND CardID=" & CardID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      jobSuccess = (Nz(rst!JobStatus, 0) = JOB_SUCCESS)
   End If
   rst.Close
   Set rst = Nothing
End Function

'move a NPC Ship one sector, preferencing a player ship sector if adjacent.    Return: sectorID of any ship encountered
Public Function moveAutoAI(ByVal ship As Integer, Optional ByVal sound As Integer = 0, Optional ByVal syncsound As Boolean = False, Optional ByVal leaveToken As Boolean = True) As Integer
Dim rst As New ADODB.Recordset, cnt
Dim SQL, SectorID As Integer, Zone As String, a() As String, b(1 To 20) As Integer, x, y, z, adjacent, NPCFlag As Boolean

   Zone = IIf(ship = 5, "A", "B") 'lock cruiser to A, treat B & R as same
   SectorID = varDLookup("SectorID", "Players", "PlayerID = " & ship)
         
   SQL = "SELECT * FROM Board WHERE SectorID = " & SectorID
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
   
      adjacent = rst!AdjacentRows
      a = Split(adjacent, ",")
      
   End If
   rst.Close
   cnt = 0
   SQL = "SELECT * FROM Players WHERE Name IS NOT NULL OR PlayerID between 6 AND " & CStr(NumOfReavers + 6)
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      'look at each adjacent Sector and record if there is a Player as a count and B()
      For x = LBound(a) To UBound(a)
         If Val(a(x)) = rst!SectorID And IIf(getZone(rst!SectorID) = "A", "A", "B") = Zone And ship > 4 Then   'found one, treat B & R as same
            If rst!playerID < 5 Then ' tag you're it
               'then check if Haven exists and if this is Cruiser
               If Not (getHaven(rst!SectorID) > 0 And ship = 5) Then
                  cnt = cnt + 1
                  b(cnt) = rst!SectorID
               End If
            ElseIf rst!playerID > 5 Then 'collect corvette or reaver locations for later
               b(rst!playerID) = rst!SectorID
            End If
         End If
      Next x
      rst.MoveNext
   Wend
   rst.Close
   adjacent = 0
   'lock onto players sector
   If cnt = 1 Then
      adjacent = b(1)
   ElseIf cnt > 1 Then
      cnt = RollDice(cnt)
      adjacent = b(cnt)
   End If
   moveAutoAI = adjacent  'return sectorID
   NPCFlag = False
   For z = 6 To 6 + NumOfReavers  'check if a NPC ship already there
      If adjacent = b(z) Then
         NPCFlag = True
         If ship > 6 Then adjacent = 0
         Exit For
      End If
   Next z
   
   If adjacent = 0 Then  'we found no players adjacent
      If ship < 7 Or Not NPCFlag Then
         adjacent = getPursuitSector(SectorID, ship)
      End If
      If adjacent = 0 Then 'no AI solution, just move randomally
         y = 0
         For x = LBound(a) To UBound(a)
            NPCFlag = True
            For z = 6 To 6 + NumOfReavers  'check if a NPC ship already there
               If Val(a(x)) = b(z) Then
                  NPCFlag = False
                  Exit For
               End If
            Next z
            If NPCFlag And (IIf(getZone(Val(a(x))) = "A", "A", "B") = Zone Or ship < 5) And Not (getHaven(a(x)) > 0 And ship = 5) Then
               y = 1 'we have at least one possible solution
               Exit For
            End If
         Next x
         If y = 1 Then
            Do
               cnt = RollDice(UBound(a) - LBound(a) + 1) - 1
               If cnt > UBound(a) Then cnt = UBound(a)
               'make sure you don't move a reaver to another reavers sector
               NPCFlag = True
               For z = 6 To 6 + NumOfReavers  'check if a NPC ship already there
                  If Val(a(cnt)) = b(z) Then
                     NPCFlag = False
                     Exit For
                  End If
               Next z
               If NPCFlag And (IIf(getZone(Val(a(cnt))) = "A", "A", "B") = Zone Or ship < 5) And Not (getHaven(a(cnt)) > 0 And ship = 5) Then
                  adjacent = Val(a(cnt))
                  Exit Do
               End If
            Loop
         ElseIf ship > 6 Then  'no where to go, locked in
            doMoveCutterPlanetary ship
         
         End If
      End If
   End If
   
   If adjacent > 0 Then 'valid move ok, otherwise no move
      MoveShip ship, adjacent, sound, syncsound, leaveToken
   End If
   
   Set rst = Nothing
End Function

Public Function getAdjacentRows(ByVal SectorID) As String
   If SectorID < 1 Then Exit Function
   getAdjacentRows = varDLookup("AdjacentRows", "Board", "SectorID=" & SectorID)

End Function

'move the Corvette 2 sector2, preferencing a player ship sector if adjacent.    Return: sectorID of any ship encountered
Public Function moveAutoCorvette2(Optional ByVal sound As Integer = 0, Optional ByVal syncsound As Boolean = False, Optional ByVal avoid As Integer = 0) As Integer
Dim rst As New ADODB.Recordset, cnt, ship, found As Boolean, Cruiser
Dim SQL, SectorID As Integer, a() As String, b(1 To 20, 1 To 2) As Integer, c() As String, x, y, z, adjacent As String, vector As Integer, vector2 As Integer

   ship = 6
   SectorID = getCorvetteSector()
   y = Nz(varDLookup("PlayerID", "Players", "Name IS NOT NULL AND SectorID=" & avoid), 0)
   Cruiser = getCruiserSector()
         
   adjacent = getAdjacentRows(SectorID)
   a = Split(adjacent, ",")
   cnt = 0
   SQL = "SELECT * FROM Players WHERE Name IS NOT NULL OR PlayerID between 4 AND " & CStr(NumOfReavers + 6)
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      'look at each adjacent Sector and record if there is a Player as a count and B()
      
      For x = LBound(a) To UBound(a) 'first step array
         If a(x) = "-1" Then 'ignore

         ElseIf (avoid > 0 And avoid = Val(a(x)) And y = 0) Then 'take out this start sector from scan, unless a player is there
            a(x) = "-1"
         ElseIf Val(a(x)) = rst!SectorID And (rst!SectorID < 120 Or rst!SectorID > 122) Then   'found one, treat B & R as same
            If rst!playerID < 5 Then ' tag you're it
               cnt = cnt + 1
               b(cnt, 1) = rst!SectorID
               Exit For
            End If
         Else 'nothing found, check second step sectors
            adjacent = getAdjacentRows(Val(a(x)))
            c = Split(adjacent, ",")
            For z = LBound(c) To UBound(c)
            
               If avoid > 0 And avoid = Val(c(z)) And y = 0 Then 'take out this start sector from scan, unless a player is there
                  c(z) = "-1"
               ElseIf Val(c(z)) = rst!SectorID And (rst!SectorID < 120 Or rst!SectorID > 122) Then   'found one, treat B & R as same
                  If rst!playerID < 5 Then ' tag you're it
                     cnt = cnt + 1
                     b(cnt, 1) = Val(a(x)) 'breadcrumb
                     b(cnt, 2) = rst!SectorID 'destination
                     found = True
                     Exit For
                  End If
               End If
            
            Next z
            If found Then 'found a solution
               found = False
               Exit For
            End If
         End If
      Next x
      rst.MoveNext
   Wend
   rst.Close
   vector = 0
   'lock onto players sector
   If cnt = 1 Then
      vector = b(1, 1)
      vector2 = b(1, 2)
   ElseIf cnt > 1 Then
      cnt = RollDice(cnt)
      vector = b(cnt, 1)
      vector2 = b(cnt, 2)
   End If
   If vector2 > 0 Then
      moveAutoCorvette2 = vector2  'return sectorID
   Else
      moveAutoCorvette2 = vector  'return sectorID
   End If
   If vector = 0 Or (vector > 0 And vector = b(5, 1)) Then 'we found no players adjacent or the Cruiser is already there
      vector = getPursuitSector(SectorID, ship)
   End If
   
   If vector > 0 Then 'valid move ok, otherwise no move
      MoveShip ship, vector, sound, syncsound
   End If
   If vector2 > 0 Then 'valid move ok, otherwise no move
      MoveShip ship, vector2, sound, syncsound
   ElseIf vector > 0 And moveAutoCorvette2 = 0 Then 'no 2nd solution, so keep pursuiting
      vector2 = getPursuitSector(vector, ship)
      MoveShip ship, vector2, sound, syncsound
   End If
   
   Set rst = Nothing
End Function

' return the pursuit sector to move to.- Sector of ship moving,   ID of ship where 6 = Corvette
Private Function getPursuitSector(ByVal SectorID, Optional ByVal ship As Integer = 6) As Integer
Dim rst As New ADODB.Recordset
Dim SQL, b(1 To 20, 1 To 3) As Long, x As Integer, y As Long, z As Long, adjacent

   SQL = "SELECT Players.PlayerID, Players.SectorID, Board.STop, Board.SLeft, Board.SHeight, Board.SWidth, Board.Zones "
   SQL = SQL & "FROM Board INNER JOIN Players ON Board.SectorID = Players.SectorID "
   SQL = SQL & "WHERE Players.Name Is Not Null OR PlayerID = " & ship
   SQL = SQL & " ORDER BY PlayerID DESC"
   x = 0
   y = 0
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      'Find the closest Player
      b(rst!playerID, 1) = Int(rst!SHeight / 2 + rst!STop)  'X
      b(rst!playerID, 2) = Int(rst!SWidth / 2 + rst!SLeft)  'Y
      If rst!playerID < 5 And (ship = 6 Or (ship > 6 And rst!Zones <> "A") Or (ship = 5 And rst!Zones = "A")) Then 'start comparing with players
         b(rst!playerID, 3) = Int(Sqr((b(ship, 1) - b(rst!playerID, 1)) ^ 2 + (b(ship, 2) - b(rst!playerID, 2)) ^ 2))
         If y = 0 Or y > b(rst!playerID, 3) Then
            y = b(rst!playerID, 3)
            x = rst!playerID   'the closest Player
         End If
      End If
      rst.MoveNext
   Wend
   rst.Close
   If x = 0 Then 'noone found
      Exit Function
   End If
   y = -1
   adjacent = getAdjacentRows(SectorID)
   SQL = "SELECT Board.*, Players.PlayerID FROM Board LEFT JOIN Players ON Board.SectorID = Players.SectorID WHERE Board.SectorID IN (" & adjacent & ")"
   If ship = 5 And useHavens(Logic!StoryID) Then
      SQL = SQL & " AND Haven = 0"
   End If
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If (ship = 5 And rst!Zones = "A") Or (ship = 6 And (rst!SectorID < 120 Or rst!SectorID > 122 Or getCorvetteSector() = 123)) Or (ship > 6 And rst!Zones <> "A" And Nz(rst!playerID, 0) < 5) Then 'path rules for the ships
         'find the adjacent sector closest to the closest player
         z = Int(Sqr((b(x, 1) - Int(rst!SHeight / 2 + rst!STop)) ^ 2 + (b(x, 2) - Int(rst!SWidth / 2 + rst!SLeft)) ^ 2))
         If y = -1 Or y > z Then
            y = z
            getPursuitSector = rst!SectorID
         End If
      End If
      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing
End Function

Public Function getSectorCount(ByVal SectorID, ByVal Target) As Integer
Dim rst As New ADODB.Recordset
Dim SQL, b(1 To 3) As Long, y As Long, z As Long, adjacent

   If Target = 1 Then Target = getCruiserSector()
   If Target = 2 Then Target = getCorvetteSector()
   If SectorID = Target Then Exit Function

   SQL = "SELECT Board.STop, Board.SLeft, Board.SHeight, Board.SWidth, Board.Zones "
   SQL = SQL & "FROM Board WHERE Board.SectorID = " & Target
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      'Find the closest Player
      b(1) = Int(rst!SHeight / 2 + rst!STop) 'X
      b(2) = Int(rst!SWidth / 2 + rst!SLeft)  'Y
    
   End If
   rst.Close
   
   Do
      y = -1
      adjacent = getAdjacentRows(SectorID)
      SQL = "SELECT Board.* FROM Board WHERE Board.SectorID IN (" & adjacent & ")"
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
            'find the adjacent sector closest to the closest player
            z = Int(Sqr((b(1) - Int(rst!SHeight / 2 + rst!STop)) ^ 2 + (b(2) - Int(rst!SWidth / 2 + rst!SLeft)) ^ 2))
            If y = -1 Or y > z Then
               y = z
               
               SectorID = rst!SectorID
            End If
   
         rst.MoveNext
      Wend
      rst.Close
      getSectorCount = getSectorCount + 1
   Loop While SectorID <> Target
   
   Set rst = Nothing
End Function

Public Function canRemoveUpgrade(ByVal playerID, ByVal CardID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL

   canRemoveUpgrade = True
   
   SQL = "SELECT ShipUpgrade.* "
   SQL = SQL & "FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND SupplyDeck.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      If rst!DriveCore = 1 Then
         canRemoveUpgrade = False
      ElseIf rst!ExtraCrewSpace > 0 Then
         If (CrewCapacity(playerID) - getCrewCount(playerID)) < rst!ExtraCrewSpace Then
            canRemoveUpgrade = False
            MessBox "Not enough remaining capacity to carry the existing crew", "Cannot Discard", "Ooops", "", 0, 0, 2
         End If
      ElseIf rst!ExtraStashSpace > 0 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < rst!ExtraStashSpace Then
            canRemoveUpgrade = False
            MessBox "Not enough remaining capacity to carry the existing cargo", "Cannot Discard", "Ooops", "", 0, 0, 6
         End If
      ElseIf rst!ExtraCargoSpace > 0 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < rst!ExtraCargoSpace Then
            canRemoveUpgrade = False
            MessBox "Not enough remaining capacity to carry the existing cargo", "Cannot Discard", "Ooops", "", 0, 0, 6
         End If
      End If
      
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function doRemoveUpgrade(ByVal playerID, ByVal CardID) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL, frmSeize As frmSeized, frmSalvage As frmSalvaging, cnt

   
   SQL = "SELECT ShipUpgrade.* "
   SQL = SQL & "FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & playerID & " AND SupplyDeck.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      If rst!DriveCore = 1 Then
         doRemoveUpgrade = False
      ElseIf rst!ExtraCrewSpace > 0 Then
         If (CrewCapacity(playerID) - getCrewCount(playerID)) < rst!ExtraCrewSpace Then
            cnt = rst!ExtraCrewSpace - (CrewCapacity(playerID) - getCrewCount(playerID))
            
            Set frmSeize = New frmSeized
            frmSeize.Caption = "Select " & cnt & " Crew Member/s to leave the Ship as there is not enough Crewspace"
            frmSeize.nbrSelect = cnt
            If frmSeize.RefreshDiscardList() > 0 Then 'crew exist
               frmSeize.Show 1, Main
            End If
            
            doRemoveUpgrade = True
         End If
      ElseIf rst!ExtraStashSpace > 0 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < rst!ExtraStashSpace Then
            cnt = rst!ExtraStashSpace - (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            Set frmSalvage = New frmSalvaging
            frmSalvage.mode = 3
            frmSalvage.salvageCount = cnt
            frmSalvage.Show 1, Main
            doRemoveUpgrade = True
            
         End If
      ElseIf rst!ExtraCargoSpace > 0 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < rst!ExtraCargoSpace Then
            cnt = rst!ExtraCargoSpace - (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            Set frmSalvage = New frmSalvaging
            frmSalvage.mode = 3
            frmSalvage.salvageCount = cnt
            frmSalvage.Show 1, Main
            doRemoveUpgrade = True
            
         End If
      End If
      
   End If
   rst.Close
   Set rst = Nothing
End Function

'returns which ship turns up if any
Public Function resolveToken(ByVal SectorID, Optional ByVal adjacent As Boolean = False) As Integer
Dim rst As New ADODB.Recordset, x, alliance As Boolean
Dim SQL
  
   SQL = "SELECT * FROM Board WHERE SectorID= " & SectorID
   rst.Open SQL, DB, adOpenStatic, adLockReadOnly
   If Not rst.EOF Then
      'do alliance tokens first
      If rst!AToken > 0 Then ' we gotta roll
         x = RollDice(6, True)
         If x <= rst!AToken Then ' we not good
            alliance = True
            If getZone(SectorID) = "A" Then 'roll for which one
               resolveToken = 4 + RollDice(2)
               If resolveToken > 6 Then resolveToken = 6
            Else
               resolveToken = 6 'corvette
            End If
            'clear any reaver tokens  'these are permanent min 1
            changeToken SectorID, -1, False
            MoveShip resolveToken, SectorID
            PutMsg player.PlayName & " " & IIf(adjacent, "scanned", "entered") & " a Sector on Alliance Alert Level " & CStr(rst!AToken) & ", and got a nasty surprise by rolling a " & x, player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, x
         Else
            PutMsg player.PlayName & " " & IIf(adjacent, "scanned", "entered") & " a Sector on Alliance Alert Level " & CStr(rst!AToken) & ", but found it all clear by rolling a " & x, player.ID, Logic!Gamecntr
         End If
         'rst!AToken = 0  'clear Alliance tokens
         'rst.Update
         DB.Execute "UPDATE Board SET AToken = 0 WHERE SectorID = " & SectorID
      End If
      rst.Requery
      'is token and check a cutter is not in this sector already
      If rst!Token > 0 And getCutterSector(SectorID) = 0 And Not alliance Then ' we gotta roll
         x = RollDice(6, True)
         If x <= rst!Token Then ' we not good, reaver incoming
            resolveToken = 6 + RollDice(NumOfReavers)
            If resolveToken > 6 + NumOfReavers Then resolveToken = 6 + NumOfReavers
            PutMsg player.PlayName & " " & IIf(adjacent, "scanned", "entered") & " a Sector on Reaver Alert Level " & CStr(rst!Token) & ", and got a nasty surprise by rolling a " & x, player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, x
            MoveShip resolveToken, SectorID
            
         Else
            PutMsg player.PlayName & " " & IIf(adjacent, "scanned", "entered") & " a Sector on Reaver Alert Level " & CStr(rst!Token) & ", but found it all clear by rolling a " & x, player.ID, Logic!Gamecntr
         End If
         DB.Execute "UPDATE Board SET Token = " & IIf(SectorID > 119 And SectorID < 123, "1", "0") & " WHERE SectorID = " & SectorID

      ElseIf rst!Token > 0 And getCutterSector(SectorID) > 0 And Not alliance Then 'reaver already there, just clear the token
         DB.Execute "UPDATE Board SET Token = " & IIf(SectorID > 119 And SectorID < 123, "1", "0") & " WHERE SectorID = " & SectorID
         PutMsg player.PlayName & " " & IIf(adjacent, "scanned", "entered") & " a Sector on Reaver Alert Level " & CStr(rst!Token) & ", clearing the Alert and noting the known threat there", player.ID, Logic!Gamecntr
      End If

   
   End If
   rst.Close
   Set rst = Nothing
End Function

'can the player fullburn without hitting Reavers, looking for 1 free sector
Public Function hasAdjacentAlert(ByVal playerID) As Boolean
Dim currentSectorID, adjacent, a() As String, x
   
   currentSectorID = Nz(varDLookup("SectorID", "Players", "PlayerID=" & playerID), 0)
   
   adjacent = varDLookup("AdjacentRows", "Board", "SectorID=" & currentSectorID)
   a = Split(adjacent, ",")
   For x = LBound(a) To UBound(a)
      If (varDLookup("Token", "Board", "SectorID=" & a(x)) > 0 And getCutterSector(a(x)) = 0) Or (varDLookup("AToken", "Board", "SectorID=" & a(x)) > 0 And getCruiserCorvette(a(x)) = 0) Then
         hasAdjacentAlert = True
         Exit For
      End If
   Next x

End Function

'check Atherton's perk of no companions
Public Function companionsOK(ByVal playerID, ByVal typeID, ByVal CrewID) As Boolean

   companionsOK = True
   If typeID <> 1 Then Exit Function
   If Not hasCrew(playerID, 69) Then Exit Function
   
   If varDLookup("Companion", "Crew", "CrewID=" & CrewID) > 0 Then
      companionsOK = False
   End If
   
End Function

Public Function hasHigginsJayneGrudge(Optional ByVal isHiggins As Boolean = False) As Boolean
Dim SectorID, ContactID
   If isHiggins Then
      ContactID = 8
   Else  'see where we are
      SectorID = getPlayerSector(player.ID)
      ContactID = Nz(varDLookup("ContactID", "Contact", "SectorID=" & SectorID), 0)
   End If
   If ContactID <> 8 Then Exit Function 'not higgy
   If Not hasCrew(player.ID, 22) Then Exit Function  'not Jayne
   PutMsg player.PlayName & " cannot Deal with Higgins with Jayne in the Crew", player.ID, Logic!Gamecntr, True, 0, 0, 0, ContactID

   hasHigginsJayneGrudge = True
   
End Function

Public Function hasHigginsJayneWork(ByVal CardID) As Boolean
Dim ContactID
   If CardID = 0 Then Exit Function
   ContactID = Nz(varDLookup("ContactID", "ContactDeck", "CardID=" & CardID), 0)

   If ContactID <> 8 Then Exit Function 'not higgy
   If Not hasCrew(player.ID, 22) Then Exit Function  'not Jayne
   PutMsg player.PlayName & " cannot Work for Higgins with Jayne in the Crew", player.ID, Logic!Gamecntr, True, 0, 0, 0, ContactID
   hasHigginsJayneWork = True
   
End Function

'SoloGame
Public Function isSoloGame() As Boolean
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Count(playerID) as cnt FROM Players WHERE Name is Not Null AND AI = 0"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      isSoloGame = (rst!cnt < 2)
   End If
   rst.Close
   Set rst = Nothing
End Function


Public Function getExcludeCrew() As String
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT ExcludeCrew FROM Story WHERE StoryID =" & Logic!StoryID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getExcludeCrew = Nz(rst!ExcludeCrew)
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function getPlanetSector(ByVal JobID) As String
Dim rst As New ADODB.Recordset
Dim SQL
   If Nz(JobID, -1) = -1 Then Exit Function
   SQL = "SELECT Planet.PlanetName, Planet.System "
   SQL = SQL & "FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID  "
   SQL = SQL & "WHERE JobID =" & JobID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getPlanetSector = rst!PlanetName & " - " & rst!System
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Sub removeJob(ByVal playerID, ByVal CardID)
   DB.Execute "UPDATE ContactDeck Set Seq = 5 WHERE CardID = " & CardID
   DB.Execute "DELETE FROM PlayerJobs WHERE PlayerID = " & playerID & " AND CardID = " & CardID
End Sub

Public Sub setBackColour(cntrl As Control)
      If cntrl.Enabled Then
         cntrl.BackColor = &H80000005
      Else
         cntrl.BackColor = &HCBE1ED
      End If
End Sub

'Public Function getsumtin(ByVal playerID) As Integer
'Dim rst As New ADODB.Recordset
'Dim SQL
'   SQL = ""
'   SQL = SQL & ""
'   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
'   If Not rst.EOF Then
'
'      'rst.MoveNext
'   End If
'   rst.Close
'   Set rst = Nothing
'End Function
