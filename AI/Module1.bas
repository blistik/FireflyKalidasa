Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function PlaySound Lib "WINMM.DLL" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_ASYNC = &H1
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
Public actionSeq As actionSeqCntr, NumOfReavers As Integer, ContactList As String
Public Trail(0 To 8) As Integer 'record the trail of sectors travelled in a burn
Public MoseyMovesDone As Integer, FullburnMovesDone As Integer
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
Public datab                        'database for game

Public Function Logon() As Boolean
Dim ConStr As String
On Error Resume Next
  If Command$ = "" Then
     datab = App.Path & "\FireflyKalidasa.mdb"
     ConStr = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & datab & ";Persist Security Info=False"
  ElseIf Left(Command$, 16) = "Provider=MSDASQL" Then
     'use commandline>> Provider=MSDASQL;Driver={MariaDB ODBC 3.1 Driver};Server=localhost;Port=3306;
     ConStr = Command$ & "DATABASE=FireflyDB;UID=firefly;PWD=Firefly.2000"
     datab = ConStr
  Else
     datab = Command$
     ConStr = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & datab & ";Persist Security Info=False"
  End If
    
  Set DB = New ADODB.Connection
  DB.ConnectionString = ConStr
  
  DB.Open
  
  If Err Then
     Logon = False
     MsgBox "Unable to open game datasource at " & datab, vbCritical
  Else
     Logon = True
  End If
  
  
End Function

'STR_TO_DATE('11/10/2023 13:22','%e/%c/%Y %H:%i')
Public Function SQLDate(ByVal datetime As Date) As String
   If Left(datab, 16) = "Provider=MSDASQL" Then
      SQLDate = "STR_TO_DATE('" & Format(datetime, "DD/MM/YYYY HH:nn") & "','%e/%c/%Y %H:%i')"
   Else
      SQLDate = "#" & Format(datetime, "MM-DD-YY HH:nn") & "#"
   End If
End Function

Public Function SQLNow() As String
   If Left(datab, 16) = "Provider=MSDASQL" Then
      SQLNow = "Now()"
   Else
      SQLNow = "#" & Format(Now(), "MM-DD-YY HH:nn") & "#"
   End If
End Function

Public Sub PutMsg(msg, Optional playerID = 0, Optional turn = 0, Optional ByVal force As Boolean = False, Optional ByVal CrewID As Integer = 0, Optional ByVal GearID As Integer = 0, Optional ByVal ShipUpgradeID As Integer = 0, Optional ByVal ContactID As Integer = 0, Optional ByVal refreshShip As Integer = 0, Optional ByVal Dice As Integer = 0)
Dim SQL
On Error GoTo err_handler

   If Left(msg, 3) <> "Wai" Then 'waiting for game to start
      SQL = "INSERT INTO Events (Eventtime, Event, PlayerID, Turn, RefreshShip"
      SQL = SQL & ") Values (" & SQLNow & ", '" & SQLFilter(msg, 255) & "', " & playerID & ", " & turn & ", " & refreshShip
      SQL = SQL & ")"
      DB.Execute SQL
   End If

   If force Then
      
   End If

   Main.Stat.Panels(1).Text = msg
   
normal_exit:

   Exit Sub
   
err_handler:
   MsgBox "PutMsg Error: " & vbCrLf & Err.Description
   Resume normal_exit
   
End Sub
Public Function SQLFilter(ByVal Source As String, Optional ByVal size As Integer = 0) As String
Dim x, y

   If Source = "" Then Exit Function

   If size > 0 Then Source = Left(Source, size)

'  Looks for single quotes and doubles them ('') to create a literal
   x = 1
   Do
      y = InStr(x, Source, "'")
      If y Then
        Source = Left(Source, y) & Mid(Source, y)
        x = y + 2
      End If
   Loop While y
      
   Source = Replace(Source, "%", "-")
   Source = Replace(Source, "#", "-")
'   Source = Replace(Source, "*", "-")
'   Source = Replace(Source, "^", "-")
'   Source = Replace(Source, "$", "-")
   Source = Replace(Source, "!", "-")
      
   SQLFilter = Source
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
   Case "F"
      If Logic!Trader > 0 Then
         msg = "Waiting for a Showdown to complete between " & PlayCode(Logic!player).PlayName & " and " & PlayCode(Logic!Trader).PlayName
      Else
         msg = "Waiting for a Showdown to complete by " & PlayCode(Logic!player).PlayName
      End If
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

Public Function getFuel(ByVal playerID)
  getFuel = varDLookup("Fuel", "Players", "PlayerID=" & playerID)
End Function

Public Function getCrewCardID(ByVal CrewID) As Integer
   getCrewCardID = Nz(varDLookup("CardID", "SupplyDeck", "CrewID=" & CrewID), 0)
End Function

Public Function getSkillCrew(ByVal CrewID, ByVal skill As String) As Integer
   getSkillCrew = Nz(varDLookup(skill, "Crew", "CrewID=" & CrewID), 0)
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

Public Function setNextLeader(ByVal lastplayer, ByVal leader)
Dim rst As New ADODB.Recordset
    'set leader for outgoing player
    DB.Execute "UPDATE Players SET Leader = " & leader & " WHERE PlayerID = " & lastplayer
    rst.CursorLocation = adUseClient
    rst.Open "SELECT * FROM Players WHERE NAME IS NOT NULL ORDER BY PlayerID", DB, adOpenDynamic, adLockOptimistic
    rst.Find "PlayerID = " & lastplayer

    If Not rst.EOF Then
       'rst!leader = leader
       'rst.Update
       'mark the Card as selected
       DB.Execute "UPDATE SupplyDeck SET Seq =" & lastplayer & " WHERE CrewID =" & leader
       'drop this leaders Card into the Player's supplies
       'DB.Execute "INSERT INTO PlayerSupplies (PlayerID,CardID) VALUES (" & lastplayer & ", " & varDLookup("CardID", "SupplyDeck", "CrewID =" & leader) & ")"
       'test if last record
       rst.MoveNext
       If rst.EOF Then   'end of this round
         rst.Requery
         rst.MoveFirst
       End If
      
       If rst!leader = 0 Then 'not set yet
          setNextLeader = rst!playerID
          DB.Execute "UPDATE GameSeq SET Player = " & CStr(setNextLeader)
          'Logic.Update "Player", setNextLeader
       Else 'we done here as we're back to the first player
          setNextLeader = 0
          DB.Execute "UPDATE GameSeq SET Seq = 'S', GameCntr = 1, Player = " & CStr(player.ID)
          'Logic!Seq = "S"    'start game setup in main cycle
          'Logic!GameCntr = 1 'start counter, players will be on 0
          'Logic!player = player.ID  'with this player as first
          'Logic.Update
       End If
       Logic.Requery
   End If
End Function


Public Sub SetupPlayer(ByVal playerID, ByVal StoryID)
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM Story WHERE StoryID =" & StoryID
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenStatic, adLockReadOnly
   If Not rst.EOF Then
      DB.Execute "UPDATE Players SET Pay = " & rst!StartingCash & ", Warrants=0, Fuel = " & rst!StartingFuel & ", Parts = " & rst!StartingParts & " WHERE PlayerID =" & playerID
   End If
   rst.Close

Set rst = Nothing
End Sub

Public Function getCrew(ByVal SupplyID) As Boolean
Dim rst As New ADODB.Recordset, SQL, CrewID, crewcnt, imposter
   imposter = ""
   If hasCrew(player.ID, 23) Then
      imposter = "41,54"
   ElseIf hasCrew(player.ID, 41) Then
      imposter = "23,54"
   ElseIf hasCrew(player.ID, 54) Then
      imposter = "23,41"
   End If
   SQL = "SELECT SupplyDeck.CardID, Crew.* FROM Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID WHERE Seq=5 AND Wanted = 0 AND Moral = 0"
   If imposter <> "" Then SQL = SQL & " AND Crew.CrewID NOT IN (" & imposter & ")"
   SQL = SQL & " AND SupplyID = " & SupplyID
   If getLeader = 69 Then 'add Atherton check
      SQL = SQL & " AND Crew.Companion = 0"
   End If
   SQL = SQL & " Order by Pay"
   imposter = 0
   rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
   If Not rst.EOF Then
      If getMoney(player.ID) > rst!Pay Then 'can afford it
         DB.Execute "UPDATE SupplyDeck SET Seq =" & player.ID & " WHERE CardID = " & rst!CardID
         'add the card to the players deck
         DB.Execute "INSERT INTO PlayerSupplies (PlayerID, CardID) VALUES (" & player.ID & ", " & rst!CardID & ")"
         getMoney player.ID, rst!Pay * -1
         If rst!CrewID = 23 Or rst!CrewID = 41 Or rst!CrewID = 54 Then 'we have deception
            If haveCrewAnyone(23) And rst!CrewID <> 23 Then
               doDiscardCrew 28
               imposter = 23
            ElseIf haveCrewAnyone(41) And rst!CrewID <> 41 Then
               doDiscardCrew 70
               imposter = 41
            ElseIf haveCrewAnyone(54) And rst!CrewID <> 54 Then
               doDiscardCrew 100
               imposter = 54
            End If
         End If
         PutMsg player.PlayName & " hires " & rst!CrewName, player.ID, Logic!Gamecntr
         If imposter > 0 Then
            PutMsg getCrewName(0, imposter) & " has turned up as " & rst!CrewName & " on " & player.PlayName & "'s Ship"
         End If
                  
         getCrew = True
      End If
   End If
   rst.Close
End Function

Public Sub getRandomCrew(ByVal noOfCrew As Integer, ByVal leader)
Dim rst As New ADODB.Recordset, SQL, CrewID, maxCrewID, crewcnt

   maxCrewID = varDLookup("max(CrewID) AS maxcrew", "Crew", "Leader=0", "maxcrew")
   SQL = "SELECT SupplyDeck.CardID, SupplyDeck.Seq, Crew.* FROM Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID WHERE Crew.Leader=0 AND Seq > 4 AND Wanted = 0 AND Moral = 0  AND Crew.CrewID NOT IN (23,54)"
   If leader = 69 Then 'add Atherton check
      SQL = SQL & " AND Crew.Companion = 0"
   End If
   crewcnt = 0
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
   If rst.EOF Then  'all have Seq = 0
      rst.Close
      SQL = "SELECT SupplyDeck.CardID, SupplyDeck.Seq, Crew.* FROM Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID WHERE Crew.Leader=0 AND (Seq = 0 or Seq > 4)  AND Wanted = 0 AND Moral = 0  AND Crew.CrewID NOT IN (23,54)"
      If leader = 69 Then 'add Atherton check
         SQL = SQL & " AND Crew.Companion = 0"
      End If
      rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
   End If
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

Public Function getRandomLeader()
Dim rst As New ADODB.Recordset, SQL, CrewID, maxCrewID, crewcnt, noOfCrew

   maxCrewID = varDLookup("max(CrewID) AS maxcrew", "Crew", "Leader=1", "maxcrew")
   SQL = "SELECT SupplyDeck.CardID, SupplyDeck.Seq, Crew.* FROM Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID WHERE Crew.Leader=1 AND Seq =0"
   
   noOfCrew = 1
   crewcnt = 0
   rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
   Randomize Timer
   While crewcnt < noOfCrew
      rst.Requery
      CrewID = Int((maxCrewID * Rnd)) + 1
      rst.filter = "CrewID =" & CrewID
      If Not rst.EOF Then
         getRandomLeader = CrewID
         crewcnt = crewcnt + 1
         PutMsg player.PlayName & " has chosen " & rst!CrewName, player.ID
      Else
         DoEvents
      End If
   Wend
   rst.Close
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

Public Sub dealDriveAndJobs(ByVal playerID)
Dim rst As New ADODB.Recordset
Dim startjobs As String, a() As String, x, y, msg As String

   'std Drive Core IDs 132 - 135
   DB.Execute "INSERT INTO PlayerSupplies (PlayerID,CardID) VALUES (" & playerID & ", " & 131 + playerID & ")"
   DB.Execute "Update SupplyDeck SET Seq = " & playerID & " WHERE CardID = " & 131 + playerID
   
   'get Story Issued Job
   x = Nz(varDLookup("IssueJobID", "StoryGoals", "StoryID=" & Logic!StoryID & " and Goal = 0"), 0)
   If x > 0 Then
      assignDeal playerID, x
   End If
   'msg = Nz(varDLookup("Instructions", "StoryGoals", "StoryID=" & Logic!StoryID & " and Goal = 0"))
   'If msg <> "" Then
   '   MessBox msg, "Story - First Goal", "Shiny", "", getLeader()
   'End If
   
   'Grab a Job from configured list out of Contact decks
   startjobs = Nz(varDLookup("StartingJobs", "Story", "StoryID=" & Logic!StoryID), "")
   'possible future change to give optional of ALL Contact Jobs.  Use frmDeals to select 3 from the 5 Contacts
   If startjobs = "" Then Exit Sub
   
   rst.Open "SELECT * FROM ContactDeck WHERE ContactID > 0 AND Seq > " & CStr(CONSIDERED) & " ORDER BY ContactID, Seq", DB, adOpenStatic, adLockReadOnly
   y = 0
   a = Split(startjobs, ",")
   For x = LBound(a) To UBound(a)
      rst.Find "ContactID = " & a(x)
      If Not rst.EOF Then
         assignDeal playerID, rst!CardID
         y = y + 1
      End If
      If y > 2 Then Exit For
   Next x

   rst.Close
   Set rst = Nothing
   
End Sub

Public Function setNextPlayer(ByVal playerID)
Dim rst As New ADODB.Recordset
    Logic.Requery
    rst.CursorLocation = adUseClient
    rst.Open "SELECT * FROM Players WHERE NAME IS NOT NULL ORDER BY PlayerID", DB, adOpenDynamic, adLockOptimistic
    rst.Find "PlayerID = " & playerID

    If Not rst.EOF Then
       'set my cntr to current Game Seq
       DB.Execute "UPDATE Players SET Seq = " & CStr(Logic!Gamecntr) & " WHERE PlayerID = " & playerID
       'rst!Seq = Logic!GameCntr 'set my go as done
       'rst.Update
       rst.MoveNext
       If rst.EOF Then   'end of this round
         rst.Requery
         rst.MoveFirst
       End If
       setNextPlayer = rst!playerID
       DB.Execute "UPDATE GameSeq SET Player=" & CStr(setNextPlayer)
       'Logic.Requery
       'Logic.Update "Player", setNextPlayer
       
       If rst!Seq = Logic!Gamecntr Then  'round over, increment GameCntr
          DB.Execute "UPDATE GameSeq SET GameCntr = " & CStr(Logic!Gamecntr + 1)
          'Logic!GameCntr = Logic!GameCntr + 1
          'Logic.Update
       End If
       
       
   End If
End Function

Public Function setNextPlayerREV(ByVal playerID, Optional ByVal nextStatus As String = "")
Dim rst As New ADODB.Recordset
    Logic.Requery
    rst.CursorLocation = adUseClient
    rst.Open "SELECT * FROM Players WHERE NAME IS NOT NULL ORDER BY PlayerID DESC", DB, adOpenDynamic, adLockOptimistic
    rst.Find "PlayerID = " & playerID

    If Not rst.EOF Then
       'set my cntr to current Game Seq
       DB.Execute "UPDATE Players SET Seq = " & CStr(Logic!Gamecntr) & " WHERE PlayerID = " & playerID
       'rst!Seq = Logic!GameCntr 'set my go as done
       'rst.Update
       rst.MoveNext
       If rst.EOF Then   'end of this round
         rst.Requery
         rst.MoveFirst
       End If
       
       If rst!Seq = Logic!Gamecntr Then  'round over, increment GameCntr
          setNextPlayerREV = player.ID
          If nextStatus <> "" Then
             DB.Execute "UPDATE GameSeq SET Seq = '" & nextStatus & "'"
          End If
          DB.Execute "UPDATE GameSeq SET Player=" & CStr(player.ID) & ", GameCntr = " & CStr(Logic!Gamecntr + 1)
          'Logic!player = player.ID
          'Logic!GameCntr = Logic!GameCntr + 1
          'Logic.Update
          'Logic.Requery
       Else
          setNextPlayerREV = rst!playerID
          DB.Execute "UPDATE GameSeq SET Player = " & CStr(setNextPlayerREV)
          'Logic.Requery
          'Logic.Update "Player", setNextPlayerREV
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
      DB.Execute "UPDATE GameSeq SET Seq = '" & nextStatus & "', Player = " & CStr(setPlayer)
   End If
   
End Function

Public Sub assignDeal(ByVal playerID, ByVal CardID)

      DB.Execute "UPDATE ContactDeck SET Seq =" & playerID & " WHERE CardID = " & CardID
      'add the card to the players deck
      DB.Execute "INSERT INTO PlayerJobs (PlayerID, CardID) VALUES (" & playerID & ", " & CardID & ")"

End Sub

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


Public Function getAdjacentRows(ByVal SectorID) As String
   If SectorID < 1 Then Exit Function
   getAdjacentRows = varDLookup("AdjacentRows", "Board", "SectorID=" & SectorID)

End Function

Public Function getNextSector(ByVal fromSectorID, ByVal toSectorID As Integer, ByVal canMosey) As Integer
Dim rst As New ADODB.Recordset
Dim SQL, b(1 To 2) As Long, x As Integer, y As Long, z As Long, adjacent

   SQL = "SELECT SectorID, Board.STop, Board.SLeft, Board.SHeight, Board.SWidth "
   SQL = SQL & "FROM Board "
   SQL = SQL & "WHERE SectorID =" & toSectorID
   

   rst.Open SQL, DB, adOpenDynamic, adLockReadOnly
   If Not rst.EOF Then
      'Find the closest Player
      b(1) = Int(rst!SHeight / 2 + rst!STop)  'X
      b(2) = Int(rst!SWidth / 2 + rst!SLeft)  'Y
   End If
   rst.Close
   y = -1
   adjacent = getAdjacentRows(fromSectorID)
   SQL = "SELECT Board.* FROM Board WHERE Board.SectorID IN (" & adjacent & ")"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   Do While Not rst.EOF
      If (getCutterSector(rst!SectorID) > 0 And Not (canMosey And toSectorID = rst!SectorID)) Or beenHereBefore(rst!SectorID) Then
         'Beep
      ElseIf toSectorID = rst!SectorID Then 'its next to us, just go there
         getNextSector = rst!SectorID
         Exit Do
      Else
         'find the adjacent sector closest to the closest player
         z = Int(Sqr((b(1) - Int(rst!SHeight / 2 + rst!STop)) ^ 2 + (b(2) - Int(rst!SWidth / 2 + rst!SLeft)) ^ 2))
         If y = -1 Or y > z Then
            y = z
            getNextSector = rst!SectorID
         End If
      End If
      rst.MoveNext
   Loop
   rst.Close
   Set rst = Nothing
End Function

Public Function beenHereBefore(ByVal SectorID As Integer) As Boolean
Dim x
   For x = 0 To 8
      If Trail(x) = SectorID And SectorID > 0 Then
         beenHereBefore = True
         Exit For
      End If
   Next x
End Function

Public Function getBounty(ByVal SupplyID)
Dim rst As New ADODB.Recordset, x
Dim SQL

   SQL = "SELECT ContactDeck.CardID, ContactDeck.JobName, SupplyDeck.CardID AS CrewCardID "
   SQL = SQL & "FROM ContactDeck INNER JOIN (Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON ContactDeck.FugitiveID = Crew.CrewID "
   SQL = SQL & "WHERE ContactDeck.ContactID = 10 AND ContactDeck.Seq = 5 AND SupplyDeck.SupplyID = " & SupplyID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      DB.Execute "UPDATE ContactDeck SET Seq =" & player.ID & " WHERE CardID = " & rst!CardID
      'add the card to the players deck
      DB.Execute "INSERT INTO PlayerJobs (PlayerID, CardID) VALUES (" & player.ID & ", " & rst!CardID & ")"
      'remove crew from supply
      DB.Execute "UPDATE SupplyDeck SET Seq =0 WHERE CardID = " & rst!CrewCardID
      getBounty = rst!CardID
      PutMsg player.PlayName & " claims a Bounty " & rst!JobName, player.ID, Logic!Gamecntr
   End If
   rst.Close
   If Not IsEmpty(getBounty) Then
      'draw another one
'      SQL = "SELECT * FROM ContactDeck WHERE Seq > 5 AND ContactID = 10 Order by Seq"
'      rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
'      If Not rst.EOF Then
'         rst.Update "Seq", 5
'         PutMsg "New Bounty available"
'      End If
'      rst.Close
      If DrawDeck("Contact", 10, 1) Then PutMsg "New Bounty available"
   End If
Set rst = Nothing
End Function

Public Function getJob(ByVal ContactID)
Dim rst As New ADODB.Recordset, x
Dim SQL
   SQL = "SELECT * FROM ContactDeck WHERE Illegal=0 and Immoral=0 AND Seq > 5 AND ContactID = " & ContactID & " Order by Seq DESC"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      DB.Execute "UPDATE ContactDeck SET Seq =" & player.ID & " WHERE CardID = " & rst!CardID
      'add the card to the players deck
      DB.Execute "INSERT INTO PlayerJobs (PlayerID, CardID) VALUES (" & player.ID & ", " & rst!CardID & ")"
      getJob = rst!CardID
      PutMsg player.PlayName & " picks up a new Job " & getJob, player.ID, Logic!Gamecntr
   End If
   rst.Close
   SQL = "SELECT Seq FROM ContactDeck WHERE Seq > 5 AND ContactID = " & ContactID & " Order by Seq"
   rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
   x = 0
   Do While Not rst.EOF
      rst.Update "Seq", 5
      x = x + 1
      If x = 2 Then
         Exit Do
      End If
      rst.MoveNext
   Loop
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
      getJobCrewBonus = Nz(rst!Pay, 0)
   End If
   rst.Close
   Set rst = Nothing
   
End Function

Public Function getJobSector(ByVal CardID, ByVal JobID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT J.* FROM ContactDeck C, Job J WHERE C.Job" & JobID & "ID = J.JobID AND C.CardID = " & CardID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getJobSector = rst!SectorID
      If getJobSector = 1 Then getJobSector = getCruiserSector
      If getJobSector = 2 Then getJobSector = getCorvetteSector
   End If
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

Public Function useHavens(ByVal StoryID) As Boolean
   
   useHavens = (varDLookup("Havens", "Story", "StoryID=" & StoryID) = "1")

End Function

Public Function getHaven(ByVal SectorID) As Integer
   
   getHaven = varDLookup("Haven", "Board", "SectorID=" & SectorID)

End Function

Public Sub placeHaven(ByVal playerID, ByVal SectorID)
    DB.Execute "UPDATE Board Set Haven = " & Str(playerID) & " WHERE SectorID = " & SectorID
End Sub

Public Function getCrewPay()
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Sum(Crew.Pay) AS SumOfPay "
   SQL = SQL & "FROM PlayerSupplies INNER JOIN (Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob=0 AND PlayerSupplies.PlayerID=" & player.ID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getCrewPay = rst!SumOfPay
   End If
   rst.Close
   Set rst = Nothing
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
         If rst.EOF Then
            DB.Execute "UPDATE " & Deck & "Deck SET Seq = " & CStr(y) & " WHERE CardID=" & CStr(CardID)
         Else
            rst!Seq = y
            rst.Update
            rst.MoveNext
         End If
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

Public Function isBountyEnabled() As Boolean
   isBountyEnabled = (varDLookup("Bounty", "Story", "StoryID=" & Logic!StoryID) = 1)
End Function

Public Function isSolid(ByVal playerID, ByVal ContactID) As Boolean

   If ContactID = 0 Or playerID = 0 Then Exit Function
   isSolid = (varDLookup("Solid" & ContactID, "Players", "PlayerID=" & playerID) = 1)
   
   If Not isSolid And ContactID = 5 Then 'alliance card gives solid with Harken
      isSolid = hasGear(playerID, 20) Or hasCrew(playerID, 101) Or hasCrew(playerID, 103)
   End If
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
         Exit For
      End If
   Next z
   
   If adjacent = 0 Or (adjacent > 0 And NPCFlag) Then  'we found no players adjacent
      adjacent = getPursuitSector(SectorID, ship)
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

'move the Cruiser adjacent the sector given
Public Function doMoveCutterPlanetary(ByVal ship) As Boolean
Dim x
   
      Do
         x = RollDice(152)
        
         If Nz(varDLookup("PlanetID", "Planet", "SectorID=" & x), 0) > 0 And getClearSector(x) <> "" And getZone(x) <> "A" And x > 2 Then
            MoveShip ship, x
            PutMsg "A Cutter is sighted at " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & x), "the unknown..")
            doMoveCutterPlanetary = True
            Exit Do
         End If
      Loop

   
End Function

Public Function RollDice(Optional ByVal top As Integer = 6, Optional ByVal Thrill As Boolean = False) As Integer

   Randomize Timer
   RollDice = Int((top * Rnd) + 1)
   'CrewID 55 is Bester who blocks extra roll
   If RollDice = 6 And Thrill Then
      RollDice = RollDice + Int((top * Rnd) + 1)
   End If

End Function

Public Sub changeToken(ByVal SectorID As Integer, ByVal cnt As Integer, Optional ByVal sound As Boolean = True)
   If cnt < 0 Then
      DB.Execute "UPDATE Board SET Token = " & IIf(SectorID > 119 And SectorID < 123, "1", "0") & " WHERE SectorID = " & SectorID
   Else
      If sound Then playsnd 2
      DB.Execute "UPDATE Board Set Token = Token + " & Str(cnt) & " WHERE SectorID = " & SectorID
   End If

End Sub

Public Sub MoveShip(ByVal playerID, ByVal SectorID, Optional ByVal sound As Integer = 0, Optional ByVal syncsound As Boolean = False, Optional ByVal leaveToken As Boolean = True)
Dim rst As New ADODB.Recordset
Dim coords, slot, lastSectorID As Integer, x, a, b, TimingState As Boolean
Dim c() As String
   
   If SectorID = 0 Then Exit Sub
'   TimingState = Main.Timing.Enabled
'   Main.Timing.Enabled = False
   lastSectorID = getPlayerSector(playerID)
   slot = IIf(playerID > 4, 5, playerID)
   DB.BeginTrans
   DB.Execute "Update Players Set SectorID = " & SectorID & " WHERE PlayerID = " & playerID
   DB.CommitTrans

   If playerID > 6 And SectorID <> lastSectorID And lastSectorID > 0 And leaveToken Then    'cutter 7-12
      changeToken lastSectorID, 1, False    'leave another token behind
   End If
   
   If varDLookup("AllianceTrail", "Story", "StoryID=" & Logic!StoryID) > 0 Then
      If (playerID = 5 Or playerID = 6) And SectorID <> lastSectorID And lastSectorID > 0 And leaveToken Then
         changeAToken lastSectorID, 1, False      'leave another token behind
      End If
   End If

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
   
   If playerID = 6 Then 'moving the Corvette, check a Reaver is not here
      x = getCutterSector(SectorID)
      If x > 0 Then 'move this reaver back to Reaver Space
         PutMsg "The Corvette chases a Reaver Cutter off, which hightails it back to Reaver Space", playerID, Logic!Gamecntr
         'check there is room
         If getCutterSector(120) > 0 And getCutterSector(121) > 0 And getCutterSector(122) > 0 Then 'full house, goto 121 instead
            DB.Execute "UPDATE Players SET SectorID = 121 WHERE PlayerID = " & x
         Else
            'place it at Miranda and use the AI move to get it back to the Reaver Space with preference to any Player Ship :O
            DB.Execute "UPDATE Players SET SectorID = 123 WHERE PlayerID = " & x
         End If
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
   
'   Main.Timing.Enabled = TimingState
   
   'RefreshBoard
   setRefresh
'   If playerID > 4 And getPlayerSector(player.ID) = SectorID And actionSeq <> ASNavEvade Then
'      If checkWhisperX1(SectorID) Then
'         actionSeq = ASNavEvade ' and get away
'      End If
'   End If
   
End Sub

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


Public Function outlawExists(ByVal playerID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM Players WHERE Name IS NOT NULL AND PlayerID <> " & playerID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   Do While Not rst.EOF
      If isOutlaw(rst!playerID) And getZone(Nz(rst!SectorID, 0)) = "A" Then
         outlawExists = rst!SectorID
         Exit Do
      End If
      rst.MoveNext
   Loop
   rst.Close
   Set rst = Nothing
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

Public Sub changeAToken(ByVal SectorID As Integer, ByVal cnt As Integer, Optional ByVal sound As Boolean = True)
   If getAToken(SectorID) + cnt < 0 Then
      DB.Execute "UPDATE Board Set AToken = 0 WHERE SectorID = " & SectorID
   Else
      If Not getHaven(SectorID) > 0 Then
         If sound Then playsnd 2
         DB.Execute "UPDATE Board Set AToken = AToken + " & Str(cnt) & " WHERE SectorID = " & SectorID
      End If
   End If

End Sub

Public Function getAToken(ByVal SectorID As Integer) As Integer
   getAToken = varDLookup("AToken", "Board", "SectorID=" & SectorID)
End Function

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

'move the Cruiser adjacent the sector given
Public Function doMoveAllianceAdjacent(ByVal SectorID, Optional ByVal check As Boolean = False) As Boolean
Dim adjacent, a() As String, x, y
   
   adjacent = varDLookup("AdjacentRows", "Board", "SectorID=" & SectorID)
   a = Split(adjacent, ",")
   
   y = 0
   For x = LBound(a) To UBound(a)
      If getClearSector(Val(a(x))) = "A" And getHaven(Val(a(x))) = 0 Then   'no ship in this spot
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
         If getClearSector(Val(a(x))) = "A" And getHaven(Val(a(x))) = 0 Then
            MoveShip 5, Val(a(x))
            doMoveAllianceAdjacent = True
            Exit Do
         End If
      Loop
   End If

   
End Function

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

'move the Corvette Planetary
Public Function doMoveCorvettePlanetary() As Boolean
Dim x
   
      Do
         x = RollDice(152)
         If Nz(varDLookup("PlanetID", "Planet", "SectorID=" & x), 0) > 0 And getClearSector(x) <> "" And x > 2 Then
            MoveShip 6, x
            PutMsg "The Corvette is sighted at " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & x), " the unknown..")
            doMoveCorvettePlanetary = True
            Exit Do
         End If
      Loop

   
End Function

'move the Cruiser To a Free Sector
Public Function doMoveCruiserToFreeSector() As Boolean
Dim x
   
      Do
         x = RollDice(71)
         If Nz(varDLookup("Zones", "Board", "SectorID=" & x & " And Haven=0"), "B") = "A" And getClearSector(x) <> "" Then
            MoveShip 5, x
            PutMsg "The Cruiser is sighted at " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & x), "Sector " & x)
            doMoveCruiserToFreeSector = True
            Exit Do
         End If
      Loop

   
End Function

'returns which ship turns up if any
Public Function resolveToken(ByVal SectorID, Optional ByVal adjacent As Boolean = False) As Integer
Dim rst As New ADODB.Recordset, x, alliance As Boolean
Dim SQL
  
   SQL = "SELECT * FROM Board WHERE SectorID= " & SectorID
   rst.CursorLocation = adUseClient
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
            PutMsg player.PlayName & " " & IIf(adjacent, "scanned", "entered") & " a Sector on Alliance Alert Level " & CStr(rst!AToken) & ", and got a nasty surprise by rolling a " & x, player.ID, Logic!Gamecntr
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
            MoveShip resolveToken, SectorID
            PutMsg player.PlayName & " " & IIf(adjacent, "scanned", "entered") & " a Sector on Reaver Alert Level " & CStr(rst!Token) & ", and got a nasty surprise by rolling a " & x, player.ID, Logic!Gamecntr
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
      If mode = 2 Then
         SQL = SQL & " AND Crew.Leader =0"
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

Public Function doKillCrew(ByVal killCrew As Integer)
Dim rst As New ADODB.Recordset
Dim SQL, Dice As Integer, cnt As Integer
         
   SQL = "SELECT SupplyDeck.CardID, Crew.*"
   SQL = SQL & " FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID"
   SQL = SQL & " Where PlayerSupplies.playerID = " & player.ID & " And Crew.leader = 0"
   SQL = SQL & " ORDER BY Crew.Pay"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   Do While Not rst.EOF
      If hasCrewAttribute(player.ID, "Medic") Then
         Dice = RollDice(6)
      End If
      If Dice > 4 Then
         PutMsg player.PlayName & "'s Medic saved " & rst!CrewName, player.ID, Logic!Gamecntr
      Else
         'update their pile status - 0 removed, 5 -discarded
         DB.Execute "UPDATE SupplyDeck SET Seq =5 WHERE CardID = " & rst!CardID
         'remove any Gear first
         DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & rst!CrewID
         'delete the card to the players deck
         DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID =" & player.ID & " AND CardID = " & rst!CardID
         'clear disgruntled
         DB.Execute "UPDATE Crew SET Disgruntled = 0 WHERE CrewID = " & rst!CrewID
         PutMsg player.PlayName & " lost " & rst!CrewName & " in the meelee", player.ID, Logic!Gamecntr
      End If
      cnt = cnt + 1
      If cnt = killCrew Then Exit Do
      rst.MoveNext
   Loop
   rst.Close
End Function

'update their pile status - 0 removed, 5 -discarded
Public Sub doDiscardCrew(ByVal CardID, Optional ByVal status As Variant = 5)
Dim CrewID

   CrewID = varDLookup("CrewID", "SupplyDeck", "CardID=" & CardID)

   'update their pile status - 0 removed, 5 -discarded or playerid
   DB.Execute "UPDATE SupplyDeck SET Seq = " & status & " WHERE CardID = " & CardID
   'remove any Gear first
   DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & CrewID
   'delete the card to the players deck
   DB.Execute "DELETE FROM PlayerSupplies WHERE CardID = " & CardID

   
End Sub

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

'return the summated value of the Perk attribute
Public Function getPerkAttributeSum(ByVal playerID, ByVal Attrib As String) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   If Attrib = "" Then Exit Function
   'may need to manage "On Job" status
   SQL = "SELECT SUM(Perk." & Attrib & ") AS SumVal"
   SQL = SQL & " FROM Perk INNER JOIN (PlayerSupplies INNER JOIN (Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND PlayerSupplies.PlayerID=" & playerID & " AND Perk." & Attrib & " <> 0"
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       getPerkAttributeSum = Nz(rst!SumVal, 0)
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function getPlanetID(ByVal playerID) As Integer
Dim SectorID
   SectorID = varDLookup("SectorID", "Players", "PlayerID=" & playerID)
   getPlanetID = Nz(varDLookup("PlanetID", "Planet", "SectorID=" & SectorID), 0)
End Function

'get final balance, and optionally add & subtract money from player
Public Function getMoney(ByVal playerID, Optional ByVal change As Integer = 0) As Long
   If change <> 0 Then DB.Execute "UPDATE Players set Pay = Pay + " & change & " WHERE PlayerID = " & playerID
   getMoney = varDLookup("Pay", "Players", "PlayerID=" & playerID)
   
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

Public Function getNewPlayer() As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * "
   SQL = SQL & "FROM Players "
   SQL = SQL & "WHERE PlayerID < 5 and Name is Null ORDER BY PlayerID DESC"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getNewPlayer = rst!playerID
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function getStartSector() As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Planet.SectorID "
   SQL = SQL & "FROM Planet LEFT JOIN Players ON Planet.SectorID = Players.SectorID "
   SQL = SQL & "WHERE Planet.System='White Sun' AND Players.Name Is Null AND Planet.SectorID NOT IN (51,60)"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getStartSector = rst!SectorID
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function reconnectPlayer() As Integer
Dim rst As New ADODB.Recordset, cnt As Integer, c(4) As Integer, msg As String
Dim SQL
   SQL = "SELECT * "
   SQL = SQL & "FROM Players "
   SQL = SQL & "WHERE PlayerID < 5 and AI = 1"
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenStatic, adLockReadOnly
   While Not rst.EOF
      cnt = cnt + 1
      c(cnt) = rst!playerID
      reconnectPlayer = c(cnt)
      msg = msg & IIf(msg = "", "", " or ") & CStr(c(cnt))
      rst.MoveNext
   Wend
   rst.Close
   If cnt > 1 Then
      reconnectPlayer = Val(InputBox("There is more than 1 AI Bot, which one is this one?" & vbNewLine & msg, "Pick this AI Bot PlayerID", CStr(reconnectPlayer)))
   End If
   Set rst = Nothing
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

'Mode = 1 for Showdowns
Public Function getSkill(ByVal playerID, ByVal skill As String, Optional ByVal mode As Integer = 0) As Integer
Dim rst As New ADODB.Recordset, rst2 As New ADODB.Recordset
Dim SQL
   getSkill = 0
   'may need to manage "On Job" status
   SQL = "SELECT SupplyDeck.CardID, Crew.* "
   SQL = SQL & "FROM (Players INNER JOIN PlayerSupplies ON Players.PlayerID = PlayerSupplies.PlayerID) INNER JOIN (Crew INNER JOIN SupplyDeck "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob = 0 AND Players.PlayerID=" & playerID
   
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      getSkill = getSkill + rst.Fields(skill)
      'check for perk to add skill
      
      '+1 skill When carrying a Keyword (Firearm / Explosives)
      If getPerkAttributeCrew(playerID, skill, rst!CardID) > 0 And hasGearKeyword(playerID, hasPerkKeyword(playerID, rst!CardID), rst!CrewID) Then
         getSkill = getSkill + 1
      End If

      If rst!HillFolk = 1 Then
         'check for HillFolk fight bonus
         If countCrewAttribute(playerID, "HillFolk") > 2 And skill = cstrSkill(1) Then getSkill = getSkill + 1
      End If
      'Head Goon
      If countCrewAttribute(playerID, "Merc") > 2 And rst!CrewID = 65 And skill = cstrSkill(3) Then
         getSkill = getSkill + 2
      End If
      If rst!CrewID = 94 And getZone(getPlayerSector(playerID)) = "B" And skill = cstrSkill(1) Then   'Sheriff Bourne
         getSkill = getSkill + 2
      End If
      If mode = 1 And (rst!PerkID = 62 Or rst!PerkID = 64) And skill = cstrSkill(1) Then    'Marshal & Deputy & Ensign & Jubal
         getSkill = getSkill + IIf(rst!PerkID = 64, 2, 1)
      End If
     

      'grab skill from gear crew is carrying-----------------------------
      SQL = "SELECT Gear.* "
      SQL = SQL & "FROM (Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID) INNER JOIN Crew ON PlayerSupplies.CrewID = Crew.CrewID "
      SQL = SQL & "WHERE PlayerSupplies.CrewID=" & rst!CrewID

      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst2.EOF
        getSkill = getSkill + rst2.Fields(skill)
        rst2.MoveNext
      Wend
      rst2.Close
      '------------------------------------------------------------

      
      rst.MoveNext
   Wend
   rst.Close

   'Foreman + 2 mudders
   If countCrewAttribute(playerID, "Mudder") > 2 And skill = cstrSkill(1) And hasCrew(playerID, 76) Then getSkill = getSkill + 2
   
   Set rst = Nothing
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

Public Function hasWarrant() As Boolean
   hasWarrant = (varDLookup("Warrants", "Players", "PlayerID=" & player.ID) > 0)

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


'this Controls all the Story Goals, their Jobs and the WIN
Public Function CheckWon(ByVal playerID) As Boolean
Dim rst As New ADODB.Recordset, SQL, goaldone As Boolean
   
   goaldone = True
   SQL = "SELECT * FROM Players WHERE PlayerID=" & playerID
   
   rst.Open SQL, DB, adOpenDynamic, adLockReadOnly
   If Not rst.EOF Then
      While Not CheckWon And goaldone
         CheckWon = doGoalCheck(playerID, Logic!StoryID, rst!Goals, rst!Seq, goaldone)
         rst.Requery
      Wend
   End If

   rst.Close

   If CheckWon = True Then
      playsnd 5, True
      DB.Execute "INSERT INTO Scores (StoryID,PlayerName,Turns,StartDate,PlayDate) Values (" & CStr(Logic!StoryID) & ",'" & SQLFilter(player.PlayName) & "'," & CStr(Logic!Gamecntr - 1) & ", " & SQLDate(varDLookup("EventTime", "Events", "Event ='" & player.PlayName & "''s on the Map'")) & ", " & SQLNow & ")"

      'DB.Execute "INSERT INTO Scores (StoryID,PlayerName,Turns,StartDate,PlayDate) Values (" & CStr(Logic!StoryID) & ",'" & SQLFilter(player.PlayName) & "'," & CStr(Logic!GameCntr - 1) & ", #" & Format(varDLookup("EventTime", "Events", "Event ='" & player.PlayName & "''s on the Map'"), "MM-DD-YY HH:nn") & "#, #" & Format(Now, "MM-DD-YY HH:nn") & "#)"
      PutMsg PlayCode(playerID).PlayName & " has WON the Game in " & Logic!Gamecntr - 1 & " turns", playerID, Logic!Gamecntr
   End If


Set rst = Nothing
End Function

Private Function doGoalCheck(ByVal playerID, ByVal StoryID, ByVal Goal, ByVal Seq, ByRef goaldone As Boolean) As Boolean
Dim rst As New ADODB.Recordset, a() As String
Dim SQL, x, cnt As Integer, ContactID As Integer
   goaldone = False
   If Goal = -1 Then Exit Function
   goaldone = True 'until proven otherwise
   SQL = "SELECT * FROM StoryGoals WHERE StoryID=" & StoryID & " AND Goal = " & CStr(Goal + 1)
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
   
         'Bounties
      If goaldone And rst!Bounties > 0 Then
         If countBounties(playerID) < rst!Bounties Then
            goaldone = False
         End If
      End If
      
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
      If goaldone And rst!SectorID > 0 Then
         x = rst!SectorID
         If x = 1 Then x = getCruiserSector
         If x = 2 Then x = getCorvetteSector
         goaldone = (getPlayerSector(player.ID) = x)
      End If
      
      If rst!SolidCount = 0 And Nz(rst!Solid) = "" And rst!Cash = 0 And rst!Bounties = 0 And rst!IssueJobID = 0 And rst!CompleteJobID = 0 Then goaldone = False 'AI can never reach goal
     
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
     
      ' END of positive Tests ================================================
      
      'If we still good and we have Win flag, we WIN
      If goaldone And rst!win > 0 Then
         doGoalCheck = True
      End If
            
            'Negative tests ---- TurnLimit
      If rst!TurnLimit > 0 And Not doGoalCheck Then
         If Seq >= rst!TurnLimit Then
            addGoal playerID, -1
            PutMsg player.PlayName & " has Failed to meet the Story Goal Turn limit of " & rst!TurnLimit & ". GAME OVER!", player.ID, Seq
            goaldone = False
         End If
      End If
            
            'load any Passengers if there is room
      If goaldone And rst!Passenger > 0 Then
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= rst!Passenger Then
            DB.Execute "UPDATE Players SET Passenger = Passenger + " & CStr(rst!Passenger) & " WHERE PlayerID = " & playerID
         End If
      End If
      
      'if we here and goaldone then Goal IS Done
      If goaldone Then
         addGoal playerID, 1
         ContactList = getContactList(StoryID)
         'check if we get solid
         ContactID = rst!doSolid
         If ContactID = 0 Then
            'pass
         ElseIf ContactID = 5 And hasWarrant() Then
            'no solid
         Else
            SQL = "UPDATE Players SET "
            SQL = SQL & " Solid" & ContactID & "=1 "  'setting SOLID with the Contact
            PutMsg player.PlayName & " is solid with " & varDLookup("ContactName", "Contact", "ContactID=" & ContactID), playerID, Logic!Gamecntr
            SQL = SQL & " Where playerID = " & playerID
            DB.Execute SQL
         End If
      End If
      
      'if we here and goaldone, we good to deliver job
      If goaldone And rst!IssueJobID > 0 Then
         assignDeal playerID, rst!IssueJobID
      End If

      
       ' we good to give new instructions
      If goaldone And Nz(rst!Instructions) <> "" And Not doGoalCheck Then
         PutMsg player.PlayName & " has completed Goal " & Goal + 1 & vbNewLine & rst!Instructions, playerID, Logic!Gamecntr
      End If
   Else
      goaldone = False
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

Public Function countBounties(ByVal playerID) As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Count(*) AS cnt FROM PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID "
   SQL = SQL & "WHERE ContactDeck.ContactID=10 AND PlayerJobs.JobStatus=3 AND PlayerJobs.PlayerID=" & playerID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      countBounties = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
End Function

Public Function getContactList(ByVal StoryID) As String
Dim Goal As Integer, tmpCL As String
   Goal = varDLookup("Goals", "Players", "PlayerID = " & player.ID) + 1
   tmpCL = varDLookup("Solid", "StoryGoals", "StoryID = " & StoryID & " AND Goal =" & Goal) & ""
   If tmpCL <> "" Then
      getContactList = tmpCL
   Else
      getContactList = "1,2,4,5"
   End If
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

Public Function getBountyMaxSeq() As Integer
Dim SQL As String
Dim rst As New ADODB.Recordset

   SQL = "SELECT max(Seq) as MaxSeq "
   SQL = SQL & "FROM ContactDeck "
   SQL = SQL & "Where ContactDeck.ContactID = 10 and Seq > " & DISCARDED
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getBountyMaxSeq = rst!MaxSeq + 1
   End If
   rst.Close
   'in case there are none left face down
   If getBountyMaxSeq = 0 Then getBountyMaxSeq = 100

End Function

Public Function pushBounties() As Boolean 'back into bottom of deck
Dim SQL As String, MaxSeq As Integer
Dim rst As New ADODB.Recordset

   MaxSeq = getBountyMaxSeq
   
   DB.BeginTrans
   SQL = "SELECT CardID "
   SQL = SQL & "FROM ContactDeck "
   SQL = SQL & "Where ContactDeck.ContactID = 10 and ContactDeck.Seq = " & DISCARDED
   rst.Open SQL, DB, adOpenStatic, adLockReadOnly
   While Not rst.EOF
      MaxSeq = MaxSeq + 1
      'rst.Update "Seq", MaxSeq
      DB.Execute "UPDATE ContactDeck SET Seq =" & MaxSeq & " WHERE CardID =" & rst!CardID
      pushBounties = True
      rst.MoveNext
   Wend
   rst.Close
   DB.CommitTrans
End Function

Public Function DrawDeck(ByVal Deck As String, ByVal ID As Integer, ByVal draw As Integer, Optional ByVal Seq As Integer = DISCARDED) As Boolean
Dim rst As New ADODB.Recordset
Dim SQL, cnt
   cnt = 0
   SQL = "SELECT * FROM " & Deck & "Deck WHERE Seq > 6 AND " & Deck & "ID =" & CStr(ID) & " ORDER BY Seq"
   rst.Open SQL, DB, adOpenDynamic, adLockOptimistic
   Do While Not rst.EOF
      DrawDeck = True
      cnt = cnt + 1
      rst!Seq = Seq
      rst.Update
      If draw = cnt Then Exit Do
      rst.MoveNext
   Loop
   rst.Close
   
   Set rst = Nothing

End Function

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

Public Function bountiesDone() As Boolean
Dim rst As New ADODB.Recordset, x As Integer, y As Integer
Dim SQL
   
   SQL = "SELECT Max(Bounties) as cnt FROM StoryGoals WHERE StoryGoals.StoryID=" & Logic!StoryID
   rst.Open SQL, DB, adOpenStatic, adLockReadOnly
   If Not rst.EOF Then
      x = rst!cnt
   End If
   rst.Close
   Set rst = Nothing
   If x = 0 Then Exit Function 'no bounty count required for goals, so leave them active
   
   'ok there is a bounty requirement, keep going until reached
   y = countBounties(player.ID)
   bountiesDone = (y >= x)
   
End Function

Public Function hasGoalSector(ByRef goalSector As Integer) As Boolean
Dim rst As New ADODB.Recordset, x As Integer, y As Integer, a
Dim SQL, Goal As Integer

   Goal = varDLookup("Goals", "Players", "PlayerID = " & player.ID) + 1
   hasGoalSector = True
   SQL = "SELECT  SectorID, CompleteJobID, Solid, SolidCount, Bounties, NoUnfinished, Fight, Tech, Negotiate, Cash "
   SQL = SQL & "FROM StoryGoals WHERE StoryID=" & Logic!StoryID & " AND SectorID>0 AND Goal=" & Goal
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenStatic, adLockReadOnly
   If rst.EOF Then
      hasGoalSector = False
   Else
      goalSector = 0
      If hasGoalSector And rst!CompleteJobID > 0 Then
         If jobSuccess(player.ID, rst!CompleteJobID) Then
            hasGoalSector = False
         End If
      End If
      
      If hasGoalSector And Nz(rst!Solid) <> "" Then
         a = Split(rst!Solid, ",")
         For x = LBound(a) To UBound(a)
            If Not isSolid(player.ID, a(x)) Then
               hasGoalSector = False
               Exit For
            End If
         Next x
      End If
      y = 0
      If hasGoalSector And rst!SolidCount > 0 Then
         For x = 1 To NO_OF_CONTACTS
            If isSolid(player.ID, x) Then
               y = y + 1
            End If
         Next x
         If y < rst!SolidCount Then
            hasGoalSector = False
         End If
      End If
      If hasGoalSector And rst!Bounties > 0 Then
         If countBounties(player.ID) < rst!Bounties Then
            hasGoalSector = False
         End If
      End If
      
      If hasGoalSector Then goalSector = rst!SectorID
      If goalSector = 1 Then goalSector = getCruiserSector
      If goalSector = 2 Then goalSector = getCorvetteSector
      
   End If
   rst.Close
   Set rst = Nothing

End Function

Public Sub getJob2Reqs(ByVal CardID, ByRef cargo As Integer, ByRef contra As Integer)
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Job.* FROM Job INNER JOIN ContactDeck ON Job.JobID = ContactDeck.Job2ID WHERE ContactDeck.CardID=" & CardID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      cargo = Abs(rst!cargo)
      contra = Abs(rst!Contraband)
   End If
   rst.Close
   Set rst = Nothing
End Sub

Public Sub setRefresh(Optional ByVal All As Boolean = False)
Dim x
   For x = 1 To 4
      If Not All And x = player.ID Then
         'skip
      Else
         DB.Execute "UPDATE GameSeq SET Refresh" & x & " = 1"
      End If
   Next x

End Sub

Public Sub clearRefresh(Optional ByVal All As Boolean = False)
Dim x
   For x = 1 To 4
      If Not All And x <> player.ID Then
         'skip
      Else
         DB.Execute "UPDATE GameSeq SET Refresh" & x & " = 0"
      End If
   Next x

End Sub
