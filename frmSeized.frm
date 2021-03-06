VERSION 5.00
Begin VB.Form frmSeized 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select the Crew Member seized by the Alliance Corvette"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSeized.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox crewList 
      BackColor       =   &H00CBE1ED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   2610
      Width           =   5415
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Select"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Height          =   2265
      Left            =   60
      Picture         =   "frmSeized.frx":D3822
      ScaleHeight     =   2205
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   90
      Width           =   5415
   End
End
Attribute VB_Name = "frmSeized"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nbrSelect, discard As Boolean

Private Sub cmd_Click()
Dim x, cnt
   cnt = 0
   playsnd 8
   For x = 0 To crewList.ListCount - 1
      If crewList.Selected(x) Then
         cnt = cnt + 1
      End If
   Next x
   If cnt = nbrSelect Then
      For x = 0 To crewList.ListCount - 1
         If crewList.Selected(x) Then
         
            'update their pile status - 0 removed, 5 -discarded
            DB.Execute "UPDATE SupplyDeck SET Seq =" & IIf(hasPerkAttribute(player.ID, "KillDiscard", crewList.ItemData(x)) > 0 Or discard, 5, 0) & " WHERE CardID = " & crewList.ItemData(x)
            'remove any Gear first
            DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = (SELECT CrewID FROM SupplyDeck WHERE CardID =" & crewList.ItemData(x) & ")"
            'delete the card to the players deck
            DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID =" & player.ID & " AND CardID = " & crewList.ItemData(x)
         
         End If
      Next x
      Me.Hide
   Else
      MsgBox "You need to select " & nbrSelect & " crew", vbExclamation
   End If

End Sub


Private Sub crewList_DblClick()
Dim frmCrew As New frmCrewSel
      If crewList.ListIndex > -1 Then
         frmCrew.crewFilter = " INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID WHERE SupplyDeck.CardID=" & crewList.ItemData(crewList.ListIndex)
         frmCrew.Show 1
         Set frmCrew = Nothing
      End If
End Sub

Private Sub Form_Load()
   nbrSelect = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

'return how many crew up for selection
Public Function RefreshList(ByVal check As Boolean) As Integer
Dim rst As New ADODB.Recordset
Dim SQL, crewcnt As Integer
   crewList.Clear
   SQL = "SELECT PlayerSupplies.CardID, Crew.* "
   SQL = SQL & "FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & player.ID & " AND Crew.Wanted=1"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If hasGear(player.ID, 20, rst!crewID) Then  'you're on the list
         PutMsg player.PlayName & "'s Nav log: " & rst!CrewName & " Flashes an Alliance Ident Card", player.ID, Logic!Gamecntr, True, rst!crewID
      ElseIf hasShipUpgrade(player.ID, 11) And crewcnt < 2 And check Then
         If crewcnt = 0 Then PutMsg player.PlayName & "'s Nav log: Concealed Smuggling Compartments hides up to 2 Wanted Crew", player.ID, Logic!Gamecntr, True, getLeader()
         crewcnt = crewcnt + 1
      Else
         crewList.AddItem rst!CrewName
         crewList.ItemData(crewList.NewIndex) = rst!CardID
      End If
      rst.MoveNext
   Wend
   rst.Close
   
   RefreshList = crewList.ListCount


End Function

'return how many crew up for selection
Public Function RefreshDiscardList() As Integer
Dim rst As New ADODB.Recordset
Dim SQL, crewcnt As Integer
   discard = True
   crewList.Clear
   SQL = "SELECT PlayerSupplies.CardID, Crew.* "
   SQL = SQL & "FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & player.ID & " AND Crew.Leader=0"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      crewList.AddItem rst!CrewName
      crewList.ItemData(crewList.NewIndex) = rst!CardID
      rst.MoveNext
   Wend
   rst.Close
   
   RefreshDiscardList = crewList.ListCount


End Function
