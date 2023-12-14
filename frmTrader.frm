VERSION 5.00
Begin VB.Form frmTrader 
   BackColor       =   &H00CBE1ED&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Trade"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTrader.frx":0000
   ScaleHeight     =   5940
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   13
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "0"
      Top             =   1110
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   12
      Left            =   4920
      TabIndex        =   33
      Text            =   "0"
      Top             =   1110
      Width           =   700
   End
   Begin VB.Timer Timing 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   10260
      Top             =   5400
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Offer"
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
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5400
      Width           =   1035
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
      Index           =   1
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5400
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5400
      Width           =   1035
   End
   Begin VB.ListBox lstSupplies 
      BackColor       =   &H00CBE1ED&
      Height          =   3570
      Index           =   1
      Left            =   6240
      TabIndex        =   28
      Top             =   1500
      Width           =   5685
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   7140
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   780
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   7140
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   450
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   9
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   780
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   8
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   450
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   11
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   780
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   10
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0"
      Top             =   450
      Width           =   700
   End
   Begin VB.ListBox lstSupplies 
      BackColor       =   &H00CBE1ED&
      Height          =   3660
      Index           =   0
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   1500
      Width           =   5685
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1050
      TabIndex        =   5
      Text            =   "0"
      Top             =   780
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1050
      TabIndex        =   4
      Text            =   "0"
      Top             =   450
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Text            =   "0"
      Top             =   780
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Text            =   "0"
      Top             =   450
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   1
      Text            =   "0"
      Top             =   780
      Width           =   700
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   0
      Text            =   "0"
      Top             =   450
      Width           =   700
   End
   Begin VB.Label lblAccept 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Offer Accepted"
      Height          =   315
      Left            =   1620
      TabIndex        =   37
      Top             =   5425
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      Height          =   225
      Index           =   13
      Left            =   10290
      TabIndex        =   36
      Top             =   1140
      Width           =   705
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      Height          =   225
      Index           =   12
      Left            =   4170
      TabIndex        =   34
      Top             =   1140
      Width           =   705
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "other Player's Trade Items"
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
      Index           =   1
      Left            =   6480
      TabIndex        =   29
      Top             =   90
      Width           =   5115
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraband"
      Height          =   225
      Index           =   11
      Left            =   6210
      TabIndex        =   27
      Top             =   810
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      Height          =   225
      Index           =   10
      Left            =   6210
      TabIndex        =   26
      Top             =   450
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Fugitives"
      Height          =   225
      Index           =   9
      Left            =   8220
      TabIndex        =   25
      Top             =   810
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Passengers"
      Height          =   225
      Index           =   8
      Left            =   8220
      TabIndex        =   24
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Parts"
      Height          =   225
      Index           =   6
      Left            =   10290
      TabIndex        =   23
      Top             =   810
      Width           =   705
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel"
      Height          =   225
      Index           =   5
      Left            =   10290
      TabIndex        =   22
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Crew / Gear / Upgrades"
      Height          =   285
      Index           =   4
      Left            =   6240
      TabIndex        =   21
      Top             =   1260
      Width           =   2355
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Trade Items"
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
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   90
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      Height          =   5055
      Left            =   5940
      Top             =   60
      Width           =   135
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraband"
      Height          =   225
      Index           =   15
      Left            =   90
      TabIndex        =   12
      Top             =   810
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      Height          =   225
      Index           =   14
      Left            =   90
      TabIndex        =   11
      Top             =   450
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Fugitives"
      Height          =   225
      Index           =   0
      Left            =   2100
      TabIndex        =   10
      Top             =   810
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Passengers"
      Height          =   225
      Index           =   1
      Left            =   2100
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Parts"
      Height          =   225
      Index           =   2
      Left            =   4170
      TabIndex        =   8
      Top             =   810
      Width           =   705
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel"
      Height          =   225
      Index           =   3
      Left            =   4170
      TabIndex        =   7
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Crew / Gear / Upgrades"
      Height          =   285
      Index           =   7
      Left            =   150
      TabIndex        =   6
      Top             =   1260
      Width           =   2355
   End
End
Attribute VB_Name = "frmTrader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TraderID As Integer, isHost As Boolean

Private Sub cmd_Click(Index As Integer)
   Timing.Enabled = False
   Select Case Index
   Case 0
      cmd(0).Enabled = False
      ClearTrade
      Me.Hide
   Case 1 'accept
      If validOffer Then
         cmd(1).Enabled = False
         cmd(2).Enabled = False
         cmd(2).Caption = "Done"
         acceptOffer
      End If
      Timing.Enabled = True
   Case 2 'offer
      If cmd(2).Caption = "Change" Then
         clearOffer
         LockForm False
         cmd(1).Enabled = False
         cmd(2).Caption = "Offer"
      ElseIf saveOffer Then
         LockForm
         'cmd(1).Enabled = True
         cmd(2).Caption = "Change"
      
      End If
      Timing.Enabled = True
   End Select
End Sub

Private Sub Form_Load()
   refreshSupplies
   refreshGoodsSupplies TraderID
   refreshTradeSupplies TraderID
   
   Timing.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub Timing_Timer()
   refreshGoodsSupplies TraderID
   refreshTradeSupplies TraderID
   
   If Logic!trader = 0 Then 'deals off or finished
      ClearTrade
      Me.Hide
   ElseIf cmd(2).Caption <> "Offer" Then 'offer outstanding
      Logic.Requery
      lblAccept.Visible = (isHost And Logic!ClientAccept = 1) Or (Not isHost And Logic!HostAccept = 1)
      cmd(2).Enabled = (Not lblAccept.Visible And cmd(2).Caption <> "Done")
      If Logic!HostAccept = 1 And Logic!ClientAccept = 1 Then
         cmd(0).Enabled = False
         If isHost Then
            processOffer
         End If
      End If
   End If
End Sub

Private Sub LockForm(Optional ByVal locks As Boolean = True)
Dim x
   For x = 0 To 5
      txt(x).Locked = locks
   Next x
   txt(12).Locked = locks
   lstSupplies(0).Enabled = Not locks

End Sub

Private Sub acceptOffer()
Dim errh
On Error GoTo err_handler

   Logic.Requery
   If isHost Then
      DB.Execute "UPDATE GameSeq SET HostAccept = 1"
      'Logic.Update "HostAccept", 1
   Else
      DB.Execute "UPDATE GameSeq SET ClientAccept = 1"
      'Logic.Update "ClientAccept", 1
   End If
   
   Exit Sub
   
err_handler:
  errh = MsgBox(Err.Description, vbCritical + vbAbortRetryIgnore, "Error in Accepting Offer")
  Select Case errh
  Case vbRetry
    Resume
  Case vbAbort
    'exit
  Case vbIgnore
    Resume Next
  End Select
  
End Sub

Private Sub processOffer()
Dim x, SQL, CrewID
   Timing.Enabled = False
   'transfer goods.
   SQL = "UPDATE Players SET Pay = Pay + " & CStr(Val(txt(13))) & " - " & CStr(Val(txt(12))) & ", Fuel = Fuel + " & CStr(Val(txt(10))) & " - " & CStr(Val(txt(4))) & _
   ", Parts = Parts + " & CStr(Val(txt(11))) & " - " & CStr(Val(txt(5))) & ", Cargo = Cargo + " & CStr(Val(txt(6))) & " - " & CStr(Val(txt(0))) & _
   ", Passenger = Passenger + " & CStr(Val(txt(8))) & " - " & CStr(Val(txt(2))) & ", Contraband = Contraband + " & CStr(Val(txt(7))) & " - " & CStr(Val(txt(1))) & _
   ", Fugitive = Fugitive + " & CStr(Val(txt(9))) & " - " & CStr(Val(txt(3))) & _
   " WHERE PlayerID = " & player.ID
   DB.Execute SQL
   
   
   SQL = "UPDATE Players SET Pay = Pay + " & CStr(Val(txt(12))) & " - " & CStr(Val(txt(13))) & ", Fuel = Fuel + " & CStr(Val(txt(4))) & " - " & CStr(Val(txt(10))) & _
   ", Parts = Parts + " & CStr(Val(txt(5))) & " - " & CStr(Val(txt(11))) & ", Cargo = Cargo + " & CStr(Val(txt(0))) & " - " & CStr(Val(txt(6))) & _
   ", Passenger = Passenger + " & CStr(Val(txt(2))) & " - " & CStr(Val(txt(8))) & ", Contraband = Contraband + " & CStr(Val(txt(1))) & " - " & CStr(Val(txt(7))) & _
   ", Fugitive = Fugitive + " & CStr(Val(txt(3))) & " - " & CStr(Val(txt(9))) & _
   " WHERE PlayerID = " & TraderID
   DB.Execute SQL
   
   
   With lstSupplies(0)
   For x = 0 To .ListCount - 1
      If .selected(x) Then
         SQL = "UPDATE PlayerSupplies Set CrewID=0, PlayerID = " & TraderID & " WHERE CardID=" & CStr(.ItemData(x))
         DB.Execute SQL
         CrewID = getCrewID(.ItemData(x))
         If CrewID > 0 Then DB.Execute "UPDATE PlayerSupplies Set CrewID=0 WHERE CrewID=" & CrewID
         SQL = "UPDATE SupplyDeck Set Seq = " & TraderID & " WHERE CardID=" & CStr(.ItemData(x))
         DB.Execute SQL
      End If
   Next x
   End With
   With lstSupplies(1)
   For x = 0 To .ListCount - 1
         SQL = "UPDATE PlayerSupplies Set CrewID=0, PlayerID = " & player.ID & " WHERE CardID=" & CStr(.ItemData(x))
         DB.Execute SQL
         CrewID = getCrewID(.ItemData(x))
         If CrewID > 0 Then DB.Execute "UPDATE PlayerSupplies Set CrewID=0 WHERE CrewID=" & CrewID
         SQL = "UPDATE SupplyDeck Set Seq = " & player.ID & " WHERE CardID=" & CStr(.ItemData(x))
         DB.Execute SQL
   Next x
   End With
   ClearTrade
   Me.Hide
End Sub

Private Sub ClearTrade()
   Screen.MousePointer = vbHourglass
   Timing.Enabled = False
   'clear trade
   If isHost Then
      DB.Execute "DELETE FROM TradeGoods"
      DB.Execute "DELETE FROM TradeSupplies"
   End If
   DB.Execute "UPDATE GameSeq SET HostAccept = 0, ClientAccept = 0, Trader = 0"
   'Logic!HostAccept = 0
   'Logic!ClientAccept = 0
   'Logic!trader = 0
   'Logic.Update
   Screen.MousePointer = vbNormal
End Sub

Private Sub clearOffer()
   Screen.MousePointer = vbHourglass
   Timing.Enabled = False
   'clear trade
   
   DB.Execute "DELETE FROM TradeGoods WHERE PlayerID = " & player.ID
   DB.Execute "DELETE FROM TradeSupplies WHERE PlayerID = " & player.ID
  
   Screen.MousePointer = vbNormal
End Sub

Private Function validOffer() As Boolean
Dim u, v, x, y, z
   x = Val(txt(4)) / 2
   x = x + Val(txt(5)) / 2
   x = x + Val(txt(0))
   x = x + Val(txt(2))
   x = x + Val(txt(1))
   x = x + Val(txt(3))
   y = Val(txt(10)) / 2
   y = y + Val(txt(11)) / 2
   y = y + Val(txt(6))
   y = y + Val(txt(7))
   y = y + Val(txt(8))
   y = y + Val(txt(9))
   If x > y Then  'you are giving more
      If CargoCapacity(TraderID) - CargoSpaceUsed(TraderID) < (x - y) Then 'no room
         MsgBox "Other Trader has not enough room for extra Goods", vbExclamation
         Exit Function
      End If
   ElseIf y > x Then 'you are getting more
      If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < (y - x) Then 'no room
         MsgBox "You have not enough room for extra Goods", vbExclamation
         Exit Function
      End If
   End If
   u = 0 'shpup
   v = 0
   'crew count
   x = 0
   y = 0
   With lstSupplies(0)
   For z = 0 To .ListCount - 1
      If .selected(z) And getCrewID(.ItemData(z)) > 0 Then
         x = x + 1
      End If
      
      If .selected(z) And getShipUpgradeID(.ItemData(z)) > 0 Then
         u = u + 1
      End If
   Next z
   End With
   With lstSupplies(1)
   For z = 0 To .ListCount - 1
      If getCrewID(.ItemData(z)) > 0 Then
         y = y + 1
      End If
      
      If getShipUpgradeID(.ItemData(z)) > 0 Then
         v = v + 1
      End If
   Next z
   End With

   If y > x Then
      If CrewCapacity(player.ID) - getCrewCount(player.ID) < (y - x) Then
         MsgBox "You have not enough room for the extra Crew", vbExclamation
         Exit Function
      End If
   ElseIf x > y Then 'receiver getting more
      If CrewCapacity(TraderID) - getCrewCount(TraderID) < (x - y) Then
         MsgBox "Other Trader has not enough room for the extra Crew", vbExclamation
         Exit Function
      End If
   End If
   
   If v > u Then
      If 3 - getShipUpgrades(player.ID) < (v - u) Then
         MsgBox "You have not enough room for the extra Ship Upgrade", vbExclamation
         Exit Function
      End If
   ElseIf u > v Then 'receiver getting more
      If 3 - getShipUpgrades(TraderID) < (u - v) Then
         MsgBox "Other Trader has not enough room for the extra Ship Upgrade", vbExclamation
         Exit Function
      End If
   End If
   
   
   validOffer = True
End Function

'return true if valid offer
Private Function saveOffer() As Boolean
Dim rst As New ADODB.Recordset
Dim SQL, x As Integer

   SQL = "SELECT * FROM Players "
   SQL = SQL & "WHERE PlayerID=" & player.ID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      If rst!pay < Val(txt(12)) Then
         MsgBox "You don't have that much Cash!", vbExclamation
         Exit Function
      End If
      If rst!fuel < Val(txt(4)) Then
         MsgBox "You don't have that much Fuel!", vbExclamation
         Exit Function
      End If
      If rst!parts < Val(txt(5)) Then
         MsgBox "You don't have that many Parts!", vbExclamation
         Exit Function
      End If
      If rst!cargo < Val(txt(0)) Then
         MsgBox "You don't have that much Cargo!", vbExclamation
         Exit Function
      End If
      If rst!Contraband < Val(txt(1)) Then
         MsgBox "You don't have that much Contraband!", vbExclamation
         Exit Function
      End If
      If rst!Passenger < Val(txt(2)) Then
         MsgBox "You don't have that many Passengers!", vbExclamation
         Exit Function
      End If
      If rst!Fugitive < Val(txt(3)) Then
         MsgBox "You don't have that many Fugitives!", vbExclamation
         Exit Function
      End If
   End If


   SQL = "INSERT INTO TradeGoods (PlayerID, Pay, Fuel, Parts, Cargo, Contraband, Passenger, Fugitive) Values ("
   SQL = SQL & player.ID
   SQL = SQL & ", " & CStr(Val(txt(12)))
   SQL = SQL & ", " & CStr(Val(txt(4)))
   SQL = SQL & ", " & CStr(Val(txt(5)))
   SQL = SQL & ", " & CStr(Val(txt(0)))
   SQL = SQL & ", " & CStr(Val(txt(1)))
   SQL = SQL & ", " & CStr(Val(txt(2)))
   SQL = SQL & ", " & CStr(Val(txt(3))) & ")"
   DB.Execute SQL
   
   With lstSupplies(0)
   For x = 0 To .ListCount - 1
      If .selected(x) Then
         SQL = "INSERT INTO TradeSupplies (PlayerID, CardID) Values (" & player.ID & ", " & CStr(.ItemData(x)) & ")"
         DB.Execute SQL
      End If
   Next x
   End With
   
   saveOffer = True
End Function


Private Sub refreshSupplies()
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT PlayerSupplies.CardID, Crew.CrewID, Crew.CrewName, Crew.CrewDescr, Crew.Pay, Crew.Leader, Gear.GearName, Gear.GearDescr, Gear.Pay, "
   SQL = SQL & "ShipUpgrade.DriveCore, ShipUpgrade.UpgradeName, ShipUpgrade.UpgradeDescr, ShipUpgrade.Pay "
   SQL = SQL & "FROM ShipUpgrade RIGHT JOIN (Gear RIGHT JOIN (Crew RIGHT JOIN (PlayerSupplies INNER JOIN SupplyDeck ON "
   SQL = SQL & "PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) ON Gear.GearID = SupplyDeck.GearID) ON "
   SQL = SQL & "ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
   SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & player.ID
   SQL = SQL & " ORDER BY Crew.CrewName & 'zz', Gear.GearName & 'zz', ShipUpgrade.UpgradeName & 'zz'"
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If Nz(rst!leader, 0) = 0 And Nz(rst!DriveCore, 0) = 0 Then
         lstSupplies(0).AddItem Trim(Nz(rst!CrewName) & " " & Nz(rst!CrewDescr) & Nz(rst!GearName) & " " & Nz(rst!GearDescr) & Nz(rst!UpgradeName) & " " & Nz(rst!UpgradeDescr))
         
         lstSupplies(0).ItemData(lstSupplies(0).NewIndex) = rst!CardID
      End If
      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing
End Sub

Private Sub refreshTradeSupplies(ByVal playerID)
Dim rst As New ADODB.Recordset
Dim SQL
   lstSupplies(1).Clear
   SQL = "SELECT TradeSupplies.CardID, Crew.CrewID, Crew.CrewName, Crew.CrewDescr, Crew.Pay, Crew.Leader, Gear.GearName, Gear.GearDescr, Gear.Pay, ShipUpgrade.UpgradeName, "
   SQL = SQL & "ShipUpgrade.UpgradeDescr, ShipUpgrade.Pay, ShipUpgrade.DriveCore "
   SQL = SQL & "FROM TradeSupplies INNER JOIN (ShipUpgrade RIGHT JOIN (Gear RIGHT JOIN (Crew RIGHT JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON "
   SQL = SQL & "Gear.GearID = SupplyDeck.GearID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID) ON TradeSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE TradeSupplies.PlayerID=" & playerID
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If Nz(rst!leader, 0) = 0 And Nz(rst!DriveCore, 0) = 0 Then
         lstSupplies(1).AddItem Trim(Nz(rst!CrewName) & " " & Nz(rst!CrewDescr) & Nz(rst!GearName) & " " & Nz(rst!GearDescr) & Nz(rst!UpgradeName) & " " & Nz(rst!UpgradeDescr))
         
         lstSupplies(1).ItemData(lstSupplies(1).NewIndex) = rst!CardID
      End If
      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing
End Sub

Private Sub refreshGoodsSupplies(ByVal playerID)
Dim rst As New ADODB.Recordset
Dim SQL, x
   SQL = "SELECT * FROM TradeGoods WHERE PlayerID=" & playerID
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       txt(6) = CStr(rst!cargo)
       txt(7) = CStr(rst!Contraband)
       txt(8) = CStr(rst!Passenger)
       txt(9) = CStr(rst!Fugitive)
       txt(10) = CStr(rst!fuel)
       txt(11) = CStr(rst!parts)
       txt(13) = CStr(rst!pay)
   Else
       txt(6) = "0"
       txt(7) = "0"
       txt(8) = "0"
       txt(9) = "0"
       txt(10) = "0"
       txt(11) = "0"
       txt(13) = "0"
   End If
   For x = 0 To 13
      If Val(txt(x)) > 0 Then
         txt(x).BackColor = 9109503
      Else
         txt(x).BackColor = &H80000005
      End If
   Next x
   
   cmd(1).Enabled = Not rst.EOF And cmd(2).Caption = "Change"
   rst.Close
   Set rst = Nothing
End Sub

Private Sub txt_DblClick(Index As Integer)
   If Not txt(Index).Locked And Index < 6 Then txt(Index).Text = CStr(Val(txt(Index).Text) + 1)
   If Not txt(Index).Locked And Index = 12 Then txt(Index).Text = CStr(Val(txt(Index).Text) + 100)
End Sub
