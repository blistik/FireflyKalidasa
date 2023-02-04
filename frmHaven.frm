VERSION 5.00
Begin VB.Form frmHaven 
   BackColor       =   &H00CBE1ED&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Haven Transfer"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHaven.frx":0000
   ScaleHeight     =   5940
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   20
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "0"
      Top             =   1110
      Width           =   825
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   18
      Left            =   5370
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   19
      Left            =   5370
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   16
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   17
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   14
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   15
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   25
      Left            =   11490
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   26
      Left            =   11490
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   23
      Left            =   9570
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   24
      Left            =   9570
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   21
      Left            =   7590
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   22
      Left            =   7590
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   27
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "0"
      Top             =   1110
      Width           =   825
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   13
      Left            =   10130
      TabIndex        =   34
      Text            =   "0"
      Top             =   1110
      Width           =   825
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   12
      Left            =   4000
      TabIndex        =   32
      Text            =   "0"
      Top             =   1110
      Width           =   825
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "<<=Transfer=>>"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5400
      Width           =   1485
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
      Height          =   3660
      Index           =   1
      Left            =   6240
      Style           =   1  'Checkbox
      TabIndex        =   28
      Top             =   1500
      Width           =   5685
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   7140
      TabIndex        =   20
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   7140
      TabIndex        =   19
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   9
      Left            =   9120
      TabIndex        =   18
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   8
      Left            =   9120
      TabIndex        =   17
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   11
      Left            =   11040
      TabIndex        =   16
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   10
      Left            =   11040
      TabIndex        =   15
      Text            =   "0"
      Top             =   450
      Width           =   375
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
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1050
      TabIndex        =   4
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   1
      Text            =   "0"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   0
      Text            =   "0"
      Top             =   450
      Width           =   375
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      Height          =   225
      Index           =   13
      Left            =   9600
      TabIndex        =   35
      Top             =   1140
      Width           =   705
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      Height          =   225
      Index           =   12
      Left            =   3400
      TabIndex        =   33
      Top             =   1140
      Width           =   705
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "At your Haven"
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
      Caption         =   "On your Ship"
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
Attribute VB_Name = "frmHaven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public success As Boolean

Private Sub cmd_Click(Index As Integer)

   Select Case Index
   Case 0
      cmd(0).Enabled = False
      playsnd 8
      'ClearTrade
      Me.Hide
   Case 1 'accept
      If validOffer Then
         playsnd 8
         saveOffer
         success = True
         Me.Hide
         PutMsg player.PlayName & " does some Work with a Haven transfer at " & varDLookup("PlanetName", "Planet", "SectorID=" & getPlayerSector(player.ID)), player.ID, Logic!Gamecntr
      End If
      
   End Select
End Sub

Private Sub Form_Load()
   success = False
   refreshSupplies
   refreshGoodsSupplies
   refreshHavenSupplies
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Function validOffer() As Boolean
Dim u, v, x, y, z, crewspacediff As Integer, cargospacediff As Integer

   If Val(txt(0)) > Val(txt(14)) Then
      MessBox "You don't have that much Cargo", "Cargo", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(1)) > Val(txt(15)) Then
      MessBox "You don't have that much Contraband", "Contraband", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(2)) > Val(txt(16)) Then
      MessBox "You don't have that many Passengers", "Passengers", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(3)) > Val(txt(17)) Then
      MessBox "You don't have that many Fugitives", "Fugitives", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(4)) > Val(txt(18)) Then
      MessBox "You don't have that much Fuel", "Fuel", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(5)) > Val(txt(19)) Then
      MessBox "You don't have that many Parts", "Parts", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(12)) > Val(txt(20)) Then
      MessBox "You don't have that much cash", "Cash", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(12)) Mod 100 > 0 Then
      MessBox "Cash must be in increments of $100", "Cash", "Ooops", "", getLeader()
      Exit Function
   End If
   'haven check
   If Val(txt(6)) > Val(txt(21)) Then
      MessBox "You don't have that much Cargo stashed at your Haven", "Haven Cargo", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(7)) > Val(txt(22)) Then
      MessBox "You don't have that much Contraband stashed at your Haven", "Haven Contraband", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(8)) > Val(txt(23)) Then
      MessBox "You don't have that many Passengers at your Haven", "Haven Passengers", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(9)) > Val(txt(24)) Then
      MessBox "You don't have that many Fugitives at your Haven", "Haven Fugitives", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(10)) > Val(txt(25)) Then
      MessBox "You don't have that much Fuel stashed at your Haven", "Haven Fuel", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(11)) > Val(txt(26)) Then
      MessBox "You don't have that many Parts stashed at your Haven", "Haven Parts", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(13)) > Val(txt(27)) Then
      MessBox "You don't have that much cash stashed at your Haven", "Haven Cash", "Ooops", "", getLeader()
      Exit Function
   End If
   If Val(txt(13)) Mod 100 > 0 Then
      MessBox "Haven Cash must be in increments of $100", "Haven Cash", "Ooops", "", getLeader()
      Exit Function
   End If
   x = Val(txt(0))
   x = x + Val(txt(1))
   x = x + Val(txt(2))
   x = x + Val(txt(3))
   x = x + Val(txt(4)) / 2
   x = x + Val(txt(5)) / 2
   
   y = Val(txt(10)) / 2
   y = y + Val(txt(11)) / 2
   y = y + Val(txt(6))
   y = y + Val(txt(7))
   y = y + Val(txt(8))
   y = y + Val(txt(9))
'   If x > y Then  'you are giving more
'      If CargoCapacity(TraderID) - CargoSpaceUsed(TraderID) < (x - y) Then 'no room
'         MsgBox "Other Trader has not enough room for extra Goods", vbExclamation
'         Exit Function
'      End If
   If y > x Then 'you are getting more
      If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < (y - x) Then 'no room
         MessBox "Hey, we ain't got enough room for all that!", "Ship's Cargo space", "Gorram it", "", getLeader()
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
         crewspacediff = crewspacediff - varDLookup("ExtraCrewSpace", "ShipUpgrade", "ShipUpgradeID=" & getShipUpgradeID(.ItemData(z))) 'losing space
         cargospacediff = cargospacediff - varDLookup("ExtraCargoSpace", "ShipUpgrade", "ShipUpgradeID=" & getShipUpgradeID(.ItemData(z))) 'losing space
         cargospacediff = cargospacediff - varDLookup("ExtraStashSpace", "ShipUpgrade", "ShipUpgradeID=" & getShipUpgradeID(.ItemData(z))) 'losing space
      End If
   Next z
   End With
   With lstSupplies(1)
   For z = 0 To .ListCount - 1
      If .selected(z) And getCrewID(.ItemData(z)) > 0 Then
         y = y + 1
      End If
      
      If .selected(z) And getShipUpgradeID(.ItemData(z)) > 0 Then
         v = v + 1
         crewspacediff = crewspacediff + varDLookup("ExtraCrewSpace", "ShipUpgrade", "ShipUpgradeID=" & getShipUpgradeID(.ItemData(z))) 'gaining space
         cargospacediff = cargospacediff + varDLookup("ExtraCargoSpace", "ShipUpgrade", "ShipUpgradeID=" & getShipUpgradeID(.ItemData(z))) 'gaining space
         cargospacediff = cargospacediff + varDLookup("ExtraStashSpace", "ShipUpgrade", "ShipUpgradeID=" & getShipUpgradeID(.ItemData(z))) 'gaining space
      End If
   Next z
   End With

   'check if enough room for crew left
   If CrewCapacity(player.ID) + crewspacediff - getCrewCount(player.ID) < (y - x) Then
      MessBox "Hey, we ain't got enough room for all Crew!", "Ship's Crew space", "Gorram it", "", getLeader()
      Exit Function
   End If

   
   If v > u Then
      If 3 - getShipUpgrades(player.ID) < (v - u) Then
         MessBox "Hey, we ain't got enough room for the extra Ship Upgrades!", "Ship's Upgrades", "Gorram it", "", getLeader()
         Exit Function
      End If
   End If
   
   
   validOffer = True
End Function

'return true if valid offer
Private Function saveOffer() As Boolean
Dim rst As New ADODB.Recordset
Dim SQL, x As Integer, CrewID

'   'transfer goods.
   SQL = "UPDATE Players SET Pay = Pay + " & CStr(Val(txt(13))) & " - " & CStr(Val(txt(12))) & ", Fuel = Fuel + " & CStr(Val(txt(10))) & " - " & CStr(Val(txt(4))) & _
   ", Parts = Parts + " & CStr(Val(txt(11))) & " - " & CStr(Val(txt(5))) & ", Cargo = Cargo + " & CStr(Val(txt(6))) & " - " & CStr(Val(txt(0))) & _
   ", Passenger = Passenger + " & CStr(Val(txt(8))) & " - " & CStr(Val(txt(2))) & ", Contraband = Contraband + " & CStr(Val(txt(7))) & " - " & CStr(Val(txt(1))) & _
   ", Fugitive = Fugitive + " & CStr(Val(txt(9))) & " - " & CStr(Val(txt(3))) & _
   ", HPay = HPay + " & CStr(Val(txt(12))) & " - " & CStr(Val(txt(13))) & ", HFuel = HFuel + " & CStr(Val(txt(4))) & " - " & CStr(Val(txt(10))) & _
   ", HParts = HParts + " & CStr(Val(txt(5))) & " - " & CStr(Val(txt(11))) & ", HCargo = HCargo + " & CStr(Val(txt(0))) & " - " & CStr(Val(txt(6))) & _
   ", HPassenger = HPassenger + " & CStr(Val(txt(2))) & " - " & CStr(Val(txt(8))) & ", HContraband = HContraband + " & CStr(Val(txt(1))) & " - " & CStr(Val(txt(7))) & _
   ", HFugitive = HFugitive + " & CStr(Val(txt(3))) & " - " & CStr(Val(txt(9))) & _
   " WHERE PlayerID = " & player.ID
   DB.Execute SQL
  
   With lstSupplies(0)
   For x = 0 To .ListCount - 1
      If .selected(x) Then
         SQL = "INSERT INTO HavenSupplies (PlayerID, CardID) Values (" & player.ID & ", " & CStr(.ItemData(x)) & ")"
         DB.Execute SQL
         CrewID = getCrewID(.ItemData(x))
         If CrewID > 0 Then DB.Execute "UPDATE PlayerSupplies Set CrewID=0 WHERE CrewID=" & CrewID
         SQL = "DELETE FROM PlayerSupplies WHERE CardID=" & CStr(.ItemData(x))
         DB.Execute SQL
      End If
   Next x
   End With
   
   With lstSupplies(1)
   For x = 0 To .ListCount - 1
      If .selected(x) Then
         SQL = "INSERT INTO PlayerSupplies (PlayerID, CardID) Values (" & player.ID & ", " & CStr(.ItemData(x)) & ")"
         DB.Execute SQL
         SQL = "DELETE FROM HavenSupplies WHERE CardID=" & CStr(.ItemData(x))
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

Private Sub refreshHavenSupplies()
Dim rst As New ADODB.Recordset
Dim SQL
   lstSupplies(1).Clear
   SQL = "SELECT HavenSupplies.CardID, Crew.CrewID, Crew.CrewName, Crew.CrewDescr, Crew.Pay, Crew.Leader, Gear.GearName, Gear.GearDescr, Gear.Pay, ShipUpgrade.UpgradeName, "
   SQL = SQL & "ShipUpgrade.UpgradeDescr, ShipUpgrade.Pay, ShipUpgrade.DriveCore "
   SQL = SQL & "FROM HavenSupplies INNER JOIN (ShipUpgrade RIGHT JOIN (Gear RIGHT JOIN (Crew RIGHT JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON "
   SQL = SQL & "Gear.GearID = SupplyDeck.GearID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID) ON HavenSupplies.CardID = SupplyDeck.CardID "
   SQL = SQL & "WHERE HavenSupplies.PlayerID=" & player.ID

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

' HavenGoods or Players
Private Sub refreshGoodsSupplies()
Dim rst As New ADODB.Recordset
Dim SQL, x
   SQL = "SELECT * FROM Players WHERE PlayerID=" & player.ID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       txt(14) = CStr(rst!cargo)
       txt(15) = CStr(rst!Contraband)
       txt(16) = CStr(rst!Passenger)
       txt(17) = CStr(rst!Fugitive)
       txt(18) = CStr(rst!fuel)
       txt(19) = CStr(rst!parts)
       txt(20) = CStr(rst!pay)
       txt(21) = CStr(rst!hcargo)
       txt(22) = CStr(rst!hContraband)
       txt(23) = CStr(rst!hPassenger)
       txt(24) = CStr(rst!hFugitive)
       txt(25) = CStr(rst!hfuel)
       txt(26) = CStr(rst!hparts)
       txt(27) = CStr(rst!HPay)
   End If
     
   For x = 14 To 27
      If Val(txt(x)) > 0 Then
         txt(x).ForeColor = &HC000&
     ' Else
     '    txt(x).ForeColor = &HC0C0&
      End If
   Next x
      
   rst.Close
   Set rst = Nothing
End Sub


Private Sub lstSupplies_DblClick(Index As Integer)
Dim CrewID, CardID
   CardID = lstSupplies(Index).ItemData(lstSupplies(Index).ListIndex)
   CrewID = Nz(varDLookup("CrewID", "SupplyDeck", "CardID=" & CardID), 0)
   
   If CrewID > 0 Then
      Dim frmCrew As New frmCrewSel
      frmCrew.crewFilter = " WHERE CrewID =" & CrewID
      frmCrew.Show 1
      Set frmCrew = Nothing
   ElseIf Nz(varDLookup("GearID", "SupplyDeck", "CardID=" & CardID), 0) > 0 Then
      Dim frmGear As New frmGearView
      frmGear.gearFilter = " WHERE CardID=" & CardID
      frmGear.Show 1
      Set frmGear = Nothing
   End If
   If Nz(varDLookup("ShipUpgradeID", "SupplyDeck", "CardID=" & CardID), 0) > 0 Then
      Dim frmUpGrd As New frmShipUpgrdView
      frmUpGrd.gearFilter = " WHERE CardID=" & CardID
      frmUpGrd.Show 1
      Set frmUpGrd = Nothing
   End If

End Sub

Private Sub txt_DblClick(Index As Integer)
   If Index = 12 Or Index = 13 Then
      txt(Index).Text = CStr(Val(txt(Index).Text) + 100)
   ElseIf Not txt(Index).Locked Then
      txt(Index).Text = CStr(Val(txt(Index).Text) + 1)
   End If
End Sub
