VERSION 5.00
Begin VB.Form frmShipUpgd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Grab a Free Ship Upgrade"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmShipUpgd.frx":0000
   ScaleHeight     =   1260
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   5903
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   780
      Width           =   1035
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   330
      Width           =   12615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick a ShipUpgrade from the discard piles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmShipUpgd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' discardMode = 1 discard one upgrd, discardMode = 2 buy an upgrd, discardMode = 3 Cry Baby in discard for free, discardMode = 4 take 1 discard for free
' discardMode = 5 take 1 DriveCore discard for free, discardMode = 6 take any upgrade for free
Option Explicit
Public CardID As Integer, discardMode As Integer

Private Sub cmd_Click()
Dim ShipUpgradeID, cost, pay
   playsnd 8
   CardID = GetCombo(cbo)
   If CardID = -1 Then
      If cbo.ListCount = 0 Then 'nothing to pick
         Me.Hide
         Exit Sub
      Else
         MessBox "you have to Pick something", "Pick something", "Ooops", "", getLeader()
         Exit Sub
      End If
   End If
   ShipUpgradeID = getShipUpgradeID(CardID)
   cost = Nz(Val(varDLookup("Pay", "ShipUpgrade", "ShipUpgradeID=" & ShipUpgradeID)), 0) / 2
   pay = varDLookup("Pay", "Players", "PlayerID=" & player.ID)
      
   If CardID = -1 Then
      Exit Sub
   ElseIf cost > pay And discardMode = 2 Then
      If MessBox("You cannot afford that!" & vbNewLine & "Do you want to pass?", "No Dough", "Yes", "No", getLeader()) = 0 Then
         CardID = 0
         Me.Hide
      Else
         Exit Sub
      End If
   Else
      If discardMode = 1 Then
         doRemoveUpgrade player.ID, CardID
         DB.Execute "UPDATE SupplyDeck SET Seq =5 WHERE CardID = " & CardID
         'remv the card from the players deck
         DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID = " & player.ID & " AND CardID =" & CardID
      
      Else
         'if getting a Drive Core, swap out the existing one
         If isDriveCore(CardID) Then
            removeDriveCore player.ID
         End If
         DB.Execute "UPDATE SupplyDeck SET Seq =" & player.ID & " WHERE CardID = " & CardID
         'add the card to the players deck
         DB.Execute "INSERT INTO PlayerSupplies (PlayerID, CardID) VALUES (" & player.ID & ", " & CardID & ")"
         If discardMode = 2 Then 'pay for it
            DB.Execute "UPDATE Players SET Pay = Pay - " & cost & " WHERE PlayerID=" & player.ID
            If actionSeq = ASBuySelect Or actionSeq = ASBuySelDiscard Then
               actionSeq = ASselect
               Main.showBuys False, "local"
               Main.frmBuy.RefreshBuys
            End If
            frmAction.buyIsDone
            frmAction.cmd(2).Enabled = False
            frmAction.lblMoney = CStr(pay - cost)
         End If
      End If
      Me.Hide
   End If
End Sub

Private Sub Form_Load()
   If discardMode = 1 Then
      Me.Caption = "Discard a Ship upgrade"
      Label1 = "Pick a Ship Upgrade to discard"
   ElseIf discardMode = 2 Then
      Me.Caption = "Buy a Ship upgrade"
      Label1 = "Pick a Ship Upgrade to buy at Half Listed Price"
   End If
   LoadCombo cbo, "shipupgd", IIf(discardMode = 1, "=" & CStr(player.ID) & " AND ShipUpgrade.DriveCore<>1", IIf(discardMode = 3, "=5 AND ShipUpgrade.ShipUpgradeID = 1", IIf(discardMode = 5, "=5 AND ShipUpgrade.DriveCore = 1", IIf(discardMode = 6, ">4", IIf(discardMode = 4, "=5 AND ShipUpgrade.DriveCore<>1", "=5")))))

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
      'Me.Hide
   End If
End Sub
