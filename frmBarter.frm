VERSION 5.00
Begin VB.Form frmBarter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Deals on the Go / Shrewd Bartering"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBarter.frx":0000
   ScaleHeight     =   4065
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   3
      Left            =   2790
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2790
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   4
      Left            =   2790
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3150
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   5
      Left            =   2790
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3510
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      Height          =   2235
      Left            =   240
      Picture         =   "frmBarter.frx":4F61A
      ScaleHeight     =   2175
      ScaleWidth      =   4485
      TabIndex        =   11
      Top             =   150
      Width           =   4545
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "No thanks"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3510
      Width           =   1125
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Buy"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3140
      Width           =   1125
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   2370
      TabIndex        =   2
      Text            =   "0"
      Top             =   3510
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   2370
      TabIndex        =   1
      Text            =   "0"
      Top             =   3150
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   2370
      TabIndex        =   0
      Text            =   "0"
      Top             =   2790
      Width           =   345
   End
   Begin VB.Label lblCost 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$0"
      Height          =   260
      Left            =   3705
      TabIndex        =   17
      ToolTipText     =   "Money in hand"
      Top             =   2820
      Width           =   660
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
      Height          =   315
      Left            =   3270
      TabIndex        =   16
      Top             =   2850
      Width           =   465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Goods for purchase"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   15
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      Height          =   315
      Left            =   3270
      TabIndex        =   10
      Top             =   2550
      Width           =   465
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$0"
      Height          =   260
      Left            =   3710
      TabIndex        =   9
      ToolTipText     =   "Money in hand"
      Top             =   2520
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty   Hold"
      Height          =   315
      Left            =   2430
      TabIndex        =   6
      Top             =   2550
      Width           =   765
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraband $300ea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   540
      TabIndex        =   5
      Top             =   3540
      Width           =   1845
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Parts $300ea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   540
      TabIndex        =   4
      Top             =   3165
      Width           =   1485
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel $200ea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   540
      TabIndex        =   3
      Top             =   2820
      Width           =   1485
   End
End
Attribute VB_Name = "frmBarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public trader As Integer
Private fuel As Integer, parts As Integer, contra As Integer

Private Sub cmd_Click(Index As Integer)
Dim total As Long
   playsnd 8
   If Val(txtDeal(2)) > 3 Then
      MessBox "You only can purchase up to 3 Contraband, what are ya tryin' to pull!", "Bartering", "Sorry", "", getLeader()
      txtDeal(2) = "3"
      Exit Sub
   End If
   
   Select Case Index
   Case 0 'buy
      total = calcCost
      If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < (Val(txtDeal(0)) + Val(txtDeal(1))) / 2 + Val(txtDeal(2)) Then
         MessBox "You don't have enough Cargo Space for that deal", "Tight for room", "Ooops", "", 0, 0, 6
         Exit Sub
      ElseIf total > getMoney(player.ID) Then
         MessBox "You cannot afford that", "Short Dealin'", "Ooops", "", getLeader()
         Exit Sub
      Else
         DB.Execute "UPDATE Players SET Fuel = Fuel + " & CStr(Val(txtDeal(0))) & ", Parts = Parts + " & CStr(Val(txtDeal(1))) & ", Contraband = Contraband + " & CStr(Val(txtDeal(2))) & ", Pay = Pay - " & CStr(total) & " WHERE PlayerID = " & player.ID
         PutMsg player.PlayName & " bought " & CStr(Val(txtDeal(0))) & " Fuel, " & CStr(Val(txtDeal(1))) & " Parts and " & CStr(Val(txtDeal(2))) & " Contraband from a Trader for $" & CStr(total), player.ID, Logic!Gamecntr
      End If
   Case 1 'nope
   
   End Select
   
   Me.Hide

End Sub

Private Sub Form_Load()
   Me.Caption = "Deals on the Go / Shrewd Bartering. Spare storage: " & CStr(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
   fuel = 200
   parts = 300
   If trader = 1 Then
      contra = 400
      lbl(2).Caption = "Contraband $400ea"
   Else
      contra = 300
   End If
   lblMoney = "$" & varDLookup("Pay", "Players", "PlayerID=" & player.ID)
   lblMoney.ForeColor = 16777215
   lblMoney.BackColor = 8388736
   initHeld

End Sub


Private Sub initHeld()
Dim rst As New ADODB.Recordset
Dim SQL, x
   For x = 3 To 5
      txtDeal(x).Visible = True
   Next x

   SQL = "SELECT * FROM Players WHERE PlayerID = " & player.ID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      txtDeal(3) = CStr(rst!fuel)
      txtDeal(4) = CStr(rst!parts)
      txtDeal(5) = CStr(rst!Contraband)
   End If
   rst.Close
   Set rst = Nothing
End Sub

Private Sub txtDeal_Change(Index As Integer)
   calcCost
End Sub

Private Sub txtDeal_DblClick(Index As Integer)
   txtDeal(Index).Text = CStr(Val(txtDeal(Index).Text) + 1)
   calcCost
End Sub

Private Function calcCost() As Integer
   calcCost = Val(txtDeal(0)) * fuel + Val(txtDeal(1)) * parts + Val(txtDeal(2)) * contra
   lblCost.Caption = CStr(calcCost)
End Function
