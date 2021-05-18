VERSION 5.00
Begin VB.Form frmSalvaging 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Salvage"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSalvaging.frx":0000
   ScaleHeight     =   1890
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   6
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1110
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   5
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   750
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   4
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   390
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   7
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1440
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00CBE1ED&
      Height          =   780
      Left            =   2280
      Picture         =   "frmSalvaging.frx":4F61A
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   11
      Top             =   90
      Width           =   1125
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   1380
      TabIndex        =   3
      Text            =   "0"
      Top             =   1440
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Text            =   "0"
      Top             =   390
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Text            =   "0"
      Top             =   750
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   1380
      TabIndex        =   2
      Text            =   "0"
      Top             =   1110
      Width           =   345
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Salvage"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   990
      Width           =   1125
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1410
      Width           =   1125
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraband"
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
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1485
      Width           =   1095
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel"
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
      Left            =   240
      TabIndex        =   9
      Top             =   420
      Width           =   1065
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Parts"
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
      Left            =   240
      TabIndex        =   8
      Top             =   765
      Width           =   885
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmSalvaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' mode = 1 Salvage, 2 Load Goods, 3 - Toss x Cargospace Goods overboard, 4 - discard x goods
Option Explicit
Public salvageCount As Variant, mode As Integer

Private Sub cmd_Click(Index As Integer)
Dim total, x
   playsnd 8
   For x = 0 To 3
      total = total + Val(txtDeal(x))
   Next x
   
   If total > salvageCount And mode < 3 Then
      MsgBox "You can only " & IIf(mode = 2, "take", "salvage") & " up to " & salvageCount & " Goods only, what are ya tryin' to pull!", vbExclamation, IIf(mode = 2, "Goods", "Salvage") & " Limits"
      Exit Sub
   End If
   
   If total < salvageCount And mode = 4 Then
      MsgBox "You must discard " & salvageCount & " Goods!", vbExclamation, "Discard Quota"
      Exit Sub
   End If
   
   Select Case Index
   Case 0 'grab
      If mode < 3 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < (Val(txtDeal(0)) + Val(txtDeal(1))) / 2 + Val(txtDeal(2)) + Val(txtDeal(3)) Then
            MsgBox "You don't have enough Cargo Space for that amount of Salvage!", vbExclamation, "Tight for room"
            Exit Sub
         Else
            DB.Execute "UPDATE Players SET Fuel = Fuel + " & CStr(Val(txtDeal(0))) & ", Parts = Parts + " & CStr(Val(txtDeal(1))) & ", Contraband = Contraband + " & CStr(Val(txtDeal(3))) & ", Cargo = Cargo + " & CStr(Val(txtDeal(2))) & " WHERE PlayerID = " & player.ID
            PutMsg player.PlayName & " salvaged " & CStr(Val(txtDeal(0))) & " Fuel, " & CStr(Val(txtDeal(1))) & " Parts, " & CStr(Val(txtDeal(3))) & " Contraband and " & CStr(Val(txtDeal(2))) & " Cargo", player.ID, Logic!Gamecntr
         End If
      ElseIf mode = 3 Then 'dump
         If (Val(txtDeal(0)) + Val(txtDeal(1))) / 2 + Val(txtDeal(2)) + Val(txtDeal(3)) >= salvageCount Then
            DB.Execute "UPDATE Players SET Fuel = Fuel - " & CStr(Val(txtDeal(0))) & ", Parts = Parts - " & CStr(Val(txtDeal(1))) & ", Contraband = Contraband - " & CStr(Val(txtDeal(3))) & ", Cargo = Cargo - " & CStr(Val(txtDeal(2))) & " WHERE PlayerID = " & player.ID
            PutMsg player.PlayName & " dumped " & CStr(Val(txtDeal(0))) & " Fuel, " & CStr(Val(txtDeal(1))) & " Parts, " & CStr(Val(txtDeal(3))) & " Contraband and " & CStr(Val(txtDeal(2))) & " Cargo overboard", player.ID, Logic!Gamecntr
         Else
            MsgBox "That's only " & CStr((Val(txtDeal(0)) + Val(txtDeal(1))) / 2 + Val(txtDeal(2)) + Val(txtDeal(3))) & " space, need " & CStr(salvageCount), vbExclamation, "Discard goods"
            Exit Sub
         End If
      ElseIf mode = 4 Then
         DB.Execute "UPDATE Players SET Fuel = Fuel - " & CStr(Val(txtDeal(0))) & ", Parts = Parts - " & CStr(Val(txtDeal(1))) & ", Contraband = Contraband - " & CStr(Val(txtDeal(3))) & ", Cargo = Cargo - " & CStr(Val(txtDeal(2))) & " WHERE PlayerID = " & player.ID
         PutMsg player.PlayName & " discarded " & CStr(Val(txtDeal(0))) & " Fuel, " & CStr(Val(txtDeal(1))) & " Parts, " & CStr(Val(txtDeal(3))) & " Contraband and " & CStr(Val(txtDeal(2))) & " Cargo", player.ID, Logic!Gamecntr
      End If
   Case 1 'nope
   
   End Select
   Me.Hide


End Sub

Private Sub Form_Load()
   Select Case mode
   Case 1
      Me.Caption = "Load up to " & salvageCount & " Salvage. Spare storage: " & CStr(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
      cmd(0).Caption = "Salvage"
   Case 2
      Me.Caption = "Load up to " & salvageCount & " Goods. Spare storage: " & CStr(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
      cmd(0).Caption = "Load Goods"
   Case 3
      Me.Caption = "Dump at least " & salvageCount & " Cargospace"
      cmd(0).Caption = "Dump Goods"
      cmd(1).Visible = False

   Case 4
      Me.Caption = "Discard at least " & salvageCount & " Goods"
      cmd(0).Caption = "Discard"
      cmd(1).Visible = False
   End Select
   Label1 = "Qty   Hold"
   initHeld
End Sub

Private Sub initHeld()
Dim rst As New ADODB.Recordset
Dim SQL, x
   For x = 4 To 7
      txtDeal(x).Visible = True
   Next x

   SQL = "SELECT * FROM Players WHERE PlayerID = " & player.ID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      txtDeal(4) = CStr(rst!fuel)
      txtDeal(5) = CStr(rst!parts)
      txtDeal(6) = CStr(rst!cargo)
      txtDeal(7) = CStr(rst!Contraband)
   End If
   rst.Close
   Set rst = Nothing
End Sub
