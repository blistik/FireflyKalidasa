VERSION 5.00
Begin VB.Form frmStash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Arrange the Stash before Hold Decompression"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStash.frx":0000
   ScaleHeight     =   4350
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   225
      Index           =   3
      Left            =   2640
      TabIndex        =   26
      ToolTipText     =   "minus"
      Top             =   3930
      Width           =   285
   End
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   225
      Index           =   2
      Left            =   2640
      TabIndex        =   25
      ToolTipText     =   "minus"
      Top             =   3600
      Width           =   285
   End
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   225
      Index           =   1
      Left            =   2640
      TabIndex        =   24
      ToolTipText     =   "minus"
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdminus 
      Caption         =   "-"
      Height          =   225
      Index           =   0
      Left            =   2640
      TabIndex        =   23
      ToolTipText     =   "minus"
      Top             =   2880
      Width           =   285
   End
   Begin VB.CommandButton cmdplus 
      Caption         =   "+"
      Height          =   225
      Index           =   3
      Left            =   1560
      TabIndex        =   22
      ToolTipText     =   "plus"
      Top             =   3930
      Width           =   285
   End
   Begin VB.CommandButton cmdplus 
      Caption         =   "+"
      Height          =   225
      Index           =   2
      Left            =   1560
      TabIndex        =   21
      ToolTipText     =   "plus"
      Top             =   3600
      Width           =   285
   End
   Begin VB.CommandButton cmdplus 
      Caption         =   "+"
      Height          =   225
      Index           =   1
      Left            =   1560
      TabIndex        =   20
      ToolTipText     =   "plus"
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdplus 
      Caption         =   "+"
      Height          =   225
      Index           =   0
      Left            =   1560
      TabIndex        =   19
      ToolTipText     =   "plus"
      Top             =   2880
      Width           =   285
   End
   Begin VB.TextBox txtStash 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   1
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3180
      Width           =   440
   End
   Begin VB.TextBox txtStash 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   0
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2820
      Width           =   440
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Done"
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
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   1125
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   1860
      TabIndex        =   8
      Text            =   "0"
      Top             =   3570
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1860
      TabIndex        =   7
      Text            =   "0"
      Top             =   3210
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1860
      TabIndex        =   6
      Text            =   "0"
      Top             =   2850
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   1860
      TabIndex        =   5
      Text            =   "0"
      Top             =   3900
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   7
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3900
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   4
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2850
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   5
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3210
      Width           =   345
   End
   Begin VB.TextBox txtDeal 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Height          =   285
      Index           =   6
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3570
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      Height          =   2235
      Left            =   270
      Picture         =   "frmStash.frx":4F61A
      ScaleHeight     =   2175
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   120
      Width           =   4545
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stash Left"
      Height          =   195
      Left            =   3150
      TabIndex        =   18
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stash Size"
      Height          =   195
      Left            =   3150
      TabIndex        =   17
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stash/Total"
      Height          =   315
      Left            =   1860
      TabIndex        =   13
      Top             =   2610
      Width           =   1125
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   270
      TabIndex        =   12
      Top             =   3570
      Width           =   1035
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Parts"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   270
      TabIndex        =   11
      Top             =   3195
      Width           =   885
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   270
      TabIndex        =   10
      Top             =   2850
      Width           =   1065
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraband"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   270
      TabIndex        =   9
      Top             =   3915
      Width           =   1245
   End
End
Attribute VB_Name = "frmStash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
Dim SQL As String
   'validate
   If Val(txtDeal(0)) > Val(txtDeal(4)) Then
      MessBox "you don't have that much Fuel", "Quantity Issue?", "Ooops"
      Exit Sub
   End If
   If Val(txtDeal(1)) > Val(txtDeal(5)) Then
      MessBox "you don't have that many Parts", "Quantity Issue?", "Ooops"
      Exit Sub
   End If
   If Val(txtDeal(2)) > Val(txtDeal(6)) Then
      MessBox "you don't have that much Cargo", "Quantity Issue?", "Ooops"
      Exit Sub
   End If
   If Val(txtDeal(3)) > Val(txtDeal(7)) Then
      MessBox "you don't have that much Contraband", "Quantity Issue?", "Ooops"
      Exit Sub
   End If
   If StashCapacity(player.ID) < Val(txtDeal(0)) / 2 + Val(txtDeal(1)) / 2 + Val(txtDeal(2)) + Val(txtDeal(3)) Then
      MessBox "you don't have that much Stash Capacity", "Quantity Issue?", "Ooops"
      Exit Sub
   End If
   
   SQL = "UPDATE Players Set Fuel = " & Val(txtDeal(0)) & ", "
   SQL = SQL & "Parts = " & Val(txtDeal(1)) & ", "
   SQL = SQL & "Cargo = " & Val(txtDeal(2)) & ", "
   SQL = SQL & "Contraband = " & Val(txtDeal(3)) & " "
   SQL = SQL & "WHERE PlayerID = " & player.ID
   
   DB.Execute SQL
   
   Me.Hide
   
End Sub

Private Sub cmdminus_Click(Index As Integer)
   If Val(txtDeal(Index)) > 0 Then
      txtDeal(Index) = CStr(Val(txtDeal(Index)) - 1)
   End If
End Sub

Private Sub cmdplus_Click(Index As Integer)
   If Val(txtDeal(Index)) < Val(txtDeal(Index + 4)) Then
      txtDeal(Index) = CStr(Val(txtDeal(Index)) + 1)
   End If

End Sub

Private Sub Form_Load()
   initHeld

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub initHeld()
Dim rst As New ADODB.Recordset
Dim SQL

   SQL = "SELECT * FROM Players WHERE PlayerID = " & player.ID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      txtDeal(4) = CStr(rst!fuel)
      txtDeal(5) = CStr(rst!parts)
      txtDeal(6) = CStr(rst!cargo)
      txtDeal(7) = CStr(rst!Contraband)
   End If
   rst.Close
   txtStash(0) = StashCapacity(player.ID)
   
   If StashCapacity(player.ID) >= Val(txtDeal(4)) / 2 + Val(txtDeal(5)) / 2 + Val(txtDeal(6)) + Val(txtDeal(7)) Then
      txtDeal(0) = txtDeal(4)
      txtDeal(1) = txtDeal(5)
      txtDeal(2) = txtDeal(6)
      txtDeal(3) = txtDeal(7)
      refreshStashLeft
   Else
      txtStash(1) = txtStash(0)
   End If
   
   Set rst = Nothing
End Sub

Private Sub refreshStashLeft()
    txtStash(1) = CStr(StashCapacity(player.ID) - (Val(txtDeal(0)) / 2 + Val(txtDeal(1)) / 2 + Val(txtDeal(2)) + Val(txtDeal(3))))
    If Val(txtStash(1)) > 0 Then
      txtStash(1).BackColor = QBColor(14)
    ElseIf Val(txtStash(1)) = 0 Then
      txtStash(1).BackColor = QBColor(10)
    Else
      txtStash(1).BackColor = QBColor(12)
    End If
End Sub

Private Sub txtDeal_Change(Index As Integer)
   If Index < 4 And Me.Visible Then refreshStashLeft
End Sub
