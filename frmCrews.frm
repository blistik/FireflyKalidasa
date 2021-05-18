VERSION 5.00
Begin VB.Form frmCrewSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crew Selector"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrews.frx":0000
   ScaleHeight     =   5220
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BackColor       =   &H00CBE1ED&
      Height          =   1200
      Left            =   2520
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   8
      Top             =   120
      Width           =   1305
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "select"
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
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   510
      Width           =   795
   End
   Begin VB.ComboBox cboCrew 
      BackColor       =   &H00CBE1ED&
      Height          =   315
      Left            =   150
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2265
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   9
      Left            =   150
      TabIndex        =   12
      ToolTipText     =   "Origin"
      Top             =   3570
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   8
      Left            =   150
      TabIndex        =   11
      ToolTipText     =   "Origin"
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   7
      Left            =   150
      TabIndex        =   10
      ToolTipText     =   "Origin"
      Top             =   2910
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   6
      Left            =   150
      TabIndex        =   9
      Top             =   2580
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   150
      TabIndex        =   7
      Top             =   2250
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   150
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   5
      Top             =   1590
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   2
      Left            =   150
      TabIndex        =   4
      Top             =   1260
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   465
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   4680
      Width           =   5715
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   930
      Width           =   2295
   End
End
Attribute VB_Name = "frmCrewSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public crewFilter As String

Private Sub cboCrew_Click()

   If cboCrew.ListIndex = -1 Then Exit Sub
   
   refreshCrew GetCombo(cboCrew)
      
End Sub

Private Sub cmd_Click()
   playsnd 8
   If cboCrew.ListIndex = -1 Then Exit Sub
      
   If player.PlayName <> "" And actionSeq = 0 Then
      PutMsg player.PlayName & " has chosen " & cboCrew.Text, player.ID
   End If
   
   Me.Hide
End Sub

Private Sub Form_Load()
   LoadCombo cboCrew, "crew", crewFilter
   If cboCrew.ListCount > 0 Then
      cboCrew.ListIndex = 0
   End If

End Sub

Private Sub lbl_DblClick(Index As Integer)
   If crewFilter = " Order By CrewName" Then
      If lbl(Index).Tag <> "" Then
         LoadCombo cboCrew, "crew", " INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID WHERE " & lbl(Index).Tag
      Else
         LoadCombo cboCrew, "crew", crewFilter
      End If
      If cboCrew.ListCount > 0 Then
         cboCrew.ListIndex = 0
      End If
   End If
End Sub

Private Sub refreshCrew(ByVal CrewID)
Dim rst As New ADODB.Recordset, SQL
   SQL = "SELECT Crew.*, Perk.PerkDescription, SupplyDeck.CardID, SupplyDeck.SupplyID, Supply.Colour, Supply.SupplyName FROM Supply RIGHT JOIN ((Perk RIGHT JOIN Crew ON Perk.PerkID = Crew.PerkID) LEFT JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON Supply.SupplyID = SupplyDeck.SupplyID WHERE Crew.CrewID=" & CrewID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      lbl(0) = rst!CrewDescr
      lbl(1) = rst!PerkDescription
      lbl(2) = IIf(rst!Mechanic = 1, "Mechanic  ", "") & IIf(rst!Pilot = 1, "Pilot  ", "") & IIf(rst!Companion = 1, "Companion  ", "") & _
               IIf(rst!Merc = 1, "Merc  ", "") & IIf(rst!Soldier = 1, "Soldier  ", "") & IIf(rst!HillFolk = 1, "HillFolk  ", "") & _
               IIf(rst!Grifter = 1, "Grifter ", "") & IIf(rst!Medic = 1, "Medic ", "")
               
      lbl(2).Tag = IIf(rst!Mechanic = 1, "Crew.Mechanic = 1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Pilot = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Pilot=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Companion = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Companion=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Merc = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Merc=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Soldier = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Soldier=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!HillFolk = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.HillFolk=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Grifter = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Grifter=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Medic = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Medic=1", "")
      
      lbl(3) = IIf(rst!Moral = 1, "Moral    ", "") & IIf(rst!Wanted > 0, "Wanted ", "")
      If rst!Moral = 1 Then
         lbl(3).BackColor = &HC0FFC0
      ElseIf rst!Wanted > 0 Then
         lbl(3).BackColor = &HC0C0FF
      Else
         lbl(3).BackColor = 13361645
      End If
      lbl(4) = Trim(IIf(rst!fight >= 1, rst!fight & " Fight  ", "") & IIf(rst!tech >= 1, rst!tech & " Tech  ", "") & IIf(rst!Negotiate >= 1, rst!Negotiate & " Negotiate", ""))
      lbl(5) = Nz(rst!KeyWords)
      If IsNull(rst!KeyWords) Then
         lbl(5).Visible = False
      Else
         lbl(5).Visible = True
         lbl(5).BackColor = 12574908
      End If
      
      lbl(6) = IIf(rst!leader = 1, "LEADER", "") & IIf(rst!pay >= 1, "$" & rst!pay & " hire/job", "")
      If rst!leader = 1 Then
         lbl(6).Visible = True
         lbl(6).BackColor = &HC0C0&
      ElseIf rst!pay >= 1 Then
         lbl(6).Visible = True
         lbl(6).BackColor = &HFFFF00
      Else
         lbl(6).Visible = False
      End If
      
      lbl(7) = rst!SupplyName
      lbl(7).BackColor = rst!Colour
      lbl(7).Tag = "SupplyDeck.SupplyID=" & rst!SupplyID
      
      lbl(8) = "CardID: " & rst!CardID & "    CrewID: " & rst!CrewID
      
      lbl(9) = IIf(rst!Disgruntled > 0, "Disgruntled ", "")
      If rst!Disgruntled > 0 Then
         lbl(9).Visible = True
         lbl(9).BackColor = &HC0C0FF
      Else
         lbl(9).Visible = False
      End If
      
      If Not IsNull(rst!Picture) Then
         Set pic.Picture = LoadPicture(App.Path & "\pictures\" & rst!Picture)
      Else
         Set pic.Picture = LoadPicture()
      End If
   End If


End Sub


