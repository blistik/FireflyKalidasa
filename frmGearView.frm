VERSION 5.00
Begin VB.Form frmGearView 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gear Viewer"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGearView.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboGear 
      BackColor       =   &H00CBE1ED&
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2505
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Close"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   510
      Width           =   795
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BackColor       =   &H00CBE1ED&
      Height          =   1200
      Left            =   2670
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   0
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   930
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   465
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   5715
      WordWrap        =   -1  'True
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
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   9
      Top             =   4290
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1260
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
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
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
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
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Origin"
      Top             =   2250
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
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Origin"
      Top             =   2580
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
      Index           =   9
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Origin"
      Top             =   2910
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmGearView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gearFilter As String

Private Sub cboGear_Click()
   If cboGear.ListIndex = -1 Then Exit Sub
   
   refreshGear GetCombo(cboGear)
End Sub

Private Sub cmd_Click()
   playsnd 8
   Me.Hide
End Sub

Private Sub Form_Load()
   LoadCombo cboGear, "gear", gearFilter
   If cboGear.ListCount > 0 Then
      cboGear.ListIndex = 0
   End If

End Sub

Private Sub lbl_DblClick(Index As Integer)
   If gearFilter = " Order By GearName" Then
      If lbl(Index).Tag <> "" Then
         LoadCombo cboGear, "gear", " WHERE " & lbl(Index).Tag
      Else
         LoadCombo cboGear, "gear", gearFilter
      End If
      If cboGear.ListCount > 0 Then
         cboGear.ListIndex = 0
      End If
   End If
End Sub

Private Sub refreshGear(ByVal CardID)
Dim rst As New ADODB.Recordset, SQL
   SQL = "SELECT Gear.*, SupplyDeck.CardID, SupplyDeck.SupplyID, Supply.Colour, Supply.SupplyName, PlayerSupplies.PlayerID, Players.Name "
   SQL = SQL & "FROM Players RIGHT JOIN (PlayerSupplies RIGHT JOIN (Supply RIGHT JOIN (Gear LEFT JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID) ON Supply.SupplyID = SupplyDeck.SupplyID) ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Players.PlayerID = PlayerSupplies.PlayerID "
   SQL = SQL & "WHERE SupplyDeck.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      lbl(0) = rst!GearDescr
      lbl(0).ToolTipText = rst!GearDescr
      lbl(1) = rst!GearDescr
      lbl(2).Visible = (Nz(rst!Name) <> "")
      If Nz(rst!Name) <> "" Then
         lbl(2) = "owned by: " & rst!Name
      End If
      lbl(3) = ""
      lbl(4) = Trim(IIf(rst!fight >= 1, rst!fight & " Fight  ", "") & IIf(rst!tech >= 1, rst!tech & " Tech  ", "") & IIf(rst!Negotiate >= 1, rst!Negotiate & " Negotiate", ""))
      lbl(4).Visible = (lbl(4) <> "")
      lbl(5) = Nz(rst!KeyWords)
      If lbl(5) <> "" Then lbl(5).Tag = "Gear.KeyWords = '" & Nz(rst!KeyWords) & "'"
      
      If IsNull(rst!KeyWords) Then
         lbl(5).Visible = False
      Else
         lbl(5).Visible = True
         lbl(5).BackColor = 12574908
      End If
      
      lbl(6) = "$" & rst!pay
      lbl(6).Visible = True
      
      lbl(7) = rst!SupplyName
      lbl(7).BackColor = rst!Colour
      lbl(7).Tag = "SupplyDeck.SupplyID=" & rst!SupplyID
      
      lbl(8) = "CardID: " & rst!CardID & "    GearID: " & rst!GearID
      
      lbl(9) = ""
      
      If Not IsNull(rst!Picture) Then
         Set pic.Picture = LoadPicture(App.Path & "\pictures\" & rst!Picture)
      Else
         Set pic.Picture = LoadPicture()
      End If
   End If
   rst.Close

End Sub

