VERSION 5.00
Begin VB.Form frmKillCrew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crew Member lost in action"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmKillCrew.frx":0000
   ScaleHeight     =   5325
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   2265
      Left            =   180
      Picture         =   "frmKillCrew.frx":BB00
      ScaleHeight     =   2205
      ScaleWidth      =   5355
      TabIndex        =   2
      Top             =   60
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4890
      Width           =   1035
   End
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
      Height          =   2370
      Left            =   180
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   2430
      Width           =   5415
   End
End
Attribute VB_Name = "frmKillCrew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nbrSelect, killed As Integer, extrafilter As String

Private Sub cmd_Click()
Dim x, cnt
   cnt = 0
   For x = 0 To crewList.ListCount - 1
      If crewList.Selected(x) Then
         cnt = cnt + 1
      End If
   Next x
   playsnd 8
   If cnt = nbrSelect Then
      For x = 0 To crewList.ListCount - 1
         If crewList.Selected(x) Then
            
            killed = killed + doKillCrew(player.ID, crewList.ItemData(x))
         
         End If
      Next x
      Me.Hide
   Else
      MessBox "You need to select " & nbrSelect & " crew", "Choose wisely", "Ooops", "", getLeader()
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

   Me.Caption = "Select " & nbrSelect & " Crew that were lost in Action"
  
   LoadCombo crewList, "killcrew", CStr(player.ID) & extrafilter ' AND Crew.Merc = 1
   If nbrSelect > crewList.ListCount Then nbrSelect = crewList.ListCount

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub
