VERSION 5.00
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmStories 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "View/Edit Story"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStories.frx":0000
   ScaleHeight     =   4980
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SftTree.SftTree sftTree 
      Height          =   1755
      Left            =   60
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2880
      Width           =   11895
      _Version        =   262144
      _ExtentX        =   20981
      _ExtentY        =   3096
      _StockProps     =   237
      ForeColor       =   8833235
      BackColor       =   3353720
      BorderStyle     =   1
      ItemPictureExpanded=   "frmStories.frx":14B5B
      ItemPictureExpandable=   "frmStories.frx":14B77
      ItemPictureLeaf =   "frmStories.frx":14B93
      PlusMinusPictureExpanded=   "frmStories.frx":14BAF
      PlusMinusPictureExpandable=   "frmStories.frx":14BCB
      PlusMinusPictureLeaf=   "frmStories.frx":14BE7
      ButtonPicture   =   "frmStories.frx":14C03
      BeginProperty ColHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty RowHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ItemEditFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColHeaderAppearance=   2
      ButtonStyle     =   0
      Columns         =   15
      ColTitle0       =   "Goal"
      ColBmp0         =   "frmStories.frx":14C1F
      ColWidth1       =   267
      ColTitle1       =   "Instructions"
      ColBmp1         =   "frmStories.frx":14C3B
      ColWidth2       =   133
      ColTitle2       =   "IssueJob"
      ColBmp2         =   "frmStories.frx":14C57
      ColWidth3       =   133
      ColTitle3       =   "CompleteJob"
      ColBmp3         =   "frmStories.frx":14C73
      ColWidth4       =   47
      ColTitle4       =   "Solid"
      ColBmp4         =   "frmStories.frx":14C8F
      ColWidth5       =   47
      ColStyle5       =   10
      ColTitle5       =   "Money"
      ColBmp5         =   "frmStories.frx":14CAB
      ColWidth6       =   30
      ColTitle6       =   "Win"
      ColBmp6         =   "frmStories.frx":14CC7
      ColWidth7       =   33
      ColTitle7       =   "TurnLimit"
      ColBmp7         =   "frmStories.frx":14CE3
      ColWidth8       =   33
      ColStyle8       =   9
      ColTitle8       =   "Fight"
      ColBmp8         =   "frmStories.frx":14CFF
      ColWidth9       =   33
      ColStyle9       =   9
      ColTitle9       =   "Tech"
      ColBmp9         =   "frmStories.frx":14D1B
      ColWidth10      =   33
      ColStyle10      =   9
      ColTitle10      =   "Negot"
      ColBmp10        =   "frmStories.frx":14D37
      ColWidth11      =   33
      ColTitle11      =   "Misbehaves"
      ColBmp11        =   "frmStories.frx":14D53
      ColWidth12      =   33
      ColTitle12      =   "SectorID"
      ColBmp12        =   "frmStories.frx":14D6F
      ColWidth13      =   33
      ColTitle13      =   "Add Crew"
      ColBmp13        =   "frmStories.frx":14D8B
      ColWidth14      =   30
      ColTitle14      =   "Passengers"
      ColBmp14        =   "frmStories.frx":14DA7
      MouseIcon       =   "frmStories.frx":14DC3
      ColHeaderBackColor=   0
      ColHeaderForeColor=   65280
      ForeColor       =   8833235
      BackColor       =   3353720
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmStories.frx":14DDF
      LeftButtonOnly  =   0   'False
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      ColPict0        =   "frmStories.frx":14DFB
      ColPict1        =   "frmStories.frx":14E17
      ColPict2        =   "frmStories.frx":14E33
      ColPict3        =   "frmStories.frx":14E4F
      ColPict4        =   "frmStories.frx":14E6B
      ColPict5        =   "frmStories.frx":14E87
      ColPict6        =   "frmStories.frx":14EA3
      ColPict7        =   "frmStories.frx":14EBF
      ColPict8        =   "frmStories.frx":14EDB
      ColPict9        =   "frmStories.frx":14EF7
      ColPict10       =   "frmStories.frx":14F13
      ColPict11       =   "frmStories.frx":14F2F
      ColPict12       =   "frmStories.frx":14F4B
      ColPict13       =   "frmStories.frx":14F67
      ColPict14       =   "frmStories.frx":14F83
      BackgroundPicture=   "frmStories.frx":14F9F
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Scores"
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
      Index           =   8
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "show the scores for this story"
      Top             =   2340
      Width           =   1035
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
      Index           =   7
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   720
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Add Goal"
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "add a Goal to this story"
      Top             =   1800
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Save"
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "save Story and close"
      Top             =   210
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Delete"
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "delete this Story"
      Top             =   1260
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Priming the Pump"
      Height          =   2775
      Left            =   30
      TabIndex        =   12
      Top             =   60
      Width           =   10605
      Begin VB.CheckBox chkMoveCutter 
         BackColor       =   &H00CBE1ED&
         Caption         =   "move a Reaver Cutter after Fullburns"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         ToolTipText     =   "randomally move a cutter after Fullburns"
         Top             =   2430
         Width           =   3045
      End
      Begin VB.CheckBox chkBounty 
         BackColor       =   &H00CBE1ED&
         Caption         =   "Bounties"
         Height          =   195
         Left            =   5250
         TabIndex        =   34
         ToolTipText     =   "enable Bounty Hunts"
         Top             =   2160
         Width           =   1080
      End
      Begin VB.CheckBox chkHavenStorage 
         BackColor       =   &H00CBE1ED&
         Caption         =   "w/storage"
         Height          =   195
         Left            =   5250
         TabIndex        =   32
         ToolTipText     =   "use Havens for storing stuff"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CheckBox chkUpgrade 
         BackColor       =   &H00CBE1ED&
         Caption         =   "custom Drive or Upgrade"
         Height          =   195
         Left            =   3840
         TabIndex        =   9
         ToolTipText     =   "pick Upgrade or Drive"
         Top             =   2400
         Width           =   2265
      End
      Begin VB.CheckBox chkRandomCrew 
         BackColor       =   &H00CBE1ED&
         Caption         =   "Random Crew"
         Height          =   195
         Left            =   3840
         TabIndex        =   8
         ToolTipText     =   "assigned a random crew instead of selecting them"
         Top             =   2160
         Width           =   1425
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   5340
         TabIndex        =   4
         Text            =   "3"
         ToolTipText     =   "min 1, max 6"
         Top             =   1170
         Width           =   405
      End
      Begin VB.CheckBox chkHavens 
         BackColor       =   &H00CBE1ED&
         Caption         =   "use Havens"
         Height          =   195
         Left            =   3840
         TabIndex        =   7
         ToolTipText     =   "use Havens for starting locations"
         Top             =   1920
         Width           =   1185
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00FF8080&
         Caption         =   "inv"
         Height          =   195
         Index           =   6
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "invert selection"
         Top             =   2500
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00FF8080&
         Caption         =   "moral"
         Height          =   195
         Index           =   5
         Left            =   9180
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "select Moral Crew"
         Top             =   2500
         Width           =   525
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00FF8080&
         Caption         =   "clr"
         Height          =   195
         Index           =   4
         Left            =   10100
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "clear selection"
         Top             =   2500
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00FF8080&
         Caption         =   "wanted"
         Height          =   195
         Index           =   3
         Left            =   8490
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "select WANTED Crew"
         Top             =   2500
         Width           =   675
      End
      Begin VB.ListBox lstCrew 
         BackColor       =   &H00CBE1ED&
         Height          =   2085
         Left            =   8490
         Style           =   1  'Checkbox
         TabIndex        =   24
         Top             =   390
         Width           =   1995
      End
      Begin VB.TextBox txt 
         Height          =   1395
         Index           =   6
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frmStories.frx":14FBB
         Top             =   930
         Width           =   3525
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   5700
         TabIndex        =   6
         Text            =   "1000"
         ToolTipText     =   "maximum cost of Crew"
         Top             =   1560
         Width           =   525
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   5
         Left            =   180
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "Add your Story Title here.."
         Top             =   390
         Width           =   3525
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   5340
         TabIndex        =   5
         Text            =   "0"
         Top             =   1560
         Width           =   315
      End
      Begin VB.ListBox lstContacts 
         BackColor       =   &H00CBE1ED&
         Height          =   2085
         Left            =   6360
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   390
         Width           =   2025
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   5800
         TabIndex        =   3
         Text            =   "2"
         ToolTipText     =   "Parts"
         Top             =   780
         Width           =   405
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   5340
         TabIndex        =   2
         Text            =   "6"
         ToolTipText     =   "Fuel"
         Top             =   780
         Width           =   405
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   5340
         TabIndex        =   10
         Text            =   "3000"
         Top             =   390
         Width           =   885
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Exclude Crew"
         Height          =   225
         Left            =   8490
         TabIndex        =   25
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Story Description"
         Height          =   285
         Index           =   4
         Left            =   210
         TabIndex        =   20
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Story Title"
         Height          =   285
         Index           =   12
         Left            =   210
         TabIndex        =   19
         Top             =   200
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crew && Max $"
         Height          =   285
         Index           =   3
         Left            =   3840
         TabIndex        =   17
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Jobs (discard to 3)"
         Height          =   225
         Left            =   6360
         TabIndex        =   16
         Top             =   195
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No. of Cutters"
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   15
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fuel && Parts"
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   14
         Top             =   780
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cash"
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   13
         Top             =   390
         Width           =   1365
      End
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   60
      TabIndex        =   30
      Top             =   4680
      Width           =   11895
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPop 
         Caption         =   "Open"
         Index           =   0
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Add new"
         Index           =   1
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Delete"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmStories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StoryID As Integer

Private Sub chkHavens_Click()
   chkHavenStorage.Visible = (chkHavens.Value = 1)
   If Not chkHavenStorage.Visible Then chkHavenStorage.Value = 0
End Sub

Private Sub cmd_Click(Index As Integer)
Dim frmGoal As frmGoals, SQL
Dim frmScores As frmScore

   Select Case Index
   Case 0 ' save
      'validation
      If Val(txt(7)) > 6 Or Val(txt(7)) = 0 Then txt(7) = 3
      
      'Save
      SQL = "UPDATE Story SET StoryTitle= " & "'" & SQLFilter(txt(5)) & "',"
      SQL = SQL & " StoryDesc = " & " '" & SQLFilter(txt(6)) & "',"
      SQL = SQL & " StartingCash = " & CStr(Val(txt(0))) & ","
      SQL = SQL & " StartingFuel = " & CStr(Val(txt(1))) & ","
      SQL = SQL & " StartingParts = " & CStr(Val(txt(2))) & ","
      SQL = SQL & " StartingCrew = " & CStr(Val(txt(3))) & ","
      SQL = SQL & " CrewCostLimit = " & CStr(Val(txt(4))) & ","
      SQL = SQL & " NoOfReavers = " & CStr(Val(txt(7))) & ","
      SQL = SQL & " StartingJobs = " & IIf(getList(lstContacts) = "", "NULL", "'" & getList(lstContacts) & "'") & ","
      SQL = SQL & " ExcludeCrew = " & IIf(getList(lstCrew) = "", "NULL", "'" & getList(lstCrew) & "'") & ","
      SQL = SQL & " Havens = " & chkHavens.Value & ","
      SQL = SQL & " HavenStorage = " & chkHavenStorage.Value & ","
      SQL = SQL & " UpgradeDrive = " & chkUpgrade.Value & ","
      SQL = SQL & " RandomCrew = " & chkRandomCrew.Value & ","
      SQL = SQL & " Bounty = " & chkBounty.Value & ","
      SQL = SQL & " MoveCutter = " & chkMoveCutter.Value
      SQL = SQL & " WHERE StoryID = " & StoryID
      DB.Execute SQL
      
      Me.Hide
   
   Case 1 ' add Goal
      Set frmGoal = New frmGoals
      frmGoal.StoryID = StoryID
      frmGoal.Goal = sftTree.ListCount
      frmGoal.chkGoal.Value = 1
      frmGoal.Show 1, Me
      RefreshGoals
   
   Case 2 'delete
      If MsgBox("Are you sure you want to Delete this Story?", vbYesNo + vbQuestion, "Delete Story") = vbNo Then Exit Sub
      DB.Execute "Delete from StoryGoals WHERE StoryID = " & StoryID
      DB.Execute "Delete from Story WHERE StoryID = " & StoryID
      StoryID = 0
      Me.Hide
   
   Case 3 ' outlaws
      SetCrewSel "Wanted"
      
   Case 4 'clear
      SetCrewSel "", True
      
   Case 5
      SetCrewSel "Moral"
         
   Case 6 'invert
      SetCrewSel "", False, True
      
   Case 7 'close
      Me.Hide
   
   Case 8 'scores
      Set frmScores = New frmScore
      frmScores.StoryID = StoryID
      frmScores.Show 1
      
   End Select
End Sub

Private Sub Form_Load()
   LoadCombo lstContacts, "contact", " WHERE ContactID > 0 and ContactID < 10"
   LoadCombo lstCrew, "crew", " Order by CrewName"
   refreshHeader
   RefreshGoals
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Public Sub readonly()
Dim x
   For x = 0 To 6
      cmd(x).Visible = False
   Next x
End Sub

Private Sub refreshHeader()
Dim rst As New ADODB.Recordset
Dim SQL, Index

      SQL = "SELECT * FROM Story WHERE StoryID =" & StoryID

      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      If Not rst.EOF Then
         txt(5) = Nz(rst!StoryTitle)
         txt(6) = Nz(rst!StoryDesc & "")
         txt(0) = Nz(rst!StartingCash)
         txt(1) = Nz(rst!StartingFuel)
         txt(2) = Nz(rst!StartingParts)
         txt(3) = Nz(rst!StartingCrew)
         txt(4) = Nz(rst!CrewCostLimit)
         txt(7) = Nz(rst!NoOfReavers)
         SetList lstContacts, Nz(rst!StartingJobs)
         Label2 = "Exclude Crew (" & CStr(SetList(lstCrew, Nz(rst!ExcludeCrew))) & " selected)"
         lbl(5).Caption = getHighScorer
         chkHavens.Value = rst!Havens
         chkHavenStorage.Visible = (chkHavens.Value = 1)
         chkHavenStorage.Value = rst!HavenStorage
         chkRandomCrew.Value = rst!RandomCrew
         chkBounty.Value = rst!Bounty
         chkMoveCutter.Value = rst!MoveCutter
         chkUpgrade.Value = rst!UpgradeDrive
      Else 'new
         DB.Execute "Insert into Story (StoryID,StoryTitle, Active) VALUES (" & StoryID & ",'add a new story title here..',1)"
      End If
      rst.Close
      Set rst = Nothing


End Sub

Private Function getHighScorer() As String
Dim rst As New ADODB.Recordset
Dim SQL

      SQL = "SELECT * FROM Scores WHERE StoryID =" & StoryID & " ORDER BY turns"

      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         getHighScorer = rst!PlayerName & " did this story in " & rst!Turns & " turns taking " & CStr(DateDiff("n", rst!StartDate, rst!PlayDate)) & " mins on the " & Format(rst!PlayDate, "DD-MMM-YYYY")
      End If
      rst.Close
      Set rst = Nothing


End Function

Private Sub RefreshGoals()
Dim rst As New ADODB.Recordset
Dim SQL, Index
   With sftTree
      .Clear
      SQL = "SELECT * FROM StoryGoals WHERE StoryID =" & StoryID
      SQL = SQL & " ORDER BY Goal"
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         Index = .AddItem(rst!Goal)
         .ItemData(Index) = rst!Goal
         .CellText(Index, 1) = rst!Instructions & ""
         .CellText(Index, 2) = IIf(rst!IssueJobID > 0, varDLookup("JobName", "ContactDeck", "CardID=" & rst!IssueJobID), "")
         .CellText(Index, 3) = IIf(rst!CompleteJobID > 0, varDLookup("JobName", "ContactDeck", "CardID=" & rst!CompleteJobID), "")
         .CellText(Index, 4) = IIf(rst!SolidCount > 0, "Any " & rst!SolidCount, Nz(rst!Solid))
         .CellText(Index, 5) = rst!Cash & ""
         .CellText(Index, 6) = rst!win & ""
         .CellText(Index, 7) = rst!TurnLimit & ""
         .CellText(Index, 8) = rst!fight & ""
         .CellText(Index, 9) = rst!tech & ""
         .CellText(Index, 10) = rst!Negotiate & ""
         .CellText(Index, 11) = rst!Misbehaves & ""
         .CellText(Index, 12) = rst!SectorID & ""
         .CellText(Index, 13) = rst!AddCrew & ""
         .CellText(Index, 14) = rst!Passenger & ""
         
         rst.MoveNext
      Wend
      rst.Close
      Set rst = Nothing
      .RecalcHorizontalExtent
   End With

End Sub

Private Function getList(cbo As Control) As String
Dim x
   With cbo
      For x = 0 To .ListCount - 1
         If .selected(x) Then
            getList = getList & IIf(getList = "", "", ",") & CStr(.ItemData(x))
         End If
      Next x
   End With
   
End Function


Private Function getSelected(cbo As Control) As Integer
Dim x
   With cbo
      For x = 0 To .ListCount - 1
         If .selected(x) Then
            getSelected = getSelected + 1
         End If
      Next x
   End With
   
End Function

Private Function SetList(cbo As Control, ByVal solids As String) As Integer
Dim x, y, a() As String

   If solids = "" Then Exit Function
   With cbo
   
         a = Split(solids, ",")
         For y = LBound(a) To UBound(a)
            For x = 0 To .ListCount - 1
               If .ItemData(x) = Val(a(y)) Then
                  .selected(x) = True
                  SetList = SetList + 1
                  Exit For
               End If
            Next x
         Next y
      
   End With
   
End Function

Private Function SetCrewSel(ByVal perk As String, Optional ByVal clearAll As Boolean = False, Optional ByVal invert As Boolean = False) As Integer
Dim x

   If lstCrew.ListCount = 0 Then Exit Function
   With lstCrew

      For x = 0 To .ListCount - 1
         If clearAll Then
            .selected(x) = False
         ElseIf invert Then
            .selected(x) = Not .selected(x)
         Else
            If varDLookup(perk, "Crew", "CrewID=" & CStr(.ItemData(x))) > 0 Then  '& " AND Leader = 0"
              .selected(x) = True
            End If
         End If
      Next x
      
   End With
   
End Function

Private Sub lstCrew_ItemCheck(Item As Integer)
   Label2 = "Exclude Crew (" & CStr(getSelected(lstCrew)) & " selected)"
End Sub

Private Sub mnuPop_Click(Index As Integer)
Dim frmGoal As frmGoals, x
   x = sftTree.ListIndex

   Select Case Index
   Case 0 'open
      If x > -1 Then
         Set frmGoal = New frmGoals
         frmGoal.StoryID = StoryID
         frmGoal.Goal = sftTree.ItemData(x)
         frmGoal.Show 1, Me
         RefreshGoals
      End If

   Case 1 'add
      Set frmGoal = New frmGoals
      frmGoal.StoryID = StoryID
      frmGoal.Goal = sftTree.ListCount
      frmGoal.Show 1, Me
      RefreshGoals
   
   Case 2 'delete
      DB.Execute "DELETE FROM StoryGoals WHERE StoryID = " & StoryID & " AND Goal = " & sftTree.ItemData(x)
      RefreshGoals
   
   End Select
   
End Sub

Private Sub sftTree_ItemClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
   Select Case AreaType
   Case 9
      If Button = 2 Then
         With sftTree
            mnuPop(2).Enabled = .ListCount > 0
            PopupMenu mnuPopup
        End With
      End If
   End Select
End Sub

Private Sub sftTree_ItemDblClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
Dim frmGoal As frmGoals

   If Button = constSftTreeLeftButton And AreaType = constSftTreeCellText Then
      If Index > -1 Then
         Set frmGoal = New frmGoals
         frmGoal.StoryID = StoryID
         frmGoal.Goal = sftTree.ItemData(Index)
         frmGoal.Show 1, Me
         RefreshGoals
      End If
   End If
End Sub
