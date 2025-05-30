VERSION 5.00
Begin VB.Form frmGoals 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Story Goal - Criteria to achieve"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGoals.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00CBE1ED&
      Caption         =   "post Goal variations"
      Height          =   855
      Left            =   120
      TabIndex        =   38
      Top             =   4290
      Width           =   8415
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   3
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   460
         Width           =   2145
      End
      Begin VB.CheckBox chkWarrant 
         BackColor       =   &H00CBE1ED&
         Caption         =   "receive Warrant"
         Height          =   255
         Left            =   4290
         TabIndex        =   43
         Top             =   240
         Width           =   1680
      End
      Begin VB.CheckBox chkClearReaver 
         BackColor       =   &H00CBE1ED&
         Caption         =   "Clear Reaver Alerts"
         Height          =   255
         Left            =   2370
         TabIndex        =   42
         ToolTipText     =   "limit job lists to only those specifically made for goals"
         Top             =   510
         Width           =   1815
      End
      Begin VB.CheckBox chkClearAlliance 
         BackColor       =   &H00CBE1ED&
         Caption         =   "Clear Alliance Alerts"
         Height          =   255
         Left            =   2370
         TabIndex        =   41
         ToolTipText     =   "limit job lists to only those specifically made for goals"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   1590
         TabIndex        =   39
         Text            =   "0"
         ToolTipText     =   "-ve or +ve change value"
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "become Solid with"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   16
         Left            =   6120
         TabIndex        =   44
         Top             =   260
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "change in Cutters"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   15
         Left            =   90
         TabIndex        =   40
         Top             =   260
         Width           =   1365
      End
   End
   Begin VB.CheckBox chkUnfinished 
      BackColor       =   &H00CBE1ED&
      Caption         =   "no un- finished jobs"
      Height          =   435
      Left            =   4860
      TabIndex        =   37
      Top             =   1480
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   9
      Left            =   4860
      TabIndex        =   35
      Text            =   "0"
      ToolTipText     =   "the number of Bounties delivered to meet the Goal"
      Top             =   1230
      Width           =   885
   End
   Begin VB.ListBox lstCrew 
      BackColor       =   &H00CBE1ED&
      Height          =   1635
      Left            =   6090
      Style           =   1  'Checkbox
      TabIndex        =   34
      Top             =   270
      Width           =   1995
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "view"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2730
      Width           =   645
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "view"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2160
      Width           =   645
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   4860
      TabIndex        =   13
      Text            =   "0"
      ToolTipText     =   "add Passengers on Goal completion"
      Top             =   750
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   4860
      TabIndex        =   12
      Text            =   "0"
      ToolTipText     =   "Any Solids to this count"
      Top             =   290
      Width           =   885
   End
   Begin VB.CheckBox chkWin 
      BackColor       =   &H00CBE1ED&
      Caption         =   "WIN"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   90
      Width           =   645
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   2
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Planetary Sector to be at for Goal to be met"
      Top             =   2730
      Width           =   1245
   End
   Begin VB.CheckBox chkGoal 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Goal Specific"
      Height          =   255
      Left            =   7050
      TabIndex        =   17
      ToolTipText     =   "limit job lists to only those specifically made for goals"
      Top             =   2490
      Width           =   1305
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
      Index           =   2
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1410
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
      Index           =   1
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
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
      Index           =   0
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   510
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   1
      Left            =   2670
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2730
      Width           =   5865
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   0
      Left            =   2670
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2160
      Width           =   5865
   End
   Begin VB.TextBox txt 
      Height          =   825
      Index           =   8
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3360
      Width           =   8415
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   1380
      TabIndex        =   6
      Text            =   "0"
      Top             =   2370
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   1380
      TabIndex        =   5
      Text            =   "0"
      Top             =   1980
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   1380
      TabIndex        =   4
      Text            =   "0"
      Top             =   1590
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   1380
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "proceed past this many"
      Top             =   1230
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1380
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1380
      TabIndex        =   1
      Text            =   "0"
      Top             =   450
      Width           =   885
   End
   Begin VB.ListBox lstContacts 
      BackColor       =   &H00CBE1ED&
      Height          =   1635
      Left            =   2520
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   290
      Width           =   2175
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Bounties delvd"
      Height          =   285
      Index           =   14
      Left            =   4800
      TabIndex        =   36
      Top             =   1050
      Width           =   1305
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "add Passenger/s"
      Height          =   285
      Index           =   13
      Left            =   4800
      TabIndex        =   31
      Top             =   570
      Width           =   1365
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Add Crew on Goal completion"
      Height          =   225
      Index           =   12
      Left            =   6060
      TabIndex        =   30
      Top             =   90
      Width           =   2205
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Job on Goal completion"
      Height          =   285
      Index           =   11
      Left            =   2670
      TabIndex        =   29
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Job must be completed to reach this Goal"
      Height          =   285
      Index           =   10
      Left            =   2670
      TabIndex        =   28
      Top             =   1950
      Width           =   3045
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions for the next Goal"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   27
      Top             =   3150
      Width           =   6315
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planet or Ship"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   26
      Top             =   2730
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Minimum Negot"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Minimum Tech"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Minimum Fight"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   1590
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Misbehaves"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "the number of Misbehaves to meet the Goal"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turn Limit"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   810
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Minimum Cash"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Solid with these Contacts OR this Solid Count"
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   19
      Top             =   90
      Width           =   3315
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Goal: 0"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   90
      Width           =   1215
   End
End
Attribute VB_Name = "frmGoals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StoryID As Integer, goal As Integer

Private Sub cbo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      cbo(Index).ListIndex = -1
   End If
      
End Sub

Private Sub chkGoal_Click()
   comboRefresh
End Sub

Private Sub cmd_Click(Index As Integer)
Dim frmJobEdit As frmJobEditor
Dim SQL
   Select Case Index
   Case 0 ' delete
      DB.Execute "DELETE FROM StoryGoals WHERE StoryID = " & StoryID & " AND Goal = " & goal
      Me.hide
   Case 1 ' Save
      If Nz(varDLookup("StoryID", "StoryGoals", "StoryID=" & StoryID & " AND Goal=" & goal), 0) = StoryID Then
         SQL = "UPDATE StoryGoals SET Instructions = " & "'" & SQLFilter(txt(8)) & "',"
         SQL = SQL & "Solid = " & IIf(getSolid = "", "NULL", "'" & getSolid & "'") & ","
         SQL = SQL & "AddCrew = " & IIf(getList(lstCrew) = "", "NULL", "'" & getList(lstCrew) & "'") & ","
         SQL = SQL & "SolidCount = " & CStr(Val(txt(6))) & ","
         SQL = SQL & "IssueJobID = " & IIf(GetCombo(cbo(1)) = -1, "0", GetCombo(cbo(1))) & ","
         SQL = SQL & "CompleteJobID = " & IIf(GetCombo(cbo(0)) = -1, "0", GetCombo(cbo(0))) & ","
         SQL = SQL & "Cash = " & CStr(Val(txt(0))) & ","
         SQL = SQL & "TurnLimit = " & CStr(Val(txt(1))) & ","
         SQL = SQL & "Misbehaves = " & CStr(Val(txt(2))) & ","
         SQL = SQL & "Fight = " & CStr(Val(txt(3))) & ","
         SQL = SQL & "Tech = " & CStr(Val(txt(4))) & ","
         SQL = SQL & "Negotiate = " & CStr(Val(txt(5))) & ","
         SQL = SQL & "SectorID = " & IIf(GetCombo(cbo(2)) = -1, "0", GetCombo(cbo(2))) & ","
         SQL = SQL & "Win = " & CStr(chkWin.Value) & ","
         SQL = SQL & "clearAlliance = " & CStr(chkClearAlliance.Value) & ","
         SQL = SQL & "clearReaver = " & CStr(chkClearReaver.Value) & ","
         SQL = SQL & "NoUnfinished = " & CStr(chkUnfinished.Value) & ","
         SQL = SQL & "Warrant = " & CStr(chkWarrant.Value) & ","
         SQL = SQL & "Passenger = " & CStr(Val(txt(7))) & ","
         SQL = SQL & "Bounties = " & CStr(Val(txt(9))) & ","
         SQL = SQL & "chngInCutters = " & CStr(Val(txt(10))) & ","
         If GetCombo(cbo(3)) = -1 Then
            SQL = SQL & " doSolid=0"
         Else
            SQL = SQL & " doSolid=" & GetCombo(cbo(3))
         End If
         SQL = SQL & " WHERE StoryID = " & StoryID & " AND Goal = " & goal
         
         
      Else
         SQL = "INSERT INTO StoryGoals (StoryID, Goal, Instructions, Solid, AddCrew, SolidCount, IssueJobID, CompleteJobID, Cash, TurnLimit, Misbehaves, Fight, "
         SQL = SQL & "Tech, Negotiate, SectorID, Win, NoUnfinished, Passenger, Bounties, chngInCutters, clearAlliance, clearReaver, Warrant) VALUES ("
         SQL = SQL & CStr(StoryID) & ","
         SQL = SQL & CStr(goal) & ","
         SQL = SQL & "'" & SQLFilter(txt(8)) & "',"
         SQL = SQL & IIf(getSolid = "", "NULL", "'" & getSolid & "'") & ","
         SQL = SQL & IIf(getList(lstCrew) = "", "NULL", "'" & getList(lstCrew) & "'") & ","
         SQL = SQL & CStr(Val(txt(6))) & ","
         SQL = SQL & IIf(GetCombo(cbo(1)) = -1, "0", GetCombo(cbo(1))) & ","
         SQL = SQL & IIf(GetCombo(cbo(0)) = -1, "0", GetCombo(cbo(0))) & ","
         SQL = SQL & CStr(Val(txt(0))) & ","
         SQL = SQL & CStr(Val(txt(1))) & ","
         SQL = SQL & CStr(Val(txt(2))) & ","
         SQL = SQL & CStr(Val(txt(3))) & ","
         SQL = SQL & CStr(Val(txt(4))) & ","
         SQL = SQL & CStr(Val(txt(5))) & ","
         SQL = SQL & IIf(GetCombo(cbo(2)) = -1, "0", GetCombo(cbo(2))) & ","
         SQL = SQL & CStr(chkWin.Value) & ","
         SQL = SQL & CStr(chkUnfinished.Value) & ","
         SQL = SQL & CStr(Val(txt(7))) & ","
         SQL = SQL & CStr(Val(txt(9))) & ","
         SQL = SQL & CStr(Val(txt(10))) & ","
         SQL = SQL & CStr(chkClearAlliance.Value) & ","
         SQL = SQL & CStr(chkClearReaver.Value) & ","
         SQL = SQL & CStr(chkWarrant.Value) & ")"
      End If
      
      DB.Execute SQL
      
      Me.hide
   
   Case 2 ' cancel
       Me.hide
   Case 3
      If GetCombo(cbo(0)) < 1 Then Exit Sub
      Set frmJobEdit = New frmJobEditor
      frmJobEdit.lockEdits = True
      frmJobEdit.JobCardID = GetCombo(cbo(0))
      frmJobEdit.Show 1
   Case 4
      If GetCombo(cbo(1)) < 1 Then Exit Sub
      Set frmJobEdit = New frmJobEditor
      frmJobEdit.lockEdits = True
      frmJobEdit.JobCardID = GetCombo(cbo(1))
      frmJobEdit.Show 1
   End Select
End Sub

Private Sub Form_Load()
Dim X, filter As String
   
   LoadCombo lstContacts, "contact", " WHERE ContactID > 0 and ContactID < 10"
   LoadCombo cbo(2), "planet"
   LoadCombo cbo(3), "contact", " WHERE ContactID > 0 and ContactID < 10"
   filter = Nz(varDLookup("ExcludeCrew", "Story", "StoryID=" & StoryID))
   If filter <> "" Then LoadCombo lstCrew, "crew", " WHERE CrewID IN(" & filter & ") ORDER BY CrewName"
   comboRefresh
   Me.lbl(0).Caption = "Goal " & CStr(goal)
   If Nz(varDLookup("StoryID", "StoryGoals", "StoryID=" & StoryID & " AND Goal=" & goal), 0) = StoryID Then
      refreshGoal
      cmd(0).Visible = True
   End If
   
   If goal = 0 Then
      For X = 0 To 7
         txt(X).Enabled = False
      Next X

      For X = 0 To cbo.Count - 1
         cbo(X).Enabled = False
      Next X
      cbo(1).Enabled = True
      lstContacts.Enabled = False
      lstCrew.Enabled = False
      chkWin.Enabled = False
      chkUnfinished.Enabled = False
      Me.Frame1.Enabled = False
   End If
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub comboRefresh()
   LoadCombo cbo(0), "contactdeck", IIf(chkGoal.Value = 1, " WHERE ContactID = 0 ", "") & " ORDER BY Jobname"
   LoadCombo cbo(1), "contactdeck", IIf(chkGoal.Value = 1, " WHERE ContactID = 0 ", "") & " ORDER BY Jobname"
End Sub

Private Function getSolid() As String
Dim X
   If Val(txt(6)) > 0 Then 'overrides
      getSolid = ""
      Exit Function
   End If
   With lstContacts
      For X = 0 To .ListCount - 1
         If .selected(X) Then
            getSolid = getSolid & IIf(getSolid = "", "", ",") & CStr(.ItemData(X))
         End If
      Next X
   End With
   
End Function

Private Sub setSolid(ByVal solids As String)
Dim X, Y, a() As String

   If solids = "" Then Exit Sub
   With lstContacts
   
         a = Split(solids, ",")
         For Y = LBound(a) To UBound(a)
            For X = 0 To .ListCount - 1
               If .ItemData(X) = Val(a(Y)) Then
                  .selected(X) = True
                  Exit For
               End If
            Next X
         Next Y
      
   End With
   
End Sub

Private Function getList(cbo As Control) As String
Dim X
   With cbo
      For X = 0 To .ListCount - 1
         If .selected(X) Then
            getList = getList & IIf(getList = "", "", ",") & CStr(.ItemData(X))
         End If
      Next X
   End With
   
End Function

Private Function SetList(cbo As Control, ByVal theList As String) As Integer
Dim X, Y, a() As String

   If theList = "" Then Exit Function
   With cbo
   
         a = Split(theList, ",")
         For Y = LBound(a) To UBound(a)
            For X = 0 To .ListCount - 1
               If .ItemData(X) = Val(a(Y)) Then
                  .selected(X) = True
                  SetList = SetList + 1
                  Exit For
               End If
            Next X
         Next Y
      
   End With
   
End Function


Private Sub refreshGoal()
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM StoryGoals WHERE StoryID=" & StoryID & " AND Goal=" & goal
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      setSolid Nz(rst!Solid)
      txt(8) = Nz(rst!Instructions & "")
      SetList lstCrew, Nz(rst!AddCrew)
      SetCombo cbo(1), "", rst!IssueJobID
      SetCombo cbo(0), "", rst!CompleteJobID
      txt(0) = CStr(rst!Cash)
      txt(1) = CStr(rst!TurnLimit)
      txt(2) = CStr(rst!Misbehaves)
      txt(3) = CStr(rst!fight)
      txt(4) = CStr(rst!tech)
      txt(5) = CStr(rst!Negotiate)
      txt(6) = CStr(rst!SolidCount)
      txt(7) = CStr(rst!Passenger)
      txt(9) = CStr(rst!Bounties)
      txt(10) = CStr(rst!chngInCutters)
      If rst!SectorID > 0 Then SetCombo cbo(2), "", rst!SectorID
      SetCombo cbo(3), "", rst!doSolid
      chkWin.Value = rst!win
      chkUnfinished.Value = rst!NoUnfinished
      chkClearAlliance.Value = rst!clearAlliance
      chkClearReaver.Value = rst!clearReaver
      chkWarrant.Value = rst!Warrant
   End If
   rst.Close
   Set rst = Nothing
End Sub

