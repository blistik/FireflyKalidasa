VERSION 5.00
Begin VB.Form frmGoals 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Story Goal - Criteria to achieve"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGoals.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   35
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
      TabIndex        =   34
      Top             =   2160
      Width           =   645
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   5730
      TabIndex        =   13
      Text            =   "0"
      ToolTipText     =   "add Passengers on Goal completion"
      Top             =   750
      Width           =   885
   End
   Begin VB.CheckBox chkMeet 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Meet with Cruiser"
      Height          =   405
      Index           =   1
      Left            =   5730
      TabIndex        =   15
      ToolTipText     =   "be in same Sector as Cruiser"
      Top             =   1500
      Width           =   1035
   End
   Begin VB.CheckBox chkMeet 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Meet with Corvette"
      Height          =   405
      Index           =   0
      Left            =   5730
      TabIndex        =   14
      ToolTipText     =   "be in same Sector as Corvette"
      Top             =   1080
      Width           =   1035
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   5730
      TabIndex        =   12
      Text            =   "0"
      ToolTipText     =   "Any Solids to this count"
      Top             =   270
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
      Left            =   7230
      TabIndex        =   19
      ToolTipText     =   "limit job lists to only those specifically made for goals"
      Top             =   1920
      Value           =   1  'Checked
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      Width           =   8385
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
      Left            =   2640
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   290
      Width           =   2985
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "add Passenger/s"
      Height          =   285
      Index           =   13
      Left            =   5730
      TabIndex        =   33
      Top             =   570
      Width           =   1365
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      Caption         =   "...OR this Solid Count"
      Height          =   225
      Index           =   12
      Left            =   5070
      TabIndex        =   32
      Top             =   90
      Width           =   1515
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Job on Goal completion"
      Height          =   285
      Index           =   11
      Left            =   2670
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   29
      Top             =   3150
      Width           =   6315
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planet"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Solid with these Contacts"
      Height          =   285
      Index           =   1
      Left            =   2670
      TabIndex        =   21
      Top             =   90
      Width           =   2655
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Goal: 0"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   20
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
Public StoryID As Integer, Goal As Integer

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
      DB.Execute "DELETE FROM StoryGoals WHERE StoryID = " & StoryID & " AND Goal = " & Goal
      Me.Hide
   Case 1 ' Save
      If Nz(varDLookup("StoryID", "StoryGoals", "StoryID=" & StoryID & " AND Goal=" & Goal), 0) = StoryID Then
         SQL = "UPDATE StoryGoals SET Instructions = " & "'" & SQLFilter(txt(8)) & "',"
         SQL = SQL & "Solid = " & IIf(getSolid = "", "NULL", "'" & getSolid & "'") & ","
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
         SQL = SQL & "MeetCorvette = " & CStr(chkMeet(0).Value) & ","
         SQL = SQL & "MeetCruiser = " & CStr(chkMeet(1).Value) & ","
         SQL = SQL & "Passenger = " & CStr(Val(txt(7)))
         SQL = SQL & " WHERE StoryID = " & StoryID & " AND Goal = " & Goal
         
         
      Else
         SQL = "INSERT INTO StoryGoals (StoryID, Goal, Instructions, Solid, SolidCount, IssueJobID, CompleteJobID, Cash, TurnLimit, Misbehaves, Fight, "
         SQL = SQL & "Tech, Negotiate, SectorID, Win, MeetCorvette, MeetCruiser, Passenger) VALUES ("
         SQL = SQL & CStr(StoryID) & ","
         SQL = SQL & CStr(Goal) & ","
         SQL = SQL & "'" & SQLFilter(txt(8)) & "',"
         SQL = SQL & IIf(getSolid = "", "NULL", "'" & getSolid & "'") & ","
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
         SQL = SQL & CStr(chkMeet(0).Value) & ","
         SQL = SQL & CStr(chkMeet(1).Value) & ","
         SQL = SQL & CStr(Val(txt(7))) & ")"
      End If
      
      DB.Execute SQL
      
      Me.Hide
   
   Case 2 ' cancel
       Me.Hide
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
Dim x
   LoadCombo lstContacts, "contact", " WHERE ContactID > 0"
   LoadCombo cbo(2), "planet"
   comboRefresh
   Me.lbl(0).Caption = "Goal " & CStr(Goal)
   If Nz(varDLookup("StoryID", "StoryGoals", "StoryID=" & StoryID & " AND Goal=" & Goal), 0) = StoryID Then
      refreshGoal
      cmd(0).Visible = True
   End If
   
   If Goal = 0 Then
      For x = 0 To 7
         txt(x).Enabled = False
      Next x

      For x = 0 To cbo.Count - 1
         cbo(x).Enabled = False
      Next x
      cbo(1).Enabled = True
      lstContacts.Enabled = False
      chkWin.Enabled = False
      chkMeet(0).Enabled = False
      chkMeet(1).Enabled = False
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
Dim x
   If Val(txt(6)) > 0 Then 'overrides
      getSolid = ""
      Exit Function
   End If
   With lstContacts
      For x = 0 To .ListCount - 1
         If .selected(x) Then
            getSolid = getSolid & IIf(getSolid = "", "", ",") & CStr(.ItemData(x))
         End If
      Next x
   End With
   
End Function

Private Sub setSolid(ByVal solids As String)
Dim x, y, a() As String

   If solids = "" Then Exit Sub
   With lstContacts
   
         a = Split(solids, ",")
         For y = LBound(a) To UBound(a)
            For x = 0 To .ListCount - 1
               If .ItemData(x) = Val(a(y)) Then
                  .selected(x) = True
                  Exit For
               End If
            Next x
         Next y
      
   End With
   
End Sub

Private Sub refreshGoal()
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM StoryGoals WHERE StoryID=" & StoryID & " AND Goal=" & Goal
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      txt(8) = Nz(rst!Instructions)
      setSolid Nz(rst!Solid)
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
      If rst!SectorID > 0 Then SetCombo cbo(2), "", rst!SectorID
      chkWin.Value = rst!Win
   End If
   rst.Close
   Set rst = Nothing
End Sub

