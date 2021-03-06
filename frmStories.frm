VERSION 5.00
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmStories 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "View/Edit Story"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStories.frx":0000
   ScaleHeight     =   4725
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SftTree.SftTree sftTree 
      Height          =   2325
      Left            =   60
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2340
      Width           =   11895
      _Version        =   262144
      _ExtentX        =   20981
      _ExtentY        =   4101
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
      Columns         =   16
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
      ColWidth13      =   30
      ColTitle13      =   "Corv"
      ColBmp13        =   "frmStories.frx":14D8B
      ColWidth14      =   30
      ColTitle14      =   "Cruis"
      ColBmp14        =   "frmStories.frx":14DA7
      ColWidth15      =   30
      ColTitle15      =   "Passengers"
      ColBmp15        =   "frmStories.frx":14DC3
      MouseIcon       =   "frmStories.frx":14DDF
      ColHeaderBackColor=   0
      ColHeaderForeColor=   65280
      ForeColor       =   8833235
      BackColor       =   3353720
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmStories.frx":14DFB
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      ColPict0        =   "frmStories.frx":14E17
      ColPict1        =   "frmStories.frx":14E33
      ColPict2        =   "frmStories.frx":14E4F
      ColPict3        =   "frmStories.frx":14E6B
      ColPict4        =   "frmStories.frx":14E87
      ColPict5        =   "frmStories.frx":14EA3
      ColPict6        =   "frmStories.frx":14EBF
      ColPict7        =   "frmStories.frx":14EDB
      ColPict8        =   "frmStories.frx":14EF7
      ColPict9        =   "frmStories.frx":14F13
      ColPict10       =   "frmStories.frx":14F2F
      ColPict11       =   "frmStories.frx":14F4B
      ColPict12       =   "frmStories.frx":14F67
      ColPict13       =   "frmStories.frx":14F83
      ColPict14       =   "frmStories.frx":14F9F
      ColPict15       =   "frmStories.frx":14FBB
      BackgroundPicture=   "frmStories.frx":14FD7
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
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
      TabIndex        =   19
      ToolTipText     =   "add a Goal to this story"
      Top             =   1800
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Continue"
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
      TabIndex        =   17
      ToolTipText     =   "save Story and continue"
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
      TabIndex        =   18
      ToolTipText     =   "delete this Story"
      Top             =   1260
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Priming the Pump"
      Height          =   2265
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   10605
      Begin VB.ListBox lstCrew 
         BackColor       =   &H00CBE1ED&
         Height          =   1635
         Left            =   8490
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   390
         Width           =   1995
      End
      Begin VB.TextBox txt 
         Height          =   1155
         Index           =   6
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frmStories.frx":14FF3
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
         Height          =   1635
         Left            =   6360
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   390
         Width           =   2025
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   5340
         TabIndex        =   4
         Text            =   "2"
         Top             =   1170
         Width           =   885
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   5340
         TabIndex        =   3
         Text            =   "6"
         Top             =   780
         Width           =   885
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   5340
         TabIndex        =   2
         Text            =   "3000"
         Top             =   390
         Width           =   885
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Exclude Crew"
         Height          =   225
         Left            =   8490
         TabIndex        =   21
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         Left            =   3960
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Jobs (discard to 3)"
         Height          =   225
         Left            =   6360
         TabIndex        =   12
         Top             =   195
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parts"
         Height          =   285
         Index           =   2
         Left            =   3960
         TabIndex        =   11
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fuel"
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   10
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cash"
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   9
         Top             =   390
         Width           =   1215
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

Private Sub cmd_Click(Index As Integer)
Dim frmGoal As frmGoals, SQL

   Select Case Index
   Case 0 ' cont
      SQL = "UPDATE Story SET StoryTitle= " & "'" & SQLFilter(txt(5)) & "',"
      SQL = SQL & " StoryDesc = " & " '" & SQLFilter(txt(6)) & "',"
      SQL = SQL & " StartingCash = " & CStr(Val(txt(0))) & ","
      SQL = SQL & " StartingFuel = " & CStr(Val(txt(1))) & ","
      SQL = SQL & " StartingParts = " & CStr(Val(txt(2))) & ","
      SQL = SQL & " StartingCrew = " & CStr(Val(txt(3))) & ","
      SQL = SQL & " CrewCostLimit = " & CStr(Val(txt(4))) & ","
      SQL = SQL & " StartingJobs = " & IIf(getList(lstContacts) = "", "NULL", "'" & getList(lstContacts) & "'") & ","
      SQL = SQL & " ExcludeCrew = " & IIf(getList(lstCrew) = "", "NULL", "'" & getList(lstCrew) & "'")
      SQL = SQL & " WHERE StoryID = " & StoryID
      DB.Execute SQL
      
      Me.Hide
   
   Case 1 ' add Goal
      Set frmGoal = New frmGoals
      frmGoal.StoryID = StoryID
      frmGoal.Goal = sftTree.ListCount
      frmGoal.Show 1, Me
      RefreshGoals
   
   Case 2 'delete
      If MsgBox("Are you sure you want to Delete this Story?", vbYesNo + vbQuestion, "Delete Story") = vbNo Then Exit Sub
      DB.Execute "Delete from StoryGoals WHERE StoryID = " & StoryID
      DB.Execute "Delete from Story WHERE StoryID = " & StoryID
      StoryID = 0
      Me.Hide
   

   End Select
End Sub

Private Sub Form_Load()
   LoadCombo lstContacts, "contact", " WHERE ContactID > 0"
   LoadCombo lstCrew, "crew", " Order by CrewName"
   refreshHeader
   RefreshGoals
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub refreshHeader()
Dim rst As New ADODB.Recordset
Dim SQL, Index

      SQL = "SELECT * FROM Story WHERE StoryID =" & StoryID

      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         txt(5) = Nz(rst!StoryTitle)
         txt(6) = Nz(rst!StoryDesc)
         txt(0) = Nz(rst!StartingCash)
         txt(1) = Nz(rst!StartingFuel)
         txt(2) = Nz(rst!StartingParts)
         txt(3) = Nz(rst!StartingCrew)
         txt(4) = Nz(rst!CrewCostLimit)
         SetList lstContacts, Nz(rst!StartingJobs)
         Label2 = "Exclude Crew (" & CStr(SetList(lstCrew, Nz(rst!ExcludeCrew))) & " selected)"
      Else 'new
         DB.Execute "Insert into Story (StoryID,StoryTitle, Active) VALUES (" & StoryID & ",'add a new story title here..',1)"
      End If
      rst.Close
      Set rst = Nothing


End Sub

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
         .CellText(Index, 6) = rst!Win & ""
         .CellText(Index, 7) = rst!TurnLimit & ""
         .CellText(Index, 8) = rst!fight & ""
         .CellText(Index, 9) = rst!tech & ""
         .CellText(Index, 10) = rst!Negotiate & ""
         .CellText(Index, 11) = rst!Misbehaves & ""
         .CellText(Index, 12) = rst!SectorID & ""
         .CellText(Index, 13) = rst!MeetCorvette & ""
         .CellText(Index, 14) = rst!MeetCruiser & ""
         .CellText(Index, 15) = rst!Passenger & ""
         
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
         If .Selected(x) Then
            getList = getList & IIf(getList = "", "", ",") & CStr(.ItemData(x))
         End If
      Next x
   End With
   
End Function


Private Function getSelected(cbo As Control) As Integer
Dim x
   With cbo
      For x = 0 To .ListCount - 1
         If .Selected(x) Then
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
                  .Selected(x) = True
                  SetList = SetList + 1
                  Exit For
               End If
            Next x
         Next y
      
   End With
   
End Function

Private Sub lstCrew_ItemCheck(Item As Integer)
   Label2 = "Exclude Crew (" & CStr(getSelected(lstCrew)) & " selected)"
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
