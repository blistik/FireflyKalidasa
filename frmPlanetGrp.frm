VERSION 5.00
Begin VB.Form frmPlanetGrp 
   BackColor       =   &H00CBE1ED&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "select/edit a Planet Group"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   Icon            =   "frmPlanetGrp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FF8080&
      Caption         =   "New"
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
      Left            =   200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7740
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7740
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7740
      Width           =   1035
   End
   Begin VB.ListBox lstPlanets 
      BackColor       =   &H00CBE1ED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7080
      Left            =   30
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   450
      Width           =   4365
   End
   Begin VB.ComboBox cmbGroupID 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Caption         =   "GroupID"
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   110
      Width           =   1215
   End
End
Attribute VB_Name = "frmPlanetGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public groupID As Integer

Private Sub cmbGroupID_Click()
   groupID = GetCombo(cmbGroupID)
   If groupID < 0 Then
      groupID = 0
   Else
      RefreshList
   End If
End Sub

Private Sub cmdCancel_Click()
   groupID = 0
   Me.hide
End Sub

Private Sub cmdNew_Click()

    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim nextID As Long
    Dim i As Long

    ' Determine next available GroupID
    SQL = "SELECT MAX(GroupID) AS MaxID FROM PlanetGroup"

    Set rs = New ADODB.Recordset
    rs.Open SQL, DB, adOpenForwardOnly, adLockReadOnly

    If rs.EOF Or IsNull(rs!MaxID) Then
        nextID = 1
    Else
        nextID = CLng(rs!MaxID) + 1
    End If

    rs.Close
    Set rs = Nothing

    ' Store it globally
    groupID = nextID

    ' Update dropdown
    cmbGroupID.AddItem CStr(nextID)
    cmbGroupID.ItemData(cmbGroupID.NewIndex) = nextID
    cmbGroupID.ListIndex = cmbGroupID.NewIndex

    ' Clear all checkboxes
    For i = 0 To lstPlanets.ListCount - 1
        lstPlanets.selected(i) = False
    Next i

End Sub

Private Sub cmdSave_Click()

    Dim SQL As String
    Dim i As Long

    ' Safety: must have a valid GroupID
    If groupID <= 0 Then
        MsgBox "No GroupID selected.", vbExclamation
        Exit Sub
    End If

    ' Delete existing rows for this GroupID
    SQL = "DELETE FROM PlanetGroup WHERE GroupID = " & groupID
    DB.Execute SQL

    ' Insert selected planets
    For i = 0 To lstPlanets.ListCount - 1
        If lstPlanets.selected(i) = True Then
            SQL = "INSERT INTO PlanetGroup (GroupID, SectorID) VALUES (" & _
                   groupID & ", " & lstPlanets.ItemData(i) & ")"
            DB.Execute SQL
        End If
    Next i

    ' Close the editor — frmGoals.cmd(5) will read cmbGroupID and update txt(11)
    Unload Me

End Sub


Private Sub Form_Load()
   LoadCombo lstPlanets, "planetsystem", " and PlanetID <> 26 ORDER BY System, PlanetName"
   LoadCombo cmbGroupID, "planetgroup", " ORDER BY GroupID"
   If groupID > 0 Then SetCombo cmbGroupID, "", groupID
End Sub

Public Sub RefreshList()

    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim i As Long
    Dim sectorID As Long

    ' Clear all checkboxes first
    For i = 0 To lstPlanets.ListCount - 1
        lstPlanets.selected(i) = False
    Next i

    ' If no valid group selected, exit cleanly
    If groupID <= 0 Then Exit Sub

    ' Query all SectorIDs for this GroupID
    SQL = "SELECT SectorID FROM PlanetGroup WHERE GroupID = " & groupID

    Set rs = New ADODB.Recordset
    rs.Open SQL, DB, adOpenForwardOnly, adLockReadOnly

    ' Loop through returned SectorIDs
    Do While Not rs.EOF
        sectorID = rs!sectorID

        ' Match against lstPlanets.ItemData
        For i = 0 To lstPlanets.ListCount - 1
            If lstPlanets.ItemData(i) = sectorID Then
                lstPlanets.selected(i) = True
                Exit For
            End If
        Next i

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

End Sub

