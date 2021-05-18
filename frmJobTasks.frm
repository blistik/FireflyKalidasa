VERSION 5.00
Begin VB.Form frmJobTasks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Job Task"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmJobTasks.frx":0000
   ScaleHeight     =   2610
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   1
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1830
      Width           =   2235
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "save Story and continue"
      Top             =   2190
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Cancel          =   -1  'True
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
      Index           =   1
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "save Story and continue"
      Top             =   2190
      Width           =   1035
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00CBE1ED&
      Caption         =   "Double Down"
      Height          =   255
      Left            =   3870
      TabIndex        =   20
      ToolTipText     =   "If the card suits of 2 or more of the Misbehave Cards passed match, take double pay. Must be on end task."
      Top             =   1860
      Width           =   1275
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   6930
      TabIndex        =   17
      Text            =   "0"
      Top             =   1140
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   6930
      TabIndex        =   16
      Text            =   "0"
      ToolTipText     =   "20=all you can load"
      Top             =   1470
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   4950
      TabIndex        =   13
      Text            =   "0"
      Top             =   1140
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   4950
      TabIndex        =   12
      Text            =   "0"
      Top             =   1470
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   3030
      TabIndex        =   9
      Text            =   "0"
      ToolTipText     =   "+pos=load, -ve=unload. +/-14=any amount"
      Top             =   1140
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   3030
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "+pos=load, -ve=unload. +/-14=any amount"
      Top             =   1470
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   1050
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "+pos=load, -ve=unload. +/-14=any amount"
      Top             =   1140
      Width           =   885
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   8
      Left            =   1050
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "+pos=load, -ve=unload. +/-14=any amount"
      Top             =   1470
      Width           =   885
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   7845
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   255
      Width           =   7875
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Planet"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   24
      Top             =   1860
      Width           =   585
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Misbehaves"
      Height          =   225
      Index           =   6
      Left            =   6030
      TabIndex        =   19
      Top             =   1170
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Goods"
      Height          =   225
      Index           =   5
      Left            =   6030
      TabIndex        =   18
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel"
      Height          =   225
      Index           =   3
      Left            =   4200
      TabIndex        =   15
      Top             =   1170
      Width           =   705
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Parts"
      Height          =   225
      Index           =   2
      Left            =   4200
      TabIndex        =   14
      Top             =   1500
      Width           =   705
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Passengers"
      Height          =   225
      Index           =   1
      Left            =   2130
      TabIndex        =   11
      Top             =   1170
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Fugitives"
      Height          =   225
      Index           =   0
      Left            =   2130
      TabIndex        =   10
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      Height          =   225
      Index           =   14
      Left            =   120
      TabIndex        =   7
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraband"
      Height          =   225
      Index           =   15
      Left            =   120
      TabIndex        =   6
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Task Details"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Task ID && summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   4095
   End
End
Attribute VB_Name = "frmJobTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public JobID As Integer

Private Sub cbo_Click(Index As Integer)
   If Index = 0 And cbo(0).ListIndex > -1 Then
      JobID = cbo(0).ItemData(cbo(0).ListIndex)
      RefreshJob
   End If
End Sub

Private Sub cmd_Click(Index As Integer)
   Select Case Index
   Case 0 'save
      If cbo(0).ListIndex = -1 Then
         JobID = newJob
         cbo(0).AddItem CStr(JobID) & " - " & txt(0)
         cbo(0).ItemData(cbo(0).NewIndex) = JobID
      End If
      saveJob
      SetCombo cbo(0), "", JobID
         
   Case 1 'close
      JobID = GetCombo(cbo(0))
      Me.Hide
   End Select
   
   
End Sub

Private Sub Form_Load()
   LoadCombo cbo(0), "task", "WHERE SectorID > 0"
   LoadCombo cbo(1), "planet", " ORDER BY PlanetName"
   If JobID > 1 Then
      SetCombo cbo(0), "", JobID
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub RefreshJob()
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM Job WHERE JobID =" & JobID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then

      SetCombo cbo(1), "", rst!SectorID
      txt(0) = Nz(rst!JobDesc)
      txt(1) = Nz(rst!fugitive)
      txt(2) = Nz(rst!passenger)
      txt(3) = Nz(rst!parts)
      txt(4) = Nz(rst!fuel)
      txt(5) = Nz(rst!tagnbag)
      txt(6) = Nz(rst!misbehave)
      txt(7) = Nz(rst!cargo)
      txt(8) = Nz(rst!contraband)
      Check1.Value = rst!DoubleDown
      
   End If
   rst.Close
   Set rst = Nothing


End Sub

Private Sub saveJob()
Dim SQL
On Error GoTo err_handler
   SQL = "UPDATE Job Set "
   SQL = SQL & "JobDesc='" & SQLFilter(txt(0)) & "'"
   SQL = SQL & ", Cargo=" & CStr(Val(txt(7)))
   SQL = SQL & ", Contraband=" & CStr(Val(txt(8)))
   SQL = SQL & ", Passenger=" & CStr(Val(txt(2)))
   SQL = SQL & ", Fugitive=" & CStr(Val(txt(1)))
   SQL = SQL & ", Fuel=" & CStr(Val(txt(4)))
   SQL = SQL & ", Parts=" & CStr(Val(txt(3)))
   SQL = SQL & ", Misbehave=" & CStr(Val(txt(6)))
   SQL = SQL & ", TagnBag=" & CStr(Val(txt(5)))
   SQL = SQL & ", DoubleDown=" & Check1.Value
   SQL = SQL & ", SectorID=" & GetCombo(cbo(1))
   SQL = SQL & " WHERE JobID=" & JobID
   DB.Execute SQL
normal_exit:
   Exit Sub
   
err_handler:
   MsgBox "Error: " & vbCrLf & Err.Description
   Resume normal_exit
End Sub

Private Function newJob() As Integer
Dim rst As New ADODB.Recordset
Dim SQL
On Error GoTo err_handler

   SQL = "Job"
   rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
   rst.AddNew
   rst!JobDesc = "New Task at " & Now()
   rst.Update
   newJob = rst!JobID
   rst.Close
   Set rst = Nothing
   
normal_exit:
   Exit Function
   
err_handler:
   MsgBox "Error: " & vbCrLf & Err.Description
   Resume normal_exit
End Function
