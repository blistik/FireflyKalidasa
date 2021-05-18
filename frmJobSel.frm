VERSION 5.00
Begin VB.Form frmJobSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Job Selection"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmJobSel.frx":0000
   ScaleHeight     =   1260
   ScaleWidth      =   12780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FF8080&
      Caption         =   "View"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   750
      Width           =   915
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   270
      Width           =   12555
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
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   750
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick a Job"
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
      Left            =   150
      TabIndex        =   2
      Top             =   60
      Width           =   3855
   End
End
Attribute VB_Name = "frmJobSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CardID As Integer, jobFilter As String

Private Sub cmd_Click()
   If cbo.ListIndex > -1 Then
      CardID = GetCombo(cbo)
      Me.Hide
   End If
End Sub


Private Sub cmdView_Click()
Dim frmJobEdit As frmJobEditor
   If GetCombo(cbo) > 0 Then
      Set frmJobEdit = New frmJobEditor
      frmJobEdit.lockEdits = True
      frmJobEdit.JobCardID = GetCombo(cbo)
      frmJobEdit.Show 1
   End If
End Sub

Private Sub Form_Load()
   refreshCbo
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub refreshCbo()
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Contact.ContactName, ContactDeck.CardID, ContactDeck.Pay, ContactDeck.Bonus, ContactDeck.Keywords, ContactDeck.Immoral, ContactDeck.JobName, "
   SQL = SQL & "Job.JobID,  Job.JobDesc, Job_1.JobDesc AS Job2Desc, Job_1.JobID AS Job2 "
   SQL = SQL & "FROM (Contact INNER JOIN (Job INNER JOIN (PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) ON Job.JobID = ContactDeck.Job1ID) ON Contact.ContactID = ContactDeck.ContactID) LEFT JOIN Job AS Job_1 ON ContactDeck.Job2ID = Job_1.JobID "
   SQL = SQL & "WHERE PlayerJobs.PlayerID=" & player.ID & " AND PlayerJobs.JobStatus " & jobFilter

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      cbo.AddItem rst!ContactName & ": " & rst!JobName & " $" & rst!pay & " " & IIf(rst!bonus > 0, "+$" & rst!bonus & " Bonus. ", "") & getPlanetSector(rst!JobID) & ": " & rst!JobDesc & " " & IIf(IsNull(rst!Job2), "", " -/- " & getPlanetSector(rst!Job2) & ": " & rst!Job2Desc) & IIf(IsNull(rst!KeyWords), "", " (" & rst!KeyWords & ")")
      cbo.ItemData(cbo.NewIndex) = rst!CardID
      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing
End Sub
