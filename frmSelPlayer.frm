VERSION 5.00
Begin VB.Form frmSelPlayer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Player selection"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelPlayer.frx":0000
   ScaleHeight     =   1125
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   915
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick your player"
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
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   3855
   End
End
Attribute VB_Name = "frmSelPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public playerID As Integer

Private Sub cmd_Click()
   playsnd 8
   If cbo.ListIndex > -1 Then
      playerID = cbo.ItemData(cbo.ListIndex)
      Me.Hide
   End If
End Sub

Private Sub Form_Load()
   RefreshList
End Sub

Public Sub RefreshList()
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM Players WHERE Name IS NOT NULL "
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   Do While Not rst.EOF
      cbo.AddItem rst!Name
      cbo.ItemData(cbo.NewIndex) = rst!playerID
      rst.MoveNext
   Loop
   rst.Close
   If cbo.ListCount > 0 Then
      cbo.ListIndex = 0
   End If
   
   Set rst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub
