VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmSelPlayer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Player selection"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelPlayer.frx":0000
   ScaleHeight     =   4185
   ScaleWidth      =   3105
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
      Left            =   2130
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   915
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3740
      Width           =   1995
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pic 
      Height          =   3420
      Left            =   60
      Top             =   240
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   6033
      Effects         =   "frmSelPlayer.frx":BB00
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick your player to continue:"
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
      Top             =   20
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


Private Sub cbo_Click()
   If cbo.ListIndex = -1 Then Exit Sub
   setPic cbo.ItemData(cbo.ListIndex)
End Sub

Private Sub cmd_Click()
   playsnd 8
   If cbo.ListIndex > -1 Then
      playerID = cbo.ItemData(cbo.ListIndex)
      Me.hide
   End If
End Sub

Private Sub Form_Load()
   'RefreshList
End Sub

'return how many human players
Public Function RefreshList() As Integer
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM Players WHERE Name IS NOT NULL AND AI = 0"  'remove for testing >>>AI = 0<<<<<<<
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   Do While Not rst.EOF
      cbo.AddItem rst!Name
      cbo.ItemData(cbo.NewIndex) = rst!playerID
      rst.MoveNext
   Loop
   rst.Close
   
   RefreshList = cbo.ListCount
   If cbo.ListCount = 1 Then
      playerID = cbo.ItemData(0)
   ElseIf cbo.ListCount > 1 Then
      cbo.ListIndex = 0
   End If
   
   Set rst = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub setPic(ByVal pID As Integer)
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Crew.Picture FROM Crew INNER JOIN Players ON Crew.CrewID = Players.Leader WHERE Players.PlayerID=" & pID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      Set pic.Picture = LoadPictureGDIplus(App.Path & "\pictures\" & rst!Picture)
   End If

End Sub
