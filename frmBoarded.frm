VERSION 5.00
Begin VB.Form frmBoarded 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Boarded"
   ClientHeight    =   8115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14685
   ForeColor       =   &H003DCBFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmBoarded.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgHotspot 
      Height          =   945
      Index           =   3
      Left            =   13140
      Top             =   6990
      Width           =   915
   End
   Begin VB.Image imgHotspot 
      Height          =   945
      Index           =   2
      Left            =   12060
      Top             =   6990
      Width           =   915
   End
   Begin VB.Image imgHotspot 
      Height          =   945
      Index           =   1
      Left            =   10920
      Top             =   6990
      Width           =   915
   End
   Begin VB.Image imgHotspot 
      Height          =   945
      Index           =   0
      Left            =   8880
      Top             =   6990
      Width           =   1875
   End
   Begin VB.Label lblAttacker 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "attacker"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   345
      Left            =   30
      TabIndex        =   3
      Top             =   1230
      Width           =   3345
   End
   Begin VB.Image imgHeader 
      Height          =   735
      Left            =   810
      Stretch         =   -1  'True
      Top             =   1830
      Width           =   1815
   End
   Begin VB.Label lblSkill 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   465
      Index           =   3
      Left            =   2250
      TabIndex        =   2
      Top             =   7260
      Width           =   765
   End
   Begin VB.Label lblSkill 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   465
      Index           =   2
      Left            =   2250
      TabIndex        =   1
      Top             =   6785
      Width           =   765
   End
   Begin VB.Label lblSkill 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Index           =   1
      Left            =   2250
      TabIndex        =   0
      Top             =   6280
      Width           =   765
   End
   Begin VB.Image img 
      Height          =   750
      Index           =   0
      Left            =   13650
      Stretch         =   -1  'True
      Top             =   150
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmBoarded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public thisplayer As Integer, result As Integer ' 0 = rearrange, 1 fight, 2 tech, 3 nego
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()
   playsnd 16
   If thisplayer > 0 Then
   
      refreshPlayer
   
   End If
End Sub

Private Sub refreshPlayer()
Dim rst As New ADODB.Recordset, X
Dim SQL
   Set Me.Picture = LoadPicture(App.Path & "\pictures\boarded.jpg")
   imgHeader.Picture = LoadPicture(App.Path & "\gui\FireflyBoard" & thisplayer & ".jpg")
   lblAttacker.Caption = PlayCode(thisplayer).PlayName
   X = 0
   SQL = "SELECT PlayerSupplies.CardID, SupplyDeck.CrewID, Crew.CrewName, Crew.Fight, Crew.Tech, Crew.Negotiate, Crew.Picture"
   SQL = SQL & " FROM PlayerSupplies INNER JOIN (Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON PlayerSupplies.CardID = SupplyDeck.CardID"
   SQL = SQL & " WHERE PlayerSupplies.OffJob=0 AND PlayerSupplies.PlayerID=" & thisplayer
   SQL = SQL & " ORDER BY Crew.Leader DESC, Crew.CrewName"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      X = X + 1
      If X > img.Count - 1 Then
         Load img(X)
      End If
      Set img(X).Picture = LoadPicture(App.Path & "\pictures\" & rst!Picture)
      img(X).Visible = True
      img(X).Tag = CStr(rst!CrewID)
      img(X).ToolTipText = rst!CrewName
      If X = 1 Then
         img(X).Height = 3420
         img(X).Width = 2940
         img(X).top = 2670
         img(X).Left = 240
      Else
         img(X).Height = 2107
         img(X).Width = 1811
         If X < 7 Then
            img(X).top = 1860
            img(X).Left = 4470 + (1836 * (X - 2))
         Else
            img(X).top = 3630
            img(X).Left = 4470 + (1836 * (X - 7))
         End If
      End If

      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing


   For X = 1 To 3
      'Set skillPic(X).Picture = LoadPictureGDIplus(App.Path & "\pictures\" & picSkill(X) & ".bmp")
      lblSkill(X) = getSkill(thisplayer, cstrSkill(X))
      'skillPic(X).TransparentColor = 0
      'skillPic(X).TransparentColorMode = lvicUseTransparentColor
   Next X

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub img_DblClick(Index As Integer)
Dim frmCrew As frmCrewSel
   If Val(img(Index).Tag) = 0 Then Exit Sub
   Set frmCrew = New frmCrewSel
   frmCrew.crewFilter = " WHERE CrewID =" & Val(img(Index).Tag)
   frmCrew.AlwaysOnTop = False
   frmCrew.Show 1
   Set frmCrew = Nothing
End Sub

Private Sub imgHotspot_Click(Index As Integer)
   result = Index
   Me.hide
End Sub
