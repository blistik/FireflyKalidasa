VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
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
   LockControls    =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Gorram it!"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7290
      Width           =   1995
   End
   Begin VB.Image imgHeader 
      Height          =   1065
      Left            =   940
      Top             =   420
      Width           =   2295
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
      Left            =   1740
      TabIndex        =   4
      Top             =   7060
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
      Left            =   1740
      TabIndex        =   3
      Top             =   6310
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
      Left            =   1740
      TabIndex        =   2
      Top             =   5650
      Width           =   765
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl skillPic 
      Height          =   675
      Index           =   3
      Left            =   810
      Top             =   6960
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      Effects         =   "frmBoarded.frx":0000
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl skillPic 
      Height          =   675
      Index           =   2
      Left            =   810
      Top             =   6240
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      Effects         =   "frmBoarded.frx":0018
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl skillPic 
      Height          =   675
      Index           =   1
      Left            =   810
      Top             =   5520
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      Effects         =   "frmBoarded.frx":0030
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "You have been Boarded"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1395
      Left            =   4103
      TabIndex        =   0
      Top             =   270
      Width           =   7155
      WordWrap        =   -1  'True
   End
   Begin VB.Image img 
      Height          =   750
      Index           =   0
      Left            =   12960
      Stretch         =   -1  'True
      Top             =   210
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "frmBoarded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public thisplayer As Integer
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Sub cmd_Click()
   Me.hide
End Sub

Private Sub Form_Load()
   If thisplayer > 0 Then
   
      refreshPlayer
   
   End If
End Sub

Private Sub refreshPlayer()
Dim rst As New ADODB.Recordset, X
Dim SQL
   Set Me.Picture = LoadPicture(App.Path & "\pictures\boarded.jpg")
   imgHeader.Picture = LoadPicture(App.Path & "\gui\Firefly" & thisplayer & ".jpg")
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
         img(X).top = 1920
         img(X).Left = 810
      Else
         img(X).Height = 2300
         img(X).Width = 2000
         If X < 7 Then
            img(X).top = 1920
            img(X).Left = 3810 + (2040 * (X - 2))
         Else
            img(X).top = 4290
            img(X).Left = 3810 + (2040 * (X - 7))
         End If
      End If

      rst.MoveNext
   Wend
   rst.Close
   Set rst = Nothing


   For X = 1 To 3
      Set skillPic(X).Picture = LoadPictureGDIplus(App.Path & "\pictures\" & picSkill(X) & ".bmp")
      lblSkill(X) = getSkill(thisplayer, cstrSkill(X))
      skillPic(X).TransparentColor = 0
      skillPic(X).TransparentColorMode = lvicUseTransparentColor
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
