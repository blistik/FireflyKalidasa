VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmShipUpgrdView 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "shipupgrd"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "OK"
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
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   90
      Width           =   615
   End
   Begin VB.ComboBox cboGear 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   610
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4020
      Width           =   4035
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   705
      Index           =   5
      Left            =   1050
      TabIndex        =   6
      Top             =   5850
      Width           =   2985
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   8
      Left            =   990
      TabIndex        =   5
      ToolTipText     =   "Origin"
      Top             =   6780
      Width           =   2325
   End
   Begin VB.Label lbl 
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   1470
      TabIndex        =   4
      Top             =   3690
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Index           =   1
      Left            =   690
      TabIndex        =   3
      Top             =   4560
      Width           =   3825
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Index           =   6
      Left            =   4050
      TabIndex        =   2
      Top             =   6350
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cyberpunk Is Not Dead"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   405
      Index           =   7
      Left            =   1350
      TabIndex        =   1
      ToolTipText     =   "Origin"
      Top             =   690
      Width           =   2655
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImg 
      Height          =   7170
      Left            =   0
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   12647
      Image           =   "frmShipUpgrdView.frx":0000
      Effects         =   "frmShipUpgrdView.frx":7AC9E
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pic 
      Height          =   2715
      Left            =   1110
      Top             =   1260
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4789
      Effects         =   "frmShipUpgrdView.frx":7ACB6
   End
End
Attribute VB_Name = "frmShipUpgrdView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gearFilter As String

'For use with USER32 Function SetWindowPos
Private Const HWND_TOPMOST = -&H1
Private Const HWND_NOTOPMOST = -&H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
'For use with USER32 Function SendMessage
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private bOnTopState As Boolean

Public Property Let AlwaysOnTop(bState As Boolean)
    Dim lFlag As Long
    On Error Resume Next
    If bState = True Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    bOnTopState = bState
    SetWindowPos Me.hwnd, lFlag, 0&, 0&, 0&, 0&, (SWP_NOSIZE Or SWP_NOMOVE)
End Property

Public Property Get AlwaysOnTop() As Boolean
    AlwaysOnTop = bOnTopState
End Property

Private Sub AlphaImg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub



Private Sub cboGear_Click()
   If cboGear.ListIndex = -1 Then Exit Sub
   
   refreshGear GetCombo(cboGear)
End Sub

Private Sub cmd_Click()
   playsnd 8
   Unload Me
End Sub

Private Sub Form_Load()
   LoadCombo cboGear, "shipupgrd", gearFilter
   If cboGear.ListCount > 0 Then
      cboGear.ListIndex = 0
   End If

End Sub

Private Sub lbl_DblClick(Index As Integer)
   If gearFilter = " Order By UpgradeName" Then
      If lbl(Index).Tag <> "" Then
         LoadCombo cboGear, "shipupgrd", " WHERE " & lbl(Index).Tag
      Else
         LoadCombo cboGear, "shipupgrd", gearFilter
      End If
      If cboGear.ListCount > 0 Then
         cboGear.ListIndex = 0
      End If
   End If
End Sub

Private Sub refreshGear(ByVal CardID)
Dim rst As New ADODB.Recordset, SQL, x, y

   SQL = "SELECT ShipUpgrade.*, SupplyDeck.CardID, SupplyDeck.SupplyID, Supply.Colour, Supply.SupplyName, PlayerSupplies.PlayerID, Players.Name "
   SQL = SQL & "FROM Players RIGHT JOIN (PlayerSupplies RIGHT JOIN (Supply RIGHT JOIN (ShipUpgrade LEFT JOIN SupplyDeck ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID) ON Supply.SupplyID = SupplyDeck.SupplyID) ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Players.PlayerID = PlayerSupplies.PlayerID "
   SQL = SQL & "WHERE SupplyDeck.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      cboGear.ToolTipText = rst!UpgradeName
      lbl(1) = rst!UpgradeDescr
      lbl(2).Visible = (Nz(rst!Name) <> "")
      If Nz(rst!Name) <> "" Then
         lbl(2) = "owned by: " & rst!Name
      End If
      
      lbl(5) = Nz(rst!Keyword)
      If lbl(5) <> "" Then lbl(5).Tag = "ShipUpgrade.KeyWord = '" & Nz(rst!Keyword) & "'"
      
      If IsNull(rst!Keyword) Then
         lbl(5).Visible = False
      Else
         lbl(5).Visible = True
         lbl(5).BackColor = 12574908
      End If
      
      lbl(6) = rst!pay
      lbl(6).Visible = True
      
      lbl(7) = rst!SupplyName
      lbl(7).BackColor = rst!Colour
      lbl(7).Tag = "SupplyDeck.SupplyID=" & rst!SupplyID
      
      lbl(8) = "CardID: " & rst!CardID & "    GearID: " & rst!ShipUpgradeID
            
      AlphaImg.TransparentColor = 0
      AlphaImg.TransparentColorMode = lvicUseTransparentColor
      
      If Not IsNull(rst!Picture) Then
         Set pic.Picture = LoadPictureGDIplus(App.Path & "\pictures\" & rst!Picture)
      Else
         Set pic.Picture = Nothing
      End If
   End If
   rst.Close

End Sub


