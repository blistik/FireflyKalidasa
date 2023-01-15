VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmGearView 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Gear Viewer"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   630
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4020
      Width           =   4035
   End
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
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl skillPic 
      Height          =   660
      Index           =   0
      Left            =   510
      Top             =   4500
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1164
      Effects         =   "frmGearView.frx":0000
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl skillPic 
      Height          =   660
      Index           =   1
      Left            =   510
      Top             =   5220
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1164
      Effects         =   "frmGearView.frx":0018
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl skillPic 
      Height          =   660
      Index           =   2
      Left            =   510
      Top             =   5940
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1164
      Effects         =   "frmGearView.frx":0030
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
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   8
      Left            =   1290
      TabIndex        =   10
      ToolTipText     =   "Origin"
      Top             =   6870
      Width           =   2325
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
      Left            =   4020
      TabIndex        =   9
      Top             =   6390
      Width           =   1185
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
      Left            =   1230
      TabIndex        =   8
      Top             =   5970
      Width           =   2985
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   930
      Visible         =   0   'False
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
      Height          =   1395
      Index           =   1
      Left            =   1290
      TabIndex        =   6
      Top             =   4590
      Width           =   3435
      WordWrap        =   -1  'True
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
      Left            =   1560
      TabIndex        =   5
      Top             =   3660
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1590
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   1230
      TabIndex        =   2
      ToolTipText     =   "Origin"
      Top             =   570
      Width           =   2655
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImg 
      Height          =   7170
      Left            =   0
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   12647
      Effects         =   "frmGearView.frx":0048
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pic 
      Height          =   2715
      Left            =   1110
      Top             =   1140
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4789
      Effects         =   "frmGearView.frx":0060
   End
End
Attribute VB_Name = "frmGearView"
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
   Me.Hide
End Sub

Private Sub Form_Load()
   With AlphaImg
      Set .Picture = LoadPictureGDIplus(App.Path & "\pictures\GearBlank.bmp")
      .TransparentColor = 0
      .TransparentColorMode = lvicUseTransparentColor
   End With
   LoadCombo cboGear, "gear", gearFilter
   If cboGear.ListCount > 0 Then
      cboGear.ListIndex = 0
   End If

End Sub

Private Sub lbl_DblClick(Index As Integer)
   If gearFilter = " Order By GearName" Then
      If lbl(Index).Tag <> "" Then
         LoadCombo cboGear, "gear", " WHERE " & lbl(Index).Tag
      Else
         LoadCombo cboGear, "gear", gearFilter
      End If
      If cboGear.ListCount > 0 Then
         cboGear.ListIndex = 0
      End If
   End If
End Sub

Private Sub refreshGear(ByVal CardID)
Dim rst As New ADODB.Recordset, SQL, x, y

   SQL = "SELECT Gear.*, SupplyDeck.CardID, SupplyDeck.SupplyID, Supply.Colour, Supply.SupplyName, PlayerSupplies.PlayerID, Players.Name "
   SQL = SQL & "FROM Players RIGHT JOIN (PlayerSupplies RIGHT JOIN (Supply RIGHT JOIN (Gear LEFT JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID) ON Supply.SupplyID = SupplyDeck.SupplyID) ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Players.PlayerID = PlayerSupplies.PlayerID "
   SQL = SQL & "WHERE SupplyDeck.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      cboGear.ToolTipText = rst!GearName
      lbl(0) = rst!GearDescr
      lbl(0).ToolTipText = rst!GearDescr
      lbl(1) = rst!GearDescr
      lbl(2).Visible = (Nz(rst!Name) <> "")
      If Nz(rst!Name) <> "" Then
         lbl(2) = "held by: " & rst!Name
      End If
      lbl(3) = ""
      'lbl(4) = Trim(IIf(rst!fight >= 1, rst!fight & " Fight  ", "") & IIf(rst!tech >= 1, rst!tech & " Tech  ", "") & IIf(rst!Negotiate >= 1, rst!Negotiate & " Negotiate", ""))
      'lbl(4).Visible = (lbl(4) <> "")
      
      y = 0
      For x = 1 To rst!fight
         Set skillPic(y).Picture = LoadPictureGDIplus(App.Path & "\pictures\fight.bmp")
         skillPic(y).TransparentColor = 0
         skillPic(y).TransparentColorMode = lvicUseTransparentColor
         y = y + 1
      Next x
      For x = 1 To rst!tech
         Set skillPic(y).Picture = LoadPictureGDIplus(App.Path & "\pictures\tech.bmp")
         skillPic(y).TransparentColor = 0
         skillPic(y).TransparentColorMode = lvicUseTransparentColor
         y = y + 1
      Next x
      For x = 1 To rst!Negotiate
         Set skillPic(y).Picture = LoadPictureGDIplus(App.Path & "\pictures\nego.bmp")
         skillPic(y).TransparentColor = 0
         skillPic(y).TransparentColorMode = lvicUseTransparentColor
         y = y + 1
      Next x
      If y < 3 Then
         For x = y To 2
             Set skillPic(x).Picture = Nothing
         Next x
      End If
      
      lbl(5) = Nz(rst!KeyWords)
      If lbl(5) <> "" Then lbl(5).Tag = "Gear.KeyWords = '" & Nz(rst!KeyWords) & "'"
      
      If IsNull(rst!KeyWords) Then
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
      
      lbl(8) = "CardID: " & rst!CardID & "    GearID: " & rst!GearID
            
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

