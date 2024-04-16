VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmCrewSel 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Crew Selector"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   5970
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
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   585
   End
   Begin VB.ComboBox cboCrew 
      Appearance      =   0  'Flat
      BackColor       =   &H00133C4A&
      BeginProperty Font 
         Name            =   "Cyberpunk Is Not Dead"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008AC9E1&
      Height          =   420
      Left            =   1500
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   3255
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl disgruntledPic 
      Height          =   750
      Left            =   2610
      Top             =   3330
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      Image           =   "frmCrews.frx":0000
      Trans           =   83886080
      Effects         =   "frmCrews.frx":1E06
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl skillPic 
      Height          =   660
      Index           =   2
      Left            =   780
      Top             =   6560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1164
      Effects         =   "frmCrews.frx":1E1E
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl skillPic 
      Height          =   660
      Index           =   1
      Left            =   780
      Top             =   5700
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1164
      Effects         =   "frmCrews.frx":1E36
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl skillPic 
      Height          =   660
      Index           =   0
      Left            =   780
      Top             =   4880
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1164
      Effects         =   "frmCrews.frx":1E4E
   End
   Begin VB.Label lblLeader2 
      Appearance      =   0  'Flat
      BackColor       =   &H00288C5C&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   2085
      Left            =   720
      TabIndex        =   12
      Top             =   1260
      Width           =   285
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLeader 
      Appearance      =   0  'Flat
      BackColor       =   &H00288C5C&
      Caption         =   "L E A D E R"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   2085
      Left            =   4980
      TabIndex        =   11
      Top             =   1260
      Width           =   285
      WordWrap        =   -1  'True
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
      Left            =   1920
      TabIndex        =   10
      ToolTipText     =   "Origin"
      Top             =   7830
      Width           =   2295
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   540
      TabIndex        =   9
      ToolTipText     =   "Origin"
      Top             =   7800
      Width           =   1305
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Index           =   6
      Left            =   4620
      TabIndex        =   8
      Top             =   7260
      Width           =   1305
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
      ForeColor       =   &H0000FF00&
      Height          =   675
      Index           =   5
      Left            =   2670
      TabIndex        =   7
      Top             =   6870
      Width           =   1635
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H0000FF00&
      Height          =   1545
      Index           =   4
      Left            =   60
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   915
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H003863D0&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   525
      Index           =   3
      Left            =   870
      TabIndex        =   5
      Top             =   3570
      Width           =   1275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00288C5C&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   525
      Index           =   2
      Left            =   3840
      TabIndex        =   4
      Top             =   3570
      Width           =   1245
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Index           =   1
      Left            =   2310
      TabIndex        =   3
      Top             =   5250
      Width           =   2715
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H008AC9E1&
      Height          =   525
      Index           =   0
      Left            =   2490
      TabIndex        =   2
      Top             =   4740
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImg 
      Height          =   8190
      Left            =   0
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   14446
      Effects         =   "frmCrews.frx":1E66
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl pic 
      Height          =   3555
      Left            =   1560
      Top             =   540
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   6271
      Effects         =   "frmCrews.frx":1E7E
   End
End
Attribute VB_Name = "frmCrewSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public crewFilter As String

'For use with USER32 Function SetWindowPos
Private Const HWND_TOPMOST = -&H1
Private Const HWND_NOTOPMOST = -&H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
'For use with USER32 Function SendMessage
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
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


Private Sub AlphaImg_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub cboCrew_Click()

   If cboCrew.ListIndex = -1 Then Exit Sub
   
   refreshCrew GetCombo(cboCrew)
      
End Sub

Private Sub cmd_Click()
   playsnd 8
   If cboCrew.ListIndex = -1 Then Exit Sub
      
'   If player.PlayName <> "" And actionSeq = 0 And AlwaysOnTop = False Then
'      PutMsg player.PlayName & " has chosen " & cboCrew.Text, player.ID
'   End If
   
   Me.hide
End Sub

Private Sub Form_Load()
   With AlphaImg
      Set .Picture = LoadPictureGDIplus(App.Path & "\pictures\CrewTemplate.bmp")
      .TransparentColor = 0
      .TransparentColorMode = lvicUseTransparentColor
   End With
   LoadCombo cboCrew, "crew", crewFilter
   If cboCrew.ListCount > 0 Then
      cboCrew.ListIndex = 0
   End If

End Sub

Private Sub lbl_DblClick(Index As Integer)
   If crewFilter = " Order By CrewName" Then
      If lbl(Index).Tag <> "" Then
         LoadCombo cboCrew, "crew", " INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID WHERE " & lbl(Index).Tag
      Else
         LoadCombo cboCrew, "crew", crewFilter
      End If
      If cboCrew.ListCount > 0 Then
         cboCrew.ListIndex = 0
      End If
   End If
End Sub

Private Sub refreshCrew(ByVal CrewID)
Dim rst As New ADODB.Recordset, SQL, x, Y
   SQL = "SELECT Crew.*, Perk.PerkDescription, SupplyDeck.CardID, SupplyDeck.SupplyID, Supply.Colour, Supply.SupplyName FROM Supply RIGHT JOIN ((Perk RIGHT JOIN Crew ON Perk.PerkID = Crew.PerkID) LEFT JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON Supply.SupplyID = SupplyDeck.SupplyID WHERE Crew.CrewID=" & CrewID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      lbl(0) = rst!CrewDescr
      lbl(1) = Nz(rst!PerkDescription)
      lbl(2) = Trim(IIf(rst!Mechanic = 1, "Mechanic  ", "") & IIf(rst!Pilot = 1, "Pilot  ", "") & IIf(rst!Companion = 1, "Companion  ", "") & _
               IIf(rst!Merc = 1, "Merc  ", "") & IIf(rst!Soldier = 1, "Soldier  ", "") & IIf(rst!HillFolk = 1, "HillFolk  ", "") & _
               IIf(rst!Grifter = 1, "Grifter ", "") & IIf(rst!Medic = 1, "Medic ", "") & IIf(rst!Mudder = 1, "Mudder ", "") & IIf(rst!Lawman = 1, "Lawman", ""))
      lbl(2).Visible = (Len(lbl(2)) > 0)
      lbl(2).Tag = IIf(rst!Mechanic = 1, "Crew.Mechanic = 1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Pilot = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Pilot=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Companion = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Companion=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Merc = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Merc=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Soldier = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Soldier=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!HillFolk = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.HillFolk=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Grifter = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Grifter=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Medic = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Medic=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Mudder = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Mudder=1", "")
      lbl(2).Tag = lbl(2).Tag & IIf(rst!Lawman = 1, IIf(lbl(2).Tag = "", "", " AND ") & "Crew.Lawman=1", "")
           
      lbl(3) = IIf(rst!Moral = 1, "Moral    ", "") & IIf(rst!wanted > 0, "Wanted ", "")
      lbl(3).Visible = (rst!Moral = 1 Or rst!wanted > 0)
      lbl(3).Tag = IIf(rst!Moral = 1, "Crew.Moral = 1", "")
      lbl(3).Tag = lbl(3).Tag & IIf(rst!wanted = 1, IIf(lbl(3).Tag = "", "", " AND ") & "Crew.Wanted=1", "")

'         lbl(3).BackColor = &HC0FFC0
'      ElseIf rst!Wanted > 0 Then
'         lbl(3).BackColor = &HC0C0FF
'      Else
'         lbl(3).BackColor = 13361645
'      End If
      'lbl(4) = Trim(IIf(rst!fight >= 1, rst!fight & " Fight  ", "") & IIf(rst!tech >= 1, rst!tech & " Tech  ", "") & IIf(rst!Negotiate >= 1, rst!Negotiate & " Negotiate", ""))
      Y = 0
      For x = 1 To rst!fight
         Set skillPic(Y).Picture = LoadPictureGDIplus(App.Path & "\pictures\fight.bmp")
         skillPic(Y).TransparentColor = 0
         skillPic(Y).TransparentColorMode = lvicUseTransparentColor
         Y = Y + 1
      Next x
      For x = 1 To rst!tech
         Set skillPic(Y).Picture = LoadPictureGDIplus(App.Path & "\pictures\tech.bmp")
         skillPic(Y).TransparentColor = 0
         skillPic(Y).TransparentColorMode = lvicUseTransparentColor
         Y = Y + 1
      Next x
      For x = 1 To rst!Negotiate
         Set skillPic(Y).Picture = LoadPictureGDIplus(App.Path & "\pictures\nego.bmp")
         skillPic(Y).TransparentColor = 0
         skillPic(Y).TransparentColorMode = lvicUseTransparentColor
         Y = Y + 1
      Next x
      If Y < 3 Then
         For x = Y To 2
             Set skillPic(x).Picture = Nothing
         Next x
      End If
      
      lbl(5) = Nz(rst!KeyWords)
      If IsNull(rst!KeyWords) Then
         lbl(5).Visible = False
      Else
         lbl(5).Visible = True
         lbl(5).BackColor = 12574908
      End If
      
      lbl(6) = IIf(rst!leader = 1, "---", CStr(rst!pay))
      lblLeader.Visible = (rst!leader = 1)
      lblLeader2.Visible = lblLeader.Visible
      If rst!leader = 1 Then
         cboCrew.ForeColor = &H80FF80
         cboCrew.BackColor = &H288C5C
         
      Else
         cboCrew.ForeColor = &H8AC9E1
         cboCrew.BackColor = &H133C4A
      End If
      
      lbl(7) = rst!SupplyName
      lbl(7).BackColor = Nz(rst!Colour, &HC0C0&)
      lbl(7).Tag = "SupplyDeck.SupplyID=" & Nz(rst!SupplyID, 0)
      
      lbl(8) = "CardID: " & Nz(rst!CardID, 0) & "   CrewID: " & rst!CrewID
      
      disgruntledPic.Visible = (rst!Disgruntled > 0)
      'lbl(9) = IIf(rst!Disgruntled > 0, "Disgruntled ", "")
'      If rst!Disgruntled > 0 Then
'         lbl(9).Visible = True
'         'lbl(9).BackColor = &HC0C0FF
'      Else
'         lbl(9).Visible = False
'      End If
      
      AlphaImg.TransparentColor = 0
      AlphaImg.TransparentColorMode = lvicUseTransparentColor

      'If IsNull(rst!Picture) Then
      '   Set pic.Picture = Nothing
      'Else
         Set pic.Picture = LoadPictureGDIplus(App.Path & "\pictures\" & rst!Picture)
      'End If
   End If


End Sub


