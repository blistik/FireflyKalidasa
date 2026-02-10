VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmBounty 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   3630
      Top             =   7440
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   345
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   600
      Index           =   1
      Left            =   5250
      Top             =   5340
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1058
      Effects         =   "frmBounty.frx":0000
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Card ID"
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   15
      Left            =   2000
      TabIndex        =   8
      Top             =   8300
      Width           =   1275
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Armed and Dangerous"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Index           =   2
      Left            =   2670
      TabIndex        =   7
      Top             =   3900
      Width           =   3105
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   540
      Index           =   4
      Left            =   4320
      Top             =   7040
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   953
      Effects         =   "frmBounty.frx":0018
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Attemp Botched"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Index           =   5
      Left            =   1900
      TabIndex        =   6
      Top             =   5895
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Apprehend"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   4
      Left            =   1900
      TabIndex        =   5
      Top             =   5480
      Width           =   4005
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Londinium"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003B80B4&
      Height          =   750
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   7800
      Width           =   2565
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LAST SEEN:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003B80B4&
      Height          =   1860
      Index           =   3
      Left            =   2670
      TabIndex        =   3
      Top             =   2130
      Width           =   2565
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Saffron"
      BeginProperty Font 
         Name            =   "Cyberpunk Is Not Dead"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Index           =   6
      Left            =   4050
      TabIndex        =   1
      Top             =   7650
      Width           =   1905
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   2280
      Index           =   7
      Left            =   660
      Top             =   2080
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   4022
      Attr            =   516
      FixedCx         =   192
      FixedCy         =   186
      Effects         =   "frmBounty.frx":0030
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   8595
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   15161
      Effects         =   "frmBounty.frx":0048
   End
End
Attribute VB_Name = "frmBounty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public JobCardID As Integer

Private LastSectorID As Integer, isOwner As Boolean
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

Private Sub cmd_Click()
   playsnd 8
   Unload Me
End Sub

Private Sub Form_Load()
Dim rgn As Long
Dim w As Long, h As Long

   ' Convert twips ? pixels
   w = Me.ScaleX(Me.Width, vbTwips, vbPixels)
   h = Me.ScaleY(Me.Height, vbTwips, vbPixels)
   
   ' 40,40 = corner roundness
   rgn = CreateRoundRectRgn(0, 0, w, h, 24, 24)
   
   SetWindowRgn Me.hwnd, rgn, True
   
   RefreshJob
   Set img(1).Picture = LoadPictureGDIplus(App.Path & "\gui-job\tick.bmp")
   img(1).TransparentColor = 0
   img(1).TransparentColorMode = lvicUseTransparentColor
End Sub

Private Sub img_DblClick(Index As Integer)
   Unload Me
End Sub

Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_Click(Index As Integer)
   If Val(lbl(1).Tag) > 0 And Index = 1 Then
      Main.drawLine IIf(isOwner, 0, 1), Val(lbl(1).Tag), varDLookup("SectorID", "Players", "PlayerID=" & player.ID), False
   End If
   
   If Val(lbl(3).Tag) > 0 And Index = 3 Then
      Main.drawLine IIf(isOwner, 0, 1), Val(lbl(3).Tag), varDLookup("SectorID", "Players", "PlayerID=" & player.ID), False
   End If
End Sub

Private Sub lbl_DblClick(Index As Integer)
   Unload Me
End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Index <> 3 And Index <> 1 Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
   End If
End Sub

Private Sub RefreshJob()
Dim Index, SQL, showBounty As Boolean
Dim rst As New ADODB.Recordset
Dim x
      SQL = "SELECT ContactDeck.*, Job.JobDesc, Job.SectorID, Crew.CrewName, Crew.Picture, Planet.PlanetName, Planet.System, Planet_1.SectorID AS LastSector, Planet_1.PlanetName AS LastPlanet, Planet_1.System AS LastSystem"
      SQL = SQL & " FROM (Supply INNER JOIN ((((Job INNER JOIN ContactDeck ON Job.JobID = ContactDeck.Job1ID) INNER JOIN Crew ON ContactDeck.FugitiveID = Crew.CrewID) INNER JOIN Planet ON Job.SectorID = Planet.SectorID)"
      SQL = SQL & " INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON Supply.SupplyID = SupplyDeck.SupplyID) INNER JOIN Planet AS Planet_1 ON Supply.SectorID = Planet_1.SectorID"
      SQL = SQL & " Where ContactDeck.CardID = " & JobCardID
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         Set img(0).Picture = LoadPictureGDIplus(App.Path & "\gui-job\JobBackBounty.jpg")
         If rst!FugitiveID > 0 Then
            Set img(7).Picture = LoadPictureGDIplus(App.Path & "\pictures\" & rst!Picture)
         End If
         lbl(0).Caption = rst!CrewName
         x = getSectorCount(getPlayerSector(player.ID), rst!sectorID)
         lbl(1).Caption = rst!PlanetName & IIf(x > 0, "  (" & x & ")", "") & vbNewLine & Nz(rst!System)
         lbl(1).Tag = CStr(rst!sectorID) 'store destination sector
         lbl(1).ForeColor = IIf(x > 0, &H3B80B4, &HC000&)
         x = haveCrewWho(rst!FugitiveID)
         If x = vbNullString Then
            lbl(3).Caption = "LAST SEEN:" & vbNewLine & rst!LastPlanet & vbNewLine & rst!LastSystem
            lbl(3).Tag = CStr(rst!LastSector)
         Else
            lbl(3).Caption = "LAST SEEN:" & vbNewLine & "Firefly class" & vbNewLine & Chr(34) & x & Chr(34)
            lbl(3).Tag = varDLookup("SectorID", "Players", "Ship='" & x & "'") 'pickup sector
         End If
         
         lbl(4).Caption = "Apprehend " & rst!CrewName
         If rst!FailKillCrew = 0 Then
            lbl(5).Caption = "Attempt Botched"
         Else
            lbl(5).Caption = "Kill " & rst!FailKillCrew & " of Attacker's Crew" & vbNewLine & "Attempt Botched"
         End If
         lbl(6).Caption = rst!pay
         
         If rst!Immoral = 1 Then
            img(4).Visible = True
            Set img(4).Picture = LoadPictureGDIplus(App.Path & "\gui-job\immoral.bmp")
            img(4).TransparentColor = 0
            img(4).TransparentColorMode = lvicUseTransparentColor
         End If
         lbl(15).Caption = "Card ID " & JobCardID
         
         
      End If
      If checkThisJob Then Unload Me
      LastSectorID = player.sectorID

End Sub

Private Sub Timer1_Timer()
   If LastJobDone = JobCardID Then
      LastJobDone = 0 'clear the flag
      If checkThisJob Then Unload Me
   End If
   If LastSectorID <> player.sectorID Then RefreshJob

End Sub

Private Function checkThisJob() As Boolean 'set true if completed
Dim rst As New ADODB.Recordset, status As Integer, x
Dim SQL
   SQL = "SELECT PlayerID, JobStatus FROM PlayerJobs "
   SQL = SQL & "WHERE CardID=" & JobCardID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      status = rst!JobStatus
      x = rst!playerID
      isOwner = (x = player.ID)
      If x = player.ID Then
         img(1).Visible = True
         lbl(3).Tag = ""
      Else 'another player has the bounty
         img(1).Visible = False
         lbl(3).Tag = varDLookup("SectorID", "Players", "PlayerID=" & x) 'pickup sector
      End If
      lbl(3).Caption = "LAST SEEN:" & vbNewLine & "Firefly class" & vbNewLine & Chr(34) & varDLookup("Ship", "Players", "PlayerID=" & x) & Chr(34)

   Else
      img(1).Visible = False
   End If
   
   rst.Close
   Set rst = Nothing
   
   If status = 3 Then checkThisJob = True

End Function
