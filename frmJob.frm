VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmJob 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "2x"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   5790
      Top             =   4500
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
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   345
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Card ID"
      ForeColor       =   &H00808080&
      Height          =   345
      Index           =   15
      Left            =   3870
      TabIndex        =   16
      Top             =   5610
      Width           =   1275
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   600
      Index           =   11
      Left            =   3240
      Top             =   5340
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1058
      Effects         =   "frmJob.frx":0000
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   600
      Index           =   7
      Left            =   2070
      Top             =   870
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1058
      Effects         =   "frmJob.frx":0018
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   7850
      TabIndex        =   15
      Top             =   1470
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   7090
      TabIndex        =   14
      Top             =   1470
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   6330
      TabIndex        =   13
      Top             =   1470
      Visible         =   0   'False
      Width           =   300
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   435
      Index           =   10
      Left            =   8100
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      Attr            =   516
      FixedCx         =   31
      FixedCy         =   25
      Effects         =   "frmJob.frx":0030
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   435
      Index           =   9
      Left            =   7350
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      Attr            =   516
      FixedCx         =   31
      FixedCy         =   25
      Effects         =   "frmJob.frx":0048
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   435
      Index           =   8
      Left            =   6600
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      Attr            =   516
      FixedCx         =   31
      FixedCy         =   25
      Effects         =   "frmJob.frx":0060
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "invalid Job"
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
      Height          =   885
      Index           =   11
      Left            =   150
      TabIndex        =   12
      Top             =   4170
      Visible         =   0   'False
      Width           =   6075
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "invalid Job"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003B80B4&
      Height          =   600
      Index           =   10
      Left            =   150
      TabIndex        =   11
      Top             =   3550
      Visible         =   0   'False
      Width           =   2805
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drop Off"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084B7DF&
      Height          =   495
      Index           =   9
      Left            =   150
      TabIndex        =   10
      Top             =   3060
      Visible         =   0   'False
      Width           =   3285
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   570
      Index           =   6
      Left            =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   1005
      Effects         =   "frmJob.frx":0078
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "invalid Job"
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
      Height          =   885
      Index           =   8
      Left            =   150
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   6075
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Needs"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3105
      Index           =   7
      Left            =   6690
      TabIndex        =   8
      Top             =   2100
      Visible         =   0   'False
      Width           =   1845
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   3105
      Index           =   5
      Left            =   6300
      Top             =   1830
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5477
      Effects         =   "frmJob.frx":0090
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   540
      Index           =   4
      Left            =   6900
      Top             =   780
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   953
      Effects         =   "frmJob.frx":00A8
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   5
      Left            =   90
      TabIndex        =   7
      Top             =   5480
      Visible         =   0   'False
      Width           =   3735
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   600
      Index           =   3
      Left            =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   1058
      Effects         =   "frmJob.frx":00C0
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HEAVY LOAD"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Index           =   4
      Left            =   3315
      TabIndex        =   6
      Top             =   1440
      Width           =   3300
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "invalid Job"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003B80B4&
      Height          =   600
      Index           =   3
      Left            =   150
      TabIndex        =   5
      Top             =   1450
      Width           =   2805
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   750
      Index           =   2
      Left            =   6390
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1323
      Effects         =   "frmJob.frx":00D8
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Transport"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00163B5F&
      Height          =   495
      Index           =   2
      Left            =   3225
      TabIndex        =   4
      Top             =   960
      Width           =   3510
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pickup"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084B7DF&
      Height          =   495
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   960
      Width           =   3285
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
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
      Height          =   945
      Index           =   6
      Left            =   5970
      TabIndex        =   2
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "invalid Job"
      BeginProperty Font 
         Name            =   "Cyberpunk Is Not Dead"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A6C5CD&
      Height          =   705
      Index           =   0
      Left            =   1125
      TabIndex        =   1
      Top             =   100
      Width           =   5205
      WordWrap        =   -1  'True
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   270
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   873
      Attr            =   640
      Effects         =   "frmJob.frx":00F0
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl img 
      Height          =   6000
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   10583
      Effects         =   "frmJob.frx":0108
   End
End
Attribute VB_Name = "frmJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public JobCardID As Integer

Private LastSectorID As Integer
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
   'Me.hide
End Sub

Private Sub Form_Load()
Dim rgn As Long
Dim w As Long, h As Long

   ' Convert twips ? pixels
   w = Me.ScaleX(Me.Width, vbTwips, vbPixels)
   h = Me.ScaleY(Me.Height, vbTwips, vbPixels)
   
   ' 40,40 = corner roundness
   rgn = CreateRoundRectRgn(0, 0, w, h, 20, 20)
   
   SetWindowRgn Me.hwnd, rgn, True
    
   RefreshJob
End Sub

Private Sub img_DblClick(Index As Integer)
   Unload Me
End Sub

Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lbl_Click(Index As Integer)
   If Val(lbl(3).Tag) > 0 And Index = 3 Then
      Main.drawLine 0, Val(lbl(3).Tag), varDLookup("SectorID", "Players", "PlayerID=" & player.ID), False
   End If
   
   If Val(lbl(10).Tag) > 0 And Index = 10 Then
      Main.drawLine 0, Val(lbl(10).Tag), varDLookup("SectorID", "Players", "PlayerID=" & player.ID), False
   End If
End Sub

Private Sub lbl_DblClick(Index As Integer)
   Unload Me
End Sub

Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Index <> 3 And Index <> 10 Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
   End If
End Sub

Private Sub RefreshJob()
Dim Index, SQL, showBounty As Boolean
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim x

      If JobCardID < 1 Then
         Exit Sub
      End If
      SQL = "SELECT Contact.Picture, JobType.JobTypeDescr, Profession.ProfessionName, ContactDeck.*, JobType_1.JobTypeDescr AS JobType2 "
      SQL = SQL & "FROM (Contact INNER JOIN ((ContactDeck INNER JOIN JobType ON ContactDeck.JobTypeID = JobType.JobTypeID) LEFT JOIN Profession "
      SQL = SQL & "ON ContactDeck.ProfessionID = Profession.ProfessionID) ON Contact.ContactID = ContactDeck.ContactID) INNER JOIN JobType AS JobType_1 ON ContactDeck.JobType2D = JobType_1.JobTypeID "
      SQL = SQL & " WHERE ContactDeck.CardID = " & JobCardID
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         If rst!ContactID = 5 Then 'Alliance Background
            Set img(0).Picture = LoadPictureGDIplus(App.Path & "\gui-job\JobBackAlliance.jpg")
         Else
            Set img(0).Picture = LoadPictureGDIplus(App.Path & "\gui-job\JobBackStd.jpg")
         End If
         
         Set img(1).Picture = LoadPictureGDIplus(App.Path & "\gui-job\Contact" & rst!ContactID & ".bmp")
         img(1).TransparentColor = 0
         img(1).TransparentColorMode = lvicUseTransparentColor
         
         Set img(2).Picture = LoadPictureGDIplus(App.Path & "\gui-job\illegal" & rst!illegal & ".bmp")
         img(2).TransparentColor = 0
         img(2).TransparentColorMode = lvicUseTransparentColor
         
         If rst!BonusPart > 0 Or rst!bonus > 0 Then
            img(3).Visible = True
            Set img(3).Picture = LoadPictureGDIplus(App.Path & "\gui-job\bonus.bmp")
            img(3).TransparentColor = 0
            img(3).TransparentColorMode = lvicUseTransparentColor
            lbl(5).Visible = True
            lbl(5).Caption = IIf(rst!BonusPart > 0, " +" & rst!BonusPart & " part: ", "") & IIf(rst!bonus > 0, " +$" & rst!bonus & ":", "") & IIf(rst!KeywordBonus = 1, rst!KeyWords, "") & IIf(IsNull(rst!ProfessionName), "", " " & rst!ProfessionName) & IIf(rst!BonusPerSkill > 0, " /" & cstrSkill(rst!BonusPerSkill), "") & IIf(rst!Job3ID > 0, "Bonus Job", "")
         End If
         
         If rst!Immoral = 1 Then
            img(4).Visible = True
            Set img(4).Picture = LoadPictureGDIplus(App.Path & "\gui-job\immoral.bmp")
            img(4).TransparentColor = 0
            img(4).TransparentColorMode = lvicUseTransparentColor
         End If
         
         If Not IsNull(rst!JobOrder) Then
            img(5).Visible = True
            Set img(5).Picture = LoadPictureGDIplus(App.Path & "\gui-job\sidebar.bmp")
            img(5).TransparentColor = 0
            img(5).TransparentColorMode = lvicUseTransparentColor
            img(5).TransparencyPct = 30
            lbl(7).Visible = True
            lbl(7).Caption = Replace(Replace(Replace(Replace(rst!JobOrder, ".", vbNewLine), ",", vbNewLine), "/", vbNewLine), "&", "&&")
         End If
         
         lbl(0).Caption = rst!JobName
         x = rst!pay
         If x = 0 And rst!ContactID = 0 Then
            lbl(6).Caption = "Goal"
         ElseIf x = 0 Then
            lbl(6).Caption = "Special"
         Else
            lbl(6).Caption = "$" & x
         End If
         lbl(2).Caption = rst!JobTypeDescr & IIf(rst!JobType2 <> "-", "/" & rst!JobType2, "")  ' & IIf(rst!illegal = 1, "/illegal", "") & IIf(rst!Immoral = 1, "/immoral", "")
         lbl(4).Visible = (rst!ExtraFuel = 1)
         
         If rst!fight > 0 Then
            Set img(8).Picture = LoadPictureGDIplus(App.Path & "\pictures\fight.bmp")
            img(8).TransparentColor = 0
            img(8).TransparentColorMode = lvicUseTransparentColor
            lbl(12).Visible = True
            lbl(12) = rst!fight
         End If
         If rst!tech > 0 Then
            Set img(9).Picture = LoadPictureGDIplus(App.Path & "\pictures\tech.bmp")
            img(9).TransparentColor = 0
            img(9).TransparentColorMode = lvicUseTransparentColor
            lbl(13).Visible = True
            lbl(13) = rst!tech
         End If
         If rst!Negotiate > 0 Then
            Set img(10).Picture = LoadPictureGDIplus(App.Path & "\pictures\nego.bmp")
            img(10).TransparentColor = 0
            img(10).TransparentColorMode = lvicUseTransparentColor
            lbl(14).Visible = True
            lbl(14) = rst!Negotiate
         End If

         If rst!Job1ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job LEFT JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job1ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
               x = getSectorCount(getPlayerSector(player.ID), rst2!SectorID)
               lbl(3).Caption = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "") & vbNewLine & Nz(rst2!System)
               lbl(3).ForeColor = IIf(x > 0, &H3B80B4, &HC000&)
               lbl(3).Tag = CStr(rst2!SectorID)
               lbl(8).Visible = True
               lbl(8).Caption = Replace(rst2!JobDesc, "&", "&&")

            End If
            rst2.Close
         End If

         'Bonus Drop Job
         If rst!Job3ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job3ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
               lbl(7).Caption = lbl(7).Caption & vbNewLine & rst2!JobDesc & vbNewLine & rst2!PlanetName & vbNewLine & rst2!System

            End If
            rst2.Close
         End If

         If rst!Job2ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job2ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
               img(6).Visible = True
               Set img(6).Picture = LoadPictureGDIplus(App.Path & "\gui-job\dropoff.bmp")
               img(6).TransparentColor = 0
               img(6).TransparentColorMode = lvicUseTransparentColor
               lbl(9).Visible = True
               lbl(9).Caption = "Drop Off"
               x = getSectorCount(getPlayerSector(player.ID), rst2!SectorID)
               lbl(10).Visible = True
               lbl(10).Caption = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "") & vbNewLine & Nz(rst2!System)
               lbl(10).ForeColor = IIf(x > 0, &H3B80B4, &HC000&)
               lbl(10).Tag = CStr(rst2!SectorID)
               lbl(11).Visible = True
               lbl(11).Caption = Replace(rst2!JobDesc, "&", "&&")

            End If
            rst2.Close
         End If
         lbl(15).Caption = "Card ID " & JobCardID
         If checkThisJob Then Unload Me
      End If
      LastSectorID = player.SectorID
End Sub

Private Sub Timer1_Timer()
   If LastJobDone = JobCardID Then
      LastJobDone = 0 'clear the flag
      If checkThisJob Then Unload Me
   End If
   If LastSectorID <> player.SectorID Then RefreshJob
End Sub

Private Function checkThisJob() As Boolean 'set true if completed
Dim rst As New ADODB.Recordset, status As Integer
Dim SQL
   SQL = "SELECT JobStatus FROM PlayerJobs "
   SQL = SQL & "WHERE CardID=" & JobCardID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      status = rst!JobStatus
   End If
   rst.Close
   Set rst = Nothing
      
   If status > 0 Then ' half done
      Set img(7).Picture = LoadPictureGDIplus(App.Path & "\gui-job\tick.bmp")
      img(7).TransparentColor = 0
      img(7).TransparentColorMode = lvicUseTransparentColor
   End If
   If status = 2 Then
      Set img(11).Picture = LoadPictureGDIplus(App.Path & "\gui-job\tick.bmp")
      img(11).TransparentColor = 0
      img(11).TransparentColorMode = lvicUseTransparentColor
   End If
   If status = 3 Then checkThisJob = True

End Function
