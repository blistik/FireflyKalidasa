VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Main 
   BackColor       =   &H80000006&
   Caption         =   "Firefly - The PC Game"
   ClientHeight    =   12375
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20070
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "Main.frx":030A
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timing 
      Enabled         =   0   'False
      Interval        =   3200
      Left            =   1830
      Top             =   1800
   End
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   12120
      Width           =   20070
      _ExtentX        =   35401
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   34880
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   1110
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   33
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4093C
            Key             =   "ship"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":40D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":411E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":41638
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":41CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":42104
            Key             =   "grap"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":42556
            Key             =   "crew"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":431A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":435FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":43A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":43E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":441B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":444D2
            Key             =   "grapz"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":447EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":44B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":44E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4513A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4558C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":459DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":45CF8
            Key             =   "start"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":46DD2
            Key             =   "join"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":47EAC
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":48F86
            Key             =   "chat"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4A060
            Key             =   "graph"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4B13A
            Key             =   "log"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4C214
            Key             =   "crewz"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4D2EE
            Key             =   "hat"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4E3C8
            Key             =   "upgrd"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4F4A2
            Key             =   "serenity"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5057C
            Key             =   "job"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":51656
            Key             =   "deal"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":52730
            Key             =   "cash"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":5380A
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20070
      _ExtentX        =   35401
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "Images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "start"
            Object.ToolTipText     =   "Host"
            ImageKey        =   "start"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "join"
            Object.ToolTipText     =   "Join"
            ImageKey        =   "join"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "exit"
            Object.ToolTipText     =   "End Game"
            ImageKey        =   "exit"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "chat"
            Object.ToolTipText     =   "Chat"
            ImageKey        =   "chat"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "graph"
            Object.ToolTipText     =   "Game Info"
            ImageKey        =   "graph"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "log"
            Object.ToolTipText     =   "Game Log"
            ImageKey        =   "log"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "crew"
            Object.ToolTipText     =   "Crew Browser"
            ImageKey        =   "crewz"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "gear"
            Object.ToolTipText     =   "Gear Browser"
            ImageKey        =   "hat"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "upgrd"
            Object.ToolTipText     =   "Ship Upgrades Browser"
            ImageKey        =   "upgrd"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ship"
            Object.ToolTipText     =   "Ship Browser"
            ImageKey        =   "serenity"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "allships"
                  Text            =   "All Ships"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "myship"
                  Text            =   "My Ship"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "job"
            Object.ToolTipText     =   "Job Browser"
            ImageKey        =   "job"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "alljobs"
                  Text            =   "All Jobs"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "myjobs"
                  Text            =   "My Jobs"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deal"
            Object.ToolTipText     =   "Deal Browser"
            ImageKey        =   "deal"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "alldeals"
                  Text            =   "All Deals"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "localdeals"
                  Text            =   "Local Deals"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "buy"
            Object.ToolTipText     =   "Buy Browser"
            ImageKey        =   "cash"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "allbuys"
                  Text            =   "All Buys"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "localbuys"
                  Text            =   "Local Buys"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Game Rules"
            ImageKey        =   "help"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   10
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "firefly"
                  Text            =   "Firefly Rulebook"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bluesun"
                  Text            =   "Blue Sun Rulebook"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kalidasa"
                  Text            =   "Kalidasa Rulebook"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pbh"
                  Text            =   "Pirates && Bounty Hunters Rulebook"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pcguide"
                  Text            =   "Firefly for PC Guide"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "jobs"
                  Text            =   "Job View/Edit"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bot"
                  Text            =   "start an AI Player Bot"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "map"
                  Text            =   "edit Map"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "check"
                  Text            =   "Check latest Release"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "about"
                  Text            =   "About"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      Begin VB.PictureBox picMB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   14335
         Picture         =   "Main.frx":548E4
         ScaleHeight     =   285
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   13930
         Picture         =   "Main.frx":54CB6
         ScaleHeight     =   285
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   13525
         Picture         =   "Main.frx":55088
         ScaleHeight     =   285
         ScaleWidth      =   240
         TabIndex        =   12
         Top             =   160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   13120
         Picture         =   "Main.frx":5545A
         ScaleHeight     =   285
         ScaleWidth      =   240
         TabIndex        =   11
         Top             =   160
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   8
         Left            =   12340
         Picture         =   "Main.frx":5582C
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   10
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   7
         Left            =   11960
         Picture         =   "Main.frx":55E07
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   9
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   6
         Left            =   11580
         Picture         =   "Main.frx":5639B
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   8
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   5
         Left            =   11200
         Picture         =   "Main.frx":569A3
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   7
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   4
         Left            =   10820
         Picture         =   "Main.frx":56BB7
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   6
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   3
         Left            =   10440
         Picture         =   "Main.frx":57118
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   5
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   2
         Left            =   10060
         Picture         =   "Main.frx":576C8
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   4
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   9680
         Picture         =   "Main.frx":57C9C
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   3
         Top             =   90
         Width           =   360
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   9
         Left            =   12720
         Picture         =   "Main.frx":5828A
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   2
         Top             =   90
         Width           =   360
      End
   End
   Begin XDOCKFLOATLibCtl.DockFrame DockFrame1 
      Left            =   570
      Top             =   1770
      _cx             =   688
      _cy             =   688
      DragAreaStyle   =   0
      PICTCNT         =   0
      MENUCNT         =   0
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Verse As Board
Attribute Verse.VB_VarHelpID = -1
Private usedStitchSkill As Boolean, dontnagme As Boolean, rejoininprogress As Boolean
Public frmJob As frmJobs, frmShip As frmShips, frmDeal As frmDeals, frmBuy As frmSupply, frmStat As frmStats, frmSkill As frmSkillSel

Private Sub MDIForm_Load()
Dim x
   
   PlayCode(1).Color = "Orange"
   PlayCode(2).Color = "Blue"
   PlayCode(3).Color = "Yellow"
   PlayCode(4).Color = "Green"
   pickStartSector = -1
   actionSeq = ASidle

   DockFrame1.LoadStates "Firefly"
   
   Set Verse = New Board
   
   initToolbar False

   If Not Logon Then End
   
   For x = 1 To NO_OF_CONTACTS
      pic(x).Visible = False
   Next x
   
   Logic.Open "GameSeq", DB, adOpenDynamic, adLockPessimistic  ', adLockOptimistic
   x = GetSeq
   Timing.Enabled = True
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   DockFrame1.SaveStates "Firefly"
   If Me.Visible Then
      If MessBox("Are you sure you want to close the game?", "Closing Game", "I'm outta here", "Nope") = 0 Then
         If DB.State = adStateOpen Then DB.Close
         End
      Else
         Cancel = True
      End If
   End If

End Sub

'THE MAIN ENGINE of the GAME
' Game States E - Idle/End, H - Host screen, 1-4 players go.
' S - setup Game
' R - run Game, T-Trade
' F - Boarding alert for defender, init Showdown
' U - capture the Move Corvette to any planetary sector
' W - Reaver to any Rim or Border sector, X-Move a Reaver 1 sector, Y=Move the Cruiser 1 sector, Z- move the Cruiser adjacent player, V-move Corvette Adjacent player
' actionSeq States = ASidle , ASselect --- >>> , ASend, -> ASidle, <repeat>
Private Sub Timing_Timer()
Dim status As Variant, errh, thisplayer As Integer
Dim sectorID, ContactID As Integer, SupplyID As Integer, x, y As Integer
Dim maxConsider
On Error GoTo err_handler

   sectorID = getPlayerSector(player.ID)
   ContactID = Nz(varDLookup("ContactID", "Contact", "SectorID=" & sectorID), 0)
   ContactID = IIf(HigginsDealPerk, 8, IIf(HarkenDeal, 5, ContactID))
   SupplyID = Nz(varDLookup("SupplyID", "Supply", "SectorID=" & sectorID), 0)

   status = GetSeqX(thisplayer)
   
   If status = "R" And player.ID = 0 And Not dontnagme Then
      dontnagme = Not reJoin
   ElseIf status = "L" And player.ID = 0 Then
      If MsgBox("Game in Leader/Crew selection, do you need to reset it?", vbYesNo + vbCritical, "Unexpected Game State") = vbYes Then
         DB.Execute "UPDATE GameSeq SET Seq = 'E'"
      End If

   End If
   
   'aminmate the current player
   If status = "R" And player.ID > 0 Then animatePlayer thisplayer

   If status <> "H" And status <> "E" And status <> "L" And pickStartSector > -1 Then
      RefreshBoard
      refreshSolid
   End If
   If status = "E" Then 'currently in End Game
      PutMsg "Waiting to Host or Join a Game"
   ElseIf status = "S" And thisplayer = player.ID And pickStartSector = 0 Then  'your go to pick starting sector on MAP
      Verse.Caption = "the 'Verse - " & varDLookup("StoryTitle", "Story", "StoryID = " & Logic!StoryID)
      NumOfReavers = varDLookup("NoOfReavers", "Story", "StoryID = " & Logic!StoryID)
      'set game ships
      For x = 5 To 6 + NumOfReavers
         MoveShip x, varDLookup("StartSectorID", "Players", "PlayerID=" & CStr(x))
      Next x
      PutMsg player.PlayName & " selecting Start Sector", player.ID, Logic!Gamecntr
      
      If useHavens(Logic!StoryID) Then
         MessBox "Click on the Planet Sector to be your Haven", "Pick your Haven", "Will do", "", getLeader()
      Else
         MessBox "Click on the Sector you want to start in", "Place your Ship", "Will do", "", getLeader()
      End If
      
      pickStartSector = 1
      
   ElseIf status = "S" And thisplayer = player.ID And pickStartSector = 2 Then  'setup
      PutMsg player.PlayName & "'s on the Map", player.ID, Logic!Gamecntr
      
      'deal start drive core, and Jobs
      dealDriveAndJobs player.ID
      
      If varDLookup("UpgradeDrive", "Story", "StoryID=" & Logic!StoryID) = 1 Then
         Dim frmShUp As frmShipUpgd
         'present list of discarded upgrades to choose one for free
         Set frmShUp = New frmShipUpgd
         frmShUp.discardMode = 6
         frmShUp.Show 1
      End If
      
      If varDLookup("Warrant", "Story", "StoryID=" & Logic!StoryID) > 0 Then
         DB.Execute "UPDATE Players Set Warrants = 1 WHERE PlayerID=" & player.ID
      End If
      
      refreshAll
      
      'starting point selected, pass to next person, or kick the main Running Game cycle off
      setNextPlayerREV player.ID, "R"
      Logic.Requery
      If Logic!Seq = "R" Then
         PutMsg PlayCode(Logic!player).PlayName & "'s Turn", Logic!player, Logic!Gamecntr
      End If
   
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASidle Then   'MAIN Cycle - init your go
      If ((getCutterSector(sectorID) > 0 Or getCruiserCorvette(sectorID) = 5) And CruiserCutter <> sectorID) Or (getCruiserCorvette(sectorID) = 6 And CorvetteSeq <> getCorvetteSeq) Then
         If checkWhisperX1(sectorID) Then
            actionSeq = ASNavEvade ' and get away
            Exit Sub
         End If
      End If

      actionSeq = ASselect 'in limbo awaiting user to select
      
      showActions
      playsnd 6
   
   ElseIf status = "F" And thisplayer <> player.ID And actionSeq = ASidle And Logic!trader = player.ID And player.ID > 0 Then 'showdown - defend!
      'initiate the skill selection
      actionSeq = ASBountySkill
      'MessBox "You have been Boarded by " & PlayCode(thisplayer).PlayName & "!!" & vbNewLine & "Prepare your Crew for a Showdown, then select a Skill once you're ready.", "SHOWDOWN", "OK", "", getLeader()
      x = showBoarded(thisplayer)
      Set frmSkill = New frmSkillSel
      If x > 0 Then
         frmSkill.skill = x
         actionSeq = ASBountySkillSel
         playsnd 8
      Else
         frmSkill.setMode 2
         frmSkill.AlwaysOnTop = True
         frmSkill.Show 0, Me
      End If
   ElseIf status = "F" And thisplayer <> player.ID And actionSeq = ASBountySkillSel And Logic!trader = player.ID Then 'showdown - defend!
      'MessBox "You have selected Skill " & frmSkill.Skill, "Skill", "OK"
      x = frmSkill.skill
      Set frmSkill = Nothing
      y = Logic!player
      'initiate the Showdown
      doShowdownDefend y, x
      actionSeq = ASidle
   
   ElseIf status = "T" And thisplayer <> player.ID And actionSeq = ASidle And Logic!trader = player.ID Then
      doSlaveTrade Logic!player
      
   ElseIf status = "U" And thisplayer = player.ID And actionSeq = ASidle Then 'capture the Move Corvette to any planetary sector
      x = setPlayer(player.ID, "", 0, True)
      MessBox "Move the Operative's Corvette to any Planetary Sector", "Place the Corvette", "OK"
      'kick it off
      actionSeq = ASNavCorvPlanetary
      
   ElseIf status = "W" And thisplayer = player.ID And actionSeq = ASidle Then 'capture the Move Reaver Cycle from another Player's Nav move
      MessBox "Move a Reaver to any Rim or Border sector", "Place a Reaver", "OK"
      'kick it off
      actionSeq = ASNavReavBorder
      
   ElseIf status = "W" And thisplayer = player.ID And actionSeq = ASNavReavEnd Then    'fullburn Cycle
      actionSeq = ASidle
      'turn finished, push to next player (for SP thats you)
      PutMsg player.PlayName & " 'baited' the Reaver Cutter", thisplayer, Logic!Gamecntr
      'change back
      thisplayer = setPlayer(player.ID, "R", 0)
      If thisplayer <> player.ID Then
         PutMsg PlayCode(thisplayer).PlayName & "'s Turn", thisplayer, Logic!Gamecntr
      End If
      
   ElseIf status = "X" And thisplayer = player.ID And actionSeq = ASidle Then 'capture the Move Reaver Cycle from another Player's Nav move
      MessBox "Move a Reaver 1 sector", "Move a Reaver", "OK"
      'kick it off
      actionSeq = ASNavReav
      
   ElseIf status = "X" And thisplayer = player.ID And actionSeq = ASNavReavEnd Then    'fullburn Cycle
       actionSeq = ASidle
       PutMsg player.PlayName & " 'summonded' the Reaver Cutter", thisplayer, Logic!Gamecntr
      'turn finished, push to next player (for SP thats you)
      thisplayer = setPlayer(player.ID, "R", 0)
      If thisplayer <> player.ID Then
         PutMsg PlayCode(thisplayer).PlayName & "'s Turn", thisplayer, Logic!Gamecntr
      End If
      
   ElseIf status = "Z" And thisplayer = player.ID And actionSeq = ASidle Then 'capture the Move Cruiser Cycle from another Player's Nav move
      x = setPlayer(player.ID, "", 0, True)
      MessBox "Move the Alliance Cruiser adjacent to " & PlayCode(x).PlayName, "Move the Alliance Cruiser", "OK"
      'kick it off
      actionSeq = ASNavCrusAdjacent
         
   ElseIf status = "V" And thisplayer = player.ID And actionSeq = ASidle Then 'capture the Move Corvette Cycle from another Player's Nav move
      x = setPlayer(player.ID, "", 0, True)
      MessBox "Move the Operative's Corvette adjacent to " & PlayCode(x).PlayName, "Move the Operative's Corvette", "OK"
      'kick it off
      actionSeq = ASNavCorvAdjacent
         
   ElseIf status = "Y" And thisplayer = player.ID And actionSeq = ASidle Then 'capture the Move Cruiser Cycle from another Player's Nav move
      MessBox "Move the Alliance Cruiser 1 sector", "Move the Alliance Cruiser", "OK"
      'kick it off
      actionSeq = ASNavCrus
      
   ElseIf (status = "U" Or status = "V" Or status = "Y" Or status = "Z") And thisplayer = player.ID And actionSeq = ASNavCrusEnd Then      'fullburn Cycle
       actionSeq = ASidle
      'turn finished, push to next player (for SP thats you)
      thisplayer = setPlayer(player.ID, "R", 0)
      PutMsg player.PlayName & " 'directed' the " & IIf(status = "U" Or status = "V", "Operative's Corvette", "Alliance Cruiser"), thisplayer, Logic!Gamecntr
  
      If thisplayer <> player.ID Then
         PutMsg PlayCode(thisplayer).PlayName & "'s Turn", thisplayer, Logic!Gamecntr
      End If
   
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASMoseyEnd Then   'Mosey Cycle - your go
      x = getPlanetName(sectorID)
      PutMsg player.PlayName & " moseyed to " & IIf(x = "", "sector " & sectorID, x), player.ID, Logic!Gamecntr
      If resolveToken(sectorID) = 6 And isOutlaw(player.ID) Then 'no Nav card when Corvette arrives
         If actionSeq <> ASNavEvade Then
            actionSeq = ASselect 'in limbo awaiting user to select
            showActions   'throw it back to the action window
         End If
      Else 'reavers
         If actionSeq <> ASNavEvade Then
            checkFlacGun sectorID
            actionSeq = ASselect 'in limbo awaiting user to select
            showActions   'throw it back to the action window
         End If
      End If
'      resolveToken SectorID
'      checkFlacGun SectorID
'      actionSeq = ASselect 'in limbo awaiting user to select
'      showActions   'throw it back to the action window to resolve end of mosey and offer other actions
   
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASFullburnEnd Then   'fullburn Cycle - your go
      x = getPlanetName(sectorID)
      PutMsg player.PlayName & " fullburned to " & IIf(x = "", "sector " & sectorID, x), player.ID, Logic!Gamecntr
      x = resolveToken(sectorID)
      'check if Alliance is in the Sector
      If x = 0 Then x = getCruiserCorvette(sectorID)
      If (x = 5 Or x = 6) And isOutlaw(player.ID) Then 'no Nav card when Alliance arrives
         If actionSeq = ASNavEvade Then
            frmAction.fullburndone = True
         Else
            actionSeq = ASselect 'in limbo awaiting user to select
            showActions   'throw it back to the action window
         End If
      Else
         If actionSeq <> ASNavEvade Then
            If isMoveCutterEnabled Then moveAutoAI 6 + RollDice(NumOfReavers)
               If actionSeq <> ASNavEvade Then 'may be set in above line
                  checkFlacGun sectorID
                  actionSeq = ASNav 'pick a Nav card
                  showNav sectorID
               End If
         End If
      End If
      
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASAllianceCall Then   'Alliance Call to own sector
      actionSeq = ASselect 'in limbo awaiting user to select
      showActions   'throw it back to the action window
      
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASnavEnd Then   'fullburn Cycle
      'deal with the Nav option chosen
      If frmNav.NavOption = 3 Then
         PutMsg player.PlayName & " discards and redraws a NAV Card using the Surveyor's Shuttle", player.ID, Logic!Gamecntr
         SurvShuttlePerk = True
         actionSeq = ASNav 'pick next Nav card
         showNav sectorID
      ElseIf frmNav.NavOption > 0 Then
         doNav frmNav.NavCardID, frmNav.NavOption
         If hasShipUpgrade(player.ID, 20) And TheBigBlack >= 0 Then 'Emissions Recycler
            checkBigBlack frmNav.NavCardID
         End If
      End If
      frmNav.FDPane1.PaneVisible = False
      'avoid special move actions like EVADE
      If actionSeq = ASnavEnd Then 'has not been modified by special moves
        'reset OffJob status
         clearOffJob player.ID
         actionSeq = ASselect 'in limbo awaiting user to select
         showActions   'throw it back to the action window
      End If
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASNavEvadeEnd Then   'fullburn Cycle
      resolveToken sectorID
      If actionSeq <> ASNavEvade Then
         checkFlacGun sectorID
         actionSeq = ASselect 'in limbo awaiting user to select
         showActions   'throw it back to the action window
      End If
   ElseIf status = "R" And thisplayer = player.ID And (actionSeq = ASNavReavEnd Or actionSeq = ASNavCrusEnd) Then   'fullburn Cycle
      actionSeq = ASselect 'in limbo awaiting user to select
      showActions   'throw it back to the action window
            
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASDeal Then   'Deal Cycle - your go
      If showDeals(False, "locals") = 0 And getUnseenDeck("Contact", ContactID) = 0 And Not HigginsDealPerk Then
         actionSeq = ASselect
      Else
         actionSeq = ASDealSelDiscard
         frmDeal.Timer1.Enabled = False
      End If
      showActions
   
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASDealDrew Then   'Deal Cycle - your go
      'save selected card as Seq = 6
      x = frmDeal.setSelected("UN", CONSIDERED)
      maxConsider = MAXJOBCARDDRAW + getGearFeature(player.ID, "MaxJobs")
      If isSolid(player.ID, 4) And ContactID = 4 Then
         maxConsider = 4
      End If
      'and draw cards up to 3
      If x < maxConsider Then
         DrawDeck "Contact", ContactID, maxConsider - x, CONSIDERED
      End If
      actionSeq = ASDealSelect
      showDeals False, "localdeal" 'only show those considered (6)
      showActions
      
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASDealEnd Then   'Deal Cycle - your go
      'save selected (Seq=6 + selected) to players Jobs, unselected back to 5
      
      x = doDeal(player.ID)
      
      PutMsg player.PlayName & " dealt and accepted " & IIf(x = 0, "no", CStr(x)) & " deals from " & varDLookup("ContactName", "Contact", "ContactID=" & ContactID), player.ID, Logic!Gamecntr
      
      'do any Sell Cargo/Contra Dealing now----------------
      If ContactID = 6 Then  'lord Harrow
         If doBuyCargo(player.ID, Val(frmAction.lblDealCargoBuy)) > 0 Then
            PutMsg player.PlayName & " bought " & frmAction.lblDealCargoBuy & " Cargo from " & varDLookup("ContactName", "Contact", "ContactID=" & ContactID), player.ID, Logic!Gamecntr
         End If
         frmAction.lblDealCargoBuy = "0"
         
      ElseIf ContactID = 9 Then  'FANTY MINGO
         If doBuyContra(player.ID, Val(frmAction.lblDealContraBuy)) Then
            PutMsg player.PlayName & " bought " & frmAction.lblDealContraBuy & " Contraband from " & varDLookup("ContactName", "Contact", "ContactID=" & ContactID), player.ID, Logic!Gamecntr
         End If
         frmAction.lblDealContraBuy = "0"
         
      ElseIf doSellCargoContra(player.ID, ContactID, Val(frmAction.lblDealCargoSell), Val(frmAction.lblDealContraSell)) > 0 Then
         PutMsg player.PlayName & " sold " & frmAction.lblDealCargoSell & " Cargo and " & frmAction.lblDealContraSell & " Contraband to " & varDLookup("ContactName", "Contact", "ContactID=" & ContactID), player.ID, Logic!Gamecntr
         frmAction.lblDealCargoSell = "0"
         frmAction.lblDealContraSell = "0"
         
      End If
      
      'Bree sells parts to any Solids
      If hasCrew(player.ID, 34) And varDLookup("Parts", "Players", "PlayerID=" & player.ID) >= Val(frmAction.lblDealPartsSell) And Val(frmAction.lblDealPartsSell) > 0 And isSolid(player.ID, ContactID) Then
         DB.Execute "UPDATE Players SET Parts = Parts-" & frmAction.lblDealPartsSell & ", Pay = Pay + " & Val(frmAction.lblDealPartsSell) * 300 & " WHERE PlayerID=" & player.ID
         PutMsg player.PlayName & " used Bree's Black Market Ties to sell " & frmAction.lblDealPartsSell & " Parts to " & varDLookup("ContactName", "Contact", "ContactID=" & ContactID), player.ID, Logic!Gamecntr
         frmAction.lblDealPartsSell = "0"
      End If
      
      'Deal with Harken to source Fuel (not a Buy action
      If frmAction.imgDealFuel.Tag = "Y" And doBuyFuelParts(player.ID, Val(frmAction.lblDealFuelBuy), 0, True) <= getMoney(player.ID) And ContactID = 5 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= (Val(frmAction.lblDealFuelBuy) / 2) Then
            If doBuyFuelParts(player.ID, Val(frmAction.lblDealFuelBuy), 0) Then
               PutMsg player.PlayName & " bought " & frmAction.lblDealFuelBuy & " Fuel from Harken", player.ID, Logic!Gamecntr
            End If
         Else
            MessBox "Not enough Cargo Space for the Fuel order", "Cargo Space", "Ooops", "", getLeader()
         End If
         frmAction.lblDealFuelBuy = "0"
         'frmAction.lblBuyParts = "0"
      End If

      
      'clear all Warrants?
      If frmAction.imgClearWarrantsOpt.Tag = "Y" Then
         If varDLookup("Pay", "Players", "PlayerID=" & player.ID) >= 1000 Then
            DB.Execute "UPDATE Players SET Warrants = 0, Pay = Pay - 1000 WHERE PlayerID=" & player.ID
            PutMsg player.PlayName & " had Badger clear all Warrants", player.ID, Logic!Gamecntr
            frmAction.clearWarrant
         Else
            MessBox "Not enough money left to pay Badger to clear all Warrants", "Warrants", "Ooops", "", getLeader()
         End If
      End If
      
      'load pasengers & Fugitives at Amnons
      If frmAction.imgLoadPassngr.Tag = "Y" Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= Val(frmAction.lblDealPassngrLoad) + Val(frmAction.lblDealFugiLoad) Then
            DB.Execute "UPDATE Players SET Passenger = Passenger + " & CStr(Val(frmAction.lblDealPassngrLoad)) & ", Fugitive = Fugitive + " & CStr(Val(frmAction.lblDealFugiLoad)) & " WHERE PlayerID = " & player.ID
            PutMsg player.PlayName & " loaded " & CStr(Val(frmAction.lblDealPassngrLoad)) & " Passengers and " & CStr(Val(frmAction.lblDealFugiLoad)) & " Fugitives", player.ID, Logic!Gamecntr
         End If
         frmAction.lblDealPassngrLoad = "0"
         frmAction.lblDealFugiLoad = "0"
      End If
            
      
      drawLine 1, -1
      actionSeq = ASselect 'in limbo awaiting user to select
      showDeals False, "local"
      If Not (frmJob Is Nothing) Then frmJob.refreshJobs
      'frmDeal.RefreshDeals
      'frmDeal.Timer1.Enabled = True
      showActions

      
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASBuy Then   'Buy Cycle - your go
      showBuys False, "local"
      actionSeq = ASBuySelDiscard
      'frmBuy.Timer1.Enabled = False
      showActions
   
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASBuyDrew Then   'Buy Cycle - your go
      'save selected card as Seq = 6
      x = frmBuy.setSelected("UN", CONSIDERED)
      'and draw cards up to 3
      If x < MAXJOBCARDDRAW Then
         DrawDeck "Supply", SupplyID, MAXJOBCARDDRAW - x, CONSIDERED
      End If
      actionSeq = ASBuySelect
      showBuys False, "localbuy" 'only show those considered (6)
      showActions
      
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASBuyShore Then   'Buy Cycle - your go
      x = doShoreLeave(player.ID)
      actionSeq = ASselect 'in limbo awaiting user to select
      
      If getPerkAttributeCrew(player.ID, "FreeShoreLeave") > 0 Then
         PutMsg player.PlayName & " had the Barkeep shout the Crew some free Shore Leave", player.ID, Logic!Gamecntr, True, 71
      ElseIf hasShipUpgrade(player.ID, 19) Then
         PutMsg player.PlayName & " treated the Crew with a shiny Board Game for $" & CStr(Abs(x)), player.ID, Logic!Gamecntr, True, 0, 0, 19
      Else
         PutMsg player.PlayName & " went on Shore Leave at " & varDLookup("PlanetName", "Planet", "SectorID=" & sectorID) & " for $" & CStr(Abs(x)), player.ID, Logic!Gamecntr
      End If
            
      showActions
      
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASBuyHaven Then   'Buy Cycle - your go
    
      'buy fuel & parts now
      If frmAction.imgFuelBuy.Tag = "Y" Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= (Val(frmAction.lblBuyFuel)) / 2 Then
            If doBuyFuelParts(player.ID, Val(frmAction.lblBuyFuel), 0, False, IIf(getHaven(sectorID) = player.ID, 4, 0)) = 0 Then
               
            End If
            PutMsg player.PlayName & " loaded " & frmAction.lblBuyFuel & " Fuel at the Haven" & IIf(getHaven(sectorID) = player.ID, ", up to 4 for free!", ""), player.ID, Logic!Gamecntr
         Else
            MessBox "Not enough Cargo Space for the Fuel/Parts order", "Fuel/Parts order", "Ooops", "", getLeader()
         End If
         frmAction.lblBuyFuel = "0"
         frmAction.lblBuyParts = "0"
      End If
      x = 0
      'Haven Shore Leave
      If hasDisgruntled(player.ID) Then
         If MessBox("Do you want to take Shore Leave at your Haven as well?", "Haven Leave", "Yes", "No", getLeader()) = 0 Then
            x = doShoreLeave(player.ID, False, (getHaven(sectorID) = player.ID))
         End If
      End If
      
      actionSeq = ASselect 'in limbo awaiting user to select
      If x = -1 Then
         PutMsg player.PlayName & " took some free Shore Leave at the Haven", player.ID, Logic!Gamecntr, True, getLeader()
      ElseIf x > 0 Then
         PutMsg player.PlayName & " went on Shore Leave at a Haven for $" & CStr(Abs(x)), player.ID, Logic!Gamecntr
      End If
      showActions
      
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASBuyEnd Then   'Buy Cycle - your go
      'save selected (Seq=6 + selected) to players Jobs, unselected back to 5
      x = doBuy(player.ID)
      PutMsg player.PlayName & " accepted and bought " & IIf(x = 0, "no", CStr(x)) & " buys from " & varDLookup("SupplyName", "Supply", "SupplyID=" & SupplyID), player.ID, Logic!Gamecntr
      
      'buy fuel & parts now
      If frmAction.imgFuelBuy.Tag = "Y" And (Val(frmAction.lblBuyFuel) > 0 Or Val(frmAction.lblBuyParts) > 0) Then
         'If doBuyFuelParts(player.ID, Val(frmAction.lblBuyFuel), Val(frmAction.lblBuyParts), True) <= getMoney(player.ID) Then
         '   If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= (Val(frmAction.lblBuyFuel) + Val(frmAction.lblBuyParts)) / 2 Then
               If doBuyFuelParts(player.ID, Val(frmAction.lblBuyFuel), Val(frmAction.lblBuyParts)) > 0 Then
                  PutMsg player.PlayName & " bought " & frmAction.lblBuyFuel & " Fuel and " & frmAction.lblBuyParts & " Parts", player.ID, Logic!Gamecntr
               End If
         '   Else
         '      MessBox "Not enough Cargo Space for the Fuel/Parts order", "Fuel/Parts order", "Ooops", "", getLeader()
         '   End If
         'Else
         '   MessBox "Not enough money left to pay for the Fuel or Parts", "Fuel/Parts order", "Ooops", "", getLeader()
         'End If
         frmAction.lblBuyFuel = "0"
         frmAction.lblBuyParts = "0"
      End If
      
      actionSeq = ASselect 'in limbo awaiting user to select
      showBuys False, "local"
      'frmBuy.Timer1.Enabled = True
      If frmShip Is Nothing Then Set frmShip = New frmShips
      frmShip.RefreshShips
      frmBuy.RefreshBuys
      showActions
      
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASWork Then   'Work Cycle - your go
      
      If frmAction.lblJobName.Tag = "" Then  'make work
         If hasCrew(player.ID, 73) Then  'Busker adds 100
            getMoney player.ID, 300
            PutMsg player.PlayName & " made Extra Work with Busker at " & varDLookup("PlanetName", "Planet", "SectorID=" & sectorID), player.ID, Logic!Gamecntr
         Else
            getMoney player.ID, 200
            PutMsg player.PlayName & " made Work at " & varDLookup("PlanetName", "Planet", "SectorID=" & sectorID), player.ID, Logic!Gamecntr
         End If
         
         If hasCrew(player.ID, 78) And (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID)) >= 1 Then ' Holder- When you Make-Work, you may also take a Fugitive
            If MessBox("Holder has lined up a Fugitive for us" & vbNewLine & "Do you want to take them on board?", "Load 1 Fugitive?", "Yes", "No", 78) = 0 Then
               DB.Execute "UPDATE Players SET Fugitive = Fugitive + 1 WHERE PlayerID =" & player.ID
            End If
         End If
         
         actionSeq = ASselect 'in limbo awaiting user to select
         
      ElseIf Val(frmAction.lblJobName.Tag) = -1 Then 'Haven Supplies
         doHavenSupplies
         actionSeq = ASselect 'in limbo awaiting user to select
         
      ElseIf Val(frmAction.lblJobName.Tag) < -1 Then 'grab a Bounty
         doBountyHunt Abs(Val(frmAction.lblJobName.Tag))
         actionSeq = ASselect 'in limbo awaiting user to select
         'reset OffJob status
         clearOffJob player.ID
      
      Else  'do a Job
         If doWork(player.ID, Val(frmAction.lblJobName.Tag)) = 0 Then ' normal exit
            actionSeq = ASselect 'this one is conditional based on the job outcome <<<< !
         End If
         'reset OffJob status
         clearOffJob player.ID
      End If
      
      showActions
      
      
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASRemoveDisgr Then
      removeSelDisgruntled player.ID
      PutMsg player.PlayName & " removed Disgruntle from a Crew", thisplayer, Logic!Gamecntr
      actionSeq = ASselect 'in limbo awaiting user to select
      showActions
   
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASResolveAlertEnd Then
      actionSeq = ASselect 'in limbo awaiting user to select
      showActions
   
   ElseIf status = "R" And thisplayer = player.ID And actionSeq = ASend Then 'Finish up your turn
      'formAction.FDPane1.PaneVisible = False
      frmAction.actionButtonEnable "imgEndTurn", False
      wormHoleOpen = False
      drawLine 2, -1
      'option to pull top card from Supply Decks
      If varDLookup("pullSupply", "Story", "StoryID=" & Logic!StoryID) = 1 Then
         For x = 1 To 7
            DrawDeck "Supply", x, 1
         Next x
      End If
      
      'option to pull top card from Supply Decks
      If varDLookup("pullContact", "Story", "StoryID=" & Logic!StoryID) = 1 Then
         For x = 1 To 9
            DrawDeck "Contact", x, 1
         Next x
      End If
      
      'Check if WON!
      If CheckWon(player.ID) Then
         If MessBox("Do you want to end the game for all players?", "End Game?", "Yes", "No", getLeader()) = 0 Then
            PutMsg PlayCode(thisplayer).PlayName & " has ENDED the Game", thisplayer, Logic!Gamecntr
            DB.Execute "UPDATE GameSeq SET Seq = 'E'"
            'Logic.Update "Seq", "E"
            EndGame
            Exit Sub
         End If
      End If
      
      'turn finished, push to next player (for SP thats you)
      thisplayer = setNextPlayer(player.ID)
      If thisplayer <> player.ID Then
         PutMsg PlayCode(thisplayer).PlayName & "'s Turn", thisplayer, Logic!Gamecntr
      End If
      actionSeq = ASidle
      If Not (frmShip Is Nothing) Then frmShip.RefreshShips
   End If
  
   Exit Sub
  
err_handler:
  errh = MsgBox(Err.Description, vbCritical + vbAbortRetryIgnore, "Error in Main Cycle")
  Select Case errh
  Case vbRetry
    Resume
  Case vbAbort
    'exit
  Case vbIgnore
    Resume Next
  End Select
  
   
   
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim ChatTxt As String, x
  playsnd 13
  Logic.Requery
  Select Case Button.key
  Case "start"  'host
    
    Select Case Logic!Seq
     Case "E"
       SoloGame = False
       player.ID = 0
       player.Color = ""
       player.PlayName = ""
       DB.Execute "UPDATE GameSeq SET Seq = 'H', GameCntr = 0"
       'Logic!Seq = "H"
       'Logic!GameCntr = 0
       'Logic.Update
       ClearBoard
       refreshSolid
       'Starter.cbo.Enabled = True
       Starter.isHost = True
       Starter.Show 1
              
       If player.ID = 0 Then 'no active player
         DB.Execute "UPDATE GameSeq SET Seq = 'E'"
         'Logic.Update "Seq", "E"
         ClearBoard
       Else
         
         pickStartSector = 0
         actionSeq = ASidle
         initBoard
         'Verse.Timer1.Enabled = True
         showEvents
         initToolbar True
         Toolbar1.Buttons("exit").Enabled = True
       End If
       
     Case "H"
        If MsgBox("Game already being hosted, do you need to reset it?", vbYesNo + vbCritical, "Game in Host mode") = vbYes Then
            DB.Execute "UPDATE GameSeq SET Seq = 'E'"
            'Logic.Update "Seq", "E"
        End If

     Case Else
        dontnagme = True
        x = MsgBox("Game in progress. If you want to re-join, use JOIN button." & vbNewLine & "otherwise press OK to RESET the Game", vbExclamation + vbOKCancel, "Game in Progress")
        Select Case x
        Case vbOK
            DB.Execute "UPDATE GameSeq SET Seq = 'E'"
            'Logic.Update "Seq", "E"
            MsgBox "Game has been reset, press Host again to start", vbInformation
        End Select
     End Select
       
  Case "join"
     
     Select Case Logic!Seq
     Case "H", "E"
         player.ID = 0
         player.Color = ""
         player.PlayName = ""
         refreshSolid
         Starter.isHost = False
         Starter.Show 1
         
         If player.ID > 0 Then
            initBoard
            'Verse.Timer1.Enabled = True
            pickStartSector = 0
            actionSeq = ASidle
            showEvents
            initToolbar True, False
         End If

     Case Else
        dontnagme = reJoin

     End Select

  Case "exit"  'END the Game
    ' Confirm Exit msgbox ?
    If MessBox("Are you sure you want to leave this game?", "Closing Game?", "Yes", "No") = 1 Then
       Exit Sub
    End If

    EndGame
    
  Case "chat"
    ChatTxt = InputBox("Enter your message", "Chat")
    If ChatTxt <> "" Then
      PutMsg player.PlayName & " : " & ChatTxt, 0
    End If
    
  Case "graph"
   If frmStat Is Nothing Then
       Set frmStat = New frmStats
       frmStat.FDPane1.PaneVisible = False
    End If
    If frmStat.FDPane1.PaneVisible = False Then
      showStats
    Else
       frmStat.FDPane1.PaneVisible = False
       frmStat.Timer1.Enabled = False
    End If
    DockFrame1.SaveStates "Firefly"
    
  Case "log"

    Events.FDPane1.PaneVisible = Not Events.FDPane1.PaneVisible
    DockFrame1.SaveStates "Firefly"
        
  Case "crew"
    Dim frmCrew As New frmCrewSel
    frmCrew.crewFilter = " Order By CrewName"
    frmCrew.AlwaysOnTop = True
    frmCrew.Show
    Set frmCrew = Nothing
  
    
  Case "gear"
    Dim frmGear As New frmGearView
    frmGear.gearFilter = " Order By GearName"
    frmGear.AlwaysOnTop = True
    frmGear.Show
    Set frmGear = Nothing
    
  Case "upgrd"
    Dim frmUpGrd As New frmShipUpgrdView
    frmUpGrd.gearFilter = " Order By UpgradeName"
    frmUpGrd.AlwaysOnTop = True
    frmUpGrd.Show
    Set frmUpGrd = Nothing
  
    
  Case "job"
    If frmJob Is Nothing Then
       Set frmJob = New frmJobs
       frmJob.FDPane1.PaneVisible = False
    End If
    With frmJob
    If .FDPane1.PaneVisible Then
        .FDPane1.PaneVisible = False
    Else
        If .jobFilter = "" Then
            .Caption = player.PlayName & "'s Jobs"
            .jobFilter = ""
        End If
        .FDPane1.InitDockHW = 120
        .FDPane1.InitDockStyle = DockToTop
        .FDPane1.PaneVisible = True
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        .Show
        .refreshJobs
    End If
    End With

    
  Case "deal"
     If actionSeq < ASDeal Or actionSeq > ASDealEnd Then
       showDeals True
     End If
    
  Case "buy"
    If actionSeq < ASBuy Or actionSeq > ASBuyEnd Then
      showBuys True
    End If
    
  Case "ship"
    If frmShip Is Nothing Then
       Set frmShip = New frmShips
       frmShip.FDPane1.PaneVisible = False
    End If
    With frmShip
      If .FDPane1.PaneVisible Then
          .FDPane1.PaneVisible = False
      Else
           If .shipFilter = "" Then
              .Caption = player.PlayName & "'s Ship"
              .shipFilter = "me"
           End If
          .FDPane1.InitDockHW = 400
          .FDPane1.InitDockStyle = DockToTop
          .FDPane1.PaneVisible = True
          .FDPane1.PinState = Pinned
          .FDPane1.SetLayoutReference Nothing
          .Show
          .RefreshShips
      End If
    End With
  
  Case "help"
    x = ShellExecute(x, "OPEN", App.Path & "\FireflyForPC.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
  
  End Select
End Sub

Private Function reJoin() As Boolean
   If rejoininprogress Or getPlayerCount() = 0 Then
      Exit Function
   End If
   rejoininprogress = True
   If MessBox("Do you want to rejoin in " & IIf(getPlayerCount() > 1, "Multiplayer", "Single Player") & " mode?" & vbNewLine & vbNewLine & "Press OK to join the Game", "Game in Progress", "OK", "Cancel", 1) = 0 Then
       player.ID = getNewPlayer()
       player.PlayName = Nz(varDLookup("Name", "Players", "PlayerID =" & player.ID))
       getPlayerCount True
       SoloGame = isSoloGame()
       pickStartSector = 2  'flag the selection is done
       actionSeq = ASidle
       initBoard
       refreshSolid
       DB.Execute "DELETE from ShowdownScores WHERE PlayerID =" & player.ID
       DB.Execute "DELETE from ShowdownGear WHERE PlayerID =" & player.ID
       
       'Verse.Timer1.Enabled = True
       showEvents
       initToolbar True, False
       reJoin = True
   End If
   rejoininprogress = False
End Function

Private Sub EndGame()
   killAllForms
    
    If Logic!player = player.ID Then setNextPlayer player.ID
    DB.Execute "Update Players Set Name = NULL WHERE PlayerID = " & player.ID
    If Nz(varDLookup("PlayerID", "Players", "Name IS NOT NULL"), 0) = 0 Then
      DB.Execute "UPDATE GameSeq SET Seq = 'E', GameCntr = 0"
       Logic.Requery
    End If

    player.ID = 0
    player.Color = ""
    player.PlayName = ""
    pickStartSector = -1
    actionSeq = ASidle
    initToolbar False
End Sub

Private Sub initBoard()
Dim rst As New ADODB.Recordset
Dim coords, c() As String, x
   NumOfReavers = getNumOfReavers()
   If Verse Is Nothing Then
      Set Verse = New Board
   End If
   With Verse
      
      .Picture1.Picture = LoadPicture(App.Path & "\Pictures\" & Logic!BoardPicture)
      .Height = Logic!BHeight
      .Width = Logic!BWidth
   
      For x = 1 To 4
         Load .imgHaven(x)
         Set .imgHaven(x).Container = .Picture1
         .imgHaven(x).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Haven" & x & ".bmp")
         .imgHaven(x).TransparentColor = &HFFFFFF
         .imgHaven(x).TransparentColorMode = lvicUseTransparentColor
      Next x
      rst.CursorLocation = adUseClient
      rst.Open "SELECT SectorID, Slot5,STop,SLeft,SHeight,SWidth, Token, AToken, Haven FROM Board WHERE SectorID > 0 ORDER BY SectorID", DB, adOpenStatic, adLockReadOnly
      While Not rst.EOF
         Load .HotSpot(rst!sectorID)
         Set .HotSpot(rst!sectorID).Container = .Picture1
         .HotSpot(rst!sectorID).top = rst!STop
         .HotSpot(rst!sectorID).Left = rst!SLeft
         .HotSpot(rst!sectorID).Height = rst!SHeight
         .HotSpot(rst!sectorID).Width = rst!SWidth
         .HotSpot(rst!sectorID).ZOrder
         .HotSpot(rst!sectorID).Visible = True
         coords = rst.Fields("Slot5").Value
         c = Split(coords, ",")
         
         Load .imgAToken(rst!sectorID)
         .imgAToken(rst!sectorID).Left = c(0)
         .imgAToken(rst!sectorID).top = c(1)
         Set .imgAToken(rst!sectorID).Container = .Picture1
         .imgAToken(rst!sectorID).Tag = CStr(rst!AToken)
         If rst!AToken > 0 Then
            .imgAToken(rst!sectorID).Picture = LoadPictureGDIplus(App.Path & "\Pictures\AToken" & IIf(rst!AToken > 6, 6, rst!AToken) & ".gif")
            .imgAToken(rst!sectorID).Visible = True
            .imgAToken(rst!sectorID).Animate2.StartAnimation
         End If
         .imgAToken(rst!sectorID).TransparentColor = &HFFFFFF
         .imgAToken(rst!sectorID).TransparentColorMode = lvicUseTransparentColor
         
         Load .imgToken(rst!sectorID)
         .imgToken(rst!sectorID).Left = c(0) + 100
         .imgToken(rst!sectorID).top = c(1) + 100
         Set .imgToken(rst!sectorID).Container = .Picture1
         .imgToken(rst!sectorID).Tag = CStr(rst!Token)
         If rst!Token > 0 Then
            .imgToken(rst!sectorID).Picture = LoadPictureGDIplus(App.Path & "\Pictures\RToken" & IIf(rst!Token > 6, 6, rst!Token) & ".gif")
            .imgToken(rst!sectorID).Visible = True
            .imgToken(rst!sectorID).Animate2.StartAnimation
         End If
         .imgToken(rst!sectorID).TransparentColor = &HFFFFFF
         .imgToken(rst!sectorID).TransparentColorMode = lvicUseTransparentColor
         
         If rst!Haven > 0 Then
            .imgHaven(rst!Haven).Left = c(0)
            .imgHaven(rst!Haven).top = c(1)
            .imgHaven(rst!Haven).Visible = True
         End If
         
         rst.MoveNext
      Wend
      For x = 5 To 6 + NumOfReavers ' .Imag.Count
         .Imag(x).Animate2.StartAnimation
      Next x
      .Caption = "the 'Verse - " & varDLookup("StoryTitle", "Story", "StoryID=" & Logic!StoryID)
      .Show
      
      
   End With
End Sub


Private Sub initToolbar(ByVal start As Boolean, Optional ByVal admin As Boolean = True)
   With Toolbar1
      .Buttons("exit").Enabled = start
      .Buttons("chat").Enabled = start
      .Buttons("graph").Enabled = start
      .Buttons("log").Enabled = start
      .Buttons("crew").Enabled = start
      .Buttons("gear").Enabled = start
      .Buttons("upgrd").Enabled = start
      .Buttons("job").Enabled = start
      .Buttons("ship").Enabled = start
      .Buttons("deal").Enabled = start
      .Buttons("buy").Enabled = start
      .Buttons("start").Enabled = Not start
      .Buttons("join").Enabled = Not start
      .Buttons("help").ButtonMenus("jobs").Enabled = admin
   End With
End Sub

Private Sub killAllForms()
Dim x

   If Verse.Visible Then
      Verse.hide
      Unload Verse
      Set Verse = Nothing
   End If
   
   If frmJob Is Nothing Then
   Else
      Unload frmJob
      Set frmJob = Nothing
   End If
   If frmShip Is Nothing Then
   Else
      Unload frmShip
      Set frmShip = Nothing
   End If
   If frmDeal Is Nothing Then
   Else
      Unload frmDeal
      Set frmDeal = Nothing
   End If
   If frmBuy Is Nothing Then
   Else
      Unload frmBuy
      Set frmBuy = Nothing
   End If
   If frmStat Is Nothing Then
   Else
     Unload frmStat
     Set frmStat = Nothing
   End If
   If Events.FDPane1.PaneVisible Then Unload Events
   frmAction.endAction
   If frmAction.FDPane1.PaneVisible Then Unload frmAction
   
   For x = 1 To NO_OF_CONTACTS
     pic(x).Visible = False
   Next x


End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim frmJobEdit As frmJobEditor, x

   playsnd 13
   Select Case ButtonMenu.key
   Case "alljobs"
      If actionSeq < ASDeal Or actionSeq > ASDealEnd Then
         If frmJob Is Nothing Then
            Set frmJob = New frmJobs
            frmJob.FDPane1.PaneVisible = False
         End If
         frmJob.Caption = "All Jobs"
         frmJob.jobFilter = "all"
         frmJob.refreshJobs
         frmJob.Show
         frmJob.FDPane1.PaneVisible = True
      End If
      
   Case "myjobs"
      If actionSeq < ASDeal Or actionSeq > ASDealEnd Then
         If frmJob Is Nothing Then
            Set frmJob = New frmJobs
            frmJob.FDPane1.PaneVisible = False
         End If
         frmJob.Caption = player.PlayName & "'s Jobs"
         frmJob.jobFilter = ""
         frmJob.refreshJobs
         frmJob.Show
         frmJob.FDPane1.PaneVisible = True
      End If
      
   Case "alldeals"
     If actionSeq < ASDeal Or actionSeq > ASDealEnd Then
       showDeals
     End If

   Case "localdeals"
      showDeals False, "local"

   Case "allbuys"
      If actionSeq < ASBuy Or actionSeq > ASBuyEnd Then
         showBuys
      End If

   Case "localbuys"
      If actionSeq < ASBuy Or actionSeq > ASBuyEnd Then
         showBuys False, "local"
      End If
      
   Case "allships"
      If frmShip Is Nothing Then
         Set frmShip = New frmShips
         frmShip.FDPane1.PaneVisible = False
      End If
      frmShip.Caption = "All Ships"
      frmShip.shipFilter = "all"
      frmShip.RefreshShips
      frmShip.Show
      frmShip.FDPane1.PaneVisible = True
   Case "myship"
      If frmShip Is Nothing Then
         Set frmShip = New frmShips
         frmShip.FDPane1.PaneVisible = False
      End If
      frmShip.Caption = player.PlayName & "'s Ship"
      frmShip.shipFilter = "me"
      frmShip.RefreshShips
      frmShip.Show
      frmShip.FDPane1.PaneVisible = True
      
   Case "firefly"
      x = ShellExecute(x, "OPEN", App.Path & "\Firefly_rulebook.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
   Case "bluesun"
      x = ShellExecute(x, "OPEN", App.Path & "\FireflyBlueSun_rulebook.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
   Case "kalidasa"
      x = ShellExecute(x, "OPEN", App.Path & "\FireflyKalidasa_rulebook.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
   Case "pbh"
      x = ShellExecute(x, "OPEN", App.Path & "\Firefly_Pirates_Bounty_Hunters_rulebook.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind

   Case "pcguide"
      x = ShellExecute(x, "OPEN", App.Path & "\FireflyForPC.pdf", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
   Case "jobs"
    Set frmJobEdit = New frmJobEditor
    frmJobEdit.Show 1
   Case "bot"
      x = ShellExecute(x, "OPEN", App.Path & "\FireflyAIBot.exe ", datab, vbNullString, 1)               '1=normal, 2=min, 3=max, 4=behind
   Case "map"

      Logic.Requery
'      x = Logic!BoardPicture
'      If x = "KalidasaBoard.jpg" Then
'         x = "KalidasaBoard1.jpg"
'      Else
'         x = "KalidasaBoard.jpg"
'      End If
'      DB.Execute "UPDATE GameSeq SET BoardPicture ='" & x & "'"
'      Logic.Requery
'      If Verse.isLoaded Then
'         Verse.Picture1.Picture = LoadPicture(App.Path & "\Pictures\" & x)
'      End If

      showEditBoard
   Case "check"
      x = ShellExecute(x, "OPEN", "https://github.com/blistik/FireflyKalidasa/releases", vbNullString, vbNullString, 1)              '1=normal, 2=min, 3=max, 4=behind
   
   Case "about"
      MessBox "Firefly + Blue Sun/Kalidasa  V" & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & "*Open Freeware* - use at your own risk" & vbNewLine & "Coded by: Vee Bee-er " & Chr(169) & " 2021-24 BLiSoftware" & _
      vbNewLine & "All rights reserved - GF9 & Fox", "About", "Shiny", "", 0, 2
   End Select

End Sub

Private Sub showEditBoard()
    With editBoard
        .FDPane1.InitDockHW = 200
        .FDPane1.InitDockStyle = DockToLeft
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        .FDPane1.PaneVisible = True
        '.refreshform
        '.Timer1.Enabled = True
    End With
End Sub

Private Sub showStats()
    With frmStat
        .FDPane1.InitDockHW = 200
        .FDPane1.InitDockStyle = DockToLeft
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        .FDPane1.PaneVisible = True
        .refreshform
        .Timer1.Enabled = True
    End With
End Sub
Private Sub showEvents()
    With Events
        .FDPane1.InitDockHW = 200
        .FDPane1.InitDockStyle = DockToLeft
        .FDPane1.PaneVisible = True
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        .Timer1.Enabled = True
        '.Show
    End With
End Sub

Public Function showDeals(Optional ByVal toggle As Boolean = False, Optional ByVal filter As String = "all") As Variant

    If frmDeal Is Nothing Then
       Set frmDeal = New frmDeals
       frmDeal.FDPane1.PaneVisible = False
    End If
    With frmDeal
    If .FDPane1.PaneVisible And toggle Then
        .FDPane1.PaneVisible = False
    Else
        .dealFilter = filter
        .FDPane1.InitDockHW = 400
        .FDPane1.InitDockStyle = DockToTop
        .FDPane1.PaneVisible = True
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        '.Show
        showDeals = .RefreshDeals
    End If
    End With
    
End Function


Public Sub showBuys(Optional ByVal toggle As Boolean = False, Optional ByVal filter As String = "all")

    If frmBuy Is Nothing Then
       Set frmBuy = New frmSupply
       frmBuy.FDPane1.PaneVisible = False
    End If
    With frmBuy
    If .FDPane1.PaneVisible And toggle Then
        .FDPane1.PaneVisible = False
    Else
        .buyFilter = filter
        .FDPane1.InitDockHW = 400
        .FDPane1.InitDockStyle = DockToTop
        .FDPane1.PaneVisible = True
        .FDPane1.PinState = Pinned
        .FDPane1.SetLayoutReference Nothing
        .RefreshBuys
    End If
    End With
    
End Sub

Private Sub showNav(ByVal sectorID, Optional ByVal CardID As Integer = 0)
Dim SQL, reshuffle, Zone, x
Dim rst As New ADODB.Recordset

   With frmNav
      .FDPane1.InitDockHW = 200
      .FDPane1.InitDockStyle = DockToLeft
      .FDPane1.PinState = Pinned
      .FDPane1.SetLayoutReference Nothing
      
      .NavCardID = 0
      .NavOption = 0
      
      .FDPane1.PaneVisible = False
      
      'SectorID = Nz(varDLookup("SectorID", "Players", "PlayerID=" & player.ID), 0)
      Zone = varDLookup("Zones", "Board", "SectorID=" & sectorID)
      
      'Read in the next NAV card and display either 1 or 2 options
      
       'OPTION 1 ===================================================================================
      SQL = "SELECT NavDeck.CardID, NavDeck.CardName, NavDeck.Reshuffle, NavDeck.Seq, NavOption.* "
      SQL = SQL & "FROM NavOption INNER JOIN NavDeck ON NavOption.OptionID = NavDeck.Option1ID "
      If CardID <> 0 Then
         SQL = SQL & "Where NavDeck.CardID = " & CardID
      Else
         SQL = SQL & "Where NavDeck.Zones = '" & Zone & "' And NavDeck.Seq > 6 "
         SQL = SQL & "ORDER BY NavDeck.Seq"
         If Left(datab, 16) = "Provider=MSDASQL" Then SQL = SQL & " LIMIT 1"
      End If
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      If rst.EOF Then  ' this happens when the reshuffle card is in the discard pile at start of game setup
         ShuffleDeck "Nav", True, False, Zone
         PutMsg player.PlayName & " Reshuffling NavDeck " & Zone & " due to end of deck", player.ID, Logic!Gamecntr, True, getLeader()
         rst.Close
         rst.CursorLocation = adUseClient
         rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      End If
      If Not rst.EOF Then
         
         'a LOT of these tests are only applied to the 1st option only
         .cmd(0).Enabled = hasNavReqs(player.ID, rst!CardID, 1)
                  
         If (rst!CardName = "Reaver Cutter!") Then ' move cutter here and deal with it after
            PutMsg player.PlayName & " has a gorram Reaver Cutter closing in!", player.ID, Logic!Gamecntr, True, getLeader()
            If getCutterSector(sectorID) = 0 Then MoveShip 6 + RollDice(NumOfReavers), sectorID
         End If
         
         If actionSeq = ASNavEvade Then 'we are evading already
            'evade
         ElseIf (rst!CardName = "Reaver Cutter!") And getCruiserCorvette(sectorID) = 6 Then 'corvette shoos the Reavers away
            x = getCutterSector(sectorID)
            moveAutoAI x
            actionSeq = ASnavEnd
            PutMsg player.PlayName & " is Shielded from a Reaver Cutter attack by the Alliance Corvette", player.ID, Logic!Gamecntr, True, getLeader()
         ElseIf checkFlacGun(sectorID, Not (rst!CardName = "Reaver Cutter!")) Then
            actionSeq = ASnavEnd
            
         'skip Customs Inspection if solid with Harken
         ElseIf (rst!CardName = "Customs Inspection") And isSolid(player.ID, 5) Then
            DB.Execute "UPDATE NavDeck SET Seq = " & CStr(player.ID) & " WHERE CardID = " & CStr(rst!CardID)
            actionSeq = ASnavEnd
            PutMsg player.PlayName & " being Solid with Harken avoided a Customs Inspection", player.ID, Logic!Gamecntr, True, getLeader()
         Else
            .NavCardID = rst!CardID
            .lblName.Caption = rst!CardName
            .cmd(0).Caption = rst!OptionName
            .cmd(0).ToolTipText = rst!OptionName
            .lblDetail(0).Caption = rst!Details
            
            If rst!skill = 0 Then
               .SkillImg(0).Visible = False
            Else
               .SkillImg(0).Picture = LoadPictureGDIplus(App.Path & "\Pictures\skill" & rst!skill & ".bmp")
               .SkillImg(0).Visible = True
               .SkillImg(0).TransparentColor = &H0
               .SkillImg(0).TransparentColorMode = lvicUseTransparentColor
            End If
         
            'set colours & background up
            
            Select Case Zone
            Case "A"
               .lblName.ForeColor = &HFFF690
               .lblDetail(0).ForeColor = &HFFFFDD
               .lblDetail(0).BackColor = &H2C1412
               .lblDetail(1).ForeColor = &HFFFFDD
               .lblDetail(1).BackColor = &H2C1412
               .BackColor = &H400000
            Case "B"
               .lblName.ForeColor = &HAACAED
               .lblDetail(0).ForeColor = &HDBFDFB
               .lblDetail(0).BackColor = &H1E2322
               .lblDetail(1).ForeColor = &HDBFDFB
               .lblDetail(1).BackColor = &H1E2322
               .BackColor = &H90A0D
            Case "R"
               .lblName.ForeColor = &H3A92F6
               .lblDetail(0).ForeColor = &HDBFDFB
               .lblDetail(0).BackColor = &H352035
               .lblDetail(1).ForeColor = &HDBFDFB
               .lblDetail(1).BackColor = &H352035
               .BackColor = &H90A0D
            End Select
         End If
         reshuffle = rst!reshuffle
         'pull the card out of the deck, assign it to the user for debugging
         DB.Execute "UPDATE NavDeck SET Seq = " & CStr(player.ID) & " WHERE CardID = " & CStr(.NavCardID)
         'rst!Seq = player.ID
         'rst.Update
      Else
         MsgBox "Error in NavDeck"
         Exit Sub
      End If
      rst.Close
      
      'OPTION 2 ===================================================================================
      SQL = "SELECT NavOption.* "
      SQL = SQL & "FROM NavOption INNER JOIN NavDeck ON NavOption.OptionID = NavDeck.Option2ID "
      SQL = SQL & "Where NavDeck.CardID = " & .NavCardID
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      If Not rst.EOF Then
         .Picture = LoadPicture(App.Path & "\pictures\Nav2_" & Zone & ".jpg")
         .cmd(1).Visible = True
         .cmd(1).Enabled = hasNavReqs(player.ID, .NavCardID, 2)
         
         .lblDetail(0).Height = 1455
         .lblDetail(1).Visible = True

         .cmd(1).Caption = rst!OptionName
         .cmd(1).ToolTipText = rst!OptionName
         .lblDetail(1).Caption = rst!Details
         
         If rst!skill = 0 Then
            .SkillImg(1).Visible = False
         Else
            .SkillImg(1).Picture = LoadPictureGDIplus(App.Path & "\Pictures\skill" & rst!skill & ".bmp")
            .SkillImg(1).Visible = True
            .SkillImg(1).TransparentColor = &H0
            .SkillImg(1).TransparentColorMode = lvicUseTransparentColor
         End If
         
      Else 'no option 2
         .Picture = LoadPicture(App.Path & "\pictures\Nav1_" & Zone & ".jpg")
         .cmd(1).Visible = False
         .SkillImg(1).Visible = False
         .lblDetail(1).Visible = False
         .lblDetail(0).Height = 2085
      End If
      rst.Close
      
      .cmd(2).Visible = False
      
      If reshuffle = 1 Then 'ready for next turn
         ShuffleDeck "Nav", True, False, Zone
         PutMsg player.PlayName & " Reshuffling NavDeck " & Zone & " due to reshuffle card", player.ID, Logic!Gamecntr, True, getLeader()
         If Zone = "A" And isBountyEnabled Then
            If pushBounties() Then
               If DrawDeck("Contact", 10, 3) Then PutMsg "New Bounties available"
            End If
         End If
      
      ElseIf hasShipUpgrade(player.ID, 24) > 0 And CardID = 0 And Not SurvShuttlePerk Then 'discard option by surveyer shuttle
         .cmd(2).Visible = True
      End If
      
      .lblUnseen = "unseen: " & getUnseenNavDeck(Zone)
      
      .FDPane1.PaneVisible = (.NavCardID <> 0)
      
   End With
      
End Sub

Public Sub showActions()
Dim SQL, sectorID, onlyFullburn As Boolean, x As Integer, y As Integer, z, unseenDeck As Integer
Dim rst As New ADODB.Recordset, reaverActive As Boolean, moseyrng As Integer
Dim frmJoSel As frmJobSel, hide As Boolean

   sectorID = getPlayerSector(player.ID)
   If sectorID = 0 Then
      MsgBox "The Starting Sector was not set, please reset and start a new game", vbExclamation, "Setup Issue"
      Exit Sub
   End If
   SoloGame = isSoloGame() 'as a player may drop out
   
   If ignoreToken <> sectorID And Not (FullburnMovesDone = 0 And MoseyMovesDone = 0) Then 'must be moving into the sector to resolve token
      resolveToken sectorID
   End If
   
   'check that the REAVER is or is not here
   If getCutterSector(sectorID) > 0 Then
      checkFlacGun sectorID, Not (frmAction.checkNoOfActions = 0 And FullburnMovesDone = 0 And MoseyMovesDone = 0 And CruiserCutter <> sectorID) 'possibly chase it away
   End If
   If getCutterSector(sectorID) > 0 And frmAction.checkNoOfActions = 0 And FullburnMovesDone = 0 And MoseyMovesDone = 0 And CruiserCutter <> sectorID Then
      reaverActive = True
      showNav sectorID, -1
      CruiserCutter = sectorID
   End If
   
   If isShipHere(5, sectorID) = 5 And CruiserCutter <> sectorID Then
      CruiserCutter = sectorID 'set it as faced regardless of outcome
      If isOutlaw(player.ID) And actionSeq <> ASNavEvade Then  'it just arrived so face it
         showNav sectorID, -2
         Exit Sub
      End If
   End If
   If isShipHere(6, sectorID) = 6 And CorvetteSeq <> getCorvetteSeq Then
      CorvetteSeq = getCorvetteSeq
      'CruiserCutter = SectorID 'set it as faced regardless of outcome
      If isOutlaw(player.ID) And actionSeq <> ASNavEvade Then  'it just arrived so face it
         showNav sectorID, -3
         If Not (FullburnMovesDone = 0 And MoseyMovesDone = 0) Then 'only stop if Flying
            frmAction.moseydone = True 'Full Stop!
            frmAction.fullburndone = True
            
         End If
         Exit Sub
      End If
   End If
   
   'check active job limit not exceeded
   x = getPlayerJobs(player.ID, "1,2")
   y = MAXACTIVEJOBS + IIf(isSolid(player.ID, 8), 1, 0)
   If x > y Then
      Set frmJoSel = New frmJobSel
      frmJoSel.jobFilter = " IN (1,2)"
      frmJoSel.Caption = "Too many ACTIVE Jobs"
      frmJoSel.Label1 = "Select the jobs you want to keep.  (up to " & y & ")"
      frmJoSel.maxjobs = y
      
      frmJoSel.Show 1

      Set frmJoSel = Nothing
   End If
   'check inactive job limit not exceeded
   x = getPlayerJobs(player.ID, "0")
   If x > MAXINACTIVEJOBS Then
      Set frmJoSel = New frmJobSel
      frmJoSel.jobFilter = "=0"
      frmJoSel.Caption = "Too many INACTIVE Jobs"
      frmJoSel.Label1 = "Select the jobs you want to keep.  (up to " & MAXINACTIVEJOBS & ")"
      frmJoSel.maxjobs = MAXINACTIVEJOBS
      
      frmJoSel.Show 1

      Set frmJoSel = Nothing
   End If
   
   With frmAction
      'check if action limit reached
      If Not SoloGame And .checkNoOfActions > 1 And actionSeq <> ASNavEvade Then
         .endAction
         Exit Sub
      'check if we are currently in Fullburn/Mosey on the 2nd action to diable other actions
      ElseIf Not SoloGame And .checkNoOfActions = 1 And ((.fullburndone = False And FullburnMovesDone > 0) Or (.moseydone = False And MoseyMovesDone > 0)) Then
         onlyFullburn = True
      End If
      
      .FDPane1.InitDockHW = 200
      .FDPane1.InitDockStyle = DockToLeft
      .FDPane1.PinState = Pinned
      .FDPane1.SetLayoutReference Nothing
      
      'Header Info ===============================
      .setVisState .imgOutlaw, isOutlaw(player.ID)
      
      x = getTurnLimit(player.ID)
      .lblTurn.Caption = CStr(Logic!Gamecntr - 1) & IIf(x > 0, "/" & x, "")
      If Logic!Gamecntr >= x And x > 0 Then
         .lblTurn.ForeColor = &HFF&
      Else
         .lblTurn.ForeColor = &H3DCBFF
      End If
      
      z = countMisbehaves(player.ID)
      y = totalMisbehaves(player.ID, x)
      .lblMisbehaves = z & IIf(y > 0, "/" & y, "")
      
      If y > 0 And z >= y Then
         .lblMisbehaves.ForeColor = &H3E631
      Else
         .lblMisbehaves.ForeColor = &H3DCBFF
      End If
      
      y = countBounties(player.ID)
      .lblBounties = y & IIf(x > 0, "/" & x, "")
      If x > 0 And y >= x Then
         .lblBounties.ForeColor = &H3E631
      Else
         .lblBounties.ForeColor = &H3DCBFF
      End If
      
      x = getPlayerJobs(player.ID, "0")
      .lblInactiveJobs = x & "/" & MAXINACTIVEJOBS
      If x >= MAXINACTIVEJOBS Then
         .lblInactiveJobs.ForeColor = &HFF&
      Else
         .lblInactiveJobs.ForeColor = &H3DCBFF
      End If
      x = getPlayerJobs(player.ID, "1,2")
      y = (MAXACTIVEJOBS + IIf(isSolid(player.ID, 8), 1, 0))
      .lblActiveJobs = x & "/" & y
      If x >= y Then
         .lblActiveJobs.ForeColor = &HFF&
      Else
         .lblActiveJobs.ForeColor = &H3DCBFF
      End If
      
      .setVisState .imgBDProof, (hasShipUpgradeAttribute(player.ID, "IgnoreBreakdowns") > 0)
                  
      SQL = "SELECT * FROM Players WHERE PlayerID = " & player.ID
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         If rst!Goals = -1 Then 'fail
            .lblGoals.Caption = "fail"
         Else
            .setVisState .imgGoals, (rst!Goals > 0)
            .lblGoals.Caption = rst!Goals & "/" & varDLookup("max(Goal) AS maxgoal", "StoryGoals", "StoryID=" & Logic!StoryID, "maxgoal")
         End If
         .setVisState .imgWarrants, (rst!Warrants > 0)
         .lblWarrants = rst!Warrants
         .lblWarrants.Visible = (rst!Warrants > 0)
         .setPay rst!pay
'         .lblCash.Tag = rst!pay
'         .lblCash = "$" & .lblCash.Tag
'         If Val(.lblCash.Tag) < 200 Then
'            .lblCash.ForeColor = &HFF&
'         Else
'            .lblCash.ForeColor = &H3DCBFF
'         End If
         
         
         'load Fuel stats,
         .lblFuel.Caption = rst!fuel
         Select Case Val(.lblFuel.Caption)
         Case 0
            .lblFuel.ForeColor = &HFF&
         Case Else
            .lblFuel.ForeColor = &H3DCBFF
         End Select
         .lblParts.Caption = rst!parts
         .lblCargo.Caption = rst!cargo
         .lblContra.Caption = rst!Contraband
         .lblPassngr.Caption = rst!Passenger
         .lblFugitives.Caption = rst!Fugitive
      End If
      rst.Close
      z = CargoSpaceUsed(player.ID)
      y = CargoCapacity(player.ID)
      .lblHoldSpace = z & "/" & y
      If z >= y Then
         .lblHoldSpace.ForeColor = &HFF&
      Else
         .lblHoldSpace.ForeColor = &H3DCBFF
      End If
      
      x = CrewCapacity(player.ID)
      y = getCrewCount(player.ID)
      .lblCrewSpace = y & "/" & x
      If y > 9 Then
         .lblCrewSpace.Left = 2360
      Else
         .lblCrewSpace.Left = 2420
      End If
      If y >= x Then
         .lblCrewSpace.ForeColor = &HFF&
      Else
         .lblCrewSpace.ForeColor = &H3DCBFF
      End If
      x = getShipUpgrades(player.ID)
      .lblUpgrades = x & "/3"
      If x >= 3 Then
         .lblUpgrades.ForeColor = &HFF&
      Else
         .lblUpgrades.ForeColor = &H3DCBFF
      End If
      
      SQL = "SELECT SUM(ShipUpgrade.BurnRange) AS BurnRange, MAX(ShipUpgrade.BurnFuel) AS BurnFuel, MAX(ShipUpgrade.MoseyRange) AS MoseyRange"
      SQL = SQL & " FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID"
      SQL = SQL & " WHERE PlayerSupplies.PlayerID=" & player.ID   'ShipUpgrade.DriveCore=1 AND
      
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         
         '>>>>>  CMD  FULLBURN  <<<<

         .lblFBRange.Caption = 5 + rst!BurnRange + getRangeMod(player.ID, 1) + .rangeBoost + turnExtraRange - FullburnMovesDone  'ADD WASH's extra Range
         If Val(.lblFBRange.Caption) = 0 Then .fullburndone = True
         If Not SoloGame And .checkNoOfActions > 1 Then
            .endAction
            Exit Sub
         End If
         
         'is Full Burn available?
         x = getExtraBurn(player.ID)  ' Heavy Load??
         .setVisState .imgFlyHL, (x > 0)
         .actionButtonEnable "imgFullBurn", (((rst!burnFuel + x + IIf(.rangeBoost > 0, 1, 0)) <= Val(.lblFuel.Caption)) Or FullburnMovesDone > 0) And (Not .fullburndone) And (actionSeq = ASselect) And Not reaverActive And hasValidFBMove(player.ID) And Not (HemmorrhagingFuel And FullburnMovesDone > 0 And Val(.lblFuel.Caption) = 0)
         'single use extended Range
         .imgFlyBoost.Visible = (hasShipUpgrade(player.ID, 17) > 0)
         .actionButtonEnable "imgFlyBoost", (hasShipUpgrade(player.ID, 17) > 0 And FullburnMovesDone = 0 And .imgFullBurn.Tag = "Y" And Val(.lblFuel.Caption) >= rst!burnFuel + x + 1)  ', .rangeBoost > 0

         .lblFBFuel.Caption = rst!burnFuel + x + IIf(.rangeBoost > 0, 1, 0)
         If x > 0 Then
            .lblFBFuel.ForeColor = &HFF&
         Else
            .lblFBFuel.ForeColor = &H45A8D4
         End If
         
         '>>>>>  CMD  MOSEY  <<<<
         If hasShipUpgrade(player.ID, 7) Then
            moseyrng = 2 + getRangeMod(player.ID, 2)
         Else
            moseyrng = rst!MoseyRange + getRangeMod(player.ID, 2)
         End If
         If moseyrng > 2 Then moseyrng = 2 'set maximum possible
         .lblMRange.Caption = moseyrng - MoseyMovesDone
         If moseyrng = MoseyMovesDone Then .moseydone = True
         If Not SoloGame And .checkNoOfActions > 1 Then
            .endAction
            Exit Sub
         End If
         .actionButtonEnable "imgMosey", (moseyrng > MoseyMovesDone) And (actionSeq = ASselect) And (Not .moseydone) And Not reaverActive
         
      End If

      rst.Close
      
      'load Supply Graphic for this sector
      .setSupply sectorID
      
      '>>>>>  CMD  BUY  <<<<
      Select Case actionSeq
         Case ASBuy
            'Beep
         Case ASBuySelDiscard  '"Draw Cards"
            .setMultiStateButton "imgShop", "2"
         Case ASBuyDrew
            'should never happen, moves straight to ASBuySelect
         Case ASBuySelect '"Close Buy"
            .setMultiStateButton "imgShop", "3"

         Case Else '"Buy"
            .setMultiStateButton "imgShop", "1"

      End Select
      
      'flag to disable certain buttons
      hide = (actionSeq = ASDealSelDiscard Or actionSeq = ASDealSelect Or actionSeq = ASBuySelDiscard Or actionSeq = ASBuySelect)
      
      'SHORE LEAVE
      If (Not .buydone) And (Not onlyFullburn) And Not reaverActive Then  ' Buy and Shore leave *may* be active
         unseenDeck = getUnseenDeck("Supply", Val(.imgSupply.Tag))
         .actionButtonEnable "imgShore", Not hide And (Nz(varDLookup("SupplyID", "Supply", "SectorID=" & sectorID), 0) > 0 Or hasShipUpgrade(player.ID, 19) Or getHaven(sectorID) > 0) And hasDisgruntled(player.ID) And (Abs(doShoreLeave(player.ID, True)) <= getMoney(player.ID) Or getHaven(sectorID) = player.ID)
         .lblDisCost.Visible = Not hide And (Nz(varDLookup("SupplyID", "Supply", "SectorID=" & sectorID), 0) > 0 Or hasShipUpgrade(player.ID, 19) Or getHaven(sectorID) > 0) And hasDisgruntled(player.ID)
         .lblDisCost = "$" & Abs(doShoreLeave(player.ID, True))
         If (.imgSupply.Tag <> "") And (actionSeq = ASselect Or (actionSeq = ASBuySelDiscard And unseenDeck > 0) Or actionSeq = ASBuySelect) Then 'we can BUY
            'Enabled
         ElseIf getHaven(sectorID) > 0 And CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0 Then
            'Enabled
         ElseIf unseenDeck = 0 And Val(.imgSupply.Tag) > 0 Then 'modify to Consider
            .setMultiStateButton "imgShop", "2a"
         Else 'disable
            .setMultiStateButton "imgShop", "N"
         End If
      
      Else 'nothing is enabled
         .actionButtonEnable "imgShore", False
         .setMultiStateButton "imgShop", "N"
      End If
      .lblDisCnt = cntDisgruntled(player.ID)
      
      'FUEL & PARTS
      '.setVisState .imgFuelBuy, (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0 And (((Nz(varDLookup("SupplyID", "Supply", "SectorID=" & SectorID), 0) > 0) And (Not .buydone)) Or (Nz(varDLookup("ContactID", "Contact", "SectorID=" & SectorID), 0) = 5 And isSolid(player.ID, 5)) Or getHaven(SectorID) > 0))
      .setVisState .imgFuelBuy, (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0 And (((Nz(varDLookup("SupplyID", "Supply", "SectorID=" & sectorID), 0) > 0) And (Not .buydone)) Or getHaven(sectorID) > 0))
      .setVisState .imgPartsBuy, (((Nz(varDLookup("SupplyID", "Supply", "SectorID=" & sectorID), 0) > 0) And (Not .buydone) And CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0))
            
      'check Bree DEAL
      .setVisState .imgDealParts, (hasCrew(player.ID, 34) And varDLookup("Parts", "Players", "PlayerID=" & player.ID) > 0 And isSolid(player.ID, varDLookup("ContactID", "Contact", "SectorID=" & sectorID)))   'Bree sells parts to Solids
      
      'load Dealer in this sector
      .setContact sectorID
      
      .imgPhone.Visible = hasCrew(player.ID, 75) And Not hasCrew(player.ID, 22)
      If (.imgContact.Tag = "" And hasCrew(player.ID, 75) And Not hasCrew(player.ID, 22)) Or HigginsDealPerk Then
         .setContact 16  'sector for "Mag. Higgins" Tag = "8"
      Else
         .actionButtonEnable "imgPhone", Not Not hasCrew(player.ID, 75) And Not hasCrew(player.ID, 22) And Not hide And Not .dealdone
      End If
      
      If HarkenDeal Then
         .setContact -1 'code for "Harken"
      End If
      
      If .imgContact.Tag = "" Then  'nothing doing here

         .setVisState .imgDealCargo, False
         .setVisState .imgDealContra, False
         .setVisState .imgDealFuel, False
         
      ElseIf .imgContact.Tag = "5" Then 'harken
         
         .setVisState .imgDealCargo, False
         .setVisState .imgDealContra, False
         .setVisState .imgDealFuel, (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0 And Nz(varDLookup("ContactID", "Contact", "SectorID=" & sectorID), 0) = 5 And isSolid(player.ID, 5) And Not .dealdone And getMoney(player.ID) >= 100)
         
      ElseIf .imgContact.Tag = "6" Then 'harrow

         .setVisState .imgDealCargo, (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= 1) And isSolid(player.ID, 6) And getMoney(player.ID) >= 300
         .imgDealCargo.ToolTipText = "buy Cargo for $300ea"

         .setVisState .imgDealContra, False
         .setVisState .imgDealFuel, False

      ElseIf .imgContact.Tag = "9" Then 'fanty mingo

         .setVisState .imgDealContra, (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) >= 1) And isSolid(player.ID, 9) And getMoney(player.ID) >= 400

         .setVisState .imgDealCargo, False
         .imgDealContra.ToolTipText = "buy Contraband @ $400ea"
         .setVisState .imgDealFuel, False

      Else  'regular Contact

         .setVisState .imgDealCargo, (doSellCargoContra(player.ID, .imgContact.Tag, 1, 0, True) > 0)

         .setVisState .imgDealContra, (doSellCargoContra(player.ID, .imgContact.Tag, 0, 1, True) > 0)
         .imgDealCargo.ToolTipText = "sell Cargo to Contact"
         .imgDealContra.ToolTipText = "sell Contraband to Contact"
         .setVisState .imgDealFuel, False
      End If

         
      '>>>>>  CMD  DEAL  <<<<
      Select Case actionSeq
         Case ASDeal
            Beep
         Case ASDealSelDiscard
            .setMultiStateButton "imgDealer", "2"

         Case ASDealDrew
            'should never happen, moves straight to ASDealSelect
         Case ASDealSelect
            .setMultiStateButton "imgDealer", "3"

         Case Else
            .setMultiStateButton "imgDealer", "1"

      End Select
      
      If Not (.imgContact.Tag <> "" And (actionSeq = ASselect Or (actionSeq = ASDealSelDiscard And getUnseenDeck("Contact", Val(.imgContact.Tag)) > 0) Or actionSeq = ASDealSelect) And (Not .dealdone) And (Not onlyFullburn) And Not reaverActive And Not (Val(.imgContact.Tag) = 5 And varDLookup("Warrants", "Players", "PlayerID=" & player.ID) > 0)) Then
         .setMultiStateButton "imgDealer", "N"
      End If
      
      'Universal Encyclopedia
      .actionButtonEnable "imgRead", Not .dealdone 'And .imgDealer.Tag = "1"
      .imgRead.Visible = hasGear(player.ID, 60)
      
      'Remove Warrants with Badger
      .setVisState .imgClearWarrants, ((varDLookup("Warrants", "Players", "PlayerID=" & player.ID) > 0) And (Val(.imgContact.Tag) = 2) And isSolid(player.ID, 2) And (Val(.lblCash.Tag) > 1000))
      

      'Load WORK Combo with Make Work & Jobs in this Sector
      SQL = "SELECT ContactDeck.CardID, Job.JobID AS JOB1, Job.JobDesc AS JOBDES1, Job.SectorID AS SECTOR1, Job_1.JobID AS JOB2, Job_1.JobDesc AS JOBDES2, Job_1.SectorID AS SECTOR2, "
      SQL = SQL & "PlayerJobs.JobStatus, ContactDeck.Immoral, ContactDeck.JobName, Job_2.JobID AS JOB3, Job_2.JobDesc AS JOBDES3, Job_2.SectorID AS SECTOR3 "
      SQL = SQL & "FROM (Job INNER JOIN ((PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) LEFT JOIN "
      SQL = SQL & "Job AS Job_1 ON ContactDeck.Job2ID = Job_1.JobID) ON Job.JobID = ContactDeck.Job1ID) LEFT JOIN Job AS Job_2 ON ContactDeck.Job3ID = Job_2.JobID "

      SQL = SQL & "Where PlayerJobs.PlayerID = " & player.ID & " And (Job.SectorID IN (1,2," & sectorID & ") Or Job_1.SectorID IN (1,2," & sectorID & ") Or Job_2.SectorID IN (1,2," & sectorID & "))"
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      '.cbo.Clear
      
      For x = 1 To .mnuWorkPop.Count - 1
         Unload .mnuWorkPop(x)
      Next x
      x = -1
      While Not rst.EOF
         If ((rst!sector1 = 1 And getCruiserSector() = sectorID) Or (rst!sector1 = 2 And getCorvetteSector() = sectorID) Or (sectorID = rst!sector1)) And rst!JobStatus = 0 And getPlayerJobs(player.ID, "1,2") < MAXACTIVEJOBS + IIf(isSolid(player.ID, 8), 1, 0) Then ' check requirements met for job
            If hasJobReqs(player.ID, rst!CardID, rst!Job1) Then
               x = x + 1
               If .mnuWorkPop.Count < x + 1 Then Load .mnuWorkPop(x)
               .mnuWorkPop(x).Caption = rst!Jobdes1 & " (" & CStr(rst!CardID) & ")"
               .mnuWorkPop(x).Tag = CStr(rst!CardID)
            End If

         ElseIf ((rst!Sector3 = 1 And getCruiserSector() = sectorID) Or (rst!Sector3 = 2 And getCorvetteSector() = sectorID) Or (sectorID = rst!Sector3)) And rst!JobStatus = 1 Then 'Job3 must be in the sector
            If hasJobReqs(player.ID, rst!CardID, rst!Job3) Then
               x = x + 1
               If .mnuWorkPop.Count < x + 1 Then Load .mnuWorkPop(x)
               .mnuWorkPop(x).Caption = rst!Jobdes3 & " (" & CStr(rst!CardID) & ")"
               .mnuWorkPop(x).Tag = CStr(rst!CardID)
            End If

         ElseIf ((rst!sector2 = 1 And getCruiserSector() = sectorID) Or (rst!sector2 = 2 And getCorvetteSector() = sectorID) Or (sectorID = rst!sector2)) And (rst!JobStatus = 1 Or rst!JobStatus = 2) Then 'Job2 must be in the sector
            If hasJobReqs(player.ID, rst!CardID, rst!Job2) Then
               x = x + 1
               If .mnuWorkPop.Count < x + 1 Then Load .mnuWorkPop(x)
               .mnuWorkPop(x).Caption = rst!Jobdes2 & " (" & CStr(rst!CardID) & ")"
               .mnuWorkPop(x).Tag = CStr(rst!CardID)
            End If

         End If
         rst.MoveNext
      Wend
      rst.Close

      'do supply transfer at Haven
      If getHaven(sectorID) = player.ID And useHavenStorage(Logic!StoryID) Then
         x = x + 1
         If .mnuWorkPop.Count < x + 1 Then Load .mnuWorkPop(x)
         .mnuWorkPop(x).Caption = "Supplies Transfer at " & varDLookup("PlanetName", "Planet", "SectorID=" & sectorID)
         .mnuWorkPop(x).Tag = "-1"
      End If

      'look for Possible Bounty
      loadBounties .mnuWorkPop, Val(.imgSupply.Tag), sectorID, x
      
      'load onto the display
      If x > -1 Then
         .lblJobName = .mnuWorkPop(0).Caption
         .lblJobName.ToolTipText = .mnuWorkPop(0).Caption
         .lblJobName.Tag = .mnuWorkPop(0).Tag
         .timScroll.Enabled = True
      Else
         .lblJobName = ""
         .lblJobName.ToolTipText = "no jobs available"
         .lblJobName.Tag = ""
      End If
      .setVisState .imgWorkDrop, (.mnuWorkPop.Count > 1)
      If .mnuWorkPop.Count > 1 Then .imgWorkDrop.Animate2.StartAnimation
      
      'Make Work if at a Planet
      
      If Nz(varDLookup("PlanetID", "Planet", "SectorID=" & sectorID), 63) <> 63 And Nz(varDLookup("PlanetID", "Planet", "SectorID=" & sectorID), 64) <> 64 And (actionSeq = ASselect) And (Not .workdone) And (Not onlyFullburn) And Not reaverActive Then  'but not Cruiser/Corvette dummy planetID 63,64
         .actionButtonEnable "imgMakeWork", True
         .lblMakeWorkVal.Visible = True
         .lblMakeWorkVal.Caption = "$" & (200 + IIf(hasCrew(player.ID, 73), 100, 0))
      Else
         .actionButtonEnable "imgMakeWork", False
         .lblMakeWorkVal.Visible = False
      End If

         
      '>>>>>  CMD  WORK  <<<<
      .actionButtonEnable "imgWorkLocal", (.lblJobName.Tag <> "") And (actionSeq = ASselect) And (Not .workdone) And (Not onlyFullburn) And Not reaverActive
         
      '>>>>>  CMD  END TURN  <<<<
      .actionButtonEnable "imgEndTurn", Not hide And Not reaverActive
      
      '>>>>>> remove Disgruntled <<<<<
      .actionButtonEnable "imgMorale", Not hide And (getPerkAttributeCrew(player.ID, "RemoveDisgruntled") > 0 Or hasGear(player.ID, 27)) And hasDisgruntled(player.ID, True) And (Not .disgruntledone) And Not reaverActive
      .imgMorale.Visible = (getPerkAttributeCrew(player.ID, "RemoveDisgruntled") > 0 Or hasGear(player.ID, 27))
      
      '>>>>Resolve Alerts <<<<<<<<<<
      .actionButtonEnable "imgResolve", Not hide And (hasAdjacentAlert(player.ID) And hasShipUpgrade(player.ID, 16) > 0 And (Not .fullburndone Or Not .moseydone))
      .imgResolve.Visible = hasShipUpgrade(player.ID, 16) > 0 ' .imgResolve.Tag = "Y"
      
      'Cruiser Call by Dobson
      .actionButtonEnable "imgFlyMole", Not hide And (hasCrew(player.ID, 93) And getZone(sectorID) = "A" And FullburnMovesDone = 0 And MoseyMovesDone = 0 And (Not .fullburndone Or Not .moseydone) And isShipHere(5, sectorID) <> 5)
      .imgFlyMole.Visible = hasCrew(player.ID, 93)
      
      '>>>>>> load Passengers & Fugitives at Amnon's <<<<<
      .setVisState .imgLoadPassngr, (sectorID = 23) And isSolid(player.ID, 1) And .imgDealer.Tag <> "N" And (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0.6)
      .setVisState .imgLoadFugi, .imgLoadPassngr.Tag = "Y"

      .setVisState .imgFly, Not (.fullburndone And .moseydone)
      .setVisState .imgBuy, Not .buydone
      .setVisState .imgDeal, Not .dealdone
      .setVisState .imgWork, Not .workdone
      .setVisState .imgBonus, (.imgMorale.Tag = "Y" Or .imgResolve.Tag = "Y")
      
      .Visible = True
      .FDPane1.PaneVisible = True
   End With
   
   refreshAll
   
   Set rst = Nothing
End Sub

Public Sub refreshAll()
   If Not (frmShip Is Nothing) Then frmShip.RefreshShips
   If Not (frmJob Is Nothing) Then frmJob.refreshJobs
   If Not (frmBuy Is Nothing) Then frmBuy.RefreshBuys
   RefreshDeals
End Sub

Public Sub RefreshDeals()
   If Not (frmDeal Is Nothing) And actionSeq <> ASDealSelect And actionSeq <> ASDealSelDiscard Then frmDeal.RefreshDeals
End Sub


'returns doWork = 0 Normal, 1= Evade
Public Function doWork(ByVal playerID, ByVal CardID) As Integer
Dim rst As New ADODB.Recordset, x, parts As Integer, a() As String, DoubleDown As Integer
Dim SQL, sectorID, ContactID, JobID, finalstate, result As Integer, misbehaveNum, bonus, cargofit As Integer, fugifit As Integer, cargopay As Integer
Dim frmCrew As frmCrewLst, Dice As Integer, payment As Integer, KeywordInUse As Boolean
Dim skillcnt, payCrewTotal As Integer, WSkill As Integer
Dim frmSalvage As frmSalvaging, frmKillCrw As frmKillCrew, frmGamb As frmGamble, solidMsg As String

   sectorID = varDLookup("SectorID", "Players", "PlayerID=" & playerID)
   ContactID = varDLookup("ContactID", "ContactDeck", "CardID=" & CardID)
   usedStitchSkill = False
   DoubleDown = 1
   
   SQL = "SELECT ContactDeck.CardID, ContactDeck.Bonus, Job.JobID AS JOB1, Job.SectorID AS SECTOR1, Job_1.JobID AS JOB2, Job_1.SectorID AS SECTOR2, "
   SQL = SQL & "PlayerJobs.JobStatus, ContactDeck.Illegal, ContactDeck.Immoral, ContactDeck.JobName, Job_2.JobID AS JOB3, Job_2.SectorID AS SECTOR3 "
   SQL = SQL & "FROM (Job INNER JOIN ((PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) LEFT JOIN "
   SQL = SQL & "Job AS Job_1 ON ContactDeck.Job2ID = Job_1.JobID) ON Job.JobID = ContactDeck.Job1ID) LEFT JOIN Job AS Job_2 ON ContactDeck.Job3ID = Job_2.JobID "
   SQL = SQL & "WHERE PlayerJobs.PlayerID= " & playerID & " AND ContactDeck.CardID=" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      If rst!JobStatus = 0 And ((rst!sector1 = 1 And getCruiserSector() = sectorID) Or (rst!sector1 = 2 And getCorvetteSector() = sectorID) Or (sectorID = rst!sector1)) Then ' we're doing Job 1
      
         JobID = rst!Job1
         If IsNull(rst!Job2) Then
            finalstate = JOB_SUCCESS
         Else
            finalstate = 1
         End If
         PutMsg player.PlayName & " Started Job: " & rst!JobName, playerID, Logic!Gamecntr
         
      ElseIf rst!JobStatus = 1 And ((rst!Sector3 = 1 And getCruiserSector() = sectorID) Or (rst!Sector3 = 2 And getCorvetteSector() = sectorID) Or (sectorID = rst!Sector3)) And Not IsNull(rst!Job3) Then ' we're doing Bonus Job
         JobID = rst!Job3
         bonus = rst!bonus
         finalstate = 2
         
      ElseIf (rst!JobStatus = 1 Or rst!JobStatus = 2) And ((rst!sector2 = 1 And getCruiserSector() = sectorID) Or (rst!sector2 = 2 And getCorvetteSector() = sectorID) Or (sectorID = rst!sector2)) And Not IsNull(rst!Job2) Then  ' we're doing Job 2
         JobID = rst!Job2
         finalstate = JOB_SUCCESS
         
      Else
         MsgBox "Job Card " & CardID & " Error for Player " & playerID, vbCritical
         Exit Function
      End If
   Else
      MsgBox "Job Card " & CardID & " Error for Player " & playerID, vbCritical
      Exit Function
   End If
   
   If finalstate = JOB_SUCCESS And rst!Immoral = 1 And hasDisgruntled(playerID, False, True) Then
      If MessBox("Warning: Completing this Immoral Job with Disgruntle Moral Crew will likely result in Crew leaving." & vbNewLine & "Are you sure you want to continue?", "Immoral Job Consequences", "Don't Care", "Oh No", getLeader()) = 1 Then
         frmAction.workdone = False
         Exit Function
      End If
   End If
   'check for Shepherd  (93 / 47) and keep him off immoral jobs
   If rst!Immoral = 1 And hasCrew(playerID, 47) Then
      DB.Execute "UPDATE PlayerSupplies SET OffJob = 1 WHERE CardID = 93"
      If Not (frmShip Is Nothing) Then frmShip.RefreshShips  'update display
      PutMsg player.PlayName & "'s Work log: Shepherd's having none of this Immoral Job citing '..a special place in Hell!'", playerID, Logic!Gamecntr, True, 47
   End If
   
   If rst!illegal = 1 Then
      doLawmenOffJob
   End If
   
   rst.Close
   
   'assume we now have a JobID to carry out in the current Sector
   SQL = "SELECT * FROM Job WHERE JobID = " & JobID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
   
      If rst!misbehave <> 0 Then
         misbehaveNum = rst!misbehave
         'if TwoFry has a Sniper Rifle, then one less misbv
         If hasGearKeyword(playerID, "SNIPERRIFLE", 56) And misbehaveNum > 1 Then
            misbehaveNum = misbehaveNum - 1
            PutMsg player.PlayName & "'s Two-Fry used his DeadEye Sniper skills to good effect and eliminates 1 misbehave", playerID, Logic!Gamecntr, True, 56
         End If
         'go do the number of misbehaves
         result = doMisbehaves(playerID, misbehaveNum, sectorID)
         Select Case result
         Case 1, 4 'proceed
            result = 0 'reset as Win for below tests
         Case 2 'botched
            PutMsg player.PlayName & "'s Work log: Job was Botched!", playerID, Logic!Gamecntr, True, getLeader()
            clearPicMB
            Exit Function
            
         Case 3 'discard Job and clear solid, with Contact
            'and if Niska - Kill 1 Crew
            'add a Warrant and clear any Solid with Harken (5)
            doJobWarrant playerID, ContactID, CardID
            If Not (frmJob Is Nothing) Then frmJob.refreshJobs
            Main.drawLine 0, -1
            clearPicMB
            Exit Function
            
         Case 5 'double down success
            result = 0 'reset as Win for below tests
            If rst!DoubleDown > 0 Then
               DoubleDown = 2
               'payment = payment + rst!DoubleDown
               PutMsg player.PlayName & " scored the Double-Down bonus", playerID, Logic!Gamecntr
            End If
         
         End Select
      End If
      clearPicMB
   
      'pickup/drop off Cargo, etc
      If rst!cargo <> 0 Then
         If (rst!cargo * -1) > varDLookup("Cargo", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
            MessBox "Not enough cargo to meet the quota, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
         
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= rst!cargo Then 'we have room
            DB.Execute "UPDATE Players SET Cargo = Cargo + " & rst!cargo & " WHERE PlayerID = " & playerID
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      
      If rst!Contraband <> 0 Then
         If rst!Contraband = 14 Then
            If Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID)) > 0 Then
               cargofit = InputBoxx("How much Contraband do you want to load onboard?", "Load Contraband", CStr(Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID))), getLeader())
            End If
         ElseIf rst!Contraband = -14 Then
            cargofit = InputBoxx("How much Contraband do you want to deliver?", "Deliver Contraband", varDLookup("Contraband", "Players", "PlayerID=" & playerID), getLeader())
            If cargofit > varDLookup("Contraband", "Players", "PlayerID=" & playerID) Then
               cargofit = varDLookup("Contraband", "Players", "PlayerID=" & playerID)
            End If
            cargopay = cargofit * 500
            cargofit = cargofit * -1
         Else
            cargofit = rst!Contraband
            If (cargofit * -1) > varDLookup("Contraband", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
               MessBox "Not enough Contraband to meet the quota, Job botched", "Job Requirements", "Ooops", "", getLeader()
               Exit Function
            End If
         End If
         
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= cargofit Then 'we have room
            DB.Execute "UPDATE Players SET Contraband = Contraband + " & cargofit & " WHERE PlayerID = " & playerID
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job botched.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      
      End If
      
      If rst!Passenger <> 0 Then
         If rst!Passenger = 14 Then
            If Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID)) > 0 Then
               cargofit = InputBoxx("How many Passengers do you want to take onboard?", "Load Passengers", CStr(Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID))), getLeader())
            End If
         ElseIf rst!Passenger = -14 Then
            cargofit = InputBoxx("How many Passengers do you want to deliver?", "Deliver Passengers", varDLookup("Passenger", "Players", "PlayerID=" & playerID), getLeader())
            If cargofit > varDLookup("Passenger", "Players", "PlayerID=" & playerID) Then
               cargofit = varDLookup("Passenger", "Players", "PlayerID=" & playerID)
            End If
            cargopay = cargofit * IIf(rst!Fugitive = -14, 300, 200) 'chk being sold as Fugitives
            cargofit = cargofit * -1
            
         Else
            cargofit = rst!Passenger
            If (cargofit * -1) > varDLookup("Passenger", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
               MessBox "Not enough Passengers to meet the quota, Job botched", "Job Requirements", "Ooops", "", getLeader()
               Exit Function
            End If
         End If
                      
         'pay at the end, passing cargofit
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= cargofit Then 'we have room
            SQL = "UPDATE Players SET Passenger = Passenger + " & cargofit
            SQL = SQL & " WHERE PlayerID = " & playerID
            DB.Execute SQL
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, Job botched", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      If rst!Fugitive <> 0 Then
         If rst!Fugitive = 14 Then
         
            If Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID)) > 0 Then
               fugifit = InputBoxx("How many Fugitives do you want to take onboard?", "Load Fugitives", CStr(Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID))), getLeader())
            End If

         ElseIf rst!Fugitive = -14 Then
            fugifit = InputBoxx("How many Fugitives do you want to deliver?", "Deliver Fugitives", varDLookup("Fugitive", "Players", "PlayerID=" & playerID), getLeader())
            If fugifit > varDLookup("Fugitive", "Players", "PlayerID=" & playerID) Then
               fugifit = varDLookup("Fugitive", "Players", "PlayerID=" & playerID)
            End If
            cargopay = cargopay + fugifit * 300  'may get paid for pasngrs too
            fugifit = fugifit * -1
            
            
         Else
            fugifit = rst!Fugitive
            If (fugifit * -1) > varDLookup("Fugitive", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
               MessBox "Not enough Fugitives to meet the quota, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
               Exit Function
            End If
         End If
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= fugifit Then 'we have room
            DB.Execute "UPDATE Players SET Fugitive = Fugitive + " & fugifit & " WHERE PlayerID = " & playerID
            If fugifit < 0 Then
               If beaDirtySlaver(playerID) Then
                  cargopay = cargopay + Abs(fugifit) * 100
               End If
            End If
            
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      If rst!fuel <> 0 Then
         If (rst!fuel * -1) > varDLookup("Fuel", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
            MessBox "Not enough Fuel to meet the quota, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= (rst!fuel / 2) Then 'we have room
            DB.Execute "UPDATE Players SET Fuel = Fuel + " & rst!fuel & " WHERE PlayerID = " & playerID
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      If rst!parts <> 0 Then
         If (rst!parts * -1) > varDLookup("Parts", "Players", "PlayerID=" & playerID) Then 'you don't have the required amount
            MessBox "Not enough Parts to meet the quota, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
         
         If CargoCapacity(playerID) - CargoSpaceUsed(playerID) >= (rst!parts / 2) Then 'we have room
            DB.Execute "UPDATE Players SET Parts = Parts + " & rst!parts & " WHERE PlayerID = " & playerID
         Else 'abort the Job
            PutMsg player.PlayName & " doesn't have enough Cargo space, Job aborted.", playerID, Logic!Gamecntr
            MessBox "Not enough cargo space in the hold, aborting the Job", "Job Requirements", "Ooops", "", getLeader()
            Exit Function
         End If
      End If
      
      'TAG and BAG
      If rst!tagnbag > 0 Then  ' And CargoCapacity(playerID) > CargoSpaceUsed(playerID) Then 'load to your capacity
         If rst!tagnbag = 1 Then
            skillcnt = getSkill(playerID, cstrSkill(2), 0, False) + RollDice(6)
            PutMsg player.PlayName & " Tech Test comes to " & skillcnt & " for the Tag and Bag", playerID, Logic!Gamecntr
            Select Case skillcnt
            Case 1 - 4
               x = 3
            Case 5 - 7
               x = 6
            Case Else
               x = 20
            End Select
         ElseIf rst!tagnbag = 20 Then
            x = 20
            PutMsg player.PlayName & " does a Tag and Bag to grab some goods", playerID, Logic!Gamecntr
         Else
            x = rst!tagnbag
         End If
         If frmSalvage Is Nothing Then
            Set frmSalvage = New frmSalvaging
         End If
         frmSalvage.mode = 2
         frmSalvage.salvageCount = x
         frmSalvage.Show 1
         
      End If
      
      
   End If
   
   rst.Close
   
   If finalstate = 1 Then

      PutMsg player.PlayName & " has completed the first Work Part of " & varDLookup("JobName", "ContactDeck", "CardID=" & CardID) & " at " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & sectorID)), playerID, Logic!Gamecntr, True, getLeader()
      
   ElseIf finalstate = 2 Then 'Bonus Job done
      'Pay Bonus
      bonus = bonus + cargopay
      getMoney playerID, bonus
      PutMsg player.PlayName & " has completed the $" & bonus & " Bonus Work Part of " & varDLookup("JobName", "ContactDeck", "CardID=" & CardID) & " at " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & sectorID)), playerID, Logic!Gamecntr, True, getLeader()

   ElseIf finalstate = JOB_SUCCESS Then 'job is ending, but do any remaining challenges Negotiate Pay or Cover your Tracks, Gamble
      'Open the Contact Deck to access end of Job info and flags
      SQL = "SELECT * FROM ContactDeck WHERE CardID=" & CardID
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         If rst!RemoveDisgruntled <> 0 Then
            doDisgruntled playerID, 3
         End If
         'KEYWORD CHECKS - ========================
         If rst!WinOptKeyword > 0 Then ' give option to use Keyword to Win -  only applies to HACKINGRIG or Explosives (getting paid half)
               
            'check if the keyword was single use, and discard
            
            If discardGearKeyword(playerID, rst!KeyWords, True) Then
               If MessBox("In this final Work Challenge, do you want to use your discardable " & rst!KeyWords & " instead of the Skill Test?" & IIf(rst!KeyWords = "EXPLOSIVES", vbNewLine & "This would result in Half Pay.", ""), "Final Work Challenge", "Do It", "Nope", getLeader()) = 0 Then
                  discardGearKeyword playerID, rst!KeyWords
                  KeywordInUse = True
                  If rst!KeyWords = "EXPLOSIVES" Then
                     result = 2 'half pay
                  Else
                     result = 0
                  End If
               End If
            Else
               If MessBox("In this final Work Challenge, do you want to use your " & rst!KeyWords & " instead of the Skill Test?" & IIf(rst!KeyWords = "EXPLOSIVES", vbNewLine & "This would result in Half Pay.", ""), "Final Work Challenge", "Do It", "Nope", getLeader()) = 0 Then
                  KeywordInUse = True
                  If rst!KeyWords = "EXPLOSIVES" Then
                     result = 2 'half pay
                  Else
                     result = 0
                  End If
               End If
            End If
             
         ElseIf rst!WinOptKeyword = 0 And rst!KeywordBonus = 0 And Not IsNull(rst!KeyWords) And Not (rst!KeywordOrSkill > 0 And hasCrewAttribute(playerID, cstrProfession(rst!RequireProfession))) Then   'check for discard
            a = Split(rst!KeyWords, " ")
            For x = LBound(a) To UBound(a)
               If discardGearKeyword(playerID, a(x), True) Then
                  MessBox "Discarding spent " & a(x), "Job Required Gear Keyword", "OK", "", getLeader()
                  discardGearKeyword playerID, a(x)
               End If
            Next x
            
         End If
         
         WSkill = rst!skill
         If WSkill > 0 And Not KeywordInUse Then 'we have a skill test
            result = doWorkSkillTest(Dice, WSkill, rst!win, rst!Intermediate)
         End If  'end of the skill Tests

         Select Case result
         Case 0 'win results
            Select Case rst!WinResult
            Case 1 'passngr  - now done above using job's passngr count -14
            
            Case 2 'fugi - now done above using job's fugi count -14
               
            Case 3 ' 1 cargo per Crew on Job
               skillcnt = Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID))
               'limit to what we can fit, or the number of crew on job
               If skillcnt >= getCrewCount(playerID, True) Then skillcnt = getCrewCount(playerID, True)
               If skillcnt <> 0 Then
                  DB.Execute "UPDATE Players Set Cargo = Cargo + " & skillcnt & " WHERE PlayerID = " & playerID
                  PutMsg player.PlayName & IIf(skillcnt > 0, " scored ", " lost ") & skillcnt & " Cargo", playerID, Logic!Gamecntr
               End If
               
            Case 4 'move Cruiser to sector and EVADE - work done
               MoveShip 5, sectorID
               If getHaven(sectorID) > 0 Then
                  PutMsg player.PlayName & "'s Nav log: refuge found at this Haven, the Alliance Cruiser sails on by", playerID, Logic!Gamecntr, True, 0, 0, 1
                  moveAutoAI 5
               Else
                  PutMsg player.PlayName & " needs to EVADE!", playerID, Logic!Gamecntr, True
                  actionSeq = ASNavEvade
                  doWork = 1 ' Evade
               End If
               
               
            Case 5 ' 1 contraband per Crew on Job
               skillcnt = Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID))
               'limit to what we can fit, or the number of crew on job
               If skillcnt >= getCrewCount(playerID, True) Then skillcnt = getCrewCount(playerID, True)
               If skillcnt <> 0 Then
                  DB.Execute "UPDATE Players Set Contraband = Contraband + " & skillcnt & " WHERE PlayerID = " & playerID
                  PutMsg player.PlayName & IIf(skillcnt > 0, " scored ", " lost ") & skillcnt & " Contraband", playerID, Logic!Gamecntr
               End If
               
            Case 6 'kill a merc
               Set frmKillCrw = New frmKillCrew
         
               frmKillCrw.nbrSelect = 1
               frmKillCrw.extrafilter = " AND Crew.Merc = 1"
               frmKillCrw.Show 1
               Set frmKillCrw = Nothing
               
            Case 7 ' EVADE - work done
               PutMsg player.PlayName & " needs to EVADE!", playerID, Logic!Gamecntr, True
               actionSeq = ASNavEvade
               doWork = 1 ' Evade
               
            Case 8 'move Corvette to sector and EVADE - work done
               MoveShip 6, sectorID
               PutMsg player.PlayName & " needs to EVADE!", playerID, Logic!Gamecntr, True
               actionSeq = ASNavEvade
               doWork = 1 ' Evade
              
               
            Case Is > 99 'paid extra
               payment = payment + rst!WinResult
            
            Case Is < 0 'make payment

               If getMoney(playerID) >= Abs(rst!WinResult) Then
                  payment = payment + rst!WinResult
               Else
                  PutMsg player.PlayName & " doesn't have enough Money to pay the $" & CStr(Abs(rst!WinResult)) & " fee, Job botched.", playerID, Logic!Gamecntr, True, getLeader()
                  Exit Function 'fail
               End If

            End Select
            
            
         Case 1 'inter
            'no change
            Select Case rst!IntermediateResult
            Case 4 'attempt botched
               PutMsg player.PlayName & " botches the final part of the job", playerID, Logic!Gamecntr, True, getLeader()
               Exit Function
            
            Case 5 'move Cruiser to sector and EVADE - work done
               MoveShip 5, sectorID
               If getHaven(sectorID) > 0 Then
                  PutMsg player.PlayName & "'s Nav log: refuge found at this Haven, the Alliance Cruiser sails on by", playerID, Logic!Gamecntr, True, 0, 0, 1
                  moveAutoAI 5
               Else
                  PutMsg player.PlayName & " needs to EVADE!", playerID, Logic!Gamecntr, True
                  actionSeq = ASNavEvade
                  doWork = 1 ' Evade
               End If
            Case 6 'move Corvette to sector and EVADE - work done
               MoveShip 6, sectorID
               PutMsg player.PlayName & " needs to EVADE!", playerID, Logic!Gamecntr, True
               actionSeq = ASNavEvade
               doWork = 1 ' Evade
               
            Case Is < -99 'make payment
               If getMoney(playerID) >= (rst!IntermediateResult * -1) Then
                  payment = payment + rst!IntermediateResult
               Else
                  PutMsg player.PlayName & " doesn't have enough Money to pay the $" & CStr(Abs(rst!IntermediateResult)) & " fee, Job botched.", playerID, Logic!Gamecntr, True, getLeader()
                  Exit Function 'fail
               End If
            End Select
   
         Case 2 'half pay
            'handled below
            
         Case 3 ' lose results -0 = continue
            Select Case rst!FailResult
            Case 1  'lose rep only
               If rst!FailLoseRep > 0 Then
                  If Not discardRoberta(playerID) Then
                     DB.Execute "UPDATE Players SET Solid" & rst!FailLoseRep & "=0 WHERE PlayerID =" & playerID
                     PutMsg player.PlayName & " loses any Rep with " & varDLookup("ContactName", "Contact", "ContactID=" & rst!FailLoseRep), playerID, Logic!Gamecntr, True, 0, 0, 0, rst!FailLoseRep
                  End If
               End If
               
            Case 2 'warrant issued - attempt botched
               doJobWarrant playerID, ContactID, CardID
               Main.drawLine 0, -1
               Exit Function
                              
            Case 3 ' pay 1000 attempt botched
               If getMoney(playerID) >= 1000 Then
                  getMoney playerID, -1000
                  PutMsg player.PlayName & " botches the final negotiation and the job, and loses $1000", playerID, Logic!Gamecntr, True, getLeader()
                  Exit Function 'fail
               Else 'take it all
                  getMoney playerID, -1 * getMoney(playerID)
                  PutMsg player.PlayName & " botches the final negotiation and the job, and loses all their money", playerID, Logic!Gamecntr, True, getLeader()
                  Exit Function 'fail
               End If
               
            Case 4 'attempt botched
               PutMsg player.PlayName & " botches the final negotiation and the job", playerID, Logic!Gamecntr, True, getLeader()
               Exit Function
             
            Case 5, 6 'move Cruiser/Corvette to sector
               PutMsg player.PlayName & " botches the job and has attracted some attention from the Alliance", playerID, Logic!Gamecntr, True, getLeader()
               MoveShip rst!FailResult, sectorID
               Exit Function
               
            Case Is < -99
               If getMoney(playerID) + payment < Abs(rst!FailResult) Then
                  payment = getMoney(playerID) * -1
               Else
                  payment = payment + rst!FailResult
               End If
            End Select
            
            'ignore for Niska + Warrant, as this already applies
            If rst!FailKillCrew > 0 And Not (ContactID = 3 And rst!FailResult = 2) Then
               doKillCrews playerID, rst!FailKillCrew, True
            End If
   
         End Select
      
         '-----------------------------------------
         'if complete - finish up
         If rst!Immoral = 1 Then
            doDisgruntled playerID, 1
            PutMsg player.PlayName & " does an immoral Job and any Moral Crew will be disgruntled", playerID, Logic!Gamecntr
         End If
   
         If ContactID = 0 Then  ' a Goal job
            payCrewTotal = 0
         ElseIf getCrewCount(playerID) > 1 Then 'someone to pay?
            'show a list of Crew to choose who gets paid, then return deduct amt
            Set frmCrew = New frmCrewLst
            frmCrew.noMoralDisgruntle = (rst!GoodDeeds = 1)
            frmCrew.Label1 = "Job Pay: " & "$" & rst!pay & "  " & IIf(rst!BonusPart > 0, " +" & rst!BonusPart & " part: ", "") & IIf(rst!bonus > 0, " +$" & rst!bonus & ":", "") & _
               IIf(rst!KeywordBonus = 1, rst!KeyWords, "") & IIf(rst!ProfessionID = 0, "", " " & cstrProfession(rst!ProfessionID)) & IIf(rst!BonusPerSkill > 0, " /" & cstrSkill(rst!BonusPerSkill), "") & _
               IIf(rst!Job3ID > 0, "Bonus Job", "") & IIf(payment > 0, "  plus $" & CStr(payment), "") & IIf(payment < 0, "  minus $" & CStr(Abs(payment)), "") & _
               IIf(getPerkAttributeSum(playerID, "BountyBonus") > 0 And ContactID = 10, "  Bounty bonus $" & getPerkAttributeSum(playerID, "BountyBonus"), "")
            frmCrew.cmd.Caption = "Pay"
            frmCrew.Show 1
            payCrewTotal = frmCrew.payTotal
         End If
         '<<<< Get Paid, with any Bonus, less deductions & Go Solid with you Contact >>>>
         SQL = "UPDATE Players SET "
         bonus = getJobBonus(playerID, CardID, parts)
         If parts > 0 Then
            SQL = SQL & "Parts= Parts + " & CStr(parts) & ", "
         End If
         'final pay with Leader Perk Bonus, and on the job profession bonus added, less crew hire
         'if result = 1 Then 'half pay
         bonus = (rst!pay * IIf(result = 2, 0.5, 1) * DoubleDown) + getJobCrewBonus(playerID, rst!JobTypeID) + getJobCrewBonus(playerID, rst!JobType2D) + bonus + payment + cargopay - payCrewTotal
         SQL = SQL & "Pay = Pay + " & bonus
         
         'are we Solid?
         If ContactID = 5 And hasWarrant(playerID) Then
            solidMsg = ", yet cannot be Solid with Harken due to outstanding Warrants"
         ElseIf ContactID > 0 And ContactID < 10 Then
            SQL = SQL & ", Solid" & ContactID & "=1 "  'setting SOLID with the Contact
            solidMsg = " and is solid with " & varDLookup("ContactName", "Contact", "ContactID=" & rst!ContactID)
         Else
            solidMsg = ""
         End If
         
         SQL = SQL & " Where playerID = " & playerID
         DB.Execute SQL
         
         If hasGear(playerID, 31) And (rst!JobTypeID = 1 Or rst!JobType2D = 1) Then 'MF-813 Flying Mule After completing a Crime Job, Load 6 Goods, minus 1 per Crew Working the Job.
            x = getCrewCount(playerID, True)
            If x < 6 Then
               If frmSalvage Is Nothing Then
                  Set frmSalvage = New frmSalvaging
               End If
               frmSalvage.mode = 2
               frmSalvage.salvageCount = (6 - x)
               frmSalvage.Show 1
               PutMsg player.PlayName & " uses the MF-813 Flying Mule to grab some goods (" & CStr(6 - x) & ")", playerID, Logic!Gamecntr, True, 0, 31
            End If
         End If
         If hasGear(playerID, 7) And (rst!JobTypeID = 1 Or rst!JobType2D = 1) Then '4WD Mule, after completing a Crime Job, Load 1 Cargo
            If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0.6 Then
               DB.Execute "UPDATE Players Set Cargo = Cargo + 1 WHERE PlayerID =" & playerID
               PutMsg player.PlayName & "'s 4WD Mule adds 1 Cargo", playerID, Logic!Gamecntr, True, 0, 7
            End If
         End If
         'do this last as Cargo may have changed.
         If hasShipUpgrade(playerID, 21) And (rst!JobTypeID = 1 Or rst!JobType2D = 1) Then
            PutMsg player.PlayName & "'s Hydraulic Docking Clamps can grab Salvage for this Crime Job", playerID, Logic!Gamecntr
            doSalvage playerID
         ElseIf (rst!JobTypeID = 6 Or rst!JobType2D = 6) Then ' SalvageOps (+ Hydraulic Docking Clamps/Crime)
            doSalvage playerID
         End If
         
         PutMsg player.PlayName & IIf(ContactID = 10, " cashed in the Bounty for the ", " completed the Job: ") & rst!JobName & " for $" & Abs(bonus) & IIf(bonus > 0, " profit", " loss") & IIf(parts > 0, ", picks up " & parts & " part" & IIf(parts > 1, "s,", ","), "") & solidMsg, playerID, Logic!Gamecntr, True, 0, 0, 0, rst!ContactID
         refreshSolid
                  
         'Gamble
         If InStr(rst!JobOrder & "", "Gamble") > 0 And getMoney(playerID) >= 500 Then
            If MessBox(varDLookup("ContactName", "Contact", "ContactID=" & rst!ContactID) & " is offering to Gamble $500 of the takings. Guess the next Card Suit and Win $3000" & vbNewLine & "Wanna Gamble?", "Gamble", "You're On", "I'll sit", 0, 0, 0, 0, rst!ContactID) = 0 Then
               Set frmGamb = New frmGamble
               frmGamb.Show 1
               x = 0
               While x = 0
                  x = doGamble() 'return suit No.
               Wend
               If frmGamb.mySuit = x Then
                  DB.Execute "UPDATE Players Set Pay = Pay + 3000 WHERE PlayerID =" & playerID
                  PutMsg player.PlayName & " gambles and Wins $3000", playerID, Logic!Gamecntr, True, getLeader()
               Else
                  DB.Execute "UPDATE Players Set Pay = Pay - 500 WHERE PlayerID =" & playerID
                  PutMsg player.PlayName & " gambles and loses $500.  It was a " & IIf(x = 1, "Spade", IIf(x = 2, "Club", IIf(x = 3, "Diamond", "Heart"))), playerID, Logic!Gamecntr, True, 0, 0, 0, rst!ContactID
               End If
               
            End If
         End If
         
      End If 'close off record
      rst.Close
      
   End If 'end of finalstate results check 1/2/3

   'update the status of the job
   DB.Execute "UPDATE PlayerJobs SET JobStatus =" & finalstate & " WHERE PlayerID = " & playerID & " AND CardID = " & CardID
   If ContactID = 10 Then frmAction.lblBounties = CStr(countBounties(playerID))
   If Not (frmJob Is Nothing) Then
      frmJob.refreshJobs
   End If
   
   Main.drawLine 0, -1

   'if we got this far, we're good!
   'doWork = 0 only set this to a value above to change the exit actionSeq behavior.  1=EVADE
   
   Set rst = Nothing
End Function

'opt 1,2, 3-ace in hole
Private Function getMisbehave(opt, cnt, total, suit, dalin As Boolean) As Integer
Dim SQL, reshuffle, ace As String
Dim rst As New ADODB.Recordset
Dim frmMB As New frmMisbehave

   With frmMB

      .MBCardID = 0
      .MBOption = 0

      'Read in the next NAV card and display either 1 or 2 options
      SQL = "SELECT MisbehaveDeck.CardID, MisbehaveDeck.CardName, MisbehaveDeck.Reshuffle, MisbehaveDeck.Seq, MisbehaveDeck.Suit, "
      SQL = SQL & "MisbehaveDeck.Keyword AS Keywords , MisbehaveDeck.CrewID, MisbehaveDeck.GearID, MisbehaveDeck.ProfessionID AS ProfesionID, MisOption.* "
      SQL = SQL & "FROM MisOption INNER JOIN MisbehaveDeck ON MisOption.OptionID = MisbehaveDeck.Option1ID "
      SQL = SQL & "Where MisbehaveDeck.Seq > 5 "
      SQL = SQL & "ORDER BY MisbehaveDeck.Seq"
      If Left(datab, 16) = "Provider=MSDASQL" Then SQL = SQL & " LIMIT 1"
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      If Not rst.EOF Then
         suit = rst!suit
         'throw the suit up on the toolbar
         picMB(cnt).Picture = LoadPicture(App.Path & "\pictures\smsuit" & suit & ".bmp")
         picMB(cnt).Visible = True
         'pull the card out of the deck
         DB.Execute "UPDATE MisbehaveDeck SET Seq = 5 WHERE CardID = " & CStr(rst!CardID)
         'rst!Seq = 5
         'rst.Update
         reshuffle = rst!reshuffle
         If reshuffle = 1 Then 'ready for next turn
            PutMsg player.PlayName & " Reshuffling MisbehaveDeck due to " & rst!CardName, player.ID, Logic!Gamecntr, True, getLeader()
            ShuffleDeck "Misbehave"
         End If
         
         'check for an ACE in the HOLE
         If Nz(rst!KeyWords, "") <> "" Then
            ace = rst!KeyWords
            If hasKeyword(player.ID, rst!KeyWords) Then
               'check if the keyword was single use, and discard
               If discardGearKeyword(player.ID, rst!KeyWords, True) Then
                  .setAce ace, ace
               Else 'just your regular multi-use Keyword
                  .setAce ace
               End If
            Else
               .setAcelbl ace
            End If
         End If
         If rst!CrewID <> 0 Then
            If hasCrew(player.ID, rst!CrewID) Then
               .setAce getCrewName(0, rst!CrewID)
            Else
               .setAcelbl getCrewName(0, rst!CrewID)
            End If
         End If
         If rst!GearID <> 0 Then
             If hasGear(player.ID, rst!GearID) Then
               .setAce getGearName(0, rst!GearID)
            Else
               .setAcelbl getGearName(0, rst!GearID)
            End If
         End If
         If rst!ProfesionID <> 0 Then
             If hasCrewAttribute(player.ID, cstrProfession(rst!ProfesionID)) Then
               .setAce cstrProfession(rst!ProfesionID)
            Else
               .setAcelbl cstrProfession(rst!ProfesionID)
            End If
         End If
         
         If Not .hasAce And hasGear(player.ID, 33) Then
            .setAce "Operative's Sword", "", True
         End If
         
         'other, soldier on..
         .MBCardID = rst!CardID
         .cmd(0).Enabled = True
         
         If Nz(rst!keyword, "") <> "" Then
            .lblKey(0).Caption = rst!keyword
            .lblKey(0).Visible = True
            If Not hasKeyword(player.ID, rst!keyword) And rst!WinOptKeyword = 0 And rst!KeywordOrSkill = 0 Then
               .cmd(0).Enabled = False
            End If
         End If
         If rst!Disgruntled = -1 And hasDisgruntled(player.ID) Then
            .cmd(0).Enabled = False
         End If
         'if this option requires Cargo, check we have enough to honour it
         If rst!cargo < 0 And varDLookup("Cargo", "Players", "PlayerID=" & player.ID) < Abs(rst!cargo) Then
            .cmd(0).Enabled = False
         End If
               
         .lblName.Caption = rst!CardName
         .cmd(0).Caption = rst!OptionName
         .cmd(0).ToolTipText = rst!OptionName
         .lblDetail(0).Caption = Replace(rst!Details, "^", vbNewLine)
         
         If rst!skill = 0 Then
            .SkillImg(0).Visible = False
         Else
            .SkillImg(0).Picture = LoadPictureGDIplus(App.Path & "\Pictures\skill" & rst!skill & ".bmp")
            .SkillImg(0).Visible = True
            .SkillImg(0).TransparentColor = &H0
            .SkillImg(0).TransparentColorMode = lvicUseTransparentColor
         End If

      Else
         PutMsg player.PlayName & " Reshuffling MisbehaveDeck due to end of deck", player.ID, Logic!Gamecntr, True, getLeader()
         ShuffleDeck "Misbehave"
         Exit Function
      End If
      rst.Close
      
      '-- Option 2--------------------------------------------------------------
      SQL = "SELECT MisOption.* "
      SQL = SQL & "FROM MisOption INNER JOIN MisbehaveDeck ON MisOption.OptionID = MisbehaveDeck.Option2ID "
      SQL = SQL & "Where MisbehaveDeck.CardID = " & .MBCardID
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      If Not rst.EOF Then
         '.cmd(1).Visible = True
         '.cmd(1).Enabled = True
         If Nz(rst!keyword, "") <> "" Then
            .lblKey(1).Caption = rst!keyword
            .lblKey(1).Visible = True
            If Not hasKeyword(player.ID, rst!keyword) And rst!WinOptKeyword = 0 And rst!KeywordOrSkill = 0 Then
               .cmd(1).Enabled = False
            End If
         End If
         '.lblDetail(0).Height = 1125
         '.lblDetail(1).Visible = True

         .cmd(1).Caption = rst!OptionName
         .cmd(1).ToolTipText = rst!OptionName
         .lblDetail(1).Caption = Replace(rst!Details, "^", vbNewLine)
         If rst!skill = 0 Then
            .SkillImg(1).Visible = False
         Else
            .SkillImg(1).Picture = LoadPictureGDIplus(App.Path & "\Pictures\skill" & rst!skill & ".bmp")
            .SkillImg(1).Visible = True
            .SkillImg(1).TransparentColor = &H0
            .SkillImg(1).TransparentColorMode = lvicUseTransparentColor
         End If

      Else 'no option 2??  never!
         .cmd(1).Visible = False
         .lblDetail(1).Visible = False
         '.lblDetail(0).Height = 2985
      End If
      rst.Close
      .Caption = "have fun Misbehavin' " & cnt & " of " & total
      .Picture = LoadPicture(App.Path & "\pictures\MisbehaveTemplate.bmp")
      .lblUnseen = "unseen: " & getUnseenMBDeck()
      .Alpha.Picture = LoadPictureGDIplus(App.Path & "\Pictures\suit" & suit & ".bmp")
      .Alpha.Visible = True
      .Alpha.TransparentColor = &HFFFFFF
      .Alpha.TransparentColorMode = lvicUseTransparentColor
      .AlphaInv.Picture = LoadPictureGDIplus(App.Path & "\Pictures\suit" & suit & ".bmp")
      .AlphaInv.Visible = True
      .AlphaInv.TransparentColor = &HFFFFFF
      .AlphaInv.TransparentColorMode = lvicUseTransparentColor
      .AlphaInv.Mirror = lvicMirrorVertical
      .setDalin dalin
      .Show 1
      
      getMisbehave = .MBCardID
      opt = .MBOption
      dalin = .dalin
      
   End With
      
End Function

'returns the Result Flag passed from doMisbehave: 1=proceed, 2=botched, 3=warrant, 4=load 1 contra per crew wit no gear, 5-double down
Public Function doMisbehaves(ByVal playerID, ByVal cnt As Integer, ByVal sectorID) As Integer
Dim x, CardID As Integer, opt, actualcnt As Integer, suit, c(1 To 4) As Integer, dalin As Boolean
   clearPicMB
   actualcnt = 0
   dalin = hasCrew(playerID, 102)
   For x = 1 To cnt
      CardID = 0
      While CardID = 0 'allow for reshuffle
         CardID = getMisbehave(opt, x, cnt, suit, dalin)
      Wend
      c(suit) = c(suit) + 1

      If opt = 3 Then
         actualcnt = actualcnt + 1
         'skip - ace in the hole
         DB.Execute "UPDATE MisbehaveDeck SET Seq =" & playerID & " WHERE CardID =" & CardID
         frmAction.lblMisbehaves = CStr(countMisbehaves(playerID))
      Else
         doMisbehaves = doMisbehave(playerID, CardID, opt)
         Select Case doMisbehaves
         Case 1, 4 ' proceed
            actualcnt = actualcnt + 1
            'stamp the card so it can be counted for goals
            DB.Execute "UPDATE MisbehaveDeck SET Seq =" & playerID & " WHERE CardID =" & CardID
            frmAction.lblMisbehaves = CStr(countMisbehaves(playerID))
         Case 2 'botched

            Exit Function
         
         Case 3 'warrant issued-done, job discarded
            Exit Function

         End Select
      End If
      'refresh as stuff may have changed
      If Not (frmShip Is Nothing) Then frmShip.RefreshShips
   Next x
   'do double down check
   If doMisbehaves = 1 Then 'normal success (not 4)
      For x = 1 To 4
         If c(x) > 1 Then
            doMisbehaves = 5
            Exit For
         End If
      Next x
   End If
   
   'inform of the success
   If IsNull(varDLookup("PlanetName", "Planet", "SectorID=" & sectorID)) Then
      If getCruiserSector() = sectorID Then
         PutMsg player.PlayName & " Misbehaved successfully " & actualcnt & " times at the Alliance Cruiser", playerID, Logic!Gamecntr, True, getLeader()
      ElseIf getCorvetteSector() = sectorID Then
         PutMsg player.PlayName & " Misbehaved successfully " & actualcnt & " times at the Operative's Corvette", playerID, Logic!Gamecntr, True, getLeader()
      End If
   Else
      PutMsg player.PlayName & " Misbehaved successfully " & actualcnt & " times at " & varDLookup("PlanetName", "Planet", "SectorID=" & sectorID), playerID, Logic!Gamecntr, True, getLeader()
   End If

End Function

' CardID is from MisbehaveDeck, opt is which option selected returns: 1=proceed, 2=botched, 3=warrant, 4=load 1 contra per crew wit no gear
Public Function doMisbehave(ByVal playerID, ByVal CardID, ByVal opt) As Integer
Dim SQL, skillcnt, skillwin, skillint, skilldiscards, x, bribe As Integer, riverskill As Integer
Dim Dice As Integer, WSkill As Integer, extraSkill As Integer, KeywordSkill As Integer, result '0=win,1-inter,2=fail
Dim rst As New ADODB.Recordset, frmDiscardGr As frmDiscardGear
Dim frmCrew As frmCrewSel, oneOnOne As Integer

   If opt = 0 Then
      MsgBox "option error", vbCritical, "Misbehave"
      Exit Function
   End If

   'grab the Nav Option chosen
   SQL = "SELECT MisOption.* "
   SQL = SQL & "FROM MisOption INNER JOIN MisbehaveDeck ON MisOption.OptionID = MisbehaveDeck.Option" & opt & "ID "
   SQL = SQL & "Where MisbehaveDeck.CardID = " & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
   
      PutMsg player.PlayName & "'s Misbehavin': " & rst!Details, playerID, Logic!Gamecntr
      Events.getNewEvents
      WSkill = rst!skill
      'let the tests begin ... :O  WIN, INTER OR FAIL ?
      If rst!win = 0 Then 'no test, just do Win outcomes
         result = 0
      ElseIf rst!ProfessionID > 0 And hasCrewAttribute(playerID, cstrProfession(rst!ProfessionID)) Then 'has Profession?->proceed
         result = 0
      ElseIf rst!KeywordOrSkill > 0 And hasKeyword(playerID, rst!keyword & "") Then 'has a Keyword?->proceed
         result = 0
         
      ElseIf WSkill > 0 Then 'we have a skill test
         '-----------------------------------------
         'Stitch & Sheydra can change a Fight to a Nego once per Job
         If WSkill = 3 And hasCrew(playerID, 27) And Not usedStitchSkill Then 'Stitch
            If MessBox("Stitch wants to Fight instead of Negotiation.  Do you want to use those skills instead?", "Negotiate -> Fight", "Yes", "No", 27) = 0 Then
               WSkill = 1
               usedStitchSkill = True
               PutMsg player.PlayName & " uses Stitch's one time Negotiation to Fight Skills", playerID, Logic!Gamecntr, True, 27
            End If
         End If
         If WSkill = 1 And hasCrew(playerID, 66) And Not usedStitchSkill Then  'Sheydra
            If MessBox("Sheydra wants to Negotiate instead of Fight.  Do you want to use those skills instead?", "Fight -> Negotiate", "Yes", "No", 66) = 0 Then
               WSkill = 3
               usedStitchSkill = True
               PutMsg player.PlayName & " uses Sheydra's one time Fight to Negotiation Skills", playerID, Logic!Gamecntr, True, 66
            End If
         End If
         
         If rst!OptionID = 52 Then  'One On One test
            Set frmCrew = New frmCrewSel
            frmCrew.crewFilter = " INNER JOIN (PlayerSupplies INNER JOIN  SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID WHERE PlayerSupplies.PlayerID=" & playerID & " AND PlayerSupplies.OffJob=0"
            frmCrew.Caption = "Pick a Crew to " & cstrSkill(WSkill) & " alone"
            frmCrew.Show 1
            oneOnOne = GetCombo(frmCrew.cboCrew)
         End If
      
         'if card accepts a bribe, ask for $100 a point
         If WSkill = 3 And (rst!bribe = 1 Or hasPerkAttributeValue(playerID, "Bribe", WSkill)) Then
            Do
               bribe = InputBoxx("They accept Bribes, $100 per skill point" & vbNewLine & vbNewLine & "Enter the number of POINTS you would bribe with..", "Money Talks", "0", getLeader())
               If bribe > 20 Then
                  MessBox "Seems a bit much don't ya think? Try that again..", "Too much!", "Ooops", "", getLeader()
               ElseIf bribe * 100 <= getMoney(playerID) Then 'can pay
                  getMoney playerID, (bribe * 100 * -1)
                  Exit Do
               Else
                  MessBox "Why you low-down thief, whatcha tryin' to pull?  Try again!", "Insufficient dough!", "Sorry", "", getLeader()
               End If
            Loop
         End If
         
         'Crazy River Tam (cardID 51/CrewID 32)
         If hasCrew(playerID, 32) And (oneOnOne = 0 Or oneOnOne = 32) Then
            Dice = RollDice(6)
            If hasCrew(playerID, 33) Then  'simon adds 2 to her rolls
               Dice = Dice + 2
            End If
            Select Case Dice
            Case 1, 2 'stay onboard
               DB.Execute "UPDATE PlayerSupplies SET OffJob = 1 WHERE CardID = 51"
               If Not (frmShip Is Nothing) Then frmShip.RefreshShips  'update display
               PutMsg player.PlayName & "'s River Tam cowers onboard and won't be misbehavin' any further on this job", playerID, Logic!Gamecntr, True, 32
            Case 3 'fight
               If WSkill = 1 Then
                  riverskill = 2
               End If
            Case 4 'Tech
               If WSkill = 2 Then
                  riverskill = 2
               End If
            Case 5 'negot
               If WSkill = 3 Then
                  riverskill = 2
               End If
            Case Else 'any skill
                  riverskill = 2
            End Select
            If riverskill = 2 Then
               'If hasCrew(playerID, 33) Then riverskill = 4
               PutMsg player.PlayName & "'s River Tam" & IIf(hasCrew(playerID, 33), ", encouraged by Simon,", "") & " channels the " & cstrSkill(WSkill) & " skill + 2", playerID, Logic!Gamecntr, True, 32
            ElseIf Dice > 2 Then
               PutMsg player.PlayName & "'s River Tam ain't misbehavin' this time", playerID, Logic!Gamecntr, True, 32
            End If
         End If
         
         extraSkill = hasGearCard(playerID, 24)
         If extraSkill > 0 Then 'we got one or more
            If MessBox("Do you wish to Eat the Fruity Bar and add 1 to the Test Roll?", "Extra Bite", "Yes", "No", 0, 24) = 0 Then
               doDiscardGear playerID, extraSkill
               extraSkill = 1
            Else
               extraSkill = 0
            End If
         End If
         
         x = hasGearCrew(playerID, 28) 'Mal's Brown Coat
         If x > 0 And varDLookup("Disgruntled", "Crew", "CrewID=" & x) > 0 And varDLookup("Fight", "Crew", "CrewID=" & x) > 0 And WSkill = 3 Then
            extraSkill = extraSkill + varDLookup("Fight", "Crew", "CrewID=" & x)
            PutMsg player.PlayName & "'s Disgruntled Crew wearing the Brown Coat adds their Fight skills to the Negotiation", playerID, Logic!Gamecntr, True, 0, 28
         End If
         
         If WSkill = 1 Then
            removeDigruntled playerID, WSkill ' Mal's Frontier Model B -Before each Fight Test, remove Disgruntled from the Owner.
         End If
         
         '<<<<<<<<<< ROLL THE DICE >>>>>>>>>>>>>>>>>
         Dice = RollDice(6, IIf(WSkill = 2 And hasCrew(playerID, 55), False, True)) 'Bester -On tech test, +6 "Thillin' Heroics" bonus dice does not apply
         
         
         If WSkill = 1 And hasGear(playerID, 47) Then ' Zoe's Mare's Leg Rifle -When making a Fight Test, roll two dice and use the highest.
            x = RollDice(6, True)
            If x > Dice Then
               PutMsg player.PlayName & " had rolled a " & CStr(Dice) & " so using Zoe's Mare's Leg Rifle rerolled a " & CStr(x), playerID, Logic!Gamecntr, True, 0, 47, 0, 0, 0, Dice
               Dice = x
            End If
         End If
         
         If Dice = 1 Then  'reroll ones?
            If hasGear(playerID, 6) Then ' has Jaynes Cunning Hat
               Do While Dice = 1
                  Dice = RollDice(6, IIf(WSkill = 2 And hasCrew(playerID, 55), False, True))
               Loop
               PutMsg player.PlayName & " uses Jaynes Cunning Hat to reRoll a 1 and got a " & CStr(Dice), playerID, Logic!Gamecntr, True, 0, 6, 0, 0, 0, Dice
               
            ElseIf hasGear(playerID, 35) And WSkill = 1 Then 'Inara's Bow
               x = hasGearCrew(playerID, 35)
               If x > 0 Then
                  If hasCrewAttribute(playerID, "Companion", 0, x) Then
                     Do While Dice = 1
                        Dice = RollDice(6, True)
                     Loop
                     PutMsg player.PlayName & " uses Inara's Bow to reRoll a 1 and got a " & CStr(Dice), playerID, Logic!Gamecntr, True, 0, 35, 0, 0, 0, Dice
                  End If
               End If
            End If
         End If
         
         'Zoe's skill can reroll a Fight test
         If WSkill = 1 And Dice < 6 Then
            x = getPerkAttributeCrew(playerID, "RerollFight")
            If x > 0 Then
               If MessBox("You rolled a " & Dice & vbNewLine & "Your Fight Skills allow you a re-roll, do you want to take that chance?", "Re-Roll option", "Re-roll", "Keep", x, 0, 0, Dice) = 0 Then
                  Dice = RollDice(6, True)
                  PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(Dice), playerID, Logic!Gamecntr, True, x, 0, 0, 0, 0, Dice
               End If
            End If
         End If
         
         'Kaylee can reroll a Tech test
         If WSkill = 2 And Dice < 6 Then
            x = getPerkAttributeCrew(playerID, "RerollTech")
            If x > 0 Then
               If MessBox("You rolled a " & Dice & vbNewLine & "Your Tech Skills allow you a re-roll, do you want to take that chance?", "Re-Roll option", "Re-roll", "Keep", x, 0, 0, Dice) = 0 Then
                  Dice = RollDice(6, IIf(WSkill = 2 And hasCrew(playerID, 55), False, True))
                  PutMsg player.PlayName & " uses extra Tech Skills to reRoll and got a " & CStr(Dice), playerID, Logic!Gamecntr, True, x, 0, 0, 0, 0, Dice
               End If
            End If
         End If

         'Inara can reroll a negotiate test
         If WSkill = 3 And Dice < 6 Then
            x = getPerkAttributeCrew(playerID, "RerollNegotiate")
            If x > 0 Then
               If MessBox("You rolled a " & Dice & vbNewLine & "Your Negotiation Skills allow you a re-roll, do you want to take that chance?", "Re-Roll option", "Re-roll", "Keep", x, 0, 0, Dice) = 0 Then
                  Dice = RollDice(6, True)
                  PutMsg player.PlayName & " uses extra Negotiation Skills to reRoll and got a " & CStr(Dice), playerID, Logic!Gamecntr, True, x, 0, 0, 0, 0, Dice
               End If
            End If
         End If
         
         If WSkill = 1 And hasGear(playerID, 45) And Dice < 6 Then 'yolanda's pistol - Discard to re-roll a Fight Test.
            If MessBox("You rolled a " & Dice & vbNewLine & "Yolanda's pistol allows you a re-roll, do you want to Discard the Pistol to take that extra chance?", "Re-Roll option", "Re-roll", "Keep", 0, 45, 0, Dice) = 0 Then
               doDiscardGear playerID, hasGearCard(playerID, 45)
               Dice = RollDice(6, True)
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(Dice), playerID, Logic!Gamecntr, True, 0, 45, 0, 0, 0, Dice
            End If
         End If
         
         If WSkill = 1 And hasGear(playerID, 48) And Dice < 6 Then 'Extra Ammo Clip - Discard to re-roll a Fight Test.
            If MessBox("You rolled a " & Dice & vbNewLine & "Extra Ammo Clips allow you a re-roll, do you want to Discard the Clips to take that extra chance?", "Re-Roll option", "Re-roll", "Keep", 0, 48, 0, Dice) = 0 Then
               doDiscardGear playerID, hasGearCard(playerID, 48)
               Dice = RollDice(6, True)
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(Dice), playerID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, Dice
            End If
         End If
         
         '----------------------------------------- see if we need to use the discardable skills & keywords...
         
         If rst!WinOptKeyword > 0 And hasKeyword(playerID, rst!keyword & "") Then 'keyword reduces win minimum - also could be discardable ???
            KeywordSkill = rst!WinOptKeyword
         End If
         
         skillwin = rst!win
         
         skillint = rst!Intermediate

         'get our skill totals, exclude gear from Kosherized rules
         skillcnt = getSkill(playerID, cstrSkill(WSkill), 0, True, (rst!kosher = 1), oneOnOne) + Dice + bribe + riverskill + extraSkill + KeywordSkill
         skilldiscards = getSkillDiscards(playerID, cstrSkill(WSkill), (rst!kosher = 1))

         
         If skillcnt < skillwin And skillcnt + skilldiscards >= skillwin Then  'we're in trouble 'we could use some help
            If MessBox("With the help of " & skillwin - skillcnt & " skill points, we can succeed" & vbNewLine & "Do you want to use a discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
               'show a list of gear to pick from up to or exceeding the value skillwin - skillcnt
               Set frmDiscardGr = New frmDiscardGear
               frmDiscardGr.kosher = (rst!kosher = 1)
               frmDiscardGr.nbrSelect = skillwin - skillcnt
               frmDiscardGr.skill = cstrSkill(WSkill)
               frmDiscardGr.Caption = "Select single use Gear to provide at least " & CStr(frmDiscardGr.nbrSelect) & " skill points"
               frmDiscardGr.Show 1
               'then add selected skill points to skillcnt, discard gear, and go on...
               skillcnt = skillcnt + frmDiscardGr.nbrSelected
            End If
                  
         ElseIf skillcnt < skillint And skillcnt + skilldiscards >= skillint Then  'we're in trouble 'we could use some help
            If MessBox("With the help of " & skillint - skillcnt & " skill points, we can make the intermediate outcome" & vbNewLine & "Do you want to use discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
               'show a list of gear to pick from up to or exceeding the value skillint - skillcnt
               Set frmDiscardGr = New frmDiscardGear
               frmDiscardGr.kosher = (rst!kosher = 1)
               frmDiscardGr.nbrSelect = skillint - skillcnt
               frmDiscardGr.skill = cstrSkill(WSkill)
               frmDiscardGr.Caption = "Select single use Gear to provide at least " & CStr(frmDiscardGr.nbrSelect) & " skill points"
               frmDiscardGr.Show 1
               'then add selected skill points to skillcnt, discard gear, and go on...
               skillcnt = skillcnt + frmDiscardGr.nbrSelected
            End If
         
         End If
         
         
         If hasGear(playerID, 32) And WSkill = 1 And skillcnt < rst!win Then
            If MessBox("The Fights not going so well with a skill score of " & skillcnt & vbNewLine & "Simon's Sonic Stun Baton might turn things around, wanna try another Thrillin' Heroics Roll and Discard the Baton?", "Stun Baton to the Fight", "Yes", "No", 0, 32) = 0 Then
               skillcnt = RollDice(6) + 6
               doDiscardGear playerID, hasGearCard(playerID, 32)
               PutMsg player.PlayName & " used Simon's Sonic Stun Baton to try and turn the Fight around ", playerID, Logic!Gamecntr
            End If
         End If
         '-----------------------------------------
         
         If skillcnt >= rst!WinOptKeyword And rst!WinOptKeyword > 0 And hasKeyword(playerID, rst!keyword & "") Then
            If skillcnt < rst!win Then 'needed the keyword to win
               'check if the keyword was single use, and discard
               discardGearKeyword playerID, rst!keyword
            End If
            result = 0
         ElseIf skillcnt >= rst!win Then
            result = 0
         ElseIf skillcnt >= rst!Intermediate And rst!Intermediate > 0 Then
            result = 1
         Else 'you lose
            result = 2
         End If
         PutMsg player.PlayName & "'s MB log: Rolls a " & Dice & " with added " & cstrSkill(WSkill) & " skill points of " & CStr(skillcnt - Dice) & " for a total of " & skillcnt & " to " & IIf(result = 0, "succeed :^)", IIf(result = 1, "partially succeed :^|", "lose :^(")), playerID, Logic!Gamecntr, True, IIf(oneOnOne > 0, oneOnOne, getLeader()), 0, 0, 0, 0, Dice, WSkill
         
      End If  'end of the initial Tests
         
      Select Case result
      Case 0 ' winners are grinners :D
         doMisbehave = rst!WinResult
         
         If rst!WinCash > 0 Then
            DB.Execute "UPDATE Players Set Pay = Pay + " & rst!WinCash & " WHERE PlayerID = " & playerID
         ElseIf rst!WinCash < 0 Then
            Dice = rst!WinCash
            If getMoney(playerID) <= Abs(Dice) Then
               Dice = getMoney(playerID) * -1
               PutMsg player.PlayName & "'s MB log: Funds depleted!", playerID, Logic!Gamecntr, True, getLeader()
            End If
            
            DB.Execute "UPDATE Players Set Pay = Pay + " & Dice & " WHERE PlayerID = " & playerID
            
         End If

         If rst!WinKillCrew <> 0 Then
            doKillCrews playerID, rst!WinKillCrew
         End If
         
      Case 1 'intermediate outcomes  :|
         doMisbehave = rst!InterResult
    
         If rst!InterKillCrew <> 0 Then
            doKillCrews playerID, rst!InterKillCrew
         End If
         
      Case 2 'loser outcomes :(
         doMisbehave = rst!FailResult

         If rst!FailKillCrew = 99 Then ':((
            doKillAllCrew playerID
         ElseIf rst!FailKillCrew = 1 And oneOnOne > 0 Then
            doKillCrew playerID, getCrewCardID(oneOnOne)
         ElseIf rst!FailKillCrew <> 0 Then
            doKillCrews playerID, rst!FailKillCrew
         End If
         
         If rst!Disgruntled = 4 Then 'discard all Mercs
            If doMercDiscard(playerID) Then
               PutMsg player.PlayName & " had the Mercs mutiny and leave", playerID, Logic!Gamecntr, True, getLeader()
            End If
         End If
            
      End Select
       
      'DO the tests that run whatever the above outcome -----------------------------------------
      
      'check if the keyword was single use, and discard
      If Nz(rst!keyword, "") <> "" Then
         If discardGearKeyword(playerID, Nz(rst!keyword), True) Then
            discardGearKeyword playerID, rst!keyword
            PutMsg player.PlayName & " used up the " & rst!keyword, playerID, Logic!Gamecntr
         End If
      End If
      
      If rst!cargo <> 0 Then ' could be -neg
         skillcnt = Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID))
         If skillcnt > rst!cargo Then skillcnt = rst!cargo
         If skillcnt <> 0 Then
            DB.Execute "UPDATE Players Set Cargo = Cargo + " & skillcnt & " WHERE PlayerID = " & playerID
            PutMsg player.PlayName & IIf(skillcnt > 0, " scored ", " lost ") & skillcnt & " Cargo", playerID, Logic!Gamecntr
         End If
      End If
      
      If rst!Contraband <> 0 Then ' could be -neg
         If doMisbehave = 4 Then ' one per Crew with no gear (doMisbehave=4)
            x = getCrewWithNoGear(playerID)
         Else
            x = rst!Contraband
         End If
         skillcnt = Int(CargoCapacity(playerID) - CargoSpaceUsed(playerID))
         If skillcnt > x Then skillcnt = x
         If skillcnt <> 0 Then
            DB.Execute "UPDATE Players Set Contraband = Contraband + " & skillcnt & " WHERE PlayerID = " & playerID
            PutMsg player.PlayName & IIf(skillcnt > 0, " scored ", " lost ") & skillcnt & " Contraband", playerID, Logic!Gamecntr
         End If
      End If
      
      '1-Moral only, 2-All Crew, 3-remove from Moral Crew,
      '4=Discard all Mercs, 5=Discard Mercs if fight is greater than crew
      
      If (rst!Disgruntled > 0 And rst!Disgruntled < 4 And (result = 2 Or rst!skill = 0)) Or (rst!Disgruntled = 6 And result = 0) Or (rst!Disgruntled = 3 And result = 0) Then 'apply disgruntled changes if lose or no test
         doDisgruntled playerID, rst!Disgruntled
         result = 0
      End If
      
      If rst!Disgruntled = 5 Then 'discard all Mercs if Merc Fight higher than others
         If getSkill(playerID, cstrSkill(1), 1) > getSkill(playerID, cstrSkill(1), 2) Then  'discard all Mercs
            If doMercDiscard(playerID) Then
               PutMsg player.PlayName & " had the Mercs outgun the crew and leave", playerID, Logic!Gamecntr, True, getLeader()
            End If
         End If
      End If
                  
   Else
      MsgBox "Error: Nav Card " & CardID & " Option " & opt & " not found!", vbCritical
   End If

Set rst = Nothing

End Function

'this is where we apply the 1000 rules and outcomes of the nav option :(
'these apply to FULLBURN only. to Full Stop, set fullburndone = True
Public Function doNav(ByVal CardID, ByVal opt) As Boolean
Dim SQL, sectorID, skillcnt, x, y, z
Dim result   '0=win,1-inter,2=fail
Dim rst As New ADODB.Recordset
Dim frmShUp As frmShipUpgd, frmBart As frmBarter
Dim frmSalvage As frmSalvaging, frmCrewList As frmCrewLst, frmSeize As frmSeized, frmStsh As frmStash

   'grab the Nav Option chosen
   SQL = "SELECT NavOption.* "
   SQL = SQL & "FROM NavOption INNER JOIN NavDeck ON NavOption.OptionID = NavDeck.Option" & opt & "ID "
   SQL = SQL & "Where NavDeck.CardID = " & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      PutMsg player.PlayName & "'s Nav log: " & rst!Details, player.ID, Logic!Gamecntr
      'let the tests begin ... :O  WIN, INTER OR FAIL ?
      x = 0
      If rst!Breakdown = 1 Then x = hasShipUpgradeAttribute(player.ID, "IgnoreBreakdowns")
      'has breakdown insurance ?
      If rst!Breakdown = 1 And x > 0 Then
         result = 0
         PutMsg player.PlayName & "'s Ship is Breakdown Proof!", player.ID, Logic!Gamecntr, True, 0, 0, x
         Exit Function
         
      ElseIf rst!WinProfession > 0 And Not hasCrewAttribute(player.ID, cstrProfession(rst!WinProfession)) Then
         result = 2 'go to lose option
         
      ElseIf rst!skill = 0 Then 'no test, just do Win outcomes
         result = 0
     
      ElseIf rst!skill > 0 Then 'we have a skill test
         result = doSkillTest(rst!skill, rst!win, rst!Intermediate, rst!bribe)
      
      End If
         
      Select Case result
      Case 0 ' winners are grinners :D
         If rst!WinKeepFlying = 0 Then  'full stop
            frmAction.fullburndone = True
            frmAction.moseydone = True
         End If
         If rst!WinCash <> 0 Then
            DB.Execute "UPDATE Players Set Pay = Pay + " & rst!WinCash & " WHERE PlayerID = " & player.ID
         End If
         If rst!WinCargo <> 0 Then ' could be -neg
            skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            If skillcnt > rst!WinCargo Then skillcnt = rst!WinCargo
            If skillcnt <> 0 Then
               DB.Execute "UPDATE Players Set Cargo = Cargo + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!WinPassenger <> 0 Then ' could be -neg
            skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            If skillcnt < rst!WinPassenger And rst!WinPassenger = 2 Then 'cannot fit eryone
               PutMsg player.PlayName & " couldn't fit all Passengers, Moral Crew are not going to be happy!", player.ID, Logic!Gamecntr, True, getLeader()
               doDisgruntled player.ID, 1
            ElseIf skillcnt > rst!WinPassenger Then
               skillcnt = rst!WinPassenger
            End If
            If skillcnt <> 0 Then
               DB.Execute "UPDATE Players Set Passenger = Passenger + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!WinFugitive <> 0 Then ' could be -neg
            skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            If skillcnt < rst!WinFugitive And rst!WinFugitive = 4 Then 'cannot fit eryone
               PutMsg player.PlayName & " couldn't fit all Fugitives, Moral Crew are not going to be happy!", player.ID, Logic!Gamecntr, True, getLeader()
               doDisgruntled player.ID, 1
            ElseIf skillcnt > rst!WinFugitive Then
               skillcnt = rst!WinFugitive
            End If
            If skillcnt <> 0 Then
               DB.Execute "UPDATE Players Set Fugitive = Fugitive + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!WinFuel < 0 Then ' -neg
            skillcnt = rst!WinFuel
            If varDLookup("Fuel", "Players", "PlayerID=" & player.ID) >= Abs(skillcnt) Then  'check we're not going -ve
               DB.Execute "UPDATE Players Set Fuel = Fuel + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         ElseIf rst!WinFuel = 14 Then ' all you can load
            skillcnt = (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID)) * 2
            Do
               x = InputBoxx("Select up to " & skillcnt & " Fuel to salvage", "Salvage Fuel", getLeader())
               If x <= skillcnt And x > -1 Then
                  skillcnt = x
                  Exit Do
               Else
                  MessBox "Invalid Fuel quantity", "Fuel Requirements", "Ooops", "", getLeader()
               End If
            Loop
            DB.Execute "UPDATE Players Set Fuel = Fuel + " & skillcnt & " WHERE PlayerID = " & player.ID
         ElseIf rst!WinFuel > 0 Then ' small +ve
            skillcnt = (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID)) * 2
            If skillcnt > rst!WinFuel Then skillcnt = rst!WinFuel
            DB.Execute "UPDATE Players Set Fuel = Fuel + " & skillcnt & " WHERE PlayerID = " & player.ID
         End If
         
         If rst!WinParts = -99 Then 'sell up to 3 parts for $500ea
            Do
               y = varDLookup("Parts", "Players", "PlayerID=" & player.ID)
               If y = 0 Then Exit Do
               x = InputBoxx("How many Parts (you have " & y & ") would you like to sell for $500ea?", "Sell Parts", "0", getLeader())
               If x > y Then
                  MessBox "Invalid Parts quantity", "Parts Requirements", "Ooops", "", getLeader()
               Else
                  If x > 0 Then
                     DB.Execute "UPDATE Players SET Parts = Parts - " & x & ", Pay = Pay + " & CStr(x * 500) & " WHERE PlayerID=" & player.ID
                  End If
                  Exit Do
               End If
            Loop
            
         ElseIf rst!WinParts <> 0 Then ' could be -neg  . skillcnt re-used to count parts here
            skillcnt = (CargoCapacity(player.ID) - CargoSpaceUsed(player.ID)) * 2
            If skillcnt > rst!WinParts Or rst!WinParts < 0 Then skillcnt = rst!WinParts
            If skillcnt * -1 > varDLookup("Parts", "Players", "PlayerID=" & player.ID) Then 'stop going neg
               skillcnt = Val(varDLookup("Parts", "Players", "PlayerID=" & player.ID)) * -1
               If rst!Breakdown = 1 Then 'no parts, fullstop as breakdown proof is tested at start /|\
                  frmAction.fullburndone = True
                  frmAction.moseydone = True
               End If
            End If
            If skillcnt <> 0 Then DB.Execute "UPDATE Players Set Parts = Parts + " & skillcnt & " WHERE PlayerID = " & player.ID
         End If
         If rst!WinContraband <> 0 Then ' could be -neg
            skillcnt = Int(CargoCapacity(player.ID) - CargoSpaceUsed(player.ID))
            If skillcnt > rst!WinContraband Then skillcnt = rst!WinContraband
            If skillcnt <> 0 Then
               DB.Execute "UPDATE Players Set Contraband = Contraband + " & skillcnt & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!WinShipUpgrade <> 0 Then
            
            'present list of discarded upgrades to choose one for free
            Set frmShUp = New frmShipUpgd
            If getShipUpgrades(player.ID) < 3 Then
               frmShUp.discardMode = rst!WinShipUpgrade
            Else 'DriveCores only, no spare slots
               frmShUp.discardMode = 5
            End If
            frmShUp.Show 1
            
         End If
         
         If rst!WinGoods > 0 Then
            Set frmSalvage = New frmSalvaging
            frmSalvage.mode = 1
            frmSalvage.salvageCount = rst!WinGoods
            frmSalvage.Show 1
         ElseIf rst!WinGoods = -99 Then
            Set frmStsh = New frmStash
            frmStsh.Show 1
            PutMsg player.PlayName & " lost any Goods in hold overboard", player.ID, Logic!Gamecntr
         End If
         
         If rst!WinKillCrew <> 0 Then
            x = doKillCrews(player.ID, rst!WinKillCrew)
            If rst!OptionName = "If we're very lucky" And hasShipUpgrade(player.ID, 18) > 0 And x > 0 Then
               doDiscardGear player.ID, hasShipUpgrade(player.ID, 18)
               PutMsg player.PlayName & " lost the Reaver-Flage upgrade in the Reaver skuffle", player.ID, Logic!Gamecntr
            End If

         End If
         
         'UNIQUE WIN OPTIONS-----------------------------------
         If rst!WinFunction > 0 Then  'here lies all the new weird functions
            Select Case rst!WinFunction
               Case 1 ' Add 1 to the Range of this Fly Action for each Moral Crew on board
                  turnExtraRange = countCrewAttribute(player.ID, "Moral")
                  frmAction.lblFBRange.Caption = CStr(Val(frmAction.lblFBRange.Caption) + turnExtraRange)
               
               Case 2 'Gambling
                  If getMoney(player.ID) < 1000 Then
                     MessBox "You don't have the Cash to make the bet", "Cashflow Problem", "Ooops", "", getLeader()
                  Else
                     x = RollDice(6)
                     If x > 4 Then
                        getMoney player.ID, 2000
                        PutMsg player.PlayName & " rolls a " & x & " and wins $2000", player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, x
                     Else
                        getMoney player.ID, -1000
                        PutMsg player.PlayName & " rolls a " & x & " and loses $1000", player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, x
                     End If
                  End If
               
               Case 3 'Passengers and Fugitives dispute
                  x = RollDice(6, True) 'use Thrillin heroics roll
                  skillcnt = x
                  y = varDLookup("Fugitive", "Players", "PlayerID=" & player.ID)
                  z = varDLookup("Passenger", "Players", "PlayerID=" & player.ID)
                  If x < (y + z) Then  'out they go -auto mode. preference to Passengers go first
                     If x >= y Then
                        x = x - y
                        y = 0
                        If x >= z Then
                           z = 0
                        Else
                           z = z - x
                        End If
                     Else
                        y = y - x
                     End If
                     DB.Execute "UPDATE Players SET Passenger =" & z & ", Fugitive =" & y & " WHERE PlayerID=" & player.ID
                     PutMsg player.PlayName & " rolls a " & skillcnt & " and is left with " & z & " Passengers and " & y & " Fugitives", player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, skillcnt
                  Else
                     PutMsg player.PlayName & " rolls a " & skillcnt & " and retains any Passengers and Fugitives", player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, skillcnt
                  End If
                  
               Case 4   'She's Hemmorrhaging Fuel!
                  HemmorrhagingFuel = True  'set global - picked up by frmAction
                  
               Case 5   'Slingshot Roundhouse
                  If getCrewAttribute(player.ID, cstrProfession(2)) > 0 And getPlanetID(player.ID) > 0 Then
                     turnExtraRange = 3
                     frmAction.lblFBRange.Caption = CStr(Val(frmAction.lblFBRange.Caption) + turnExtraRange)
                  End If
                  
               Case 6
                  'Fancy Meetin' You Here, take 1 Crew Card from any discard pile for free
                  If getCrewCount(player.ID) < CrewCapacity(player.ID) Then
                     Set frmCrewList = New frmCrewLst
                     frmCrewList.selectCrew = -1
                     frmCrewList.Caption = "Select 1 Crew from Discards"
                     frmCrewList.Show 1
                  End If
                                    
               Case 7  'Shanghai Surprise! Take 1 Crew from Regina's Discard Pile.
                  If getCrewCount(player.ID) < CrewCapacity(player.ID) Then
                     Set frmCrewList = New frmCrewLst
                     frmCrewList.selectCrew = -1
                     frmCrewList.SupplyID = 5 'Regina
                     frmCrewList.Caption = "Select 1 Regina Crew from Discards"
                     frmCrewList.Show 1
                  End If
                  
            End Select
         End If
         
         If rst!SalvageOp <> 0 Then 'last win function, ignored if lose
            'load any Crew modifiers to add salvage due to Perk (SOCargo, SOContra...)
            'if can fit them of course
            doSalvage player.ID
         End If
         
      Case 1 'intermediate outcomes  :|
         If rst!InterKeepFlying = 0 Then  'full stop
            frmAction.fullburndone = True
            frmAction.moseydone = True
         End If
         If rst!InterGoods > 0 Then
            Set frmSalvage = New frmSalvaging
            frmSalvage.mode = 1
            frmSalvage.salvageCount = rst!InterGoods
            frmSalvage.Show 1
         End If
         If rst!InterCargo <> 0 Then ' could be -neg
            DB.Execute "UPDATE Players Set Cargo = Cargo + " & rst!InterCargo & " WHERE PlayerID = " & player.ID
         End If
         If rst!InterKillCrew <> 0 Then
            x = doKillCrews(player.ID, rst!InterKillCrew)
            If rst!OptionName = "If we're very lucky" And hasShipUpgrade(player.ID, 18) > 0 And x > 0 Then
               doDiscardGear player.ID, hasShipUpgrade(player.ID, 18)
               PutMsg player.PlayName & " lost the Reaver-Flage upgrade in the Reaver skuffle", player.ID, Logic!Gamecntr
            End If

         End If
         
      Case 2 'loser outcomes :(
         If rst!FailKeepFlying = 0 Then  'full stop
            frmAction.fullburndone = True
            frmAction.moseydone = True
         End If
         If rst!FailCargo <> 0 Then ' could be -neg
            If (rst!FailCargo * -1) <= varDLookup("Cargo", "Players", "PlayerID = " & player.ID) Then
               DB.Execute "UPDATE Players Set Cargo = Cargo + " & rst!FailCargo & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!FailFuel <> 0 Then ' could be -neg
            If (rst!FailFuel * -1) <= varDLookup("Fuel", "Players", "PlayerID = " & player.ID) Then
               DB.Execute "UPDATE Players Set Fuel = Fuel + " & rst!FailFuel & " WHERE PlayerID = " & player.ID
            End If
         End If
         If rst!FailParts <> 0 Then ' could be -neg
            If (rst!FailParts * -1) <= varDLookup("Parts", "Players", "PlayerID = " & player.ID) Then
               DB.Execute "UPDATE Players Set Parts = Parts + " & rst!FailParts & " WHERE PlayerID = " & player.ID
            End If
         End If
         
         If rst!FailGoods = -99 Then 'goods seized not in Stash
            'allow for stash modifiers.  reduce by 4+mods
            If SeizeAllContraCargo(player.ID) Then   'this is a compromise - todo - rework for ALL Goods
               PutMsg player.PlayName & "'s Nav log: lost some Contraband/Cargo not in Stash", player.ID, Logic!Gamecntr, True, getLeader()
            End If
            
         ElseIf rst!FailGoods <> 0 Then ' could be -neg
            Set frmSalvage = New frmSalvaging
            frmSalvage.mode = IIf(rst!FailGoods > 0, 1, 4) 'add/discard
            frmSalvage.salvageCount = Abs(rst!FailGoods)
            frmSalvage.Show 1

         End If
         
         If rst!FailShipUpgrade <> 0 Then
            If getShipUpgrades(player.ID) > 0 Then
               'present list of player upgrades to discard one
               Set frmShUp = New frmShipUpgd
               frmShUp.discardMode = 1
               frmShUp.Show 1
            End If
         End If
         If rst!FailKillCrew <> 0 Then
            x = doKillCrews(player.ID, rst!FailKillCrew)
            If rst!OptionName = "If we're very lucky" And hasShipUpgrade(player.ID, 18) > 0 And x > 0 Then
               doDiscardGear player.ID, hasShipUpgrade(player.ID, 18)
               PutMsg player.PlayName & " lost the Reaver-Flage upgrade in the Reaver skuffle", player.ID, Logic!Gamecntr
            End If
         End If
         
         If rst!FailNestedTest > 0 Then 'go all Inception on its ass
            doNav CardID, rst!FailNestedTest
         End If
            
      End Select
       
      'DO the tests that run whatever the above outcome -----------------------------------------
      sectorID = varDLookup("SectorID", "Players", "PlayerID=" & player.ID)
       
      If rst!Disgruntled <> 0 Then 'apply disgruntled changes
         doDisgruntled player.ID, rst!Disgruntled
      End If
      
      Select Case rst!MoveReaver
         Case 1   ' 1 - move 1
            If Logic!AutoAI = 0 Then
               setPlayer player.ID, "X", 1
               If isSoloGame(True) Then
                  actionSeq = ASNavReav
               Else
                  actionSeq = ASNavReavEnd
               End If
            
            Else
               moveAutoAI 6 + RollDice(NumOfReavers)
      
            End If
            
         Case 2    '2-you move reaver to any B zone,
            MessBox "Move a Reaver to any Rim or Border sector", "Reavers on the Move", "OK", "", getLeader()
            actionSeq = ASNavReavBorder
            
         Case 3    '3-move to your location  (evade done later)
            If getCutterSector(sectorID) = 0 Then
               MoveShip 6 + RollDice(NumOfReavers), sectorID
            End If
            
         Case 4  'other player move reaver to any B zone,
            If Logic!AutoAI = 0 Then
               setPlayer player.ID, "W", 1
               If isSoloGame(True) Then
                  MessBox "Move a Reaver to any Rim or Border sector", "Reavers on the Move", "OK", "", getLeader()
                  actionSeq = ASNavReavBorder
               Else
                  actionSeq = ASNavReavEnd
               End If
            
            Else
               doMoveCutterPlanetary 6 + RollDice(NumOfReavers)
      
            End If
      End Select
      
      Select Case rst!MoveAlliance
         Case 1   ' 1 - move 1
             If Logic!AutoAI = 0 Then
               setPlayer player.ID, "Y", 1
               If isSoloGame(True) Then
                  MessBox "Move the Alliance Cruiser one sector", "Cruiser on the Move", "OK", "", getLeader()
                  actionSeq = ASNavCrus
               Else
                  actionSeq = ASNavCrusEnd
               End If
             
            Else
               moveAutoAI 5
      
            End If
            
         Case 2   '2- move to any
            MessBox "Move the Alliance Crusier to any Alliance sector not occupied by a Firefly", "Wild Gosling Chase", "OK", "", getLeader()
            actionSeq = ASNavCrusBorder
            
         Case 3   '3-move to outlaw ship
            If outlawExists(player.ID) Then
               MessBox "Move the Crusier to a sector with a rival Outlaw Ship", "A Legitimate Tip", "OK", "", getLeader()
               actionSeq = ASNavCrusOutlaw
            End If
         Case 4 'alliance pays you a visit
            'for each Wanted Crew: 1-Remove Crew, 2+ Crew safe
            'may use Cry Baby - or other modifiers? eg. Concealed Smuggling Compartments
            x = isOutlaw(player.ID)
            If doMoveAlliance(player.ID, sectorID) Then
               CruiserCutter = sectorID 'set it as faced
               If Not (FullburnMovesDone = 0 And MoseyMovesDone = 0) And x Then 'only stop if Flying
                  frmAction.moseydone = True 'Full Stop!
                  frmAction.fullburndone = True
               End If
            End If
         
         Case 5 'move adjacent if failed
            If result = 2 Then
               If Logic!AutoAI = 0 And doMoveAllianceAdjacent(sectorID, True) Then  'there is a valid solution
                  setPlayer player.ID, "Z", 1
                  If isSoloGame(True) Then
                    MessBox "Move the Alliance Cruiser adjacent your Ship", "Cruiser on the Move", "OK", "", getLeader()
                    actionSeq = ASNavCrusAdjacent
                  Else
                    actionSeq = ASNavCrusEnd
                  End If
               
               Else
                  doMoveAllianceAdjacent sectorID
                 
               End If
            End If
            
         Case 6 'alert tokens adjacent your posn
            doAddTokensAdjacent sectorID
            RefreshBoard
         Case 7 'corvette contact
            Set frmSeize = New frmSeized
            If frmSeize.RefreshList > 0 Then 'some are not stashed
               frmSeize.Show 1
            End If
            If SeizeAllFugi(player.ID) Then
               PutMsg player.PlayName & " lost some Fugitives not in Stash", player.ID, Logic!Gamecntr, True, getLeader()
            End If
         
         Case 8 'discard 1 crew
            Set frmSeize = New frmSeized
            frmSeize.Caption = "Select the Crew Member detained by the Alliance"
            If frmSeize.RefreshDiscardList() > 0 Then 'crew exist
               frmSeize.Show 1
            End If
         
         Case 9 'alert tokens at every Outlaw Ship
            doAddTokensOutlaws
            If isOutlaw(player.ID) Then ignoreToken = sectorID 'so as to not trip on one put here
            RefreshBoard
         Case 10 ' Move Corvette Adjacent player
            If Logic!AutoAI = 0 And doMoveCorvetteAdjacent(sectorID, True) Then
               setPlayer player.ID, "V", 1
               If isSoloGame(True) Then
                 MessBox "Move the Operative's Corvette adjacent your Ship", "Corvette on the Move", "OK", "", getLeader()
                 actionSeq = ASNavCorvAdjacent
               Else
                 actionSeq = ASNavCrusEnd
               End If
            
            Else
               doMoveCorvetteAdjacent sectorID
              
            End If
            
         Case 11  'Corvette to an unoccupied Alliance, Border, or Rim Planetary Sector.
            If Logic!AutoAI = 0 Then
               setPlayer player.ID, "U", 1
               If isSoloGame(True) Then
                 MessBox "Move the Operative's Corvette to a Planetary Sector", "Corvette on the Move", "OK", "", getLeader()
                 actionSeq = ASNavCorvPlanetary
               Else
                 actionSeq = ASNavCrusEnd
               End If
            
            Else
               doMoveCorvettePlanetary
            End If
            
         Case 12  'move Operative's Corvette 1 or 2 Sectors within Alliance, Border or Rim Space
            y = getCorvetteSector
            moveAutoCorvette2 0, False, y
                  
      End Select
      
      If rst!MovePlayer > 0 Then
         For x = 1 To rst!MovePlayer
            moveAutoAI player.ID, 1, True
            drawLine 0, -2, getPlayerSector(player.ID)
            drawLine 1, -2, getPlayerSector(player.ID)
         Next x
      End If
      
      
      If rst!trader <> 0 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < 0.5 Then
            MessBox "You have no spare cargo capacity", "Trading on the go", "Ooops", "", getLeader()
         Else
            ' enable 1&2 Trader modes
            Set frmBart = New frmBarter
            frmBart.trader = rst!trader
            frmBart.Show 1
         End If
      End If
      
      If rst!KillAllPassFugi <> 0 Then
          DB.Execute "UPDATE Players SET Fugitive = 0, Passenger = 0 WHERE PlayerID = " & player.ID
          'remove active bounty jobs too
          DB.Execute "DELETE FROM PlayerJobs WHERE PlayerID = " & player.ID & " AND JobStatus = 0 AND CardID in (Select CardID From ContactDeck WHERE ContactID = 10)"
          PutMsg player.PlayName & " has lost any fugutives and passengers they may have had.", player.ID, Logic!Gamecntr
      End If
      
      If rst!SeizeGoods = 1 Then 'Contraband and Fugitives not in your Stash are seized. Full Stop
         'allow for stash modifiers.  reduce by 4+mods
         If SeizeAllContraFugi(player.ID) Then
            PutMsg player.PlayName & "'s Nav log: lost some Contraband/Fugitives not in Stash", player.ID, Logic!Gamecntr, True, getLeader()
         End If
      End If
      
      If rst!SeizeGoods = 2 Then 'Contraband and Cargo not in your Stash are seized. Full Stop
         'allow for stash modifiers.  reduce by 4+mods
         If SeizeAllContraCargo(player.ID) Then
            PutMsg player.PlayName & "'s Nav log: lost some Contraband/Cargo not in Stash", player.ID, Logic!Gamecntr, True, getLeader()
         End If
      End If
      
      If rst!Warrant = 1 Or (result = 2 And rst!Warrant = 2) Or (isOutlaw(player.ID) And rst!Warrant = 3) Then
         If Not warrantDodge(player.ID) Then
            PutMsg player.PlayName & "'s Nav log: a Warrant has been issued" & IIf(isSolid(player.ID, 5), " and you are no longer Solid with Harken", "") & "!", player.ID, Logic!Gamecntr, True, getLeader()
            'add a Warrant and clear any Solid with Harken (5)
            DB.Execute "UPDATE Players SET Warrants = Warrants + 1" & IIf(discardRoberta(player.ID), "", ", Solid5 = 0") & " WHERE PlayerID = " & player.ID
         End If
      End If
      If rst!Warrant = -1 Then
         'clear Warrants
         DB.Execute "UPDATE Players SET Warrants = 0 WHERE PlayerID = " & player.ID
         PutMsg player.PlayName & "'s Nav log: any Warrants have been cleared!", player.ID, Logic!Gamecntr, True, getLeader()
      End If
      
      If rst!Token = 1 Then
         changeToken sectorID, 1
         ignoreToken = sectorID
      ElseIf rst!Token = 2 Then
         changeAToken sectorID, 1
         ignoreToken = sectorID
      End If
      
      If rst!Evade = 1 Or (result = 0 And rst!Evade = 2) Then
         PutMsg player.PlayName & "'s Nav log: EVADE!", player.ID, Logic!Gamecntr, True, getLeader()
         actionSeq = ASNavEvade
      End If
                  
   Else
      MsgBox "Error: Nav Card " & CardID & " Option " & opt & " not found!", vbCritical
   End If

Set rst = Nothing

End Function

'save selected (Seq=6 + selected) to players Jobs, unselected back to 5 DISCARDED
Public Function doDeal(ByVal playerID As Integer) As Integer
Dim Index As Integer
   With frmDeal.sftTree
      
      For Index = 0 To .ListCount - 1
         Select Case .ItemDataString(Index)
         Case "R"  'selected
            doDeal = doDeal + 1
            assignDeal playerID, .ItemData(Index)

         Case "UN" 'place back in discard (5)
            DB.Execute "UPDATE ContactDeck SET Seq =" & CStr(DISCARDED) & " WHERE CardID = " & .ItemData(Index)
            .ItemDataString(Index) = "O"
            Set .ItemPicture(Index) = frmDeal.AssetImages.Overlay("L", "O")
         End Select
      Next Index
   
   End With
End Function

'save selected (Seq=6 + selected) to players Jobs, unselected back to 5 DISCARDED
Public Function doBuy(ByVal playerID As Integer) As Integer
Dim Index As Integer, cost As Integer, imposter As Integer
   With frmBuy.sftTree
      cost = 0
      For Index = 0 To .ListCount - 1
         imposter = 0
         Select Case .ItemDataString(Index)
         Case "R"  'selected -pay up!!
            If .ItemData(Index) = 28 Then 'If Saffron is hired by anyone, Remove the existing imposter from Play
               If haveCrewAnyone(54) Then
                  doDiscardCrew 100
                  imposter = 54
               ElseIf haveCrewAnyone(41) Then
                  doDiscardCrew 70
                  imposter = 41
               End If
            End If
            
            If .ItemData(Index) = 100 Then 'If Bridgit is hired by anyone, Remove the existing imposter from Play
               If haveCrewAnyone(23) Then
                  doDiscardCrew 28
                  imposter = 23
               ElseIf haveCrewAnyone(41) Then
                  doDiscardCrew 70
                  imposter = 41
               End If
            End If
            
            If .ItemData(Index) = 70 Then 'If Yolonda is hired by anyone, Remove the existing imposter from Play
               If haveCrewAnyone(54) Then
                  doDiscardCrew 100
                  imposter = 54
               ElseIf haveCrewAnyone(23) Then
                  doDiscardCrew 28
                  imposter = 23
               End If
            End If
            
            If .CellItemData(Index, 2) = 1 Then
               'if buying a Drive Core, swap out the existing one
               removeDriveCore player.ID
            End If
            doBuy = doBuy + 1

            cost = cost + .CellItemData(Index, 8)

            DB.Execute "UPDATE SupplyDeck SET Seq =" & playerID & " WHERE CardID = " & .ItemData(Index)
            'add the card to the players deck
            DB.Execute "INSERT INTO PlayerSupplies (PlayerID, CardID) VALUES (" & playerID & ", " & .ItemData(Index) & ")"
         
            If imposter > 0 Then
               PutMsg getCrewName(0, imposter) & " has turned up as " & getCrewName(.ItemData(Index)) & " on " & player.PlayName & "'s Ship", playerID, Logic!Gamecntr, True, getCrewID(.ItemData(Index)), 0, 0, 0, 1
            End If
         
         Case "UN" 'place back in discard (5)
            DB.Execute "UPDATE SupplyDeck SET Seq =" & CStr(DISCARDED) & " WHERE CardID = " & .ItemData(Index)
            .ItemDataString(Index) = "O"
            Set .ItemPicture(Index) = frmBuy.AssetImages.Overlay("L", "O")
         End Select
      Next Index
      If cost > 0 Then
         DB.Execute "UPDATE Players SET Pay=Pay - " & cost & " WHERE PlayerID = " & playerID
      End If
   
   End With
End Function

Public Function getBuyCost() As Integer
   getBuyCost = frmBuy.getCost("R")
End Function

Private Function getNewPlayer() As Integer
Dim frmplayer As New frmSelPlayer
   If frmplayer.RefreshList > 1 Then 'no need to pick
      frmplayer.Show 1
   End If
   getNewPlayer = frmplayer.playerID

End Function

Private Function checkFlacGun(ByVal sectorID, Optional ByVal ignore As Boolean = False) As Boolean
Dim x, g

   If ignore Then Exit Function 'use when other conditions already fail so as to not trigger the test

   x = getCutterSector(sectorID)
   If x > 0 Then  'we got company!
      g = hasShipUpgrade(player.ID, 15)
      If g > 0 Then 'Flac Gun
         If MessBox("Reaver within firing Range, do you want to use the single-use Flac Gun to fend it off?", "Reaver Cutter", "Yes", "No", 0, 0, 15) = 0 Then
            checkFlacGun = True
            doDiscardGear player.ID, g
            moveAutoAI x
            PutMsg player.PlayName & " depleted their Hull-Mounted Flak Gun to fend off a Reaver", player.ID, Logic!Gamecntr
         End If
      End If
   End If
   

End Function

Private Sub checkBigBlack(ByVal CardID)

   If varDLookup("CardName", "NavDeck", "CardID=" & CardID) = "The Big Black" Then
      TheBigBlack = TheBigBlack + 1
      If TheBigBlack = 2 Then
         If CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) > 0 Then 'got room
            If MessBox("Your Emissions Recycler is working well in the Big Black." & vbNewLine & "Do you want to recover one Fuel?", "Emissions Recycler", "Yes", "No", 0, 0, 20) = 0 Then
               DB.Execute "Update Players Set Fuel = Fuel + 1 WHERE PlayerID =" & player.ID
            End If
         End If
         TheBigBlack = -1 'disable
      End If
   Else
      TheBigBlack = 0 'reset the count
   End If
   
End Sub

Public Function doKillCrews(ByVal playerID, ByVal NoToKill As Integer, Optional ByVal onJobOnly As Boolean = False) As Integer
Dim tracey As Integer, crewCount As Integer, killed As Integer
Dim frmKillCrw As frmKillCrew
      tracey = 0
      'Tracey must be Killed first "KillFirst" perk CardID 4, CrewID 11
      If hasCrew(playerID, 11) Then
         killed = doKillCrew(playerID, 4)
         tracey = 1
      End If
      
      If NoToKill - tracey > 0 Then
         Set frmKillCrw = New frmKillCrew
         crewCount = getCrewCount(playerID, onJobOnly)
         If crewCount >= (NoToKill - tracey) Then 'more or equal crew than to be killed
            crewCount = (NoToKill - tracey)
            
         End If
         If crewCount > 0 Then
            frmKillCrw.nbrSelect = crewCount
            frmKillCrw.Show 1
            killed = frmKillCrw.killed
         End If

      End If
      doKillCrews = killed
      If killed > 0 Then
         'If Not (frmDeal Is Nothing) Then frmDeal.RefreshDeals
         PutMsg player.PlayName & " sadly lost " & CStr(killed) & " Crew", playerID, Logic!Gamecntr
      End If
End Function

Private Sub doSlaveTrade(ByVal TraderID)
Dim frmTrade As New frmTrader
   frmTrade.TraderID = Logic!player
   frmTrade.lblTitle(1).Caption = PlayCode(TraderID).PlayName & "'s Trade Items"
   frmTrade.Show 1
   If Not (frmShip Is Nothing) Then frmShip.RefreshShips
End Sub


Private Sub Verse_SectClick(ByVal Index As Integer)
Dim Havens As Boolean

   'picking starting sector
   If pickStartSector = 1 Then
      Havens = useHavens(Logic!StoryID)
      If Not CheckClash(player.ID, Index, Havens) Then
         If Havens Then placeHaven player.ID, Index
         MoveShip player.ID, Index
         pickStartSector = 2  'flag the selection is done
      End If
   End If
   
   If actionSeq = ASmosey Then
      If validMove(player.ID, Index, True) And frmAction.imgMosey.Tag = "1" Then
         frmAction.actionButtonEnable "imgCancel", False
         frmAction.actionButtonEnable "imgMosey", False
         'get players current posn and check route
         MoveShip player.ID, Index, 7
         MoseyMovesDone = MoseyMovesDone + 1
         drawLine 0, -2, Index
         drawLine 1, -2, Index
         wormHoleOpen = False
         drawLine 2, -1
         actionSeq = ASMoseyEnd 'throw to main loop
         CruiserCutter = 0
         CorvetteSeq = 0
         ignoreToken = 0
       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASfullburn Then
      If validMove(player.ID, Index, hasShipUpgrade(player.ID, 18)) And frmAction.imgFullBurn.Tag = "1" Then
         frmAction.actionButtonEnable "imgCancel", False
         frmAction.actionButtonEnable "imgFullBurn", False
         MoveShip player.ID, Index
         FullburnMovesDone = FullburnMovesDone + 1
         If FullburnMovesDone = 1 And Val(frmAction.lblFBFuel.Caption) > 0 Then burnFuel player.ID, Val(frmAction.lblFBFuel.Caption)
         If HemmorrhagingFuel Then burnFuel player.ID, 1
         If Not frmShip Is Nothing Then
            frmShip.refreshFuel player.ID
         End If
         drawLine 0, -2, Index
         drawLine 1, -2, Index
         wormHoleOpen = False
         drawLine 2, -1
         actionSeq = ASFullburnEnd 'throw to main loop
         CruiserCutter = 0
         CorvetteSeq = 0
         ignoreToken = 0

       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavEvade Then
      If validMove(player.ID, Index, hasShipUpgrade(player.ID, 18)) Then
         'if evading a reaver at the beginning of turn, then don't stop fullburn
         If FullburnMovesDone > 0 Then frmAction.fullburndone = True
         If MoseyMovesDone > 0 Then frmAction.moseydone = True
         MoveShip player.ID, Index
         drawLine 0, -2, Index
         drawLine 1, -2, Index
         actionSeq = ASNavEvadeEnd
         CruiserCutter = 0
         CorvetteSeq = 0
         ignoreToken = 0
       Else
         playsnd 9
      End If
   End If
   
   'move reaver one space - manual option
   If actionSeq = ASNavReav Then
      If Logic!player = player.ID Then
         If reaverMove(Index) Then actionSeq = ASNavReavEnd
      End If
   End If
   
   If actionSeq = ASNavReavBorder Then
      If Logic!player = player.ID And (getClearSector(Index) = "B" Or getClearSector(Index) = "R") Then
         MoveShip 6 + RollDice(NumOfReavers), Index
         actionSeq = ASNavReavEnd
       Else
         playsnd 9
      End If

   End If
   
   'move cruiser one space - manual option
   If actionSeq = ASNavCrus Then
      If validMove(5, Index) And Logic!player = player.ID And Not getHaven(Index) > 0 Then
         MoveShip 5, Index
         actionSeq = ASNavCrusEnd
       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCrusBorder Then
      If Logic!player = player.ID And getClearSector(Index) = "A" And getCruiserSector() <> Index And Not getHaven(Index) > 0 Then
         MoveShip 5, Index
         actionSeq = ASNavCrusEnd
      Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCrusOutlaw Then
      If Logic!player = player.ID And outlawExists(player.ID) And getCruiserSector() <> Index And Not getHaven(Index) > 0 Then
         MoveShip 5, Index
         actionSeq = ASNavCrusEnd
      Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCrusAdjacent Then
      If Logic!player = player.ID And getClearSector(Index) = "A" And Not getHaven(Index) > 0 And isAdjacent(player.ID, Index) Then
         If Not isAdjacent(player.ID, Index) Then PutMsg player.PlayName & " appears to be bending the rules"
         MoveShip 5, Index
         actionSeq = ASNavCrusEnd
      Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCorvAdjacent Then
      If Logic!player = player.ID And getClearSector(Index) <> "" And isAdjacent(player.ID, Index) Then
         MoveShip 6, Index
         actionSeq = ASNavCrusEnd
       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASNavCorvPlanetary Then
      If Logic!player = player.ID And getClearSector(Index) <> "" And Nz(varDLookup("PlanetID", "Planet", "SectorID=" & Index), 0) > 0 Then
         MoveShip 6, Index
         actionSeq = ASNavCrusEnd
       Else
         playsnd 9
      End If
   End If
   
   If actionSeq = ASResolveAlert Then
      If isAdjacent(player.ID, Index) Then
         playsnd 14
         resolveToken Index, True
         actionSeq = ASResolveAlertEnd
       Else
         playsnd 9
      End If
   End If
   
   
End Sub

Public Sub drawLine(ByVal mode, ByVal sector1, Optional ByVal sector2, Optional ByVal silent As Boolean = True)
Dim rst As New ADODB.Recordset
Dim SQL, X1, X2, Y1, Y2

   If sector1 = -1 Then
      Verse.LineB(mode).Visible = False
      Exit Sub
   End If
   If sector1 = -2 And Verse.LineB(mode).Visible = False Then Exit Sub
   
   If sector1 = 1 Then
      sector1 = getCruiserSector()
   ElseIf sector1 = 2 Then
      sector1 = getCorvetteSector()
   End If
   
   If sector1 = -2 Then
      X1 = Verse.LineB(mode).X1
      Y1 = Verse.LineB(mode).Y1
   Else
      SQL = "SELECT * "
      SQL = SQL & "FROM Board WHERE SectorID=" & sector1
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst.EOF Then
         X1 = rst!SLeft + Int(rst!SWidth / 2)
         Y1 = rst!STop + Int(rst!SHeight / 2)
      End If
      rst.Close
   
   End If
   
   SQL = "SELECT * "
   SQL = SQL & "FROM Board WHERE SectorID=" & sector2
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
       X2 = rst!SLeft + Int(rst!SWidth / 2)
       Y2 = rst!STop + Int(rst!SHeight / 2)
   End If
   rst.Close
      
   If Verse.LineB(mode).X1 = X1 And Verse.LineB(mode).Y1 = Y1 And Verse.LineB(mode).X2 = X2 And Verse.LineB(mode).Y2 = Y2 And Verse.LineB(mode).Visible = True Then
      Verse.LineB(mode).Visible = False
   Else
      Verse.LineB(mode).X1 = X1
      Verse.LineB(mode).Y1 = Y1
      Verse.LineB(mode).X2 = X2
      Verse.LineB(mode).Y2 = Y2
      Verse.LineB(mode).Visible = True
      'Verse.LineB(mode).ZOrder
      If Not silent Then playsnd 2
   End If
   
   Set rst = Nothing
End Sub

Private Sub animatePlayer(ByVal playerID)
Dim x
   For x = 1 To 4
      If x = playerID Then
         If Verse.Imag(x).Animate2.AnimationState = lvicAniCmdStop Then
            Verse.Imag(x).Animate2.StartAnimation
         End If
      Else
         If Verse.Imag(x).Animate2.AnimationState = lvicAniCmdStart Then
            Verse.Imag(x).Animate2.StopAnimation
            Verse.Imag(x).ImageIndex = 1
         End If
      End If
   Next x
End Sub

Private Sub refreshSolid()
Dim x, s As Boolean
Dim SQL
Dim rst As New ADODB.Recordset

   SQL = "SELECT * FROM Contact WHERE ContactID > 0 AND ContactID < 10 order by ContactID"
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenStatic, adLockReadOnly
   While Not rst.EOF
      x = rst!ContactID
      pic(x).Visible = True
      s = isSolid(player.ID, x)
      If Not (pic(x).Tag = "s" And s) Or (pic(x).Tag = "" And Not s) Then  'prevent reloading the same image
         pic(x).Picture = LoadPicture(App.Path & "\pictures\Solid" & IIf(s, "2", "1") & rst!Picture)
         pic(x).ToolTipText = IIf(s, "Solid with " & rst!ContactName & " - ", "") & rst!DealDescr & _
         IIf(x = 5, " Sells Fuel: $100", IIf(rst!cargo = 0, "", " Buys Cargo: $" & rst!cargo & " & Contraband: $" & rst!Contraband))
         pic(x).Tag = IIf(s, "s", "")
      End If
      rst.MoveNext
   Wend

End Sub

Private Sub clearPicMB()
Dim x
   For x = 1 To 4
      picMB(x).Visible = False
   Next x
End Sub

Private Function doGamble() As Integer
Dim SQL, reshuffle
Dim rst As New ADODB.Recordset

   'Read in the next NAV card and display either 1 or 2 options
   SQL = "SELECT CardID, Suit, Seq, CardName, reshuffle "
   SQL = SQL & "FROM MisbehaveDeck "
   SQL = SQL & "Where Seq > 5 "
   SQL = SQL & "ORDER BY Seq"
   If Left(datab, 16) = "Provider=MSDASQL" Then SQL = SQL & " LIMIT 1"
   rst.CursorLocation = adUseClient
   rst.Open SQL, DB, adOpenStatic, adLockReadOnly
   If Not rst.EOF Then
      doGamble = rst!suit
      
      'pull the card out of the deck
      DB.Execute "UPDATE MisbehaveDeck SET Seq = 5 WHERE CardID = " & CStr(rst!CardID)
      'rst!Seq = 5
      'rst.Update
      reshuffle = rst!reshuffle
      If reshuffle = 1 Then 'ready for next turn
         PutMsg player.PlayName & " Reshuffling MisbehaveDeck due to " & rst!CardName, player.ID, Logic!Gamecntr
         ShuffleDeck "Misbehave"
      End If

   Else
      PutMsg player.PlayName & " Reshuffling MisbehaveDeck due to end of deck", player.ID, Logic!Gamecntr
      ShuffleDeck "Misbehave"
      Exit Function
   End If
   rst.Close
      
End Function

Private Function doHavenSupplies()
Dim frmHavn As New frmHaven

   With frmHavn
      .Show 1
      frmAction.workdone = .success
   End With
   
   Set frmHavn = Nothing

End Function

Private Function loadBounties(mnuWorkPop As Object, ByVal SupplyID As Integer, ByVal sectorID, ByRef x)
Dim SQL As String
Dim rst As New ADODB.Recordset, rst2 As New ADODB.Recordset

   SQL = "SELECT ContactDeck.CardID, ContactDeck.Job1ID, ContactDeck.JobName, ContactDeck.FugitiveID "
   SQL = SQL & "FROM ContactDeck "
   SQL = SQL & "Where ContactDeck.ContactID = 10 And ContactDeck.Seq = " & DISCARDED
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      'see if the bounty is on board
      If hasCrew(player.ID, rst!FugitiveID) Then
         x = x + 1
         If mnuWorkPop.Count < x + 1 Then Load mnuWorkPop(x)
         mnuWorkPop(x).Caption = "Crew Bounty " & rst!JobName & " (" & CStr(rst!CardID) & ")"
         mnuWorkPop(x).Tag = CStr(rst!CardID * -1)
         
      ElseIf Nz(varDLookup("SupplyID", "SupplyDeck", "CrewID = " & rst!FugitiveID & " AND Seq = " & DISCARDED), 0) = SupplyID And SupplyID > 0 Then 'see if the bounty is at the current supply planet
         x = x + 1
         If mnuWorkPop.Count < x + 1 Then Load mnuWorkPop(x)
         mnuWorkPop(x).Caption = "Supply Bounty " & rst!JobName & " (" & CStr(rst!CardID) & ")"
         mnuWorkPop(x).Tag = CStr(rst!CardID * -1)
            
      Else 'or Rival crew
         SQL = "SELECT Players.PlayerID FROM Players INNER JOIN PlayerSupplies ON Players.PlayerID = PlayerSupplies.PlayerID WHERE Players.SectorID=" & sectorID & " AND PlayerSupplies.CardID = " & getCrewCardID(rst!FugitiveID)
         rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
         If Not rst2.EOF Then
            x = x + 1
            If mnuWorkPop.Count < x + 1 Then Load mnuWorkPop(x)
            mnuWorkPop(x).Caption = "Rival Crew Bounty " & rst!JobName & " (" & CStr(rst!CardID) & ")"
            mnuWorkPop(x).Tag = CStr(rst!CardID * -1)
         End If
         rst2.Close
      
      End If
      rst.MoveNext
   Wend
   rst.Close
   
   'look for Bounty Jumps
   SQL = "SELECT ContactDeck.CardID, ContactDeck.Job1ID, ContactDeck.JobName, ContactDeck.FugitiveID, Players.SectorID "
   SQL = SQL & "FROM Players INNER JOIN (PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) ON Players.PlayerID = PlayerJobs.PlayerID "
   SQL = SQL & "Where ContactDeck.ContactID = 10 And PlayerJobs.JobStatus =0 And Players.SectorID = " & sectorID & " And PlayerJobs.PlayerID <> " & player.ID & " And PlayerJobs.PlayerID > 0 And PlayerJobs.PlayerID <" & DISCARDED
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      x = x + 1
      If mnuWorkPop.Count < x + 1 Then Load mnuWorkPop(x)
      mnuWorkPop(x).Caption = "Bounty Jump " & rst!JobName & " (" & CStr(rst!CardID) & ")"
      mnuWorkPop(x).Tag = CStr(rst!CardID * -1)
      rst.MoveNext
   Wend
   rst.Close

End Function

Private Function doBountyHunt(ByVal CardID) As Integer
Dim CrewID, SupplyID, CrewCardID, killCrew, ShipID, JumpShipID

   CrewID = varDLookup("FugitiveID", "ContactDeck", "CardID=" & CardID)
   CrewCardID = getCrewCardID(CrewID)
   SupplyID = Nz(varDLookup("SupplyID", "SupplyDeck", "CardID = " & CrewCardID & " AND Seq = " & DISCARDED), 0)
   ShipID = Nz(varDLookup("PlayerID", "PlayerSupplies", "CardID = " & CrewCardID), 0)
   JumpShipID = Nz(varDLookup("PlayerID", "PlayerJobs", "CardID = " & CardID), 0)
   
   If hasCrew(player.ID, CrewID) Then 'on board fight
         'remove any Gear first
      DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & CrewID
      'delete the card to the players deck
      DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID =" & player.ID & " AND CardID = " & CrewCardID

      DB.Execute "UPDATE SupplyDeck SET Seq =0 WHERE CardID = " & CrewCardID
      PutMsg player.PlayName & " betrays " & getCrewName(CrewCardID) & " to claim a Bounty" & IIf(getCrewCount(player.ID) > 1, " and the Crew are not impressed", ""), player.ID, Logic!Gamecntr
      
      assignDeal player.ID, CardID
      'pullBounties
      If DrawDeck("Contact", 10, 1) Then PutMsg "New Bounty available"
      doDisgruntled player.ID, 7
   
   ElseIf SupplyID > 0 Then  'at a supply planet?
      If doShowDownSupply(CrewID) Then  'won
         DB.Execute "UPDATE SupplyDeck SET Seq =0 WHERE CardID = " & CrewCardID
         assignDeal player.ID, CardID
         'pullBounties
         If DrawDeck("Contact", 10, 1) Then PutMsg "New Bounty available"
      Else 'lost
         killCrew = varDLookup("FailKillCrew", "ContactDeck", "CardID=" & CardID)
         If killCrew > 0 Then
            doKillCrews player.ID, killCrew, True
         End If
      End If
      
   ElseIf ShipID > 0 Or JumpShipID > 0 Then 'attack Rival Ship or bounty jump
      PutMsg player.PlayName & " is attempting to board a rival ship, " & PlayCode(ShipID).PlayName, player.ID, Logic!Gamecntr
      'boarding test
      If doBoardingTest Then
         PutMsg player.PlayName & " has boarded a rival ship, " & PlayCode(ShipID + JumpShipID).PlayName & " to attempt to " & IIf(ShipID > 0, "claim", "jump") & " a Bounty for " & getCrewName(CrewCardID) & ". This calls for a Showdown!", player.ID, Logic!Gamecntr
         'showdown with ShipID
         doShowDownRival ShipID + JumpShipID, CardID, CrewCardID, CrewID, IIf(JumpShipID > 0, True, False)
      Else
         PutMsg player.PlayName & " failed to board the rival ship, " & PlayCode(ShipID).PlayName, player.ID, Logic!Gamecntr
      End If
  
   End If
   
End Function

Private Function doBoardingTest() As Boolean
Dim frmBT As New frmSkillSel, skill As Integer

   'select skill to use
   frmBT.setMode 1
   frmBT.Show vbModal, Me
   skill = frmBT.skill
   
   If doSkillTest(skill, 6, 0, 0, 1) = 0 Then
      doBoardingTest = True
   End If
      
End Function

Private Function doShowDownSupply(ByVal CrewID) As Boolean
Dim frmSS As New frmSkillSel, skillcnt As Integer, skill As Integer, Dice As Integer
Dim Cskillcnt As Integer, Cskill As Integer, CDice As Integer, x, cnt, msg, win As Boolean

   'select skill to use
   frmSS.Show vbModal, Me
   skill = frmSS.skill
   
   'pick highest skill of Fugitive
   For x = 1 To 3
      cnt = getSkillCrew(CrewID, cstrSkill(x))
      If cnt > Cskillcnt Then
         Cskillcnt = cnt
         Cskill = x
      End If
   Next x
      
   'Crazy River Tam (cardID 51/CrewID 32)
   If CrewID = 32 Then
      CDice = RollDice(6)
      Select Case CDice
      Case 3, 6 'fight
         Cskill = 1
         Cskillcnt = 2
      Case 4 'Tech
         Cskill = 2
         Cskillcnt = 2
      Case 5 'negot
         Cskill = 3
         Cskillcnt = 2
      End Select
   End If
   
   If Cskill = 0 Then Cskill = 3
   
   'fugitive rolls
   CDice = RollDice(6, True)
   msg = "Showdown: " & getCrewName(0, CrewID) & IIf(Cskillcnt = 0, " has no Skills", " uses the " & cstrSkill(Cskill) & " skill of " & Cskillcnt) & " and rolls a " & CDice & " for a total of " & CStr(Cskillcnt + CDice)
   
   If hasCrew(player.ID, 91) And CDice > 1 Then 'showdown re-roll
      If MessBox(msg & vbNewLine & "Chari can force them to re-roll if you want", "Showdown", "Re-Roll", "Leave", 91, 0, 0, Dice) = 0 Then
         CDice = RollDice(6, True)
         PutMsg player.PlayName & " uses Chari's Skills to force a reRoll and they got a " & CStr(CDice) & " for a total of " & CStr(Cskillcnt + CDice), player.ID, Logic!Gamecntr, True, 91, 0, 0, 0, 0, CDice
      End If
   Else
      PutMsg msg, player.ID, Logic!Gamecntr, True, CrewID, 0, 0, 0, 0, CDice, Cskill
   End If
   
   skillcnt = doWorkSkillTest(Dice, skill, Cskillcnt + CDice + 1, 0, 2)  'result includes dice score
   
   win = (skillcnt = 0)
      
   msg = "Showdown: " & player.PlayName & IIf(win, " apprehends " & getCrewName(0, CrewID), " Botches the job!")
   PutMsg msg, player.ID, Logic!Gamecntr, True, CrewID, 0, 0, 0, 0, Dice, skill
   
   doShowDownSupply = win
   
   If Not (frmJob Is Nothing) Then frmJob.refreshJobs
   If Not (frmDeal Is Nothing) Then frmDeal.RefreshDeals
End Function

Private Function doShowDownRival(ByVal DefenderID As Integer, ByVal CardID As Integer, ByVal CrewCardID As Integer, ByVal CrewID As Integer, Optional ByVal bountyJump As Boolean = False) As Boolean
Dim frmSD As frmShowdown, winShowdown As Boolean, killCrew
Dim frmSS As New frmSkillSel, Skilltype As Integer, skill As Integer, Dice As Integer
   'pick a Skill
   frmSS.Show vbModal, Me
   Skilltype = frmSS.skill
   
   skill = doWorkSkillTest(Dice, Skilltype, 0, 0, 2) 'returns skill + dice
   'clear the decks
   DB.Execute "Delete from ShowdownScores"
   DB.Execute "Delete from ShowdownGear"
   
   'throw a hook for the Rival Ship to pickup the boarding alert and init the showdown
   DB.Execute "UPDATE GameSeq SET Seq = 'F', HostAccept = 0, ClientAccept = 0, Trader = " & CStr(DefenderID)
   Logic.Requery
   'Logic!Seq = "F"
   'Logic!HostAccept = 0
   'Logic!ClientAccept = 0
   'Logic!trader = DefenderID
   'Logic.Update
   Set frmSD = New frmShowdown
   With frmSD
      .isHost = True
      .Skilltype = Skilltype
      .skill = skill - Dice
      .Dice = Dice
      .OpponentID = DefenderID
      Set .imgCrew.Picture = LoadPicture(App.Path & "\pictures\" & varDLookup("Picture", "Crew", "CrewID=" & CrewID))
      .Show vbModal
      winShowdown = .winShowdown
   End With
   
   'if won, wanted crew seized, and bounty card claimed
   If winShowdown Then
      'remove any Gear first
      If bountyJump Then
         'remove the card to the rival players deck
         DB.Execute "DELETE FROM PlayerJobs WHERE CardID  = " & CardID
         If MessBox("Do you want to Rescue " & getCrewName(CrewCardID) & " instead of claiming the bounty?", "Bounty Jump", "Rescue", "Bounty", CrewID) = 0 Then
            If CrewCapacity(player.ID) - getCrewCount(player.ID) > 0 Then 'bring into crew
               DB.Execute "UPDATE SupplyDeck SET Seq =" & player.ID & " WHERE CardID = " & CrewCardID
               DB.Execute "INSERT INTO PlayerSupplies (PlayerID, CardID) VALUES (" & player.ID & ", " & CrewCardID & ")"
               
            Else 'discard pile
               PutMsg "No room for " & getCrewName(CrewCardID) & " onboard! Sent them home instead.", player.ID, Logic!Gamecntr, True, CrewID
               DB.Execute "UPDATE SupplyDeck SET Seq =" & DISCARDED & " WHERE CardID = " & CrewCardID
                           
            End If
            'place bounty card at bottom of deck
            DB.Execute "UPDATE ContactDeck SET Seq =" & getBountyMaxSeq & " WHERE CardID = " & CardID
         Else 'take the bounty card
            assignDeal player.ID, CardID
         End If
      Else
         DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & CrewID
         'delete the card to the players deck
         DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID =" & DefenderID & " AND CardID = " & CrewCardID
      
         DB.Execute "UPDATE SupplyDeck SET Seq =0 WHERE CardID = " & CrewCardID
         assignDeal player.ID, CardID
         'pullBounties
         If DrawDeck("Contact", 10, 1) Then PutMsg "New Bounty available"
      End If
      'Agent McGinnis is a bad loser
      If hasCrew(DefenderID, 92) Then issueWarrant player.ID, DefenderID
      
   Else    'if lost, execute lost bounty orders - kill a crew?
      killCrew = varDLookup("FailKillCrew", "ContactDeck", "CardID=" & CardID)
      If killCrew > 0 Then
         doKillCrews player.ID, killCrew, True
      End If
      'Agent McGinnis is a bad loser
      If hasCrew(player.ID, 92) Then
         issueWarrant DefenderID, player.ID
      End If
   End If
   DB.Execute "UPDATE GameSeq SET Seq = 'R', Trader = 0"
   Logic.Requery
   'Logic!Seq = "R"
   'Logic!trader = 0
   'Logic.Update
   
End Function

Private Sub doShowdownDefend(ByVal AttackerID As Integer, ByVal Skilltype)
Dim frmSD As frmShowdown, skill As Integer, Dice As Integer, winShowdown As Boolean

   skill = doWorkSkillTest(Dice, Skilltype, 0, 0, 1) 'returns skill + dice

   Set frmSD = New frmShowdown
   With frmSD
      .Skilltype = Skilltype
      .skill = skill - Dice
      .Dice = Dice
      .OpponentID = AttackerID
      'Set .imgCrew.Picture = LoadPicture(App.Path & "\pictures\" & varDLookup("Picture", "Crew", "CrewID=" & CrewCardID))
      .Show vbModal
      winShowdown = .winShowdown
   End With
   DB.Execute "UPDATE GameSeq SET Trader = 0"
   Logic.Requery
   'Logic.Update "Trader", 0
   'reset OffJob status
   clearOffJob player.ID
   If Not winShowdown Then
      If Not (frmShip Is Nothing) Then frmShip.RefreshShips
      If Not (frmJob Is Nothing) Then frmJob.refreshJobs
      If Not (frmDeal Is Nothing) Then frmDeal.RefreshDeals
   End If
      
End Sub

Public Function showBoarded(ByVal AttackerID As Integer) As Integer
Dim frmB As New frmBoarded
   frmB.thisplayer = AttackerID
   frmB.Show vbModal
   showBoarded = frmB.result
   Set frmB = Nothing
End Function


Private Sub issueWarrant(ByVal takerID As Integer, ByVal giverID As Integer)

   DB.Execute "UPDATE Players set Warrants = Warrants + 1, Solid5 = 0 WHERE PlayerID = " & CStr(takerID)  'clear Harken Solid
   PutMsg PlayCode(giverID).PlayName & "'s Agent McGinnis didn't take kindly to the Showdown result and has issued a Warrant to " & PlayCode(takerID).PlayName, giverID, Logic!Gamecntr, True, 92
End Sub

' use for Nav type Skill Tests (not Work)
Private Function doSkillTest(ByVal skill As Integer, ByVal skillwin As Integer, Optional ByVal skillint As Integer = 0, Optional bribe As Integer = 0, Optional boarding As Integer = 0) As Integer
Dim skillcnt, skilldiscards, x
Dim Dice As Integer, riverskill As Integer, extraSkill As Integer, frmDiscardGr As frmDiscardGear

         'Crazy River Tam (cardID 51/CrewID 32)
         If hasCrew(player.ID, 32) Then
            Dice = RollDice(6)
            If hasCrew(player.ID, 33) Then 'simon adds 2 to her rolls
               Dice = Dice + 2
            End If
            Select Case Dice
            Case 1, 2 'stay onboard
               DB.Execute "UPDATE PlayerSupplies SET OffJob = 1 WHERE CardID = 51"
               If Not (frmShip Is Nothing) Then frmShip.RefreshShips  'update display
               PutMsg player.PlayName & "'s River Tam hides in her bunk and won't be helpin'", player.ID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, Dice
            Case 3 'fight
               If skill = 1 Then
                  riverskill = 2
               End If
            Case 4 'Tech
               If skill = 2 Then
                  riverskill = 2
               End If
            Case 5 'negot
               If skill = 3 Then
                  riverskill = 2
               End If
            Case Else 'any skill
                  riverskill = 2
            End Select
            If riverskill = 2 Then
               PutMsg player.PlayName & "'s River Tam" & IIf(hasCrew(player.ID, 33), ", encouraged by Simon,", "") & " channels the " & cstrSkill(skill) & " skill + 2", player.ID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, Dice
            ElseIf Dice > 2 Then
               PutMsg player.PlayName & "'s River Tam ain't gettin' involved this time", player.ID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, Dice, skill
            End If
         End If
         
         extraSkill = hasGearCard(player.ID, 24)
         If extraSkill > 0 Then 'we got one or more
            If MessBox("Do you wish to Eat the Fruity Bar and add 1 to the Test Roll?", "Extra Bite", "Yes", "No", 0, 24) = 0 Then
               doDiscardGear player.ID, extraSkill
               extraSkill = 1
            Else
               extraSkill = 0
            End If
         End If

         x = hasGearCrew(player.ID, 28) 'Mal's Brown Coat
         If x > 0 And varDLookup("Disgruntled", "Crew", "CrewID=" & x) > 0 And varDLookup("Fight", "Crew", "CrewID=" & x) > 0 And skill = 3 Then
            extraSkill = extraSkill + varDLookup("Fight", "Crew", "CrewID=" & x)
            PutMsg player.PlayName & "'s Disgruntled Crew wearing the Brown Coat adds their Fight skills to the Negotiation", player.ID, Logic!Gamecntr, True, 0, 28
         End If
            
         If skill = 1 Then
            removeDigruntled player.ID, skill
         End If
         'if bribes are acceptable, ask for $100 a point
         If skill = 3 And (bribe = 1 Or hasPerkAttributeValue(player.ID, "Bribe", skill)) Then
            Do
               bribe = InputBoxx("They accept Bribes, $100 per skill point" & vbNewLine & vbNewLine & "Enter the number of POINTS you would bribe with..", "Money Talks", "0", getLeader())
               If bribe > 20 Then
                  MessBox "Seems a bit much don't ya think? Try that again..", "Too much!", "Ooops", "", getLeader()
               ElseIf bribe * 100 <= getMoney(player.ID) Then 'can pay
                  getMoney player.ID, (bribe * 100 * -1)
                  Exit Do
               Else
                  MessBox "Why you low-down thief, whatcha tryin' to pull?  Try again!", "Insufficient dough!", "Sorry", "", getLeader()
               End If
            Loop
         End If
            
         'Roll the flippin Dice already!!!
         Dice = RollDice(6, IIf(skill = 2 And hasCrew(player.ID, 55), False, True))
         
         If skill = 1 And hasGear(player.ID, 47) Then ' Zoe's Mare's Leg Rifle -When making a Fight Test, roll two dice and use the highest.
            x = RollDice(6, True)
            If x > Dice Then
               PutMsg player.PlayName & " had rolled a " & CStr(Dice) & " so using Zoe's Mare's Leg Rifle rerolled a " & CStr(x), player.ID, Logic!Gamecntr, True, 0, 47, 0, 0, 0, x, skill
               Dice = x
            End If
         End If
         
         If Dice = 1 Then  'reroll ones?
            If hasGear(player.ID, 56) And boarding = 0 Then 'Wash's Dinos
              
               Do While Dice = 1
                  Dice = RollDice(6, True)
               Loop
               PutMsg player.PlayName & " uses Wash's Lucky Dinosaurs to reRoll a 1 and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, 0, 56, 0, 0, 0, Dice, skill
               
            ElseIf hasGear(player.ID, 35) And skill = 1 Then 'Inara's Bow
               x = hasGearCrew(player.ID, 35)
               If x > 0 Then
                  If hasCrewAttribute(player.ID, "Companion", 0, x) Then
                     Do While Dice = 1
                        Dice = RollDice(6, True)
                     Loop
                     PutMsg player.PlayName & " uses Inara's Bow to reRoll a 1 and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, 0, 35, 0, 0, 0, Dice, skill
                  End If
               End If
            End If
         End If

         'Zoe's skill can reroll a fight test
         If skill = 1 And Dice < 6 Then
            x = getPerkAttributeCrew(player.ID, "RerollFight")
            If x > 0 Then
               If MessBox("You rolled a " & Dice & vbNewLine & "Your Fight Skills allow you a re-roll, do you want to take that extra chance?", "Re-Roll option", "Re-roll", "Keep", x, 0, 0, Dice) = 0 Then
                  Dice = RollDice(6, True)
                  PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, x, 0, 0, 0, 0, Dice, skill
               End If
            End If
         End If
         'Kaylee can reroll a Tech test
         If skill = 2 And Dice < 6 Then
            x = getPerkAttributeCrew(player.ID, "RerollTech")
            If x > 0 Then
               If MessBox("You rolled a " & Dice & vbNewLine & "Your Tech Skills allow you a re-roll, do you want to take that chance?", "Re-Roll option", "Re-roll", "Keep", x, 0, 0, Dice) = 0 Then
                  Dice = RollDice(6, IIf(hasCrew(player.ID, 55), False, True))
                  PutMsg player.PlayName & " uses extra Tech Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, x, 0, 0, 0, 0, Dice, skill
               End If
            End If
         End If
         'Inara can reroll a negotiate test
         If skill = 3 And Dice < 6 Then
            x = getPerkAttributeCrew(player.ID, "RerollNegotiate")
            If x > 0 Then
               If MessBox("You rolled a " & Dice & vbNewLine & "Your Negotiation Skills allow you a re-roll, do you want to take that chance?", "Re-Roll option", "Re-roll", "Keep", x, 0, 0, Dice) = 0 Then
                  Dice = RollDice(6, True)
                  PutMsg player.PlayName & " uses extra Negotiation Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, x, 0, 0, 0, 0, Dice, skill
               End If
            End If
         End If
         
         If skill = 1 And hasGear(player.ID, 45) And Dice < 6 Then 'yolanda's pistol - Discard to re-roll a Fight Test.
            If MessBox("You rolled a " & Dice & vbNewLine & "Yolanda's pistol allows you a re-roll, do you want to Discard the Pistol to take that extra chance?", "Re-Roll option", "Re-roll", "Keep", 0, 45, 0, Dice) = 0 Then
               doDiscardGear player.ID, hasGearCard(player.ID, 45)
               Dice = RollDice(6, True)
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, 0, 45, 0, 0, 0, Dice, skill
            End If
         End If
         
         If skill = 1 And hasGear(player.ID, 48) And Dice < 6 Then 'Extra Ammo Clip - Discard to re-roll a Fight Test.
            If MessBox("You rolled a " & Dice & vbNewLine & "Extra Ammo Clips allow you a re-roll, do you want to Discard the Clips to take that extra chance?", "Re-Roll option", "Re-roll", "Keep", 0, 48, 0, Dice) = 0 Then
               doDiscardGear player.ID, hasGearCard(player.ID, 48)
               Dice = RollDice(6, True)
               PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, Dice, skill
            End If
         End If
         
         'skillcnt = getSkill(player.ID, cstrSkill(Skill)) + dice
         '----------------------------------------- see if we need to use the discardable skills & keywords...
         
         'skillwin = rst!win
         'skillint = rst!Intermediate

         'get our skill totals
         skillcnt = getSkill(player.ID, cstrSkill(skill), 0, True) + Dice + bribe + riverskill + extraSkill
         skilldiscards = getSkillDiscards(player.ID, cstrSkill(skill))
         
         '-----------------------------------------
         If skillcnt < skillwin And skillcnt + skilldiscards >= skillwin Then 'we're in trouble 'we could use some help
            If MessBox("Rolled a " & CStr(Dice) & vbNewLine & "With the help of " & skillwin - skillcnt & " single use skill points, we can succeed." & vbNewLine & "Do you want to use discardable Gear items for this?", "Skill Test Trouble", "Yes", "No", getLeader()) = 0 Then
               'show a list of gear to pick from up to or exceeding the value skillwin - skillcnt
               Set frmDiscardGr = New frmDiscardGear
               frmDiscardGr.nbrSelect = skillwin - skillcnt
               frmDiscardGr.skill = cstrSkill(skill)
               frmDiscardGr.Show 1
               'then add selected skill points to skillcnt, discard gear, and go on...
               skillcnt = skillcnt + frmDiscardGr.nbrSelected
            End If
                  
         ElseIf skillint > 0 And skillcnt < skillint And skillcnt + skilldiscards >= skillint Then 'we're in trouble 'we could use some help
            If MessBox("Rolled a " & CStr(Dice) & vbNewLine & "With the help of " & skillint - skillcnt & " single use skill points, we can make the intermediate outcome." & vbNewLine & "Do you want to use discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
               'show a list of gear to pick from up to or exceeding the value skillint - skillcnt
               Set frmDiscardGr = New frmDiscardGear
               frmDiscardGr.nbrSelect = skillint - skillcnt
               frmDiscardGr.skill = cstrSkill(skill)
               frmDiscardGr.Show 1
               'then add selected skill points to skillcnt, discard gear, and go on...
               skillcnt = skillcnt + frmDiscardGr.nbrSelected
            End If
         
         End If
         
         If hasGear(player.ID, 32) And skill = 1 And skillcnt < skillwin Then
            If MessBox("The Fights not going so well with a skill score of " & skillcnt & vbNewLine & "Simon's Sonic Stun Baton might turn things around, wanna try another Thrillin' Heroics Roll and Discard the Baton?", "Stun Baton to the Fight", "Yes", "No", 0, 32) = 0 Then
               skillcnt = RollDice(6) + 6
               doDiscardGear player.ID, hasGearCard(player.ID, 32)
            End If
         End If
         '-----------------------------------------
         
         If skillcnt >= skillwin Then
            doSkillTest = 0
         ElseIf skillcnt >= skillint And skillint > 0 Then
            doSkillTest = 1
         Else 'you lose
            doSkillTest = 2
         End If
         
         PutMsg player.PlayName & "'s Nav log: Rolls a " & Dice & " with added " & cstrSkill(skill) & " skill points of " & CStr(skillcnt - Dice) & " for a total of " & skillcnt & " to " & IIf(doSkillTest = 0, "succeed :^)", IIf(doSkillTest = 1, "partially succeed :^|", "lose :^(")), player.ID, Logic!Gamecntr, True, getLeader(), 0, 0, 0, 0, Dice, skill

End Function

'mode = 0 normal work test. 1= showdown defend. 2= showdown attack
Private Function doWorkSkillTest(ByRef Dice As Integer, ByVal WSkill As Integer, Optional ByVal skillwin As Integer = 0, Optional ByVal skillint As Integer = 0, Optional ByVal mode As Integer = 0) As Integer
Dim x, bribe As Integer
Dim riverskill As Integer, frmDiscardGr As frmDiscardGear
Dim skillcnt, skilldiscards, extraSkill As Integer

            'Stitch & Sheydra can change a Fight to a Nego once per Job
            If WSkill = 3 And hasCrew(player.ID, 27) And Not usedStitchSkill Then
               If MessBox("Stitch wants to change this Negotiation to a Fight.  Do you want to use those one-time skills now?", "Negotiate -> Fight", "Bring it", "Not now", 27) = 0 Then
                  WSkill = 1
                  usedStitchSkill = True
                  PutMsg player.PlayName & " uses Stitch's one time Negotiation to Fight Skills", player.ID, Logic!Gamecntr, True, 27
               End If
            End If
            If WSkill = 1 And getPerkAttributeCrew(player.ID, "ChangeTestType") > 0 And Not usedStitchSkill Then
               If MessBox("Sheydra wants to Negotiate instead of Fight.  Do you want to use those one-time skills now?", "Fight -> Negotiate", "Yes", "Not now", 66) = 0 Then
                  WSkill = 3
                  usedStitchSkill = True
                  PutMsg player.PlayName & " uses Sheydra's one time Fight to Negotiation Skills", player.ID, Logic!Gamecntr, True, 66
               End If
            End If
            
            'Crazy River Tam (cardID 51/CrewID 32)
            If hasCrew(player.ID, 32) Then
               Dice = RollDice(6)
               If hasCrew(player.ID, 33) Then  'simon adds 2 to her rolls
                  Dice = Dice + 2
               End If
               Select Case Dice
               Case 1, 2 'stay onboard
                  DB.Execute "UPDATE PlayerSupplies SET OffJob = 1 WHERE CardID = 51"
                  If Not (frmShip Is Nothing) Then frmShip.RefreshShips  'update display
                  PutMsg player.PlayName & "'s River Tam cowers onboard and won't be workin' anymore today", player.ID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, Dice
               Case 3 'fight
                  If WSkill = 1 Then
                     riverskill = 2
                  End If
               Case 4 'Tech
                  If WSkill = 2 Then
                     riverskill = 2
                  End If
               Case 5 'negot
                  If WSkill = 3 Then
                     riverskill = 2
                  End If
               Case Else 'any skill
                     riverskill = 2
               End Select
               If riverskill = 2 Then
                  PutMsg player.PlayName & "'s River Tam" & IIf(hasCrew(player.ID, 33), ", encouraged by Simon,", "") & " channels the " & cstrSkill(WSkill) & " skill + 2", player.ID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, Dice
               ElseIf Dice > 2 Then
                  PutMsg player.PlayName & "'s River Tam ain't workin' this time", player.ID, Logic!Gamecntr, True, 32, 0, 0, 0, 0, Dice
               End If
            End If
            
            extraSkill = hasGearCard(player.ID, 24)
            If extraSkill > 0 Then 'we got one or more
               If MessBox("Do you wish to Eat the Fruity Bar and add 1 to the Test Roll?", "Extra Bite", "Yes", "Not now", 0, 24) = 0 Then
                  doDiscardGear player.ID, extraSkill
                  extraSkill = 1
               Else
                  extraSkill = 0
               End If
            End If
            
            x = hasGearCrew(player.ID, 28) 'Mal's Brown Coat
            If x > 0 And varDLookup("Disgruntled", "Crew", "CrewID=" & x) > 0 And varDLookup("Fight", "Crew", "CrewID=" & x) > 0 And WSkill = 3 Then
               extraSkill = extraSkill + varDLookup("Fight", "Crew", "CrewID=" & x)
               PutMsg player.PlayName & "'s Disgruntled Crew wearing the Brown Coat adds their Fight skills to the Negotiation", player.ID, Logic!Gamecntr, True, 0, 28
            End If
            
            If mode > 0 Then 'Showdown with Fight skill
               extraSkill = extraSkill + getPerkCount(cstrSkill(WSkill), "SHOWDOWN", mode) 'Posse Lawman
            End If
            
            If WSkill = 1 Then
               removeDigruntled player.ID, WSkill
            End If
            
            'if card accepts a bribe, ask for $100 a point
            If WSkill = 3 And mode = 0 And hasPerkAttributeValue(player.ID, "Bribe", WSkill) Then
               Do
                  bribe = InputBoxx("They accept Bribes, $100 per skill point" & vbNewLine & vbNewLine & "Enter the number of POINTS you would bribe with..", "Money Talks", "0", getLeader())
                  If bribe > 20 Then
                     MessBox "Seems a bit much don't ya think? Try that again..", "Too much!", "Ooops", "", getLeader()
                  ElseIf bribe * 100 <= getMoney(player.ID) Then 'can pay
                     getMoney player.ID, (bribe * 100 * -1)
                     Exit Do
                  Else
                     MessBox "Why you low-down thief, whatcha tryin' to pull?  Try again!", "Insufficient dough!", "Sorry", "", getLeader()
                  End If
               Loop
            End If
            
            '<<<<<<<<<<<<<< ROLL THE DICE >>>>>>>>>>>>>>>>>>>>>>>>>
            Dice = RollDice(6, IIf(WSkill = 2 And hasCrew(player.ID, 55), False, True))
            
            If WSkill = 1 And hasGear(player.ID, 47) Then ' Zoe's Mare's Leg Rifle -When making a Fight Test, roll two dice and use the highest.
               x = RollDice(6, True)
               If x > Dice Then
                  PutMsg player.PlayName & " had rolled a " & CStr(Dice) & " so using Zoe's Mare's Leg Rifle rerolled a " & CStr(x), player.ID, Logic!Gamecntr, True, 0, 47, 0, 0, 0, x
                  Dice = x
               End If
            End If
            
            If Dice = 1 Then  'reroll ones?
               If hasGear(player.ID, 35) And WSkill = 1 Then 'Inara's Bow
                  x = hasGearCrew(player.ID, 35)
                  If x > 0 Then
                     If hasCrewAttribute(player.ID, "Companion", 0, x) Then
                        Do While Dice = 1
                           Dice = RollDice(6, True)
                        Loop
                        PutMsg player.PlayName & "'s Companion uses Inara's Bow to reRoll a 1 and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, 0, 35, 0, 0, 0, Dice
                     End If
                  End If
               End If
            End If
            
            'Zoe's skill can reroll a Fight test
            If WSkill = 1 And Dice < 6 Then
               x = getPerkAttributeCrew(player.ID, "RerollFight")
               If x > 0 Then
                  If MessBox("You rolled a " & Dice & vbNewLine & "Your Fight Skills allow you a re-roll, do you want to take that extra chance?", "Re-Roll option", "Re-roll", "Keep", x, 0, 0, Dice) = 0 Then
                     Dice = RollDice(6, True)
                     PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, x, 0, 0, 0, 0, Dice
                  End If
               End If
            End If
            
            'Kaylee can reroll a Tech test
            If WSkill = 2 And Dice < 6 Then
               x = getPerkAttributeCrew(player.ID, "RerollTech")
               If x > 0 Then
                  If MessBox("You rolled a " & Dice & vbNewLine & "Your Tech Skills allow you a re-roll, do you want to take that chance?", "Re-Roll option", "Re-roll", "Keep", x, 0, 0, Dice) = 0 Then
                     Dice = RollDice(6, IIf(WSkill = 3 And hasCrew(player.ID, 55), False, True))
                     PutMsg player.PlayName & " uses extra Tech Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, x, 0, 0, 0, 0, Dice
                  End If
               End If
            End If
         
            'Inara can reroll a negotiate test
            If WSkill = 3 And Dice < 6 Then
               x = getPerkAttributeCrew(player.ID, "RerollNegotiate")
               If x > 0 Then
                  If MessBox("You rolled a " & Dice & vbNewLine & "Your Negotiation Skills allow you a re-roll, do you want to take that chance?", "Re-Roll option", "Re-roll", "Keep", x, 0, 0, Dice) = 0 Then
                     Dice = RollDice(6, True)
                     PutMsg player.PlayName & " uses extra Negotiation Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, x, 0, 0, 0, 0, Dice
                  End If
               End If
            End If
            
            If WSkill = 1 And hasGear(player.ID, 45) And Dice < 6 Then 'yolanda's pistol - Discard to re-roll a Fight Test.
               If MessBox("You rolled a " & Dice & vbNewLine & "Yolanda's pistol allows you a re-roll, do you want to Discard the Pistol to take that extra chance?", "Re-Roll option", "Re-roll", "Keep", 0, 45, 0, Dice) = 0 Then
                  doDiscardGear player.ID, hasGearCard(player.ID, 45)
                  Dice = RollDice(6, True)
                  PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, 0, 45, 0, 0, 0, Dice
               End If
            End If
            
            If WSkill = 1 And hasGear(player.ID, 48) And Dice < 6 Then 'Extra Ammo Clip - Discard to re-roll a Fight Test.
               If MessBox("You rolled a " & Dice & vbNewLine & "Extra Ammo Clips allow you a re-roll, do you want to Discard the Clips to take that extra chance?", "Re-Roll option", "Re-roll", "Keep", 0, 48, 0, Dice) = 0 Then
                  doDiscardGear player.ID, hasGearCard(player.ID, 48)
                  Dice = RollDice(6, True)
                  PutMsg player.PlayName & " uses extra Fight Skills to reRoll and got a " & CStr(Dice), player.ID, Logic!Gamecntr, True, 0, 48, 0, 0, 0, Dice
               End If
            End If
            
                                    
            '----------------------------------------- see if we need to use the discardable skills & keywords...
           
            'skillwin = rst!win
            'skillint = rst!Intermediate
   
            'get our skill totals, no Kosherized rules in play for Jobs, only MB
            skillcnt = getSkill(player.ID, cstrSkill(WSkill), 0, True) + Dice + riverskill + extraSkill + bribe
            
            If skillwin = 0 Then  ' skip the discard process and return the skill score - use for Rival Showdown
               doWorkSkillTest = skillcnt
               Exit Function
            End If
            
            skilldiscards = getSkillDiscards(player.ID, cstrSkill(WSkill))
               
            '-----------------------------------------
            If skillcnt < skillwin And skillcnt + skilldiscards >= skillwin Then 'we're in trouble 'we could use some help
               If MessBox("With the help of " & skillwin - skillcnt & " skill points, we can succeed" & vbNewLine & "Do you want to use a discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
                  'show a list of gear to pick from up to or exceeding the value skillwin - skillcnt
                  Set frmDiscardGr = New frmDiscardGear
                  frmDiscardGr.nbrSelect = skillwin - skillcnt
                  frmDiscardGr.skill = cstrSkill(WSkill)
                  frmDiscardGr.Caption = "Select single use Gear to provide at least " & CStr(frmDiscardGr.nbrSelect) & " skill points"
                  frmDiscardGr.Show 1
                  'then add selected skill points to skillcnt, discard gear, and go on...
                  skillcnt = skillcnt + frmDiscardGr.nbrSelected
               End If
            ElseIf skillcnt < skillint And skillcnt + skilldiscards >= skillint Then 'we're in trouble 'we could use some help
               If MessBox("With the help of " & skillint - skillcnt & " skill points, we can make the intermediate outcome" & vbNewLine & "Do you want to use discardable Gear items for this?", "Skill Test", "Yes", "No", getLeader()) = 0 Then
                  'show a list of gear to pick from up to or exceeding the value skillint - skillcnt
                  Set frmDiscardGr = New frmDiscardGear
                  frmDiscardGr.nbrSelect = skillint - skillcnt
                  frmDiscardGr.skill = cstrSkill(WSkill)
                  frmDiscardGr.Caption = "Select single use Gear to provide at least " & CStr(frmDiscardGr.nbrSelect) & " skill points"
                  frmDiscardGr.Show 1
                  'then add selected skill points to skillcnt, discard gear, and go on...
                  skillcnt = skillcnt + frmDiscardGr.nbrSelected
               End If
            End If
            '-----------------------------------------
            
            If hasGear(player.ID, 32) And WSkill = 1 And skillcnt < skillwin Then '  use Simon's Sonic Stun Baton??
               If MessBox("The Fights not going so well with a skill score of " & skillcnt & vbNewLine & "Simon's Sonic Stun Baton might turn things around, wanna try another Thrillin' Heroics Roll and Discard the Baton?", "Stun Baton to the Fight", "Yes", "No", 0, 32) = 0 Then
                  Dice = RollDice(6) + 6
                  skillcnt = getSkill(player.ID, cstrSkill(WSkill), 0, True) + Dice + riverskill + extraSkill
                  doDiscardGear player.ID, hasGearCard(player.ID, 32)
               End If
            End If
            
            If skillcnt >= skillwin Then
               doWorkSkillTest = 0
            ElseIf skillcnt >= skillint And skillint > 0 Then
               doWorkSkillTest = 1
            Else 'you lose :(
               doWorkSkillTest = 3
            End If
            PutMsg player.PlayName & "'s Work log: Rolls a " & Dice & " with added " & cstrSkill(WSkill) & " skill points of " & CStr(skillcnt - Dice) & " for a total of " & skillcnt & " to " & IIf(doWorkSkillTest = 0, "succeed :^)", IIf(doWorkSkillTest = 1, "part win", "lose :^(")), player.ID, Logic!Gamecntr, True, getLeader(), getLeader(), 0, 0, 0, Dice, WSkill
            
End Function

Public Function getPerkCount(ByVal Attrib As String, ByVal keyword As String, ByVal mode As Integer) As Integer
Dim rst As New ADODB.Recordset
Dim SQL, x As Integer

   SQL = "SELECT Perk.Fight, Crew.CrewName, SupplyDeck.CrewID "
   SQL = SQL & "FROM Perk INNER JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) "
   SQL = SQL & "ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
   SQL = SQL & "WHERE PlayerSupplies.OffJob=0 AND Perk." & Attrib & " > 0 AND Perk." & Attrib & " <= " & mode & " AND Perk.Keyword='" & keyword & "' AND PlayerSupplies.PlayerID=" & player.ID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      x = rst.Fields(Attrib)
      getPerkCount = getPerkCount + x
      PutMsg player.PlayName & "'s " & rst!CrewName & " adds " & x & " more to the Showdown " & Attrib & " skill", player.ID, Logic!Gamecntr, True, rst!CrewID
      rst.MoveNext
   Wend
   rst.Close
End Function

Public Sub doLawmenOffJob()
Dim rst As New ADODB.Recordset
Dim SQL, refrsh As Boolean

   SQL = "SELECT Crew.*, SupplyDeck.CardID FROM Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck "
   SQL = SQL & "ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID WHERE "
   SQL = SQL & "PlayerSupplies.OffJob=0 AND Crew.Lawman = 1 AND PlayerSupplies.playerID = " & player.ID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF

      DB.Execute "UPDATE PlayerSupplies SET OffJob = 1 WHERE CardID = " & rst!CardID
      
      PutMsg player.PlayName & "'s Lawman " & rst!CrewName & " refuses to work this illegal Job and stays on the ship.", player.ID, Logic!Gamecntr, True, rst!CrewID
      refrsh = True
      rst.MoveNext
   Wend
   rst.Close
   If refrsh Then
      If Not (frmShip Is Nothing) Then frmShip.RefreshShips  'update display
   End If
End Sub

