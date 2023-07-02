VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmNav 
   BackColor       =   &H00400000&
   Caption         =   "Nav Deck"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   Icon            =   "frmNav.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   4755
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "discard"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "discard and re-draw the next Nav Card"
      Top             =   6510
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "option 2"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   1
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "option 2"
      Top             =   4240
      Width           =   3585
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Those Stars Look Right to You?"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   0
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "option 1"
      Top             =   1650
      Width           =   3615
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl SkillImg 
      Height          =   390
      Index           =   1
      Left            =   3210
      Top             =   6160
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   688
      Trans           =   100663295
      Effects         =   "frmNav.frx":030A
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl SkillImg 
      Height          =   390
      Index           =   0
      Left            =   3210
      Top             =   3560
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   688
      Trans           =   100663295
      Effects         =   "frmNav.frx":0322
   End
   Begin VB.Label lblUnseen 
      BackStyle       =   0  'Transparent
      Caption         =   "unseen"
      ForeColor       =   &H0000C0C0&
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label lblDetail 
      BackColor       =   &H002C1412&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNav.frx":033A
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFDD&
      Height          =   1695
      Index           =   1
      Left            =   210
      TabIndex        =   5
      Top             =   5040
      Width           =   3765
      WordWrap        =   -1  'True
   End
   Begin XDOCKFLOATLibCtl.FDPane FDPane1 
      Height          =   420
      Left            =   5130
      TabIndex        =   3
      Top             =   2790
      Visible         =   0   'False
      Width           =   420
      _cx             =   2010972901
      _cy             =   2010972901
      DockType        =   0
      PaneVisible     =   -1  'True
      DockStyle       =   0
      CanDockLeft     =   -1  'True
      CanDockTop      =   -1  'True
      CanDockRight    =   -1  'True
      CanDockBottom   =   -1  'True
      AutoHide        =   1
      InitDockHW      =   150
      InitFloatLeft   =   200
      InitFloatTop    =   200
      InitFloatWidth  =   200
      InitFloatHeight =   200
   End
   Begin VB.Label lblDetail 
      BackColor       =   &H002C1412&
      BackStyle       =   0  'Transparent
      Caption         =   "Requires 1 or more Moral Crew: Add 1 to the Range of this Fly Action for each Moral Crew on board. Keep Flying."
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFDD&
      Height          =   1455
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   2430
      Width           =   3765
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "What's going on in the Engine Room?"
      BeginProperty Font 
         Name            =   "Cyberpunk Is Not Dead"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFF690&
      Height          =   1335
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   3675
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public NavCardID, NavOption

Private Sub cmd_Click(Index As Integer)
Dim x
   playsnd 8
   Select Case Index
      Case 0
         NavOption = 1
         actionSeq = ASnavEnd
      Case 1
         NavOption = 2
         actionSeq = ASnavEnd
      Case 2
         NavOption = 3 ' discard
         actionSeq = ASnavEnd
   End Select
   
   For x = 0 To 1
      cmd(x).Enabled = False
   Next x
   
End Sub

Private Sub FDPane1_OnHidden()
   Select Case actionSeq
   Case ASNav, ASselect
      playsnd 9
      FDPane1.PaneVisible = True
   Case ASnavEnd, ASNavEvade, ASNavEvadeEnd, ASNavReav, ASNavReavBorder, ASNavReavEnd, ASNavCrus, ASNavCrusBorder, ASNavCrusOutlaw, ASNavCrusAdjacent, ASNavCorvAdjacent, ASNavCorvPlanetary, ASNavCrusEnd

   Case Else
     ' Beep
   End Select
End Sub
