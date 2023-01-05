VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Begin VB.Form frmNav 
   BackColor       =   &H00400000&
   Caption         =   "Nav Deck"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   Icon            =   "frmNav.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmNav.frx":030A
   ScaleHeight     =   4140
   ScaleWidth      =   4335
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "option 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2340
      Width           =   3585
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "option 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblUnseen 
      BackStyle       =   0  'Transparent
      Caption         =   "unseen"
      ForeColor       =   &H0000C0C0&
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   3570
      Width           =   1815
   End
   Begin VB.Label lblDetail 
      BackColor       =   &H00800000&
      Caption         =   "Keep Flying"
      ForeColor       =   &H0000FFFF&
      Height          =   760
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3585
      WordWrap        =   -1  'True
   End
   Begin XDOCKFLOATLibCtl.FDPane FDPane1 
      Height          =   420
      Left            =   4680
      TabIndex        =   3
      Top             =   3480
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
      BackColor       =   &H00800000&
      Caption         =   "Keep Flying"
      ForeColor       =   &H0000FFFF&
      Height          =   765
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   3585
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   885
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   3825
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
