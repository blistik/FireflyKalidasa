VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmPopup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Activity Notice"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPopup.frx":0000
   ScaleHeight     =   1830
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1300
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Height          =   1200
      Left            =   6630
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1300
      Width           =   1995
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl picDice 
      Height          =   915
      Index           =   2
      Left            =   1260
      Top             =   2520
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      Effects         =   "frmPopup.frx":CDCAA
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl picDice 
      Height          =   915
      Index           =   1
      Left            =   2300
      Top             =   2430
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      Effects         =   "frmPopup.frx":CDCC2
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl picDice 
      Height          =   915
      Index           =   0
      Left            =   330
      Top             =   2430
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1614
      Effects         =   "frmPopup.frx":CDCDA
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   6405
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public result As Integer

Private Sub cmd_Click(Index As Integer)
   result = Index
   playsnd 8
   Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub
