VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Activity Notice"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPopup.frx":0000
   ScaleHeight     =   1710
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   2370
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1110
      Width           =   1905
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   800
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

Private Sub cmd_Click(index As Integer)
   playsnd 8
   Me.Hide
End Sub
