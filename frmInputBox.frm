VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "InputBox"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInputBox.frx":0000
   ScaleHeight     =   1710
   ScaleWidth      =   6615
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
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtVal 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   435
      Left            =   2723
      TabIndex        =   2
      Text            =   "0"
      Top             =   1110
      Width           =   1185
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
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1110
      Width           =   1635
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
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   6405
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public result As Integer

Private Sub cmd_Click(Index As Integer)
   result = Val(txtVal)
   playsnd 8
   Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub
