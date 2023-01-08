VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmGamble 
   BackColor       =   &H00CBE1ED&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pick a Suit"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LaVolpeAlphaImg.AlphaImgCtl Alpha 
      Height          =   645
      Index           =   4
      Left            =   1110
      Top             =   870
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1138
      Effects         =   "frmGamble.frx":0000
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Alpha 
      Height          =   645
      Index           =   2
      Left            =   360
      Top             =   870
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1138
      Effects         =   "frmGamble.frx":0018
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Alpha 
      Height          =   645
      Index           =   3
      Left            =   1110
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1138
      Effects         =   "frmGamble.frx":0030
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Alpha 
      Height          =   645
      Index           =   1
      Left            =   360
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1138
      Effects         =   "frmGamble.frx":0048
   End
End
Attribute VB_Name = "frmGamble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mySuit As Integer

Private Sub Alpha_Click(Index As Integer)
   mySuit = Index
   Me.Hide
End Sub

Private Sub Form_Load()
Dim suit As Integer

   For suit = 1 To 4
      Alpha(suit).Picture = LoadPictureGDIplus(App.Path & "\Pictures\suit" & suit & ".bmp")
      Alpha(suit).Visible = True
      Alpha(suit).TransparentColor = &HFFFFFF
      Alpha(suit).TransparentColorMode = lvicUseTransparentColor
   Next suit

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
      MsgBox "Click on a Suit", vbExclamation
   End If
End Sub
