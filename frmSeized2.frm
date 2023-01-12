VERSION 5.00
Begin VB.Form frmSeized2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select 2 crew to hide in the Smuggling Compartments"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSeized2.frx":0000
   ScaleHeight     =   5115
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   2265
      Left            =   0
      Picture         =   "frmSeized2.frx":D3822
      ScaleHeight     =   2205
      ScaleWidth      =   5355
      TabIndex        =   2
      Top             =   0
      Width           =   5415
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Select"
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
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4710
      Width           =   1035
   End
   Begin VB.ListBox crewList 
      BackColor       =   &H00CBE1ED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   2520
      Width           =   5415
   End
End
Attribute VB_Name = "frmSeized2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fullstash

Private Sub cmd_Click()
Dim x, cnt
   cnt = 0
   playsnd 8
   For x = 0 To crewList.ListCount - 1
      If crewList.selected(x) Then
         cnt = cnt + 1
      End If
   Next x
   If cnt <> fullstash Then
      MessBox "You need to select " & fullstash & " crew to go into the Concealed Smuggling Compartments", "Select Crew", "Ooops", "", 0, 0, 11
   Else
      For x = 0 To crewList.ListCount - 1
         If crewList.selected(x) Then
            PutMsg player.PlayName & "'s Crew member " & crewList.List(x) & " hides in the Concealed Smuggling Compartments", player.ID, Logic!Gamecntr
         Else
            doSeizeCrew player.ID, crewList.ItemData(x), varDLookup("Wanted", "Crew", "CrewID=" & getCrewID(crewList.ItemData(x)))
         End If
      Next x
      Me.Hide
   End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub
