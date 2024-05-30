VERSION 5.00
Begin VB.Form frmNavDecks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select a Nav Deck to Reshuffle"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNavDecks.frx":0000
   ScaleHeight     =   3075
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton opt 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Rim Region"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   983
      TabIndex        =   3
      Tag             =   "R"
      Top             =   1650
      Width           =   2595
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Border Region"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   983
      TabIndex        =   2
      Tag             =   "B"
      Top             =   1080
      Width           =   2595
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Alliance Space"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   983
      TabIndex        =   1
      Tag             =   "A"
      Top             =   510
      Value           =   -1  'True
      Width           =   2595
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "select"
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
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2370
      Width           =   795
   End
   Begin VB.Label lblUnseen 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   315
      Index           =   2
      Left            =   3840
      TabIndex        =   7
      Top             =   1710
      Width           =   525
   End
   Begin VB.Label lblUnseen 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   315
      Index           =   1
      Left            =   3840
      TabIndex        =   6
      Top             =   1140
      Width           =   525
   End
   Begin VB.Label lblUnseen 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   315
      Index           =   0
      Left            =   3840
      TabIndex        =   5
      Top             =   570
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unseen"
      Height          =   225
      Left            =   3660
      TabIndex        =   4
      Top             =   300
      Width           =   825
   End
End
Attribute VB_Name = "frmNavDecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public navOpt

Private Sub cmd_Click()
Dim x
   For x = 0 To 2
      If opt(x).Value Then
         navOpt = opt(x).Tag
         Exit For
      End If
   Next x
   playsnd 8
   Me.hide
   
End Sub

Private Sub Form_Load()
Dim x
   For x = 0 To 2
      lblUnseen(x).Caption = getUnseen(opt(x).Tag)
   Next x
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Function getUnseen(ByVal Zone) As Variant
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT Count(CardID) AS cnt "
   SQL = SQL & "FROM NavDeck WHERE Zones ='" & Zone & "' AND Seq > 6"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      getUnseen = rst!cnt
      
   End If
   rst.Close
   Set rst = Nothing
End Function
