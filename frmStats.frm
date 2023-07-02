VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Begin VB.Form frmStats 
   Caption         =   "Game Info"
   ClientHeight    =   10365
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   4155
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmStats.frx":030A
   ScaleHeight     =   10365
   ScaleWidth      =   4155
   Begin VB.CommandButton cmdStory 
      BackColor       =   &H00FF8080&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "view Story Details"
      Top             =   420
      Width           =   405
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3690
      Top             =   1800
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Top             =   400
      Width           =   3580
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   825
      Index           =   9
      Left            =   60
      TabIndex        =   12
      Top             =   9225
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   825
      Index           =   8
      Left            =   60
      TabIndex        =   11
      Top             =   8340
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   825
      Index           =   6
      Left            =   60
      TabIndex        =   10
      Top             =   6570
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   825
      Index           =   7
      Left            =   60
      TabIndex        =   9
      Top             =   7455
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   825
      Index           =   5
      Left            =   60
      TabIndex        =   8
      Top             =   5700
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   825
      Index           =   4
      Left            =   60
      TabIndex        =   7
      Top             =   4810
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   825
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   3930
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   825
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   3040
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   820
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   2160
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      Caption         =   "Story"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   465
   End
   Begin VB.Label lblStory 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Height          =   1065
      Left            =   60
      TabIndex        =   1
      Top             =   780
      Width           =   4005
   End
   Begin XDOCKFLOATLibCtl.FDPane FDPane1 
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   8040
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
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub refreshform()
Dim rst As ADODB.Recordset, SQL, x

   If Val(lblTitle.Tag) <> Logic!StoryID Then
      LoadStory Logic!StoryID
      Set rst = New ADODB.Recordset
      SQL = "SELECT * FROM Contact WHERE ContactID > 0  and ContactID < 10 "
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      While Not rst.EOF
         lblContact(rst!ContactID) = rst!ContactName & ":  " & rst!DealDescr & vbNewLine & IIf(rst!ContactID = 5, "Sells Fuel: $100", IIf(rst!cargo = 0, "", "Buys Cargo: $" & rst!cargo & " && Contraband: $" & rst!Contraband))
         If isSolid(player.ID, rst!ContactID) Then
            lblContact(rst!ContactID).BackColor = &HC0FFC0
         Else
            lblContact(rst!ContactID).BackColor = &HCBE1ED
         End If
         rst.MoveNext
      Wend
      rst.Close
   Else
      For x = 1 To NO_OF_CONTACTS
         If isSolid(player.ID, x) Then
            lblContact(x).BackColor = &HC0FFC0
         Else
            lblContact(x).BackColor = &HCBE1ED
         End If
      Next x
   End If
  
   Set rst = Nothing
      
End Sub

Private Sub cmdStory_Click()
   doCustomStory True
End Sub

Private Sub Form_Load()
   refreshform
End Sub

Private Sub Form_Resize()
Dim x
   lblTitle.Width = Abs(Me.Width - 475)
   lblStory.Width = Abs(Me.Width - 50)
   
   For x = 1 To NO_OF_CONTACTS
      lblContact(x).Width = Abs(Me.Width - 50)
   Next x
   cmdStory.Left = Abs(Me.Width - 450)
End Sub

Private Sub Timer1_Timer()
   refreshform
   
End Sub

Private Sub LoadStory(ByVal StoryID)
Dim rst As New ADODB.Recordset
Dim SQL
   SQL = "SELECT * FROM Story "
   SQL = SQL & "WHERE StoryID =" & StoryID
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      lblTitle = rst!StoryTitle
      lblTitle.Tag = CStr(StoryID)
      lblStory.Caption = rst!StoryDesc
   End If
   rst.Close
   Set rst = Nothing

End Sub
