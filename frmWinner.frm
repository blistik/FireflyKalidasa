VERSION 5.00
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmWinner 
   BorderStyle     =   0  'None
   Caption         =   "win"
   ClientHeight    =   9585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "frmWinner.frx":0000
   ScaleHeight     =   9585
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SftTree.SftTree Grid 
      Height          =   3585
      Left            =   270
      TabIndex        =   1
      Top             =   5010
      Visible         =   0   'False
      Width           =   3765
      _Version        =   262144
      _ExtentX        =   6641
      _ExtentY        =   6324
      _StockProps     =   237
      ForeColor       =   12648447
      BackColor       =   3355725
      BorderStyle     =   1
      Appearance      =   1
      Appearance      =   1
      ItemPictureExpanded=   "frmWinner.frx":2541E
      ItemPictureExpandable=   "frmWinner.frx":2543A
      ItemPictureLeaf =   "frmWinner.frx":25456
      PlusMinusPictureExpanded=   "frmWinner.frx":25472
      PlusMinusPictureExpandable=   "frmWinner.frx":2548E
      PlusMinusPictureLeaf=   "frmWinner.frx":254AA
      ButtonPicture   =   "frmWinner.frx":254C6
      BeginProperty ColHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty RowHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ItemEditFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColHeaderAppearance=   0
      GridStyle       =   2
      ButtonStyle     =   0
      ItemLines       =   10
      TreeLineStyle   =   0
      Columns         =   3
      ColWidth0       =   67
      ColTitle0       =   "Name"
      ColWidth1       =   33
      ColStyle1       =   9
      ColTitle1       =   "Turns"
      ColBmp1         =   "frmWinner.frx":254E2
      ColWidth2       =   67
      ColTitle2       =   "Date Played"
      ColBmp2         =   "frmWinner.frx":254FE
      MouseIcon       =   "frmWinner.frx":2551A
      ForeColor       =   12648447
      BackColor       =   3355725
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmWinner.frx":25536
      LeftButtonOnly  =   0   'False
      RowHeaderAppearance=   0
      ColPict1        =   "frmWinner.frx":25552
      ColPict2        =   "frmWinner.frx":2556E
      ItemStyle       =   1
      BackgroundPicture=   "frmWinner.frx":2558A
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin VB.CommandButton cmdScores 
      BackColor       =   &H00FF8080&
      Caption         =   "Show Scores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8790
      Width           =   1365
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "YOU  HAVE  WON!"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4718
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8670
      Width           =   2565
   End
End
Attribute VB_Name = "frmWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
   playsnd 8
   Me.Hide
   
End Sub

Private Sub cmdScores_Click()
Dim rst As New ADODB.Recordset
Dim SQL, index

With Grid
   If .Visible Then
      .Visible = False
      cmdScores.Caption = "Show Scores"
   
   Else
      .Clear
      SQL = "SELECT * FROM Scores WHERE StoryID=" & Logic!StoryID
      SQL = SQL & " ORDER BY Turns, PlayDate DESC"
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         index = .AddItem(rst!PlayerName)
         .CellText(index, 1) = CStr(rst!Turns)
         .CellText(index, 2) = Format(rst!PlayDate, "DD Mmm YYYY HH:nn")
         rst.MoveNext
      Wend
      rst.Close
      Set rst = Nothing
      '.RecalcHorizontalExtent
      '.ScrollBars = scrollSftTreeVert
      .Visible = True
      cmdScores.Caption = "Hide Scores"
   End If
   
End With

End Sub

