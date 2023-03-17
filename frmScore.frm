VERSION 5.00
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SftTreeX.ocx"
Begin VB.Form frmScore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scores"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SftTree.SftTree Grid 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4125
      _Version        =   262144
      _ExtentX        =   7276
      _ExtentY        =   6324
      _StockProps     =   237
      ForeColor       =   12648447
      BackColor       =   3355725
      BorderStyle     =   1
      Appearance      =   1
      Appearance      =   1
      ItemPictureExpanded=   "frmScore.frx":0000
      ItemPictureExpandable=   "frmScore.frx":001C
      ItemPictureLeaf =   "frmScore.frx":0038
      PlusMinusPictureExpanded=   "frmScore.frx":0054
      PlusMinusPictureExpandable=   "frmScore.frx":0070
      PlusMinusPictureLeaf=   "frmScore.frx":008C
      ButtonPicture   =   "frmScore.frx":00A8
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
      ColHeaderAppearance=   2
      GridStyle       =   2
      ButtonStyle     =   0
      ItemLines       =   10
      TreeLineStyle   =   0
      Columns         =   4
      ColWidth0       =   67
      ColTitle0       =   "Name"
      ColWidth1       =   33
      ColStyle1       =   9
      ColTitle1       =   "Turns"
      ColBmp1         =   "frmScore.frx":00C4
      ColWidth2       =   53
      ColStyle2       =   9
      ColTitle2       =   "Minutes"
      ColBmp2         =   "frmScore.frx":00E0
      ColWidth3       =   67
      ColTitle3       =   "Date Won"
      ColBmp3         =   "frmScore.frx":00FC
      MouseIcon       =   "frmScore.frx":0118
      ColHeaderBackColor=   -2147483639
      ColHeaderForeColor=   32768
      ForeColor       =   12648447
      BackColor       =   3355725
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmScore.frx":0134
      LeftButtonOnly  =   0   'False
      RowHeaderAppearance=   0
      ColPict1        =   "frmScore.frx":0150
      ColPict2        =   "frmScore.frx":016C
      ColPict3        =   "frmScore.frx":0188
      ItemStyle       =   1
      BackgroundPicture=   "frmScore.frx":01A4
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear Scores"
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StoryID As Integer

Private Sub Form_Load()
Dim rst As New ADODB.Recordset
Dim SQL, Index

With Grid
      .Clear
      SQL = "SELECT PlayerName, Turns, PlayDate,DateDiff('n', StartDate, PlayDate) AS Mins FROM Scores WHERE StoryID=" & StoryID
      SQL = SQL & " ORDER BY Turns, DateDiff('n', StartDate, PlayDate)"
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         Index = .AddItem(rst!PlayerName)
         .CellText(Index, 1) = CStr(rst!Turns)
         .CellText(Index, 2) = rst!Mins
         .CellText(Index, 3) = Format(rst!PlayDate, "DD Mmm YYYY HH:nn")
         rst.MoveNext
      Wend
      rst.Close
      Set rst = Nothing
   
End With
End Sub

Private Sub mnuClear_Click()
      
   DB.Execute "DELETE FROM Scores WHERE StoryID=" & StoryID
   Grid.Clear
   MessBox "Scores Cleared!", "Scores", "OK"
   Me.Hide
End Sub
