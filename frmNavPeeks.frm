VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmNavPeeks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Re-Order the top 5 Nav Cards as you wish"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SftTree.SftTree sftTree 
      DragIcon        =   "frmNavPeeks.frx":0000
      Height          =   4065
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   12315
      _Version        =   262144
      _ExtentX        =   21722
      _ExtentY        =   7170
      _StockProps     =   237
      ForeColor       =   16777215
      BackColor       =   8388669
      BorderStyle     =   1
      ItemPictureExpanded=   "frmNavPeeks.frx":030A
      ItemPictureExpandable=   "frmNavPeeks.frx":0326
      ItemPictureLeaf =   "frmNavPeeks.frx":0342
      PlusMinusPictureExpanded=   "frmNavPeeks.frx":035E
      PlusMinusPictureExpandable=   "frmNavPeeks.frx":037A
      PlusMinusPictureLeaf=   "frmNavPeeks.frx":0396
      ButtonPicture   =   "frmNavPeeks.frx":03B2
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
      ButtonStyle     =   2
      TreeLineColor   =   -2147483632
      Columns         =   2
      ColWidth0       =   167
      ColTitle0       =   "Card Name and Options"
      ColBmp0         =   "frmNavPeeks.frx":03CE
      ColWidth1       =   187
      ColTitle1       =   "Details"
      ColBmp1         =   "frmNavPeeks.frx":03EA
      MouseIcon       =   "frmNavPeeks.frx":0406
      ColHeaderBackColor=   0
      ColHeaderForeColor=   65280
      ForeColor       =   16777215
      BackColor       =   8388669
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmNavPeeks.frx":0422
      DropHighlightStyle=   2
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      ColPict0        =   "frmNavPeeks.frx":043E
      ColFlag1        =   4
      ColPict1        =   "frmNavPeeks.frx":045A
      BackgroundPicture=   "frmNavPeeks.frx":0476
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin VB.CheckBox chkDiscard 
      Caption         =   "Discard All"
      Height          =   225
      Left            =   9510
      TabIndex        =   2
      Top             =   4230
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "OK"
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
      Left            =   11580
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4140
      Width           =   1035
   End
   Begin MSComctlLib.ImageList AssetImages 
      Left            =   30
      Top             =   3870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavPeeks.frx":0492
            Key             =   "UN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavPeeks.frx":0724
            Key             =   "L"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavPeeks.frx":1376
            Key             =   "U"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavPeeks.frx":1FC8
            Key             =   "LN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavPeeks.frx":22E2
            Key             =   "R"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavPeeks.frx":2F34
            Key             =   "D"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavPeeks.frx":3B86
            Key             =   "O"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNavPeeks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public NavZone, minSeq As Integer

Private Sub cmd_Click()
Dim index, cnt
   cnt = minSeq
   With sftTree
      For index = 0 To .ListCount - 1
         If .CellItemData(index, 0) = 1 Then 'header
            If NavZone = "M" Then
               DB.Execute "UPDATE MisbehaveDeck Set Seq =" & IIf(chkDiscard.Value = 1, 5, cnt) & " WHERE CardID =" & .ItemData(index)
            Else
               DB.Execute "UPDATE NavDeck Set Seq =" & cnt & " WHERE CardID =" & .ItemData(index)
            End If
            cnt = cnt + 1
         End If
      Next index
   End With
   playsnd 8
   Me.Hide
End Sub

Private Sub Form_Load()
   With sftTree
      Set .ItemPictureExpandable = AssetImages.Overlay("U", "U")
      Set .ItemPictureExpanded = AssetImages.Overlay("U", "D")
      Set .ItemPictureLeaf = AssetImages.Overlay("LN", "LN")
   
      .LeftButtonOnly = False
      .AutoRespond = True
      .ButtonStyle = buttonsSftTreeAll
      If NavZone = "M" Then
         Me.Caption = "Re-Order the top 3 Misbehave Cards as you wish.   (click, then click-drag and drop)"
         chkDiscard.Visible = True
         RefreshMB
      Else
         RefreshNav
         Me.Caption = "Re-Order the top 5 '" & NavZone & "' zone Nav Cards as you wish.   (click, then click-drag and drop)"
      End If
   
   
   End With
End Sub

Private Sub Form_Resize()
  sftTree.Move sftTree.Left, sftTree.top, Abs(Me.Width - 200), Abs(Me.Height - sftTree.top - 1000)
End Sub


Private Sub sftTree_DragDrop(Source As Control, x As Single, y As Single)
Dim index As Long, CardID
   With sftTree
      
      index = .DropHighlight
      CardID = .ListIndex
      If index = -1 Then Exit Sub 'dropped on original drag
      If index >= .ListCount Then
         .MoveItems CardID, .DependentCount(CardID, 0) + 1, -1
         
      Else
         If .CellItemData(index, 0) = 1 Then
            .MoveItems CardID, .DependentCount(CardID, 0) + 1, index
         End If
      End If
      
      .DropHighlight = -1

   End With
End Sub

Private Sub sftTree_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Dim index As Long
   With sftTree
      index = .HitTest(x, y)
      If index = -1 Then Exit Sub
      .DropHighlightStyle = dropSftTreeBetween  ' = dropSftTreeOnTop
      If State = 1 Then
            ' Leaving this tree control
            .DropHighlight = -1
      Else
            .DropHighlight = index
            
      End If
   End With
End Sub

Private Sub sftTree_DragStarting(ByVal Button As Integer, ByVal Shift As Integer)
   If sftTree.CellItemData(sftTree.ListIndex, 0) = 1 Then  'any title
      'sftTree.DragIcon = DragIcon.Picture
      sftTree.drag 1
   End If
End Sub

Private Sub RefreshMB()
Dim rst As New ADODB.Recordset
Dim SQL, index, cnt As Integer

With sftTree

   .Clear
   
   SQL = "SELECT MisbehaveDeck.CardID, MisbehaveDeck.CardName, MisOption.OptionID, MisOption.OptionName, MisOption.Details, MisOption_1.OptionID AS Option2, "
   SQL = SQL & "MisOption_1.OptionName AS Option2Name, MisOption_1.Details AS Details2, MisbehaveDeck.Seq "
   SQL = SQL & "FROM (MisOption INNER JOIN MisbehaveDeck ON MisOption.OptionID = MisbehaveDeck.Option1ID) LEFT JOIN MisOption AS MisOption_1 ON MisbehaveDeck.Option2ID = MisOption_1.OptionID "
   SQL = SQL & "WHERE MisbehaveDeck.Seq > 6 ORDER BY  MisbehaveDeck.Seq"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      minSeq = rst!Seq
   End If
   Do While Not rst.EOF
      cnt = cnt + 1
      index = .AddItem(rst!CardName)
      .ItemData(index) = rst!CardID
       .CellItemData(index, 0) = 1
       .CellItemData(index, 1) = rst!Seq
      .ItemLevel(index) = 0
      .CellForeColor(index, 0) = 0
      .CellBackColor(index, 0) = 13236739
      
      index = .AddItem(rst!OptionName)
      .ItemData(index) = rst!OptionID
      .CellText(index, 1) = rst!Details
      .CellItemData(index, 0) = 2
      .ItemLevel(index) = 1
      
      If Not IsNull(rst!Option2) Then
         index = .AddItem(rst!Option2Name)
         .ItemData(index) = rst!Option2
         .CellText(index, 1) = rst!Details2
         .CellItemData(index, 0) = 2
         .ItemLevel(index) = 1
      End If
      If cnt > 2 Then Exit Do
      rst.MoveNext
   Loop
   rst.Close
   Set rst = Nothing
   
End With

End Sub

Private Sub RefreshNav()
Dim rst As New ADODB.Recordset
Dim SQL, index, cnt As Integer

With sftTree

   .Clear
   
   SQL = "SELECT NavDeck.CardID, NavDeck.CardName, NavOption.OptionID, NavOption.OptionName, NavOption.Details, NavOption_1.OptionID AS Option2, "
   SQL = SQL & "NavOption_1.OptionName AS Option2Name, NavOption_1.Details AS Details2, NavDeck.Seq "
   SQL = SQL & "FROM (NavOption INNER JOIN NavDeck ON NavOption.OptionID = NavDeck.Option1ID) LEFT JOIN NavOption AS NavOption_1 ON NavDeck.Option2ID = NavOption_1.OptionID "
   SQL = SQL & "WHERE NavDeck.Zones='" & NavZone & "' AND NavDeck.Seq > 6 ORDER BY  NavDeck.Seq"

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      minSeq = rst!Seq
   End If
   Do While Not rst.EOF
      cnt = cnt + 1
      index = .AddItem(rst!CardName)
      .ItemData(index) = rst!CardID
       .CellItemData(index, 0) = 1
       .CellItemData(index, 1) = rst!Seq
      .ItemLevel(index) = 0
      .CellForeColor(index, 0) = 0
      .CellBackColor(index, 0) = 13236739
      
      index = .AddItem(rst!OptionName)
      .ItemData(index) = rst!OptionID
      .CellText(index, 1) = rst!Details
      .CellItemData(index, 0) = 2
      .ItemLevel(index) = 1
      
      If Not IsNull(rst!Option2) Then
         index = .AddItem(rst!Option2Name)
         .ItemData(index) = rst!Option2
         .CellText(index, 1) = rst!Details2
         .CellItemData(index, 0) = 2
         .ItemLevel(index) = 1
      End If
      If cnt > 4 Then Exit Do
      rst.MoveNext
   Loop
   rst.Close
   Set rst = Nothing
   
End With

End Sub
