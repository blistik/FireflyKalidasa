VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmDiscardGear 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Discard the single use item"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SftTree.SftTree sftTree 
      Height          =   2205
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   11445
      _Version        =   262144
      _ExtentX        =   20188
      _ExtentY        =   3889
      _StockProps     =   237
      ForeColor       =   16777215
      BackColor       =   8388669
      BorderStyle     =   1
      ItemPictureExpanded=   "frmDiscardGear.frx":0000
      ItemPictureExpandable=   "frmDiscardGear.frx":001C
      ItemPictureLeaf =   "frmDiscardGear.frx":0038
      PlusMinusPictureExpanded=   "frmDiscardGear.frx":0054
      PlusMinusPictureExpandable=   "frmDiscardGear.frx":0070
      PlusMinusPictureLeaf=   "frmDiscardGear.frx":008C
      ButtonPicture   =   "frmDiscardGear.frx":00A8
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
      Columns         =   9
      ColTitle0       =   "ShipID"
      ColBmp0         =   "frmDiscardGear.frx":00C4
      ColWidth1       =   133
      ColTitle1       =   "Ship Name"
      ColBmp1         =   "frmDiscardGear.frx":00E0
      ColWidth2       =   253
      ColTitle2       =   "Functions"
      ColBmp2         =   "frmDiscardGear.frx":00FC
      ColWidth3       =   67
      ColTitle3       =   "-"
      ColBmp3         =   "frmDiscardGear.frx":0118
      ColWidth4       =   53
      ColStyle4       =   9
      ColTitle4       =   "-"
      ColBmp4         =   "frmDiscardGear.frx":0134
      ColWidth5       =   33
      ColStyle5       =   9
      ColTitle5       =   "Fight"
      ColBmp5         =   "frmDiscardGear.frx":0150
      ColWidth6       =   33
      ColStyle6       =   9
      ColTitle6       =   "Tech"
      ColBmp6         =   "frmDiscardGear.frx":016C
      ColWidth7       =   37
      ColStyle7       =   9
      ColTitle7       =   "Nego"
      ColBmp7         =   "frmDiscardGear.frx":0188
      ColWidth8       =   47
      ColTitle8       =   "Status"
      ColBmp8         =   "frmDiscardGear.frx":01A4
      MouseIcon       =   "frmDiscardGear.frx":01C0
      ColHeaderBackColor=   0
      ColHeaderForeColor=   65280
      ForeColor       =   16777215
      BackColor       =   8388669
      SelectStyle     =   2
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmDiscardGear.frx":01DC
      LeftButtonOnly  =   0   'False
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      ColPict0        =   "frmDiscardGear.frx":01F8
      ColPict1        =   "frmDiscardGear.frx":0214
      ColFlag2        =   4
      ColPict2        =   "frmDiscardGear.frx":0230
      ColFlag3        =   12
      ColPict3        =   "frmDiscardGear.frx":024C
      ColFlag4        =   8
      ColPict4        =   "frmDiscardGear.frx":0268
      ColFlag5        =   8
      ColPict5        =   "frmDiscardGear.frx":0284
      ColFlag6        =   8
      ColPict6        =   "frmDiscardGear.frx":02A0
      ColFlag7        =   8
      ColPict7        =   "frmDiscardGear.frx":02BC
      ColFlag8        =   8
      ColPict8        =   "frmDiscardGear.frx":02D8
      BackgroundPicture=   "frmDiscardGear.frx":02F4
      ShowFocusRectangle=   0   'False
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
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
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2340
      Width           =   1035
   End
   Begin MSComctlLib.ImageList AssetImages 
      Left            =   3030
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":0310
            Key             =   "UN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":05A2
            Key             =   "ST"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":0834
            Key             =   "NT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":1486
            Key             =   "CS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":1CD8
            Key             =   "ZS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":252A
            Key             =   "L"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":317C
            Key             =   "U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":3DCE
            Key             =   "SG"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":4620
            Key             =   "R"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":5272
            Key             =   "D"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":5EC4
            Key             =   "O"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":6B16
            Key             =   "P"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":6C70
            Key             =   "PS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":6F8A
            Key             =   "LN"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":72A4
            Key             =   "CN"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":75BE
            Key             =   "GR"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":78D8
            Key             =   "UP"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiscardGear.frx":7D2A
            Key             =   "LD"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDiscardGear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nbrSelect, nbrSelected, skill, kosher As Boolean

Private Sub cmd_Click()
   playsnd 8
   If nbrSelected >= nbrSelect Then
      'discard selected gear
      discardGear
      Me.hide
   Else
      If MessBox("Not enough Skill Points selected, did you change your mind?", "Single Use Gear", "Yes", "No", getLeader()) = 0 Then
      'If MsgBox("Not enough Skill Points selected, did you change your mind?", vbYesNo + vbQuestion, "Single Use Gear") = vbYes Then
         Me.hide
      End If
   End If

End Sub

Private Sub Form_Load()
  Set Me.Picture = LoadPicture(App.Path & "\pictures\showdown.jpg")
    With sftTree
       Set .ItemPictureExpandable = AssetImages.Overlay("U", "U")
       Set .ItemPictureExpanded = AssetImages.Overlay("U", "D")
       Set .ItemPictureLeaf = AssetImages.Overlay("LN", "LN")
       
       'set the splitter to a scrollbar's width from the right side
       '.SplitterOffset = .Width - 1400  '390.165
      
       .LeftButtonOnly = False
       .AutoRespond = True
       .ButtonStyle = buttonsSftTreeAll
       
       nbrSelected = 0
       RefreshList
       updateSkill
       Me.Caption = "Select single use Gear to provide at least " & CStr(nbrSelect) & " " & skill & " skill points"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub sftTree_ItemClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)

With sftTree

  If Button = constSftTreeLeftButton And (AreaType = constSftTreeItem Or AreaType = constSftTreeCellText) Then
         Select Case .ItemDataString(Index)
         Case "R"  'no Skill

            .ItemDataString(Index) = "O"
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
            updateSkill
            
         Case "O"  'use Skill

            .ItemDataString(Index) = "R"
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "R")
            updateSkill
            
         End Select
      
   End If
   
End With

End Sub
Private Sub updateSkill()
Dim Index
nbrSelected = 0
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = "R" Then
            nbrSelected = nbrSelected + .CellItemData(Index, 5)
         End If
      Next Index
      .CellText(0, 8) = "Remaining=" & CStr(nbrSelect - nbrSelected)
   End With
End Sub

Private Sub discardGear()
Dim Index
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = "R" Then
            doDiscardGear player.ID, .CellItemData(Index, 1)
         End If
      Next Index

   End With
End Sub

Public Sub RefreshList()
Dim rst As New ADODB.Recordset
Dim SQL, Index
   With sftTree
      .Clear
'      SQL = "SELECT SupplyDeck.CardID, Gear.* "
'      SQL = SQL & "FROM Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
'      SQL = SQL & "WHERE PlayerSupplies.CrewID > 0 AND PlayerSupplies.PlayerID=" & player.ID & " AND Gear.Discard=1 and Gear." & skill & " > 0"
      
      SQL = "SELECT SupplyDeck.CardID, Gear.* "
      SQL = SQL & "FROM ((Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID) "
      SQL = SQL & "INNER JOIN SupplyDeck AS SupplyDeck_1 ON PlayerSupplies.CrewID = SupplyDeck_1.CrewID) INNER JOIN PlayerSupplies AS PlayerSupplies_1 ON SupplyDeck_1.CardID = PlayerSupplies_1.CardID "
      SQL = SQL & "Where PlayerSupplies.playerID = " & player.ID & " And Gear.discard = 1 And Gear." & skill & " > 0 And PlayerSupplies_1.OffJob = 0"
      
      If kosher Then
         SQL = SQL & " AND (PlayerSupplies.CrewID = 60 or Gear.GearID = 59)" ' Lund & Inaras knife
      End If
      
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         Index = .AddItem(CStr(rst!CardID))
         .CellItemData(Index, 0) = 3 'gear unlinked
         .CellItemData(Index, 1) = rst!CardID
         .CellItemData(Index, 4) = player.ID
         .ItemDataString(Index) = "O"
         Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
         
         .CellText(Index, 1) = rst!GearName
         .CellText(Index, 2) = rst!GearDescr
         '.CellText(Index, 3) =
         '.CellText(index, 4) =
         .CellItemData(Index, 5) = rst.Fields(skill)
         .CellText(Index, 5) = IIf(rst!fight > 0, CStr(rst!fight), "")
         .CellForeColor(Index, 5) = 65280
         If rst!fight > 0 Then .CellBackColor(Index, 5) = 6052315
         
         .CellText(Index, 6) = IIf(rst!tech > 0, CStr(rst!tech), "")
         .CellForeColor(Index, 6) = 255
         If rst!tech > 0 Then .CellBackColor(Index, 6) = 16382208
     
         .CellText(Index, 7) = IIf(rst!Negotiate > 0, CStr(rst!Negotiate), "")
         .CellForeColor(Index, 7) = 255
         If rst!Negotiate > 0 Then .CellBackColor(Index, 7) = 5373777
     
         rst.MoveNext
      Wend
      rst.Close
   End With
End Sub
