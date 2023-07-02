VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmShips 
   Caption         =   "Fireflies"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "frmShips.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin SftTree.SftTree sftTree 
      Height          =   2325
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4485
      _Version        =   262144
      _ExtentX        =   7911
      _ExtentY        =   4101
      _StockProps     =   237
      ForeColor       =   8833235
      BackColor       =   4587520
      BorderStyle     =   1
      ItemPictureExpanded=   "frmShips.frx":030A
      ItemPictureExpandable=   "frmShips.frx":0326
      ItemPictureLeaf =   "frmShips.frx":0342
      PlusMinusPictureExpanded=   "frmShips.frx":035E
      PlusMinusPictureExpandable=   "frmShips.frx":037A
      PlusMinusPictureLeaf=   "frmShips.frx":0396
      ButtonPicture   =   "frmShips.frx":03B2
      BeginProperty ColHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cyberpunk Is Not Dead"
         Size            =   9
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
      Columns         =   10
      ColTitle0       =   "ID"
      ColBmp0         =   "frmShips.frx":03CE
      ColWidth1       =   167
      ColTitle1       =   "Names & Titles"
      ColBmp1         =   "frmShips.frx":03EA
      ColWidth2       =   227
      ColTitle2       =   "Perks & Quirks"
      ColBmp2         =   "frmShips.frx":0406
      ColWidth3       =   67
      ColTitle3       =   "Ability"
      ColBmp3         =   "frmShips.frx":0422
      ColWidth4       =   77
      ColStyle4       =   9
      ColTitle4       =   "Status"
      ColBmp4         =   "frmShips.frx":043E
      ColWidth5       =   33
      ColStyle5       =   9
      ColTitle5       =   "Fight"
      ColBmp5         =   "frmShips.frx":045A
      ColWidth6       =   33
      ColStyle6       =   9
      ColTitle6       =   "Tech"
      ColBmp6         =   "frmShips.frx":0476
      ColWidth7       =   37
      ColStyle7       =   9
      ColTitle7       =   "Nego"
      ColBmp7         =   "frmShips.frx":0492
      ColWidth8       =   47
      ColStyle8       =   10
      ColTitle8       =   "Pay/job"
      ColBmp8         =   "frmShips.frx":04AE
      ColWidth9       =   200
      ColTitle9       =   "Keywords"
      ColBmp9         =   "frmShips.frx":04CA
      MouseIcon       =   "frmShips.frx":04E6
      ColHeaderBackColor=   0
      ColHeaderForeColor=   65280
      ForeColor       =   8833235
      BackColor       =   4587520
      SelectStyle     =   2
      NoFocusStyle    =   2
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmShips.frx":0502
      LeftButtonOnly  =   0   'False
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      OpenEnded       =   0   'False
      ColFlag0        =   4
      ColPict0        =   "frmShips.frx":051E
      ColFlag1        =   8
      ColPict1        =   "frmShips.frx":053A
      ColFlag2        =   4
      ColPict2        =   "frmShips.frx":0556
      ColFlag3        =   12
      ColPict3        =   "frmShips.frx":0572
      ColFlag4        =   8
      ColPict4        =   "frmShips.frx":058E
      ColFlag5        =   8
      ColPict5        =   "frmShips.frx":05AA
      ColFlag6        =   8
      ColPict6        =   "frmShips.frx":05C6
      ColFlag7        =   8
      ColPict7        =   "frmShips.frx":05E2
      ColFlag8        =   8
      ColPict8        =   "frmShips.frx":05FE
      ColFlag9        =   8
      ColPict9        =   "frmShips.frx":061A
      BackgroundPicture=   "frmShips.frx":0636
      CharSearchMode  =   2
      ShowFocusRectangle=   0   'False
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   510
      Top             =   2640
   End
   Begin MSComctlLib.ImageList AssetImages 
      Left            =   2790
      Top             =   2310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":0652
            Key             =   "UN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":08E4
            Key             =   "ST"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":0B76
            Key             =   "NT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":17C8
            Key             =   "haven"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":1B1A
            Key             =   "CS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":236C
            Key             =   "ZS"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":2BBE
            Key             =   "L"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":3810
            Key             =   "U"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":4462
            Key             =   "SG"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":4CB4
            Key             =   "R"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":5906
            Key             =   "D"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":6558
            Key             =   "O"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":71AA
            Key             =   "P"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":7304
            Key             =   "PS"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":761E
            Key             =   "LN"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":7938
            Key             =   "CN"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":7C52
            Key             =   "GR"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":7F6C
            Key             =   "UP"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":83BE
            Key             =   "LD"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":86D8
            Key             =   "SU"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":8B2A
            Key             =   "MA"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":8F7C
            Key             =   "fight"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":92CE
            Key             =   "negot"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":9820
            Key             =   "tech"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":9B72
            Key             =   "dis"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShips.frx":9EC4
            Key             =   "serenity"
         EndProperty
      EndProperty
   End
   Begin VB.Image DragIcon 
      Height          =   480
      Left            =   1560
      Picture         =   "frmShips.frx":A216
      Top             =   2490
      Visible         =   0   'False
      Width           =   480
   End
   Begin XDOCKFLOATLibCtl.FDPane FDPane1 
      Height          =   420
      Left            =   3840
      TabIndex        =   0
      Top             =   2580
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
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "View"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Discard"
         Index           =   1
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Hire Crew"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Trade"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "make a Better Offer"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Off next Job"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "On next Job"
         Index           =   6
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmShips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public shipFilter As String, sftTreeListIndex As Integer

Private Sub Form_Load()
    With sftTree
       Set .ItemPictureExpandable = AssetImages.Overlay("U", "U")
       Set .ItemPictureExpanded = AssetImages.Overlay("U", "D")
       Set .ItemPictureLeaf = AssetImages.Overlay("LN", "LN")
       
       'set the splitter to a scrollbar's width from the right side
       '.SplitterOffset = .Width - 1400  '390.165
      
       .LeftButtonOnly = False
       .AutoRespond = True
       .ButtonStyle = buttonsSftTreeAll

    End With
    loadImages
    
End Sub

Private Sub loadImages()
Dim Index, SQL
Dim rst As New ADODB.Recordset
   SQL = "SELECT Distinct Picture FROM Crew"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If Dir(App.Path & "\Pictures\Sm" & rst!Picture) <> "" Then
         AssetImages.ListImages.Add , Left(rst!Picture, Len(rst!Picture) - 4), LoadPicture(App.Path & "\Pictures\Sm" & rst!Picture)
      End If
      rst.MoveNext
   Wend
   rst.Close
   
End Sub
Private Function findImageKey(ByVal key As String) As Integer
Dim x
   key = Left(key, Len(key) - 4) 'remove .jpg
   With AssetImages
      For x = 1 To .ListImages.Count
         If key = .ListImages.Item(x).key Then
            findImageKey = x
            Exit For
         End If
      Next x
   End With
End Function

Public Sub RefreshShips()
   'keep yours at the top
   refreshShip " WHERE PlayerID = " & player.ID

   If shipFilter = "all" Then
      If useHavenStorage(Logic!StoryID) Then refreshHavens " WHERE PlayerID = " & player.ID, False  '" WHERE Players.Name IS NOT NULL AND Players.AI = 0"
      refreshShip " WHERE Players.Name IS NOT NULL AND PlayerID <> " & player.ID, False
   End If
      
End Sub

Private Sub refreshShip(filter, Optional ByVal doClear As Boolean = True)
Dim Index, SQL, v, w, x, y, z
Dim totalfight, totaltech, totalnego, totalpay, lastplayer, fight As Integer, tech As Integer, nego As Integer
Dim discardF As Boolean, discardT As Boolean, discardN As Boolean
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
    
SQL = "SELECT Board.Zones, P.PlanetName, Players.*"
SQL = SQL & " FROM (Board INNER JOIN Players ON Board.SectorID = Players.SectorID) LEFT JOIN (select Planet.SectorID, min(Planet.PlanetName) AS PlanetName FROM Planet group by Planet.SectorID) P ON Players.SectorID = P.SectorID "
SQL = SQL & filter
SQL = SQL & " ORDER BY PlayerID"
    
'SQL = "SELECT Board.Zones, Planet.PlanetName, Players.* FROM (Board INNER JOIN Players ON Board.SectorID = Players.SectorID) LEFT JOIN Planet ON Players.SectorID = Planet.SectorID "
'SQL = SQL & filter
    
With sftTree

   For Index = 0 To .ListCount - 1
      If .ItemExpand(Index) = False And .DependentCount(Index, 1) > 0 And Index > 2 Then
         z = Index
         Exit For
      End If
   Next Index

   If doClear Then .Clear  'otherwise Append
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      totalfight = 0
      totaltech = 0
      totalnego = 0
      totalpay = 0
      Index = .AddItem(CStr(rst!playerID) & IIf(isOutlaw(rst!playerID), " - outlaw", ""))
      lastplayer = Index
      .CellBackColor(Index, 0) = getPlayerColor(rst!playerID)
      .CellForeColor(Index, 0) = 0
      .ItemLevel(Index) = 0
      .CellText(Index, 1) = rst!ship & " - " & PlayCode(rst!playerID).PlayName ' & IIf(rst!playerID = player.ID, " [me]", "")
      .CellForeColor(Index, 1) = 0
      .CellBackColor(Index, 1) = getPlayerColor(rst!playerID)
      Set .ItemPicture(Index) = AssetImages.ListImages("serenity").Picture
      If Logic!player = rst!playerID Then
         .CellText(Index, 2) = " << IN PLAY >>"
      Else
         .CellText(Index, 2) = "Cash in Hand: $" & rst!pay
      End If
         
      .CellForeColor(Index, 2) = 0
      .CellBackColor(Index, 2) = getPlayerColor(rst!playerID)
      
      .CellText(Index, 3) = "Warrants: " & CStr(rst!Warrants)
      If rst!Warrants > 0 Then
         .CellBackColor(Index, 3) = 3355647
      End If
      If Nz(rst!PlanetName, "Cruiser") = "Cruiser" Or Nz(rst!PlanetName, "Corvette") = "Corvette" Then
         .CellText(Index, 4) = "Sector " & CStr(rst!SectorID)
      Else
         .CellText(Index, 4) = rst!PlanetName
      End If
      .CellItemData(Index, 4) = rst!playerID
      .CellItemData(Index, 6) = rst!SectorID
      If rst!Zones = "B" Then
         .CellBackColor(Index, 4) = 0
      ElseIf rst!Zones = "R" Then
         .CellBackColor(Index, 4) = 79
      Else
         .CellBackColor(Index, 4) = 16711680
      End If
      .CellText(Index, 9) = "Goals: " & CStr(rst!Goals)
      
      'CREW---------------------------------------------
      Index = .AddItem("Crew")

      'Display actual Crew Number and Capacity (6) with modifiers
      x = CrewCapacity(rst!playerID)
      y = getCrewCount(rst!playerID)
      .CellText(Index, 2) = "Crew Cap: " & CStr(x) & " Crew: " & CStr(y) & "  Spare: " & CStr(x - y)
      If getCrewCount(rst!playerID) >= CrewCapacity(rst!playerID) Then
         .CellForeColor(Index, 2) = QBColor(12)
      End If
      .ItemLevel(Index) = 1
      SQL = "SELECT PlayerSupplies.CardID, PlayerSupplies.OffJob, Crew.*, Perk.PerkDescription"
      SQL = SQL & " FROM Perk INNER JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
      SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & rst!playerID
      SQL = SQL & " ORDER BY Crew.Leader DESC, Crew.CrewName"
      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst2.EOF
          Index = .AddItem(CStr(rst2!CrewID))
         .CellItemData(Index, 0) = 1 'crew
         .CellItemData(Index, 1) = rst2!CardID
         .CellItemData(Index, 2) = rst2!CrewID
         .CellItemData(Index, 3) = rst2!leader
         .CellItemData(Index, 4) = rst!playerID
         .CellItemData(Index, 6) = rst!SectorID
         .CellItemData(Index, 7) = rst2!Disgruntled
         .CellItemData(Index, 8) = rst2!pay
         .ItemLevel(Index) = 2
         'set Crew's Avatar
         If rst2!OffJob = 1 Then
            Set .ItemPicture(Index) = AssetImages.Overlay(findImageKey(rst2!Picture), IIf(rst2!leader = 1, "LD", "O"))  '"L"
         ElseIf findImageKey(rst2!Picture) > 0 Then
            'Set .ItemPicture(Index) =  LoadPicture(App.Path & "\Pictures\Sm" & rst2!Picture)
            Set .ItemPicture(Index) = AssetImages.ListImages(findImageKey(rst2!Picture)).Picture
         Else
            Set .ItemPicture(Index) = AssetImages.Overlay("L", IIf(rst2!leader = 1, "LD", "P"))
         End If

         .CellText(Index, 1) = rst2!CrewName & "  -  " & rst2!CrewDescr

         .CellText(Index, 2) = rst2!PerkDescription
         
         .CellText(Index, 3) = Trim(IIf(rst2!Mechanic = 1, "Mechanic  ", "") & IIf(rst2!Pilot = 1, "Pilot  ", "") & IIf(rst2!Companion = 1 Or hasGearCrew(rst!playerID, 36) = rst2!CrewID, "Companion  ", "") & _
               IIf(rst2!Merc = 1, "Merc  ", "") & IIf(rst2!Soldier = 1, "Soldier  ", "") & IIf(rst2!HillFolk = 1, "HillFolk  ", "") & _
               IIf(rst2!Grifter = 1, "Grifter ", "") & IIf(rst2!Medic = 1, "Medic ", "") & IIf(rst2!Mudder = 1, "Mudder", ""))
         .CellForeColor(Index, 3) = 65280

         .CellText(Index, 4) = IIf(rst2!wanted > 0, "Wanted", "") & IIf(rst2!Moral = 1, IIf(rst2!wanted > 0, "/", "") & "Moral ", "")
         .CellForeColor(Index, 4) = 0

         If rst2!wanted > 0 Then
            .CellBackColor(Index, 4) = &HC0C0FF
         ElseIf rst2!Moral = 1 Then
            .CellBackColor(Index, 4) = &HC0FFC0
         End If
         
         'FIGHT
         fight = rst2!fight
         If rst2!HillFolk = 1 Then 'see if there are 3 or more total
            If countCrewAttribute(rst!playerID, "HillFolk") > 2 Then
               fight = fight + 1
               .CellFont(Index, 5).Bold = True
            End If
         End If
         If rst2!CrewID = 76 Then
            If countCrewAttribute(rst!playerID, "Mudder") > 2 Then
               fight = fight + 2
               .CellFont(Index, 5).Bold = True
            End If
         End If
         
         If getPerkAttributeCrew(rst!playerID, "fight", rst2!CardID) > 0 Then
            If hasGearKeyword(rst!playerID, hasPerkKeyword(rst!playerID, rst2!CardID), rst2!CrewID) Then 'crow's special Knife rule
               fight = fight + 1
               .CellFont(Index, 5).Bold = True
            End If
         End If
         .CellText(Index, 5) = IIf(fight > 0, CStr(fight), "")
         .CellForeColor(Index, 5) = 0
         If fight > 0 Then .CellBackColor(Index, 5) = 6052315
         If rst2!OffJob = 0 Then
            totalfight = totalfight + fight
         Else
            .CellFont(Index, 5).Strikethrough = True
         End If
         
         'TECH
         tech = rst2!tech
         If getPerkAttributeCrew(rst!playerID, "tech", rst2!CardID) > 0 Then
            If hasGearKeyword(rst!playerID, hasPerkKeyword(rst!playerID, rst2!CardID), rst2!CrewID) Then 'no one with this rule yet
               tech = tech + 1
               .CellFont(Index, 6).Bold = True
            End If
         End If
         .CellText(Index, 6) = IIf(tech > 0, CStr(tech), "")
         .CellForeColor(Index, 6) = 0
         If tech > 0 Then .CellBackColor(Index, 6) = 16382208
         If rst2!OffJob = 0 Then
            totaltech = totaltech + tech
         Else
            .CellFont(Index, 6).Strikethrough = True
         End If
         
         'NEGOTIATE
         nego = rst2!Negotiate
         x = hasGearCrew(rst!playerID, 28)  'Mal's Brown Coat
         If x = rst2!CrewID Then
            If varDLookup("Disgruntled", "Crew", "CrewID=" & x) > 0 Then
               nego = nego + fight
               .CellFont(Index, 7).Bold = True
            End If
         End If
         'Head Goon
         If countCrewAttribute(rst!playerID, "Merc") > 2 And rst2!CrewID = 65 Then
            nego = nego + 2
            .CellFont(Index, 7).Bold = True
         End If
         If getPerkAttributeCrew(rst!playerID, "negotiate", rst2!CardID) > 0 And hasGearKeyword(rst!playerID, hasPerkKeyword(rst!playerID, rst2!CardID), rst2!CrewID) Then
            nego = nego + 1
            .CellFont(Index, 7).Bold = True
         End If
         
         .CellText(Index, 7) = IIf(nego > 0, CStr(nego), "")
         
         .CellForeColor(Index, 7) = 0
         If Val(.CellText(Index, 7)) > 0 Then .CellBackColor(Index, 7) = 5373777
         If rst2!OffJob = 0 Then
            totalnego = totalnego + Val(.CellText(Index, 7))
         Else
            .CellFont(Index, 7).Strikethrough = True
         End If
         
         .CellText(Index, 8) = IIf(rst2!leader = 1, "Leader ", "$" & CStr(rst2!pay))
         If rst2!leader = 0 Then
            .CellBackColor(Index, 8) = 8388736
            .CellForeColor(Index, 8) = 16777215
         End If
         If rst2!OffJob = 0 Then
            totalpay = totalpay + rst2!pay
         Else
            .CellFont(Index, 8).Strikethrough = True
         End If
         'IIf(rst2!Disgruntled > 0, "Disgruntled ", "") &
         .CellText(Index, 9) = Nz(rst2!KeyWords) & IIf(rst2!Pilot = 1 And hasShipUpgrade(rst!playerID, 10), "TRANSPORT", "")
         .CellForeColor(Index, 9) = 0
         If rst2!Disgruntled > 0 Then
            .CellBackColor(Index, 9) = 8898502 ' 11468799
            Set .CellPicture(Index, 9) = AssetImages.ListImages("dis").Picture
         ElseIf Not IsNull(rst2!KeyWords) Or (rst2!Pilot = 1 And hasShipUpgrade(rst!playerID, 10)) Then
            .CellForeColor(Index, 9) = 65280
         End If
         If rst2!OffJob = 1 Then
            .CellFont(Index, 9).Strikethrough = True
         End If
         
         'Crew's GEAR ---------------------------
         SQL = "SELECT SupplyDeck.CardID, Gear.* FROM Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
         SQL = SQL & "Where PlayerSupplies.CrewID = " & rst2!CrewID
         rst3.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
         While Not rst3.EOF
            Index = .AddItem(CStr(rst3!CardID))
            .CellItemData(Index, 0) = 2 'gear
            .CellItemData(Index, 1) = rst3!CardID
            .CellItemData(Index, 2) = rst3!GearID
            .CellItemData(Index, 4) = rst!playerID
            .CellItemData(Index, 5) = rst2!CrewID
            .ItemLevel(Index) = 3
            If rst2!OffJob = 1 Then
               Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
            ElseIf InStr(rst3!GearName, "Charts") > 0 Or InStr(rst3!GearName, "Contract") > 0 Then
               Set .ItemPicture(Index) = AssetImages.Overlay("L", "MA")
            ElseIf rst3!fight > 0 Then
               Set .ItemPicture(Index) = AssetImages.Overlay("L", "fight")
            ElseIf rst3!tech > 0 Then
               Set .ItemPicture(Index) = AssetImages.Overlay("L", "tech")
            ElseIf rst3!Negotiate > 0 Then
               Set .ItemPicture(Index) = AssetImages.Overlay("L", "negot")
            Else
               Set .ItemPicture(Index) = AssetImages.Overlay("L", "GR")
            End If
            .CellText(Index, 1) = rst3!GearName
            .CellForeColor(Index, 1) = 16685961
            .CellText(Index, 2) = rst3!GearDescr
            .CellForeColor(Index, 2) = 16685961
            .CellForeColor(Index, 3) = 16685961
            '.CellText(Index, 3) =
            '.CellText(index, 4) =
            .CellText(Index, 5) = IIf(rst3!fight > 0, CStr(rst3!fight), "")
            If rst3!discard = 1 And rst3!fight > 0 And rst2!OffJob = 0 Then
               discardF = True
               .CellForeColor(Index, 5) = 65280
            Else
               .CellForeColor(Index, 5) = 0
            End If
            If rst3!fight > 0 Then .CellBackColor(Index, 5) = 6052315
            If rst2!OffJob = 0 Then
               totalfight = totalfight + rst3!fight
            Else
               .CellFont(Index, 5).Strikethrough = True
            End If
                        
            .CellText(Index, 6) = IIf(rst3!tech > 0, CStr(rst3!tech), "")
            If rst3!discard = 1 And rst3!tech > 0 And rst2!OffJob = 0 Then
               discardT = True
               .CellForeColor(Index, 6) = 255
            Else
               .CellForeColor(Index, 6) = 0
            End If
            If rst3!tech > 0 Then .CellBackColor(Index, 6) = 16382208
            If rst2!OffJob = 0 Then
               totaltech = totaltech + rst3!tech
            Else
               .CellFont(Index, 6).Strikethrough = True
            End If
            
            .CellText(Index, 7) = IIf(rst3!Negotiate > 0, CStr(rst3!Negotiate), "")
            If rst3!discard = 1 And rst3!Negotiate > 0 And rst2!OffJob = 0 Then
               discardN = True
               .CellForeColor(Index, 7) = 255
            Else
               .CellForeColor(Index, 7) = 0
            End If
            If rst3!Negotiate > 0 Then .CellBackColor(Index, 7) = 5373777
            If rst2!OffJob = 0 Then
               totalnego = totalnego + rst3!Negotiate
            Else
               .CellFont(Index, 7).Strikethrough = True
            End If
                        
            'Keywords
            .CellText(Index, 9) = Nz(rst3!KeyWords, "")
            .CellForeColor(Index, 9) = 65280
            If rst2!OffJob = 1 Then
               .CellFont(Index, 9).Strikethrough = True
            End If
            rst3.MoveNext
         Wend
         rst3.Close
         rst2.MoveNext
      Wend
      rst2.Close
      'fill the heading totals
      .CellText(lastplayer, 5) = IIf(totalfight > 0, CStr(totalfight), "")
      If discardF Then
         .CellForeColor(lastplayer, 5) = 65280
      Else
         .CellForeColor(lastplayer, 5) = 0
      End If
      If totalfight > 0 Then .CellBackColor(lastplayer, 5) = 6052315
      
      .CellText(lastplayer, 6) = IIf(totaltech > 0, CStr(totaltech), "")
       If discardT Then
         .CellForeColor(lastplayer, 6) = 255
      Else
         .CellForeColor(lastplayer, 6) = 0
      End If
      If totaltech > 0 Then .CellBackColor(lastplayer, 6) = 16382208
      
      .CellText(lastplayer, 7) = IIf(totalnego > 0, CStr(totalnego), "")
      If discardN Then
         .CellForeColor(lastplayer, 7) = 255
      Else
         .CellForeColor(lastplayer, 7) = 0
      End If
      If totalnego > 0 Then .CellBackColor(lastplayer, 7) = 5373777
      
      .CellText(lastplayer, 8) = "$" & CStr(totalpay)
      .CellBackColor(lastplayer, 8) = 8388736
      .CellForeColor(lastplayer, 8) = 16777215
      


       'Unlinked GEAR-----------------------------------
      Index = .AddItem("Gear")
       .CellItemData(Index, 0) = 4 'gear title
       .CellItemData(Index, 4) = rst!playerID
      .ItemLevel(Index) = 1
      SQL = "SELECT SupplyDeck.CardID, Gear.* "
      SQL = SQL & "FROM Gear INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
      SQL = SQL & "WHERE PlayerSupplies.CrewID = 0 AND PlayerSupplies.PlayerID=" & rst!playerID
      
      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst2.EOF
         Index = .AddItem(CStr(rst2!CardID))
         .CellItemData(Index, 0) = 3 'gear unlinked
         .CellItemData(Index, 1) = rst2!CardID
         .CellItemData(Index, 2) = rst2!GearID
         .CellItemData(Index, 4) = rst!playerID
         .ItemLevel(Index) = 2
         If InStr(rst2!GearName, "Charts") > 0 Or InStr(rst2!GearName, "Contract") > 0 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "MA")
         ElseIf rst2!fight > 0 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "fight")
         ElseIf rst2!tech > 0 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "tech")
         ElseIf rst2!Negotiate > 0 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "negot")
         Else
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "GR")
         End If
         .CellText(Index, 1) = rst2!GearName
         .CellForeColor(Index, 1) = 16685961
         .CellText(Index, 2) = rst2!GearDescr
         .CellForeColor(Index, 2) = 16685961
         '.CellText(Index, 3) =
         '.CellText(index, 4) =
         .CellText(Index, 5) = IIf(rst2!fight > 0, CStr(rst2!fight), "")
         .CellForeColor(Index, 5) = 0
         If rst2!fight > 0 Then .CellBackColor(Index, 5) = 6052315
         
         .CellText(Index, 6) = IIf(rst2!tech > 0, CStr(rst2!tech), "")
         .CellForeColor(Index, 6) = 0
         If rst2!tech > 0 Then .CellBackColor(Index, 6) = 16382208
     
         .CellText(Index, 7) = IIf(rst2!Negotiate > 0, CStr(rst2!Negotiate), "")
         .CellForeColor(Index, 7) = 0
         If rst2!Negotiate > 0 Then .CellBackColor(Index, 7) = 5373777
     
         rst2.MoveNext
      Wend
      rst2.Close
       
      'CARGO-----------------------------------
      y = .AddItem("Cargo Hold / Stash")
      .ItemLevel(y) = 1
      
      SQL = "SELECT * FROM Players WHERE PlayerID=" & rst!playerID
      
      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst2.EOF Then
         x = 0
         If rst2!fuel > 0 Then
            x = x + Int(rst2!fuel / 2) + (rst2!fuel Mod 2)
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 6 'fuel
            .CellItemData(Index, 4) = rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "SG")
            .CellText(Index, 1) = "Fuel: " & CStr(rst2!fuel)
            '0=0, 1=1,2=1,3=2,4=2
            
         End If
         If rst2!parts > 0 Then
            x = x + Int(rst2!parts / 2) + (rst2!parts Mod 2)
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 7 'parts
            .CellItemData(Index, 4) = rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "ST")
            .CellText(Index, 1) = "Parts: " & CStr(rst2!parts)
            '0=0, 1=1,2=1,3=2,4=2
            
         End If
         If rst2!cargo > 0 Then
            x = x + rst2!cargo
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 8 'cargo
            .CellItemData(Index, 4) = rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "NT")
            .CellText(Index, 1) = "Cargo: " & CStr(rst2!cargo)
         End If
         If rst2!Passenger > 0 Then
            x = x + rst2!Passenger
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 9 'Passengers
            .CellItemData(Index, 4) = rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "PS")
            .CellText(Index, 1) = "Passenger: " & CStr(rst2!Passenger)
            
         End If
         If rst2!Contraband > 0 Then
            x = x + rst2!Contraband
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 10 'Contraband
            .CellItemData(Index, 4) = rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "CN")
            .CellText(Index, 1) = "Contraband: " & CStr(rst2!Contraband)

         End If
         If rst2!Fugitive > 0 Then
            x = x + rst2!Fugitive
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 11 ' Fugitives
            .CellItemData(Index, 4) = rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "P")
            .CellText(Index, 1) = "Fugitive: " & CStr(rst2!Fugitive)
            
         End If

      End If
      v = StashCapacity(rst!playerID)
      w = CargoCapacity(rst!playerID)
      x = CargoSpaceUsed(rst!playerID)
      .CellText(y, 2) = "Hold Capacity: " & CStr(w - v) & ",  Stash Capacity: " & CStr(v) & ",  Carrying: " & CStr(x) & "  Spare: " & CStr((w - x))
      If (w - CargoSpaceUsed(rst!playerID)) < 1 Then .CellForeColor(y, 2) = QBColor(12)
      
      If z = y Then .Collapse y, True
      rst2.Close
      'SHIP UPDGRADES-----------------------------------
      y = .AddItem("Drive Core & Ship Upgrades")
      .ItemLevel(y) = 1
      SQL = "SELECT PlayerSupplies.CardID, ShipUpgrade.* "
      SQL = SQL & "FROM ShipUpgrade INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
      SQL = SQL & "WHERE PlayerSupplies.PlayerID=" & rst!playerID
      
      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst2.EOF
         Index = .AddItem(CStr(rst2!CardID))
         .CellItemData(Index, 0) = 5 'ship upgds
         .CellItemData(Index, 1) = rst2!CardID
         .CellItemData(Index, 4) = rst!playerID
         .CellText(Index, 1) = rst2!UpgradeName
         .CellForeColor(Index, 1) = 8823762
         .CellText(Index, 2) = IIf(rst2!DriveCore = 1, "DriveCore: ", "") & rst2!UpgradeDescr
         .CellForeColor(Index, 2) = 8823762
         .CellText(Index, 3) = IIf(rst2!burnFuel > 0, "Full Burn Fuel:" & rst2!burnFuel & ", ", "") & IIf(rst2!DriveCore = 1, "BurnRange: " & CStr(rst2!BurnRange + 5) & ", MoseyRange: " & CStr(rst2!MoseyRange), "")
         'Keywords
         .CellText(Index, 9) = IIf(rst2!ShipUpgradeID = 10 And Not hasCrewAttribute(rst!playerID, "Pilot"), "", Nz(rst2!Keyword, ""))
         .CellForeColor(Index, 9) = 65280
         .ItemLevel(Index) = 2
         Set .ItemPicture(Index) = AssetImages.Overlay("LN", IIf(rst2!DriveCore = 1, "SU", "UP"))
         rst2.MoveNext
      Wend
      If z = y Then .Collapse y, True
      rst2.Close
      w = getShipUpgrades(rst!playerID)
      .CellText(y, 2) = "Upgrade Slots Spare: " & (3 - w)
      If w > 2 Then .CellForeColor(y, 2) = QBColor(12)
      '--------------------------------------------------
      rst.MoveNext
   Wend
   
 End With
   
End Sub

Private Sub refreshHavens(filter, Optional ByVal doClear As Boolean = True)
Dim Index, SQL, w, x, y, z
Dim totalfight, totaltech, totalnego, totalpay, lastplayer, fight As Integer, tech As Integer, nego As Integer
Dim discardF As Boolean, discardT As Boolean, discardN As Boolean
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
    
SQL = "SELECT Board.Zones, P.PlanetName, Players.*"
SQL = SQL & " FROM (Board INNER JOIN Players ON Board.Haven = Players.PlayerID)"
SQL = SQL & " LEFT JOIN (select Planet.SectorID, min(Planet.PlanetName) AS PlanetName FROM Planet group by Planet.SectorID) P ON Board.SectorID = P.SectorID"
SQL = SQL & filter
SQL = SQL & " ORDER BY PlayerID"

With sftTree

   For Index = 0 To .ListCount - 1
      If .ItemExpand(Index) = False And .DependentCount(Index, 1) > 0 And Index > 2 Then
         z = Index
         Exit For
      End If
   Next Index

   If doClear Then .Clear  'otherwise Append
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      totalfight = 0
      totaltech = 0
      totalnego = 0
      totalpay = 0
      Index = .AddItem(CStr(rst!playerID) & " - Haven")
      lastplayer = Index
      .CellBackColor(Index, 0) = getPlayerColor(rst!playerID)
      .CellForeColor(Index, 0) = 0
      .ItemLevel(Index) = 0
      .CellText(Index, 1) = rst!ship & " - " & PlayCode(rst!playerID).PlayName & IIf(rst!playerID = player.ID, " [mine]", "")
      .CellForeColor(Index, 1) = 0
      .CellBackColor(Index, 1) = getPlayerColor(rst!playerID)
      Set .ItemPicture(Index) = AssetImages.ListImages("haven").Picture
     .CellText(Index, 2) = "Cash in the Safe: $" & rst!HPay
         
      .CellForeColor(Index, 2) = 0
      .CellBackColor(Index, 2) = getPlayerColor(rst!playerID)
      
      .CellText(Index, 4) = rst!PlanetName
      .CellItemData(Index, 4) = rst!playerID
      .CellItemData(Index, 6) = rst!SectorID
      If rst!Zones = "B" Then
         .CellBackColor(Index, 4) = 0
      ElseIf rst!Zones = "R" Then
         .CellBackColor(Index, 4) = 79
      Else
         .CellBackColor(Index, 4) = 16711680
      End If
      
      'CREW---------------------------------------------
      Index = .AddItem("Crew")

      'Display actual Crew Number and Capacity (6) with modifiers
      .ItemLevel(Index) = 1
      SQL = "SELECT HavenSupplies.CardID, Crew.*, Perk.PerkDescription"
      SQL = SQL & " FROM Perk INNER JOIN (Crew INNER JOIN (HavenSupplies INNER JOIN SupplyDeck ON HavenSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
      SQL = SQL & "WHERE HavenSupplies.PlayerID=" & rst!playerID
      SQL = SQL & " ORDER BY Crew.CrewName"
      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst2.EOF
          Index = .AddItem(CStr(rst2!CrewID))
         .CellItemData(Index, 0) = 1 'crew
         .CellItemData(Index, 1) = rst2!CardID
         .CellItemData(Index, 2) = rst2!CrewID
         .CellItemData(Index, 3) = rst2!leader
         .CellItemData(Index, 4) = 0 ' rst!playerID
         .CellItemData(Index, 6) = rst!SectorID
         .CellItemData(Index, 7) = rst2!Disgruntled
         .CellItemData(Index, 8) = rst2!pay
         .ItemLevel(Index) = 2
         'set Crew's Avatar
         If findImageKey(rst2!Picture) > 0 Then
            'Set .ItemPicture(Index) =  LoadPicture(App.Path & "\Pictures\Sm" & rst2!Picture)
            Set .ItemPicture(Index) = AssetImages.ListImages(findImageKey(rst2!Picture)).Picture
         Else
            Set .ItemPicture(Index) = AssetImages.Overlay("L", IIf(rst2!leader = 1, "LD", "P"))
         End If

         .CellText(Index, 1) = rst2!CrewName & "  -  " & rst2!CrewDescr

         .CellText(Index, 2) = rst2!PerkDescription
         
         .CellText(Index, 3) = Trim(IIf(rst2!Mechanic = 1, "Mechanic  ", "") & IIf(rst2!Pilot = 1, "Pilot  ", "") & IIf(rst2!Companion = 1 Or hasGearCrew(rst!playerID, 36) = rst2!CrewID, "Companion  ", "") & _
               IIf(rst2!Merc = 1, "Merc  ", "") & IIf(rst2!Soldier = 1, "Soldier  ", "") & IIf(rst2!HillFolk = 1, "HillFolk  ", "") & _
               IIf(rst2!Grifter = 1, "Grifter ", "") & IIf(rst2!Medic = 1, "Medic ", "") & IIf(rst2!Mudder = 1, "Mudder", ""))
         .CellForeColor(Index, 3) = 65280

         .CellText(Index, 4) = IIf(rst2!wanted > 0, "Wanted", "") & IIf(rst2!Moral = 1, IIf(rst2!wanted > 0, "/", "") & "Moral ", "")
         .CellForeColor(Index, 4) = 0

         If rst2!wanted > 0 Then
            .CellBackColor(Index, 4) = &HC0C0FF
         ElseIf rst2!Moral = 1 Then
            .CellBackColor(Index, 4) = &HC0FFC0
         End If
         
         'FIGHT
         fight = rst2!fight
         If rst2!HillFolk = 1 Then 'see if there are 3 or more total
            If countCrewAttribute(rst!playerID, "HillFolk") > 2 Then
               fight = fight + 1
               .CellFont(Index, 5).Bold = True
            End If
         End If
         If rst2!CrewID = 76 Then
            If countCrewAttribute(rst!playerID, "Mudder") > 2 Then
               fight = fight + 2
               .CellFont(Index, 5).Bold = True
            End If
         End If
         
         If getPerkAttributeCrew(rst!playerID, "fight", rst2!CardID) > 0 Then
            If hasGearKeyword(rst!playerID, hasPerkKeyword(rst!playerID, rst2!CardID), rst2!CrewID) Then 'crow's special Knife rule
               fight = fight + 1
               .CellFont(Index, 5).Bold = True
            End If
         End If
         .CellText(Index, 5) = IIf(fight > 0, CStr(fight), "")
         .CellForeColor(Index, 5) = 0
         If fight > 0 Then .CellBackColor(Index, 5) = 6052315
         totalfight = totalfight + fight
         
         'TECH
         tech = rst2!tech
         If getPerkAttributeCrew(rst!playerID, "tech", rst2!CardID) > 0 Then
            If hasGearKeyword(rst!playerID, hasPerkKeyword(rst!playerID, rst2!CardID), rst2!CrewID) Then 'no one with this rule yet
               tech = tech + 1
               .CellFont(Index, 6).Bold = True
            End If
         End If
         .CellText(Index, 6) = IIf(tech > 0, CStr(tech), "")
         .CellForeColor(Index, 6) = 0
         If tech > 0 Then .CellBackColor(Index, 6) = 16382208
         totaltech = totaltech + tech
         
         'NEGOTIATE
         nego = rst2!Negotiate
         x = hasGearCrew(rst!playerID, 28)  'Mal's Brown Coat
         If x = rst2!CrewID Then
            If varDLookup("Disgruntled", "Crew", "CrewID=" & x) > 0 Then
               nego = nego + fight
               .CellFont(Index, 7).Bold = True
            End If
         End If
         'Head Goon
         If countCrewAttribute(rst!playerID, "Merc") > 2 And rst2!CrewID = 65 Then
            nego = nego + 2
            .CellFont(Index, 7).Bold = True
         End If
         If getPerkAttributeCrew(rst!playerID, "negotiate", rst2!CardID) > 0 And hasGearKeyword(rst!playerID, hasPerkKeyword(rst!playerID, rst2!CardID), rst2!CrewID) Then
            nego = nego + 1
            .CellFont(Index, 7).Bold = True
         End If
         
         .CellText(Index, 7) = IIf(nego > 0, CStr(nego), "")
         
         .CellForeColor(Index, 7) = 0
         If Val(.CellText(Index, 7)) > 0 Then .CellBackColor(Index, 7) = 5373777
         totalnego = totalnego + Val(.CellText(Index, 7))
         
         .CellText(Index, 8) = IIf(rst2!leader = 1, "Leader ", "$" & CStr(rst2!pay))
         If rst2!leader = 0 Then
            .CellBackColor(Index, 8) = 8388736
            .CellForeColor(Index, 8) = 16777215
         End If
         totalpay = totalpay + rst2!pay
         'IIf(rst2!Disgruntled > 0, "Disgruntled ", "") &
         .CellText(Index, 9) = Nz(rst2!KeyWords) & IIf(rst2!Pilot = 1 And hasShipUpgrade(rst!playerID, 10), "TRANSPORT", "")
         .CellForeColor(Index, 9) = 0
         If rst2!Disgruntled > 0 Then
            .CellBackColor(Index, 9) = 8898502 ' 11468799
            Set .CellPicture(Index, 9) = AssetImages.ListImages("dis").Picture
         ElseIf Not IsNull(rst2!KeyWords) Or (rst2!Pilot = 1 And hasShipUpgrade(rst!playerID, 10)) Then
            .CellForeColor(Index, 9) = 65280
         End If
                 
         rst2.MoveNext
      Wend
      rst2.Close
      'fill the heading totals
      .CellText(lastplayer, 5) = IIf(totalfight > 0, CStr(totalfight), "")
      If discardF Then
         .CellForeColor(lastplayer, 5) = 65280
      Else
         .CellForeColor(lastplayer, 5) = 0
      End If
      If totalfight > 0 Then .CellBackColor(lastplayer, 5) = 6052315
      
      .CellText(lastplayer, 6) = IIf(totaltech > 0, CStr(totaltech), "")
       If discardT Then
         .CellForeColor(lastplayer, 6) = 255
      Else
         .CellForeColor(lastplayer, 6) = 0
      End If
      If totaltech > 0 Then .CellBackColor(lastplayer, 6) = 16382208
      
      .CellText(lastplayer, 7) = IIf(totalnego > 0, CStr(totalnego), "")
      If discardN Then
         .CellForeColor(lastplayer, 7) = 255
      Else
         .CellForeColor(lastplayer, 7) = 0
      End If
      If totalnego > 0 Then .CellBackColor(lastplayer, 7) = 5373777
      
      .CellText(lastplayer, 8) = "$" & CStr(totalpay)
      .CellBackColor(lastplayer, 8) = 8388736
      .CellForeColor(lastplayer, 8) = 16777215
      


       'Unlinked GEAR-----------------------------------
      Index = .AddItem("Gear")
       .CellItemData(Index, 0) = 4 'gear title
       .CellItemData(Index, 4) = 0 ' rst!playerID
      .ItemLevel(Index) = 1
      SQL = "SELECT SupplyDeck.CardID, Gear.* "
      SQL = SQL & "FROM Gear INNER JOIN (HavenSupplies INNER JOIN SupplyDeck ON HavenSupplies.CardID = SupplyDeck.CardID) ON Gear.GearID = SupplyDeck.GearID "
      SQL = SQL & "WHERE HavenSupplies.PlayerID=" & rst!playerID
      
      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst2.EOF
         Index = .AddItem(CStr(rst2!CardID))
         .CellItemData(Index, 0) = 3 'gear unlinked
         .CellItemData(Index, 1) = rst2!CardID
         .CellItemData(Index, 2) = rst2!GearID
         .CellItemData(Index, 4) = 0 ' rst!playerID
         .ItemLevel(Index) = 2
         If InStr(rst2!GearName, "Charts") > 0 Or InStr(rst2!GearName, "Contract") > 0 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "MA")
         ElseIf rst2!fight > 0 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "fight")
         ElseIf rst2!tech > 0 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "tech")
         ElseIf rst2!Negotiate > 0 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "negot")
         Else
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "GR")
         End If
         .CellText(Index, 1) = rst2!GearName
         .CellForeColor(Index, 1) = 16685961
         .CellText(Index, 2) = rst2!GearDescr
         .CellForeColor(Index, 2) = 16685961
         '.CellText(Index, 3) =
         '.CellText(index, 4) =
         .CellText(Index, 5) = IIf(rst2!fight > 0, CStr(rst2!fight), "")
         .CellForeColor(Index, 5) = 0
         If rst2!fight > 0 Then .CellBackColor(Index, 5) = 6052315
         
         .CellText(Index, 6) = IIf(rst2!tech > 0, CStr(rst2!tech), "")
         .CellForeColor(Index, 6) = 0
         If rst2!tech > 0 Then .CellBackColor(Index, 6) = 16382208
     
         .CellText(Index, 7) = IIf(rst2!Negotiate > 0, CStr(rst2!Negotiate), "")
         .CellForeColor(Index, 7) = 0
         If rst2!Negotiate > 0 Then .CellBackColor(Index, 7) = 5373777
     
         rst2.MoveNext
      Wend
      rst2.Close
       
      'CARGO-----------------------------------
      y = .AddItem("Cargo Hold")
      .ItemLevel(y) = 1
      
      SQL = "SELECT * FROM Players WHERE PlayerID=" & rst!playerID
      
      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      If Not rst2.EOF Then
         x = 0
         If rst2!hfuel > 0 Then
            x = x + Int(rst2!hfuel / 2) + (rst2!hfuel Mod 2)
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 6 'fuel
            .CellItemData(Index, 4) = 0 ' rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "SG")
            .CellText(Index, 1) = "Fuel: " & CStr(rst2!hfuel)
            '0=0, 1=1,2=1,3=2,4=2
            
         End If
         If rst2!hparts > 0 Then
            x = x + Int(rst2!hparts / 2) + (rst2!hparts Mod 2)
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 7 'parts
            .CellItemData(Index, 4) = 0  ' rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "ST")
            .CellText(Index, 1) = "Parts: " & CStr(rst2!hparts)
            '0=0, 1=1,2=1,3=2,4=2
            
         End If
         If rst2!hcargo > 0 Then
            x = x + rst2!hcargo
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 8 'cargo
            .CellItemData(Index, 4) = 0  ' rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "NT")
            .CellText(Index, 1) = "Cargo: " & CStr(rst2!hcargo)
         End If
         If rst2!hPassenger > 0 Then
            x = x + rst2!hPassenger
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 9 'Passengers
            .CellItemData(Index, 4) = 0  ' rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "PS")
            .CellText(Index, 1) = "Passenger: " & CStr(rst2!hPassenger)
            
         End If
         If rst2!hContraband > 0 Then
            x = x + rst2!hContraband
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 10 'Contraband
            .CellItemData(Index, 4) = 0  ' rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "CN")
            .CellText(Index, 1) = "Contraband: " & CStr(rst2!hContraband)

         End If
         If rst2!hFugitive > 0 Then
            x = x + rst2!hFugitive
            Index = .AddItem(CStr(x))
            .ItemLevel(Index) = 2
            .CellItemData(Index, 0) = 11 ' Fugitives
            .CellItemData(Index, 4) = 0  ' rst!playerID
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "P")
            .CellText(Index, 1) = "Fugitive: " & CStr(rst2!hFugitive)
            
         End If

      End If

      If z = y Then .Collapse y, True
      rst2.Close
      'SHIP UPDGRADES-----------------------------------
      y = .AddItem("Drive Core & Ship Upgrades")
      .ItemLevel(y) = 1
      SQL = "SELECT HavenSupplies.CardID, ShipUpgrade.* "
      SQL = SQL & "FROM ShipUpgrade INNER JOIN (HavenSupplies INNER JOIN SupplyDeck ON HavenSupplies.CardID = SupplyDeck.CardID) ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
      SQL = SQL & "WHERE HavenSupplies.PlayerID=" & rst!playerID
      
      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst2.EOF
         Index = .AddItem(CStr(rst2!CardID))
         .CellItemData(Index, 0) = 5 'ship upgds
         .CellItemData(Index, 1) = rst2!CardID
         .CellItemData(Index, 4) = 0  ' rst!playerID
         .CellText(Index, 1) = rst2!UpgradeName
         .CellForeColor(Index, 1) = 8823762
         .CellText(Index, 2) = IIf(rst2!DriveCore = 1, "DriveCore: ", "") & rst2!UpgradeDescr
         .CellForeColor(Index, 2) = 8823762
         .CellText(Index, 3) = IIf(rst2!burnFuel > 0, "Full Burn Fuel:" & rst2!burnFuel & ", ", "") & IIf(rst2!DriveCore = 1, "BurnRange: " & CStr(rst2!BurnRange + 5) & ", MoseyRange: " & CStr(rst2!MoseyRange), "")
         .ItemLevel(Index) = 2
         Set .ItemPicture(Index) = AssetImages.Overlay("LN", IIf(rst2!DriveCore = 1, "SU", "UP"))
         rst2.MoveNext
      Wend
      If z = y Then .Collapse y, True
      rst2.Close
      '--------------------------------------------------
      rst.MoveNext
   Wend
   
 End With
   
End Sub

Private Sub Form_Resize()
   sftTree.Move sftTree.Left, sftTree.top, Abs(Me.Width - 100), Abs(Me.Height - sftTree.top)
   'Timer1.Enabled = True

End Sub

Private Sub sftTree_DragDrop(Source As Control, x As Single, y As Single)
Dim Index As Long, CrewID, CardID
   With sftTree
   
      Index = .DropHighlight
      If Index = -1 Then Exit Sub 'dropped on original drag
         
      If .CellItemData(Index, 4) <> .CellItemData(.ListIndex, 4) Then
          'tried to drag on to another user - however could be used for Trade in future
      ElseIf .CellItemData(Index, 0) = 1 And (.CellItemData(.ListIndex, 0) = 3 Or .CellItemData(.ListIndex, 0) = 2) Then 'unlinked/linked Gear dropped on a Crew
         CrewID = .CellItemData(Index, 2)
         CardID = .CellItemData(.ListIndex, 1)
         'if CardID = 21 then Jaynes hat can double up
         doChangeGear player.ID, CrewID, CardID, 1
         
      ElseIf .CellItemData(Index, 0) = 4 And .CellItemData(.ListIndex, 0) = 2 Then 'linked Gear dropped on gear title
         CrewID = .CellItemData(Index, 2)
         CardID = .CellItemData(.ListIndex, 1)
         doChangeGear player.ID, CrewID, CardID, 0
         
      End If
      
      .DropHighlight = -1
      RefreshShips
      If actionSeq > ASidle And actionSeq < ASend And actionSeq <> ASBountySkill Then
         Main.showActions
      Else
         If Not (Main.frmJob Is Nothing) Then
            Main.frmJob.RefreshJobs
         End If
         
      End If
      
      'Timer1.Enabled = True
   End With
End Sub

Private Sub sftTree_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Dim Index As Long
   With sftTree
      Index = .HitTest(x, y)
      If Index = -1 Then Exit Sub
      .DropHighlightStyle = dropSftTreeOnTop
      If State = 1 Then
         ' Leaving this tree control
         .DropHighlight = -1
      Else
         .DropHighlight = Index
      End If
   End With
End Sub

Private Sub sftTree_DragStarting(ByVal Button As Integer, ByVal Shift As Integer)
   If sftTree.CellItemData(sftTree.ListIndex, 0) = 2 Or sftTree.CellItemData(sftTree.ListIndex, 0) = 3 And sftTree.CellItemData(sftTree.ListIndex, 4) = player.ID Then   'any Gear
      sftTree.DragIcon = DragIcon.Picture
      sftTree.Drag 1
      'Timer1.Enabled = False
   End If
End Sub

Private Sub sftTree_ItemClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
On Error GoTo err_handler
   Select Case AreaType
   Case 9
      If Button = 2 Then
         With sftTree
            If .CellItemData(Index, 4) <> player.ID And .CellItemData(Index, 6) = .CellItemData(0, 6) And .CellItemData(Index, 0) = 1 And .CellItemData(Index, 3) = 0 And .CellItemData(Index, 7) > 0 And (CrewCapacity(player.ID) - getCrewCount(player.ID) >= 1) And getMoney(player.ID) >= .CellItemData(Index, 8) And actionSeq = ASselect Then 'poach disgruntled other player
               mnuPopup(6).Visible = False
               mnuPopup(5).Visible = False
               mnuPopup(4).Visible = True 'poach crew
               mnuPopup(3).Visible = False
               mnuPopup(2).Visible = False
               mnuPopup(1).Visible = False
               mnuPopup(0).Visible = False
               sftTreeListIndex = Index
            ElseIf .CellItemData(Index, 4) <> player.ID And .CellItemData(Index, 6) = .CellItemData(0, 6) And .CellItemData(Index, 0) = 0 And actionSeq = ASselect Then 'trade with player
               mnuPopup(6).Visible = False
               mnuPopup(5).Visible = False
               mnuPopup(3).Visible = True  'Trade
               mnuPopup(4).Visible = False
               mnuPopup(2).Visible = False
               mnuPopup(1).Visible = False
               mnuPopup(0).Visible = False
               sftTreeListIndex = Index
            ElseIf .CellItemData(Index, 4) <> player.ID Then 'not yours
               Exit Sub
            Else
               mnuPopup(0).Visible = True 'View
               mnuPopup(3).Visible = False
               mnuPopup(4).Visible = False
               sftTreeListIndex = Index
               'discard check
               If .CellItemData(Index, 0) = 1 And getPlanetID(player.ID) > 0 And .CellItemData(Index, 3) = 0 Then 'crew
                  mnuPopup(1).Visible = True
               ElseIf .CellItemData(Index, 0) = 2 Or .CellItemData(Index, 0) = 3 And actionSeq = ASselect Then  'gear
                  mnuPopup(1).Visible = True
               ElseIf .CellItemData(Index, 0) = 5 And Not isDriveCore(.CellItemData(Index, 1)) Then   'shipUpgrds
                  mnuPopup(1).Visible = True
               ElseIf .CellItemData(Index, 0) = 6 Or .CellItemData(Index, 0) = 7 Or .CellItemData(Index, 0) = 8 Or .CellItemData(Index, 0) = 10 Then  'goods
                  mnuPopup(1).Visible = True
               ElseIf (.CellItemData(Index, 0) = 9 Or .CellItemData(Index, 0) = 11) And getPlanetID(player.ID) > 0 Then  'passengers/fugi
                  mnuPopup(1).Visible = True
               Else
                  mnuPopup(1).Visible = False
               End If
               
               mnuPopup(2).Visible = (hasGearAttribute(player.ID, "LabourContract", .CellItemData(Index, 2)) > 0) And (Not frmAction.buydone) And (actionSeq = ASselect) And .CellItemData(Index, 2) > 0 ' is gear with LabourContract
               mnuPopup(0).Visible = (.CellItemData(Index, 0) < 6 Or mnuPopup(1).Visible = False)
               
               If .CellItemData(Index, 0) = 1 And getLeader() <> .CellItemData(Index, 2) Then
                  mnuPopup(5).Visible = Not isOffJob(.CellItemData(Index, 1)) 'Off Job
                  mnuPopup(6).Visible = Not mnuPopup(5).Visible  'On Job
               Else
                  mnuPopup(6).Visible = False
                  mnuPopup(5).Visible = False
               End If
            End If
                
            If .CellItemData(Index, 0) = 1 Or (mnuPopup(0).Visible And .CellItemData(Index, 0) < 6 And .CellItemData(Index, 0) > 0 And .CellItemData(Index, 0) <> 4) Or mnuPopup(1).Visible Or mnuPopup(3).Visible Or mnuPopup(4).Visible Then PopupMenu mnuPop
            
        End With
      End If
   End Select
   
normal_exit:
   Exit Sub
   
err_handler:
   MsgBox Err.Description
   Resume normal_exit
     
   
End Sub

Private Sub sftTree_ItemDblClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)

   If Button = constSftTreeLeftButton And AreaType = constSftTreeCellText Then
      If sftTree.CellItemData(Index, 0) = 1 Then
         Dim frmCrew As New frmCrewSel
         frmCrew.crewFilter = " WHERE CrewID =" & sftTree.CellItemData(Index, 2)
         frmCrew.AlwaysOnTop = True
         frmCrew.Show
         Set frmCrew = Nothing
      End If
      If sftTree.CellItemData(Index, 0) = 2 Or sftTree.CellItemData(Index, 0) = 3 Then
         Dim frmGear As New frmGearView
         frmGear.gearFilter = " WHERE CardID=" & sftTree.CellItemData(Index, 1)
         frmGear.AlwaysOnTop = True
         frmGear.Show
         Set frmGear = Nothing
      End If
      If sftTree.CellItemData(Index, 0) = 5 Then
         Dim frmUpGrd As New frmShipUpgrdView
         frmUpGrd.gearFilter = " WHERE CardID=" & sftTree.CellItemData(Index, 1)
         frmUpGrd.AlwaysOnTop = True
         frmUpGrd.Show
         Set frmUpGrd = Nothing
      End If
   End If
End Sub

Private Sub mnuPopUp_Click(Index As Integer)
Dim frmCrew As frmCrewSel, x, y, z, frmShUp As frmShipUpgd, frmCrewList As frmCrewLst
Dim frmNavDeck As frmNavDecks, frmNavPeek As frmNavPeeks, status
Dim frmGear As frmGearView, frmTrade As frmTrader

   x = sftTreeListIndex
   sftTreeListIndex = -1
   If x < 1 Then Exit Sub

   With sftTree
      Select Case Index 'menu option
         Case 0 'view
            If .CellItemData(x, 0) = 1 Then
               Set frmCrew = New frmCrewSel
               frmCrew.crewFilter = " WHERE CrewID =" & .CellItemData(x, 2)
               frmCrew.AlwaysOnTop = True
               frmCrew.Show
               Set frmCrew = Nothing
            End If
            If .CellItemData(x, 0) = 2 Or .CellItemData(x, 0) = 3 Then
               Set frmGear = New frmGearView
               frmGear.gearFilter = " WHERE Gear.GearID=" & .CellItemData(x, 2)
               frmGear.AlwaysOnTop = True
               frmGear.Show
               Set frmGear = Nothing
            End If
            If sftTree.CellItemData(x, 0) = 5 Then
               Dim frmUpGrd As New frmShipUpgrdView
               frmUpGrd.gearFilter = " WHERE CardID=" & sftTree.CellItemData(x, 1)
               frmUpGrd.AlwaysOnTop = True
               frmUpGrd.Show
               Set frmUpGrd = Nothing
            End If
         
         Case 1 'DISCARD <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'MsgBox "CardID = " & .CellItemData(x, 1)
            Select Case .CellItemData(x, 0) ' 1-crew 2-linked gear, 3-unlinked gear, 5-ship upgd
            Case 1 'CREW
               status = DISCARDED
               If .CellItemData(x, 2) = 68 Then 'buy ShipUpGrd at half price
                  If frmAction.buydone Then
                     MessBox "Must have a Buy Action available!", "Buy Action", "Ooops", "", getLeader()
                     Exit Sub
                  End If
                  'present list of upgrades to buy one at half price
                  Set frmShUp = New frmShipUpgd
                  frmShUp.discardMode = 2
                  frmShUp.Show 1
                  If frmShUp.CardID = 0 Then Exit Sub ' didn't happen
                  PutMsg player.PlayName & " sent The Salesman off to buy a nice Ship Upgrade", player.ID, Logic!Gamecntr

               'Remove from Play at Harvest, Red Sun to take $500. This counts as Immoral.
               ElseIf getPerkAttributeCrew(player.ID, "Indentured", .CellItemData(x, 1)) > 0 And getPlayerSector(player.ID) = 16 Then
                  getMoney player.ID, 500
                  doDisgruntled player.ID, 1
                  PutMsg player.PlayName & " indentured a Mudder to Higgins for $500" & IIf(hasCrewAttribute(player.ID, "Moral"), " and some of the Crew aren't impressed", ""), player.ID, Logic!Gamecntr, True, 0, 0, 0, 8
                  status = 0
               
               Else
                  PutMsg player.PlayName & " decided to part company with " & getCrewName(sftTree.CellItemData(x, 1)), player.ID, Logic!Gamecntr
                  
               End If
               
               doDiscardCrew .CellItemData(x, 1), status
               
               
               RefreshShips
            Case 2, 3  '2-linked gear, 3-unlinked gear
               Select Case .CellItemData(x, 2)
               Case 30  'Maque Tiles
                  If RollDice(6) > 4 Then '1:3 chance
                     getMoney player.ID, 1200
                     
                     PutMsg player.PlayName & " gambled with Maque Tiles and won $1200", player.ID, Logic!Gamecntr, True, 0, 30
                     frmAction.lblMoney.Caption = "$" & getMoney(player.ID)
                  Else
                     PutMsg player.PlayName & " gambled with Maque Tiles and had no luck", player.ID, Logic!Gamecntr, True, getLeader()
                  End If
               Case 24  'eating a fruity bar
                  If .CellItemData(x, 0) = 2 Then 'must be on a crew
                     DB.Execute "UPDATE Crew SET Disgruntled=0 WHERE CrewID=" & .CellItemData(x, 5)
                  End If
                  
               Case 25
                  Set frmNavPeek = New frmNavPeeks
                  frmNavPeek.NavZone = "M"
                  frmNavPeek.Show 1
                  PutMsg player.PlayName & " used the " & .CellText(x, 1) & " to fiddle with the Misbehave deck", player.ID, Logic!Gamecntr
                  
               Case 34 'Billiards Betting -  roll two dice. Take $100 times the total of the dice and discard this card.
                  y = RollDice(6) + RollDice(6)
                  getMoney player.ID, (y * 100)
                  
                  PutMsg player.PlayName & " gambled with Billiards Betting, rolled " & y & " and won $" & CStr(y * 100), player.ID, Logic!Gamecntr, True, 0, 34
                  frmAction.lblMoney.Caption = "$" & getMoney(player.ID)
                  
               Case 43 'Wash's Hawaiian Shirt
                  Set frmNavDeck = New frmNavDecks
                  frmNavDeck.Show 1
                  ShuffleDeck "Nav", True, False, frmNavDeck.navOpt
                  PutMsg player.PlayName & " used Wash's Hawaiian Shirt to reshuffle the " & frmNavDeck.navOpt & " deck", player.ID, Logic!Gamecntr
               
               Case 44 'Discard to look at the top 5 cards of ANY Nav Deck
                  Set frmNavDeck = New frmNavDecks
                  frmNavDeck.Caption = "Pick a Nav Deck to peek"
                  frmNavDeck.Show 1
                  Set frmNavPeek = New frmNavPeeks
                  frmNavPeek.NavZone = frmNavDeck.navOpt
                  frmNavPeek.Show 1
                  PutMsg player.PlayName & " used Wash's Nav Charts to fiddle with the " & frmNavPeek.NavZone & " deck", player.ID, Logic!Gamecntr
                  
               Case 50, 51, 52 'Nav Chart
                  Set frmNavPeek = New frmNavPeeks
                  frmNavPeek.NavZone = IIf(.CellItemData(x, 2) = 50, "A", IIf(.CellItemData(x, 2) = 51, "B", "R"))
                  frmNavPeek.Show 1
                  PutMsg player.PlayName & " used the " & .CellText(x, 1) & " to fiddle with the " & frmNavPeek.NavZone & " deck", player.ID, Logic!Gamecntr
               
               Case 53 'wormhole
                  Main.drawLine 2, 133, 104
                  wormHoleOpen = True
                  PutMsg player.PlayName & " used the " & .CellText(x, 1), player.ID, Logic!Gamecntr
               
               Case 57 '"The Best in the House"
                  doDisgruntled player.ID, -1
                  PutMsg player.PlayName & " shouted the crew " & .CellText(x, 1) & ". Everyone's happy again!", player.ID, Logic!Gamecntr
               End Select
               
               doDiscardGear player.ID, .CellItemData(x, 1)
               RefreshShips
            Case 5   '5-ship upgd
               If canRemoveUpgrade(player.ID, .CellItemData(x, 1)) Then
                  doDiscardGear player.ID, .CellItemData(x, 1)
               End If
               
            Case 6 ' fuel
               y = varDLookup("Fuel", "Players", "PlayerID=" & player.ID)
               Do
                  z = InputBoxx("How much Fuel do you want to toss overboard?", "Make room in the Cargo Hold", "1", getLeader())
                  If z >= 0 And z <= y Then
                     Exit Do
                  End If
               Loop
               DB.Execute "UPDATE Players SET Fuel = Fuel - " & z & " WHERE PlayerID=" & player.ID
               
            Case 7 'parts
               y = varDLookup("Parts", "Players", "PlayerID=" & player.ID)
               Do
                  z = InputBoxx("How many Parts do you want to toss overboard?", "Make room in the Cargo Hold", CStr(y), getLeader())
                  If z >= 0 And z <= y Then
                     Exit Do
                  End If
               Loop
               DB.Execute "UPDATE Players SET Parts = Parts - " & z & " WHERE PlayerID=" & player.ID
            
            Case 8 'Cargo
               y = varDLookup("Cargo", "Players", "PlayerID=" & player.ID)
               Do
                  z = InputBoxx("How much Cargo do you want to toss overboard?", "Make room in the Cargo Hold", CStr(y), getLeader())
                  If z >= 0 And z <= y Then
                     Exit Do
                  End If
               Loop
               DB.Execute "UPDATE Players SET Cargo = Cargo - " & z & " WHERE PlayerID=" & player.ID
            
            Case 10 'Contraband
               y = varDLookup("Contraband", "Players", "PlayerID=" & player.ID)
               Do
                  z = InputBoxx("How much Contraband do you want to toss overboard?", "Make room in the Cargo Hold", CStr(y), getLeader())
                  If z >= 0 And z <= y Then
                     Exit Do
                  End If
               Loop
               DB.Execute "UPDATE Players SET Contraband = Contraband - " & z & " WHERE PlayerID=" & player.ID
                                          
            
            Case 9 'Passenger
               y = varDLookup("Passenger", "Players", "PlayerID=" & player.ID)
               Do
                  z = InputBoxx("How many Passengers do you want to set ashore?", "Make room in the Cargo Hold", CStr(y), getLeader())
                  If z >= 0 And z <= y Then
                     Exit Do
                  End If
               Loop
               DB.Execute "UPDATE Players SET Passenger = Passenger - " & z & " WHERE PlayerID=" & player.ID
             
            Case 11 'Fugitive
               y = varDLookup("Fugitive", "Players", "PlayerID=" & player.ID)
               Do
                  z = InputBoxx("How many Fugitives do you want to set ashore?", "Make room in the Cargo Hold", CStr(y), getLeader())
                  If z >= 0 And z <= y Then
                     Exit Do
                  End If
               Loop
               DB.Execute "UPDATE Players SET Fugitive = Fugitive - " & z & " WHERE PlayerID=" & player.ID
            End Select
            RefreshShips
         
         Case 2 'Labour Contract
            If getCrewCount(player.ID) < CrewCapacity(player.ID) Then
               Set frmCrewList = New frmCrewLst
               frmCrewList.selectCrew = -1
               frmCrewList.SupplyID = hasGearAttribute(player.ID, "LabourContract", .CellItemData(x, 2))
               frmCrewList.Caption = "Select 1 " & getSupplyName(frmCrewList.SupplyID) & " Crew from Discards"
               frmCrewList.Show 1
               If frmCrewList.crewcnt = 1 Then
                  DB.Execute "UPDATE Players SET Pay = Pay - " & frmCrewList.payTotal & " WHERE PlayerID = " & player.ID
                  frmAction.buyIsDone
                  PutMsg player.PlayName & " used a Labour Contract to hire " & getCrewName(0, frmCrewList.lastCrewID), player.ID, Logic!Gamecntr, True, frmCrewList.lastCrewID
               End If
            Else
               MessBox "No Crewspace available!", "No room", "Ooops", "", 0, 0, 2
            End If
         
         Case 3 'Trade
            PutMsg player.PlayName & " requesting to Trade with " & PlayCode(.CellItemData(x, 4)).PlayName, player.ID, Logic!Gamecntr
   
            Logic.Requery
            Logic!Seq = "T"
            Logic!HostAccept = 0
            Logic!ClientAccept = 0
            Logic!trader = .CellItemData(x, 4)
            Logic.Update
            Set frmTrade = New frmTrader
            frmTrade.isHost = True
            frmTrade.TraderID = .CellItemData(x, 4)
            frmTrade.lblTitle(1).Caption = PlayCode(.CellItemData(x, 4)).PlayName & "'s Trade Items"
            frmTrade.Show 1
            Logic.Requery
            Logic.Update "Seq", "R"
            PutMsg player.PlayName & " ended Trading with " & PlayCode(.CellItemData(x, 4)).PlayName, player.ID, Logic!Gamecntr
            RefreshShips
            Main.showActions
         
         Case 4 'poach
            'check can afford Crew pay
            If getMoney(player.ID) < .CellItemData(x, 8) Then
               MessBox "You cannot afford " & getCrewName(.CellItemData(x, 1)), "Money Problems", "Ooops", "", .CellItemData(x, 2)
               Exit Sub
            End If
            
            'update their pile status
            DB.Execute "UPDATE SupplyDeck SET Seq =" & player.ID & " WHERE CardID = " & .CellItemData(x, 1)
            'remove any Gear first
            DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & .CellItemData(x, 2)
            'delete the card to the players deck
            DB.Execute "UPDATE PlayerSupplies SET PlayerID = " & player.ID & " WHERE CardID = " & .CellItemData(x, 1)
            'remove Disgruntled
            DB.Execute "UPDATE Crew SET Disgruntled = 0 WHERE CrewID = " & .CellItemData(x, 2)
            getMoney player.ID, .CellItemData(x, 8) * -1
            PutMsg player.PlayName & " gave disgruntled " & getCrewName(.CellItemData(x, 1)) & " a BETTER OFFER and poached them", player.ID, Logic!Gamecntr, False, 0, 0, 0, 0, 1
            
            RefreshShips
            Main.showActions
         Case 5, 6 'Off / On next Job
            DB.Execute "UPDATE PlayerSupplies SET OffJob =" & IIf(Index = 5, "1", "0") & " WHERE CardID = " & .CellItemData(x, 1)
            RefreshShips
            Main.showActions
         End Select
   End With
   If frmAction.FDPane1.PaneVisible Then Main.showActions
End Sub

Public Sub refreshFuel(ByVal playerID As Integer)
Dim Index
With sftTree

   For Index = 0 To .ListCount - 1
      If .CellItemData(Index, 0) = 6 And .CellItemData(Index, 4) = playerID Then
        .CellText(Index, 1) = "Fuel: " & CStr(Nz(varDLookup("Fuel", "Players", "PlayerID = " & playerID), 0))
         Exit For
      End If
   Next Index
   
   End With
End Sub


Private Sub Timer1_Timer()
   RefreshShips
End Sub

