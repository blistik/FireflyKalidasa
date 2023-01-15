VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmSupply 
   Caption         =   "Supplies"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5025
   Icon            =   "frmSupply.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   5025
   Begin SftTree.SftTree sftTree 
      Height          =   2325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4485
      _Version        =   262144
      _ExtentX        =   7911
      _ExtentY        =   4101
      _StockProps     =   237
      ForeColor       =   8833235
      BackColor       =   7360778
      BorderStyle     =   1
      ItemPictureExpanded=   "frmSupply.frx":030A
      ItemPictureExpandable=   "frmSupply.frx":0326
      ItemPictureLeaf =   "frmSupply.frx":0342
      PlusMinusPictureExpanded=   "frmSupply.frx":035E
      PlusMinusPictureExpandable=   "frmSupply.frx":037A
      PlusMinusPictureLeaf=   "frmSupply.frx":0396
      ButtonPicture   =   "frmSupply.frx":03B2
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
      Columns         =   10
      ColWidth0       =   27
      ColTitle0       =   "ID"
      ColBmp0         =   "frmSupply.frx":03CE
      ColWidth1       =   167
      ColTitle1       =   "Names & Titles"
      ColBmp1         =   "frmSupply.frx":03EA
      ColWidth2       =   227
      ColTitle2       =   "Perks & Quirks"
      ColBmp2         =   "frmSupply.frx":0406
      ColWidth3       =   67
      ColTitle3       =   "Ability"
      ColBmp3         =   "frmSupply.frx":0422
      ColWidth4       =   77
      ColStyle4       =   9
      ColTitle4       =   "Status"
      ColBmp4         =   "frmSupply.frx":043E
      ColWidth5       =   33
      ColStyle5       =   9
      ColTitle5       =   "Fight"
      ColBmp5         =   "frmSupply.frx":045A
      ColWidth6       =   33
      ColStyle6       =   9
      ColTitle6       =   "Tech"
      ColBmp6         =   "frmSupply.frx":0476
      ColWidth7       =   34
      ColStyle7       =   9
      ColTitle7       =   "Nego"
      ColBmp7         =   "frmSupply.frx":0492
      ColWidth8       =   60
      ColStyle8       =   10
      ColTitle8       =   "Hire/Pay"
      ColBmp8         =   "frmSupply.frx":04AE
      ColWidth9       =   200
      ColTitle9       =   "Keywords"
      ColBmp9         =   "frmSupply.frx":04CA
      MouseIcon       =   "frmSupply.frx":04E6
      ColHeaderBackColor=   0
      ColHeaderForeColor=   65280
      ForeColor       =   8833235
      BackColor       =   7360778
      SelectStyle     =   2
      NoFocusStyle    =   2
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmSupply.frx":0502
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      OpenEnded       =   0   'False
      ColPict0        =   "frmSupply.frx":051E
      ColPict1        =   "frmSupply.frx":053A
      ColFlag2        =   4
      ColPict2        =   "frmSupply.frx":0556
      ColFlag3        =   12
      ColPict3        =   "frmSupply.frx":0572
      ColFlag4        =   8
      ColPict4        =   "frmSupply.frx":058E
      ColFlag5        =   8
      ColPict5        =   "frmSupply.frx":05AA
      ColFlag6        =   8
      ColPict6        =   "frmSupply.frx":05C6
      ColFlag7        =   8
      ColPict7        =   "frmSupply.frx":05E2
      ColFlag8        =   8
      ColPict8        =   "frmSupply.frx":05FE
      ColPict9        =   "frmSupply.frx":061A
      BackgroundPicture=   "frmSupply.frx":0636
      CharSearchMode  =   2
      ShowFocusRectangle=   0   'False
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1050
      Top             =   2970
   End
   Begin MSComctlLib.ImageList AssetImages 
      Left            =   2730
      Top             =   2910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":0652
            Key             =   "UN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":08E4
            Key             =   "ST"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":0B76
            Key             =   "NT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":17C8
            Key             =   "CS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":201A
            Key             =   "ZS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":286C
            Key             =   "L"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":34BE
            Key             =   "U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":4110
            Key             =   "SG"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":4962
            Key             =   "R"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":55B4
            Key             =   "D"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":6206
            Key             =   "O"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":6E58
            Key             =   "P"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":6FB2
            Key             =   "GR"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupply.frx":72CC
            Key             =   "UP"
         EndProperty
      EndProperty
   End
   Begin XDOCKFLOATLibCtl.FDPane FDPane1 
      Height          =   420
      Left            =   3840
      TabIndex        =   1
      Top             =   2970
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
Attribute VB_Name = "frmSupply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public buyFilter As String

Private Sub Form_Load()
   With sftTree
       Set .ItemPictureExpandable = AssetImages.Overlay("D", "R")
       Set .ItemPictureExpanded = AssetImages.Overlay("D", "R")
       Set .ItemPictureLeaf = AssetImages.Overlay("UN", "O")
       
       'set the splitter to a scrollbar's width from the right side
       '.SplitterOffset = .Width - 1400  '390.165
      
       .LeftButtonOnly = False
       .AutoRespond = True
       .ButtonStyle = buttonsSftTreeAll

    End With
    'Timer1.Enabled = True
End Sub


Public Sub RefreshBuys()
Dim Index, SQL
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim SectorID, SupplyID As Integer, x, first As Boolean, discount As Single
    
With sftTree
   .Clear
   
   SectorID = varDLookup("SectorID", "Players", "PlayerID=" & player.ID)
   If Left(buyFilter, 5) = "local" Then
      If buyFilter = "localbuy" Then
         Me.Caption = "Local Buys for Consideration"
      Else
         Me.Caption = "Local Buys"
      End If
      SupplyID = Nz(varDLookup("SupplyID", "Supply", "SectorID=" & SectorID), 0)

      If SupplyID = 0 Then Exit Sub 'no Deals in this Sector - cater for Alliance/Harken - Fuel

   Else
      Me.Caption = "All Buys"
   End If
   
   SQL = "SELECT * FROM Supply "
   If Left(buyFilter, 5) = "local" Then
      SQL = SQL & "WHERE SupplyID = " & SupplyID
   Else
      SQL = SQL & "WHERE SupplyID >0 "
   End If
   SQL = SQL & " ORDER BY SupplyID"
   rst3.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst3.EOF
      Index = .AddItem(CStr(rst3!SupplyID))
      .ItemLevel(Index) = 0
      .CellText(Index, 1) = rst3!SupplyName
      .CellText(Index, 2) = CStr(getUnseenDeck("Supply", rst3!SupplyID)) & " unseen"
      For x = 0 To 8
         .CellForeColor(Index, x) = 0
         .CellBackColor(Index, x) = rst3!Colour
      Next x
      Set .ItemPicture(Index) = AssetImages.Overlay("L", "U")
      
      'Crew for Hire ____________________________________
      
      SQL = "SELECT SupplyDeck.CardID, SupplyDeck.Seq , Perk.PerkDescription , Crew.* "
      SQL = SQL & "FROM Perk INNER JOIN (Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
      If buyFilter = "localbuy" Then
         SQL = SQL & "WHERE  SupplyDeck.Seq = " & CStr(CONSIDERED)  'only for consideration (6)
      Else
         SQL = SQL & "WHERE (SupplyDeck.Seq = " & CStr(DISCARDED) & " or SupplyDeck.Seq = " & CStr(CONSIDERED) & " ) " 'either discarded (5) or for consideration (6)
      End If
      SQL = SQL & "AND SupplyDeck.CardType=1 AND SupplyDeck.SupplyID = " & rst3!SupplyID
      SQL = SQL & " ORDER BY Crew.CrewID"
      first = True
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         If first Then
            Index = .AddItem("Crew")
            .ItemLevel(Index) = 1
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "P")
            first = False
         End If
         Index = .AddItem(CStr(rst!CardID))
         .ItemData(Index) = rst!CardID
         .CellItemData(Index, 0) = 1 'crew
         .CellItemData(Index, 1) = rst!CrewID
         .ItemLevel(Index) = 2
         If rst!Seq = 6 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
            .ItemDataString(Index) = "UN"
         Else
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
            .ItemDataString(Index) = "O"
         End If
         .CellText(Index, 1) = rst!CrewName & "  -  " & rst!CrewDescr

         .CellText(Index, 2) = rst!PerkDescription
         .CellText(Index, 3) = Trim(IIf(rst!Mechanic = 1, "Mechanic  ", "") & IIf(rst!Pilot = 1, "Pilot  ", "") & IIf(rst!Companion = 1, "Companion  ", "") & _
               IIf(rst!Merc = 1, "Merc  ", "") & IIf(rst!Soldier = 1, "Soldier  ", "") & IIf(rst!HillFolk = 1, "HillFolk  ", "") & _
               IIf(rst!Grifter = 1, "Grifter ", "") & IIf(rst!Medic = 1, "Medic", ""))
         .CellForeColor(Index, 3) = 65280
         
         .CellText(Index, 4) = IIf(rst!wanted > 0, "Wanted", "") & IIf(rst!Moral = 1, IIf(rst!wanted > 0, "/", "") & "Moral ", "")
         .CellForeColor(Index, 4) = 0
         If rst!wanted > 0 Then
            .CellBackColor(Index, 4) = &HC0C0FF
         ElseIf rst!Moral = 1 Then
            .CellBackColor(Index, 4) = &HC0FFC0
         End If
         
         .CellText(Index, 5) = IIf(rst!fight > 0, CStr(rst!fight), "")
         .CellForeColor(Index, 5) = 0
         If rst!fight > 0 Then .CellBackColor(Index, 5) = 6052315
         
         .CellText(Index, 6) = IIf(rst!tech > 0, CStr(rst!tech), "")
         .CellForeColor(Index, 6) = 0
         If rst!tech > 0 Then .CellBackColor(Index, 6) = 16382208
         
         .CellText(Index, 7) = IIf(rst!Negotiate > 0, CStr(rst!Negotiate), "")
         .CellForeColor(Index, 7) = 0
         If rst!Negotiate > 0 Then .CellBackColor(Index, 7) = 5373777
         
         If freeCrew(player.ID) Then
            .CellText(Index, 8) = "free/$" & rst!pay
            .CellForeColor(Index, 8) = 65280
         Else
            .CellText(Index, 8) = "$" & rst!pay
            .CellItemData(Index, 8) = rst!pay
            .CellForeColor(Index, 8) = 16777215
            If rst!pay > getMoney(player.ID) Then
               .CellForeColor(Index, 8) = 255
            End If
         End If
         .CellBackColor(Index, 8) = 8388736
         
         .CellText(Index, 9) = Nz(rst!KeyWords)   'IIf(rst!Wanted > 0, "Wanted ", "") &
         .CellForeColor(Index, 9) = 0
         If Not IsNull(rst!KeyWords) Then
            .CellForeColor(Index, 9) = 65280
         End If
         
         rst.MoveNext
      Wend
      rst.Close
      
      'Gear for Purchase ____________________________________
      
      SQL = "SELECT SupplyDeck.CardID, SupplyDeck.Seq, Gear.* "
      SQL = SQL & "FROM Gear INNER JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID "
      If buyFilter = "localbuy" Then
         SQL = SQL & "WHERE  SupplyDeck.Seq = " & CStr(CONSIDERED)  'only for consideration (6)
      Else
         SQL = SQL & "WHERE (SupplyDeck.Seq = " & CStr(DISCARDED) & " or SupplyDeck.Seq = " & CStr(CONSIDERED) & " ) " 'either discarded (5) or for consideration (6)
      End If
      SQL = SQL & "AND SupplyDeck.CardType=2 AND SupplyDeck.SupplyID = " & rst3!SupplyID

      first = True
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         If first Then
            Index = .AddItem("Gear")
            .ItemLevel(Index) = 1
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "GR")
            first = False
         End If
         Index = .AddItem(CStr(rst!CardID))
         .ItemData(Index) = rst!CardID
         .CellItemData(Index, 0) = 2 'gear
         .CellItemData(Index, 1) = rst!GearID
         .ItemLevel(Index) = 2
         If rst!Seq = 6 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
            .ItemDataString(Index) = "UN"
         Else
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
            .ItemDataString(Index) = "O"
         End If
         .CellText(Index, 1) = rst!GearName
         .CellForeColor(Index, 1) = 16685961
         .CellText(Index, 2) = rst!GearDescr
         .CellForeColor(Index, 2) = 16685961
         
         .CellText(Index, 5) = IIf(rst!fight > 0, CStr(rst!fight), "")

         If rst!discard = 1 Then
            .CellForeColor(Index, 5) = 65280
         Else
            .CellForeColor(Index, 5) = 0
         End If
         If rst!fight > 0 Then .CellBackColor(Index, 5) = 6052315
         
         
         .CellText(Index, 6) = IIf(rst!tech > 0, CStr(rst!tech), "")
         If rst!discard = 1 Then
            .CellForeColor(Index, 6) = 255
         Else
            .CellForeColor(Index, 6) = 0
         End If
         If rst!tech > 0 Then .CellBackColor(Index, 6) = 16382208
         
         .CellText(Index, 7) = IIf(rst!Negotiate > 0, CStr(rst!Negotiate), "")
         If rst!discard = 1 Then
            .CellForeColor(Index, 7) = 255
         Else
            .CellForeColor(Index, 7) = 0
         End If
         If rst!Negotiate > 0 Then .CellBackColor(Index, 7) = 5373777
         
         If InStr(rst!KeyWords, "EXPLOSIVES") > 0 Then
            discount = discounts(player.ID, "ExplosivesDiscount")
         ElseIf InStr(rst!KeyWords, "FIREARM") > 0 Then
            discount = discounts(player.ID, "FirearmDiscount")
         Else
            discount = 0
         End If
         
         If discount > 0 Then
           .CellText(Index, 8) = "$" & CStr(Int(rst!pay * discount))
           .CellItemData(Index, 8) = Int(rst!pay * discount)
           .CellForeColor(Index, 8) = 65280
         Else
           .CellText(Index, 8) = "$" & rst!pay
           .CellItemData(Index, 8) = rst!pay
           .CellForeColor(Index, 8) = 16777215
         End If
         .CellBackColor(Index, 8) = 8388736
         If .CellItemData(Index, 8) > getMoney(player.ID) Then
            .CellForeColor(Index, 8) = 255
         End If
         
         'Keywords
         .CellText(Index, 9) = Nz(rst!KeyWords, "")
         .CellForeColor(Index, 9) = 65280
         
         rst.MoveNext
      Wend
      rst.Close
      
      
      'ShipUpgrades for Purchase ____________________________________
      
      SQL = "SELECT SupplyDeck.CardID, SupplyDeck.Seq, ShipUpgrade.* "
      SQL = SQL & "FROM ShipUpgrade INNER JOIN SupplyDeck ON ShipUpgrade.ShipUpgradeID = SupplyDeck.ShipUpgradeID "
      If buyFilter = "localbuy" Then
         SQL = SQL & "WHERE  SupplyDeck.Seq = " & CStr(CONSIDERED)  'only for consideration (6)
      Else
         SQL = SQL & "WHERE (SupplyDeck.Seq = " & CStr(DISCARDED) & " or SupplyDeck.Seq = " & CStr(CONSIDERED) & " ) " 'either discarded (5) or for consideration (6)
      End If
      SQL = SQL & "AND SupplyDeck.CardType=3 AND SupplyDeck.SupplyID = " & rst3!SupplyID

      first = True
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         If first Then
            Index = .AddItem("Upgrades")
            .ItemLevel(Index) = 1
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "UP")
            first = False
         End If
         Index = .AddItem(CStr(rst!CardID))
         .ItemData(Index) = rst!CardID
         .CellItemData(Index, 0) = 3 'upgrd
         .CellItemData(Index, 1) = rst!ShipUpgradeID
         .ItemLevel(Index) = 2
         If rst!Seq = 6 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
            .ItemDataString(Index) = "UN"
         Else
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
            .ItemDataString(Index) = "O"
         End If
         .CellText(Index, 1) = rst!UpgradeName
         .CellForeColor(Index, 1) = 8823762
         'Drive Cores are marked for swapping out
         .CellText(Index, 2) = IIf(rst!DriveCore = 1, "DriveCore: ", "") & rst!UpgradeDescr
         .CellItemData(Index, 2) = rst!DriveCore
         .CellForeColor(Index, 2) = 8823762
         
         .CellText(Index, 3) = IIf(rst!burnFuel > 0, "Full Burn Fuel:" & rst!burnFuel & ", ", "") & IIf(rst!DriveCore = 1, "BurnRange: " & CStr(rst!BurnRange + 5) & ", MoseyRange: " & CStr(rst!MoseyRange), "")
         
          discount = discounts(player.ID, "ShipUpgrades")
          If discount > 0 Then
            .CellText(Index, 8) = "$" & CStr(Int(rst!pay * discount))
            .CellItemData(Index, 8) = Int(rst!pay * discount)
            .CellForeColor(Index, 8) = 65280
          Else
            .CellText(Index, 8) = "$" & rst!pay
            .CellItemData(Index, 8) = rst!pay
            .CellForeColor(Index, 8) = 16777215
          End If
         .CellBackColor(Index, 8) = 8388736
         If .CellItemData(Index, 8) > getMoney(player.ID) Then
            .CellForeColor(Index, 8) = 255
         End If
         
         
         rst.MoveNext
      Wend
      rst.Close
      rst3.MoveNext
   Wend
 End With
   
End Sub

Private Sub Form_Resize()
   sftTree.Move sftTree.Left, sftTree.top, Abs(Me.Width - 100), Abs(Me.Height - sftTree.top)
   
End Sub

Private Sub sftTree_ItemClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
Dim max As Integer
With sftTree

  If Button = constSftTreeLeftButton And (AreaType = constSftTreeItem Or AreaType = constSftTreeCellText) Then
      
         Select Case .ItemDataString(Index)
         Case "UN" 'consider
            Select Case actionSeq
               Case ASBuySelDiscard
                  .ItemDataString(Index) = "O"
                  Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
                  If getSelected("UN") = MAXJOBCARDDRAW Then
                     frmAction.cmd(2).Caption = "Consider"
                     frmAction.cmd(2).Enabled = True
                  Else
                     frmAction.cmd(2).Caption = "Draw Cards"
                     frmAction.cmd(2).Enabled = (getUnseenDeck("Supply", Val(frmAction.lblSupply.Tag)) > 0)
                  End If
                  
               Case ASBuySelect
                  'determine how many cards can be accepted
                  max = MAXJOBCARDACCEPT + getGearFeature(player.ID, "MaxSupplies") 'accept only up to 2 cards + modifier
                  ' check for ship upgrades that there are Slots left. Drives OK (but only 1), and only 3 upgrades. if (getShipUpgrades(player.id)< 3 and .CellItemData(index, 0) = 3) 'upgrd
                  If getSelected("R") < max And companionsOK(player.ID, .CellItemData(Index, 0), .CellItemData(Index, 1)) And _
                     ((getCost("R") + .CellItemData(Index, 8)) <= getMoney(player.ID)) And _
                     ((.CellItemData(Index, 0) = 1 And (getCrewCount(player.ID) + getCrewSelected("R")) < CrewCapacity(player.ID) + getCrewSpaceSelected("R")) Or _
                       .CellItemData(Index, 0) = 2 Or _
                      (.CellItemData(Index, 0) = 3 And ((getShipUpgrades(player.ID) + getUpgradesSelected("R") < 3 And Not isDriveCore(.ItemData(Index))) Or (isDriveCore(.ItemData(Index)) And getDriveCoresSelected("R") = 0)))) Then
                     
                     .ItemDataString(Index) = "R"
                     Set .ItemPicture(Index) = AssetImages.Overlay("L", "R")
                  Else
                     playsnd 9
                  End If
                  
            End Select
            
         Case "R"  'deal
            If actionSeq = ASBuySelect Then
                  .ItemDataString(Index) = "UN"
                  'check if trying to de-select crew expansion which leaves too many crew to carry
                  If .CellItemData(Index, 0) = 3 And (getCrewCount(player.ID) + getCrewSelected("R")) > CrewCapacity(player.ID) + getCrewSpaceSelected("R") Then
                     .ItemDataString(Index) = "R" 'block de-selection until crew de-selected
                     playsnd 9
                  Else
                     Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
                  End If
            End If
            
         Case "O"  'discard
            If actionSeq = ASBuySelDiscard And getSelected("UN") < MAXJOBCARDDRAW Then 'can consider up to 3 cards
                  .ItemDataString(Index) = "UN"
                  Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
                  If getSelected("UN") = MAXJOBCARDDRAW Then
                     frmAction.cmd(2).Caption = "Consider"
                     frmAction.cmd(2).Enabled = True
                  Else
                     frmAction.cmd(2).Caption = "Draw Cards"
                     frmAction.cmd(2).Enabled = (getUnseenDeck("Supply", Val(frmAction.lblSupply.Tag)) > 0)
                  End If
                  
            End If
            
         End Select
         getCost "R"
   End If
   
End With

End Sub

Private Function getCost(ByVal status As String) As Integer
Dim Index As Integer, pay As Integer
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = status Then
            getCost = getCost + .CellItemData(Index, 8)
            pay = pay + .CellItemData(Index, 8)
         End If
      Next Index
      .CellText(0, 8) = "$" & CStr(pay)
   End With

End Function

Private Function getCrewSelected(ByVal status As String) As Integer
Dim Index As Integer
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = status And .CellItemData(Index, 0) = 1 Then
            getCrewSelected = getCrewSelected + 1
         End If
      Next Index
   
   End With

End Function

Private Function getUpgradesSelected(ByVal status As String) As Integer
Dim Index As Integer
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = status And .CellItemData(Index, 0) = 3 Then
            getUpgradesSelected = getUpgradesSelected + 1
         End If
      Next Index
   
   End With

End Function

Private Function getCrewSpaceSelected(ByVal status As String) As Integer
Dim Index As Integer
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = status And .CellItemData(Index, 0) = 3 Then
            getCrewSpaceSelected = getCrewSpaceSelected + varDLookup("ExtraCrewSpace", "ShipUpgrade", "ShipUpgradeID=" & .CellItemData(Index, 1))
         End If
      Next Index
   
   End With

End Function

Private Function getDriveCoresSelected(ByVal status As String) As Integer
Dim Index As Integer
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = status And .CellItemData(Index, 0) = 3 Then
            If isDriveCore(.ItemData(Index)) Then
               getDriveCoresSelected = getDriveCoresSelected + 1
            End If
         End If
      Next Index
   
   End With

End Function

Private Function getSelected(ByVal status As String) As Integer
Dim Index As Integer
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = status Then
            getSelected = getSelected + 1
            
         End If
      Next Index
      
   End With

End Function

Public Function setSelected(ByVal status As String, ByVal Seq As Integer) As Integer
Dim Index As Integer
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = status Then
            'mark the card as the players
            DB.Execute "UPDATE SupplyDeck SET Seq =" & Seq & " WHERE CardID = " & .ItemData(Index)
            setSelected = setSelected + 1
         End If
      Next Index
   
   End With


End Function

Private Sub sftTree_ItemDblClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)


   If Button = constSftTreeLeftButton And AreaType = constSftTreeCellText Then
      If sftTree.CellItemData(Index, 0) = 1 Then
         Dim frmCrew As New frmCrewSel
         frmCrew.crewFilter = " WHERE CrewID =" & sftTree.CellItemData(Index, 1)
         frmCrew.AlwaysOnTop = True
         frmCrew.Show
         Set frmCrew = Nothing
      End If
      If sftTree.CellItemData(Index, 0) = 2 Then
         Dim frmGear As New frmGearView
         frmGear.gearFilter = " WHERE CardID=" & sftTree.ItemData(Index)
         frmGear.AlwaysOnTop = True
         frmGear.Show
         Set frmGear = Nothing
      End If
      If sftTree.CellItemData(Index, 0) = 3 Then
         Dim frmUpGrd As New frmShipUpgrdView
         frmUpGrd.gearFilter = " WHERE CardID=" & sftTree.ItemData(Index)
         frmUpGrd.AlwaysOnTop = True
         frmUpGrd.Show
         Set frmUpGrd = Nothing
      End If
   End If
End Sub

Private Sub Timer1_Timer()
   If FDPane1.PaneVisible Then RefreshBuys
End Sub


