VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmDeals 
   Caption         =   "Deals"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   Icon            =   "frmDeals.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   4875
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
      BackColor       =   3353720
      BorderStyle     =   1
      ItemPictureExpanded=   "frmDeals.frx":030A
      ItemPictureExpandable=   "frmDeals.frx":0326
      ItemPictureLeaf =   "frmDeals.frx":0342
      PlusMinusPictureExpanded=   "frmDeals.frx":035E
      PlusMinusPictureExpandable=   "frmDeals.frx":037A
      PlusMinusPictureLeaf=   "frmDeals.frx":0396
      ButtonPicture   =   "frmDeals.frx":03B2
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
      Columns         =   9
      ColTitle0       =   "Card ID"
      ColBmp0         =   "frmDeals.frx":03CE
      ColWidth1       =   200
      ColTitle1       =   "Contact / Instructions"
      ColBmp1         =   "frmDeals.frx":03EA
      ColWidth2       =   117
      ColTitle2       =   "Job Type / Planet"
      ColBmp2         =   "frmDeals.frx":0406
      ColWidth3       =   120
      ColTitle3       =   "Needs / System"
      ColBmp3         =   "frmDeals.frx":0422
      ColWidth4       =   41
      ColStyle4       =   10
      ColTitle4       =   "Pay"
      ColBmp4         =   "frmDeals.frx":043E
      ColWidth5       =   99
      ColTitle5       =   "Bonus"
      ColBmp5         =   "frmDeals.frx":045A
      ColWidth6       =   33
      ColStyle6       =   9
      ColTitle6       =   "Fight"
      ColBmp6         =   "frmDeals.frx":0476
      ColWidth7       =   33
      ColStyle7       =   9
      ColTitle7       =   "Tech"
      ColBmp7         =   "frmDeals.frx":0492
      ColWidth8       =   37
      ColStyle8       =   9
      ColTitle8       =   "Negot"
      ColBmp8         =   "frmDeals.frx":04AE
      MouseIcon       =   "frmDeals.frx":04CA
      ColHeaderBackColor=   0
      ColHeaderForeColor=   10937324
      ForeColor       =   8833235
      BackColor       =   3353720
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmDeals.frx":04E6
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      OpenEnded       =   0   'False
      ColPict0        =   "frmDeals.frx":0502
      ColPict1        =   "frmDeals.frx":051E
      ColPict2        =   "frmDeals.frx":053A
      ColPict3        =   "frmDeals.frx":0556
      ColPict4        =   "frmDeals.frx":0572
      ColPict5        =   "frmDeals.frx":058E
      ColPict6        =   "frmDeals.frx":05AA
      ColPict7        =   "frmDeals.frx":05C6
      ColPict8        =   "frmDeals.frx":05E2
      BackgroundPicture=   "frmDeals.frx":05FE
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":061A
            Key             =   "UN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":08AC
            Key             =   "ST"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":0B3E
            Key             =   "NT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":1790
            Key             =   "CS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":1FE2
            Key             =   "ZS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":2834
            Key             =   "L"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":3486
            Key             =   "U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":40D8
            Key             =   "SG"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":492A
            Key             =   "R"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":557C
            Key             =   "D"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":61CE
            Key             =   "O"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeals.frx":6E20
            Key             =   "LN"
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
Attribute VB_Name = "frmDeals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dealFilter As String


Private Sub FDPane1_OnHidden()
   Main.drawLine 1, -1
End Sub

Private Sub Form_Load()
   With sftTree
       Set .ItemPictureExpandable = AssetImages.Overlay("D", "R")
       Set .ItemPictureExpanded = AssetImages.Overlay("D", "R")
       Set .ItemPictureLeaf = AssetImages.Overlay("LN", "LN")
       
       'set the splitter to a scrollbar's width from the right side
       '.SplitterOffset = .Width - 1400  '390.165
      
       .LeftButtonOnly = False
       .AutoRespond = True
       .ButtonStyle = buttonsSftTreeAll

    End With
    'Timer1.Enabled = True
End Sub


Public Function RefreshDeals() As Variant
Dim Index, SQL
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim SectorID, ContactID As Integer, x, cnt As Integer
    
With sftTree
   .Clear
   
   SectorID = varDLookup("SectorID", "Players", "PlayerID=" & player.ID)
   If Left(dealFilter, 5) = "local" Then
      If dealFilter = "localdeal" Then
         Me.Caption = "Local Deals for Consideration"
      Else
         Me.Caption = "Local Deals"
      End If
      ContactID = Nz(varDLookup("ContactID", "Contact", "SectorID=" & SectorID), 0)
      
      If HigginsDealPerk Or (hasCrew(player.ID, 75) And ContactID = 0) And dealFilter = "locals" And Not hasCrew(player.ID, 22) Then
         ContactID = 8
         HigginsDealPerk = True
      ElseIf hasCrew(player.ID, 75) And dealFilter = "locals" And Not hasCrew(player.ID, 22) And SectorID <> 16 Then
         If MessBox("Do you want Deal with Higgins instead?", "Fess - Phone Home Deals", "Yes", "No", 75) = 0 Then
         'If MsgBox("Do you want Deal with Higgins instead?", vbQuestion + vbYesNo, "Fess - Phone Home Deals") = vbYes Then
            ContactID = 8
            HigginsDealPerk = True
         End If
      Else
         If ContactID = 0 Then Exit Function 'no Deals in this Sector
      End If
      
   Else
      Me.Caption = "All Deals"
   End If
   
   SQL = "SELECT * FROM Contact "
   If Left(dealFilter, 5) = "local" Then
      SQL = SQL & "WHERE ContactID = " & ContactID
   Else
      SQL = SQL & "WHERE ContactID <> 0"
   End If
   SQL = SQL & " ORDER BY ContactName"
   rst3.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst3.EOF
      Index = .AddItem(CStr(rst3!ContactID) & IIf(isSolid(player.ID, rst3!ContactID), " - Solid", ""))
      .ItemLevel(Index) = 0
      .CellText(Index, 1) = rst3!ContactName  '& IIf(isSolid(player.ID, rst3!ContactID), " - Solid", "")
      .CellText(Index, 2) = CStr(getUnseenDeck("Contact", rst3!ContactID)) & " unseen"
      For x = 0 To 8
         .CellForeColor(Index, x) = 0
         .CellBackColor(Index, x) = rst3!Colour
      Next x
      Set .ItemPicture(Index) = LoadPicture(App.Path & "\Pictures\Sm" & Nz(varDLookup("Picture", "Contact", "ContactID=" & rst3!ContactID)))
      'Set .ItemPicture(Index) = AssetImages.Overlay("L", "U")
    
      SQL = "SELECT Contact.Colour, JobType.JobTypeDescr, Profession.ProfessionName, ContactDeck.*, JobType_1.JobTypeDescr AS JobType2 "
      SQL = SQL & "FROM (Contact INNER JOIN ((ContactDeck INNER JOIN JobType ON ContactDeck.JobTypeID = JobType.JobTypeID) LEFT JOIN Profession "
      SQL = SQL & "ON ContactDeck.ProfessionID = Profession.ProfessionID) ON Contact.ContactID = ContactDeck.ContactID) INNER JOIN JobType AS JobType_1 ON ContactDeck.JobType2D = JobType_1.JobTypeID "
      'SQL = SQL & "LEFT JOIN Profession ON ContactDeck.ProfessionID = Profession.ProfessionID) ON Contact.ContactID = ContactDeck.ContactID "
      If dealFilter = "localdeal" Then
         SQL = SQL & "WHERE  Seq = " & CStr(CONSIDERED)  'only for consideration (6)
      Else
         SQL = SQL & "WHERE (Seq = " & CStr(DISCARDED) & " or Seq = " & CStr(CONSIDERED) & " ) " 'either discarded (5) or for consideration (6)
      End If
      SQL = SQL & "AND ContactDeck.ContactID = " & rst3!ContactID
      SQL = SQL & " ORDER BY Contact.ContactName,ContactDeck.CardID"
      
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         cnt = cnt + 1
         Index = .AddItem(CStr(rst!CardID))
         .ItemData(Index) = rst!CardID
         .ItemLevel(Index) = 1
         If rst!Seq = 6 Then
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
            .ItemDataString(Index) = "UN"
         Else
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
            .ItemDataString(Index) = "O"
         End If
         .CellText(Index, 1) = rst!JobName
         .CellForeColor(Index, 1) = 0
         .CellBackColor(Index, 1) = rst!Colour
         .CellText(Index, 2) = rst!JobTypeDescr & IIf(rst!JobType2 <> "-", "/" & rst!JobType2, "") & IIf(rst!illegal = 1, "/illegal", "") & IIf(rst!Immoral = 1, "/immoral", "")
         If rst!illegal = 1 Or rst!Immoral Then
            .CellBackColor(Index, 2) = 3355647
         End If
         .CellText(Index, 3) = Nz(rst!JobOrder)
         .CellForeColor(Index, 3) = 51712
         .CellText(Index, 4) = "$" & rst!pay
         .CellBackColor(Index, 4) = 8388736
         .CellForeColor(Index, 4) = 16777215
                              
         .CellText(Index, 5) = IIf(rst!BonusPart > 0, " +" & rst!BonusPart & " part: ", "") & IIf(rst!bonus > 0, " +$" & rst!bonus & ":", "") & IIf(rst!KeywordBonus = 1, rst!KeyWords, "") & IIf(IsNull(rst!ProfessionName), "", " " & rst!ProfessionName) & IIf(rst!BonusPerSkill > 0, " /" & cstrSkill(rst!BonusPerSkill), "") & IIf(rst!Job3ID > 0, "Bonus Job", "")
         If rst!BonusPart > 0 Or rst!bonus > 0 Then
             .CellForeColor(Index, 5) = 0
            .CellBackColor(Index, 5) = 1900316
         End If
         .CellText(Index, 6) = IIf(rst!fight > 0, CStr(rst!fight), "")
         .CellForeColor(Index, 6) = 0
         If rst!fight > 0 Then .CellBackColor(Index, 6) = 6052315
         .CellText(Index, 7) = IIf(rst!tech > 0, CStr(rst!tech), "")
         .CellForeColor(Index, 7) = 0
         If rst!tech > 0 Then .CellBackColor(Index, 7) = 16382208
         .CellText(Index, 8) = IIf(rst!Negotiate > 0, CStr(rst!Negotiate), "")
         .CellForeColor(Index, 8) = 0
         If rst!Negotiate > 0 Then .CellBackColor(Index, 8) = 5373777

         
         If rst!Job1ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job1ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
                Index = .AddItem(CStr(rst2!JobID))
                .ItemLevel(Index) = 2
               .CellText(Index, 1) = rst2!JobDesc
               x = getSectorCount(getPlayerSector(player.ID), rst2!SectorID)
               .CellText(Index, 2) = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "")
               If SectorID = rst2!SectorID Then
                  .CellFont(Index, 2).Bold = True
                  .CellFont(Index, 3).Bold = True
                  If hasJobReqs(player.ID, rst!CardID, rst!Job1ID) Then
                     .CellForeColor(Index, 2) = 0
                     .CellForeColor(Index, 3) = 0
                  Else
                     .CellForeColor(Index, 2) = 255
                     .CellForeColor(Index, 3) = 255
                  End If
                  .CellBackColor(Index, 2) = &HC0FFC0
                  
                  .CellBackColor(Index, 3) = &HC0FFC0
               End If
               .CellText(Index, 3) = rst2!System
               Set .ItemPicture(Index) = AssetImages.Overlay("LN", "LN")
               .CellItemData(Index, 1) = rst2!SectorID
            Else
               MsgBox "Job Card " & rst!Job1ID & " Error", vbCritical
            End If
            rst2.Close
         End If
         
                  'Bonus Drop Job
         
         If rst!Job3ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job3ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
                Index = .AddItem(CStr(rst2!JobID))
                .ItemLevel(Index) = 3
               .CellText(Index, 1) = rst2!JobDesc
               x = getSectorCount(getPlayerSector(player.ID), rst2!SectorID)
               .CellText(Index, 2) = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "")
               If SectorID = rst2!SectorID Then
                  .CellFont(Index, 2).Bold = True
                  .CellFont(Index, 3).Bold = True
                  If hasJobReqs(player.ID, rst!CardID, rst!Job3ID) Then
                     .CellForeColor(Index, 2) = 0
                     .CellForeColor(Index, 3) = 0
                  Else
                     .CellForeColor(Index, 2) = 255
                     .CellForeColor(Index, 3) = 255
                  End If
                  .CellBackColor(Index, 2) = &HC0FFC0
                  
                  .CellBackColor(Index, 3) = &HC0FFC0
               End If
               .CellText(Index, 3) = rst2!System
               Set .ItemPicture(Index) = AssetImages.Overlay("LN", "LN")
               .CellItemData(Index, 1) = rst2!SectorID
            Else
               MsgBox "Job Card " & rst!Job3ID & " Error", vbCritical
            End If
            rst2.Close
         End If
         
         
         If rst!Job2ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job2ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
                Index = .AddItem(CStr(rst2!JobID))
                .ItemLevel(Index) = 2
               .CellText(Index, 1) = rst2!JobDesc
               x = getSectorCount(getPlayerSector(player.ID), rst2!SectorID)
               .CellText(Index, 2) = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "")
               If SectorID = rst2!SectorID Then
                  .CellFont(Index, 2).Bold = True
                  .CellFont(Index, 3).Bold = True
                  If hasJobReqs(player.ID, rst!CardID, rst!Job2ID) Then
                     .CellForeColor(Index, 2) = 0
                     .CellForeColor(Index, 3) = 0
                  Else
                     .CellForeColor(Index, 2) = 255
                     .CellForeColor(Index, 3) = 255
                  End If
                  .CellBackColor(Index, 2) = &HC0FFC0
                  
                  .CellBackColor(Index, 3) = &HC0FFC0
               End If
               .CellText(Index, 3) = rst2!System
               Set .ItemPicture(Index) = AssetImages.Overlay("LN", "LN")
               .CellItemData(Index, 1) = rst2!SectorID
            Else
               MsgBox "Job Card " & rst!Job2ID & " Error", vbCritical
            End If
            rst2.Close
         End If
         
         rst.MoveNext
      Wend
      rst.Close
      rst3.MoveNext
   Wend
 End With
 RefreshDeals = cnt
 
End Function

Private Sub Form_Resize()
   sftTree.Move sftTree.Left, sftTree.top, Abs(Me.Width - 100), Abs(Me.Height - sftTree.top)
   
End Sub

Private Sub sftTree_ItemClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
Dim max As Integer, maxConsider
With sftTree

  If Button = constSftTreeLeftButton And (AreaType = constSftTreeItem Or AreaType = constSftTreeCellText) Then
         maxConsider = MAXJOBCARDDRAW + getGearFeature(player.ID, "MaxJobs")
         If isSolid(player.ID, 4) And Val(sftTree.CellText(0, 0)) = 4 Then
            maxConsider = 4
         End If
         
         If sftTree.CellItemData(Index, 1) > 0 Then
             'draw a line
             Main.drawLine 1, sftTree.CellItemData(Index, 1), varDLookup("SectorID", "Players", "PlayerID=" & player.ID), False
          Else
             Main.drawLine 1, -1
         End If
         
         Select Case .ItemDataString(Index)
         Case "UN" 'consider
            Select Case actionSeq
               Case ASDealSelDiscard
                  .ItemDataString(Index) = "O"
                  Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
                  If getSelected("UN") = maxConsider Then
                     frmAction.cmd(3).Caption = "Consider"
                     frmAction.cmd(3).Enabled = True
                  Else
                     frmAction.cmd(3).Caption = "Draw Cards"
                     frmAction.cmd(3).Enabled = (getUnseenDeck("Contact", Val(sftTree.CellText(0, 0))) > 0)
                  End If
               Case ASDealSelect
                  'determine how many cards can be accepted
                  
                  max = MAXINACTIVEJOBS - getPlayerJobs(player.ID, "0")

                  If getSelected("R") < max Then  'accept only up to 2 cards, not exceeding 6 in hand
                     .ItemDataString(Index) = "R"
                     Set .ItemPicture(Index) = AssetImages.Overlay("L", "R")
                  Else
                     playsnd 9
                  End If
            End Select
            
         Case "R"  'deal
            If actionSeq = ASDealSelect Then
                  .ItemDataString(Index) = "UN"
                  Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
            End If
            
         Case "O"  'discard
            If actionSeq = ASDealSelDiscard And getSelected("UN") < maxConsider Then 'can consider up to 3 cards
                  .ItemDataString(Index) = "UN"
                  Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
                  If getSelected("UN") = maxConsider Or getUnseenDeck("Contact", Val(sftTree.CellText(0, 0))) = 0 Then
                     frmAction.cmd(3).Caption = "Consider"
                     frmAction.cmd(3).Enabled = True
                  Else
                     frmAction.cmd(3).Caption = "Draw Cards"
                     frmAction.cmd(3).Enabled = (getUnseenDeck("Contact", Val(sftTree.CellText(0, 0))) > 0)
                  End If
                  
            End If
            
         End Select
      
   End If
   
End With

End Sub

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
            DB.Execute "UPDATE ContactDeck SET Seq =" & Seq & " WHERE CardID = " & .ItemData(Index)
            setSelected = setSelected + 1
         End If
      Next Index
   
   End With


End Function

Private Sub Timer1_Timer()
   If FDPane1.PaneVisible Then RefreshDeals
End Sub

