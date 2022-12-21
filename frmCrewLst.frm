VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmCrewLst 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Select Crew to Pay"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SftTree.SftTree sftTree 
      Height          =   3435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11715
      _Version        =   262144
      _ExtentX        =   20664
      _ExtentY        =   6059
      _StockProps     =   237
      ForeColor       =   16777215
      BackColor       =   8388669
      BorderStyle     =   1
      ItemPictureExpanded=   "frmCrewLst.frx":0000
      ItemPictureExpandable=   "frmCrewLst.frx":001C
      ItemPictureLeaf =   "frmCrewLst.frx":0038
      PlusMinusPictureExpanded=   "frmCrewLst.frx":0054
      PlusMinusPictureExpandable=   "frmCrewLst.frx":0070
      PlusMinusPictureLeaf=   "frmCrewLst.frx":008C
      ButtonPicture   =   "frmCrewLst.frx":00A8
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
      Columns         =   10
      ColTitle0       =   "CardID"
      ColBmp0         =   "frmCrewLst.frx":00C4
      ColWidth1       =   133
      ColTitle1       =   "Names and Titles"
      ColBmp1         =   "frmCrewLst.frx":00E0
      ColWidth2       =   187
      ColTitle2       =   "Perks and Quirks"
      ColBmp2         =   "frmCrewLst.frx":00FC
      ColWidth3       =   67
      ColTitle3       =   "Ability"
      ColBmp3         =   "frmCrewLst.frx":0118
      ColWidth4       =   53
      ColStyle4       =   9
      ColTitle4       =   "Status"
      ColBmp4         =   "frmCrewLst.frx":0134
      ColWidth5       =   33
      ColStyle5       =   9
      ColTitle5       =   "Fight"
      ColBmp5         =   "frmCrewLst.frx":0150
      ColWidth6       =   33
      ColStyle6       =   9
      ColTitle6       =   "Tech"
      ColBmp6         =   "frmCrewLst.frx":016C
      ColWidth7       =   37
      ColStyle7       =   9
      ColTitle7       =   "Nego"
      ColBmp7         =   "frmCrewLst.frx":0188
      ColWidth8       =   47
      ColStyle8       =   10
      ColTitle8       =   "Pay/job"
      ColBmp8         =   "frmCrewLst.frx":01A4
      ColWidth9       =   107
      ColStyle9       =   9
      ColTitle9       =   "Special Info"
      ColBmp9         =   "frmCrewLst.frx":01C0
      MouseIcon       =   "frmCrewLst.frx":01DC
      ColHeaderBackColor=   0
      ColHeaderForeColor=   65280
      ForeColor       =   16777215
      BackColor       =   8388669
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmCrewLst.frx":01F8
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      OpenEnded       =   0   'False
      ColPict0        =   "frmCrewLst.frx":0214
      ColPict1        =   "frmCrewLst.frx":0230
      ColFlag2        =   4
      ColPict2        =   "frmCrewLst.frx":024C
      ColFlag3        =   12
      ColPict3        =   "frmCrewLst.frx":0268
      ColFlag4        =   8
      ColPict4        =   "frmCrewLst.frx":0284
      ColFlag5        =   8
      ColPict5        =   "frmCrewLst.frx":02A0
      ColFlag6        =   8
      ColPict6        =   "frmCrewLst.frx":02BC
      ColFlag7        =   8
      ColPict7        =   "frmCrewLst.frx":02D8
      ColFlag8        =   8
      ColPict8        =   "frmCrewLst.frx":02F4
      ColFlag9        =   8
      ColPict9        =   "frmCrewLst.frx":0310
      BackgroundPicture=   "frmCrewLst.frx":032C
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin VB.CheckBox chk 
      Caption         =   "Others"
      Height          =   255
      Index           =   2
      Left            =   8100
      TabIndex        =   5
      Top             =   5070
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CheckBox chk 
      Caption         =   "Moral"
      Height          =   255
      Index           =   1
      Left            =   6750
      TabIndex        =   4
      Top             =   5070
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CheckBox chk 
      Caption         =   "Wanted"
      Height          =   255
      Index           =   0
      Left            =   5250
      TabIndex        =   3
      Top             =   5070
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1155
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
      Left            =   10620
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4980
      Width           =   1035
   End
   Begin MSComctlLib.ImageList AssetImages 
      Left            =   10950
      Top             =   3990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":0348
            Key             =   "UN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":05DA
            Key             =   "ST"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":086C
            Key             =   "NT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":14BE
            Key             =   "CS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":1D10
            Key             =   "ZS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":2562
            Key             =   "L"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":31B4
            Key             =   "U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":3E06
            Key             =   "SG"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":4658
            Key             =   "R"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":52AA
            Key             =   "D"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":5EFC
            Key             =   "O"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":6B4E
            Key             =   "P"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":6CA8
            Key             =   "PS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":6FC2
            Key             =   "LN"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":72DC
            Key             =   "CN"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":75F6
            Key             =   "GR"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrewLst.frx":7910
            Key             =   "UP"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   5070
      Width           =   3615
   End
End
Attribute VB_Name = "frmCrewLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' selectCrew: 0= job payment mode.  >0 = No. of Crew to Select for recruiting.  <0 = Select 1 from Discard Pile
Option Explicit
Public payTotal As Integer, selectCrew As Integer, crewcnt As Integer, noMoralDisgruntle As Boolean, costLimit As Integer, SupplyID As Integer, crewFilter As String, lastCrewID

Private Sub chk_Click(Index As Integer)
Dim selected As String
   selected = getSelected()
   refreshShip
   restoreSelected selected
   updatePay
End Sub

Private Sub cmd_Click()
Dim Index As Integer, imposter As Integer

   With sftTree
      crewcnt = 0
      payTotal = 0
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = "R" Then
            payTotal = payTotal + .CellItemData(Index, 8)
            crewcnt = crewcnt + 1
         ElseIf .ItemDataString(Index) = "O" And selectCrew = 0 Then
            'check was actually expecting pay. eg. River = $0
            If .CellItemData(Index, 8) > 0 Then doDisgruntled player.ID, 2, .CellItemData(Index, 1)
         End If
      Next Index
   
   playsnd 8
   If selectCrew <> 0 Then
      If (selectCrew > 0 And crewcnt <= selectCrew And payTotal <= costLimit) Or (selectCrew = -1 And crewcnt <= 1) Then 'valid selection
         For Index = 0 To .ListCount - 1
            If .ItemDataString(Index) = "R" Then
            
               If .ItemData(Index) = 28 Then 'If Saffron is hired by anyone, Remove the existing imposter from Play
                  If haveCrewAnyone(54) Then
                     doDiscardCrew 100
                     imposter = 54
                  ElseIf haveCrewAnyone(41) Then
                     doDiscardCrew 70
                     imposter = 41
                  End If
               End If
               
               If .ItemData(Index) = 100 Then 'If Bridgit is hired by anyone, Remove the existing imposter from Play
                  If haveCrewAnyone(23) Then
                     doDiscardCrew 28
                     imposter = 23
                  ElseIf haveCrewAnyone(41) Then
                     doDiscardCrew 70
                     imposter = 41
                  End If
               End If
               
               If .ItemData(Index) = 70 Then 'If Yolonda is hired by anyone, Remove the existing imposter from Play
                  If haveCrewAnyone(54) Then
                     doDiscardCrew 100
                     imposter = 54
                  ElseIf haveCrewAnyone(23) Then
                     doDiscardCrew 28
                     imposter = 23
                  End If
               End If
               
               lastCrewID = .CellItemData(Index, 1)
               DB.Execute "UPDATE SupplyDeck SET Seq =" & player.ID & " WHERE CardID = " & .ItemData(Index)
               'add the card to the players deck
               DB.Execute "INSERT INTO PlayerSupplies (PlayerID, CardID) VALUES (" & player.ID & ", " & .ItemData(Index) & ")"

               If imposter > 0 Then
                  PutMsg getCrewName(0, imposter) & " has turned up as " & getCrewName(.ItemData(Index)) & " on " & player.PlayName & "'s Ship", player.ID, Logic!Gamecntr, True, getCrewID(.ItemData(Index)), 0, 0, 0, 1
               End If

            End If
         Next Index
      
         Me.Hide
      ElseIf selectCrew > 0 Then
         MessBox "No more than " & selectCrew & " crew and less than $" & costLimit, "Choose wisely", "Ooops", "", getLeader()
      Else
         MessBox "No more than 1 crew", "Choose wisely", "Ooops", "", getLeader()
      End If
   Else
      Me.Hide
   End If
   End With
End Sub

Private Sub Form_Resize()
Dim x
  sftTree.Move sftTree.Left, sftTree.top, Abs(Me.Width - 240), Abs(Me.Height - sftTree.top - 1000)
  For x = 0 To 2
      chk(x).Move Abs(Me.ScaleWidth - (1200 * (x + 1)) - 2000), Abs(Me.ScaleHeight - 330), chk(x).Width, chk(x).Height
  Next x
  'chk(1).top = Me.ScaleHeight - 300
  'chk(2).top = Me.ScaleHeight - 300
  cmd.Move Abs(Me.ScaleWidth - 1500), Abs(Me.ScaleHeight - 390), cmd.Width, cmd.Height
  Label1.Move Label1.Left, Abs(Me.ScaleHeight - 330), Label1.Width, Label1.Height
  
End Sub

Private Sub sftTree_ItemClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
With sftTree

  If Button = constSftTreeLeftButton And (AreaType = constSftTreeItem Or AreaType = constSftTreeCellText) Then
         Select Case .ItemDataString(Index)
         Case "R"  'no pay

            .ItemDataString(Index) = "O"
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
            updatePay
            
         Case "O"  'pay
         
            .ItemDataString(Index) = "R"
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "R")
            updatePay
            
         End Select
      
   End If
   
End With

End Sub
Private Sub sftTree_ItemDblClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
Dim frmCrew As New frmCrewSel
   If Button = constSftTreeLeftButton And AreaType = constSftTreeCellText Then
      If sftTree.CellItemData(Index, 0) = 1 Then
         frmCrew.crewFilter = " WHERE CrewID =" & sftTree.CellItemData(Index, 1)
         frmCrew.Show 1
         Set frmCrew = Nothing
      End If
   End If
End Sub

Private Sub Form_Load()
    With sftTree
       Set .ItemPictureExpandable = AssetImages.Overlay("U", "U")
       Set .ItemPictureExpanded = AssetImages.Overlay("U", "D")
       Set .ItemPictureLeaf = AssetImages.Overlay("LN", "LN")
    
       .LeftButtonOnly = False
       .AutoRespond = True
       .ButtonStyle = buttonsSftTreeAll
       
       refreshShip

    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub updatePay()
Dim Index, pay As Integer
Dim totalfight As Integer, totaltech As Integer, totalnego As Integer

   payTotal = 0
   crewcnt = 0
   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = "R" Then
            pay = pay + .CellItemData(Index, 8)
            payTotal = payTotal + .CellItemData(Index, 8)
            totalfight = totalfight + .CellItemData(Index, 5)
            totaltech = totaltech + .CellItemData(Index, 6)
            totalnego = totalnego + .CellItemData(Index, 7)
            crewcnt = crewcnt + 1
            .CellText(0, 1) = "selected=" & CStr(crewcnt)
         End If
      Next Index
      .CellText(0, 5) = IIf(totalfight > 0, CStr(totalfight), "")
      .CellText(0, 6) = IIf(totaltech > 0, CStr(totaltech), "")
      .CellText(0, 7) = IIf(totalnego > 0, CStr(totalnego), "")
      .CellText(0, 8) = "$" & CStr(pay)
   End With
End Sub

Public Sub refreshShip()
Dim Index, SQL
Dim totalfight, totaltech, totalnego, totalpay, lastplayer
Dim rst2 As New ADODB.Recordset
    
With sftTree

      .Clear  'otherwise Append
      totalfight = 0
      totaltech = 0
      totalnego = 0
      totalpay = 0

      
      'CREW---------------------------------------------
      Index = .AddItem("Crew")
      lastplayer = Index
      'Display actual Crew Number and Capacity (6) with modifiers
      .CellText(Index, 2) = "Crew Cap: " & CStr(CrewCapacity(player.ID)) & " Crew: " & CStr(getCrewCount(player.ID))
      If getCrewCount(player.ID) = CStr(CrewCapacity(player.ID)) Then
         .CellForeColor(Index, 2) = QBColor(12)
      End If
      .ItemLevel(Index) = 0
      Select Case selectCrew
      Case Is <> 0
         chk(0).Visible = True
         chk(1).Visible = True
         chk(2).Visible = True
      
         SQL = "SELECT SupplyDeck.CardID, Crew.*, Perk.PerkDescription "
         SQL = SQL & "FROM PlayerSupplies RIGHT JOIN (Perk INNER JOIN (Crew INNER JOIN SupplyDeck ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID) ON PlayerSupplies.CardID = SupplyDeck.CardID "
         SQL = SQL & "WHERE Crew.Leader=0 AND PlayerSupplies.PlayerID Is Null" & IIf(selectCrew = -1, " AND SupplyDeck.Seq = 5", " AND Crew.CrewID <> 41 AND Crew.CrewID <> 54")
         If selectCrew = -1 And (hasCrew(player.ID, 23) Or hasCrew(player.ID, 41) Or hasCrew(player.ID, 54)) Then
             SQL = SQL & " AND Crew.CrewID <> 23 AND Crew.CrewID <> 41 AND Crew.CrewID <> 54"
         End If
         If crewFilter <> "" Then
            SQL = SQL & " AND Crew.CrewID NOT IN (" & crewFilter & ")"
         End If
         If hasCrew(player.ID, 69) Then
            SQL = SQL & " AND Crew.Companion = 0"
         End If
         If SupplyID > 0 Then
             SQL = SQL & " AND SupplyDeck.SupplyID=" & SupplyID
         End If
         'filters
         If chk(0).Value = 0 Then
            SQL = SQL & " AND Crew.Wanted = 0"
         End If
         If chk(1).Value = 0 Then
            SQL = SQL & " AND Crew.Moral = 0"
         End If
         If chk(2).Value = 0 Then
            SQL = SQL & " AND (Crew.Wanted > 0 OR Crew.Moral = 1)"
         End If
         
         SQL = SQL & " ORDER BY Pilot DESC, Mechanic DESC, Companion DESC, Soldier DESC, Merc DESC, HillFolk DESC, Medic DESC, Grifter DESC, CrewName"
      Case Else 'pay crew
         SQL = "SELECT PlayerSupplies.CardID, Crew.*, Perk.PerkDescription"
         SQL = SQL & " FROM Perk INNER JOIN (Crew INNER JOIN (PlayerSupplies INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON Crew.CrewID = SupplyDeck.CrewID) ON Perk.PerkID = Crew.PerkID "
         SQL = SQL & "WHERE Leader=0 AND PlayerSupplies.OffJob=0 AND Crew.Pay>0 AND PlayerSupplies.PlayerID=" & player.ID
         If noMoralDisgruntle Then SQL = SQL & " AND Crew.Moral = 0"
      End Select
      rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst2.EOF
          Index = .AddItem(CStr(rst2!CrewID))
          .ItemData(Index) = rst2!CardID
          If selectCrew <> 0 Then
            .ItemDataString(Index) = "O"
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "O")
          Else
            .ItemDataString(Index) = "R"
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "R")
         End If
         .CellItemData(Index, 0) = 1 'crew
         .CellItemData(Index, 1) = rst2!CrewID
         .ItemLevel(Index) = 1
         .CellText(Index, 1) = rst2!CrewName & "  -  " & rst2!CrewDescr
         .CellText(Index, 2) = rst2!PerkDescription
         
         .CellText(Index, 3) = Trim(IIf(rst2!Mechanic = 1, "Mechanic  ", "") & IIf(rst2!Pilot = 1, "Pilot  ", "") & IIf(rst2!Companion = 1, "Companion  ", "") & _
               IIf(rst2!Merc = 1, "Merc  ", "") & IIf(rst2!Soldier = 1, "Soldier  ", "") & IIf(rst2!HillFolk = 1, "HillFolk  ", "") & _
               IIf(rst2!Grifter = 1, "Grifter ", "") & IIf(rst2!Medic = 1, "Medic ", "") & IIf(rst2!Mudder = 1, "Mudder", ""))
         .CellForeColor(Index, 3) = 65280
         .CellText(Index, 4) = IIf(rst2!Moral = 1, "Moral ", "")
         If rst2!Moral = 1 Then
            .CellForeColor(Index, 4) = 0
            .CellBackColor(Index, 4) = &HC0FFC0
         End If
         
         .CellText(Index, 5) = IIf(rst2!fight > 0, CStr(rst2!fight), "")
         .CellForeColor(Index, 5) = 0
         If rst2!fight > 0 Then .CellBackColor(Index, 5) = 6052315
         totalfight = totalfight + rst2!fight
         .CellItemData(Index, 5) = rst2!fight
         
         .CellText(Index, 6) = IIf(rst2!tech > 0, CStr(rst2!tech), "")
         .CellForeColor(Index, 6) = 0
         If rst2!tech > 0 Then .CellBackColor(Index, 6) = 16382208
         totaltech = totaltech + rst2!tech
         .CellItemData(Index, 6) = rst2!tech
         
         If getPerkAttributeCrew(player.ID, "negotiate", rst2!CardID) > 0 And hasGearKeyword(player.ID, "FIREARM", rst2!CrewID) Then
             .CellText(Index, 7) = CStr(rst2!Negotiate + 1)
         Else
            .CellText(Index, 7) = IIf(rst2!Negotiate > 0, CStr(rst2!Negotiate), "")
         End If
         .CellForeColor(Index, 7) = 0
         If Val(.CellText(Index, 7)) > 0 Then .CellBackColor(Index, 7) = 5373777
         totalnego = totalnego + Val(.CellText(Index, 7))
         .CellItemData(Index, 7) = rst2!Negotiate
         
         .CellText(Index, 8) = IIf(rst2!leader = 1, "Leader ", "$" & CStr(rst2!pay))
         .CellItemData(Index, 8) = rst2!pay
         If rst2!leader = 0 Then
            .CellBackColor(Index, 8) = 8388736
            .CellForeColor(Index, 8) = 16777215
         End If
         totalpay = totalpay + rst2!pay
         
         .CellText(Index, 9) = IIf(rst2!wanted > 0, "Wanted", "") & IIf(rst2!Disgruntled > 0, IIf(rst2!wanted > 0, " / ", "") & "Disgruntled", "")
         .CellForeColor(Index, 9) = 0
         If rst2!wanted > 0 Then
            .CellBackColor(Index, 9) = &HC0C0FF
         ElseIf rst2!Disgruntled > 0 Then
            .CellBackColor(Index, 9) = 11468799
         End If
         
         rst2.MoveNext
      Wend
      rst2.Close
      'fill the heading totals
      .CellText(lastplayer, 5) = IIf(totalfight > 0, CStr(totalfight), "")
      .CellForeColor(lastplayer, 5) = 0
      If totalfight > 0 Then .CellBackColor(lastplayer, 5) = 6052315

      .CellText(lastplayer, 6) = IIf(totaltech > 0, CStr(totaltech), "")
      .CellForeColor(lastplayer, 6) = 0
      If totaltech > 0 Then .CellBackColor(lastplayer, 6) = 16382208

      .CellText(lastplayer, 7) = IIf(totalnego > 0, CStr(totalnego), "")
      .CellForeColor(lastplayer, 7) = 0
      If totalnego > 0 Then .CellBackColor(lastplayer, 7) = 5373777

      .CellText(lastplayer, 8) = "$" & CStr(totalpay)
      .CellBackColor(lastplayer, 8) = 128
      .CellForeColor(lastplayer, 8) = 16777215
      '--------------------------------------------------
   
 End With
   
End Sub

Private Function getSelected() As String
Dim Index As Long

   With sftTree
      For Index = 0 To .ListCount - 1
         If .ItemDataString(Index) = "R" Then
            getSelected = getSelected & IIf(getSelected = "", "", ",") & CStr(.ItemData(Index))
         End If
      Next Index
   End With
End Function


Private Sub restoreSelected(ByVal selected As String)
Dim Index As Long
Dim y, a() As String

   If selected = "" Then Exit Sub
   a = Split(selected, ",")
   
   With sftTree
      For Index = 0 To .ListCount - 1
         For y = LBound(a) To UBound(a)
            If a(y) = CStr(.ItemData(Index)) Then
               .ItemDataString(Index) = "R"
               Set .ItemPicture(Index) = AssetImages.Overlay("L", "R")
            End If
         Next y
      Next Index
   End With
End Sub
