VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmJobs 
   Caption         =   "Jobs"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   Icon            =   "frmJobs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   4575
   Begin SftTree.SftTree sftTree 
      Height          =   2325
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   4485
      _Version        =   262144
      _ExtentX        =   7911
      _ExtentY        =   4101
      _StockProps     =   237
      ForeColor       =   8833235
      BackColor       =   855618
      BorderStyle     =   1
      ItemPictureExpanded=   "frmJobs.frx":030A
      ItemPictureExpandable=   "frmJobs.frx":0326
      ItemPictureLeaf =   "frmJobs.frx":0342
      PlusMinusPictureExpanded=   "frmJobs.frx":035E
      PlusMinusPictureExpandable=   "frmJobs.frx":037A
      PlusMinusPictureLeaf=   "frmJobs.frx":0396
      ButtonPicture   =   "frmJobs.frx":03B2
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
      ColBmp0         =   "frmJobs.frx":03CE
      ColWidth1       =   200
      ColTitle1       =   "Contact / Job Details"
      ColBmp1         =   "frmJobs.frx":03EA
      ColWidth2       =   107
      ColTitle2       =   "Job Type / Planet"
      ColBmp2         =   "frmJobs.frx":0406
      ColWidth3       =   120
      ColTitle3       =   "Needs / System"
      ColBmp3         =   "frmJobs.frx":0422
      ColWidth4       =   41
      ColStyle4       =   10
      ColTitle4       =   "Pay"
      ColBmp4         =   "frmJobs.frx":043E
      ColWidth5       =   87
      ColTitle5       =   "Bonus"
      ColBmp5         =   "frmJobs.frx":045A
      ColWidth6       =   33
      ColStyle6       =   9
      ColTitle6       =   "Fight"
      ColBmp6         =   "frmJobs.frx":0476
      ColWidth7       =   33
      ColStyle7       =   9
      ColTitle7       =   "Tech"
      ColBmp7         =   "frmJobs.frx":0492
      ColWidth8       =   34
      ColStyle8       =   9
      ColTitle8       =   "Nego"
      ColBmp8         =   "frmJobs.frx":04AE
      MouseIcon       =   "frmJobs.frx":04CA
      ColHeaderBackColor=   0
      ColHeaderForeColor=   10937324
      ForeColor       =   8833235
      BackColor       =   855618
      SelectStyle     =   2
      NoFocusStyle    =   2
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmJobs.frx":04E6
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      OpenEnded       =   0   'False
      ColPict0        =   "frmJobs.frx":0502
      ColPict1        =   "frmJobs.frx":051E
      ColPict2        =   "frmJobs.frx":053A
      ColPict3        =   "frmJobs.frx":0556
      ColPict4        =   "frmJobs.frx":0572
      ColPict5        =   "frmJobs.frx":058E
      ColPict6        =   "frmJobs.frx":05AA
      ColPict7        =   "frmJobs.frx":05C6
      ColPict8        =   "frmJobs.frx":05E2
      BackgroundPicture=   "frmJobs.frx":05FE
      CharSearchMode  =   2
      ShowFocusRectangle=   0   'False
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1080
      Top             =   3030
   End
   Begin MSComctlLib.ImageList AssetImages 
      Left            =   2760
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":061A
            Key             =   "UN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":08AC
            Key             =   "ST"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":0B3E
            Key             =   "NT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":1790
            Key             =   "CS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":1FE2
            Key             =   "ZS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":2834
            Key             =   "L"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":3486
            Key             =   "U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":40D8
            Key             =   "SG"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":492A
            Key             =   "R"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":557C
            Key             =   "D"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobs.frx":61CE
            Key             =   "O"
         EndProperty
      EndProperty
   End
   Begin XDOCKFLOATLibCtl.FDPane FDPane1 
      Height          =   420
      Left            =   3870
      TabIndex        =   0
      Top             =   3030
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
      Begin VB.Menu mnuPopUp 
         Caption         =   "Discard"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public jobFilter As String

Private Sub Form_Load()
    With sftTree
       Set .ItemPictureExpandable = AssetImages.Overlay("D", "R")
       Set .ItemPictureExpanded = AssetImages.Overlay("D", "R")
       Set .ItemPictureLeaf = AssetImages.Overlay("UN", "O")
     
       .LeftButtonOnly = False
       .AutoRespond = True
       .ButtonStyle = buttonsSftTreeAll

    End With
End Sub

Public Sub RefreshJobs()
   'keep yours at the top
   RefreshJob " WHERE PlayerID = " & player.ID

   If jobFilter = "all" Then
      RefreshJob " WHERE Players.Name IS NOT NULL AND PlayerID <> " & player.ID, False
   End If
      
End Sub

Private Sub RefreshJob(filter, Optional ByVal doClear As Boolean = True)
Dim Index, SQL
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim SectorID, x
     
SQL = "SELECT Board.Zones, Players.* FROM (Board INNER JOIN Players ON Board.SectorID = Players.SectorID) "
SQL = SQL & filter
SQL = SQL & " ORDER BY PlayerID"
    
With sftTree
   If doClear Then .Clear  'otherwise Append
   'add the Player details
   rst3.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst3.EOF
      Index = .AddItem(CStr(rst3!playerID) & IIf(isOutlaw(rst3!playerID), " - outlaw", ""))
      .ItemLevel(Index) = 0
      .CellText(Index, 1) = rst3!ship & " - " & PlayCode(rst3!playerID).PlayName & IIf(rst3!playerID = player.ID, " [me]", "")
      For x = 0 To 8
         .CellForeColor(Index, x) = 0
         .CellBackColor(Index, x) = getPlayerColor(rst3!playerID)
      Next x
     Set .ItemPicture(Index) = AssetImages.Overlay("L", "U")
      
      SQL = "SELECT PlayerJobs.PlayerID, PlayerJobs.JobStatus, Contact.ContactName, Contact.Colour, Contact.Picture, JobType.JobTypeDescr, Profession.ProfessionName, ContactDeck.*, JobType_1.JobTypeDescr AS JobType2 "
      SQL = SQL & "FROM (Contact INNER JOIN (((PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) INNER JOIN JobType ON ContactDeck.JobTypeID = JobType.JobTypeID) "
      SQL = SQL & "LEFT JOIN Profession ON ContactDeck.ProfessionID = Profession.ProfessionID) ON Contact.ContactID = ContactDeck.ContactID) INNER JOIN JobType AS JobType_1 ON ContactDeck.JobType2D = JobType_1.JobTypeID "
      SQL = SQL & " WHERE PlayerJobs.JobStatus < " & JOB_SUCCESS & " AND PlayerJobs.PlayerID=" & rst3!playerID
      
      If player.ID <> rst3!playerID Then 'hide inactives
         SQL = SQL & " AND PlayerJobs.JobStatus IN (1,2)"
      End If
      SQL = SQL & " ORDER BY Contact.ContactName,PlayerJobs.CardID"
      
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         Index = .AddItem(CStr(rst!CardID))
         .ItemData(Index) = rst!CardID
         .CellItemData(Index, 0) = rst!JobStatus
         .CellText(Index, 1) = rst!ContactName & " - " & rst!JobName
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
          Set .ItemPicture(Index) = LoadPicture(App.Path & "\Pictures\Sm" & rst!Picture)
'         If (rst!JobStatus = 1 Or rst!JobStatus = 2) Then
'            Set .ItemPicture(Index) = AssetImages.Overlay("L", "D")
'         Else
'            Set .ItemPicture(Index) = AssetImages.Overlay("L", "U")
'         End If
         SectorID = varDLookup("SectorID", "Players", "PlayerID=" & rst!playerID)
         .ItemLevel(Index) = 1
         
         If rst!Job1ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job1ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
                Index = .AddItem(CStr(rst2!JobID))
               .CellText(Index, 1) = rst2!JobDesc
               x = getSectorCount(getPlayerSector(rst3!playerID), rst2!SectorID)
               .CellText(Index, 2) = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "")
               .ItemData(Index) = rst!playerID
               If (rst2!SectorID = 1 And getCruiserSector() = SectorID) Or (rst2!SectorID = 2 And getCorvetteSector() = SectorID) Or (rst2!SectorID > 2 And SectorID = rst2!SectorID) Then
                  .CellFont(Index, 2).Bold = True
                  .CellFont(Index, 3).Bold = True
                  If hasJobReqs(rst3!playerID, rst!CardID, rst!Job1ID) Then
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
               .ItemLevel(Index) = 2
               If (rst!JobStatus = 1 Or rst!JobStatus = 2) Then
                  Set .ItemPicture(Index) = AssetImages.Overlay("L", "R")
               Else
                  Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
                  .CellItemData(Index, 1) = rst2!SectorID
               End If
         
               '.CellText(index, 3) = rst!
            End If
            rst2.Close
         End If
         
         'Bonus Drop Job
         If rst!Job3ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job3ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
                Index = .AddItem(CStr(rst2!JobID))
               .CellText(Index, 1) = rst2!JobDesc
               x = getSectorCount(getPlayerSector(rst3!playerID), rst2!SectorID)
               .CellText(Index, 2) = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "")
               .ItemData(Index) = rst!playerID
               If (rst2!SectorID = 1 And getCruiserSector() = SectorID) Or (rst2!SectorID = 2 And getCorvetteSector() = SectorID) Or (rst2!SectorID > 2 And SectorID = rst2!SectorID) Then
                  .CellFont(Index, 2).Bold = True
                  .CellFont(Index, 3).Bold = True
                  If hasJobReqs(rst3!playerID, rst!CardID, rst!Job3ID) Then
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
               .ItemLevel(Index) = 3
               If rst!JobStatus = 2 Then
                  Set .ItemPicture(Index) = AssetImages.Overlay("L", "R")
               Else
                  Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
                  .CellItemData(Index, 1) = rst2!SectorID
               End If
         
               '.CellText(index, 3) = rst!
            End If
            rst2.Close
         End If
         
         If rst!Job2ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job2ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
                Index = .AddItem(CStr(rst2!JobID))
               .CellText(Index, 1) = rst2!JobDesc
               x = getSectorCount(getPlayerSector(rst3!playerID), rst2!SectorID)
               .CellText(Index, 2) = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "")
               .ItemData(Index) = rst!playerID
               If (rst2!SectorID = 1 And getCruiserSector() = SectorID) Or (rst2!SectorID = 2 And getCorvetteSector() = SectorID) Or (rst2!SectorID > 2 And SectorID = rst2!SectorID) Then
                  .CellFont(Index, 2).Bold = True
                  .CellFont(Index, 3).Bold = True
                  If hasJobReqs(rst3!playerID, rst!CardID, rst!Job2ID) Then
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
               .ItemLevel(Index) = 2
               Set .ItemPicture(Index) = AssetImages.Overlay("L", "UN")
               .CellItemData(Index, 1) = rst2!SectorID
            End If
            rst2.Close
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
  'Timer1.Enabled = True

End Sub


Private Sub mnuPopUp_Click(Index As Integer)
 With sftTree
   Select Case Index
   Case 0
      If .ListIndex < 1 Or Left(.CellText(.ListIndex, 1), 4) = "Goal" Then Exit Sub
      If MessBox("Are you sure you want to ditch the Job: " & .CellText(.ListIndex, 1) & "?", "Discard Job", "Ditch", "Nope", getLeader()) = 0 Then
      'If MsgBox("Are you sure you want to ditch the Job: " & .CellText(.ListIndex, 1) & "?", vbOKCancel + vbQuestion, "Discard Job") = vbOK Then
         removeJob player.ID, .ItemData(.ListIndex)
         If actionSeq > ASidle And actionSeq < ASEnd Then
            Main.showActions
         End If
         RefreshJobs
         Main.drawLine 0, -1
      End If
   End Select
 End With
End Sub

Private Sub sftTree_ItemClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
   Select Case AreaType
   Case 9
      If Button = 1 Then
         With sftTree
            If sftTree.CellItemData(Index, 1) > 0 And sftTree.ItemData(Index) = player.ID Then
               'draw a line
               Main.drawLine 0, sftTree.CellItemData(Index, 1), varDLookup("SectorID", "Players", "PlayerID=" & player.ID), False
            Else
               Main.drawLine 0, -1
            End If
            
        End With
      ElseIf Button = 2 Then
         With sftTree
            mnuPopup(0).Enabled = (.CellItemData(Index, 0) = 0 And .ItemData(Index) > 0 And .ItemLevel(Index) = 1)
            If (.CellItemData(Index, 0) = 0 And .ItemData(Index) > 0 And .ItemLevel(Index) = 1) Then PopupMenu mnuPop
            
        End With
      End If
   End Select
End Sub

Private Sub sftTree_ItemDblClick(ByVal Index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
Dim frmJobEdit As frmJobEditor
   If Button = constSftTreeLeftButton And AreaType = constSftTreeCellText And sftTree.ItemLevel(Index) = 1 Then
      Set frmJobEdit = New frmJobEditor
      frmJobEdit.lockEdits = True
      frmJobEdit.JobCardID = sftTree.ItemData(Index)
      frmJobEdit.Show 1
   End If
End Sub

Private Sub Timer1_Timer()
   If FDPane1.PaneVisible Then RefreshJobs
End Sub
