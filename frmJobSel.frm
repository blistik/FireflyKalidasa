VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form frmJobSel 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Job Selection"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SftTree.SftTree sftTree 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   14355
      _Version        =   262144
      _ExtentX        =   25321
      _ExtentY        =   8493
      _StockProps     =   237
      ForeColor       =   8833235
      BackColor       =   855618
      BorderStyle     =   1
      ItemPictureExpanded=   "frmJobSel.frx":0000
      ItemPictureExpandable=   "frmJobSel.frx":001C
      ItemPictureLeaf =   "frmJobSel.frx":0038
      PlusMinusPictureExpanded=   "frmJobSel.frx":0054
      PlusMinusPictureExpandable=   "frmJobSel.frx":0070
      PlusMinusPictureLeaf=   "frmJobSel.frx":008C
      ButtonPicture   =   "frmJobSel.frx":00A8
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
      ColWidth0       =   16
      ColTitle0       =   "Card ID"
      ColBmp0         =   "frmJobSel.frx":00C4
      ColWidth1       =   267
      ColTitle1       =   "Contact / Job Details"
      ColBmp1         =   "frmJobSel.frx":00E0
      ColWidth2       =   107
      ColTitle2       =   "Job Type / Planet"
      ColBmp2         =   "frmJobSel.frx":00FC
      ColWidth3       =   267
      ColTitle3       =   "Needs / System"
      ColBmp3         =   "frmJobSel.frx":0118
      ColWidth4       =   41
      ColStyle4       =   10
      ColTitle4       =   "Pay"
      ColBmp4         =   "frmJobSel.frx":0134
      ColWidth5       =   87
      ColTitle5       =   "Bonus"
      ColBmp5         =   "frmJobSel.frx":0150
      ColWidth6       =   33
      ColStyle6       =   9
      ColTitle6       =   "Fight"
      ColBmp6         =   "frmJobSel.frx":016C
      ColWidth7       =   33
      ColStyle7       =   9
      ColTitle7       =   "Tech"
      ColBmp7         =   "frmJobSel.frx":0188
      ColWidth8       =   34
      ColStyle8       =   9
      ColTitle8       =   "Nego"
      ColBmp8         =   "frmJobSel.frx":01A4
      MouseIcon       =   "frmJobSel.frx":01C0
      ColHeaderBackColor=   0
      ColHeaderForeColor=   10937324
      ForeColor       =   8833235
      BackColor       =   855618
      SelectStyle     =   2
      NoFocusStyle    =   2
      RowColHeaderAppearance=   0
      RowColPicture   =   "frmJobSel.frx":01DC
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      OpenEnded       =   0   'False
      ColPict0        =   "frmJobSel.frx":01F8
      ColPict1        =   "frmJobSel.frx":0214
      ColPict2        =   "frmJobSel.frx":0230
      ColPict3        =   "frmJobSel.frx":024C
      ColPict4        =   "frmJobSel.frx":0268
      ColPict5        =   "frmJobSel.frx":0284
      ColPict6        =   "frmJobSel.frx":02A0
      ColPict7        =   "frmJobSel.frx":02BC
      ColPict8        =   "frmJobSel.frx":02D8
      BackgroundPicture=   "frmJobSel.frx":02F4
      CharSearchMode  =   2
      ShowFocusRectangle=   0   'False
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Keep"
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
      Left            =   11190
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   50
      Width           =   915
   End
   Begin MSComctlLib.ImageList AssetImages 
      Left            =   8610
      Top             =   -30
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
            Picture         =   "frmJobSel.frx":0310
            Key             =   "UN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":05A2
            Key             =   "ST"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":0834
            Key             =   "NT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":1486
            Key             =   "CS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":1CD8
            Key             =   "ZS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":252A
            Key             =   "L"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":317C
            Key             =   "U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":3DCE
            Key             =   "SG"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":4620
            Key             =   "R"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":5272
            Key             =   "D"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobSel.frx":5EC4
            Key             =   "O"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick 3 Jobs you want to keep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   210
      TabIndex        =   1
      Top             =   90
      Width           =   4635
   End
End
Attribute VB_Name = "frmJobSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CardID As Integer, jobFilter As String, maxjobs As Integer

Private Sub cmd_Click()
Dim index As Long, y As Long
   For index = 0 To sftTree.ListCount - 1
      Select Case sftTree.ItemDataString(index)
      Case "O"
         removeJob player.ID, sftTree.ItemData(index)
      Case "R"
         y = y + 1
      End Select
   Next index
   If y > maxjobs Then
      MessBox "You need to have no more than " & maxjobs, "Too many Jobs", "Ooops", "", getLeader()
      refreshJobs
   Else
      Me.hide
   End If

End Sub


'Private Sub jobView()
'Dim frmJobEdit As frmJobEditor
'   If GetCombo(cbo) > 0 Then
'      Set frmJobEdit = New frmJobEditor
'      frmJobEdit.lockEdits = True
'      frmJobEdit.JobCardID = GetCombo(cbo)
'      frmJobEdit.Show 1
'   End If
'End Sub

Private Sub Form_Load()
    With sftTree
       Set .ItemPictureExpandable = AssetImages.Overlay("D", "R")
       Set .ItemPictureExpanded = AssetImages.Overlay("D", "R")
       Set .ItemPictureLeaf = AssetImages.Overlay("UN", "O")
     
       .LeftButtonOnly = False
       .AutoRespond = True
       .ButtonStyle = buttonsSftTreeAll

    End With
   refreshJobs
End Sub

Private Sub Form_Resize()
   sftTree.Move sftTree.Left, sftTree.top, Abs(Me.Width - 260), Abs(Me.Height - sftTree.top - 580)
   cmd.Left = Abs(Me.Width - 1300)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

'Private Sub refreshCbo()
'Dim rst As New ADODB.Recordset
'Dim SQL
'   SQL = "SELECT Contact.ContactName, ContactDeck.CardID, ContactDeck.Pay, ContactDeck.Bonus, ContactDeck.Keywords, ContactDeck.Immoral, ContactDeck.JobName, "
'   SQL = SQL & "Job.JobID,  Job.JobDesc, Job_1.JobDesc AS Job2Desc, Job_1.JobID AS Job2 "
'   SQL = SQL & "FROM (Contact INNER JOIN (Job INNER JOIN (PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) ON Job.JobID = ContactDeck.Job1ID) ON Contact.ContactID = ContactDeck.ContactID) LEFT JOIN Job AS Job_1 ON ContactDeck.Job2ID = Job_1.JobID "
'   SQL = SQL & "WHERE PlayerJobs.PlayerID=" & player.ID & " AND PlayerJobs.JobStatus " & jobFilter
'
'   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
'   While Not rst.EOF
'      cbo.AddItem rst!ContactName & ": " & rst!JobName & " $" & rst!pay & " " & IIf(rst!bonus > 0, "+$" & rst!bonus & " Bonus. ", "") & getPlanetSector(rst!JobID) & ": " & rst!JobDesc & " " & IIf(IsNull(rst!Job2), "", " -/- " & getPlanetSector(rst!Job2) & ": " & rst!Job2Desc) & IIf(IsNull(rst!KeyWords), "", " (" & rst!KeyWords & ")")
'      cbo.ItemData(cbo.NewIndex) = rst!CardID
'      rst.MoveNext
'   Wend
'   rst.Close
'   Set rst = Nothing
'End Sub

Private Sub refreshJobs()
Dim index, SQL
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim SectorID, x
   With sftTree
      .Clear
      SQL = "SELECT PlayerJobs.PlayerID, PlayerJobs.JobStatus, Contact.ContactName, Contact.Colour, Contact.Picture, JobType.JobTypeDescr, Profession.ProfessionName, ContactDeck.*, JobType_1.JobTypeDescr AS JobType2 "
      SQL = SQL & "FROM (Contact INNER JOIN (((PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) INNER JOIN JobType ON ContactDeck.JobTypeID = JobType.JobTypeID) "
      SQL = SQL & "LEFT JOIN Profession ON ContactDeck.ProfessionID = Profession.ProfessionID) ON Contact.ContactID = ContactDeck.ContactID) INNER JOIN JobType AS JobType_1 ON ContactDeck.JobType2D = JobType_1.JobTypeID "
      SQL = SQL & " WHERE ContactDeck.ContactID <> 10 and ContactDeck.ContactID <> 0 and PlayerJobs.JobStatus " & jobFilter & " AND PlayerJobs.PlayerID=" & player.ID
      
      SQL = SQL & " ORDER BY Contact.ContactName,PlayerJobs.CardID"
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      While Not rst.EOF
         index = .AddItem(CStr(rst!CardID))
         .ItemData(index) = rst!CardID
         '.CellItemData(Index, 0) = rst!JobStatus
         Set .ItemPicture(index) = AssetImages.Overlay("L", "O")
         .ItemDataString(index) = "O"
         .CellItemData(index, 2) = rst!ContactID
         .CellText(index, 1) = rst!ContactName & " - " & rst!JobName
         .CellForeColor(index, 1) = 0
         .CellBackColor(index, 1) = rst!Colour
         .CellText(index, 2) = rst!JobTypeDescr & IIf(rst!JobType2 <> "-", "/" & rst!JobType2, "") & IIf(rst!illegal = 1, "/illegal", "") & IIf(rst!Immoral = 1, "/immoral", "")
         If rst!illegal = 1 Or rst!Immoral Then
            .CellBackColor(index, 2) = 3355647
         End If
         .CellText(index, 3) = Nz(rst!JobOrder)
         .CellForeColor(index, 3) = 51712
         .CellText(index, 4) = "$" & rst!pay
         .CellBackColor(index, 4) = 8388736
         .CellForeColor(index, 4) = 16777215
         .CellText(index, 5) = IIf(rst!BonusPart > 0, " +" & rst!BonusPart & " part: ", "") & IIf(rst!bonus > 0, " +$" & rst!bonus & ":", "") & IIf(rst!KeywordBonus = 1, rst!KeyWords, "") & IIf(IsNull(rst!ProfessionName), "", " " & rst!ProfessionName) & IIf(rst!BonusPerSkill > 0, " /" & cstrSkill(rst!BonusPerSkill), "") & IIf(rst!Job3ID > 0, "Bonus Job", "")
         If rst!BonusPart > 0 Or rst!bonus > 0 Then
            .CellForeColor(index, 5) = 0
            .CellBackColor(index, 5) = 1900316
         End If
         .CellText(index, 6) = IIf(rst!fight > 0, CStr(rst!fight), "")
         .CellForeColor(index, 6) = 0
         If rst!fight > 0 Then .CellBackColor(index, 6) = 6052315
         .CellText(index, 7) = IIf(rst!tech > 0, CStr(rst!tech), "")
         .CellForeColor(index, 7) = 0
         If rst!tech > 0 Then .CellBackColor(index, 7) = 16382208
         .CellText(index, 8) = IIf(rst!Negotiate > 0, CStr(rst!Negotiate), "")
         .CellForeColor(index, 8) = 0
         If rst!Negotiate > 0 Then .CellBackColor(index, 8) = 5373777
         'Set .ItemPicture(index) = LoadPicture(App.Path & "\Pictures\Sm" & rst!Picture)
'         If (rst!JobStatus = 1 Or rst!JobStatus = 2) Then
'            Set .ItemPicture(Index) = AssetImages.Overlay("L", "D")
'         Else
'            Set .ItemPicture(Index) = AssetImages.Overlay("L", "U")
'         End If
         SectorID = varDLookup("SectorID", "Players", "PlayerID=" & rst!playerID)
         .ItemLevel(index) = 1
         
         If rst!Job1ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job1ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
                index = .AddItem(CStr(rst2!JobID))
               .CellText(index, 1) = rst2!JobDesc
               x = getSectorCount(getPlayerSector(player.ID), rst2!SectorID)
               .CellText(index, 2) = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "")
               .ItemData(index) = rst!playerID
               If (rst2!SectorID = 1 And getCruiserSector() = SectorID) Or (rst2!SectorID = 2 And getCorvetteSector() = SectorID) Or (rst2!SectorID > 2 And SectorID = rst2!SectorID) Then
                  .CellFont(index, 2).Bold = True
                  .CellFont(index, 3).Bold = True
                  If hasJobReqs(player.ID, rst!CardID, rst!Job1ID) Then
                     .CellForeColor(index, 2) = 0
                     .CellForeColor(index, 3) = 0
                  Else
                     .CellForeColor(index, 2) = 255
                     .CellForeColor(index, 3) = 255
                  End If
                  .CellBackColor(index, 2) = &HC0FFC0
                  
                  .CellBackColor(index, 3) = &HC0FFC0
                  
               End If
               .CellText(index, 3) = rst2!System
               .ItemLevel(index) = 2

               Set .ItemPicture(index) = AssetImages.Overlay("L", "UN")
               .CellItemData(index, 1) = rst2!SectorID

         
               '.CellText(index, 3) = rst!
            End If
            rst2.Close
         End If
         
         'Bonus Drop Job
         If rst!Job3ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job3ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
                index = .AddItem(CStr(rst2!JobID))
               .CellText(index, 1) = rst2!JobDesc
               x = getSectorCount(getPlayerSector(player.ID), rst2!SectorID)
               .CellText(index, 2) = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "")
               .ItemData(index) = rst!playerID
               If (rst2!SectorID = 1 And getCruiserSector() = SectorID) Or (rst2!SectorID = 2 And getCorvetteSector() = SectorID) Or (rst2!SectorID > 2 And SectorID = rst2!SectorID) Then
                  .CellFont(index, 2).Bold = True
                  .CellFont(index, 3).Bold = True
                  If hasJobReqs(player.ID, rst!CardID, rst!Job3ID) Then
                     .CellForeColor(index, 2) = 0
                     .CellForeColor(index, 3) = 0
                  Else
                     .CellForeColor(index, 2) = 255
                     .CellForeColor(index, 3) = 255
                  End If
                  .CellBackColor(index, 2) = &HC0FFC0
                  
                  .CellBackColor(index, 3) = &HC0FFC0
               End If
               .CellText(index, 3) = rst2!System
               .ItemLevel(index) = 3

               Set .ItemPicture(index) = AssetImages.Overlay("L", "UN")

         
               '.CellText(index, 3) = rst!
            End If
            rst2.Close
         End If
         
         If rst!Job2ID > 0 Then
            SQL = "SELECT Planet.PlanetName, Planet.System, Job.* FROM Job INNER JOIN Planet ON Job.SectorID = Planet.SectorID WHERE JobID =" & rst!Job2ID
            rst2.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
            If Not rst2.EOF Then
                index = .AddItem(CStr(rst2!JobID))
               .CellText(index, 1) = rst2!JobDesc
               x = getSectorCount(getPlayerSector(player.ID), rst2!SectorID)
               .CellText(index, 2) = rst2!PlanetName & IIf(x > 0, "  (" & x & ")", "")
               .ItemData(index) = rst!playerID
               If (rst2!SectorID = 1 And getCruiserSector() = SectorID) Or (rst2!SectorID = 2 And getCorvetteSector() = SectorID) Or (rst2!SectorID > 2 And SectorID = rst2!SectorID) Then
                  .CellFont(index, 2).Bold = True
                  .CellFont(index, 3).Bold = True
                  If hasJobReqs(player.ID, rst!CardID, rst!Job2ID) Then
                     .CellForeColor(index, 2) = 0
                     .CellForeColor(index, 3) = 0
                  Else
                     .CellForeColor(index, 2) = 255
                     .CellForeColor(index, 3) = 255
                  End If
                  .CellBackColor(index, 2) = &HC0FFC0
                  .CellBackColor(index, 3) = &HC0FFC0
               End If
               .CellText(index, 3) = rst2!System
               .ItemLevel(index) = 2
               Set .ItemPicture(index) = AssetImages.Overlay("L", "UN")
               .CellItemData(index, 1) = rst2!SectorID
            End If
            rst2.Close
         End If
         
         rst.MoveNext
      Wend
   End With
End Sub


Private Sub sftTree_ItemClick(ByVal index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
Dim max As Integer, maxConsider
With sftTree

  If Button = constSftTreeLeftButton And (AreaType = constSftTreeItem Or AreaType = constSftTreeCellText) Then
         Select Case .ItemDataString(index)
         Case "O" 'select

            'determine how many cards can be accepted

            If getSelected("R") < maxjobs Then  'accept only up to 3(or 4) cards
               .ItemDataString(index) = "R"
               Set .ItemPicture(index) = AssetImages.Overlay("L", "R")
            Else
               playsnd 9
            End If

            
         Case "R"  'de-select
            .ItemDataString(index) = "O"
            Set .ItemPicture(index) = AssetImages.Overlay("L", "O")
            
            
         End Select
      
   End If
   
End With

End Sub

Private Sub sftTree_ItemDblClick(ByVal index As Long, ByVal ColNum As Integer, ByVal AreaType As Integer, ByVal Button As Integer, ByVal Shift As Integer)
Dim frmJobEdit As frmJobEditor
   If Button = constSftTreeLeftButton And AreaType = constSftTreeCellText And sftTree.ItemLevel(index) = 1 Then
      Set frmJobEdit = New frmJobEditor
      frmJobEdit.lockEdits = True
      frmJobEdit.JobCardID = sftTree.ItemData(index)
      frmJobEdit.Show 1
   End If
End Sub

Private Function getSelected(ByVal status As String) As Integer
Dim index As Integer
   With sftTree
      For index = 0 To .ListCount - 1
         If .ItemDataString(index) = status Then
            getSelected = getSelected + 1
         End If
      Next index
   
   End With

End Function
