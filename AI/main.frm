VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form Main 
   BackColor       =   &H00000000&
   Caption         =   "Firefly AI Bot"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin SftTree.SftTree sftTree 
      Height          =   4125
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14115
      _Version        =   262144
      _ExtentX        =   24897
      _ExtentY        =   7276
      _StockProps     =   237
      ForeColor       =   8833235
      BackColor       =   4587520
      BorderStyle     =   1
      ItemPictureExpanded=   "main.frx":0442
      ItemPictureExpandable=   "main.frx":045E
      ItemPictureLeaf =   "main.frx":047A
      PlusMinusPictureExpanded=   "main.frx":0496
      PlusMinusPictureExpandable=   "main.frx":04B2
      PlusMinusPictureLeaf=   "main.frx":04CE
      ButtonPicture   =   "main.frx":04EA
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
      ColBmp0         =   "main.frx":0506
      ColWidth1       =   167
      ColTitle1       =   "Names & Titles"
      ColBmp1         =   "main.frx":0522
      ColWidth2       =   227
      ColTitle2       =   "Perks & Quirks"
      ColBmp2         =   "main.frx":053E
      ColWidth3       =   67
      ColTitle3       =   "Ability"
      ColBmp3         =   "main.frx":055A
      ColWidth4       =   77
      ColStyle4       =   9
      ColTitle4       =   "Status"
      ColBmp4         =   "main.frx":0576
      ColWidth5       =   33
      ColStyle5       =   9
      ColTitle5       =   "Fight"
      ColBmp5         =   "main.frx":0592
      ColWidth6       =   33
      ColStyle6       =   9
      ColTitle6       =   "Tech"
      ColBmp6         =   "main.frx":05AE
      ColWidth7       =   37
      ColStyle7       =   9
      ColTitle7       =   "Nego"
      ColBmp7         =   "main.frx":05CA
      ColWidth8       =   47
      ColStyle8       =   10
      ColTitle8       =   "Pay/job"
      ColBmp8         =   "main.frx":05E6
      ColWidth9       =   200
      ColTitle9       =   "Keywords"
      ColBmp9         =   "main.frx":0602
      MouseIcon       =   "main.frx":061E
      ColHeaderBackColor=   0
      ColHeaderForeColor=   65280
      ForeColor       =   8833235
      BackColor       =   4587520
      SelectStyle     =   2
      RowColHeaderAppearance=   0
      RowColPicture   =   "main.frx":063A
      LeftButtonOnly  =   0   'False
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      OpenEnded       =   0   'False
      ColFlag0        =   4
      ColPict0        =   "main.frx":0656
      ColFlag1        =   8
      ColPict1        =   "main.frx":0672
      ColFlag2        =   4
      ColPict2        =   "main.frx":068E
      ColFlag3        =   12
      ColPict3        =   "main.frx":06AA
      ColFlag4        =   8
      ColPict4        =   "main.frx":06C6
      ColFlag5        =   8
      ColPict5        =   "main.frx":06E2
      ColFlag6        =   8
      ColPict6        =   "main.frx":06FE
      ColFlag7        =   8
      ColPict7        =   "main.frx":071A
      ColFlag8        =   8
      ColPict8        =   "main.frx":0736
      ColFlag9        =   8
      ColPict9        =   "main.frx":0752
      BackgroundPicture=   "main.frx":076E
      ShowFocusRectangle=   0   'False
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin SftTree.SftTree sftTree2 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   4170
      Width           =   14145
      _Version        =   262144
      _ExtentX        =   24950
      _ExtentY        =   2566
      _StockProps     =   237
      ForeColor       =   8833235
      BackColor       =   855618
      BorderStyle     =   1
      ItemPictureExpanded=   "main.frx":078A
      ItemPictureExpandable=   "main.frx":07A6
      ItemPictureLeaf =   "main.frx":07C2
      PlusMinusPictureExpanded=   "main.frx":07DE
      PlusMinusPictureExpandable=   "main.frx":07FA
      PlusMinusPictureLeaf=   "main.frx":0816
      ButtonPicture   =   "main.frx":0832
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
      ColBmp0         =   "main.frx":084E
      ColWidth1       =   200
      ColTitle1       =   "Contact / Job Details"
      ColBmp1         =   "main.frx":086A
      ColWidth2       =   213
      ColTitle2       =   "Job Type / Planet"
      ColBmp2         =   "main.frx":0886
      ColWidth3       =   120
      ColTitle3       =   "Needs / System"
      ColBmp3         =   "main.frx":08A2
      ColWidth4       =   41
      ColStyle4       =   10
      ColTitle4       =   "Pay"
      ColBmp4         =   "main.frx":08BE
      ColWidth5       =   87
      ColTitle5       =   "Bonus"
      ColBmp5         =   "main.frx":08DA
      ColWidth6       =   33
      ColStyle6       =   9
      ColTitle6       =   "Fight"
      ColBmp6         =   "main.frx":08F6
      ColWidth7       =   33
      ColStyle7       =   9
      ColTitle7       =   "Tech"
      ColBmp7         =   "main.frx":0912
      ColWidth8       =   34
      ColStyle8       =   9
      ColTitle8       =   "Nego"
      ColBmp8         =   "main.frx":092E
      MouseIcon       =   "main.frx":094A
      ColHeaderBackColor=   0
      ColHeaderForeColor=   10937324
      ForeColor       =   8833235
      BackColor       =   855618
      RowColHeaderAppearance=   0
      RowColPicture   =   "main.frx":0966
      RowHeaderStyle  =   128
      RowHeaderAppearance=   0
      OpenEnded       =   0   'False
      ColPict0        =   "main.frx":0982
      ColPict1        =   "main.frx":099E
      ColPict2        =   "main.frx":09BA
      ColPict3        =   "main.frx":09D6
      ColPict4        =   "main.frx":09F2
      ColPict5        =   "main.frx":0A0E
      ColPict6        =   "main.frx":0A2A
      ColPict7        =   "main.frx":0A46
      ColPict8        =   "main.frx":0A62
      BackgroundPicture=   "main.frx":0A7E
      ToolTipForeColor=   -2147483640
      ToolTipBackColor=   -2147483643
   End
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25506
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timing 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   2340
      Top             =   0
   End
   Begin MSComctlLib.ImageList AssetImages 
      Left            =   9510
      Top             =   4710
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
            Picture         =   "main.frx":0A9A
            Key             =   "UN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0D2C
            Key             =   "ST"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0FBE
            Key             =   "NT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1C10
            Key             =   "CS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2462
            Key             =   "ZS"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2CB4
            Key             =   "L"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3906
            Key             =   "U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4558
            Key             =   "SG"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4DAA
            Key             =   "R"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":59FC
            Key             =   "D"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":664E
            Key             =   "O"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":72A0
            Key             =   "P"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":73FA
            Key             =   "PS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":7714
            Key             =   "LN"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":7A2E
            Key             =   "CN"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":7D48
            Key             =   "GR"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":8062
            Key             =   "UP"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":84B4
            Key             =   "LD"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":87CE
            Key             =   "SU"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":8C20
            Key             =   "MA"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":9072
            Key             =   "dis"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":93C4
            Key             =   "serenity"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":9716
            Key             =   "dino"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":9A68
            Key             =   "fight"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":9DBA
            Key             =   "tech"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":A10C
            Key             =   "negot"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSolid 
      BackStyle       =   0  'Transparent
      Caption         =   "Solid"
      BeginProperty Font 
         Name            =   "Cyberpunk Is Not Dead"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   14140
      TabIndex        =   3
      Top             =   30
      Width           =   795
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
      Height          =   240
      Index           =   9
      Left            =   14370
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Effects         =   "main.frx":A45E
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
      Height          =   240
      Index           =   8
      Left            =   14370
      Top             =   3200
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Effects         =   "main.frx":A476
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
      Height          =   240
      Index           =   7
      Left            =   14370
      Top             =   2800
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Effects         =   "main.frx":A48E
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
      Height          =   240
      Index           =   6
      Left            =   14370
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Effects         =   "main.frx":A4A6
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
      Height          =   240
      Index           =   5
      Left            =   14370
      Top             =   2000
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Effects         =   "main.frx":A4BE
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
      Height          =   240
      Index           =   4
      Left            =   14370
      Top             =   1600
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Effects         =   "main.frx":A4D6
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
      Height          =   240
      Index           =   3
      Left            =   14370
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Effects         =   "main.frx":A4EE
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
      Height          =   240
      Index           =   2
      Left            =   14370
      Top             =   800
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Effects         =   "main.frx":A506
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
      Height          =   240
      Index           =   1
      Left            =   14370
      Top             =   400
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Effects         =   "main.frx":A51E
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim targetContact As Integer, targetJobCard, targetJobID, targetSector, targetSupplySector, bountyCardID As Integer
Public moseydone As Boolean, fullburndone As Boolean, buydone As Boolean, buyfueldone As Boolean, leader
Public dealdone As Boolean, workdone As Boolean, forcererollused As Boolean
Private Const MAXFUEL As Variant = 8

Private Sub Form_Load()
Dim x, inprog As Boolean

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
    
   With sftTree2
       Set .ItemPictureExpandable = AssetImages.Overlay("D", "R")
       Set .ItemPictureExpanded = AssetImages.Overlay("D", "R")
       Set .ItemPictureLeaf = AssetImages.Overlay("UN", "O")
     
       .LeftButtonOnly = False
       .AutoRespond = True
       .ButtonStyle = buttonsSftTreeAll

   End With
    
    
   ContactList = "1,2,4,5" 'default, to be reduced as legal Jobs run out.  6-Harrow has 4, 8-Higgins has 8
      
   PlayCode(1).Color = "Orange"
   PlayCode(2).Color = "Blue"
   PlayCode(3).Color = "Yellow"
   PlayCode(4).Color = "Green"
   pickStartSector = -1
   actionSeq = ASidle

   If Not Logon Then End

   For x = 1 To NO_OF_CONTACTS
      Imag(x).Visible = False
      Imag(x).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Sm" & Nz(varDLookup("Picture", "Contact", "ContactID=" & x)))
      Imag(x).ToolTipText = varDLookup("ContactName", "Contact", "ContactID=" & x)
   Next x
   
   loadImages

   Logic.Open "GameSeq", DB, adOpenDynamic, adLockPessimistic ', adLockOptimistic
   x = GetSeq(inprog)
   If inprog Then
      player.ID = reconnectPlayer()
      If player.ID = 0 Then
         MsgBox "There are no AI slots in the current game." & vbNewLine & "Game requires reset & hosting before the Bot can join", vbExclamation
         End
      End If
      player.PlayName = Nz(varDLookup("Name", "Players", "PlayerID =" & player.ID))
      If getPlayerCount(True) = 1 Then
         MsgBox "Game requires reset & hosting before the Bot can join", vbExclamation
         End
      End If
      Me.Caption = "Firefly AI Bot " & PlayCode(player.ID).Color & " (" & CStr(player.ID) & ")" & " - " & varDLookup("StoryTitle", "Story", "StoryID = " & Logic!StoryID)
      pickStartSector = 2  'flag the selection is done
      getJobParams
      refreshShip " WHERE PlayerID = " & player.ID
      RefreshJob " WHERE PlayerID = " & player.ID
      refreshSolid
      NumOfReavers = varDLookup("NoOfReavers", "Story", "StoryID = " & Logic!StoryID)
      ContactList = getContactList(Logic!StoryID)
      actionSeq = ASidle
      If Logic!Seq = "F" And Logic!player = player.ID And Logic!HostAccept = 0 And Logic!Trader > 0 Then doShowdownAttack Logic!Trader
   End If
   Timing.Enabled = True

End Sub

Public Function GetSeq(ByRef inprog As Boolean)
Dim msg
   Logic.Requery
   GetSeq = Logic!Seq
   
   Select Case GetSeq
   Case "H"
      msg = "Waiting for players to join"
   Case "E"
      msg = "Waiting for a new game to be hosted"
   Case "L"
      msg = "Waiting for a Leaders to be chosen"
   Case "S"
      msg = "Waiting for the Game Setup to complete"
   Case "F"
      msg = "Showdown in progress"
      inprog = True
   Case "R"
      msg = "Waiting for " & PlayCode(Logic!player).PlayName & " [" & PlayCode(Logic!player).Color & "] to finish their GO"
      inprog = True
   Case Else
      msg = "Wait, game in progress!!"
      inprog = True
   End Select
      
   PutMsg msg
End Function

Private Sub Form_Resize()
Dim x
   sftTree.Move sftTree.Left, sftTree.top, Abs(Me.Width - 885), Abs(sftTree2.top - 20)
   sftTree2.Move sftTree2.Left, sftTree2.top, Abs(Me.Width - 885), Abs(Me.Height - sftTree2.top - 920)
   lblSolid.Left = Abs(Me.Width - 860)
   For x = 1 To NO_OF_CONTACTS
      Imag(x).Left = Abs(Me.Width - 700)
   Next x

End Sub

Private Sub loadImages()
Dim Index, SQL
Dim rst As New ADODB.Recordset
   On Error Resume Next
   SQL = "SELECT Distinct Picture FROM Crew"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If Dir(App.Path & "\Pictures\Sm" & rst!Picture) <> "" Then
         AssetImages.ListImages.Add , Left(rst!Picture, Len(rst!Picture) - 4), LoadPicture(App.Path & "\Pictures\Sm" & rst!Picture)
      End If
      rst.MoveNext
   Wend
   
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

'THE MAIN ENGINE of the GAME
' Game States E - Idle/End, H - Host screen, 1-4 players go. S - setup Game, R - run Game, T - Trade, F - Showdowns
' W - Reaver to any Rim or Border sector, X-Move a Reaver 1 sector, Y=Move the Cruiser 1 sector, Z- move the Cruiser adjacent player, V-move Corvette Adjacent player
' actionSeq States = ASidle , ASselect --- >>> , ASend, -> ASidle, <repeat>
' --------------------------------------------------------------
' Dear future me or anyone else who tries to modify this
' When I wrote this code, only god and I knew how it worked.
' Now only God knows.
'
' If you attempt to modify it and it fails
' please increase this counter as a warning for the next person
' TOTAL_HOURS_WASTED_here = 205
' --------------------------------------------------------------
Private Sub Timing_Timer()
Dim status As Variant, errh, thisPlayer As Integer
Dim SectorID, ContactID As Integer, SupplyID As Integer, x, y
Dim maxConsider, fuelleft, HavenID As Integer, DefenderID As Integer, BSupplyID As Integer
Dim bountyJumpSector As Integer, supplyBountySector As Integer, sbountyCardID As Integer
Dim goalSector As Integer
'On Error GoTo err_handler

   SectorID = getPlayerSector(player.ID)
   HavenID = Nz(varDLookup("Haven", "Board", "SectorID=" & SectorID), 0)
   ContactID = Nz(varDLookup("ContactID", "Contact", "SectorID=" & SectorID), 0)
   If Trail(0) = 0 Then
      Trail(0) = SectorID
   End If

   status = GetSeqX(thisPlayer)

   'If status <> "H" And status <> "E" And status <> "L" And pickStartSector > -1 Then
     ' RefreshBoard
   'End If
   If status = "E" Then 'currently in End Game
      PutMsg "Waiting to Host or Join a Game"
      player.ID = 0
   ElseIf (status = "H" Or status = "L") And player.ID = 0 Then 'ready to join
      player.ID = getNewPlayer()
      If player.ID = 0 Then
         MsgBox "No available player slots", vbExclamation, "AI Bot: Fail to join"
         End
      End If
      player.PlayName = varDLookup("Ship", "Players", "PlayerID=" & player.ID)  '"FireflyAI" & CStr(player.ID)
      DB.Execute "Update Players SET Name ='" & player.PlayName & "', AI = 1 WHERE PlayerID = " & player.ID & " AND Name IS NULL"
      

   ElseIf status = "L" And player.ID = thisPlayer Then 'pick leader
      getPlayerCount True
      SetupPlayer player.ID, Logic!StoryID
      leader = getRandomLeader
      DB.Execute "INSERT INTO PlayerSupplies (PlayerID,CardID) VALUES (" & player.ID & ", " & varDLookup("CardID", "SupplyDeck", "CrewID =" & leader) & ")"
      player.PlayName = varDLookup("CrewName", "Crew", "CrewID=" & leader)
      DB.Execute "Update Players SET Name ='" & player.PlayName & "' WHERE PlayerID = " & player.ID

      getRandomCrew 5, leader
      setNextLeader player.ID, leader   'leader
      pickStartSector = 0
      
      actionSeq = ASidle
   
   ElseIf status = "S" And thisPlayer = player.ID And pickStartSector = 0 Then  'your go to pick starting sector on MAP
      Main.Caption = "Firefly AI Bot " & PlayCode(player.ID).Color & " (" & CStr(player.ID) & ")" & " - " & varDLookup("StoryTitle", "Story", "StoryID = " & Logic!StoryID)
      NumOfReavers = varDLookup("NoOfReavers", "Story", "StoryID = " & Logic!StoryID)
      
      ContactList = getContactList(Logic!StoryID)
      
      PutMsg player.PlayName & " selecting Start Sector", player.ID, Logic!GameCntr

       pickStartSector = 1
       x = getStartSector
       DB.Execute "Update Players SET SectorID =" & x & " WHERE PlayerID = " & player.ID
       setRefresh
       If useHavens(Logic!StoryID) Then placeHaven player.ID, x
       Trail(0) = x
       pickStartSector = 2
      
   ElseIf status = "S" And thisPlayer = player.ID And pickStartSector = 2 Then  'setup
      PutMsg player.PlayName & "'s on the Map", player.ID, Logic!GameCntr

      'deal start drive core, and Jobs
      dealDriveAndJobs player.ID

      'starting point selected, pass to next person, or kick the main Running Game cycle off
      setNextPlayerREV player.ID, "R"
      Logic.Requery
      If Logic!Seq = "R" Then
         PutMsg "Next Player's Turn", Logic!player, Logic!GameCntr
      End If
   
   ElseIf status = "F" And thisPlayer <> player.ID And actionSeq = ASidle And Logic!Trader = player.ID And Logic!ClientAccept = 0 Then  'showdown - defend!
      doShowdownDefend Logic!player
      
   ElseIf status = "F" And thisPlayer <> player.ID And actionSeq = ASidle And Logic!Trader = player.ID And Logic!ClientAccept = 1 Then
      If varDLookup("forcereroll", "ShowdownScores", "PlayerID=" & thisPlayer) = 1 Then  'showdown defend - check for re-roll
         DB.Execute "UPDATE ShowdownScores Set Dice = " & RollDice(6, True) & " WHERE PlayerID = " & player.ID
         DB.Execute "UPDATE ShowdownScores Set forcereroll = 0 WHERE PlayerID = " & thisPlayer
         Logic.Update "Trader", 0
      End If
      
   ElseIf status = "F" And thisPlayer = player.ID And actionSeq = ASidle And Logic!HostAccept = 0 Then 'showdown - attack!
      If Nz(varDLookup("PlayerID", "ShowdownScores", "PlayerID=" & player.ID), 9) = player.ID Then 'we have already rolled the dice
         DefenderID = 0
         If hasCrew(player.ID, 91) And Not forcererollused Then
            If Not processBountyJumpScore(DefenderID) Then
               If DefenderID > 0 Then
                  DB.Execute "Update ShowdownScores set forcereroll = 1 WHERE PlayerID = " & player.ID
                  Logic!ClientAccept = 0
                  Logic!HostAccept = 1
                  Logic!Trader = DefenderID
                  Logic.Update
                  forcererollused = True
                  
               End If
            Else
               If DefenderID > 0 Then Logic.Update "HostAccept", 1
            End If
         Else
            Logic.Update "HostAccept", 1
         End If
      Else 'roll the dice first
         doShowdownAttack Logic!Trader
      End If
      
   ElseIf status = "F" And thisPlayer = player.ID And actionSeq = ASidle And Logic!HostAccept = 1 And Logic!ClientAccept = 0 Then
      DefenderID = 0
      If varDLookup("forcereroll", "ShowdownScores", "PlayerID=" & Logic!Trader) = 1 Then 'showdown attack - check for re-roll
         x = RollDice(6, True)
         DB.Execute "UPDATE ShowdownScores Set Dice = " & x & " WHERE PlayerID = " & player.ID
         DB.Execute "UPDATE ShowdownScores Set forcereroll = 0 WHERE PlayerID = " & Logic!Trader
         PutMsg player.PlayName & " is forced to re-roll and gets a " & x, player.ID, Logic!GameCntr
      ElseIf Not processBountyJumpScore(DefenderID) And hasCrew(player.ID, 91) And Not forcererollused Then
         If DefenderID > 0 Then
            DB.Execute "Update ShowdownScores set forcereroll = 1 WHERE PlayerID = " & player.ID
            Logic.Update "trader", DefenderID
            forcererollused = True
         End If
      End If
      
   ElseIf status = "F" And thisPlayer = player.ID And actionSeq = ASidle And Logic!HostAccept = 1 And Logic!ClientAccept = 1 Then   'showdown - attack complete - who won?

         If processBountyJump(bountyCardID) Then 'Bot wins
            targetJobID = 1
            fullburndone = (FullburnMovesDone > 0) Or fullburndone
         Else
            fullburndone = True  'set to not move as we could have another go here
         End If
         workdone = True
         Logic.Update "Seq", "R"
         
   ElseIf status = "U" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Corvette to any planetary sector
      x = setPlayer(player.ID, "", 0, True)
      doMoveCorvettePlanetary
      'turn finished, push to next player (for SP thats you)
      thisPlayer = setPlayer(player.ID, "R", 0)
         
   ElseIf status = "V" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Corvette Cycle from another Player's Nav move
      x = setPlayer(player.ID, "", 0, True)
      doMoveCorvetteAdjacent getPlayerSector(x)
      'turn finished, push to next player (for SP thats you)
      thisPlayer = setPlayer(player.ID, "R", 0)

   ElseIf status = "W" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Reaver Cycle from another Player's Nav move
      doMoveCutterPlanetary 6 + RollDice(NumOfReavers)
      'turn finished, push to next player (for SP thats you)
      thisPlayer = setPlayer(player.ID, "R", 0)
      
   ElseIf status = "X" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the Move Reaver Cycle from another Player's Nav move
      moveAutoAI 6 + RollDice(NumOfReavers)
      PutMsg "A Cutter is on the move"
      'turn finished, push to next player (for SP thats you)
      thisPlayer = setPlayer(player.ID, "R", 0)
      
   ElseIf status = "Y" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the move Cruiser cycle
      moveAutoAI 5
      PutMsg "The Cruiser is on the move"
      'turn finished, push to next player (for SP thats you)
      thisPlayer = setPlayer(player.ID, "R", 0)
      
   ElseIf status = "Z" And thisPlayer = player.ID And actionSeq = ASidle Then 'capture the move Cruiser cycle
      x = setPlayer(player.ID, "", 0, True)
      doMoveAllianceAdjacent getPlayerSector(x)
      PutMsg "The Cruiser is on the move"
      'turn finished, push to next player (for SP thats you)
      thisPlayer = setPlayer(player.ID, "R", 0)
      
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< MAIN In-Game CYCLE >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASidle Then   'MAIN Cycle - init your go
      'PutMsg player.PlayName & "'s having a go", player.ID, Logic!Gamecntr
      
      fuelleft = varDLookup("Fuel", "Players", "PlayerID=" & player.ID)
      
      targetJobCard = varDLookup("CardID", "PlayerJobs", "JobStatus < 3 AND PlayerID=" & player.ID)
      
      targetSupplySector = getNearestSupply(SectorID) 'get the sector of the closest supply or Haven
      
      resolveToken SectorID
      
      targetSector = 0
      bountyCardID = 0
      
      If getCruiserSector = SectorID And hasWarrant Then
         PutMsg player.PlayName & " pays fine for Warrants", player.ID, Logic!GameCntr
         DB.Execute "UPDATE Players SET Warrants=0, Pay = " & IIf(getMoney(player.ID) < 1000, "0", "Pay - 1000") & " WHERE PlayerID =" & player.ID
      End If
      
      'need fuel??
      If ((fuelleft < 3 And FullburnMovesDone = 0) Or fuelleft < 2) And getMoney(player.ID) > 0 And Not fullburndone Then     'head for nearest supply as our top priority
         If SectorID = targetSupplySector Then
            buyfueldone = buyFuel(fuelleft, HavenID, SectorID)
            fullburndone = (FullburnMovesDone > 0) Or fullburndone

         Else
            'go there
            targetSector = targetSupplySector
            If goToSupply(SectorID, targetSector, (FullburnMovesDone = 0)) > 0 And Not fullburndone Then 'move then check if a mosey or a fullburn was required
               SectorID = processMove ' getPlayerSector(player.ID)
               
            Else 'we are there, load fuel below
               buyfueldone = buyFuel(fuelleft, HavenID, SectorID)
               fullburndone = (FullburnMovesDone > 0) Or fullburndone

            End If
         End If

      ' do we have a job or need to go to a Contact?
      ElseIf IsNull(targetJobCard) Then 'no job, are we at next Contact?
         SupplyID = Nz(varDLookup("SupplyID", "Supply", "SectorID=" & SectorID), 0)
         targetContact = getNearestContact(SectorID)
         'put a check here to see if goal bounty requirements have been met.  If so, then ignore
         
         'do we need to go to a goal sector?
         If hasGoalSector(goalSector) Then
            targetSector = goalSector
         
         ElseIf isBountyEnabled And Not bountiesDone() Then
        
            If (IsEmpty(targetJobCard) Or IsNull(targetJobCard)) And getCrewCount(player.ID) > 3 Then 'see if there is a rival player with a claimed bounty
               bountyJumpSector = findBountyJump(SectorID, DefenderID, bountyCardID) 'we have a bounty chase
            End If
            If (IsEmpty(targetJobCard) Or IsNull(targetJobCard)) Then 'check for Supply bounty
               supplyBountySector = findSupplyBounty(SectorID, BSupplyID, sbountyCardID)
            End If
            'if both, then prioritise which one
            If bountyJumpSector > 0 And supplyBountySector = 0 Then
               targetSector = bountyJumpSector
            ElseIf bountyJumpSector = 0 And supplyBountySector > 0 Then
               targetSector = supplyBountySector
               bountyCardID = sbountyCardID
            ElseIf bountyJumpSector > 0 And supplyBountySector > 0 Then 'both
               x = getSectorCount(SectorID, bountyJumpSector)
               y = getSectorCount(SectorID, supplyBountySector)
               
               If x <= y Then  'go to closest, if same then player
                  targetSector = bountyJumpSector
                  supplyBountySector = 0
               Else
                  targetSector = supplyBountySector
                  bountyJumpSector = 0
                  bountyCardID = sbountyCardID
               End If
            End If
            
         End If
         
         
         If targetSector > 0 Then 'we have a bounty chase
            If goToPlayer(SectorID, targetSector, (FullburnMovesDone = 0), goalSector) > 0 And Not fullburndone Then 'move then check if a mosey or a fullburn was required
               SectorID = processMove ' getPlayerSector(player.ID)
               
            ElseIf SectorID = targetSector Then 'we are there, load fuel below
               
               fullburndone = (FullburnMovesDone > 0) Or fullburndone
               
               If workdone Or goalSector Then  'we have just finished a job, so wait until next turn
                  workdone = True
                  fullburndone = True
               
               ElseIf BSupplyID > 0 And supplyBountySector > 0 Then
                  If doShowDownSupply(BSupplyID, bountyCardID) Then
                     targetJobID = 1
                  Else
                     fullburndone = True 'set to not move as we could have another go here
                  End If
                  fullburndone = (FullburnMovesDone > 0) Or fullburndone
                  workdone = True
                  
               ElseIf doBoardingTest(DefenderID) Then
                  'AI has boarded! status "F"
               Else
                  targetSector = 0
                  workdone = True
                  fullburndone = True  'set to not move as we could have another go here
               End If
            Else 'we didn't make it
               workdone = True
            End If
         ElseIf SectorID = varDLookup("SectorID", "Contact", "ContactID=" & targetContact) And (IsEmpty(targetJobCard) Or IsNull(targetJobCard)) Then 'yes, we're at the Contact
            If targetContact = 2 And hasWarrant() And isSolid(player.ID, targetContact) And getMoney(player.ID) > 1999 Then
               PutMsg player.PlayName & " gets Badger to clear Warrants for $1000", player.ID, Logic!GameCntr
               DB.Execute "UPDATE Players SET Warrants=0, Pay = Pay - 1000 WHERE PlayerID =" & player.ID
            End If
            'pickup a job
            targetJobCard = getJob(targetContact)
            If IsEmpty(targetJobCard) Then
               PutMsg player.PlayName & " finds no Legal Jobs for " & varDLookup("ContactName", "Contact", "ContactID=" & targetContact), player.ID, Logic!GameCntr
               'reset for next contact
               ContactList = Replace(ContactList, targetContact & ",", "") 'for 4,5
               ContactList = Replace(ContactList, "," & targetContact, "") 'for 3,4
               ContactList = Replace(ContactList, targetContact, "") 'for 4
               If ContactList = "" Then ContactList = "1,2,4,5"  'reset
            End If
            
            targetJobID = 1
            workdone = True
            fullburndone = (FullburnMovesDone > 0) Or fullburndone
            
         ElseIf (IsEmpty(targetJobCard) Or IsNull(targetJobCard)) And Not fullburndone And targetContact > 0 Then 'Not at a Contact - Head for target contact / 1,2,4,5 that has a legal job left
            targetSector = varDLookup("SectorID", "Contact", "ContactID=" & targetContact)
            If goToContact(SectorID, targetSector, (FullburnMovesDone = 0)) > 0 And Not fullburndone Then 'move then check if a mosey or a fullburn was required
               SectorID = processMove ' getPlayerSector(player.ID)
           
            Else 'we are there
               'workdone = True
               fullburndone = (FullburnMovesDone > 0) Or fullburndone
                  
            End If
         ElseIf SectorID <> targetSupplySector And Not fullburndone And targetContact = 0 Then  'head for nearest supply as our top priority
            'go there
            targetSector = targetSupplySector
            If goToSupply(SectorID, targetSector, (FullburnMovesDone = 0)) > 0 And Not fullburndone Then 'move then check if a mosey or a fullburn was required
               SectorID = processMove ' getPlayerSector(player.ID)
               
            Else 'we are there, load fuel below
               
               fullburndone = (FullburnMovesDone > 0) Or fullburndone
                  
            End If
         Else  'run out of valid actions
            fullburndone = True
            workdone = True
         End If
         
      Else 'we have a job, go to it
         If IsEmpty(targetJobID) Then targetJobID = 1
         targetSector = getJobSector(targetJobCard, targetJobID)
         If targetSector = SectorID Then   'we there
            'do job
            If workdone And ((FullburnMovesDone > 0) Or fullburndone) Then 'already used this action
               fullburndone = True
            ElseIf targetJobID = 1 Then
               'check if this Job has a part 2?
               targetSector = getJobSector(targetJobCard, 2)
               
               If targetSector = 0 Then 'complete job (only 1 part)
                  completeJob targetJobCard, targetJobID

               Else
                  completeFirstPartJob targetJobCard
                  targetJobID = 2
               End If
               
            Else 'complete Job (part 2)
               completeJob targetJobCard, targetJobID
               targetJobID = 1
            End If
            
            workdone = True
            fullburndone = (FullburnMovesDone > 0) Or fullburndone
            
         ElseIf Not fullburndone Then    'we're not there yet
            If goToContact(SectorID, targetSector, (FullburnMovesDone = 0)) > 0 Then  'move then check if a mosey or a fullburn was required

               SectorID = processMove ' getPlayerSector(player.ID)
               
            End If
         End If
         
      End If
           
      If FullburnMovesDone > 4 Or fullburndone Then
         fullburndone = True
         If targetSector > 0 And targetSector <> SectorID Then
            workdone = True
         End If
      End If
      're-check supply & haven as we may have moved
      SupplyID = Nz(varDLookup("SupplyID", "Supply", "SectorID=" & SectorID), 0)
      HavenID = Nz(varDLookup("Haven", "Board", "SectorID=" & SectorID), 0)
      'ShoreLeave
      If (SupplyID > 0 Or HavenID > 0) And Abs(doShoreLeave(player.ID, True)) <= getMoney(player.ID) And hasDisgruntled(player.ID) Then
         x = doShoreLeave(player.ID, False, (HavenID = player.ID))
         PutMsg player.PlayName & " decides to shout the Crew some Shoreleave for " & IIf(x = -1, "Free!", "$" & Abs(x)), player.ID, Logic!GameCntr
         fullburndone = (FullburnMovesDone > 0) Or fullburndone
      ElseIf SupplyID > 0 And Not buydone Then
         pullSupplies SupplyID
         If CrewCapacity(player.ID) > getCrewCount(player.ID) Then 'need some crew
            getCrew SupplyID
            fullburndone = (FullburnMovesDone > 0) Or fullburndone
         End If
         buydone = True
      End If
      'Fuel Check & Buy if not done initially
      If (SupplyID > 0 Or HavenID > 0) And ((fullburndone And fuelleft < MAXFUEL) Or fuelleft < 3) And Not (targetSector = SectorID) And Not buyfueldone Then
         buyfueldone = buyFuel(fuelleft, HavenID, SectorID)

         workdone = True 'bought fuel
         fullburndone = (FullburnMovesDone > 0) Or fullburndone
         
      End If
      
      resolveToken SectorID
      
      If (workdone And fullburndone) Then
         movesDone
         actionSeq = ASEnd
      End If
         
   
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< END CYCLE >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   ElseIf status = "R" And thisPlayer = player.ID And actionSeq = ASEnd Then 'Finish up your turn
      'Check if WON!
      CheckWon player.ID
      

      'turn finished, push to next player (for SP thats you)
      thisPlayer = setNextPlayer(player.ID)
      If thisPlayer <> player.ID Then
         PutMsg "Next Player's Turn", thisPlayer, Logic!GameCntr
      End If

      actionSeq = ASidle
      
      ClearTrail

   End If
      
   refreshShip " WHERE PlayerID = " & player.ID
   RefreshJob " WHERE PlayerID = " & player.ID
   refreshSolid

   Exit Sub
  
err_handler:
  errh = MsgBox(Err.Description, vbCritical + vbAbortRetryIgnore, "Error in Main Cycle")
  Select Case errh
  Case vbRetry
    Resume
  Case vbAbort
    'exit
  Case vbIgnore
    Resume Next
  End Select
  
   
   
End Sub

Private Function buyFuel(ByRef fuelleft, ByVal HavenID, ByVal SectorID) As Boolean
Dim x
         If fuelleft < 0 Then fuelleft = 0
         If (MAXFUEL - fuelleft) * 100 > getMoney(player.ID) Then
            fuelleft = MAXFUEL - (getMoney(player.ID) / 100)
         End If
         If HavenID = player.ID Then
            x = MAXFUEL - fuelleft
            If x > 4 Then
               x = x - 4
            Else
               x = 0
            End If
            DB.Execute "UPDATE Players SET Fuel =Fuel + (" & MAXFUEL & "- " & CStr(fuelleft) & "),Pay = Pay - (" & CStr(x) & "*100) WHERE PlayerID = " & player.ID
            PutMsg player.PlayName & " gets " & CStr(MAXFUEL - fuelleft) & " Fuel at their Haven", player.ID, Logic!GameCntr
         Else
            'buy fuel
            DB.Execute "UPDATE Players SET Fuel =Fuel + (" & MAXFUEL & "- " & CStr(fuelleft) & "),Pay = Pay - ((" & MAXFUEL & "-" & CStr(fuelleft) & ")*100) WHERE PlayerID = " & player.ID
            If HavenID > 0 Then
               PutMsg player.PlayName & " buys " & CStr(MAXFUEL - fuelleft) & " Fuel at " & varDLookup("Name", "Players", "PlayerID=" & HavenID) & "'s Haven", player.ID, Logic!GameCntr
            Else
               PutMsg player.PlayName & " buys " & CStr(MAXFUEL - fuelleft) & " Fuel at " & varDLookup("SupplyName", "Supply", "SectorID=" & SectorID), player.ID, Logic!GameCntr
            End If

         End If
         fuelleft = MAXFUEL
         buyFuel = True
End Function

Private Function processMove() As Integer
Dim fuel As Integer

   fuel = 1 + getExtraBurn(player.ID)
   processMove = getPlayerSector(player.ID)
   If FullburnMovesDone = 0 And (targetSector = processMove Or getFuel(player.ID) < 1) Then 'Mosey
      PutMsg player.PlayName & " Moseys to Sector " & processMove, player.ID, Logic!GameCntr
      fullburndone = True
   ElseIf FullburnMovesDone = 0 Then
      'burn 1 fuel
      DB.Execute "UPDATE Players Set Fuel = Fuel - " & fuel & " WHERE PlayerID = " & player.ID
      PutMsg player.PlayName & " goes FullBurn" & IIf(fuel > 1, " with Heavy Load", "") & " and has " & varDLookup("Fuel", "Players", "PlayerID=" & player.ID) & " Fuel left", player.ID, Logic!GameCntr
      FullburnMovesDone = FullburnMovesDone + 1
      showNav processMove
   Else
      FullburnMovesDone = FullburnMovesDone + 1
      showNav processMove
   End If
End Function

Private Function goToContact(ByVal SectorID, ByVal ContactSectorID, ByVal canMosey)
Dim rst As ADODB.Recordset, SQL, x As Integer, closest As Integer, targetSectorID, playerSector

      If SectorID <> ContactSectorID Then
         goToContact = getNextSector(SectorID, ContactSectorID, canMosey)
         If goToContact > 0 Then
            DB.Execute "UPDATE Players SET SectorID = " & goToContact & " WHERE PlayerID =" & player.ID
            setRefresh
            Trail(FullburnMovesDone + 1) = goToContact
            PutMsg player.PlayName & " moving towards " & Nz(varDLookup("PlanetName", "Planet", "SectorID=" & ContactSectorID), "the Cruiser") & " via Sector " & goToContact, player.ID, Logic!GameCntr
            playsnd 1, True
            
         Else
            PutMsg player.PlayName & " has no viable path", player.ID, Logic!GameCntr
            movesDone
            actionSeq = ASEnd
         End If
      Else
         goToContact = 0 'we here already
      End If
      
      
End Function

Private Function goToSupply(ByVal SectorID, ByVal SupplySectorID, ByVal canMosey)
Dim rst As ADODB.Recordset, SQL, x As Integer, closest As Integer, targetSectorID, playerSector

      If SectorID <> SupplySectorID Then
         goToSupply = getNextSector(SectorID, SupplySectorID, canMosey)
         If goToSupply > 0 Then
            DB.Execute "UPDATE Players SET SectorID = " & goToSupply & " WHERE PlayerID =" & player.ID
            setRefresh
            Trail(FullburnMovesDone + 1) = goToSupply
            PutMsg player.PlayName & " moving towards " & varDLookup("PlanetName", "Planet", "SectorID=" & SupplySectorID) & " via Sector " & goToSupply & " to get Supplies", player.ID, Logic!GameCntr
            playsnd 1, True
         Else
            PutMsg player.PlayName & " has no viable path", player.ID, Logic!GameCntr
            movesDone
            actionSeq = ASEnd
         End If
      Else
         goToSupply = 0 'we here already
      End If
      
      
End Function

Private Function goToPlayer(ByVal SectorID, ByVal PlayerSectorID, ByVal canMosey, ByVal goalSector)
Dim rst As ADODB.Recordset, SQL, x As Integer, closest As Integer, targetSectorID, playerSector, PlanetName As String

      PlanetName = Nz(varDLookup("PlanetName", "Planet", "SectorID=" & PlayerSectorID))
      If PlanetName <> "" Then PlanetName = "towards " & PlanetName
      
      If SectorID <> PlayerSectorID Then
         goToPlayer = getNextSector(SectorID, PlayerSectorID, canMosey)
         If goToPlayer > 0 Then
            DB.Execute "UPDATE Players SET SectorID = " & goToPlayer & " WHERE PlayerID =" & player.ID
            setRefresh
            Trail(FullburnMovesDone + 1) = goToPlayer
            PutMsg player.PlayName & " moving " & PlanetName & " via Sector " & goToPlayer & IIf(goalSector = 0, " to seek a Bounty", " to reach a goal"), player.ID, Logic!GameCntr
            playsnd 1, True
         Else
            PutMsg player.PlayName & " has no viable path", player.ID, Logic!GameCntr
            movesDone
            actionSeq = ASEnd
         End If
      Else
         goToPlayer = 0 'we here already
      End If
      
      
End Function

Private Function getNearestContact(ByVal SectorID)
Dim rst As ADODB.Recordset, SQL, x As Integer, closest As Integer, targetSectorID, playerSector, tmpContacts As String
   Set rst = New ADODB.Recordset
   'ContactList = Replace(ContactList, targetContact & ",", "")
   'look at contact list and see if any are solid and remove them
   'if left with none (all solid) then reset list
      SQL = "SELECT * FROM Contact WHERE ContactID IN (" & ContactList & ")"
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      While Not rst.EOF
         If Not isSolid(player.ID, rst!ContactID) Then
            tmpContacts = tmpContacts & IIf(tmpContacts = "", "", ",") & CStr(rst!ContactID)
         End If
         rst.MoveNext
      Wend
      If tmpContacts = "" Then tmpContacts = ContactList
      rst.Close
      
      SQL = "SELECT * FROM Contact WHERE ContactID IN (" & tmpContacts & ")"
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      closest = 500
      While Not rst.EOF
         x = getSectorCount(SectorID, rst!SectorID)
         If x < closest And getCutterSector(rst!SectorID) = 0 Then 'dismiss if cutter there
            closest = x
            getNearestContact = rst!ContactID
         End If
         rst.MoveNext
      Wend
     
End Function

Private Function getNearestSupply(ByVal SectorID)
Dim rst As ADODB.Recordset, SQL, x As Integer, closest As Integer, targetSectorID, playerSector

      Set rst = New ADODB.Recordset
      SQL = "SELECT * FROM Supply WHERE SectorID > 0"
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      closest = 500
      While Not rst.EOF
         x = getSectorCount(SectorID, rst!SectorID)
         If x < closest Then
            closest = x
            getNearestSupply = rst!SectorID
         End If
         rst.MoveNext
      Wend
      rst.Close
      
      Set rst = New ADODB.Recordset
      SQL = "SELECT SectorID FROM Board WHERE Haven > 0"
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      While Not rst.EOF
         x = getSectorCount(SectorID, rst!SectorID)
         If x < closest Then
            closest = x
            getNearestSupply = rst!SectorID
         End If
         rst.MoveNext
      Wend
      rst.Close
     
End Function

Private Sub movesDone()
   MoseyMovesDone = 0
   FullburnMovesDone = 0
   moseydone = False
   fullburndone = False
   buydone = False
   buyfueldone = False
   dealdone = False
   workdone = False
   forcererollused = False

End Sub

Private Sub getJobParams()
Dim rst As ADODB.Recordset, SQL


      Set rst = New ADODB.Recordset
      SQL = "SELECT ContactID, p.CardID,JobStatus FROM ContactDeck c, PlayerJobs p WHERE c.cardID = p.CardID AND PlayerID= " & player.ID & " AND JobStatus < 2"
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      If Not rst.EOF Then
         targetContact = rst!ContactID
         targetJobCard = rst!CardID
         targetJobID = rst!JobStatus + 1
      End If
      rst.Close

End Sub

Public Function completeFirstPartJob(ByVal CardID)
Dim rst As New ADODB.Recordset
Dim SQL, msg As String, contra As Integer, passgr  As Integer, fugi As Integer, cargo As Integer
      Set rst = New ADODB.Recordset
      SQL = "SELECT Job.* FROM Job INNER JOIN ContactDeck ON Job.JobID = ContactDeck.Job1ID WHERE ContactDeck.CardID=" & CardID
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      If Not rst.EOF Then
         contra = IIf(rst!Contraband = 14, 7, rst!Contraband)
         passgr = IIf(rst!Passenger = 14, 7, rst!Passenger)
         fugi = IIf(rst!Fugitive = 14, 7, rst!Fugitive)
         cargo = rst!cargo
         
         If rst!TagnBag > 0 Then 'need to find out what to deliver on 2nd part
            getJob2Reqs CardID, cargo, contra
         End If
         
         DB.Execute "UPDATE Players SET Fuel = Fuel + " & rst!fuel & ", Parts = Parts + " & rst!parts & ", Cargo = Cargo + " & cargo & ", Contraband = Contraband + " & contra & ", Passenger = Passenger + " & passgr & ", Fugitive = Fugitive + " & fugi & " WHERE PlayerID=" & player.ID

         DB.Execute "UPDATE PlayerJobs SET JobStatus = 1 WHERE PlayerID = " & player.ID & " AND CardID = " & CardID
         
         msg = IIf(rst!fuel = 0, "", rst!fuel & " Fuel")
         msg = msg & IIf(rst!parts = 0, "", IIf(Len(msg) > 0, ", ", "") & rst!parts & " Part" & IIf(rst!parts > 1, "s", ""))
         msg = msg & IIf(rst!cargo = 0, "", IIf(Len(msg) > 0, ", ", "") & rst!cargo & " Cargo")
         msg = msg & IIf(contra = 0, "", IIf(Len(msg) > 0, ", ", "") & contra & " Contraband")
         msg = msg & IIf(passgr = 0, "", IIf(Len(msg) > 0, ", ", "") & passgr & " Passenger" & IIf(passgr > 1, "s", ""))
         msg = msg & IIf(fugi = 0, "", IIf(Len(msg) > 0, ", ", "") & fugi & " Fugitive" & IIf(fugi > 1, "s", ""))
         
         PutMsg player.PlayName & " completed the first part of Job " & CardID & IIf(msg = "", "", " and took on " & msg), player.ID, Logic!GameCntr

      End If
      rst.Close
         
End Function

Public Function completeJob(ByVal CardID, ByVal JobID)
Dim rst As New ADODB.Recordset, jobpay, crewpay, perk As Integer
Dim SQL, msg As String, contra As Integer, passgr  As Integer, fugi  As Integer, contact As Integer
      Set rst = New ADODB.Recordset
      SQL = "SELECT Pay, WinResult, ContactID, JobTypeID, JobType2D, Job.* FROM Job INNER JOIN ContactDeck ON Job.JobID = ContactDeck.Job" & JobID & "ID WHERE ContactDeck.CardID=" & CardID
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      If Not rst.EOF Then
         contra = IIf(rst!Contraband = -14, -7, rst!Contraband)
         passgr = IIf(rst!Passenger = -14, -7, rst!Passenger)
         fugi = IIf(rst!Fugitive = -14, -7, rst!Fugitive)
         contact = rst!ContactID
         
         DB.Execute "UPDATE Players SET Fuel = Fuel + " & rst!fuel & ", Parts = Parts + " & rst!parts & ", Cargo = Cargo + " & rst!cargo & ", Contraband = Contraband + " & contra & ", Passenger = 0, Fugitive = Fugitive + " & fugi & IIf(contact = 10 Or contact = 0, "", ", Solid" & rst!ContactID & "= 1") & " WHERE PlayerID=" & player.ID

         DB.Execute "UPDATE PlayerJobs SET JobStatus = 3 WHERE PlayerID = " & player.ID & " AND CardID = " & CardID
         
        
         If rst!Contraband = -14 Or rst!Passenger = -14 Or rst!Fugitive = -14 Then
            jobpay = Abs(passgr) * 200 + Abs(contra) * 300 + Abs(fugi) * 300
         Else
            jobpay = rst!Pay
         End If
         
         'check crew perks
         jobpay = jobpay + getJobCrewBonus(player.ID, rst!JobTypeID) + getJobCrewBonus(player.ID, rst!JobType2D)
         
         'look for bounty bonus
         If rst!ContactID = 10 Then
            perk = getPerkAttributeSum(player.ID, "BountyBonus")
            If perk > 0 Then
               PutMsg player.PlayName & " gets an extra $" & perk & " Bounty bonus for having Lawmen in the crew", player.ID, Logic!GameCntr
               jobpay = jobpay + perk
            End If
         End If

         If contact = 0 Then 'goal job
            crewpay = 0
         ElseIf RollDice(6) < 4 And Not hasDisgruntled(player.ID) And jobpay > 0 Then 'don't pay em
            doDisgruntled player.ID, 2
            PutMsg player.PlayName & " didn't pay the crew and they are not Happy about it", player.ID, Logic!GameCntr
            crewpay = 0
         Else
            crewpay = getCrewPay
         End If
         
         DB.Execute "UPDATE Players SET Pay = Pay + " & CStr(jobpay - crewpay) & " WHERE PlayerID = " & player.ID
         
         msg = IIf(rst!fuel = 0, "", Abs(rst!fuel) & " Fuel")
         msg = msg & IIf(rst!parts = 0, "", IIf(Len(msg) > 0, ", ", "") & Abs(rst!parts) & " Part" & IIf(rst!parts < -1, "s", ""))
         msg = msg & IIf(rst!cargo = 0, "", IIf(Len(msg) > 0, ", ", "") & Abs(rst!cargo) & " Cargo")
         msg = msg & IIf(contra = 0, "", IIf(Len(msg) > 0, ", ", "") & Abs(contra) & " Contraband")
         msg = msg & IIf(passgr = 0, "", IIf(Len(msg) > 0, ", ", "") & Abs(passgr) & " Passenger" & IIf(passgr < -1, "s", ""))
         msg = msg & IIf(fugi = 0, "", IIf(Len(msg) > 0, ", ", "") & Abs(fugi) & " Fugitive" & IIf(fugi < -1, "s", ""))
         
         PutMsg player.PlayName & IIf(msg = "", "", " unloaded " & msg & " and") & " completed Job " & targetJobCard & " for $" & CStr(jobpay - crewpay) & IIf(contact = 10 Or contact = 0, "", " and is Solid with " & varDLookup("ContactName", "Contact", "ContactID=" & rst!ContactID)), player.ID, Logic!GameCntr
         
      End If
      rst.Close
         
End Function

Private Sub showNav(ByVal SectorID)
Dim SQL, reshuffle, Zone, x
Dim rst As New ADODB.Recordset

      Zone = varDLookup("Zones", "Board", "SectorID=" & SectorID)
      
      'Read in the next NAV card and display either 1 or 2 options
      
       'OPTION 1 ===================================================================================
      SQL = "SELECT NavDeck.CardID, NavDeck.CardName, NavDeck.Reshuffle, NavDeck.Seq, NavOption.*, Opt2.WinKeepFlying AS KeepFlying "
      SQL = SQL & "FROM (NavOption INNER JOIN NavDeck ON NavOption.OptionID = NavDeck.Option1ID) LEFT JOIN NavOption as Opt2 ON Opt2.OptionID = NavDeck.Option2ID "

      SQL = SQL & "Where NavDeck.Zones = '" & Zone & "' And NavDeck.Seq > 6 "
      SQL = SQL & "ORDER BY NavDeck.Seq"

      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      If rst.EOF Then  ' this happens when the reshuffle card is in the discard pile at start of game setup
         ShuffleDeck "Nav", True, False, Zone
         PutMsg player.PlayName & " Reshuffling NavDeck " & Zone & " due to end of deck", player.ID, Logic!GameCntr
         rst.Close
         rst.Open SQL, DB, adOpenStatic, adLockReadOnly
      End If
      If Not rst.EOF Then
         'put special outcomes first >>>>>>>>>>>>>>>>
         If (rst!CardName = "Reaver Cutter!") And getCruiserCorvette(SectorID) = 6 Then 'corvette shoos the Reavers away
            movesDone
            actionSeq = ASEnd
            PutMsg player.PlayName & " is Shielded from a Reaver Cutter attack by the Alliance Corvette", player.ID, Logic!GameCntr
            
         'skip Customs Inspection if solid with Harken
         ElseIf (rst!CardName = "Customs Inspection") And isSolid(player.ID, 5) Then
            PutMsg player.PlayName & " being Solid with Harken avoided a Customs Inspection", player.ID, Logic!GameCntr
         'skip Customs Inspection if solid with Harken
         ElseIf (rst!CardName = "Customs Inspection") And Not isSolid(player.ID, 5) Then
            movesDone
            actionSeq = ASEnd
            PutMsg player.PlayName & " was stopped by a Customs Inspection", player.ID, Logic!GameCntr
         ElseIf rst!WinSolid > 0 And Not isSolid(player.ID, rst!WinSolid) And Zone = "A" Then
            If doMoveAllianceAdjacent(SectorID) Then
               PutMsg player.PlayName & " has the Cruiser move into an adjacent Sector, tipped off by " & varDLookup("ContactName", "Contact", "ContactID=" & rst!WinSolid), player.ID, Logic!GameCntr
            End If
         Else
            PutMsg player.PlayName & " Nav: " & rst!OptionName & " - " & rst!Details, player.ID, Logic!GameCntr

            '<<<<<<<<<<  INSERT NAV OUTCOMES >>>>>>>>>>>>>>
            Select Case rst!MoveReaver
               Case 1   ' 1 - move 1
                  moveAutoAI 6 + RollDice(NumOfReavers)
                  
               Case 2    '2-you move reaver to any B zone,
                  doMoveCutterPlanetary 6 + RollDice(NumOfReavers)
                  
               Case 3    '3-move to your location  (evade done later)
                  If getCutterSector(SectorID) = 0 Then
                     MoveShip 6 + RollDice(NumOfReavers), SectorID
                  End If
                  
               Case 4  'other player move reaver to any B zone,
                  doMoveCutterPlanetary 6 + RollDice(NumOfReavers)
            
            End Select

            Select Case rst!MoveAlliance
               Case 1   ' 1 - move 1
                  moveAutoAI 5
                  
               Case 2   '2- move to any
                  doMoveCruiserToFreeSector
                  
               Case 3   '3-move to outlaw ship
                  x = outlawExists(player.ID)
                  If x > 0 Then
                     MoveShip 5, x
                  End If
                  
               Case 4 'alliance pays you a visit
                  MoveShip 5, SectorID
                  fullburndone = True
               
               Case 5 'move adjacent if failed
                  doMoveAllianceAdjacent SectorID
                  
               Case 6 'alert tokens adjacent your posn
                  doAddTokensAdjacent SectorID
                  setRefresh
               Case 7 'corvette contact
                  fullburndone = True
'                  If SeizeAllFugi(player.ID) Then
'                     PutMsg player.PlayName & " lost some Fugitives not in Stash", player.ID, Logic!Gamecntr, True, getLeader()
'                  End If
               
               Case 8 'discard 1 crew
'                  Set frmSeize = New frmSeized
'                  frmSeize.Caption = "Select the Crew Member detained by the Alliance"
'                  If frmSeize.RefreshDiscardList() > 0 Then 'crew exist
'                     frmSeize.Show 1
'                  End If
               
               Case 9 'alert tokens at every Outlaw Ship
                  doAddTokensOutlaws
                  setRefresh
               Case 10 ' Move Corvette Adjacent player

                  doMoveCorvetteAdjacent SectorID

                  
               Case 11  'Corvette to an unoccupied Alliance, Border, or Rim Planetary Sector.

                  doMoveCorvettePlanetary

                  
               Case 12  'move Operative's Corvette 1 or 2 Sectors within Alliance, Border or Rim Space
                  x = getCorvetteSector
                  moveAutoCorvette2 0, False, x
                        
            End Select
            
            If rst!MovePlayer > 0 Then
               For x = 1 To rst!MovePlayer
                  moveAutoAI player.ID, 1, True
               Next x
            End If
            
            If rst!WinKeepFlying = 0 And Nz(rst!KeepFlying, 0) = 0 Then 'stop
               movesDone
               actionSeq = ASEnd
               If rst!Evade > 0 Then
                  moveAutoAI player.ID, 1, True
                  PutMsg player.PlayName & " has had to Evade", player.ID, Logic!GameCntr
               Else
                  PutMsg player.PlayName & " has come to an abrupt halt", player.ID, Logic!GameCntr
               End If
            End If

         End If
         reshuffle = rst!reshuffle
         'pull the card out of the deck, assign it to the user for debugging
         DB.Execute "UPDATE NavDeck SET Seq = " & CStr(player.ID) & " WHERE CardID = " & CStr(rst!CardID)
         'rst!Seq = player.ID
         'rst.Update
      End If
      rst.Close
     
      If reshuffle = 1 Then 'ready for next turn
         ShuffleDeck "Nav", True, False, Zone
         PutMsg player.PlayName & " Reshuffling NavDeck " & Zone & " due to reshuffle card", player.ID, Logic!GameCntr
         If Zone = "A" Then
            If pushBounties() Then
               If DrawDeck("Contact", 10, 3) Then PutMsg "New Bounties available"
            End If
         End If
      End If


      
End Sub

Private Sub pullSupplies(ByVal SupplyID)
Dim rst As ADODB.Recordset, SQL, cnt As Integer, ldr As Integer, kword As String, x As Integer
Dim Skilltype As Integer, skill As Integer

      Set rst = New ADODB.Recordset
      SQL = "SELECT CardID FROM SupplyDeck WHERE Seq > 6 AND SupplyID =" & SupplyID & " ORDER BY Seq"
      rst.CursorLocation = adUseClient
      rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
      cnt = 0
      While Not rst.EOF And cnt < 3
         cnt = cnt + 1
         DB.Execute "UPDATE SupplyDeck SET Seq = 5 WHERE CardID = " & rst!CardID
         rst.MoveNext
      Wend
      rst.Close
      
      If isBountyEnabled Then  'find the Captain some gear
         ldr = getLeader()
         SQL = "SELECT PlayerSupplies.CardID "
         SQL = SQL & "FROM Players INNER JOIN PlayerSupplies ON (PlayerSupplies.CrewID = Players.Leader) AND (Players.PlayerID = PlayerSupplies.PlayerID) "
         SQL = SQL & "WHERE PlayerSupplies.PlayerID = " & player.ID
         rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
         If Not rst.EOF Then Exit Sub
         rst.Close
      
         skill = getMySkill(Skilltype)
         SQL = "SELECT SupplyDeck.CardID, Gear.GearName, Gear.Fight, Gear.Pay, Gear.GearID, Gear.Keywords "
         SQL = SQL & "FROM Gear INNER JOIN SupplyDeck ON Gear.GearID = SupplyDeck.GearID "
         SQL = SQL & "WHERE Gear.Discard=0 AND SupplyDeck.Seq=5 AND Gear." & cstrSkill(Skilltype) & "> 0 AND  SupplyDeck.SupplyID= " & SupplyID
         SQL = SQL & " ORDER BY Gear.Pay"
         rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
         If Not rst.EOF Then
            DB.Execute "UPDATE SupplyDeck SET Seq =" & player.ID & " WHERE CardID = " & rst!CardID
            'add the card to the players deck
            DB.Execute "INSERT INTO PlayerSupplies (PlayerID, CardID, CrewID) VALUES (" & player.ID & ", " & rst!CardID & ", " & ldr & ")"
            kword = rst!KeyWords & ""
            x = 1
            If ldr = 4 And (InStr(kword, "FIREARM") > 0 Or InStr(kword, "EXPLOSIVES") > 0) Then x = 2
            getMoney player.ID, -1 * (rst!Pay / x)
            PutMsg player.PlayName & " buys " & rst!GearName & " at " & getSupplyName(SupplyID), player.ID, Logic!GameCntr
         End If
         rst.Close
      End If
      
      
End Sub

Private Sub refreshShip(filter, Optional ByVal doClear As Boolean = True)
Dim Index, SQL, w, x, y, z
Dim totalfight, totaltech, totalnego, totalpay, lastplayer, fight As Integer, tech As Integer, nego As Integer
Dim discardF As Boolean, discardT As Boolean, discardN As Boolean, planet As String
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
    
SQL = "SELECT Board.Zones, P.PlanetName, Players.*"
SQL = SQL & " FROM (Board INNER JOIN Players ON Board.SectorID = Players.SectorID) LEFT JOIN (select Planet.SectorID, min(Planet.PlanetName) AS PlanetName FROM Planet  group by Planet.SectorID) P ON Players.SectorID = P.SectorID "
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
      Set .ItemPicture(Index) = AssetImages.Overlay("L", "serenity")
      .CellBackColor(Index, 0) = getPlayerColor(rst!playerID)
      .CellForeColor(Index, 0) = 0
      .ItemLevel(Index) = 0
      .CellText(Index, 1) = rst!ship & " [AI]"  ' PlayCode(rst!playerID).PlayName & IIf(rst!playerID = player.ID, " [AI]", "")
      .CellFont(Index, 1).Name = "BankGothic Md BT"
      .CellForeColor(Index, 1) = 0
      .CellBackColor(Index, 1) = getPlayerColor(rst!playerID)
      If Logic!player = rst!playerID Then
         .CellText(Index, 2) = " << IN PLAY >> $" & rst!Pay
         Set .CellPicture(Index, 2) = AssetImages.ListImages("dino").Picture
      Else
         .CellText(Index, 2) = "Cash in Hand: $" & rst!Pay
      End If
         
      .CellForeColor(Index, 2) = 0
      .CellBackColor(Index, 2) = getPlayerColor(rst!playerID)
      .CellFont(Index, 2).Name = "BankGothic Md BT"
      .CellText(Index, 3) = "Warrants: " & CStr(rst!Warrants)
      If rst!Warrants > 0 Then
         .CellBackColor(Index, 3) = 3355647
      End If
      
      planet = rst!PlanetName & ""
      If planet = "" Or planet = "Cruiser" Or planet = "Corvette" Then
          .CellText(Index, 4) = "Sector " & CStr(rst!SectorID)
      Else
         .CellText(Index, 4) = planet
      End If
'
'      If Nz(rst!PlanetName, "Cruiser") = "Cruiser" Or Nz(rst!PlanetName, "Corvette") = "Corvette" Then
'         .CellText(Index, 4) = "Sector " & CStr(rst!SectorID)
'      Else
'         .CellText(Index, 4) = rst!PlanetName
'      End If
      .CellItemData(Index, 4) = rst!playerID
      .CellItemData(Index, 6) = rst!SectorID
      If rst!Zones = "B" Then
         .CellBackColor(Index, 4) = 0
      ElseIf rst!Zones = "R" Then
         .CellBackColor(Index, 4) = 79
      Else
         .CellBackColor(Index, 4) = 16711680
      End If
      .CellText(Index, 9) = "Goals: " & CStr(rst!Goals) & " Turns: " & CStr(Logic!GameCntr - 1) & IIf(isBountyEnabled, " Bounties: " & CStr(countBounties(player.ID)), "")
      
      'CREW---------------------------------------------
      Index = .AddItem("Crew")
      .CellFont(Index, 0).Name = "BankGothic Md BT"

      'Display actual Crew Number and Capacity (6) with modifiers
      x = CrewCapacity(rst!playerID)
      y = getCrewCount(rst!playerID)
      .CellText(Index, 2) = "Crew Cap: " & CStr(x) & " Crew: " & CStr(y) & "  Spare: " & CStr(x - y)
      .CellFont(Index, 2).Name = "BankGothic Md BT"
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
         .CellItemData(Index, 8) = rst2!Pay
         .ItemLevel(Index) = 2
         'set Crew's Avatar
         If rst2!OffJob = 1 Then
            Set .ItemPicture(Index) = AssetImages.Overlay(findImageKey(rst2!Picture), IIf(rst2!leader = 1, "LD", "O"))  '"L"
         ElseIf findImageKey(rst2!Picture) > 0 Then
            'Set .ItemPicture(Index) =  LoadPicture(App.Path & "\Pictures\Sm" & rst2!Picture)
            Set .ItemPicture(Index) = AssetImages.ListImages(findImageKey(rst2!Picture)).Picture
         Else
            Set .ItemPicture(Index) = AssetImages.Overlay("L", "P")
         End If
'         If rst2!leader = 1 Then
'            Set .ItemPicture(Index) = LoadPicture(App.Path & "\Pictures\Sm" & rst2!Picture)
'         ElseIf rst2!OffJob = 0 Then
'            Set .ItemPicture(Index) = AssetImages.Overlay("L", IIf(rst2!leader = 1, "LD", "P"))
'         Else
'            Set .ItemPicture(Index) = AssetImages.Overlay("L", IIf(rst2!leader = 1, "LD", "O"))
'         End If

         .CellText(Index, 1) = rst2!CrewName & "  -  " & rst2!CrewDescr

         .CellText(Index, 2) = rst2!PerkDescription
         
         .CellText(Index, 3) = Trim(IIf(rst2!Mechanic = 1, "Mechanic  ", "") & IIf(rst2!Pilot = 1, "Pilot  ", "") & IIf(rst2!Companion = 1 Or hasGearCrew(rst!playerID, 36) = rst2!CrewID, "Companion  ", "") & _
               IIf(rst2!Merc = 1, "Merc  ", "") & IIf(rst2!Soldier = 1, "Soldier  ", "") & IIf(rst2!HillFolk = 1, "HillFolk  ", "") & _
               IIf(rst2!Grifter = 1, "Grifter ", "") & IIf(rst2!Medic = 1, "Medic ", "") & IIf(rst2!Mudder = 1, "Mudder ", "") & IIf(rst2!Lawman = 1, "Lawman", ""))
         .CellForeColor(Index, 3) = 65280
         '.CellBackColor(Index, 3) = 6553600
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
         If rst2!CrewID = 94 And getZone(rst!SectorID) = "B" Then 'Sheriff Bourne
            fight = fight + 2
            .CellFont(Index, 5).Bold = True
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
         
         .CellText(Index, 8) = IIf(rst2!leader = 1, "Leader ", "$" & CStr(rst2!Pay))
         If rst2!leader = 0 Then
            .CellBackColor(Index, 8) = 8388736
            .CellForeColor(Index, 8) = 16777215
         End If
         If rst2!OffJob = 0 Then
            totalpay = totalpay + rst2!Pay
         Else
            .CellFont(Index, 8).Strikethrough = True
         End If
         
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
         rst3.CursorLocation = adUseClient
         rst3.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
         While Not rst3.EOF
            Index = .AddItem(CStr(rst3!CardID))
            .CellItemData(Index, 0) = 2 'gear
            .CellItemData(Index, 1) = rst3!CardID
            .CellItemData(Index, 2) = rst3!GearID
            .CellItemData(Index, 4) = rst!playerID
            .CellItemData(Index, 5) = rst2!CrewID
            .ItemLevel(Index) = 3
            
            If InStr(rst3!GearName, "Charts") > 0 Or InStr(rst3!GearName, "Contract") > 0 Then
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
            If rst3!Discard = 1 And rst3!fight > 0 Then
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
            If rst3!Discard = 1 And rst3!tech > 0 Then
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
            If rst3!Discard = 1 And rst3!Negotiate > 0 Then
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
      .CellFont(Index, 0).Name = "BankGothic Md BT"
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
      .CellFont(y, 0).Name = "BankGothic Md BT"
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
      w = CargoCapacity(rst!playerID)
      x = CargoSpaceUsed(rst!playerID)
      .CellText(y, 2) = "Cargo Cap: " & w & ",  Cargo: " & CStr(x) & "  Spare: " & CStr((w - x))
      .CellFont(y, 2).Name = "BankGothic Md BT"
      If (w - CargoSpaceUsed(rst!playerID)) < 1 Then .CellForeColor(y, 2) = QBColor(12)
      
      If z = y Then .Collapse y, True
      rst2.Close
      'SHIP UPDGRADES-----------------------------------
      y = .AddItem("Drive Core & Ship Upgrades")
      .CellFont(y, 0).Name = "BankGothic Md BT"
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
         .ItemLevel(Index) = 2
         Set .ItemPicture(Index) = AssetImages.Overlay("LN", IIf(rst2!DriveCore = 1, "SU", "UP"))
         rst2.MoveNext
      Wend
      If z = y Then .Collapse y, True
      rst2.Close
      w = getShipUpgrades(rst!playerID)
      .CellText(y, 2) = "Upgrade Slots Spare: " & (3 - w)
      .CellFont(y, 2).Name = "BankGothic Md BT"
      If w > 2 Then .CellForeColor(y, 2) = QBColor(12)
      '--------------------------------------------------
      rst.MoveNext
   Wend
   
 End With
   
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
    
With sftTree2
   If doClear Then .Clear  'otherwise Append
   'add the Player details
   rst3.CursorLocation = adUseClient
   rst3.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst3.EOF
      Index = .AddItem(CStr(rst3!playerID) & IIf(isOutlaw(rst3!playerID), " - outlaw", ""))
      .ItemLevel(Index) = 0
      .CellText(Index, 1) = rst3!ship & " [AI]" ' PlayCode(rst3!playerID).PlayName & IIf(rst3!playerID = player.ID, " [AI]", "")
      .CellFont(Index, 1).Name = "BankGothic Md BT"
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
      
      rst.CursorLocation = adUseClient
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
         .CellText(Index, 4) = "$" & rst!Pay
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

Private Sub doShowdownDefend(ByVal AttackerID)
Dim Skilltype As Integer, skill As Integer, Dice As Integer
Dim x As Integer, y As Integer

   For x = 1 To 3
      y = getSkill(player.ID, cstrSkill(x), 1)
      If y > skill Then
         Skilltype = x
         skill = y
      End If
   Next x

   Dice = RollDice(6, True)
   If hasCrew(player.ID, 90) And Dice = 1 Then
      Dice = RollDice(6, True)
      PutMsg player.PlayName & " rolls a 1 and uses The Guardian to re-roll a " & Dice, player.ID, Logic!GameCntr
   End If

   DB.Execute "Insert into ShowdownScores (PlayerID,SkillType,Skill,Dice) Values (" & player.ID & "," & Skilltype & "," & skill & "," & Dice & ")"
   Logic.Requery
   Logic!ClientAccept = 1
   Logic!Trader = 0
   Logic.Update

End Sub

Private Sub doShowdownAttack(ByVal DefenderID)
Dim Skilltype As Integer, skill As Integer, Dice As Integer
Dim x As Integer, y As Integer

   skill = getMySkill(Skilltype)
'   For x = 1 To 3
'      y = getSkill(player.ID, cstrSkill(x), 1)
'      If y > skill Then
'         Skilltype = x
'         skill = y
'      End If
'   Next x

   Dice = RollDice(6, True)
   If hasCrew(player.ID, 90) And Dice = 1 Then
      Dice = RollDice(6, True)
      PutMsg player.PlayName & " rolls a 1 and uses The Guardian to re-roll a " & Dice, player.ID, Logic!GameCntr
   End If
   DB.Execute "Insert into ShowdownScores (PlayerID,SkillType,Skill,Dice) Values (" & player.ID & "," & Skilltype & "," & skill & "," & Dice & ")"
'   Logic.Requery
'   Logic!HostAccept = 1
'   Logic.Update

End Sub

Private Function getMySkill(ByRef Skilltype As Integer) As Integer
Dim x, y
   For x = 1 To 3
      y = getSkill(player.ID, cstrSkill(x), 1)
      If y > getMySkill Then
         Skilltype = x
         getMySkill = y
      End If
   Next x
End Function

Private Function doShowDownSupply(ByVal SupplyID, ByVal CardID As Integer) As Boolean
Dim Skilltype As Integer, skill As Integer, Dice As Integer
Dim x As Integer, y As Integer, FugitiveID As Integer, CrewCardID As Integer, killCrew As Integer
Dim CSkilltype As Integer, Cskill As Integer, CDice As Integer, cnt, msg, win As Boolean

   FugitiveID = varDLookup("FugitiveID", "ContactDeck", "CardID = " & CardID)
   CrewCardID = getCrewCardID(FugitiveID)

   skill = getMySkill(Skilltype)
   '   For x = 1 To 3
   '      y = getSkill(player.ID, cstrSkill(x), 1)
   '      If y > skill Then
   '         Skilltype = x
   '         skill = y
   '      End If
   '   Next x

   Dice = RollDice(6, True)
   
   'pick highest skill of Fugitive
   For x = 1 To 3
      cnt = getSkillCrew(FugitiveID, cstrSkill(x))
      If cnt > CSkilltype Then
         CSkilltype = cnt
         Cskill = x
      End If
   Next x
      
      'Crazy River Tam (cardID 51/CrewID 32)
   If FugitiveID = 32 Then
      CDice = RollDice(6)
      Select Case CDice
      Case 3, 6 'fight
         CSkilltype = 1
         Cskill = 2
      Case 4 'Tech
         CSkilltype = 2
         Cskill = 2
      Case 5 'negot
         CSkilltype = 3
         Cskill = 2
      End Select
   End If
   
   If CSkilltype = 0 Then CSkilltype = 3
   'fugitive rolls
   CDice = RollDice(6, True)
   
   PutMsg "Showdown: " & getCrewName(0, FugitiveID) & IIf(Cskill = 0, " has no Skills", " uses the " & cstrSkill(CSkilltype) & " skill of " & Cskill) & " and rolls a " & CDice & " for a total of " & CStr(Cskill + CDice), player.ID, Logic!GameCntr
    
   win = ((skill + Dice) > (Cskill + CDice))
   
   msg = "Showdown: " & player.PlayName & " uses the " & cstrSkill(Skilltype) & " skill of " & skill & " and rolls a " & Dice & " for a total of " & CStr(skill + Dice) & IIf(win, " to apprehend " & getCrewName(0, FugitiveID), " and Botches the job!")
   PutMsg msg, player.ID, Logic!GameCntr
   
   If win Then
      DB.Execute "UPDATE SupplyDeck SET Seq =0 WHERE CardID = " & CrewCardID
      assignDeal player.ID, CardID
      doShowDownSupply = True
      'pullBounties
      If DrawDeck("Contact", 10, 1) Then PutMsg "New Bounty available"
   Else 'lost
      killCrew = varDLookup("FailKillCrew", "ContactDeck", "CardID=" & bountyCardID)
      If killCrew > 0 Then
         doKillCrew killCrew
      End If
   End If
   
   
End Function

Private Function doBoardingTest(ByVal DefenderID As Integer) As Boolean
Dim Skilltype As Integer, skill As Integer, Dice As Integer
Dim x As Integer, y As Integer

   For x = 2 To 3
      y = getSkill(player.ID, cstrSkill(x))
      If y > skill Then
         Skilltype = x
         skill = y
      End If
   Next x

   Dice = RollDice(6, True)
   If Dice + skill > 5 Then
      doBoardingTest = True
      DB.Execute "Delete from ShowdownScores"
      DB.Execute "Delete from ShowdownGear"
      'throw a hook for the Rival Ship to pickup the boarding alert and init the showdown
      Logic.Requery
      Logic!Seq = "F"
      Logic!HostAccept = 0
      Logic!ClientAccept = 0
      Logic!Trader = DefenderID
      Logic.Update
   End If
      
   PutMsg player.PlayName & IIf(doBoardingTest, " boards", " fails to board") & " " & PlayCode(DefenderID).PlayName & "'s Ship using " & skill & " " & cstrSkill(Skilltype) & " skills and a dice roll of " & Dice, player.ID, Logic!GameCntr

End Function

Private Function processBountyJump(ByVal bountyCardID As Integer) As Boolean
Dim DefenderID As Integer, killCrew, FugitiveID As Integer, CardID As Integer
'Dim rst As New ADODB.Recordset,SQL As String
'Dim ASkill, DSkill, ADice, DDice
Dim msg As String

   If bountyCardID = 0 Then Exit Function
   

   If processBountyJumpScore(DefenderID) Then
      msg = " wins"
      
      'is this a claimed bounty or a crew member?
      FugitiveID = varDLookup("FugitiveID", "ContactDeck", "CardID = " & bountyCardID)
      If hasCrew(DefenderID, FugitiveID) Then  'a crew member
         CardID = getCrewCardID(FugitiveID)
         'update their pile status - 0 removed, 5 -discarded
         DB.Execute "UPDATE SupplyDeck SET Seq =0 WHERE CardID = " & CardID
         'remove any Gear first
         DB.Execute "UPDATE PlayerSupplies SET CrewID = 0 WHERE CrewID = " & FugitiveID
         'delete the card to the players deck
         DB.Execute "DELETE FROM PlayerSupplies WHERE PlayerID =" & DefenderID & " AND CardID = " & CardID
         'clear disgruntled
         DB.Execute "UPDATE Crew SET Disgruntled = 0 WHERE CrewID = " & FugitiveID
         'pullBounties
         If DrawDeck("Contact", 10, 1) Then PutMsg "New Bounty available"
         
      Else 'claimed bounty
         'remove the card to the rival players deck
         DB.Execute "DELETE FROM PlayerJobs WHERE CardID  = " & bountyCardID
      End If
      'take the bounty card
      assignDeal player.ID, bountyCardID
      processBountyJump = True
      'Agent McGinnis is a bad loser
      If hasCrew(DefenderID, 92) Then issueWarrant player.ID, DefenderID
      'PutMsg PlayCode(DefenderID).PlayName & "s Agent McGinnis tries to issue a Warrant, but they don't apply to Robots!", DefenderID, Logic!Gamecntr
     
   Else
      msg = " has lost"
      killCrew = varDLookup("FailKillCrew", "ContactDeck", "CardID=" & bountyCardID)
      If killCrew > 0 Then
         doKillCrew killCrew
      End If
      'Agent McGinnis is a bad loser
      If hasCrew(player.ID, 92) Then issueWarrant DefenderID, player.ID
   End If
   PutMsg player.PlayName & msg & " the Showdown against " & PlayCode(DefenderID).PlayName, player.ID, Logic!GameCntr


End Function

Private Function processBountyJumpScore(ByRef DefenderID As Integer) As Boolean
Dim SQL As String
Dim rst As New ADODB.Recordset
Dim ASkill, DSkill, ADice, DDice
  
   SQL = "SELECT * FROM ShowdownScores"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      If rst!playerID = player.ID Then
         ASkill = rst!skill
         ADice = rst!Dice
      Else
         DefenderID = rst!playerID
         DSkill = rst!skill
         DDice = rst!Dice
      End If
      rst.MoveNext
   Wend
   rst.Close

   If ASkill + ADice > DSkill + DDice Then
      processBountyJumpScore = True
   End If
End Function

Private Sub issueWarrant(ByVal takerID As Integer, ByVal giverID As Integer)

   DB.Execute "UPDATE Players set Warrants = Warrants + 1 WHERE PlayerID = " & CStr(takerID)
   PutMsg "Agent McGinnis didn't take kindly to the Showdown result and has issued a Warrant to " & PlayCode(takerID).PlayName, giverID, Logic!GameCntr
   
End Sub

Private Function findBountyJump(ByVal SectorID As Integer, ByRef DefenderID As Integer, ByRef CardID As Integer) As Integer
Dim SQL As String, dist As Integer, RSectorID As Integer
Dim rst As New ADODB.Recordset
Dim x
   dist = 1000
   
   'return Players with a Claimed Fugitive or Crew with a bounty on them
'   SQL = "SELECT Players.PlayerID, Players.SectorID, ContactDeck.CardID "
'   SQL = SQL & "FROM Players INNER JOIN (PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) ON Players.PlayerID = PlayerJobs.PlayerID"
'   SQL = SQL & " WHERE ContactDeck.ContactID=10 AND PlayerJobs.JobStatus=0 ORDER BY Players.Pay DESC"
   SQL = "SELECT Players.PlayerID, Players.SectorID, ContactDeck.CardID, Players.Pay, ContactDeck.Pay"
   SQL = SQL & " FROM ContactDeck INNER JOIN ((Players INNER JOIN PlayerSupplies ON Players.PlayerID = PlayerSupplies.PlayerID) INNER JOIN SupplyDeck ON PlayerSupplies.CardID = SupplyDeck.CardID) ON ContactDeck.FugitiveID = SupplyDeck.CrewID"
   SQL = SQL & " Where ContactDeck.ContactID = 10 And ContactDeck.Seq = 5 AND Players.PlayerID <> " & player.ID
   SQL = SQL & " Union"
   SQL = SQL & " SELECT Players.PlayerID, Players.SectorID, ContactDeck.CardID, Players.Pay, ContactDeck.Pay"
   SQL = SQL & " FROM Players INNER JOIN (PlayerJobs INNER JOIN ContactDeck ON PlayerJobs.CardID = ContactDeck.CardID) ON Players.PlayerID = PlayerJobs.PlayerID"
   SQL = SQL & " Where ContactDeck.ContactID = 10 And PlayerJobs.JobStatus = 0 AND Players.PlayerID <> " & player.ID
   SQL = SQL & " ORDER BY 3 DESC, 4 DESC" 'target player with most money and best bounty$$
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      x = getSectorCount(SectorID, rst!SectorID)
      If x < dist Then
         dist = x
         RSectorID = rst!SectorID
         DefenderID = rst!playerID
         CardID = rst!CardID
      End If
      rst.MoveNext
   Wend
   rst.Close
   If dist = 0 Or (dist < 6 - FullburnMovesDone And Not fullburndone) Then
      findBountyJump = RSectorID
   End If


End Function

Private Function findSupplyBounty(ByVal SectorID As Integer, ByRef SupplyID As Integer, ByRef CardID As Integer) As Integer
Dim SQL As String, dist As Integer, RSectorID As Integer
Dim rst As New ADODB.Recordset
Dim x
   dist = 1000
   
      SQL = "SELECT ContactDeck.CardID, SupplyDeck.SupplyID, Supply.SectorID"
   SQL = SQL & " FROM Supply INNER JOIN (ContactDeck INNER JOIN SupplyDeck ON ContactDeck.FugitiveID = SupplyDeck.CrewID) ON Supply.SupplyID = SupplyDeck.SupplyID"
   SQL = SQL & " WHERE ContactDeck.Seq=5 AND ContactDeck.ContactID=10 AND SupplyDeck.Seq=5"
   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   While Not rst.EOF
      x = getSectorCount(SectorID, rst!SectorID)
      If x < dist Then
         dist = x
         RSectorID = rst!SectorID
         SupplyID = rst!SupplyID
         CardID = rst!CardID
      End If
      rst.MoveNext
   Wend
   rst.Close
   If dist = 0 Or (dist < 6 - FullburnMovesDone And Not fullburndone) Then
      findSupplyBounty = RSectorID
   Else
      SupplyID = 0
      CardID = 0
   End If
   
   
End Function

Private Sub refreshSolid()
Dim x
      For x = 1 To NO_OF_CONTACTS
         Imag(x).Visible = isSolid(player.ID, x)
      Next x

End Sub

Private Function ClearTrail()
Dim x
   For x = 0 To 8
      Trail(x) = 0
   Next x
End Function


