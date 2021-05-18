VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form Action 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actions"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "Action.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SftTree.SftTree Grid 
      Height          =   1560
      Left            =   2280
      TabIndex        =   24
      Top             =   120
      Width           =   1575
      _Version        =   262144
      _ExtentX        =   2778
      _ExtentY        =   2752
      _StockProps     =   237
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      ItemPictureExpanded=   "Action.frx":0442
      ItemPictureExpandable=   "Action.frx":045E
      ItemPictureLeaf =   "Action.frx":047A
      PlusMinusPictureExpanded=   "Action.frx":0496
      PlusMinusPictureExpandable=   "Action.frx":04B2
      PlusMinusPictureLeaf=   "Action.frx":04CE
      ButtonPicture   =   "Action.frx":04EA
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
      ColHeaderAppearance=   0
      ButtonStyle     =   0
      TreeLineStyle   =   0
      ColStyle0       =   20
      MouseIcon       =   "Action.frx":0506
      SelectionBackColor=   16744703
      Scrollbars      =   0
      RowColHeaderAppearance=   0
      RowColPicture   =   "Action.frx":0522
      LeftButtonOnly  =   0   'False
      RowHeaderAppearance=   0
      BackgroundPicture=   "Action.frx":053E
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   1560
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":055A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":066C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":0890
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":09A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":0AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":0BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":0CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":0DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":0EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":100E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":1120
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":1232
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":1344
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":1456
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":1568
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":167A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":178C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":189E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":19B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":1AC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":1BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":1CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Action.frx":1DF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Set option && move"
      Height          =   1695
      Left            =   2280
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CommandButton cmd 
         Caption         =   "Go"
         DownPicture     =   "Action.frx":1F0A
         Height          =   1335
         Index           =   9
         Left            =   960
         Picture         =   "Action.frx":2574
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Green"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   795
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Yellow"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   795
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Blue"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   795
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Red"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame fraMove 
      Caption         =   "Direction"
      Height          =   1695
      Left            =   2280
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   5
         Left            =   120
         Picture         =   "Action.frx":2BDE
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   495
         Index           =   1
         Left            =   600
         Picture         =   "Action.frx":3020
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   6
         Left            =   1080
         Picture         =   "Action.frx":3462
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   2
         Left            =   120
         Picture         =   "Action.frx":38A4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   3
         Left            =   960
         Picture         =   "Action.frx":3CE6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   7
         Left            =   120
         Picture         =   "Action.frx":4128
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   495
         Index           =   4
         Left            =   600
         Picture         =   "Action.frx":456A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   8
         Left            =   1080
         Picture         =   "Action.frx":49AC
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Carry Bag"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1850
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "End Move"
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Action"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      Begin VB.OptionButton Opt 
         Caption         =   "Mission Orders"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Move onto Barrier"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Move Any Direction"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Diagonal Move"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Straight Move"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2280
      Width           =   2055
   End
End
Attribute VB_Name = "Action"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MoveMode, MovesLeft
Public MyTop, MyLeft

Private Sub cmd_Click(Index As Integer)
'    If Grid.ListCount = 1 Then Grid.Selected(0) = True
'    If Index = 0 Then     'end move
'        Unload Me
'    ElseIf Index = 9 And getGridSelected > 0 Then 'jump to new location
'        For z = 5 To 8
'            If Opt(z).Value Then
'                x = Val(Left(Opt(z).Tag, 2))
'                y = Val(Right(Opt(z).Tag, 2))
'                bittype = getGridSelected
'                If CheckMoveBit(Player.ID, bittype, x, y) Then
'                    MoveBit bittype, x, y, Check1.Value
'                    Unload Me
'                End If
'                Exit For
'            End If
'        Next z
'
'    ElseIf getGridSelected > 0 And Grid.ListCount > 0 Then
'        bittype = getGridSelected
'        If GetBitPos(Player.ID, bittype, x, y) = "-1" Then Exit Sub
'        Select Case Index
'        Case 1
'          a = 0
'          b = -1
'        Case 2
'          a = -1
'          b = 0
'        Case 3
'          a = 1
'          b = 0
'        Case 4
'          a = 0
'          b = 1
'        Case 5
'          a = -1
'          b = -1
'        Case 6
'          a = 1
'          b = -1
'        Case 7
'          a = -1
'          b = 1
'        Case 8
'          a = 1
'          b = 1
'        End Select
'
'        'check for out of bounds
'        If x + a < 1 Or x + a > 11 Or y + b < 1 Or y + b > 11 Then
'            playsnd 8
'            Exit Sub
'        End If
'
'        'check if jumping a barrier and modify a & b
'        CheckMoveBarrier Player.ID, bittype, x, y, a, b
'
'        'set new position
'        x = x + a
'        y = y + b
'
'        'check if move is legal
'        If CheckMoveBit(Player.ID, bittype, x, y) Then
'            MoveBit bittype, x, y, Check1.Value
'            MovesLeft = MovesLeft - 1
'            Select Case MoveMode
'            Case 10
'               removebit bittype
'            Case 20
'               setbits bittype, bittype
'            End Select
'            If MovesLeft = 0 Then
'               Unload Me
'            Else
'               DrawBoard
'            End If
'        End If
'    End If
End Sub

Private Sub Form_Load()
Dim rst As New ADODB.Recordset
 Me.Top = MyTop
 Me.Left = MyLeft
 Me.Caption = PlayCode(Player.ID).PlayName & "'s Turn"
 'List1.BackColor = getPlayerColor(Player.ID)
 
 Randomize
 x = Int((6 * Rnd)) - 1
' If x = -1 Then x = 0 'double chance for straight
' Opt(x).Value = True
' 'msg to move a piece of their choice
' If x < 4 Then  'not mission card
'    setbits 1, 6
'    PutMsg PlayCode(Player.ID).ID & ": " & Opt(x).Caption, Player.ID
'    MovesLeft = 1
' End If
' Select Case x
' Case 4  'mission
'     Randomize
'     x = Int((NumberOfMissions * Rnd)) + 1
'     rst.Open "SELECT * FROM Mission WHERE MissionID = " & x, DB, adOpenStatic, adLockReadOnly
'     Text1 = rst!Mission
'     PutMsg PlayCode(Player.ID).ID & " Mission Order: " & Text1, Player.ID
'     code = rst!result
'     Select Case Left(code, 1)
'     Case "A"  '1 General Mobilisation.  Move every piece of your equipment one square in any direction.
'               'Q Move every piece of your equipment back one square towards your own HQ
'        setbits 1, 6
'        MoveMode = 10
'        MovesLeft = 6
'        Check1.Value = 1
'        playsnd 17
''     Case "B"  'Bribe Card + & -
''        MoveMode = 0
''        MovesLeft = 0
''        setcontols 1, 8
''        Check1.Visible = False
''        If Mid(code, 2, 1) = "+" Then   'add
''            BribeCard Player.ID, 1
''            playsnd 24
''        Else                            ' - remove
''            If BribeCard(Player.ID, 0) > 0 Then
''                BribeCard Player.ID, -1
''            End If
''            playsnd 15
''        End If
'     Case "C"  'Move Passport to checkpoint.
'        MoveMode = 0
'        If CheckMoveBit(Player.ID, 1, 6, 6) Then
'            MovesLeft = 0
'            MoveBit 1, 6, 6, 0
'            setcontols 1, 8
'            Check1.Visible = False
'            DrawBoard
'        Else
'            MovesLeft = 2
'            setbits 4, 4
'            Check1.Value = 1
'        End If
'        playsnd 18
'
'     Case "E"  'Move Nightvision to any Foreign Embassy.
'        MoveMode = 0
'        MovesLeft = 0
'        SetMoveOptions "E"
'        setbits 4, 4
'        setcontols 1, 8
'        Check1.Value = 1
'        playsnd 21
'     Case "F"  'Move any one piece of your equipment on foreign territory two squares.
'        setbits 1, 6
'        MoveMode = 20
'        MovesLeft = 2
'        Check1.Value = 1
'        playsnd 17
'     Case "G"  'Move Dagger5 two spaces in any direction , Move Gun6 two squares in any direction.
'        setbits Val(Mid(code, 2, 1)), Val(Mid(code, 2, 1))
'        MoveMode = 0
'        MovesLeft = 2
'        Check1.Value = 1
'        playsnd 17
'     Case "L"  '2,3 Move Wirecutters adjacent to any enemy barrier.
'               'A     Move either Wirecutters or Ladder onto any barrier
'        MovesLeft = 6
'        If Val(Mid(code, 2, 1)) > 0 Then
'            setbits Val(Mid(code, 2, 1)), Val(Mid(code, 2, 1))
'            MoveMode = 0
'        Else
'            setbits 2, 3
'            MoveMode = 20
'        End If
'        Check1.Visible = False
'        DrawBoard
'        playsnd 26
'     Case "M"  'Miss this turn.
'        MoveMode = 0
'        MovesLeft = 0
'        setcontols 1, 8
'        setgoes -1
'        Check1.Visible = False
'        playsnd 9
'     Case "Q"  '1,2,3, A Return to HQ
'        MoveMode = 0
'        MovesLeft = 0
'        setcontols 1, 8
'        If Val(Mid(code, 2, 1)) > 0 Then
'            MoveBit Val(Mid(code, 2, 1)), 0, 0, 0
'            PlaceNewBit Val(Mid(code, 2, 1))
'            playsnd 22
'        Else
'            MoveBit 1, 0, 0, 0
'            MoveBit 2, 0, 0, 0
'            MoveBit 3, 0, 0, 0
'            MoveBit 4, 0, 0, 0
'            MoveBit 5, 0, 0, 0
'            MoveBit 6, 0, 0, 0
'            PlaceNewBit 6
'            PlaceNewBit 5
'            PlaceNewBit 4
'            PlaceNewBit 3
'            PlaceNewBit 2
'            PlaceNewBit 1
'            playsnd 23
'        End If
'        Check1.Visible = False
'        DrawBoard
'     End Select
' Case 0  'straight
'     setcontols 5, 8     'disable diag
'     Check1.Value = 1
'     playsnd 19
' Case 1  'diag
'     setcontols 1, 4     'disable straight
'     Check1.Value = 1
'     playsnd 17
' Case 2 'any dir
'     Check1.Value = 1
'     playsnd 25
' Case 3 'move onto barrier
'     setbits 2, 3
'     playsnd 26
' End Select
'
End Sub

Private Sub setcontols(ByVal x, ByVal y)
     For z = x To y   'disable diag
       cmd(z).Visible = False
     Next z
End Sub

Private Sub setbits(ByVal x, ByVal y)
 Dim Index
    Grid.Clear
 
 For z = x To y
   Index = Grid.AddItem(Bits(z))
   Grid.ItemData(Index) = z
   Grid.CellBackColor(Index, 0) = getPlayerColor(Player.ID)
   Set Grid.CellPicture(Index, 0) = Images.ListImages((Player.ID - 1) * 6 + z).Picture
 Next z
End Sub

Private Sub removebit(ByVal x)
 If Grid.ListCount = 0 Then Exit Sub
 For z = 0 To Grid.ListCount - 1
    If Grid.ItemData(z) = x Then
       Grid.RemoveItem z
       Exit For
    End If
 Next z

End Sub

Private Sub SetMoveOptions(mode As String)
'    fraOptions.Visible = True
'    fraMove.Enabled = False
'    If mode = "E" Then
'        If CheckMoveBit(Player.ID, 4, 8, 8) Then  'red
'            Opt(5).Tag = "0808"
'        Else
'            Opt(5).Visible = False
'        End If
'        If CheckMoveBit(Player.ID, 4, 4, 8) Then  'blue
'            Opt(6).Tag = "0408"
'        Else
'            Opt(6).Visible = False
'        End If
'        If CheckMoveBit(Player.ID, 4, 4, 4) Then 'yellow
'            Opt(7).Tag = "0404"
'        Else
'            Opt(7).Visible = False
'        End If
'        If CheckMoveBit(Player.ID, 4, 8, 4) Then  'green
'            Opt(8).Tag = "0804"
'        Else
'            Opt(8).Visible = False
'        End If
'    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    MyTop = Me.Top
    MyLeft = Me.Left
End Sub

Private Function getGridSelected() As Integer
   getGridSelected = 0
    For x = 0 To Grid.ListCount - 1
        If Grid.Selected(x) Then
            getGridSelected = Grid.ItemData(x)
            Exit For
        End If
    Next x

End Function
