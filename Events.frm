VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{6ABB9000-48F8-11CF-AC42-0040332ED4E5}#4.0#0"; "SFTTREEX.OCX"
Begin VB.Form Events 
   Caption         =   "Game Events"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Events.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin SftTree.SftTree Grid 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      _Version        =   262144
      _ExtentX        =   3413
      _ExtentY        =   3201
      _StockProps     =   237
      ForeColor       =   -2147483640
      BackColor       =   3355725
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      ItemPictureExpanded=   "Events.frx":0442
      ItemPictureExpandable=   "Events.frx":045E
      ItemPictureLeaf =   "Events.frx":047A
      PlusMinusPictureExpanded=   "Events.frx":0496
      PlusMinusPictureExpandable=   "Events.frx":04B2
      PlusMinusPictureLeaf=   "Events.frx":04CE
      ButtonPicture   =   "Events.frx":04EA
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
      GridStyle       =   2
      ButtonStyle     =   0
      ItemLines       =   10
      TreeLineStyle   =   0
      ColStyle0       =   20
      MouseIcon       =   "Events.frx":0506
      BackColor       =   3355725
      NoFocusStyle    =   2
      RowColHeaderAppearance=   0
      RowColPicture   =   "Events.frx":0522
      LeftButtonOnly  =   0   'False
      RowHeaderAppearance=   0
      ItemStyle       =   1
      BackgroundPicture=   "Events.frx":053E
      CharSearchMode  =   2
      ShowFocusRectangle=   0   'False
      ToolTipBackColor=   -2147483643
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1080
      Top             =   3120
   End
   Begin XDOCKFLOATLibCtl.FDPane FDPane1 
      Height          =   420
      Left            =   3105
      TabIndex        =   0
      Top             =   3060
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
Attribute VB_Name = "Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LastEventNumber As Long

Private Sub Form_Load()
   Grid.BackgroundPicture = LoadPicture(App.Path & "\gui\tileEvents.bmp")
    LastEventNumber = 0
End Sub

Private Sub Form_Resize()
  Grid.Width = Me.Width
  Grid.Height = Me.Height
End Sub

Private Sub Timer1_Timer()
   getNewEvents
End Sub

Public Function getNewEvents() As Boolean
Dim rst As New ADODB.Recordset
Dim X, ship As Boolean, evtDesc As String
   rst.CursorLocation = adUseClient
   rst.Open "SELECT * FROM Events WHERE EventID > " & LastEventNumber & " ORDER BY EventID", DB, adOpenStatic, adLockReadOnly
   While Not rst.EOF
      evtDesc = rst!event
      If Logic!player <> player.ID And Left(evtDesc, 9) = "New Bount" Then
         Main.RefreshDeals
      End If
      X = Grid.InsertItem(0, Replace(evtDesc, "^", " "))
      If Not IsNull(rst!playerID) Then
         'Grid.CellBackColor(X, 0) = getPlayerColor(rst!playerID)
         Grid.CellForeColor(X, 0) = getPlayerColor(rst!playerID)
      Else
         Grid.CellForeColor(X, 0) = QBColor(15)
      End If
      LastEventNumber = rst!EventID
      If Not ship And rst!refreshShip > 0 And actionSeq = ASidle Then
         ship = True
         If Not (Main.frmShip Is Nothing) Then Main.frmShip.RefreshShips
      End If
      rst.MoveNext
   Wend

End Function
