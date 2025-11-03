VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form Board 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "the 'Verse"
   ClientHeight    =   15150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   22710
   Icon            =   "Board.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   15150
   ScaleWidth      =   22710
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   15060
      Left            =   0
      ScaleHeight     =   15000
      ScaleWidth      =   22575
      TabIndex        =   0
      Top             =   0
      Width           =   22635
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   300
         Index           =   6
         Left            =   16920
         Top             =   2190
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Effects         =   "Board.frx":030A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   750
         Index           =   5
         Left            =   17970
         Top             =   2190
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   1323
         Effects         =   "Board.frx":0322
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   330
         Index           =   10
         Left            =   3000
         Top             =   270
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         Effects         =   "Board.frx":033A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   330
         Index           =   11
         Left            =   1890
         Top             =   2040
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         Effects         =   "Board.frx":0352
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   330
         Index           =   12
         Left            =   570
         Top             =   3360
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         Effects         =   "Board.frx":036A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   330
         Index           =   9
         Left            =   5550
         Top             =   60
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         Effects         =   "Board.frx":0382
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   330
         Index           =   8
         Left            =   4770
         Top             =   60
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         Effects         =   "Board.frx":039A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   330
         Index           =   7
         Left            =   3990
         Top             =   60
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         Effects         =   "Board.frx":03B2
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   540
         Index           =   1
         Left            =   18570
         Top             =   30
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   953
         Effects         =   "Board.frx":03CA
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   540
         Index           =   2
         Left            =   19470
         Top             =   30
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   953
         Effects         =   "Board.frx":03E2
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   540
         Index           =   3
         Left            =   20340
         Top             =   30
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   953
         Effects         =   "Board.frx":03FA
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   540
         Index           =   4
         Left            =   21240
         Top             =   30
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   953
         Effects         =   "Board.frx":0412
      End
      Begin VB.Label lblSolid 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   9
         Left            =   21930
         TabIndex        =   10
         Top             =   7440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSolid 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   8
         Left            =   21930
         TabIndex        =   9
         Top             =   7170
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSolid 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   7
         Left            =   21930
         TabIndex        =   8
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSolid 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   6
         Left            =   21930
         TabIndex        =   7
         Top             =   6540
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSolid 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   21930
         TabIndex        =   6
         Top             =   6210
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSolid 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   21930
         TabIndex        =   5
         Top             =   5850
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSolid 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   21930
         TabIndex        =   4
         Top             =   5490
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSolid 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   21930
         TabIndex        =   3
         Top             =   5190
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSolid 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Solid"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   21930
         TabIndex        =   2
         Top             =   4890
         Visible         =   0   'False
         Width           =   615
      End
      Begin XDOCKFLOATLibCtl.FDPane FDPane1 
         Height          =   420
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Visible         =   0   'False
         Width           =   420
         _cx             =   741
         _cy             =   741
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
      Begin VB.Line LineB 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   2
         Visible         =   0   'False
         X1              =   11730
         X2              =   13950
         Y1              =   9480
         Y2              =   9690
      End
      Begin VB.Line LineB 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   1
         Visible         =   0   'False
         X1              =   11790
         X2              =   14010
         Y1              =   8640
         Y2              =   8850
      End
      Begin VB.Line LineB 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   11820
         X2              =   14040
         Y1              =   7830
         Y2              =   8040
      End
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgAToken 
      Height          =   705
      Index           =   0
      Left            =   10000
      Top             =   15000
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1244
      Effects         =   "Board.frx":042A
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgHaven 
      Height          =   705
      Index           =   0
      Left            =   15000
      Top             =   15060
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1244
      Effects         =   "Board.frx":0442
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgToken 
      Height          =   705
      Index           =   0
      Left            =   19380
      Top             =   15060
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1244
      Effects         =   "Board.frx":045A
   End
   Begin VB.Image HotSpot 
      Height          =   735
      Index           =   0
      Left            =   0
      MouseIcon       =   "Board.frx":0472
      MousePointer    =   4  'Icon
      Top             =   14500
      Width           =   555
   End
End
Attribute VB_Name = "Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public isLoaded As Boolean
Public Event SectClick(ByVal Index As Integer)


Private Sub Form_Load()
   isLoaded = True
   initImages
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      MsgBox "Probably not a good idea to close the Map", vbInformation

   End If
End Sub

Private Sub HotSpot_Click(Index As Integer)
  ' MsgBox "You clicked on " & Index
    RaiseEvent SectClick(Index)
     
End Sub

Private Sub Imag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
'allow title drag
Private Sub HotSpot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If pickStartSector <> 1 Then
      Select Case actionSeq
         Case ASmosey, ASfullburn, ASNavEvade, ASNavReav, ASNavReavBorder, ASNavCrus, ASNavCrusBorder, ASNavCrusOutlaw, ASNavCrusAdjacent, ASNavCorvAdjacent, ASNavCorvPlanetary, ASResolveAlert
         Case Else
            ReleaseCapture
            SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
      End Select
   End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub initImages()
Dim x
   Imag(1).Picture = LoadPictureGDIplus(App.Path & "\Pictures\FireflyOrange.gif")
   Imag(2).Picture = LoadPictureGDIplus(App.Path & "\Pictures\FireflyBlue.gif")
   Imag(3).Picture = LoadPictureGDIplus(App.Path & "\Pictures\FireflyYellow.gif")
   Imag(4).Picture = LoadPictureGDIplus(App.Path & "\Pictures\FireflyGreen.gif")
   Imag(5).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Cruiser.gif")
   Imag(6).Picture = LoadPictureGDIplus(App.Path & "\Pictures\corvette.gif")
   For x = 7 To 12  '6 Reavers
      Imag(x).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Cutter.gif")
   Next x
   For x = 1 To 12   'all ships set invisible to start
      Imag(x).TransparentColor = 0
      Imag(x).TransparentColorMode = lvicUseTransparentColor
   Next x
End Sub
