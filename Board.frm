VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form Board 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "the 'Verse"
   ClientHeight    =   15150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   22710
   Icon            =   "Board.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   15150
   ScaleWidth      =   22710
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
         Height          =   465
         Index           =   6
         Left            =   21690
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   820
         Effects         =   "Board.frx":030A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   465
         Index           =   9
         Left            =   20820
         Top             =   5430
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   820
         Effects         =   "Board.frx":0322
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   465
         Index           =   8
         Left            =   20850
         Top             =   4800
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   820
         Effects         =   "Board.frx":033A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   465
         Index           =   7
         Left            =   21720
         Top             =   4110
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   820
         Effects         =   "Board.frx":0352
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   555
         Index           =   1
         Left            =   21780
         Top             =   0
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Effects         =   "Board.frx":036A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   555
         Index           =   2
         Left            =   21780
         Top             =   720
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Effects         =   "Board.frx":0382
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   555
         Index           =   3
         Left            =   21810
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Effects         =   "Board.frx":039A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   555
         Index           =   4
         Left            =   21780
         Top             =   2130
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Effects         =   "Board.frx":03B2
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   765
         Index           =   5
         Left            =   21810
         Top             =   2730
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1349
         Effects         =   "Board.frx":03CA
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
         Left            =   20820
         TabIndex        =   1
         Top             =   150
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
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgAToken 
      Height          =   705
      Index           =   0
      Left            =   17310
      Top             =   15090
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1244
      Effects         =   "Board.frx":03E2
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
      Effects         =   "Board.frx":03FA
   End
   Begin VB.Image HotSpot 
      Height          =   735
      Index           =   0
      Left            =   0
      MouseIcon       =   "Board.frx":0412
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

'For use with USER32 Function SetWindowPos
Private Const HWND_TOPMOST = -&H1
Private Const HWND_NOTOPMOST = -&H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
'For use with USER32 Function SendMessage
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Event SectClick(ByVal Index As Integer)


Private Sub Form_Load()
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
   Imag(1).Picture = LoadPictureGDIplus(App.Path & "\Pictures\FireflyOrange.bmp")
   Imag(2).Picture = LoadPictureGDIplus(App.Path & "\Pictures\FireflyBlue.bmp")
   Imag(3).Picture = LoadPictureGDIplus(App.Path & "\Pictures\FireflyYellow.bmp")
   Imag(4).Picture = LoadPictureGDIplus(App.Path & "\Pictures\FireflyGreen.bmp")
   Imag(5).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Crusier.bmp")
   Imag(5).AutoSize = lvicMultiAngle
   Imag(6).Picture = LoadPictureGDIplus(App.Path & "\Pictures\corvette.bmp")
   Imag(7).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Cutter.bmp")
   Imag(8).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Cutter.bmp")
   Imag(9).Picture = LoadPictureGDIplus(App.Path & "\Pictures\Cutter.bmp")
   For x = 1 To 9
      Imag(x).TransparentColor = 0
      Imag(x).TransparentColorMode = lvicUseTransparentColor
   Next x
End Sub
