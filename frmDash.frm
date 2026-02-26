VERSION 5.00
Begin VB.Form frmDash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   -30
      Width           =   345
   End
   Begin VB.Label lblDistance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   495
      Left            =   1740
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label lblToken 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1920
      TabIndex        =   3
      Top             =   2550
      Width           =   855
   End
   Begin VB.Label lblAToken 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   435
      Left            =   510
      TabIndex        =   2
      Top             =   2550
      Width           =   855
   End
   Begin VB.Label lblPlanet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Melbourne"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   435
      Left            =   330
      TabIndex        =   1
      Top             =   1660
      Width           =   2595
   End
   Begin VB.Label lblSector 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   495
      Left            =   330
      TabIndex        =   0
      Top             =   540
      Width           =   1215
   End
End
Attribute VB_Name = "frmDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SectorID As Integer, Distance As Integer
Public PlanetName As String
Public Atoken As Integer, Token As Integer

Private Sub cmd_Click()
   Unload Me
End Sub

'Public Property Let AlwaysOnTop(bState As Boolean)
'    Dim lFlag As Long
'    On Error Resume Next
'    If bState = True Then
'        lFlag = HWND_TOPMOST
'    Else
'        lFlag = HWND_NOTOPMOST
'    End If
'    bOnTopState = bState
'    SetWindowPos Me.hwnd, lFlag, 0&, 0&, 0&, 0&, (SWP_NOSIZE Or SWP_NOMOVE)
'End Property

Private Sub Form_Load()
Dim rgn As Long
Dim w As Long, h As Long

   ' Convert twips ? pixels
   w = Me.ScaleX(Me.Width, vbTwips, vbPixels)
   h = Me.ScaleY(Me.Height, vbTwips, vbPixels)
   
   ' 40,40 = corner roundness
   rgn = CreateRoundRectRgn(0, 0, w, h, 20, 20)
   
   SetWindowRgn Me.hwnd, rgn, True
   
   Me.Picture = LoadPicture(App.Path & "\gui\dash.jpg")
   
   RefreshDash
End Sub

Public Sub RefreshDash()
   lblSector = CStr(SectorID)
   lblDistance = CStr(Distance)
   lblPlanet = PlanetName
   lblAToken = CStr(Atoken)
   lblToken = CStr(Token)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub lblPlanet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
