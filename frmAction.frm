VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmAction 
   BackColor       =   &H00000000&
   Caption         =   "Actions"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4515
   Icon            =   "frmAction.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   4515
   Begin VB.Timer timScroll 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   4380
      Top             =   7440
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgWorkDrop 
      Height          =   255
      Left            =   2780
      ToolTipText     =   "click here for more jobs"
      Top             =   6390
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   450
      Attr            =   640
      Effects         =   "frmAction.frx":030A
   End
   Begin XDOCKFLOATLibCtl.FDPane FDPane1 
      Height          =   420
      Left            =   4380
      TabIndex        =   57
      Top             =   7860
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
   Begin LaVolpeAlphaImg.AlphaImgCtl imgCancel 
      Height          =   510
      Left            =   500
      Top             =   7830
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   900
      Attr            =   640
      Effects         =   "frmAction.frx":0322
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgEndTurn 
      Height          =   510
      Left            =   2340
      Top             =   7830
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   900
      Attr            =   640
      Effects         =   "frmAction.frx":033A
   End
   Begin VB.Image imgBonus 
      Height          =   405
      Left            =   0
      Top             =   7410
      Width           =   900
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgResolve 
      Height          =   405
      Left            =   2610
      Top             =   7410
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Attr            =   640
      Effects         =   "frmAction.frx":0352
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgMorale 
      Height          =   405
      Left            =   900
      Top             =   7410
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      Attr            =   640
      Effects         =   "frmAction.frx":036A
   End
   Begin VB.Label lblMakeWorkVal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0003E631&
      Height          =   195
      Left            =   3330
      TabIndex        =   56
      Top             =   6390
      UseMnemonic     =   0   'False
      Width           =   705
   End
   Begin VB.Label lblJobName 
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   180
      Left            =   1080
      TabIndex        =   55
      Top             =   6390
      UseMnemonic     =   0   'False
      Width           =   1860
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgMakeWork 
      Height          =   375
      Left            =   3150
      Top             =   6020
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      Attr            =   640
      Effects         =   "frmAction.frx":0382
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgWorkLocal 
      Height          =   375
      Left            =   900
      Top             =   6020
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   661
      Attr            =   640
      Effects         =   "frmAction.frx":039A
   End
   Begin VB.Label lblActiveJobs 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   3820
      TabIndex        =   54
      Top             =   5720
      Width           =   615
   End
   Begin VB.Label lblInactiveJobs 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   1780
      TabIndex        =   53
      Top             =   5720
      Width           =   615
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgClearWarrantsOpt 
      Height          =   150
      Left            =   3850
      ToolTipText     =   "Clear All Warrants as part of the Deal"
      Top             =   5500
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   265
      Attr            =   640
      Effects         =   "frmAction.frx":03B2
   End
   Begin VB.Label lblDealLoadFugi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   3830
      TabIndex        =   51
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label lblDealLoadFugi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   4110
      TabIndex        =   50
      Top             =   5300
      Width           =   195
   End
   Begin VB.Label lblDealLoadPassngr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   3830
      TabIndex        =   48
      Top             =   5115
      Width           =   195
   End
   Begin VB.Label lblDealLoadPassngr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   4110
      TabIndex        =   47
      Top             =   5130
      Width           =   195
   End
   Begin VB.Label lblDealSellParts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   2340
      TabIndex        =   45
      Top             =   5450
      Width           =   195
   End
   Begin VB.Label lblDealSellParts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   2625
      TabIndex        =   44
      Top             =   5460
      Width           =   195
   End
   Begin VB.Label lblDealSellContra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   2340
      TabIndex        =   42
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label lblDealSellContra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   2625
      TabIndex        =   41
      Top             =   5300
      Width           =   195
   End
   Begin VB.Label lblDealSellCargo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   2340
      TabIndex        =   39
      Top             =   5115
      Width           =   195
   End
   Begin VB.Label lblDealSellCargo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   2625
      TabIndex        =   38
      Top             =   5130
      Width           =   195
   End
   Begin VB.Label lblDealBuyFuel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   870
      TabIndex        =   36
      Top             =   5450
      Width           =   195
   End
   Begin VB.Label lblDealBuyFuel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   1160
      TabIndex        =   35
      Top             =   5460
      Width           =   195
   End
   Begin VB.Label lblDealBuyContra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   870
      TabIndex        =   33
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label lblDealBuyContra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   1160
      TabIndex        =   32
      Top             =   5300
      Width           =   195
   End
   Begin VB.Label lblDealBuyCargo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   870
      TabIndex        =   30
      Top             =   5115
      Width           =   195
   End
   Begin VB.Label lblDealBuyCargo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   1160
      TabIndex        =   29
      Top             =   5130
      Width           =   195
   End
   Begin VB.Image imgClearWarrants 
      Height          =   150
      Left            =   2970
      ToolTipText     =   "Clear All Warrants"
      Top             =   5505
      Width           =   870
   End
   Begin VB.Image imgLoadFugi 
      Height          =   165
      Left            =   3080
      Top             =   5280
      Width           =   780
   End
   Begin VB.Image imgLoadPassngr 
      Height          =   150
      Left            =   2880
      Top             =   5130
      Width           =   975
   End
   Begin VB.Image imgDealCargo 
      Height          =   150
      Left            =   1580
      Top             =   5130
      Width           =   750
   End
   Begin VB.Image imgDealContra 
      Height          =   165
      Left            =   1320
      Top             =   5280
      Width           =   1065
   End
   Begin VB.Image imgDealParts 
      Height          =   195
      Left            =   1740
      ToolTipText     =   "Sell Parts for $300ea"
      Top             =   5455
      Width           =   645
   End
   Begin VB.Image imgDealFuel 
      Height          =   195
      Left            =   1320
      Top             =   5455
      Width           =   420
   End
   Begin VB.Label lblHoldSpace 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   30
      TabIndex        =   28
      Top             =   7110
      Width           =   900
   End
   Begin VB.Label lblFugitives 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003C39F3&
      Height          =   195
      Left            =   3980
      TabIndex        =   27
      Top             =   6940
      Width           =   465
   End
   Begin VB.Label lblContra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003C39F3&
      Height          =   195
      Left            =   2700
      TabIndex        =   26
      Top             =   6940
      Width           =   465
   End
   Begin VB.Label lblPassngr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FA54B3&
      Height          =   195
      Left            =   3340
      TabIndex        =   25
      Top             =   6940
      Width           =   465
   End
   Begin VB.Label lblCargo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FA54B3&
      Height          =   195
      Left            =   2130
      TabIndex        =   24
      Top             =   6940
      Width           =   465
   End
   Begin VB.Label lblPartsBuy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   2780
      TabIndex        =   23
      Top             =   3810
      Width           =   195
   End
   Begin VB.Label lblPartsBuy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   2490
      TabIndex        =   22
      Top             =   3790
      Width           =   195
   End
   Begin VB.Label lblFuelBuy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   2
      Left            =   1780
      TabIndex        =   21
      Top             =   3810
      Width           =   195
   End
   Begin VB.Label lblFuelBuy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   1500
      TabIndex        =   20
      Top             =   3780
      Width           =   195
   End
   Begin VB.Label lblUpgrades 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/3"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   3940
      TabIndex        =   19
      Top             =   4080
      Width           =   405
   End
   Begin VB.Label lblCrewSpace 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   2420
      TabIndex        =   18
      Top             =   4080
      Width           =   540
   End
   Begin VB.Image imgPartsBuy 
      Height          =   165
      Left            =   1900
      Top             =   3810
      Width           =   630
   End
   Begin VB.Image imgFuelBuy 
      Height          =   165
      Left            =   880
      Top             =   3810
      Width           =   645
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   165
      Left            =   2670
      TabIndex        =   17
      Top             =   2420
      Width           =   255
   End
   Begin VB.Label lblDisCost 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3330
      TabIndex        =   16
      Top             =   3800
      Width           =   735
   End
   Begin VB.Label lblDisCnt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   165
      Left            =   3750
      TabIndex        =   15
      Top             =   3610
      Width           =   225
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgShore 
      Height          =   555
      Left            =   3150
      Top             =   3040
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   979
      Attr            =   640
      Effects         =   "frmAction.frx":03CA
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgRead 
      Height          =   555
      Left            =   3700
      ToolTipText     =   "Universal Encyclopedia"
      Top             =   4350
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   979
      Attr            =   640
      Effects         =   "frmAction.frx":03E2
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgPhone 
      Height          =   555
      Left            =   3150
      ToolTipText     =   "phone Deal with Higgins"
      Top             =   4350
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      Attr            =   640
      Effects         =   "frmAction.frx":03FA
   End
   Begin VB.Image imgContact 
      Height          =   555
      Left            =   1740
      Top             =   4350
      Width           =   1410
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgDealer 
      Height          =   555
      Left            =   900
      Top             =   4350
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   979
      Attr            =   640
      Effects         =   "frmAction.frx":0412
   End
   Begin VB.Image imgWork 
      Height          =   630
      Left            =   0
      Top             =   6030
      Width           =   900
   End
   Begin VB.Image imgDeal 
      Height          =   1170
      Left            =   0
      Top             =   4350
      Width           =   900
   End
   Begin VB.Image imgOutlaw 
      Height          =   1065
      Left            =   3680
      Top             =   0
      Width           =   825
   End
   Begin VB.Label lblWarrants 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3065
      TabIndex        =   14
      Top             =   700
      Width           =   525
   End
   Begin VB.Image imgWarrants 
      Height          =   1065
      Left            =   2990
      Top             =   0
      Width           =   705
   End
   Begin VB.Image imgBuy 
      Height          =   930
      Left            =   0
      Top             =   3030
      Width           =   900
   End
   Begin VB.Image imgSupply 
      Height          =   390
      Left            =   900
      Top             =   3390
      Width           =   2250
   End
   Begin VB.Image imgBDProof 
      Height          =   300
      Left            =   1920
      Top             =   2730
      Width           =   2595
   End
   Begin VB.Label lblPartsChg 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   2
      Left            =   2920
      TabIndex        =   13
      Top             =   3880
      Width           =   165
   End
   Begin VB.Label lblPartsChg 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   0
      Left            =   2550
      TabIndex        =   12
      Top             =   3880
      Width           =   165
   End
   Begin VB.Label lblBuyParts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   165
      Left            =   2570
      TabIndex        =   11
      Top             =   3780
      Width           =   340
   End
   Begin VB.Label lblBuyFuel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Left            =   1580
      TabIndex        =   10
      Top             =   3780
      Width           =   340
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgShop 
      Height          =   375
      Left            =   900
      Top             =   3040
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   661
      Attr            =   640
      Effects         =   "frmAction.frx":042A
   End
   Begin VB.Image imgFly 
      Height          =   1050
      Left            =   0
      Top             =   1380
      Width           =   900
   End
   Begin VB.Image imgFlyHL 
      Height          =   300
      Left            =   0
      Top             =   2740
      Width           =   1905
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgFlyMole 
      Height          =   690
      Left            =   3180
      Top             =   2020
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1217
      Attr            =   640
      Effects         =   "frmAction.frx":0442
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgFlyBoost 
      Height          =   675
      Left            =   3180
      Top             =   1350
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1191
      Attr            =   640
      Effects         =   "frmAction.frx":045A
   End
   Begin VB.Label lblMRange 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   165
      Left            =   1860
      TabIndex        =   9
      Top             =   2420
      Width           =   255
   End
   Begin VB.Label lblFBFuel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   165
      Left            =   2680
      TabIndex        =   8
      Top             =   1740
      Width           =   255
   End
   Begin VB.Label lblFBRange 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   165
      Left            =   1860
      TabIndex        =   7
      Top             =   1740
      Width           =   255
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgMosey 
      Height          =   375
      Left            =   900
      Top             =   2020
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   661
      Attr            =   640
      Effects         =   "frmAction.frx":0472
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl imgFullBurn 
      Height          =   375
      Left            =   900
      Top             =   1350
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   661
      Attr            =   640
      Effects         =   "frmAction.frx":048A
   End
   Begin VB.Label lblBounties 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   3970
      TabIndex        =   6
      Top             =   1070
      Width           =   615
   End
   Begin VB.Label lblMisbehaves 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   2380
      TabIndex        =   5
      Top             =   1070
      Width           =   640
   End
   Begin VB.Label lblTurn 
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   640
      TabIndex        =   4
      Top             =   1070
      Width           =   640
   End
   Begin VB.Label lblGoals 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   2360
      TabIndex        =   3
      Top             =   690
      Width           =   525
   End
   Begin VB.Label lblParts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0003E631&
      Height          =   195
      Left            =   1530
      TabIndex        =   2
      Top             =   6940
      Width           =   465
   End
   Begin VB.Label lblFuel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   1040
      TabIndex        =   1
      Top             =   6940
      Width           =   465
   End
   Begin VB.Label lblCash 
      BackStyle       =   0  'Transparent
      Caption         =   "$0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003DCBFF&
      Height          =   195
      Left            =   795
      TabIndex        =   0
      Top             =   4065
      Width           =   1065
   End
   Begin VB.Image imgHeader 
      Height          =   1065
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image imgGoals 
      Height          =   1065
      Left            =   2300
      Top             =   0
      Width           =   675
   End
   Begin VB.Label lblDealCargoBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Left            =   950
      TabIndex        =   31
      Top             =   5100
      Width           =   345
   End
   Begin VB.Label lblDealCargoSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Left            =   2415
      TabIndex        =   40
      Top             =   5100
      Width           =   345
   End
   Begin VB.Label lblDealPassngrLoad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Left            =   3900
      TabIndex        =   49
      Top             =   5100
      Width           =   345
   End
   Begin VB.Label lblDealContraBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Left            =   950
      TabIndex        =   34
      Top             =   5270
      Width           =   345
   End
   Begin VB.Label lblDealContraSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Left            =   2415
      TabIndex        =   43
      Top             =   5270
      Width           =   345
   End
   Begin VB.Label lblDealFuelBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Left            =   950
      TabIndex        =   37
      Top             =   5430
      Width           =   345
   End
   Begin VB.Label lblDealPartsSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Left            =   2415
      TabIndex        =   46
      Top             =   5430
      Width           =   345
   End
   Begin VB.Label lblDealFugiLoad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   180
      Left            =   3900
      TabIndex        =   52
      Top             =   5270
      Width           =   345
   End
   Begin VB.Menu mnuWorkPopup 
      Caption         =   "mnuWorkPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuWorkPop 
         Caption         =   "empty"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public moseydone As Boolean, fullburndone As Boolean, buydone As Boolean
Public dealdone As Boolean, workdone As Boolean, disgruntledone As Boolean, rangeBoost As Integer

Private Sub Form_Load()
   Me.Picture = LoadPicture(App.Path & "\gui\Actionbg.jpg")
   imgHeader.Picture = LoadPicture(App.Path & "\gui\Firefly" & player.ID & ".jpg")
   imgGoals.Picture = LoadPicture(App.Path & "\gui\Header2GoalsActive.jpg")
   imgWarrants.Picture = LoadPicture(App.Path & "\gui\Header3WarrantsActive.jpg")
   imgOutlaw.Picture = LoadPicture(App.Path & "\gui\Header4OutlawActive.jpg")
   imgFly.Picture = LoadPicture(App.Path & "\gui\Action1FlyActive.jpg")
   imgFlyHL.Picture = LoadPicture(App.Path & "\gui\FlyHeavyLoadActive.jpg")
   imgBDProof.Picture = LoadPicture(App.Path & "\gui\FlyBreakdownProofActive.jpg")
   imgBuy.Picture = LoadPicture(App.Path & "\gui\Action2BuyActive.jpg")
   'imgSupply.Picture = LoadPicture(App.Path & "\gui\Buy3Beaumonde.jpg")
   'imgContact.Picture = LoadPicture(App.Path & "\gui\Deal2Fanty.jpg")
      
   imgDeal.Picture = LoadPicture(App.Path & "\gui\Action3DealActive.jpg")
   imgWork.Picture = LoadPicture(App.Path & "\gui\Action4WorkActive.jpg")
   
   imgFuelBuy.Picture = LoadPicture(App.Path & "\gui\Buy3FuelActive.jpg")
   imgPartsBuy.Picture = LoadPicture(App.Path & "\gui\Buy3PartsActive.jpg")
   
   imgDealCargo.Picture = LoadPicture(App.Path & "\gui\Deal5CargoActive.jpg")
   imgDealContra.Picture = LoadPicture(App.Path & "\gui\Deal5ContraActive.jpg")
   imgDealFuel.Picture = LoadPicture(App.Path & "\gui\Deal5FuelActive.jpg")
   imgDealParts.Picture = LoadPicture(App.Path & "\gui\Deal5PartsActive.jpg")
   imgLoadPassngr.Picture = LoadPicture(App.Path & "\gui\Deal5PassActive.jpg")
   imgLoadFugi.Picture = LoadPicture(App.Path & "\gui\Deal5FugiActive.jpg")
   imgClearWarrants.Picture = LoadPicture(App.Path & "\gui\Deal5WarrantsActive.jpg")
   imgWorkDrop.Picture = LoadPictureGDIplus(App.Path & "\gui\Work1LocalArrow.gif")
   imgBonus.Picture = LoadPicture(App.Path & "\gui\Action5BonusActive.jpg")
   
   setVisState imgFlyHL, False
   setVisState imgFly, False
   setVisState imgBuy, False
   setVisState imgDeal, False
   setVisState imgWork, False
   setVisState imgBonus, False
   setVisState imgGoals, False
   setVisState imgWarrants, False
   setVisState imgOutlaw, False
   setVisState imgFuelBuy, False
   setVisState imgPartsBuy, False
   setVisState imgBDProof, False
      
   setVisState imgDealCargo, False
   setVisState imgDealContra, False
   setVisState imgDealFuel, False
   setVisState imgDealParts, False
   setVisState imgLoadPassngr, False
   setVisState imgLoadFugi, False
   setVisState imgClearWarrants, False
   setVisState imgWorkDrop, False
      
   'alpha imgs
   disableAllButtons
 
End Sub

Public Sub disableAllButtons(Optional ByVal except As String = vbNullString)

   actionButtonEnable "imgFullBurn", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly1FullBurnInactive.jpg")
   actionButtonEnable "imgMosey", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly2MoseyInactive.jpg")
   If except <> "imgFlyBoost" Then
      actionButtonEnable "imgFlyBoost", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly3BoostInactive.jpg")
   End If
   actionButtonEnable "imgFlyMole", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly4MoleInactive.jpg")
   actionButtonEnable "imgShop", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ShopInactive.jpg")
   actionButtonEnable "imgDealer", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1LocalInactive.jpg")
   actionButtonEnable "imgPhone", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal3PhoneInactive.jpg")
   actionButtonEnable "imgRead", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal4ReadInactive.jpg")
   actionButtonEnable "imgShore", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy2ShoreLeaveInactive.jpg")
   actionButtonEnable "imgClearWarrantsOpt", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal5WarrantsOptInactive.jpg")
   actionButtonEnable "imgWorkLocal", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Work1LocalInactive.jpg")
   actionButtonEnable "imgMakeWork", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Work2MakeInactive.jpg")
   actionButtonEnable "imgMorale", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus1MoraleInactive.jpg")
   actionButtonEnable "imgResolve", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus2AlertsInactive.jpg")
   actionButtonEnable "imgEndTurn", False  '.Picture = LoadPictureGDIplus(App.Path & "\gui\End1EndInactive.jpg")
   actionButtonEnable "imgCancel", False

End Sub

Public Sub disableAllActions()
   setVisState imgFly, False
   setVisState imgBuy, False
   setVisState imgDeal, False
   setVisState imgWork, False
   setVisState imgBonus, False
End Sub

Public Function checkNoOfActions() As Integer

   If moseydone And fullburndone Then
      checkNoOfActions = checkNoOfActions + 1
   End If
   If buydone Then
      checkNoOfActions = checkNoOfActions + 1
   End If
   If dealdone Then
      checkNoOfActions = checkNoOfActions + 1
   End If
   If workdone Then
      checkNoOfActions = checkNoOfActions + 1
   End If
   If disgruntledone Then
      checkNoOfActions = checkNoOfActions + 1
   End If
End Function



Private Sub FDPane1_OnHidden()
   Select Case actionSeq
   Case ASend
   
   Case Else
      playsnd 9
      FDPane1.PaneVisible = True
   End Select
End Sub

Private Sub imgCancel_Click()
   With imgCancel
       If .Tag = "Y" Then
          playsnd 8
          .Picture = LoadPictureGDIplus(App.Path & "\gui\End2CancelClick.jpg")
          
          If actionSeq = ASmosey Then
             If MoseyMovesDone = 0 Then fullburndone = False
             actionSeq = ASidle
          ElseIf actionSeq = ASfullburn Then
            If FullburnMovesDone = 0 Then moseydone = False
            actionSeq = ASidle
          End If
          
          .Visible = False
          .Tag = "N"
          
      End If
   End With
End Sub

Private Sub imgCancel_MouseEnter()
 If imgCancel.Tag = "Y" Then imgCancel.Picture = LoadPictureGDIplus(App.Path & "\gui\End2CancelMouseover.jpg")
End Sub

Private Sub imgCancel_MouseExit()
If imgCancel.Tag = "Y" Then imgCancel.Picture = LoadPictureGDIplus(App.Path & "\gui\End2CancelActive.jpg")
End Sub


Private Sub imgEndTurn_Click()
   With imgEndTurn
       If .Tag = "Y" Then
         disableAllButtons
         .Tag = "1"
          .Picture = LoadPictureGDIplus(App.Path & "\gui\End1EndClick.jpg")
          
         If actionSeq = ASNavEvade Then
            MessBox "You need to EVADE!", "Evade", "Ooops", "", getLeader()
            Exit Sub
         End If
         playsnd 8
         endAction
          
         .Tag = "N"
         '.Picture = LoadPictureGDIplus(App.Path & "\gui\End1EndInactive.jpg")
         
          
      End If
   End With
End Sub

Private Sub imgEndTurn_MouseEnter()
 If imgEndTurn.Tag = "Y" Then imgEndTurn.Picture = LoadPictureGDIplus(App.Path & "\gui\End1EndMouseover.jpg")
End Sub

Private Sub imgEndTurn_MouseExit()
 If imgEndTurn.Tag = "Y" Then imgEndTurn.Picture = LoadPictureGDIplus(App.Path & "\gui\End1EndActive.jpg")
End Sub

Private Sub imgMorale_Click()
   With imgMorale
      If .Tag = "Y" Then
         If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
         If FullburnMovesDone > 0 Then fullburndone = True
         
         disableAllButtons
         .Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus1MoraleClick.jpg")
   
         .Tag = "1"
          
          disgruntledone = True
          actionSeq = ASRemoveDisgr
          playsnd 8
         
      End If
   End With
End Sub

Private Sub imgMorale_MouseEnter()
 If imgMorale.Tag = "Y" Then imgMorale.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus1MoraleMouseover.jpg")
End Sub

Private Sub imgMorale_MouseExit()
 If imgMorale.Tag = "Y" Then imgMorale.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus1MoraleActive.jpg")
End Sub

Private Sub imgResolve_Click()
   With imgResolve
       If .Tag = "Y" Then
         playsnd 8
         disableAllButtons
         .Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus2AlertsClick.jpg")
          
         .Tag = "1"
         
         MessBox "Select an Alert Token to resolve", "Alert Token", "Will Do", "", getLeader()
         actionSeq = ASResolveAlert
         
          
      End If
   End With
End Sub

Private Sub imgResolve_MouseEnter()
 If imgResolve.Tag = "Y" Then imgResolve.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus2AlertsMouseover.jpg")
End Sub

Private Sub imgResolve_MouseExit()
 If imgResolve.Tag = "Y" Then imgResolve.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus2AlertsActive.jpg")
End Sub

Private Sub imgWorkDrop_Click()
  PopupMenu mnuWorkPopup
End Sub

Private Sub imgClearWarrantsOpt_Click()
   If imgClearWarrants.Tag = "N" Then Exit Sub
   With imgClearWarrantsOpt
       If .Tag = "N" Then
          playsnd 13
          .Picture = LoadPictureGDIplus(App.Path & "\gui\Deal5WarrantsOptClick.jpg")
          
          .Tag = "Y"
      Else
           .Picture = LoadPictureGDIplus(App.Path & "\gui\Deal5WarrantsOptInactive.jpg")

          .Tag = "N"
      End If
   End With
End Sub

Private Sub imgClearWarrantsOpt_MouseEnter()
   If imgClearWarrantsOpt.Tag = "N" And imgClearWarrants.Tag = "Y" Then imgClearWarrantsOpt.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal5WarrantsOptActive.jpg")
End Sub

Private Sub imgClearWarrantsOpt_MouseExit()
   If imgClearWarrantsOpt.Tag = "N" Then imgClearWarrantsOpt.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal5WarrantsOptInactive.jpg")
End Sub

Private Sub imgShore_Click()
   With imgShore
       If .Tag = "Y" Then
          playsnd 8
          disableAllButtons
          setVisState imgFuelBuy, False
          setVisState imgPartsBuy, False
          actionButtonEnable "imgShore", True, True
          actionSeq = ASBuyShore
          buydone = True
      End If
   End With

End Sub

Private Sub imgShore_MouseEnter()
   If imgShore.Tag = "Y" Then imgShore.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy2ShoreLeaveMouseover.jpg")
End Sub

Private Sub imgShore_MouseExit()
   If imgShore.Tag = "Y" Then imgShore.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy2ShoreLeaveActive.jpg")
End Sub



Private Sub imgRead_Click()
Dim frmNavPeek As frmNavPeeks
   With imgRead
      If .Tag = "Y" Then
         disableAllButtons
         .Picture = LoadPictureGDIplus(App.Path & "\gui\Deal4ReadClick.jpg")
          
         .Tag = "1"
          
         dealdone = True
         Set frmNavPeek = New frmNavPeeks
         frmNavPeek.NavZone = "E"
         frmNavPeek.Show 1
         PutMsg player.PlayName & " used the Universal Encyclopedia to fiddle with the Misbehave deck", player.ID, Logic!Gamecntr
         actionSeq = ASidle
       
          
      End If
   End With
   
End Sub

Private Sub imgRead_MouseEnter()
   If imgRead.Tag = "Y" Then imgRead.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal4ReadMouseover.jpg")
End Sub

Private Sub imgRead_MouseExit()
   If imgRead.Tag = "Y" Then imgRead.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal4ReadActive.jpg")
End Sub

Private Sub imgPhone_Click()
   With imgPhone
       If .Tag = "Y" Then
          playsnd 13
          .Picture = LoadPictureGDIplus(App.Path & "\gui\Deal3PhoneClick.jpg")
          imgContact.Picture = LoadPicture(App.Path & "\gui\Contact8.jpg")
          .Tag = "1"
          HigginsDealPerk = True
          setVisState imgClearWarrants, False
      End If
   End With
   
End Sub

Private Sub imgPhone_MouseEnter()
   If imgPhone.Tag = "Y" Then imgPhone.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal3PhoneMouseover.jpg")
End Sub

Private Sub imgPhone_MouseExit()
   If imgPhone.Tag = "Y" Then imgPhone.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal3PhoneActive.jpg")
End Sub


Private Sub imgFlyBoost_Click()
   With imgFlyBoost
      Select Case .Tag
      Case "Y"
         playsnd 13
          .Picture = LoadPictureGDIplus(App.Path & "\gui\Fly3BoostClick.jpg")
          .Tag = "1" 'click down
          rangeBoost = 2
          lblFBFuel = Val(lblFBFuel) + 1
          lblFBFuel.ForeColor = &HFF&
          lblFBRange = Trim(Str(Val(lblFBRange) + 2))
          lblFBRange.ForeColor = &H3E631
      Case "1"
         If FullburnMovesDone = 0 Then 'only when not moved yet
            rangeBoost = 0
            imgFlyBoost.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly3BoostActive.jpg")
            lblFBFuel = Val(lblFBFuel) - 1
            If Val(lblFBFuel) < 2 Then lblFBFuel.ForeColor = &H3DCBFF
            lblFBRange = Trim(Str(Val(lblFBRange) - 2))
            lblFBRange.ForeColor = &H3DCBFF
            .Tag = "Y"
         End If
      End Select
   End With
End Sub

Private Sub imgFlyBoost_MouseEnter()
   If imgFlyBoost.Tag = "Y" Then imgFlyBoost.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly3BoostMouseover.jpg")
End Sub

Private Sub imgFlyBoost_MouseExit()
   If imgFlyBoost.Tag = "Y" Then imgFlyBoost.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly3BoostActive.jpg")
End Sub

Private Sub imgFlyMole_Click()
   With imgFlyMole
       If .Tag = "Y" Then
         disableAllButtons
         actionButtonEnable "imgFlyMole", True, True

         PutMsg player.PlayName & "'s Lawman Dobson calls the Alliance Cruiser to his location", player.ID, Logic!Gamecntr
         moseydone = True
         fullburndone = True
         doMoveAlliance player.ID, getPlayerSector(player.ID)
         actionSeq = ASAllianceCall
        
      End If
   End With
End Sub

Private Sub imgFlyMole_MouseEnter()
   If imgFlyMole.Tag = "Y" Then imgFlyMole.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly4MoleMouseover.jpg")
End Sub

Private Sub imgFlyMole_MouseExit()
   If imgFlyMole.Tag = "Y" Then imgFlyMole.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly4MoleActive.jpg")
End Sub

Private Sub imgFullBurn_Click()
   With imgFullBurn
       If .Tag = "Y" Then
         playsnd 8
         disableAllButtons "imgFlyBoost"
         actionButtonEnable "imgFullBurn", True, True
         moseydone = True
         actionSeq = ASfullburn
         actionButtonEnable "imgCancel", True

      End If
   End With
End Sub

Private Sub imgFullBurn_MouseEnter()
   If imgFullBurn.Tag = "Y" Then imgFullBurn.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly1FullBurnMouseover.jpg")
End Sub

Private Sub imgFullBurn_MouseExit()
   If imgFullBurn.Tag = "Y" Then imgFullBurn.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly1FullBurnActive.jpg")
End Sub

Private Sub imgMosey_Click()
   With imgMosey
      If .Tag = "Y" Then
         playsnd 8
         disableAllButtons
         actionButtonEnable "imgMosey", True, True
         fullburndone = True
         actionSeq = ASmosey
         
         actionButtonEnable "imgCancel", True
         
      End If
  End With
End Sub

Private Sub imgMosey_MouseEnter()
   If imgMosey.Tag = "Y" Then imgMosey.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly2MoseyMouseover.jpg")

End Sub

Private Sub imgMosey_MouseExit()
  If imgMosey.Tag = "Y" Then imgMosey.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly2MoseyActive.jpg")

End Sub

Private Sub imgShop_Click()
   With imgShop
       Select Case .Tag
       Case "1" 'init buy
          If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
          If FullburnMovesDone > 0 Then fullburndone = True
          disableAllButtons
          
          .Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ShopClick.jpg")
          .Tag = "2"
          
          If getHaven(getPlayerSector(player.ID)) > 0 Then
            actionSeq = ASBuyHaven
          End If
          
       Case "2" 'draw cards
          .Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1DrawClick.jpg")
          
          .Tag = "3"
          
       Case "2a" 'draw cards
          .Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ConsiderClick.jpg")
      
          .Tag = "3"
       
       Case "3" 'close buy
          .Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1FinishClick.jpg")
          
         .Tag = "4"
         setVisState imgBuy, False
       End Select

      If .Tag <> "N" Then
         playsnd 8
         Select Case actionSeq
         Case ASBuyHaven
            buydone = True
         Case ASBuySelect
            'validate the fuel/parts purchase
            If Main.getBuyCost + doBuyFuelParts(player.ID, Val(lblBuyFuel), Val(lblBuyParts), True) > getMoney(player.ID) Then
               MessBox "Not enough money left to pay for the Fuel/Parts order", "Fuel/Parts order", "Ooops", "", getLeader()
               setMultiStateButton "imgShop", "3"
            ElseIf CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < (Val(lblBuyFuel) + Val(lblBuyParts)) / 2 Then
               MessBox "Not enough Cargo Space for the Fuel/Parts order", "Fuel/Parts order", "Ooops", "", getLeader()
               setMultiStateButton "imgShop", "3"
            Else
               buydone = True
               actionSeq = ASBuyEnd
            End If
            
         Case ASBuySelDiscard
      '            'save selected card as Seq = 6 and draw cards up to 3
            actionSeq = ASBuyDrew 'bounce back and refresh formAction via timer
         Case Else
            actionSeq = ASBuy
         End Select
         
      End If
   
   End With
End Sub

Private Sub imgShop_MouseEnter()
   Select Case imgShop.Tag
      Case "1"
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ShopMouseover.jpg")
      Case "2"
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1DrawMouseover.jpg")
      Case "2a"
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ConsiderMouseover.jpg")
      Case "3"
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1FinishMouseover.jpg")
   End Select
End Sub

Private Sub imgShop_MouseExit()
   Select Case imgShop.Tag
      Case "1"
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ShopActive.jpg")
      Case "2"
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1DrawActive.jpg")
      Case "2a"
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ConsiderActive.jpg")
      Case "3"
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1FinishActive.jpg")
   End Select
End Sub

Private Sub imgDealer_Click()
   With imgDealer
      Select Case .Tag
      Case "1" 'init buy
         If hasHigginsJayneGrudge(imgContact.Tag = "8") Then Exit Sub
         If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
         If FullburnMovesDone > 0 Then fullburndone = True
         
         disableAllButtons
         
         .Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1LocalClick.jpg")
         .Tag = "2"
                   
      Case "2" 'draw cards
         .Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1DrawClick.jpg")
         .Tag = "3"
      
      Case "2a" 'draw cards
         .Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1ConsiderClick.jpg")
         .Tag = "3"
         
      Case "3" 'close buy
         .Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1FinishClick.jpg")
         .Tag = "4"
         setVisState imgDeal, False
      End Select
      If .Tag <> "N" Then
         
         Select Case actionSeq
         Case ASDealSelect
         
            'validate the fuel/parts purchase
            If doBuyFuelParts(player.ID, Val(lblDealFuelBuy), 0, True) > getMoney(player.ID) Then
               MessBox "Not enough money to pay for the Fuel order", "Fuel order", "Ooops", "", getLeader()
               setMultiStateButton "imgDealer", "3"
            ElseIf CargoCapacity(player.ID) - CargoSpaceUsed(player.ID) < (Val(lblDealFuelBuy)) / 2 Then
               MessBox "Not enough Cargo Space for the Fuel order", "Fuel order", "Ooops", "", getLeader()
               setMultiStateButton "imgDealer", "3"
            Else
               dealdone = True
               actionSeq = ASDealEnd
            End If
         
            'save selected (Seq=6 + selected) to players Jobs, unselected back to 5

         Case ASDealSelDiscard
            'save selected card as Seq = 6 and draw cards up to 3
            actionSeq = ASDealDrew 'bounce back and refresh formAction via timer
         Case Else
            actionSeq = ASDeal
         End Select
   
         playsnd 8
      End If
   End With
   
    
End Sub

Private Sub imgDealer_MouseEnter()
   Select Case imgDealer.Tag
      Case "1"
         imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1LocalMouseover.jpg")
      Case "2"
         imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1DrawMouseover.jpg")
      Case "3"
         imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1FinishMouseover.jpg")
   End Select
End Sub

Private Sub imgDealer_MouseExit()
   Select Case imgDealer.Tag
      Case "1"
         imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1LocalActive.jpg")
      Case "2"
         imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1DrawAvailable.jpg")
      Case "3"
         imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1FinishAvailable.jpg")
   End Select
End Sub

Private Sub imgMakeWork_Click()
   With imgMakeWork
      If .Tag = "Y" Then
         If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
         If FullburnMovesDone > 0 Then fullburndone = True
         
         disableAllButtons
          
         .Picture = LoadPictureGDIplus(App.Path & "\gui\Work2MakeClick.jpg")
          
         .Tag = "1"
         
         lblJobName.Tag = ""
          
         workdone = True
         actionSeq = ASWork
         playsnd 8
         
      End If
   End With
End Sub

Private Sub imgMakeWork_MouseEnter()
 If imgMakeWork.Tag = "Y" Then imgMakeWork.Picture = LoadPictureGDIplus(App.Path & "\gui\Work2MakeMouseover.jpg")
End Sub

Private Sub imgMakeWork_MouseExit()
 If imgMakeWork.Tag = "Y" Then imgMakeWork.Picture = LoadPictureGDIplus(App.Path & "\gui\Work2MakeActive.jpg")
End Sub

Private Sub imgWorkLocal_Click()
   With imgWorkLocal
      If .Tag = "Y" Then
         If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
         If FullburnMovesDone > 0 Then fullburndone = True
         
         disableAllButtons
         .Picture = LoadPictureGDIplus(App.Path & "\gui\Work1LocalClick.jpg")
         .Tag = "1"
          
         If hasHigginsJayneWork(Val(lblJobName.Tag)) Then Exit Sub

         workdone = True
         actionSeq = ASWork
         playsnd 8
          
      End If
   End With
End Sub

Private Sub imgWorkLocal_MouseEnter()
 If imgWorkLocal.Tag = "Y" Then imgWorkLocal.Picture = LoadPictureGDIplus(App.Path & "\gui\Work1LocalMouseover.jpg")
End Sub

Private Sub imgWorkLocal_MouseExit()
 If imgWorkLocal.Tag = "Y" Then imgWorkLocal.Picture = LoadPictureGDIplus(App.Path & "\gui\Work1LocalActive.jpg")
End Sub

Private Sub lblDealBuyFuel_Click(Index As Integer)
   If imgDealFuel.Tag = "N" Then Exit Sub
   If Val(lblDealFuelBuy.Caption) = 0 And Index = 0 Then Exit Sub
   lblDealFuelBuy.Caption = Val(lblDealFuelBuy.Caption) + (Index - 1)
End Sub

Private Sub lblDealBuyCargo_Click(Index As Integer)
   If imgDealCargo.Tag = "N" Or imgContact.Tag <> "6" Then Exit Sub  'Harrow only
   If Val(lblDealCargoBuy.Caption) = 0 And Index = 0 Then Exit Sub
   lblDealCargoBuy.Caption = Val(lblDealCargoBuy.Caption) + (Index - 1)
End Sub

Private Sub lblDealBuyContra_Click(Index As Integer)
   If imgDealContra.Tag = "N" Or imgContact.Tag <> "9" Then Exit Sub  'Twins only
   If Val(lblDealContraBuy.Caption) = 0 And Index = 0 Then Exit Sub
   lblDealContraBuy.Caption = Val(lblDealContraBuy.Caption) + (Index - 1)
End Sub

Private Sub lblDealLoadFugi_Click(Index As Integer)
   If imgLoadFugi.Tag = "N" Then Exit Sub
   If Val(lblDealFugiLoad.Caption) = 0 And Index = 0 Then Exit Sub
   lblDealFugiLoad.Caption = Val(lblDealFugiLoad.Caption) + (Index - 1)
End Sub

Private Sub lblDealLoadPassngr_Click(Index As Integer)
   If imgLoadPassngr.Tag = "N" Then Exit Sub
   If Val(lblDealPassngrLoad.Caption) = 0 And Index = 0 Then Exit Sub
   lblDealPassngrLoad.Caption = Val(lblDealPassngrLoad.Caption) + (Index - 1)
End Sub

Private Sub lblDealSellCargo_Click(Index As Integer)
   If imgDealCargo.Tag = "N" Then Exit Sub
   If Val(lblDealCargoSell.Caption) = 0 And Index = 0 Then Exit Sub
   If Val(lblDealCargoSell.Caption) >= Val(lblCargo.Caption) And Index = 2 Then Exit Sub
   lblDealCargoSell.Caption = Val(lblDealCargoSell.Caption) + (Index - 1)
End Sub

Private Sub lblDealSellContra_Click(Index As Integer)
  If imgDealContra.Tag = "N" Then Exit Sub
  If Val(lblDealContraSell.Caption) = 0 And Index = 0 Then Exit Sub
  If Val(lblDealContraSell.Caption) >= Val(lblContra.Caption) And Index = 2 Then Exit Sub
   lblDealContraSell.Caption = Val(lblDealContraSell.Caption) + (Index - 1)
End Sub

Private Sub lblDealSellParts_Click(Index As Integer)
   If imgDealParts.Tag = "N" Then Exit Sub
   If Val(lblDealPartsSell.Caption) = 0 And Index = 0 Then Exit Sub
   If Val(lblDealPartsSell.Caption) >= Val(lblParts.Caption) And Index = 2 Then Exit Sub
   lblDealPartsSell.Caption = Val(lblDealPartsSell.Caption) + (Index - 1)
End Sub

Private Sub lblFuelBuy_Click(Index As Integer)
   If imgFuelBuy.Tag = "N" Then Exit Sub
   If Val(lblBuyFuel.Caption) = 0 And Index = 0 Then Exit Sub
   lblBuyFuel.Caption = Val(lblBuyFuel.Caption) + (Index - 1)
   
End Sub

Private Sub lblPartsBuy_Click(Index As Integer)
   If imgPartsBuy.Tag = "N" Then Exit Sub
   If Val(lblBuyParts.Caption) = 0 And Index = 0 Then Exit Sub
   lblBuyParts.Caption = Val(lblBuyParts.Caption) + (Index - 1)
End Sub

Private Sub lblPartsChg_Click(Index As Integer)

   lblBuyParts = Trim(Str(Val(lblBuyParts) + Index - 1))
   If Val(lblBuyParts) < 0 Then lblBuyParts = "0"
   
End Sub

Private Sub mnuWorkPop_Click(Index As Integer)
   lblJobName = mnuWorkPop(Index).Caption
   lblJobName.Tag = mnuWorkPop(Index).Tag
   lblJobName.ToolTipText = mnuWorkPop(Index).Caption
   timScroll.Enabled = True
End Sub

Public Sub setVisState(ByRef cntl As Control, ByVal enable As Boolean)
   
   If enable Then
      cntl.Visible = True
      cntl.Tag = "Y"
   Else
      cntl.Visible = False
      cntl.Tag = "N"
   End If

End Sub

Public Sub clearWarrant()

      imgClearWarrantsOpt.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal5WarrantsOptInactive.jpg")
      imgClearWarrantsOpt.Tag = "N"
      lblWarrants = ""
      setVisState imgWarrants, False
      setVisState imgClearWarrants, False

End Sub

Public Sub setSupply(ByVal SectorID)
Dim s
   s = varDLookup("SupplyID", "Supply", "SectorID=" & SectorID)
   If s > 0 Then
      imgSupply.Picture = LoadPicture(App.Path & "\gui\Supply" & s & ".jpg")
      imgSupply.Tag = s
      imgSupply.Visible = True
   Else
      imgSupply.Tag = ""
      imgSupply.Visible = False
   End If

End Sub

Public Sub setContact(ByVal SectorID)
Dim s
   If SectorID = -1 Then 'deal with Harken
      s = 5 'harken's contactID
   Else
      s = varDLookup("ContactID", "Contact", "SectorID=" & SectorID)
   End If
   If s > 0 Then
      imgContact.Picture = LoadPicture(App.Path & "\gui\Contact" & s & ".jpg")
      imgContact.Tag = s
      imgContact.Visible = True
   Else
      imgContact.Tag = ""
      imgContact.Visible = False
   End If

End Sub

Public Sub actionButtonEnable(ByVal cntrl As String, ByVal enable As Boolean, Optional ByVal clicked As Boolean = False)

   Select Case cntrl
   Case "imgFullBurn"
      If clicked Then
         imgFullBurn.Tag = "1"
         imgFullBurn.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly1FullBurnClick.jpg")
      ElseIf enable Then
         imgFullBurn.Tag = "Y"
         imgFullBurn.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly1FullBurnActive.jpg")
      Else
         imgFullBurn.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly1FullBurnInactive.jpg")
         imgFullBurn.Tag = "N"
      End If
      
   Case "imgMosey"
      If clicked Then
         imgMosey.Tag = "1"
         imgMosey.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly2MoseyClick.jpg")
      ElseIf enable Then
         imgMosey.Tag = "Y"
         imgMosey.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly2MoseyActive.jpg")
      Else
         imgMosey.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly2MoseyInactive.jpg")
         imgMosey.Tag = "N"
      End If
      
   Case "imgFlyBoost"
      If clicked Then
      ElseIf enable Then
         If imgFlyBoost.Tag <> "1" Then 'already set
            imgFlyBoost.Tag = "Y"
            imgFlyBoost.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly3BoostActive.jpg")
         End If
      Else
         imgFlyBoost.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly3BoostInactive.jpg")
         imgFlyBoost.Tag = "N"
      End If
   
   Case "imgFlyMole"
      If clicked Then
         imgFlyMole.Tag = "1"
         imgFlyMole.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly4MoleClick.jpg")
      ElseIf enable Then
         imgFlyMole.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly4MoleActive.jpg")
         imgFlyMole.Tag = "Y"
      Else
         imgFlyMole.Picture = LoadPictureGDIplus(App.Path & "\gui\Fly4MoleInactive.jpg")
         imgFlyMole.Tag = "N"
      End If
      
   Case "imgShop"
      If clicked Then
      ElseIf enable Then
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ShopActive.jpg")
         imgShop.Tag = "1"
      Else
         imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ShopInactive.jpg")
         imgShop.Tag = "N"
      End If
      
    Case "imgDealer"
      If clicked Then
      ElseIf enable Then
         imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1LocalActive.jpg")
         imgDealer.Tag = "1"
      Else
         imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1LocalInactive.jpg")
         imgDealer.Tag = "N"
      End If
      
   Case "imgPhone"
      If clicked Then
      ElseIf enable Then
         imgPhone.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal3PhoneActive.jpg")
         imgPhone.Tag = "Y"
      Else
         imgPhone.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal3PhoneInactive.jpg")
         imgPhone.Tag = "N"
      End If
      
   Case "imgRead"
      If clicked Then
      ElseIf enable Then
         imgRead.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal4ReadActive.jpg")
         imgRead.Tag = "Y"
      Else
         imgRead.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal4ReadInactive.jpg")
         imgRead.Tag = "N"
      End If
      
   Case "imgShore"
      If clicked Then
         imgShore.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy2ShoreLeaveClick.jpg")
         imgShore.Tag = "1"
      ElseIf enable Then
         imgShore.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy2ShoreLeaveActive.jpg")
         imgShore.Tag = "Y"
         'lblDisCost.Visible = True
         'lblDisCost = "$" & Abs(doShoreLeave(player.ID, True))
      Else
         imgShore.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy2ShoreLeaveInactive.jpg")
         imgShore.Tag = "N"
         'lblDisCost.Visible = False
         'lblDisCost = ""
      End If
      
   Case "imgClearWarrantsOpt"
      If clicked Then
      ElseIf enable Then
         imgClearWarrantsOpt.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal5WarrantsOptActive.jpg")
         imgClearWarrantsOpt.Tag = "Y"
      Else
         imgClearWarrantsOpt.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal5WarrantsOptInactive.jpg")
         imgClearWarrantsOpt.Tag = "N"
      End If
      
   Case "imgWorkLocal"
      If clicked Then
      ElseIf enable Then
         imgWorkLocal.Picture = LoadPictureGDIplus(App.Path & "\gui\Work1LocalActive.jpg")
         imgWorkLocal.Tag = "Y"
      Else
         imgWorkLocal.Picture = LoadPictureGDIplus(App.Path & "\gui\Work1LocalInactive.jpg")
         imgWorkLocal.Tag = "N"
      End If
      
   Case "imgMakeWork"
      If clicked Then
      ElseIf enable Then
         imgMakeWork.Picture = LoadPictureGDIplus(App.Path & "\gui\Work2MakeActive.jpg")
         imgMakeWork.Tag = "Y"
      Else
         imgMakeWork.Picture = LoadPictureGDIplus(App.Path & "\gui\Work2MakeInactive.jpg")
         imgMakeWork.Tag = "N"
      End If
      
   Case "imgMorale"
      If clicked Then
      ElseIf enable Then
         imgMorale.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus1MoraleActive.jpg")
         imgMorale.Tag = "Y"
      Else
         imgMorale.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus1MoraleInactive.jpg")
         imgMorale.Tag = "N"
      End If
      
   Case "imgResolve"
      If clicked Then
      ElseIf enable Then
         imgResolve.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus2AlertsActive.jpg")
         imgResolve.Tag = "Y"
      Else
         imgResolve.Picture = LoadPictureGDIplus(App.Path & "\gui\Bonus2AlertsInactive.jpg")
         imgResolve.Tag = "N"
      End If
   

   Case "imgEndTurn"
      If clicked Then
      ElseIf enable Then
         imgEndTurn.Picture = LoadPictureGDIplus(App.Path & "\gui\End1EndActive.jpg")
         imgEndTurn.Tag = "Y"
      Else
         imgEndTurn.Picture = LoadPictureGDIplus(App.Path & "\gui\End1EndInactive.jpg")
         imgEndTurn.Tag = "N"
      End If

      
   Case "imgCancel"
      If clicked Then
      ElseIf enable Then
         imgCancel.Visible = True
         imgCancel.Picture = LoadPictureGDIplus(App.Path & "\gui\End2CancelActive.jpg")
         imgCancel.Tag = "Y"
      Else
         imgCancel.Visible = False
         imgCancel.Tag = "N"
      End If
      
   End Select
End Sub

Public Sub setMultiStateButton(ByVal cntrl As String, ByVal State As String)
   Select Case cntrl
   Case "imgShop"
      imgShop.Tag = State
       Select Case State
       Case "1" 'init buy
          imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ShopActive.jpg")
          
       Case "2" 'draw cards
          imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1DrawActive.jpg")
       
       Case "2a" 'draw cards
          imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ConsiderActive.jpg")
       
       Case "3" 'close buy
          imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1FinishActive.jpg")
       
       Case "N"
          imgShop.Picture = LoadPictureGDIplus(App.Path & "\gui\Buy1ShopInactive.jpg")
       
       End Select
       
   Case "imgDealer"
      imgDealer.Tag = State
       Select Case State
       Case "1" 'init deal
          imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1LocalActive.jpg")
          
       Case "2" 'draw cards
          imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1DrawAvailable.jpg")
       
       Case "2a" 'draw cards
          imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1ConsiderActive.jpg")
       
       Case "3" 'close deal
          imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1FinishAvailable.jpg")
       
       Case "N"
          imgDealer.Picture = LoadPictureGDIplus(App.Path & "\gui\Deal1LocalInactive.jpg")
       
       End Select
  
   End Select
End Sub

Public Sub endAction()
   disableAllButtons
   disableAllActions
   lblBuyFuel = "0"
   lblBuyParts = "0"
   lblDealCargoBuy = "0"
   lblDealContraBuy = "0"
   lblDealCargoSell = "0"
   lblDealContraSell = "0"
   lblDealPassngrLoad = "0"
   lblDealFugiLoad = "0"
   
   MoseyMovesDone = 0
   FullburnMovesDone = 0
   moseydone = False
   fullburndone = False
   buydone = False
   dealdone = False
   workdone = False
   disgruntledone = False
   HemmorrhagingFuel = False
   SurvShuttlePerk = False
   rangeBoost = 0
   turnExtraRange = 0
   TheBigBlack = 0
   HigginsDealPerk = False
   HarkenDeal = False
   actionSeq = ASend
End Sub

Public Sub buyIsDone()
   If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
   If FullburnMovesDone > 0 Then fullburndone = True
   actionButtonEnable "imgShop", False
   actionButtonEnable "imgShore", False
   lblCash = "$" & getMoney(player.ID)
   buydone = True
End Sub

Private Sub timScroll_Timer()
Static csr As Integer
   If lblJobName.Tag = "" Then
      csr = 0 'reset
      timScroll.Enabled = False
   ElseIf Len(lblJobName.ToolTipText) < 29 Then
      lblJobName.Caption = lblJobName.ToolTipText
      csr = 0 'reset
      timScroll.Enabled = False
   Else
      csr = csr + 1
      If csr > Len(lblJobName.ToolTipText) Then csr = 1
      If Len(Mid(lblJobName.ToolTipText, csr)) < 20 Then
         lblJobName.Caption = Mid(lblJobName.ToolTipText, csr) & " - " & lblJobName.ToolTipText
      Else
         lblJobName.Caption = Mid(lblJobName.ToolTipText, csr)
      End If
   End If
End Sub

Public Sub setPay(ByVal pay)
   lblCash.Tag = pay
   lblCash = "$" & lblCash.Tag
   If Val(lblCash.Tag) < 200 Then
      lblCash.ForeColor = &HFF&
   Else
      lblCash.ForeColor = &H3DCBFF
   End If

End Sub
