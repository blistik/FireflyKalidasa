VERSION 5.00
Begin VB.Form frmJobEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "View/Edit Jobs"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmJobEditor.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   15540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   11
      Left            =   11430
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   80
      Top             =   2730
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2175
      Left            =   10100
      ScaleHeight     =   2115
      ScaleWidth      =   4050
      TabIndex        =   79
      Top             =   3120
      Width           =   4110
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Cancel          =   -1  'True
      Caption         =   "new"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "create a new Task"
      Top             =   5040
      Width           =   550
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "new"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "create a new Task"
      Top             =   4650
      Width           =   550
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "new"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "create a new Task"
      Top             =   4270
      Width           =   550
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "edit this task"
      Top             =   5040
      Width           =   465
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "edit this task"
      Top             =   4650
      Width           =   465
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "edit this task"
      Top             =   4270
      Width           =   465
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Close"
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
      Index           =   3
      Left            =   14410
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "close the form without saving"
      Top             =   5010
      Width           =   1035
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   8
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   5040
      Width           =   7695
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   7
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   4650
      Width           =   7695
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Index           =   6
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   4270
      Width           =   7695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Skill Test on Job Completion"
      Height          =   1515
      Left            =   90
      TabIndex        =   59
      Top             =   2640
      Width           =   9795
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   13
         Left            =   6600
         TabIndex        =   82
         Text            =   "0"
         ToolTipText     =   "do not use with Niska Job + Warrant"
         Top             =   660
         Width           =   885
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   4
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   970
         Width           =   1875
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Fail Kill Crew"
         DataField       =   "do not use for Niska job with Warrant"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   3
         Left            =   8040
         TabIndex        =   30
         ToolTipText     =   "do not use with Niska Job + Warrant"
         Top             =   210
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   12
         Left            =   6600
         TabIndex        =   29
         Text            =   "0"
         ToolTipText     =   $"frmJobEditor.frx":1A191
         Top             =   330
         Width           =   885
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   4560
         TabIndex        =   28
         Text            =   "0"
         ToolTipText     =   " 4-attempt botched. 5- Move Cruiser to sector, EVADE.  6-Move Corvette to sector,EVADE.  - 100x pay $"
         Top             =   660
         Width           =   885
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   4560
         TabIndex        =   27
         Text            =   "0"
         ToolTipText     =   "min skill value to achieve middle result"
         Top             =   330
         Width           =   885
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Win Opt Keyword"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   1740
         TabIndex        =   63
         ToolTipText     =   "tick to optionally Win with a Keyword (1/2 Pay for Explosives)"
         Top             =   1020
         Width           =   1575
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   9
         Left            =   2670
         TabIndex        =   26
         Text            =   "0"
         ToolTipText     =   $"frmJobEditor.frx":1A24C
         Top             =   660
         Width           =   885
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   2670
         TabIndex        =   25
         Text            =   "0"
         ToolTipText     =   "minimum skill value to win"
         Top             =   330
         Width           =   885
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   3
         Left            =   420
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Fail Kill Crew"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   23
         Left            =   5580
         TabIndex        =   83
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "---==== FAIL ====---"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5760
         TabIndex        =   76
         Top             =   120
         Width           =   1545
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "---== Intermediate ==---"
         Height          =   195
         Left            =   3750
         TabIndex        =   75
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "---==== WIN ====---"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2070
         TabIndex        =   74
         Top             =   120
         Width           =   1545
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Lose Rep"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   17
         Left            =   5590
         TabIndex        =   67
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Fail Result"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   16
         Left            =   5590
         TabIndex        =   66
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Inter Result"
         Height          =   225
         Index           =   15
         Left            =   3660
         TabIndex        =   65
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll Interm"
         Height          =   225
         Index           =   14
         Left            =   3660
         TabIndex        =   64
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Win Result"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   13
         Left            =   1770
         TabIndex        =   62
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll to Win"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   11
         Left            =   1770
         TabIndex        =   61
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         Height          =   225
         Index           =   10
         Left            =   60
         TabIndex        =   60
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Delete"
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
      Index           =   2
      Left            =   14410
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "delete this Job"
      Top             =   3450
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Save"
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
      Index           =   0
      Left            =   14410
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "save Job"
      Top             =   4500
      Width           =   1035
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "New Job"
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
      Index           =   1
      Left            =   14410
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "add a new Job"
      Top             =   3990
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CBE1ED&
      Caption         =   "Contact's Job"
      Height          =   2565
      Left            =   90
      TabIndex        =   45
      Top             =   90
      Width           =   15375
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   10
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   77
         ToolTipText     =   "Bonus Pay x skill points"
         Top             =   2070
         Width           =   915
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   9
         Left            =   11490
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2160
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Illegal"
         Height          =   285
         Index           =   8
         Left            =   5160
         TabIndex        =   9
         Top             =   1470
         Width           =   945
      End
      Begin VB.ListBox lstReqProf 
         BackColor       =   &H00CBE1ED&
         Height          =   2085
         Left            =   13230
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   390
         Width           =   1995
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   5
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1500
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Good Deeds"
         Height          =   165
         Index           =   9
         Left            =   11790
         TabIndex        =   21
         ToolTipText     =   "no need to pay Moral Crew"
         Top             =   1740
         Width           =   1245
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Heavy Load Fuel"
         Height          =   270
         Index           =   7
         Left            =   11520
         TabIndex        =   20
         ToolTipText     =   "+1 Fuel for a Full Burn"
         Top             =   1400
         Width           =   1515
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Rmv Disgruntled"
         Height          =   285
         Index           =   6
         Left            =   11520
         TabIndex        =   19
         ToolTipText     =   "remove Disgruntled from Moral Crew"
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Keyword or Solid"
         Height          =   285
         Index           =   5
         Left            =   11520
         TabIndex        =   18
         ToolTipText     =   "Needs Keyword OR being Solid with Contact"
         Top             =   765
         Width           =   1515
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Keyword or Prof."
         Height          =   285
         Index           =   4
         Left            =   11520
         TabIndex        =   17
         ToolTipText     =   "needs Keyword Or Profession"
         Top             =   450
         Width           =   1515
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   2
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   400
         Width           =   3675
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H0051FF51&
         Height          =   285
         Index           =   7
         Left            =   8460
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "Negotiate Skill required"
         Top             =   1365
         Width           =   795
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H00F9F900&
         Height          =   285
         Index           =   6
         Left            =   8460
         TabIndex        =   13
         Text            =   "0"
         Top             =   885
         Width           =   795
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H005C59DB&
         Height          =   285
         Index           =   5
         Left            =   8460
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "Fight  Skill required"
         Top             =   405
         Width           =   795
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   5340
         TabIndex        =   8
         Text            =   "0"
         Top             =   1110
         Width           =   885
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Keyword Bonus"
         Height          =   285
         Index           =   1
         Left            =   11520
         TabIndex        =   16
         ToolTipText     =   "get the Bonus if holds the Keyword"
         Top             =   150
         Width           =   1515
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   1
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1500
         Width           =   1425
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1500
         Width           =   1425
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CBE1ED&
         Caption         =   "Immoral"
         Height          =   285
         Index           =   0
         Left            =   5130
         TabIndex        =   10
         Top             =   1740
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   2100
         Width           =   6075
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   5340
         TabIndex        =   6
         Text            =   "0"
         Top             =   390
         Width           =   885
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   5340
         TabIndex        =   7
         Text            =   "0"
         Top             =   750
         Width           =   885
      End
      Begin VB.ListBox lstKeyword 
         BackColor       =   &H00CBE1ED&
         Height          =   2085
         Left            =   9360
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   390
         Width           =   2025
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   1
         Top             =   975
         Width           =   3675
      End
      Begin VB.ListBox lstProf 
         BackColor       =   &H00CBE1ED&
         Height          =   2085
         Left            =   6345
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   390
         Width           =   1995
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus per Skill Point"
         Height          =   405
         Left            =   8490
         TabIndex        =   78
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Need to be Solid with"
         Height          =   285
         Index           =   25
         Left            =   11520
         TabIndex        =   73
         Top             =   1950
         Width           =   1545
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Requires: Profession"
         Height          =   225
         Left            =   13230
         TabIndex        =   72
         Top             =   180
         Width           =   2205
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         Height          =   285
         Index           =   18
         Left            =   3120
         TabIndex        =   68
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Card ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   58
         Top             =   200
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Negot"
         Height          =   285
         Index           =   8
         Left            =   8520
         TabIndex        =   57
         Top             =   1170
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Tech"
         Height          =   285
         Index           =   7
         Left            =   8520
         TabIndex        =   56
         ToolTipText     =   "Tech Skill required"
         Top             =   690
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Fight"
         Height          =   285
         Index           =   6
         Left            =   8460
         TabIndex        =   55
         Top             =   210
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bonus Parts"
         Height          =   285
         Index           =   5
         Left            =   3960
         TabIndex        =   54
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "JobType 2"
         Height          =   285
         Index           =   3
         Left            =   1650
         TabIndex        =   53
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "JobType 1"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   52
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Payment"
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   51
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00CBE1ED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bonus Pay"
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   50
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword - Needs OR . . ."
         Height          =   225
         Left            =   9360
         TabIndex        =   49
         Top             =   180
         Width           =   2115
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Name"
         Height          =   285
         Index           =   12
         Left            =   210
         TabIndex        =   48
         Top             =   765
         Width           =   1545
      End
      Begin VB.Label lbl 
         BackColor       =   &H00CBE1ED&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Special Details"
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   47
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Profession for Bonus"
         Height          =   225
         Left            =   6350
         TabIndex        =   46
         Top             =   180
         Width           =   2205
      End
   End
   Begin VB.Label lbl 
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      Caption         =   "Wanted Fugitive"
      Height          =   285
      Index           =   22
      Left            =   10110
      TabIndex        =   81
      Top             =   2790
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bonus"
      Height          =   285
      Index           =   21
      Left            =   120
      TabIndex        =   71
      Top             =   5055
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Task 2"
      Height          =   285
      Index           =   20
      Left            =   120
      TabIndex        =   70
      Top             =   4665
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Task 1"
      Height          =   285
      Index           =   19
      Left            =   120
      TabIndex        =   69
      Top             =   4290
      Width           =   735
   End
End
Attribute VB_Name = "frmJobEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lockEdits As Boolean, JobCardID As Integer

Private Sub cbo_Click(Index As Integer)
   Select Case Index
   Case 2
      RefreshJob GetCombo(cbo(Index))
   Case 5
      If GetCombo(cbo(5)) = 10 Then
         Picture1.top = 3120
         cbo(11).Visible = True
      Else
         cbo(11).Visible = False
         cbo(11).ListIndex = -1
         Picture1.top = 2820
         If Picture1.Tag <> "s" Then Picture1.Picture = LoadPicture(App.Path & "\Pictures\Salvage.jpg")
         Picture1.Tag = "s"
      End If
   Case 11
      If GetCombo(cbo(11)) > 0 Then
         Picture1.Picture = LoadPicture(App.Path & "\Pictures\" & varDLookup("picture", "Crew", "CrewID=" & GetCombo(cbo(11))))
         Picture1.Tag = ""
      End If
   
   End Select
End Sub
Private Sub cbo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      cbo(Index).ListIndex = -1
   End If
      
End Sub
Private Sub cmd_Click(Index As Integer)
Dim x
   Select Case Index
   Case 0 'save
      If GetCombo(cbo(6)) = -1 Then
         MsgBox "Must have Task 1 set", vbExclamation
         Exit Sub
      End If
      
      saveJob
   Case 1
      ClearForm
      x = newJob
      cbo(2).AddItem x
      cbo(2).ItemData(cbo(2).NewIndex) = x
      cbo(2).ListIndex = cbo(2).NewIndex
      
   Case 2 'delete
      If GetCombo(cbo(2)) < 264 Then
         MsgBox "This is a Protected Job", vbExclamation
      Else
         If MsgBox("Are you sure you want to Delete Job " & GetCombo(cbo(2)) & "?", vbYesNo + vbQuestion, "Delete job") = vbYes Then
            DB.Execute "DELETE FROM ContactDeck WHERE CardID=" & GetCombo(cbo(2))
            MsgBox "Job No. " & GetCombo(cbo(2)) & " deleted", vbInformation
            LoadCombo cbo(2), "contactdeck"
         End If
      End If
   Case 3 'cancel
      Me.Hide
      
   Case 4, 5, 6
      editTask Index + 2
   Case 7, 8, 9
      editTask Index - 1, True

   End Select
End Sub

Private Sub Form_Load()
Dim x
   LoadCombo lstProf, "profession"
   LoadCombo lstReqProf, "profession"
   LoadCombo cbo(0), "jobtype"
   LoadCombo cbo(1), "jobtype"
   LoadCombo cbo(2), "contactdeck"
   LoadCombo cbo(3), "skill"
   LoadCombo cbo(4), "contact", " WHERE ContactID > 0 and ContactID < 10"
   LoadCombo cbo(5), "contact"
   LoadCombo cbo(6), "task", "WHERE SectorID > 0"
   LoadCombo cbo(7), "task", "WHERE SectorID > 0"
   LoadCombo cbo(8), "task", "WHERE SectorID > 0"
   LoadCombo cbo(9), "contact", " WHERE ContactID > 0 and ContactID < 10"
   LoadCombo cbo(10), "skill"
   LoadCombo cbo(11), "crew", " WHERE Leader = 0 Order by CrewName"
   
   lstKeyword.AddItem "EXPLOSIVES"
   lstKeyword.AddItem "FAKEID"
   lstKeyword.AddItem "FANCYDUDS"
   lstKeyword.AddItem "FIREARM"
   lstKeyword.AddItem "FIREARMX2"
   lstKeyword.AddItem "HACKINGRIG"
   lstKeyword.AddItem "KNIFE"
   lstKeyword.AddItem "SNIPERRIFLE"
   lstKeyword.AddItem "TRANSPORT"
   For x = cmd.LBound To cmd.UBound
    cmd(x).Visible = Not lockEdits
   Next x
   cmd(3).Visible = True
   If JobCardID > 0 Then
      SetCombo cbo(2), "", JobCardID
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.Visible Then
      Cancel = True
   End If
End Sub

Private Sub RefreshJob(ByVal CardID)
Dim rst As New ADODB.Recordset, MP
Dim SQL
   MP = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   SQL = "SELECT * FROM ContactDeck WHERE CardID =" & CardID

   rst.Open SQL, DB, adOpenForwardOnly, adLockReadOnly
   If Not rst.EOF Then
      SetCombo cbo(5), "", rst!ContactID
      SetCombo cbo(0), "", rst!JobTypeID
      SetCombo cbo(1), "", rst!JobType2D
      SetCombo cbo(3), "", rst!skill
      SetCombo cbo(4), "", rst!FailLoseRep
      SetCombo cbo(9), "", rst!Solid
      SetCombo cbo(6), "", rst!Job1ID
      SetCombo cbo(7), "", rst!Job2ID
      SetCombo cbo(8), "", rst!Job3ID
      SetCombo cbo(10), "", rst!BonusPerSkill
      SetCombo cbo(11), "", rst!FugitiveID
      
      SetProflist lstProf, rst!ProfessionID
      SetProflist lstReqProf, rst!RequireProfession
      SetChklist lstKeyword, Nz(rst!KeyWords)
      txt(0) = Nz(rst!pay)
      txt(1) = Nz(rst!bonus)
      txt(2) = Nz(rst!JobName)
      txt(3) = Nz(rst!JobOrder)
      txt(4) = Nz(rst!BonusPart)
      txt(5) = Nz(rst!fight)
      txt(6) = Nz(rst!tech)
      txt(7) = Nz(rst!Negotiate)
      txt(8) = Nz(rst!win)
      txt(9) = Nz(rst!WinResult)
      txt(10) = Nz(rst!Intermediate)
      txt(11) = Nz(rst!IntermediateResult)
      txt(12) = Nz(rst!FailResult)
      txt(13) = Nz(rst!FailKillCrew)
      chk(0).Value = rst!Immoral
      chk(1).Value = rst!KeywordBonus
      chk(2).Value = rst!WinOptKeyword
      'chk(3).Value = rst!FailKillCrew
      chk(4).Value = rst!KeywordOrSkill
      chk(5).Value = rst!KeywordOrSolid
      chk(6).Value = rst!RemoveDisgruntled
      chk(7).Value = rst!ExtraFuel
      chk(8).Value = rst!illegal
      chk(9).Value = rst!GoodDeeds
      

   End If
   rst.Close
   Set rst = Nothing
   Screen.MousePointer = MP

End Sub

Private Sub SetProflist(cmbo As Control, ByVal itemVal)
Dim x
On Error GoTo err_handler

   With cmbo
      For x = 0 To .ListCount - 1
         .selected(x) = False
         If Val(itemVal) = 12 And (.ItemData(x) = 1 Or .ItemData(x) = 2) Then
            .selected(x) = True
         ElseIf .ItemData(x) = Val(itemVal) Then
           .selected(x) = True
         End If
      Next x

   End With
   
normal_exit:
   Exit Sub

err_handler:
   MsgBox "SetChecklist " & itemVal & vbCrLf & Err.Description
   Resume normal_exit
   
End Sub


Private Sub ClearForm()
Dim x
   For x = cbo.LBound To cbo.UBound
      cbo(x).ListIndex = -1
   Next x
   clearList lstProf
   clearList lstReqProf
   clearList lstKeyword
   For x = 0 To 1
      txt(x) = "0"
   Next x
   For x = 2 To 3
      txt(x) = ""
   Next x
   For x = 4 To txt.UBound
      txt(x) = "0"
   Next x
   
   For x = chk.LBound To chk.UBound
      chk(x).Value = 0
   Next x
      
End Sub

Private Function newJob() As Integer
'Dim rst As New ADODB.Recordset
Dim SQL, CardID As Integer
On Error GoTo err_handler
   CardID = varDLookup("max(CardID)", "ContactDeck", "") + 1
   SQL = "INSET INTO ContactDeck (CardID, JobName, ContactID, Job1ID) VALUES (" & CardID & ", 'New Job at " & Now() & "', 0,1)"
   DB.Execute SQL
   newJob = varDLookup("max(CardID)", "ContactDeck", "")
'   rst.Open SQL, DB, adOpenDynamic, adLockPessimistic
'   rst.AddNew
'   rst!JobName = "New Job at " & Now()
'   rst!ContactID = 0
'   rst!Job1ID = 1
'   rst.Update
'   newJob = rst!CardID
'   rst.Close
'   Set rst = Nothing
normal_exit:
   Exit Function
   
err_handler:
   MsgBox "Error: " & vbCrLf & Err.Description
   Resume normal_exit
End Function

Private Function saveJob() As Boolean
Dim CardID As Integer, SQL
On Error GoTo err_handler
   CardID = cbo(2).ItemData(cbo(2).ListIndex)
   If CardID = 0 Then Exit Function
   SQL = "UPDATE ContactDeck Set "
   SQL = SQL & "ContactID=" & GetCombo(cbo(5))
   SQL = SQL & ", JobTypeID=" & GetCombo(cbo(0))
   SQL = SQL & ", JobType2D=" & GetCombo(cbo(1))
   SQL = SQL & ", JobName='" & SQLFilter(txt(2)) & "'"
   If txt(3) = "" Then
      SQL = SQL & ", JobOrder=Null"
   Else
      SQL = SQL & ", JobOrder='" & SQLFilter(txt(3)) & "'"
   End If
   SQL = SQL & ", Job1ID=" & GetCombo(cbo(6))
   If GetCombo(cbo(7)) = -1 Then
      SQL = SQL & ", Job2ID=0"
   Else
      SQL = SQL & ", Job2ID=" & GetCombo(cbo(7))
   End If
   If GetCombo(cbo(8)) = -1 Then
      SQL = SQL & ", Job3ID=0"
   Else
      SQL = SQL & ", Job3ID=" & GetCombo(cbo(8))
   End If
   SQL = SQL & ", Illegal=" & chk(8).Value
   SQL = SQL & ", Immoral=" & chk(0).Value
   SQL = SQL & ", Pay=" & CStr(Val(txt(0)))
   SQL = SQL & ", Fight=" & CStr(Val(txt(5)))
   SQL = SQL & ", Tech=" & CStr(Val(txt(6)))
   SQL = SQL & ", Negotiate=" & CStr(Val(txt(7)))
   SQL = SQL & ", Keywords=" & IIf(getList(lstKeyword) = "", "NULL", "'" & getList(lstKeyword) & "'")
   If GetCombo(cbo(9)) = -1 Then
      SQL = SQL & ", Solid=0"
   Else
      SQL = SQL & ", Solid=" & GetCombo(cbo(9))
   End If
   If GetCombo(cbo(3)) = -1 Then
      SQL = SQL & ", Skill=0"
   Else
      SQL = SQL & ", Skill=" & GetCombo(cbo(3))
   End If
   If GetCombo(cbo(10)) = -1 Then
      SQL = SQL & ", BonusPerSkill=0"
   Else
      SQL = SQL & ", BonusPerSkill=" & GetCombo(cbo(10))
   End If
   If GetCombo(cbo(11)) = -1 Then
      SQL = SQL & ", FugitiveID=0"
   Else
      SQL = SQL & ", FugitiveID=" & GetCombo(cbo(11))
   End If
   SQL = SQL & ", Win=" & CStr(Val(txt(8)))
   SQL = SQL & ", WinResult=" & CStr(Val(txt(9)))
   SQL = SQL & ", WinOptKeyword=" & chk(2).Value
   SQL = SQL & ", FailResult=" & CStr(Val(txt(12)))
   SQL = SQL & ", FailKillCrew=" & CStr(Val(txt(13))) 'chk(3).Value
   If GetCombo(cbo(4)) = -1 Then
      SQL = SQL & ", FailLoseRep=0"
   Else
      SQL = SQL & ", FailLoseRep=" & GetCombo(cbo(4))
   End If
   SQL = SQL & ", Intermediate=" & CStr(Val(txt(10)))
   SQL = SQL & ", IntermediateResult=" & CStr(Val(txt(11)))
   SQL = SQL & ", ExtraFuel=" & chk(7).Value
   SQL = SQL & ", ProfessionID=" & getProfItem(lstProf)
   SQL = SQL & ", Bonus=" & CStr(Val(txt(1)))
   SQL = SQL & ", BonusPart=" & CStr(Val(txt(4)))
   SQL = SQL & ", RemoveDisgruntled=" & chk(6).Value
   SQL = SQL & ", KeywordOrSkill=" & chk(4).Value
   SQL = SQL & ", KeywordBonus=" & chk(1).Value
   SQL = SQL & ", KeywordOrSolid=" & chk(5).Value
   SQL = SQL & ", RequireProfession=" & getProfItem(lstReqProf)
   SQL = SQL & ", GoodDeeds=" & chk(9).Value
   SQL = SQL & " WHERE CardID =" & CardID
   DB.Execute SQL
   
normal_exit:
   Exit Function
   
err_handler:
   MsgBox "Error: " & vbCrLf & Err.Description
   Resume normal_exit
End Function

Private Sub clearList(cbo As Control)
Dim x
   With cbo
      For x = 0 To .ListCount - 1
         .selected(x) = False
      Next x
   End With
   
End Sub

Private Function getProfItem(cbo As Control) As Integer
Dim x
   With cbo
      If .selected(0) And .selected(1) Then 'mech & pilot
         getProfItem = 12
         Exit Function
      End If
      For x = 0 To .ListCount - 1
         If .selected(x) Then
            getProfItem = CStr(.ItemData(x))
            Exit For
         End If
      Next x
   End With
   
End Function



Private Function getList(cbo As Control) As String
Dim x
   With cbo
      For x = 0 To .ListCount - 1
         If .selected(x) Then
            getList = getList & IIf(getList = "", "", " ") & CStr(.List(x))
         End If
      Next x
   End With
   
End Function


Private Function getSelected(cbo As Control) As Integer
Dim x
   With cbo
      For x = 0 To .ListCount - 1
         If .selected(x) Then
            getSelected = getSelected + 1
         End If
      Next x
   End With
   
End Function

Private Function SetChklist(cbo As Control, ByVal solids As String) As Integer
Dim x, y, a() As String

   With cbo
      a = Split(solids, " ")
      For x = 0 To .ListCount - 1
         .selected(x) = False
         For y = LBound(a) To UBound(a)
            If .List(x) = a(y) Then
               .selected(x) = True
               SetChklist = SetChklist + 1
               'Exit For
            End If
         Next y
      Next x
   End With
   
End Function

Private Sub editTask(ByVal Index As Integer, Optional ByVal newtask As Boolean = False)
Dim frmJobtask As New frmJobTasks
   With frmJobtask
      If Not newtask And GetCombo(cbo(Index)) = -1 Then Exit Sub
      .JobID = IIf(newtask, 0, GetCombo(cbo(Index)))
      .Show 1, Me
      If .JobID > 0 Then
         LoadCombo cbo(Index), "task", "WHERE SectorID > 0"
         SetCombo cbo(Index), "", .JobID
      End If
      
      
   End With
End Sub

