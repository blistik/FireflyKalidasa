VERSION 5.00
Object = "{714D09E3-B193-11D3-A192-00A0CC26207F}#1.0#0"; "XDockFloat.dll"
Begin VB.Form frmAction 
   Caption         =   "Actions"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3900
   Icon            =   "frmAction.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmAction.frx":030A
   ScaleHeight     =   3150
   ScaleWidth      =   3900
   Begin VB.TextBox txtFuel 
      Height          =   345
      Left            =   2450
      TabIndex        =   38
      Text            =   "0"
      ToolTipText     =   "Buy Fuel qty $100ea, dbl-clk +1"
      Top             =   1020
      Width           =   315
   End
   Begin VB.CheckBox chkRange2 
      BackColor       =   &H00CBE1ED&
      Height          =   225
      Left            =   3400
      TabIndex        =   36
      ToolTipText     =   "extra 2 range for 1 extra fuel"
      Top             =   580
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Resolve Alerts"
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
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "resolve Alert Tokens"
      Top             =   2380
      Width           =   1545
   End
   Begin VB.TextBox txtFug 
      Height          =   345
      Left            =   3540
      TabIndex        =   29
      Text            =   "0"
      ToolTipText     =   "load Fugitives as part of a Deal"
      Top             =   1530
      Width           =   315
   End
   Begin VB.TextBox txtPass 
      Height          =   345
      Left            =   3180
      TabIndex        =   28
      Text            =   "0"
      ToolTipText     =   "load Passengers as part of a Deal"
      Top             =   1530
      Width           =   315
   End
   Begin VB.CheckBox chkWarrant 
      Height          =   225
      Left            =   3345
      TabIndex        =   31
      ToolTipText     =   "pay Badger $1000 to remove all Warrants"
      Top             =   1650
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Remove Disgruntled"
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
      Left            =   20
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "morale booster"
      Top             =   2760
      Width           =   2115
   End
   Begin VB.TextBox txtCargo 
      Height          =   345
      Left            =   2460
      TabIndex        =   25
      Text            =   "0"
      ToolTipText     =   "sell Cargo to Contact"
      Top             =   1530
      Width           =   315
   End
   Begin VB.TextBox txtContra 
      Height          =   345
      Left            =   2820
      TabIndex        =   27
      Text            =   "0"
      ToolTipText     =   "sell Contraband to Contact"
      Top             =   1530
      Width           =   315
   End
   Begin VB.TextBox txtParts 
      Height          =   345
      Left            =   2810
      TabIndex        =   39
      Text            =   "0"
      ToolTipText     =   "Buy Parts qty $300ea"
      Top             =   1020
      Width           =   315
   End
   Begin VB.CheckBox chkShore 
      Height          =   225
      Left            =   3400
      TabIndex        =   19
      ToolTipText     =   "$100 per Crew"
      Top             =   1140
      Width           =   195
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "End Turn"
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
      Left            =   20
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "end your turn now"
      Top             =   2380
      Width           =   1125
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Work"
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
      Left            =   20
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1950
      Width           =   1125
   End
   Begin VB.ComboBox cbo 
      BackColor       =   &H00CBE1ED&
      Height          =   315
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1950
      Width           =   2660
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Deal"
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
      Index           =   3
      Left            =   20
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1530
      Width           =   1125
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Buy"
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
      Index           =   2
      Left            =   20
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1020
      Width           =   1125
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Full Burn"
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
      Index           =   1
      Left            =   20
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   540
      Width           =   1125
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FF8080&
      Caption         =   "Mosey"
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
      Index           =   0
      Left            =   20
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   70
      Width           =   1125
   End
   Begin VB.Label lblMis 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   1650
      TabIndex        =   40
      ToolTipText     =   "Misbehaves"
      Top             =   2400
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblRange2 
      BackStyle       =   0  'Transparent
      Caption         =   "+2 Rng"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3350
      TabIndex        =   37
      Top             =   360
      Width           =   555
   End
   Begin VB.Label lblPassFugi 
      BackStyle       =   0  'Transparent
      Caption         =   "Pass Fugi"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3150
      TabIndex        =   34
      Top             =   1350
      Width           =   795
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Warrants"
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   3180
      TabIndex        =   33
      Top             =   1485
      Width           =   675
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Remv"
      Height          =   195
      Left            =   3225
      TabIndex        =   32
      Top             =   1350
      Width           =   525
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Carg Cont"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2430
      TabIndex        =   26
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Shore"
      Height          =   195
      Left            =   3300
      TabIndex        =   20
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Leave"
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   3300
      TabIndex        =   23
      Top             =   980
      Width           =   525
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$0"
      Height          =   285
      Left            =   2990
      TabIndex        =   22
      ToolTipText     =   "Money in hand"
      Top             =   50
      Width           =   660
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel Parts"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2460
      TabIndex        =   21
      Top             =   840
      Width           =   825
   End
   Begin VB.Label lblGo 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   1170
      TabIndex        =   18
      ToolTipText     =   "Turns"
      Top             =   2400
      Width           =   435
   End
   Begin VB.Label lblSupply 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1170
      TabIndex        =   16
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label lblMosey 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   50
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Range"
      Height          =   195
      Left            =   1350
      TabIndex        =   13
      Top             =   110
      Width           =   705
   End
   Begin VB.Label lblContact 
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1170
      TabIndex        =   11
      Top             =   1530
      Width           =   1215
   End
   Begin VB.Label lblRange 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   1920
      TabIndex        =   10
      Top             =   560
      Width           =   645
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Range"
      Height          =   195
      Left            =   1950
      TabIndex        =   9
      Top             =   360
      Width           =   705
   End
   Begin VB.Label lblFuelRq 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   1170
      TabIndex        =   6
      Top             =   560
      Width           =   645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel Req"
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   705
   End
   Begin VB.Label lblFuelOn 
      Alignment       =   2  'Center
      BackColor       =   &H00CBE1ED&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   2670
      TabIndex        =   4
      Top             =   555
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel left"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   3
      Top             =   360
      Width           =   555
   End
   Begin XDOCKFLOATLibCtl.FDPane FDPane1 
      Height          =   420
      Left            =   4080
      TabIndex        =   0
      Top             =   2520
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
      Height          =   225
      Left            =   2595
      TabIndex        =   24
      Top             =   110
      Width           =   465
   End
End
Attribute VB_Name = "frmAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public moseydone As Boolean, fullburndone As Boolean, buydone As Boolean
Public dealdone As Boolean, workdone As Boolean, disgruntledone As Boolean

Private Sub cbo_Click()
   cbo.ToolTipText = cbo.Text
End Sub

Private Sub chkShore_Click()
      txtFuel.Enabled = (chkShore.Value = 0 Or getHaven(getPlayerSector(player.ID)) > 0)
      txtParts.Enabled = (chkShore.Value = 0 And getHaven(getPlayerSector(player.ID)) = 0)
      setBackColour txtFuel
      setBackColour txtParts
End Sub

Private Sub FDPane1_OnHidden()
   Select Case actionSeq
   Case ASend
   
   Case Else
      playsnd 9
      FDPane1.PaneVisible = True
   End Select
End Sub

Private Sub Form_Load()
   Dim x As Integer
   x = CStr(countMisbehaves(player.ID))
   If x > 0 Then
      lblMis.Caption = CStr(x)
      lblMis.Visible = True
   End If
End Sub

Private Sub txtCargo_DblClick()
   If txtCargo.Enabled Then txtCargo.Text = CStr(Val(txtCargo.Text) + 1)
End Sub

Private Sub txtContra_DblClick()
   If txtContra.Enabled Then txtContra.Text = CStr(Val(txtContra.Text) + 1)
End Sub

Private Sub txtFuel_DblClick()
   If txtFuel.Enabled Then txtFuel.Text = CStr(Val(txtFuel.Text) + 1)
End Sub

Private Sub txtFug_DblClick()
   If txtFug.Enabled Then txtFug.Text = CStr(Val(txtFug.Text) + 1)
End Sub

Private Sub txtParts_DblClick()
   If txtParts.Enabled Then txtParts.Text = CStr(Val(txtParts.Text) + 1)
End Sub

Private Sub txtPass_DblClick()
   If txtPass.Enabled Then txtPass.Text = CStr(Val(txtPass.Text) + 1)
End Sub

Private Sub chkRange2_Click()
   If chkRange2.Value = 1 Then
      lblRange.Caption = CStr(Val(lblRange.Caption) + 2)
      lblFuelRq.Caption = CStr(Val(lblFuelRq.Caption) + 1)
   ElseIf Val(lblRange.Caption) > 2 Then
      lblRange.Caption = CStr(Val(lblRange.Caption) - 2)
      lblFuelRq.Caption = CStr(Val(lblFuelRq.Caption) - 1)
   End If
End Sub

Private Sub cmd_Click(Index As Integer)
Dim x, SectorID
   
   Select Case Index
      Case 0 'mosey
         If cmd(0).Caption = "Cancel" Then
            cmd(0).Caption = "Mosey"
            If MoseyMovesDone = 0 Then fullburndone = False
            chkRange2.Enabled = True
            actionSeq = ASidle
         Else
            fullburndone = True
            chkRange2.Enabled = False
            actionSeq = ASmosey
            CruiserCutter = 0
            CorvetteSeq = 0
            ignoreToken = 0
         End If
      Case 1  'fullburn
         If cmd(1).Caption = "Cancel" Then
            cmd(1).Caption = "Full Burn"
            If FullburnMovesDone = 0 Then moseydone = False
            chkRange2.Enabled = True
            actionSeq = ASidle
         Else
            moseydone = True
            chkRange2.Enabled = False
            actionSeq = ASfullburn
            'reset any alliance visit
            CruiserCutter = 0
            CorvetteSeq = 0
            ignoreToken = 0
         End If
      Case 2 'buy
         If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
         If FullburnMovesDone > 0 Then fullburndone = True
         SectorID = getPlayerSector(player.ID)
         
         If getHaven(SectorID) > 0 Then
            actionSeq = ASBuyHaven
         
         ElseIf chkShore.Value = 1 Then
             'shore leave only
            actionSeq = ASBuyShore
         End If
         chkShore.Visible = False
         Label4.Visible = False
         Label7.Visible = False
         
         Select Case actionSeq
         Case ASBuyShore, ASBuyHaven
            buydone = True
         Case ASBuySelect
            buydone = True
            actionSeq = ASBuyEnd
            
         Case ASBuySelDiscard
'            'save selected card as Seq = 6 and draw cards up to 3
            actionSeq = ASBuyDrew 'bounce back and refresh frmAction via timer
         Case Else
            'Main.showBuys False, "local"
            actionSeq = ASBuy
         End Select
         playsnd 8
      Case 3 'Deal
         If hasHigginsJayneGrudge(frmAction.lblContact.Tag = "8") Then Exit Sub
         If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
         If FullburnMovesDone > 0 Then fullburndone = True
         Select Case actionSeq
         Case ASDealSelect
            'save selected (Seq=6 + selected) to players Jobs, unselected back to 5
            dealdone = True
            actionSeq = ASDealEnd
         Case ASDealSelDiscard
'            'save selected card as Seq = 6 and draw cards up to 3
            actionSeq = ASDealDrew 'bounce back and refresh frmAction via timer
         Case Else
            actionSeq = ASDeal
         End Select

         playsnd 8
      Case 4 'work
         If hasHigginsJayneWork(GetCombo(cbo)) Then Exit Sub
         If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
         If FullburnMovesDone > 0 Then fullburndone = True
         workdone = True
         actionSeq = ASWork
         playsnd 8
      Case 5 'end turn
         If actionSeq = ASNavEvade Then
            MessBox "You need to EVADE!", "Evade", "Ooops", "", getLeader()
            Exit Sub
         End If
         playsnd 8
         endAction
         
      Case 6 'remove disgruntled special perk
         If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
         If FullburnMovesDone > 0 Then fullburndone = True
         disgruntledone = True
         actionSeq = ASRemoveDisgr
         playsnd 8
      Case 7 'resolve alerts
         MessBox "Select an Alert Token to resolve", "Alert Token", "Will Do", "", getLeader()
         actionSeq = ASResolveAlert
         playsnd 8
      
   End Select
   For x = 0 To 7
      cmd(x).Enabled = False
   Next x
   
   If Index = 0 And actionSeq = ASmosey Then
       cmd(0).Caption = "Cancel"
       cmd(0).Enabled = True
   End If
   If Index = 1 And actionSeq = ASfullburn Then
       cmd(1).Caption = "Cancel"
       cmd(1).Enabled = True
   End If

End Sub

Public Sub endAction()
   txtFuel = "0"
   txtParts = "0"
   txtCargo = "0"
   txtContra = "0"
   txtPass = "0"
   txtFug = "0"
   MoseyMovesDone = 0
   FullburnMovesDone = 0
   moseydone = False
   fullburndone = False
   buydone = False
   dealdone = False
   workdone = False
   disgruntledone = False
   chkRange2.Value = 0
   HemmorrhagingFuel = False
   turnExtraRange = 0
   TheBigBlack = 0
   HigginsDealPerk = False
   actionSeq = ASend
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

Public Sub buyIsDone()
   If MoseyMovesDone > 0 Then moseydone = True 'already moseyed
   If FullburnMovesDone > 0 Then fullburndone = True
   buydone = True
End Sub
