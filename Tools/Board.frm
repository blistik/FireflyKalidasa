VERSION 5.00
Object = "{49801673-2EC8-456E-98B2-037B9B02A1C5}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form Board 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "the 'Verse"
   ClientHeight    =   15360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   22740
   Icon            =   "Board.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15360
   ScaleWidth      =   22740
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   15060
      Left            =   30
      ScaleHeight     =   15000
      ScaleWidth      =   22575
      TabIndex        =   0
      Top             =   30
      Width           =   22635
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000001&
         Caption         =   "\/"
         Height          =   195
         Left            =   20670
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "prev.sector"
         Top             =   14840
         Width           =   435
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000001&
         Caption         =   "+"
         Height          =   285
         Left            =   20130
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "add sector"
         Top             =   14700
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "edit Hotspots"
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   21060
         TabIndex        =   6
         ToolTipText     =   "off edits ship posn"
         Top             =   60
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000001&
         Caption         =   "/\"
         Height          =   195
         Left            =   20670
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "next sector"
         Top             =   14640
         Width           =   435
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000001&
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   20250
         TabIndex        =   3
         Top             =   14340
         Width           =   2265
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000001&
         Caption         =   "Sav"
         Height          =   285
         Left            =   22080
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "save hotspot or ship positions"
         Top             =   14700
         Width           =   435
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000001&
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   21240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "sector selector"
         Top             =   14670
         Width           =   795
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   465
         Index           =   6
         Left            =   21570
         Top             =   13320
         Width           =   850
         _ExtentX        =   1508
         _ExtentY        =   820
         Effects         =   "Board.frx":0442
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   615
         Index           =   5
         Left            =   21690
         Top             =   12570
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         Effects         =   "Board.frx":045A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   615
         Index           =   4
         Left            =   21750
         Top             =   11850
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         Effects         =   "Board.frx":0472
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   615
         Index           =   3
         Left            =   21690
         Top             =   11160
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         Effects         =   "Board.frx":048A
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   615
         Index           =   2
         Left            =   21720
         Top             =   10440
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         Effects         =   "Board.frx":04A2
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl Imag 
         Height          =   615
         Index           =   1
         Left            =   21660
         Top             =   9720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         Effects         =   "Board.frx":04BA
      End
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   20550
      TabIndex        =   5
      Top             =   15150
      Width           =   375
   End
   Begin VB.Image HotSpot 
      Height          =   330
      Index           =   0
      Left            =   22080
      MouseIcon       =   "Board.frx":04D2
      MousePointer    =   4  'Icon
      Stretch         =   -1  'True
      Top             =   15120
      Width           =   540
   End
End
Attribute VB_Name = "Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
'For use with USER32 Function SendMessage
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private DB As ADODB.Connection, Logic As New ADODB.Recordset
Private DragControl As Integer, myAppPath As String

Private Sub Check1_Click()
Dim rst As New ADODB.Recordset, slot
   With Me

   
      rst.Open "SELECT * FROM Board WHERE SectorID > 0 ORDER BY SectorID", DB, adOpenDynamic, adLockOptimistic
      While Not rst.EOF
         .HotSpot(rst!sectorID).Top = rst!STop
         .HotSpot(rst!sectorID).Left = rst!SLeft
         .HotSpot(rst!sectorID).Height = rst!SHeight
         .HotSpot(rst!sectorID).Width = rst!SWidth
         .lbl(rst!sectorID).Top = rst!STop
         .lbl(rst!sectorID).Left = rst!SLeft
         .lbl(rst!sectorID).Height = rst!SHeight
         .lbl(rst!sectorID).Width = rst!SWidth
         
         If Check1.Value = 1 Then
            .HotSpot(rst!sectorID).Picture = LoadPicture(myAppPath & "\Pictures\" & "HotSpot.jpg")
            .HotSpot(rst!sectorID).ZOrder
            .lbl(rst!sectorID).ZOrder
         Else
            .HotSpot(rst!sectorID).Picture = LoadPicture()
            .HotSpot(rst!sectorID).ZOrder
         End If
       
         rst.MoveNext
      Wend
      rst.Close
       
      If Check1.Value = 0 Then
         For slot = 1 To 6
            .Imag(slot).ZOrder
         Next slot
      End If
      Command3.Visible = (Check1.Value = 1)
   
   End With
   
   
End Sub

Private Sub Combo1_Click()
   If Combo1.ListIndex > -1 Then
      Text1 = ""
      LoadCoords Combo1.List(Combo1.ListIndex)
   End If
End Sub

Private Sub Command1_Click()
   'shiftCoords
   'SaveHotPoints
   SaveCoords Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Command2_Click()
   If Combo1.ListIndex < Combo1.ListCount - 1 Then
      Combo1.ListIndex = Combo1.ListIndex + 1
   Else
      Combo1.ListIndex = 1
   End If
End Sub
Private Sub Command4_Click()
   
   If Combo1.ListIndex = -1 Then
      Combo1.ListIndex = Combo1.ListCount - 1
   ElseIf Combo1.ListIndex = 1 Then
      Combo1.ListIndex = Combo1.ListCount - 1
   Else
      Combo1.ListIndex = Combo1.ListIndex - 1
   End If
End Sub
Private Sub Command3_Click()
Dim sectorID, SQL
   With Me
      sectorID = Combo1.List(Combo1.ListCount - 1) + 1
      Load .HotSpot(sectorID)
      Set .HotSpot(sectorID).Container = .Picture1
      .HotSpot(sectorID).Top = 90
      .HotSpot(sectorID).Left = Abs(.Width - 8840)
      .HotSpot(sectorID).Height = 885
      .HotSpot(sectorID).Width = 1440
      .HotSpot(sectorID).ZOrder
      .HotSpot(sectorID).Visible = True
      .HotSpot(sectorID).Picture = LoadPicture(myAppPath & "\Pictures\" & "HotSpot.jpg")
      .HotSpot(sectorID).BorderStyle = 0
      .HotSpot(sectorID).ToolTipText = CStr(sectorID)
      Load .lbl(sectorID)
      Set .lbl(sectorID).Container = .Picture1
      .lbl(sectorID).Top = 90
      .lbl(sectorID).Left = Abs(.Width - 8840)
      .lbl(sectorID).Height = 885
      .lbl(sectorID).Width = 1440
      .lbl(sectorID).Visible = True
      .lbl(sectorID).ZOrder
      .lbl(sectorID).ToolTipText = sectorID
      .lbl(sectorID) = sectorID
      SQL = "INSERT INTO Board (SectorID, Slot1, Slot2, Slot3, Slot4, Slot5, Zones, STop, SLeft, SHeight, SWidth) VALUES ("
      SQL = SQL & sectorID & ", '" & HotSpot(sectorID).Left & "," & HotSpot(sectorID).Top & "', '" & HotSpot(sectorID).Left & "," & HotSpot(sectorID).Top & "', "
      SQL = SQL & "'" & HotSpot(sectorID).Left & "," & HotSpot(sectorID).Top & "', '" & HotSpot(sectorID).Left & "," & HotSpot(sectorID).Top & "', "
      SQL = SQL & "'" & HotSpot(sectorID).Left & "," & HotSpot(sectorID).Top & "', 'R', "
      SQL = SQL & HotSpot(sectorID).Top & ", " & HotSpot(sectorID).Left & ", " & HotSpot(sectorID).Height & ", " & HotSpot(sectorID).Width & ")"
      DB.Execute SQL
      Combo1.AddItem sectorID
      Combo1.ListIndex = Combo1.ListCount - 1
   End With
End Sub



Private Sub Form_Load()
Dim x
   If Not Logon Then End
   
   Imag(1).Picture = LoadPictureGDIplus(myAppPath & "\Pictures\FireflyOrange.bmp")
   Imag(2).Picture = LoadPictureGDIplus(myAppPath & "\Pictures\FireflyBlue.bmp")
   Imag(3).Picture = LoadPictureGDIplus(myAppPath & "\Pictures\FireflyYellow.bmp")
   Imag(4).Picture = LoadPictureGDIplus(myAppPath & "\Pictures\FireflyGreen.bmp")
   Imag(5).Picture = LoadPictureGDIplus(myAppPath & "\Pictures\Crusier.bmp")
   Imag(5).AutoSize = lvicMultiAngle
   Imag(6).Picture = LoadPictureGDIplus(myAppPath & "\Pictures\Cutter.bmp") 'corvette.bmp")  '
   For x = 1 To 6
      Imag(x).TransparentColor = 0
      Imag(x).TransparentColorMode = lvicUseTransparentColor
   Next x
   
   Logic.Open "GameSeq", DB, adOpenDynamic, adLockOptimistic
   initBoard
   InitCbo

End Sub

Private Sub Form_Resize()
   Check1.Left = Me.Width - 1680
   Text1.Left = Me.Width - 2490
   Command1.Left = Me.Width - 660
   Combo1.Left = Me.Width - 1500
   Command2.Left = Me.Width - 2070
   Command4.Left = Me.Width - 2070
   Command3.Left = Me.Width - 2600
   
   
End Sub

Private Sub HotSpot_Click(Index As Integer)
      'MsgBox "You clicked on " & Index
      'Picture1.Visible = Not Picture1.Visible
      Text1 = Text1 & IIf(Text1 = "", "", ",") & CStr(Index)
End Sub


Private Sub lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 And Check1.Value = 1 Then
      DragControl = Shift
      Screen.MousePointer = vbCrosshair
   End If
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim z, w

   z = lbl(Index).Left
   w = lbl(Index).Top

   Select Case DragControl
   Case vbShiftMask
      
      lbl(Index).Left = z + x - 40
      lbl(Index).Top = w + y - 40
      
   Case vbCtrlMask

      lbl(Index).Width = Abs(x)
      lbl(Index).Height = Abs(y)
  
   End Select
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   DragControl = 0
   Screen.MousePointer = vbDefault
   HotSpot(Index).Left = lbl(Index).Left
   HotSpot(Index).Top = lbl(Index).Top
   HotSpot(Index).Width = lbl(Index).Width
   HotSpot(Index).Height = lbl(Index).Height
End Sub

Private Sub Imag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 And Check1.Value = 0 Then
      DragControl = Shift
      Screen.MousePointer = vbCrosshair
   End If
End Sub

Private Sub Imag_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim z, w

   z = Imag(Index).Left
   w = Imag(Index).Top

   Select Case DragControl
   Case vbShiftMask
      
      Imag(Index).Left = z + x - 40
      Imag(Index).Top = w + y - 40
      
   End Select
End Sub

Private Sub Imag_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   DragControl = 0
   Screen.MousePointer = vbDefault
End Sub

Public Function Logon() As Boolean
Dim datab
  
  On Error Resume Next
  
  If Command$ = "" Then
     myAppPath = App.Path
  Else
     myAppPath = Command$
  End If
  
  datab = myAppPath & "\FireflyKalidasa.mdb"
  Set DB = New ADODB.Connection
  DB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & datab & ";Persist Security Info=False"
  If Err Then
     Logon = False
  Else
     Logon = True
  End If
  
  
End Function
Public Sub InitCbo()
Dim rst As New ADODB.Recordset
Dim coords, slot
Dim c()
   Combo1.Clear
     
   rst.Open "SELECT SectorID FROM Board ORDER BY SectorID ", DB, adOpenDynamic, adLockOptimistic
   While Not rst.EOF
      Combo1.AddItem rst!sectorID
      rst.MoveNext
   Wend
   rst.Close


End Sub

Public Sub LoadCoords(ByVal sectorID)
Dim rst As New ADODB.Recordset
Dim coords, slot, more
Dim c() As String
   
   rst.Open "SELECT * FROM Board WHERE SectorID = " & sectorID, DB, adOpenDynamic, adLockOptimistic
   If Not rst.EOF Then
      For slot = 1 To 6
         Board.Imag(slot).Visible = Not IsNull(rst!Zones)
      Next slot
      If IsNull(rst!Zones) Then

         Exit Sub
      End If
      Text1 = rst!AdjacentRows & ""
      Board.Imag(5).Visible = (rst!Zones = "A")
      Board.Imag(6).Visible = Not (rst!Zones = "A")
      Board.Imag(5).Tag = rst!Zones
      more = 0
      For slot = 1 To 5
         coords = rst.Fields("Slot" & slot).Value
         If IsNull(coords) Then
            Exit For
         Else
            c = Split(coords, ",")
            If (rst!Zones = "B" Or rst!Zones = "R") And slot = 5 Then more = 1
            Board.Imag(slot + more).Left = c(0)
            Board.Imag(slot + more).Top = c(1)
         End If
      Next slot
   End If


End Sub
Public Sub SaveCoords(ByVal sectorID)
Dim more
   If Check1.Value = 0 Then 'saving ships
      more = 0
      If Board.Imag(5).Tag = "B" Or Board.Imag(5).Tag = "R" Then more = 1
         
      DB.Execute "Update Board Set Slot1 = '" & Imag(1).Left & "," & Imag(1).Top & "',  Slot2 = '" & Imag(2).Left & "," & Imag(2).Top & _
      "', Slot3 = '" & Imag(3).Left & "," & Imag(3).Top & "', Slot4 = '" & Imag(4).Left & "," & Imag(4).Top & "', Slot5 = '" & _
      Imag(5 + more).Left & "," & Imag(5 + more).Top & "', AdjacentRows = '" & Text1 & "' WHERE SectorID = " & sectorID
   
   Else 'saving hotspots
      DB.Execute "Update Board Set STop = " & HotSpot(sectorID).Top & ", SLeft = " & HotSpot(sectorID).Left & ", SHeight = " & HotSpot(sectorID).Height & ", SWidth = " & HotSpot(sectorID).Width & " WHERE SectorID = " & sectorID
   End If
End Sub

'Private Sub SaveHotPoints()
'Dim x
'   For x = 1 To 100
'      DB.Execute "UPDATE Board SET SLeft = " & HotSpot(x).Left & ", STop=" & HotSpot(x).Top & ", SWidth=" & HotSpot(x).Width & ", SHeight=" & HotSpot(x).Height & " WHERE SectorID =" & CStr(x)
'   Next x
'
'End Sub

Private Sub RefreshBoard()
Dim rst As New ADODB.Recordset
   With Me

    
      rst.Open "SELECT * FROM Board WHERE SectorID > 0 ORDER BY SectorID", DB, adOpenDynamic, adLockOptimistic
      While Not rst.EOF
         .HotSpot(rst!sectorID).Top = rst!STop
         .HotSpot(rst!sectorID).Left = rst!SLeft
         .HotSpot(rst!sectorID).Height = rst!SHeight
         .HotSpot(rst!sectorID).Width = rst!SWidth
         '.HotSpot(rst!sectorID).ZOrder
         '.HotSpot(rst!sectorID).Visible = True
         '.HotSpot(rst!sectorID).Picture = LoadPicture(App.Path & "\Pictures\" & "HotSpot.jpg")
         '.HotSpot(rst!sectorID).BorderStyle = 0
         .lbl(rst!sectorID).Top = rst!STop
         .lbl(rst!sectorID).Left = rst!SLeft
         .lbl(rst!sectorID).Height = rst!SHeight
         .lbl(rst!sectorID).Width = rst!SWidth
         '.lbl(rst!sectorID).Visible = True
         '.lbl(rst!sectorID).ZOrder
         .lbl(rst!sectorID) = rst!sectorID
         rst.MoveNext
      Wend
      rst.Close
      
      
   End With
End Sub

Private Sub initBoard()
Dim rst As New ADODB.Recordset
   With Me

      .Picture1.Picture = LoadPicture(myAppPath & "\Pictures\" & Logic!BoardPicture)
      .Height = Logic!BHeight + 100
      .Width = Logic!BWidth + 50
   
      rst.Open "SELECT * FROM Board WHERE SectorID > 0 ORDER BY SectorID", DB, adOpenDynamic, adLockOptimistic
      While Not rst.EOF
         Load .HotSpot(rst!sectorID)
         Set .HotSpot(rst!sectorID).Container = .Picture1
         .HotSpot(rst!sectorID).Top = rst!STop
         .HotSpot(rst!sectorID).Left = rst!SLeft
         .HotSpot(rst!sectorID).Height = rst!SHeight
         .HotSpot(rst!sectorID).Width = rst!SWidth
         .HotSpot(rst!sectorID).ZOrder
         .HotSpot(rst!sectorID).Visible = True
         .HotSpot(rst!sectorID).Picture = LoadPicture(myAppPath & "\Pictures\" & "HotSpot.jpg")
         .HotSpot(rst!sectorID).BorderStyle = 0
         .HotSpot(rst!sectorID).ToolTipText = CStr(rst!sectorID)
         Load .lbl(rst!sectorID)
         Set .lbl(rst!sectorID).Container = .Picture1
         .lbl(rst!sectorID).Top = rst!STop
         .lbl(rst!sectorID).Left = rst!SLeft
         .lbl(rst!sectorID).Height = rst!SHeight
         .lbl(rst!sectorID).Width = rst!SWidth
         .lbl(rst!sectorID).Visible = True
         .lbl(rst!sectorID).ZOrder
         .lbl(rst!sectorID).ToolTipText = CStr(rst!sectorID)
         .lbl(rst!sectorID) = rst!sectorID
         rst.MoveNext
      Wend
      rst.Close
            
      .Caption = "the 'Verse - " & varDLookup("StoryTitle", "Story", "StoryID=" & Logic.Fields("StoryID"))
      .Show
      
      
   End With
End Sub


Public Sub shiftCoords()
Dim rst As New ADODB.Recordset
Dim coords, slot, more
Dim c() As String
   
   rst.Open "SELECT * FROM Board WHERE SectorID > 0 ORDER BY SectorID ", DB, adOpenDynamic, adLockOptimistic
   While Not rst.EOF
      For slot = 1 To 5
         coords = rst.Fields("Slot" & slot).Value
         If IsNull(coords) Then
            Exit For
         Else
            c = Split(coords, ",")
            rst.Fields("Slot" & slot).Value = CStr(Val(c(0)) + 7480) & "," & CStr(c(1))
            
         End If
      Next slot
      rst.Fields("SLeft").Value = rst.Fields("SLeft").Value + 7480
      rst.Update
      
      rst.MoveNext
   Wend

End Sub

Public Function varDLookup(ByVal vstrField As String, ByVal vstrDomain As String, Optional ByVal vstrCriteria As String = vbNullString) As Variant

 Dim rstLookup As ADODB.Recordset

  'The SQL to locate the status code from the schema
  Dim strSQL As String

  On Error GoTo errhandler

  'Assume no record will be found
  varDLookup = Null

  'Prefix the where clause to the criteria if supplied
  If Len(vstrCriteria) > 0 Then vstrCriteria = " WHERE " & vstrCriteria

  'Generate the SQL statement to return the status code
  'for the currently displayed outage
  strSQL = "SELECT " & vstrField & " FROM " & vstrDomain & vstrCriteria

  'Generate a new instance of the recordset object
  Set rstLookup = New ADODB.Recordset

  'Return all the data to the client machine
  rstLookup.CursorLocation = adUseClient
  
  'Open the selected record
  rstLookup.Open strSQL, DB

  'Provided a record was returned, set the return
  'value to the value of the required field
  If Not rstLookup.EOF Then _
    varDLookup = rstLookup.Fields(vstrField)
  
  'Close the recordset and clean up memory
  rstLookup.Close
  Set rstLookup = Nothing
  
normalexit:
  Exit Function
  
errhandler:
  
  'Display the error description for the moment.
  MsgBox "varDLookup:" & strSQL & vbCrLf & Err.Description
  Resume normalexit
  
End Function

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

