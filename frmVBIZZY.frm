VERSION 5.00
Begin VB.Form frmVBIZZY 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "VBIZZY"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   6315
   Icon            =   "frmVBIZZY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmVBIZZY.frx":030A
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vscrTiles 
      Height          =   4335
      Left            =   4920
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picTiles 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   4
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   5220
      ScaleHeight     =   4980
      ScaleMode       =   0  'User
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      Begin VB.Shape shpSel 
         BorderColor     =   &H0080C0FF&
         BorderWidth     =   4
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox picBoard 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   0
      MouseIcon       =   "frmVBIZZY.frx":0614
      MousePointer    =   99  'Custom
      ScaleHeight     =   4980
      ScaleMode       =   0  'User
      ScaleWidth      =   4860
      TabIndex        =   0
      Top             =   0
      Width           =   4860
      Begin VB.Timer Timer1 
         Left            =   3840
         Top             =   3960
      End
      Begin VB.PictureBox picCur 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   720
         ScaleHeight     =   100
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Shape shpCur 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   4
         Height          =   735
         Left            =   2280
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Options"
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
         Shortcut        =   ^T
      End
      Begin VB.Menu SP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSolve 
         Caption         =   "Solve"
      End
      Begin VB.Menu mnuHighScore 
         Caption         =   "High Scores"
         Visible         =   0   'False
         Begin VB.Menu mnuScore 
            Caption         =   "Score"
            Index           =   0
         End
      End
      Begin VB.Menu SP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
         Begin VB.Menu mnuColors 
            Caption         =   "Colors = 2"
         End
         Begin VB.Menu mnuSize 
            Caption         =   "Size = 8X8"
         End
      End
      Begin VB.Menu SP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInstructions 
         Caption         =   "Instructions"
      End
      Begin VB.Menu SP4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuSP5 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuGTime 
      Caption         =   "TIME:"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuGPoints 
      Caption         =   "Points"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmVBIZZY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const vbDkGrey = &H808080
Const vbLtGrey = &HC0C0C0

Private Type Tile
    Index As Integer
    Visible As Boolean
    Slats(7) As Integer
    Tries() As String 'this is for the automatic solver
End Type

Private Type ScoreInfo
    GColors As Integer
    GSize As Integer
    GPercent As Single
    GName As String
End Type

Dim CurX As Integer
Dim CurY As Integer
Dim TileSet() As Tile
Dim Board() As Tile
Dim CurTile As Tile
Dim BSize As Integer
Dim Colors As Integer
Dim TSel As Integer
Dim SSECS As Double
Dim Paused As Boolean
Dim Scores() As ScoreInfo
Dim TileCount As Integer
Dim Stopped As Boolean
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub AddTry(MyTile As Tile, X As Integer, Y As Integer, Rot As Integer)
On Local Error GoTo eTrap
Dim i As Integer
Dim j As Integer
i = UBound(MyTile.Tries) + 1
For j = 0 To i - 1
    If MyTile.Tries(j) = "" Then
        MyTile.Tries(j) = Format(X, "00") & Format(Y, "00") & Rot
        Exit Sub
    End If
Next j
'if there are no blank ones, then add a new one.
'This is done so that we dont have to itterate through a HUGE
'array of blank 'tries'
ReDim Preserve MyTile.Tries(i)
MyTile.Tries(i) = Format(X, "00") & Format(Y, "00") & Rot
Exit Sub
eTrap:
    i = 0
        Resume Next
End Sub

Sub RemoveTries(X As Integer, Y As Integer)
On Local Error GoTo eTrap
Dim i As Integer
Dim k As Long
Dim j As Long
For i = 0 To UBound(TileSet)
    k = UBound(TileSet(i).Tries)
    For j = 0 To k
        'If Val(Left(TileSet(i).Tries(j), 1)) >= X And Val(Mid(TileSet(i).Tries(j), 3, 2)) >= Y Then
        If Val(Left(TileSet(i).Tries(j), 2)) >= X And Val(Mid(TileSet(i).Tries(j), 3, 2)) >= Y Then
            TileSet(i).Tries(j) = ""
        End If
    Next j
Next i
Exit Sub
eTrap:
    k = -1
    Resume Next
End Sub

Private Function TileTried(MyTile As Tile, X As Integer, Y As Integer, Rot As Integer) As Boolean
On Local Error GoTo eTrap
Dim i As Integer
For i = 0 To UBound(MyTile.Tries)
    If MyTile.Tries(i) = Format(X, "00") & Format(Y, "00") & Rot Then
        TileTried = True
        Exit Function
    End If
Next i
Exit Function
eTrap:
End Function
Private Function CheckPlacement(MyTile As Tile, X As Integer, Y As Integer) As Boolean
Dim Up As Integer
Dim Down As Integer
Dim Left As Integer
Dim Right As Integer
CheckPlacement = True 'assume a good placement
Up = GetUp(Y)
Down = GetDown(Y)
Left = GetLeft(X)
Right = GetRight(X)
If Up <> -1 Then
    If Board(X, Up).Slats(5) <> MyTile.Slats(0) And Board(X, Up).Visible = True Then CheckPlacement = False
    If Board(X, Up).Slats(4) <> MyTile.Slats(1) And Board(X, Up).Visible = True Then CheckPlacement = False
End If
If Down <> -1 Then
    If Board(X, Down).Slats(0) <> MyTile.Slats(5) And Board(X, Down).Visible = True Then CheckPlacement = False
    If Board(X, Down).Slats(1) <> MyTile.Slats(4) And Board(X, Down).Visible = True Then CheckPlacement = False
End If
If Left <> -1 Then
    If Board(Left, Y).Slats(2) <> MyTile.Slats(7) And Board(Left, Y).Visible = True Then CheckPlacement = False
    If Board(Left, Y).Slats(3) <> MyTile.Slats(6) And Board(Left, Y).Visible = True Then CheckPlacement = False
End If
If Right <> -1 Then
    If Board(Right, Y).Slats(7) <> MyTile.Slats(2) And Board(Right, Y).Visible = True Then CheckPlacement = False
    If Board(Right, Y).Slats(6) <> MyTile.Slats(3) And Board(Right, Y).Visible = True Then CheckPlacement = False
End If
End Function




Sub CreateTiles()
Dim i As Integer
Dim j As Integer
Randomize
For i = 0 To UBound(TileSet)
    For j = 0 To 7
        TileSet(i).Slats(j) = CInt(Rnd * Colors)
    Next j
    TileSet(i).Index = i
    TileSet(i).Visible = True
Next i
End Sub

Sub DrawBoard()
Dim i As Integer
Dim j As Integer
picBoard.Cls
For i = 0 To BSize
    For j = 0 To BSize
        DrawTile picBoard, i, j, Board(i, j)
    Next j
Next i
picCur.Cls
DrawTile picCur, 0, 0, CurTile
DoEvents
End Sub

Private Sub DrawTile(MyBox As PictureBox, X As Integer, Y As Integer, ByRef MyTile As Tile)
Dim cX As Single
Dim cY As Single
Dim Color As Long
If MyTile.Visible = False Then Exit Sub
'-----------POSITION 0-------------------
Color = TColor(MyTile.Slats(0))
PrepareBox MyBox, Color
MyBox.Line (X + 0.5, Y + 0.5)-(X, Y), Color
MyBox.Line (X, Y)-(X + 0.5, Y), Color
MyBox.Line (X + 0.5, Y)-(X + 0.5, Y + 0.5), Color
cX = MyBox.ScaleX(X - MyBox.ScaleLeft + 0.375, vbUser, vbPixels)
cY = MyBox.ScaleY(Y - MyBox.ScaleTop + 0.125, vbUser, vbPixels)
FloodFill MyBox.hDC, cX, cY, Color
'-----------POSITION 5-------------------
Color = TColor(MyTile.Slats(5))
PrepareBox MyBox, Color
MyBox.Line (X + 0.5, Y + 0.5)-(X + 0.5, Y + 1), Color
MyBox.Line (X + 0.5, Y + 1)-(X, Y + 1), Color
MyBox.Line (X, Y + 1)-(X + 0.5, Y + 0.5), Color
cX = MyBox.ScaleX(X - MyBox.ScaleLeft + 0.375, vbUser, vbPixels)
cY = MyBox.ScaleY(Y - MyBox.ScaleTop + 0.875, vbUser, vbPixels)
FloodFill MyBox.hDC, cX, cY, Color
'-----------POSITION 1-------------------
Color = TColor(MyTile.Slats(1))
PrepareBox MyBox, Color
MyBox.Line (X + 0.5, Y + 0.5)-(X + 0.5, Y), Color
MyBox.Line (X + 0.5, Y)-(X + 1, Y), Color
MyBox.Line (X + 1, Y)-(X + 0.5, Y + 0.5), Color
cX = MyBox.ScaleX(X - MyBox.ScaleLeft + 0.625, vbUser, vbPixels)
cY = MyBox.ScaleY(Y - MyBox.ScaleTop + 0.125, vbUser, vbPixels)
FloodFill MyBox.hDC, cX, cY, Color
'-----------POSITION 4-------------------
Color = TColor(MyTile.Slats(4))
PrepareBox MyBox, Color
MyBox.Line (X + 0.5, Y + 0.5)-(X + 1, Y + 1), Color
MyBox.Line (X + 1, Y + 1)-(X + 0.5, Y + 1), Color
MyBox.Line (X + 0.5, Y + 1)-(X + 0.5, Y + 0.5), Color
cX = MyBox.ScaleX(X - MyBox.ScaleLeft + 0.625, vbUser, vbPixels)
cY = MyBox.ScaleY(Y - MyBox.ScaleTop + 0.875, vbUser, vbPixels)
FloodFill MyBox.hDC, cX, cY, Color
'-----------POSITION 2-------------------
Color = TColor(MyTile.Slats(2))
PrepareBox MyBox, Color
MyBox.Line (X + 0.5, Y + 0.5)-(X + 1, Y), Color
MyBox.Line (X + 1, Y)-(X + 1, Y + 0.5), Color
MyBox.Line (X + 1, Y + 0.5)-(X + 0.5, Y + 0.5), Color
cX = MyBox.ScaleX(X - MyBox.ScaleLeft + 0.875, vbUser, vbPixels)
cY = MyBox.ScaleY(Y - MyBox.ScaleTop + 0.375, vbUser, vbPixels)
FloodFill MyBox.hDC, cX, cY, Color
'-----------POSITION 7-------------------
Color = TColor(MyTile.Slats(7))
PrepareBox MyBox, Color
MyBox.Line (X + 0.5, Y + 0.5)-(X, Y + 0.5), Color
MyBox.Line (X, Y + 0.5)-(X, Y), Color
MyBox.Line (X, Y)-(X + 0.5, Y + 0.5), Color
cX = MyBox.ScaleX(X - MyBox.ScaleLeft + 0.125, vbUser, vbPixels)
cY = MyBox.ScaleY(Y - MyBox.ScaleTop + 0.375, vbUser, vbPixels)
FloodFill MyBox.hDC, cX, cY, Color
'-----------POSITION 3-------------------
Color = TColor(MyTile.Slats(3))
PrepareBox MyBox, Color
MyBox.Line (X + 0.5, Y + 0.5)-(X + 1, Y + 0.5), Color
MyBox.Line (X + 1, Y + 0.5)-(X + 1, Y + 1), Color
MyBox.Line (X + 1, Y + 1)-(X + 0.5, Y + 0.5), Color
cX = MyBox.ScaleX(X - MyBox.ScaleLeft + 0.875, vbUser, vbPixels)
cY = MyBox.ScaleY(Y - MyBox.ScaleTop + 0.625, vbUser, vbPixels)
FloodFill MyBox.hDC, cX, cY, Color
'-----------POSITION 6-------------------
Color = TColor(MyTile.Slats(6))
PrepareBox MyBox, Color
MyBox.Line (X + 0.5, Y + 0.5)-(X, Y + 1), Color
MyBox.Line (X, Y + 1)-(X, Y + 0.5), Color
MyBox.Line (X, Y + 0.5)-(X + 0.5, Y + 0.5), Color
cX = MyBox.ScaleX(X - MyBox.ScaleLeft + 0.125, vbUser, vbPixels)
cY = MyBox.ScaleY(Y - MyBox.ScaleTop + 0.625, vbUser, vbPixels)
FloodFill MyBox.hDC, cX, cY, Color


Exit Sub

'---------Lines---------------------------
Color = vbLtGrey
MyBox.Line (X, Y)-(X + 1, Y + 1), Color
MyBox.Line (X + 1, Y)-(X, Y + 1), Color
MyBox.Line (X, Y)-(X + 1, Y), Color
MyBox.Line (X, Y + 0.5)-(X + 1, Y + 0.5), Color
MyBox.Line (X, Y + 1)-(X + 1, Y + 1), Color
MyBox.Line (X, Y)-(X, Y + 1), Color
MyBox.Line (X + 0.5, Y)-(X + 0.5, Y + 1), Color
MyBox.Line (X + 1, Y)-(X + 1, Y + 1), Color

End Sub

Sub DrawTileSet()
On Local Error Resume Next
Dim i As Integer
picTiles.ScaleTop = vscrTiles.Value
picTiles.Cls
For i = vscrTiles.Value To vscrTiles.Value + CInt(picTiles.ScaleHeight + 1)
    DrawTile picTiles, 0, i, TileSet(i)
Next i
picTiles.DrawWidth = 4
For i = vscrTiles.Value To vscrTiles.Value + CInt(picTiles.ScaleHeight + 1)
    picTiles.Line (0, i - 1)-(0, i), RGB(120, 255, 120)
    picTiles.Line (1, i - 1)-(1, i), RGB(120, 255, 120)
    picTiles.Line (0, i)-(1, i), RGB(120, 255, 120)
    picTiles.Line (0, i - 1)-(1, i - 1), RGB(120, 255, 120)
Next i
DoEvents
End Sub

Function GameOver() As Boolean
Dim i As Integer
Dim j As Integer
GameOver = True
TileCount = 0
For i = 0 To BSize
    For j = 0 To BSize
        If Board(i, j).Visible = False Then
            GameOver = False
        Else
            TileCount = TileCount + 1
        End If
    Next j
Next i
mnuGPoints.Caption = "Points: " & TileCount * Colors * BSize
End Function

Sub GetBoardTile()
CurTile = Board(CurX, CurY)
Board(CurX, CurY).Visible = False
picCur.Visible = True
shpCur.Visible = False
picCur.Cls
DrawTile picCur, 0, 0, CurTile
TSel = CurTile.Index
DrawBoard
End Sub

Function GetFileAsString(FName As String) As String
On Error GoTo eTrap
Dim FF As Integer
Dim FL As Long
FL = FileLen(FName)
FF = FreeFile
Open FName For Binary As #FF
GetFileAsString = String(FL, " ")
Get #FF, , GetFileAsString
eTrap:
Close #FF
End Function

Sub GetRandomTile()
Dim X As Tile
Dim i As Integer
Randomize
picCur.Cls
picCur.Refresh
For i = 0 To 7
    X.Slats(i) = CInt(Rnd * 1)
Next i
CurTile = X
DrawTile picCur, 0, 0, CurTile
End Sub

Function GetUp(Y As Integer) As Integer
If Y = 0 Then GetUp = -1 Else GetUp = Y - 1
End Function

Function GetDown(Y As Integer) As Integer
If Y = BSize Then GetDown = -1 Else GetDown = Y + 1
End Function
Function GetLeft(X As Integer) As Integer
If X = 0 Then GetLeft = -1 Else GetLeft = X - 1
End Function
Function GetRight(X As Integer) As Integer
If X = BSize Then GetRight = -1 Else GetRight = X + 1
End Function
Sub InitArrays()
ReDim Board(BSize, BSize)
ReDim TileSet(((BSize + 1) ^ 2) - 1)
vscrTiles.Max = UBound(TileSet)
End Sub

Sub InitBoard()
Dim i As Integer
Dim j As Integer
Dim k As Integer
For i = 0 To BSize
    For j = 0 To BSize
        For k = 0 To 7
            Board(i, j).Slats(k) = -1
        Next k
    Next j
Next i
End Sub
Sub PlaceTile()
Board(CurX, CurY) = CurTile
DrawTile picBoard, CurX, CurY, CurTile
CurTile.Visible = False
If GameOver Then
    Timer1.Interval = 0
    MsgBox "You did it!"
End If
End Sub

Sub PrepareBox(MyBox As PictureBox, Color As Long)
MyBox.ForeColor = Color
MyBox.FillColor = Color
MyBox.FillStyle = vbSolid
MyBox.DrawWidth = 1
MyBox.DrawMode = vbCopyPen
End Sub


Sub RotateCurTile()
Dim X As Tile
Dim i As Integer
For i = 2 To 7
    X.Slats(i) = CurTile.Slats(i - 2)
Next i
X.Slats(0) = CurTile.Slats(6)
X.Slats(1) = CurTile.Slats(7)
X.Visible = True
X.Index = CurTile.Index
CurTile = X
picCur.Cls
picCur.Refresh
DrawTile picCur, 0, 0, CurTile
End Sub

Sub Solve()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim r As Integer
Dim ThroughOnce As Boolean
NextTile:
    If Stopped Then Exit Sub
    For i = 0 To BSize
        For j = 0 To BSize
            If Board(i, j).Visible = False Then
                For k = 0 To UBound(TileSet)
                    If TileSet(k).Visible = True Then
                        CurTile = TileSet(k)
                        For r = 0 To 3
                            RotateCurTile
                            If Not TileTried(TileSet(k), i, j, r) Then
                                If CheckPlacement(CurTile, i, j) Then
                                    Board(i, j) = CurTile
                                    CurTile.Visible = False
                                    TileSet(k).Visible = False
                                    DrawBoard
                                    vscrTiles.Value = k
                                    DrawTileSet
                                    AddTry TileSet(k), i, j, r
                                    GoTo NextTile
                                End If
                            End If
                        Next r
                    End If
                Next k
                If ThroughOnce Then GoTo RemoveTile
            End If
        Next j
    Next i
    If Not ThroughOnce Then
        ThroughOnce = True
        GoTo NextTile
    End If
    If GameOver() Then
        Timer1.Interval = 0
        MsgBox "Solved"
        Exit Sub
    End If
    'If we get this far, then we need to remove a tile
RemoveTile:
    For i = BSize To 0 Step -1
        For j = BSize To 0 Step -1
            If Board(i, j).Visible = True Then
                k = Board(i, j).Index
                TileSet(k).Visible = True
                Board(i, j).Visible = False
                DrawBoard
                vscrTiles.Value = k
                DrawTileSet
                If j = BSize Then
                    j = 0
                    i = i + 1
                Else
                    j = j + 1
                End If
                RemoveTries i, j
                GoTo NextTile
            End If
        Next j
    Next i
    MsgBox "No Solution Found!"
    Timer1.Interval = 0
End Sub

Sub StartVBIzzy()
picBoard.Enabled = True
InitArrays
CreateTiles
InitBoard
DrawTileSet
picCur.Visible = False
shpCur.Visible = True
SSECS = Timer()
Timer1.Interval = 500
Call Form_Resize
End Sub

Function TColor(Index As Integer) As Long
Select Case Index
    Case -1: TColor = vbLtGrey
    Case 0: TColor = vbWhite
    Case 1: TColor = vbBlack
    Case 2: TColor = vbRed
    Case 3: TColor = vbGreen
    Case 4: TColor = vbBlue
    Case 5: TColor = vbYellow
    Case 6: TColor = vbMagenta
    Case 7: TColor = vbCyan
End Select
End Function


Private Sub Form_Load()
BSize = 7
Colors = 1
StartVBIzzy
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
If Me.Width < Me.Height + 150 Then Me.Width = Me.Height + 150
picBoard.Width = picBoard.Height
picTiles.Width = Me.ScaleWidth - picBoard.Width - 17
picTiles.ScaleHeight = CInt(picTiles.Height / picTiles.Width)
picTiles.ScaleWidth = 1
vscrTiles.Height = Me.ScaleHeight
vscrTiles.Left = picBoard.Width
shpSel.Width = 1
shpSel.Height = 1
picBoard.ScaleWidth = BSize + 1
picBoard.ScaleHeight = BSize + 1
picCur.Width = 1
picCur.Height = 1
picCur.ScaleWidth = 1
picCur.ScaleHeight = 1
shpCur.Width = 1
shpCur.Height = 1
DrawBoard
DrawTileSet
End Sub


Private Sub mnuColors_Click()
Dim ret As String
ret = InputBox("Input the number of colors you want to use.", "Set Colors", Colors + 1)
If ret = "" Then Exit Sub
If Not IsNumeric(ret) Then Exit Sub
If Val(ret) < 2 Then MsgBox "Too Few Colors": Exit Sub
If Val(ret) > 7 Then MsgBox "Too Many Colors": Exit Sub
Colors = Val(ret) - 1
mnuColors.Caption = "Colors = " & Val(ret)
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub


Private Sub mnuInstructions_Click()
Dim TXT As String
TXT = GetFileAsString(App.Path & "\Instructions.txt")
MsgBox TXT, vbInformation, "Instructions"
End Sub

Private Sub mnuPause_Click()
If Paused Then
    picBoard.Visible = True
    picTiles.Visible = True
    DrawBoard
    DrawTileSet
    Me.Caption = "VBIZZY"
    Timer1.Interval = 500
    Paused = False
Else
    picBoard.Visible = False
    picTiles.Visible = False
    Timer1.Interval = 0
    Me.Caption = "VBIZZY (paused)"
    Paused = True
End If
End Sub

Private Sub mnuSize_Click()
Dim ret As String
ret = InputBox("Input the number of rows and columns you want to use." & Chr$(10) & "X x X", "Set Board Size", BSize + 1)
If ret = "" Then Exit Sub
If Not IsNumeric(ret) Then Exit Sub
If Val(ret) < 2 Then Exit Sub
If Val(ret) < 2 Then MsgBox "Too Small": Exit Sub
If Val(ret) > 150 Then MsgBox "Too Big": Exit Sub
BSize = Val(ret) - 1
mnuColors.Caption = "Size = " & Val(ret) & "X" & Val(ret)
End Sub

Private Sub mnuSolve_Click()
Stopped = False
Solve
End Sub

Private Sub mnuStart_Click()
StartVBIzzy
mnuSettings.Enabled = False
Stopped = False
End Sub

Private Sub mnuStop_Click()
mnuSettings.Enabled = True
Timer1.Interval = 0
Stopped = True
End Sub


Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Stopped Then Exit Sub
If Button = 1 Then
    If picCur.Visible = True Then
        If CheckPlacement(CurTile, CurX, CurY) = False Then Exit Sub
        If Board(CurX, CurY).Visible = True Then
            Dim Swap As Tile
            Swap = Board(CurX, CurY)
            PlaceTile
            CurTile = Swap
            picCur.Visible = True
            shpCur.Visible = False
            picCur.Cls
            DrawTile picCur, 0, 0, CurTile
            TSel = CurTile.Index
        Else
            PlaceTile
            DrawTileSet
            picCur.Visible = False
            shpCur.Visible = True
        End If
    ElseIf Board(CurX, CurY).Visible Then
        GetBoardTile
    End If
End If
If Button = 2 Then
    If CurTile.Visible Then
        RotateCurTile
    ElseIf Board(CurX, CurY).Visible Then
        TileSet(Board(CurX, CurY).Index).Visible = True
        Board(CurX, CurY).Visible = False
        DrawBoard
        DrawTileSet
    End If
End If
End Sub
Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 0 Then X = 0
If X > BSize + 1 Then X = BSize + 1
If Y < 0 Then Y = 0
If Y > BSize + 1 Then Y = BSize + 1
CurX = CInt(X - 0.5)
CurY = CInt(Y - 0.5)
picCur.Left = CurX
picCur.Top = CurY
shpCur.Left = CurX
shpCur.Top = CurY
End Sub

Private Sub picTiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
If Button = 1 Then
    i = CInt(Y - 0.5)
    If TileSet(i).Visible = False Then Exit Sub
    If CurTile.Visible = True Then TileSet(TSel).Visible = True
    TSel = i
    CurTile = TileSet(TSel)
    picCur.Cls
    picCur.Visible = True
    shpCur.Visible = False
    DrawTile picCur, 0, 0, CurTile
    TileSet(TSel).Visible = False
    DrawTileSet
End If
End Sub

Private Sub picTiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpSel.Left = CInt(X - 0.5)
shpSel.Top = CInt(Y - 0.5)
End Sub

Private Sub Timer1_Timer()
mnuGTime.Caption = "TIME: " & Format(DateAdd("S", Timer - SSECS, CDate("12:00:00 AM")), "hh:nn:ss")
End Sub

Private Sub vscrTiles_Change()
DrawTileSet
End Sub


Private Sub vscrTiles_Scroll()
DrawTileSet
End Sub


