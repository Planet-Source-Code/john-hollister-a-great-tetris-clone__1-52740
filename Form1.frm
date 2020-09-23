VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "Tehtaryss"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox sscore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2880
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox slevel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2880
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox slines 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2880
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox stime 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2880
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox numbers 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2760
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.PictureBox nextup 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2760
      ScaleHeight     =   945
      ScaleWidth      =   1335
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   1365
   End
   Begin VB.PictureBox gamescreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4880
      Left            =   120
      ScaleHeight     =   242.25
      ScaleMode       =   2  'Point
      ScaleWidth      =   122.25
      TabIndex        =   0
      Top             =   120
      Width           =   2475
      Begin VB.Label howtoplay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "...How to play...                Right key = right Left key = left   Up key = rotate  Spacebar = pause"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   1575
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label aboutme 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "...About this...  Tetris is boring, but I always wanted to program it. So I did, yesterday.    -John Hollister     March, 2004"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Image igameover 
         Height          =   270
         Left            =   480
         Picture         =   "Form1.frx":0BEE
         Top             =   1920
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Image ipause 
         Height          =   270
         Left            =   720
         Picture         =   "Form1.frx":1F50
         Top             =   2280
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.PictureBox blocks 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2760
      Picture         =   "Form1.frx":2C3A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   165
      Left            =   3285
      Picture         =   "Form1.frx":447C
      Top             =   1240
      Width           =   330
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   165
      Left            =   3285
      Picture         =   "Form1.frx":46DA
      Top             =   0
      Width           =   330
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   165
      Left            =   3240
      Picture         =   "Form1.frx":4938
      Top             =   2670
      Width           =   405
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   165
      Left            =   3240
      Picture         =   "Form1.frx":4C26
      Top             =   1920
      Width           =   405
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000C&
      BorderWidth     =   7
      X1              =   2640
      X2              =   4320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   165
      Left            =   3240
      Picture         =   "Form1.frx":4F14
      Top             =   3405
      Width           =   405
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      BorderWidth     =   7
      X1              =   2640
      X2              =   4320
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000003&
      BorderWidth     =   6
      X1              =   4250
      X2              =   4250
      Y1              =   0
      Y2              =   5040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      BorderWidth     =   7
      X1              =   2640
      X2              =   4320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000003&
      BorderWidth     =   6
      X1              =   4320
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      BorderWidth     =   6
      X1              =   4320
      X2              =   0
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderWidth     =   15
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   6
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   5040
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000C&
      BorderWidth     =   7
      X1              =   2640
      X2              =   4320
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu newGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu border 
         Caption         =   "-"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu howto 
         Caption         =   "How to play..."
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written by John Hollister 2004


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private t As Long
Private board(0 To 13, -3 To 20) As Boolean
Private bcolor(0 To 12, -1 To 20) As Integer
Private nextpiece As curr
Private nextpname
Private speed As Integer
Public piece As curr
Public level As Integer
Public lines As Integer
Private sp(0 To 15) As Integer
Private gameover As Boolean
Private paused As Boolean
Private time As Integer
Private score As Long



Function gameLoop()
igameover.Visible = False
'these are just speed controls for levels
sp(0) = 33
sp(1) = 28
sp(2) = 24
sp(3) = 20
sp(4) = 16
sp(5) = 12
sp(6) = 8
sp(7) = 6
sp(8) = 4
sp(9) = 3
sp(10) = 3
sp(11) = 3
sp(12) = 3
sp(13) = 3
sp(14) = 3
sp(15) = 3
score = 0
level = 0


'set up the first pieces
Set nextpiece = New curr
Dim mat() As Integer
mat = nextpiece.returnPiece
For a = 0 To 3
    BitBlt nextup.hDC, (mat(a, 0) - 2) * 16, (mat(a, 1) + 1) * 16, 16, 16, blocks.hDC, 112, 0, vbSrcAnd
    BitBlt nextup.hDC, (mat(a, 0) - 2) * 16, (mat(a, 1) + 1) * 16, 16, 16, blocks.hDC, nextpiece.ptype * 16, 0, vbSrcPaint
Next a
Set piece = New curr

'enter game loop
Do
t = GetTickCount + 33
        
    'spin wait here while paused
    While paused
        DoEvents
    Wend
    
    '*************************
    'drop piece
    '*************************
    'this mod controls piece speed.  the lower the mod, the faster
    speed = (speed + 1) Mod sp(Int(lines / 10))
    If speed Mod sp(Int(lines / 10)) = 0 Then Call dropPiece
    
    
    '*************************
    'update stats
    '*************************
    'draws time, score, level...
    Call updateStats

    
    '*************************
    'drawgraphics
    '*************************
    Call drawScreen
        
    gamescreen.Refresh
        
    While t > GetTickCount
        DoEvents
    Wend
    time = time + 1
    If gameover = False Then gamescreen.Cls

Loop Until gameover
igameover.Visible = True


End Function

Function drawScreen()

Dim mat() As Integer
mat = piece.returnPiece

'this blits the board
For a = 0 To 10
    For b = 0 To 20
        If board(a, b) Then
            BitBlt gamescreen.hDC, a * 16, b * 16, 16, 16, blocks.hDC, 112, 0, vbSrcAnd
            BitBlt gamescreen.hDC, a * 16, b * 16, 16, 16, blocks.hDC, bcolor(a, b) * 16, 0, vbSrcPaint
        End If
    Next b
Next a

'this blits the currently falling piece
For a = 0 To 3
    BitBlt gamescreen.hDC, mat(a, 0) * 16, mat(a, 1) * 16, 16, 16, blocks.hDC, 112, 0, vbSrcAnd
    BitBlt gamescreen.hDC, mat(a, 0) * 16, mat(a, 1) * 16, 16, 16, blocks.hDC, piece.ptype * 16, 0, vbSrcPaint
Next a
End Function


Function dropPiece()
Dim mat() As Integer
mat = piece.returnPiece
'can the piece drop anymore?
If mat(2, 1) >= 19 Or canDown = False Then
    For a = 0 To 3
        board(mat(a, 0), mat(a, 1)) = True
        bcolor(mat(a, 0), mat(a, 1)) = piece.ptype
    Next a
    'check to see if rows were completed
    Call dropRows
    'set up next current piece
    Set piece = nextpiece
    Set nextpiece = New curr
    mat = nextpiece.returnPiece
    
    'get the next piece
    nextup.Cls
    For a = 0 To 3
        BitBlt nextup.hDC, (mat(a, 0) - 2) * 16, (mat(a, 1) + 1) * 16, 16, 16, blocks.hDC, 112, 0, vbSrcAnd
        BitBlt nextup.hDC, (mat(a, 0) - 2) * 16, (mat(a, 1) + 1) * 16, 16, 16, blocks.hDC, nextpiece.ptype * 16, 0, vbSrcPaint
    Next a
    Call lose

    speed = 1
    direction = 0
Else
'the piece can move down
    piece.moveDown
End If
End Function



Private Sub about_Click()
    paused = True
    aboutme.Visible = True
    ipause.Visible = True
    howtoplay.Visible = False
End Sub

Private Sub Form_Load()
    Me.Show
    speed = 0
    Call gameLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub gamescreen_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mat() As Integer
mat = piece.returnPiece

'pause the game
If KeyCode = vbKeySpace Then
    howtoplay.Visible = False
    aboutme.Visible = False
    paused = (paused + 1) Mod 2
    If paused = True Then
        If gameover = False Then
            ipause.Visible = True
        End If
    Else
        ipause.Visible = False
    End If
End If

'controls
If paused = False Then
    If KeyCode = vbKeyLeft And mat(0, 0) > 0 Then
        If canLeft Then
            piece.moveLeft
        End If
    ElseIf KeyCode = vbKeyRight And mat(1, 0) < 9 Then
        If canRight Then
            piece.moveRight
        End If
    ElseIf KeyCode = vbKeyDown And mat(2, 1) < 19 Then
        If canDown Then
            speed = 1
            piece.moveDown
        End If
    ElseIf KeyCode = vbKeyUp Then
        Call piece.rotate
        If canRotate = False Then
            For a = 0 To 2
                Call piece.rotate
            Next a
        End If
    End If
End If
End Sub

'is the piece able to be moved left?
Public Function canLeft() As Boolean
Dim mat() As Integer
mat = piece.returnPiece

Dim flag As Boolean
flag = True
For a = 0 To 3
    If board(mat(a, 0) - 1, mat(a, 1)) Then
        flag = False
        a = 4
    End If
Next a

canLeft = flag
End Function

'is the piece able to be moved right?
Public Function canRight() As Boolean
Dim mat() As Integer
mat = piece.returnPiece

Dim flag As Boolean
flag = True
For a = 0 To 3
    If board(mat(a, 0) + 1, mat(a, 1)) Then
        flag = False
        a = 4
    End If
Next a

canRight = flag
End Function

'can the piece be moved down (via arrow key)
Public Function canDown() As Boolean
Dim mat() As Integer
mat = piece.returnPiece

Dim flag As Boolean
flag = True
For a = 0 To 3
    If board(mat(a, 0), mat(a, 1) + 1) Then
        flag = False
        a = 4
    End If
Next a

canDown = flag
End Function

'can the piece be rotated
Public Function canRotate() As Boolean
Dim mat() As Integer
mat = piece.returnPiece

Dim flag As Boolean
flag = True
For a = 0 To 3
    If board(mat(a, 0), mat(a, 1)) Then
        flag = False
        a = 4
    End If
Next a
canRotate = flag
End Function

'this checks to see if rows have been completed
Public Function dropRows() As Integer
    Dim flag As Boolean, mult As Integer
    flag = True
    'more points for doubles, triples, quads
    For b = 0 To 19
        For a = 0 To 9
            If board(a, b) = False Then
                flag = False
                a = 10
            End If
        Next a
        'scored a row!
        If flag Then
            lines = lines + 1
            mult = mult + 1
            For d = b To 0 Step -1
                For c = 0 To 9
                    board(c, d) = board(c, d - 1)
                    bcolor(c, d) = bcolor(c, d - 1)
                Next c
            Next d
        End If
        flag = True
    Next b
    
    score = score + ((mult ^ 2) * 50)
End Function

'is your board too high and no more pieces can be dropped?
Public Function lose()
    Dim mat() As Integer
    mat = piece.returnPiece
    For a = 0 To 3
        If board(mat(a, 0), mat(a, 1)) Then
            gameover = True
        End If
    Next a
End Function

'draw current stats
Public Function updateStats()
Dim ttime As Integer
    ttime = Int(time / 33)
    gamescreen.Enabled = True
        stime.Cls
        slines.Cls
        slevel.Cls
        sscore.Cls
        
        'time-------------------------------------------------------
        'thousands
        BitBlt stime.hDC, 27, 5, 5, 9, numbers.hDC, ((Int(ttime / 1000) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt stime.hDC, 27, 5, 5, 9, numbers.hDC, ((Int(ttime / 1000) Mod 10) * 5), 0, vbSrcPaint
        
        ttime = ttime - ((Int(ttime / 1000)) * 1000)
        'hundreds
        BitBlt stime.hDC, 32, 5, 5, 9, numbers.hDC, ((Int(ttime / 100) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt stime.hDC, 32, 5, 5, 9, numbers.hDC, ((Int(ttime / 100) Mod 10) * 5), 0, vbSrcPaint
        
        ttime = ttime - ((Int(ttime / 100)) * 100)
        'tens
        BitBlt stime.hDC, 37, 5, 5, 9, numbers.hDC, ((Int(ttime / 10) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt stime.hDC, 37, 5, 5, 9, numbers.hDC, ((Int(ttime / 10) Mod 10) * 5), 0, vbSrcPaint
        
        ttime = ttime - ((Int(ttime / 10)) * 10)
        'ones
        BitBlt stime.hDC, 42, 5, 5, 9, numbers.hDC, (ttime Mod 10) * 5 + 55, 0, vbSrcAnd
        BitBlt stime.hDC, 42, 5, 5, 9, numbers.hDC, (ttime Mod 10) * 5, 0, vbSrcPaint
        
        
        'lines---------------------------------------------------------
    Dim tlines As Integer
        tlines = lines
        'thousands
        BitBlt slines.hDC, 27, 5, 5, 9, numbers.hDC, ((Int(tlines / 1000) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt slines.hDC, 27, 5, 5, 9, numbers.hDC, ((Int(tlines / 1000) Mod 10) * 5), 0, vbSrcPaint
        
        tlines = tlines - ((Int(tlines / 1000)) * 1000)
        'hundreds
        BitBlt slines.hDC, 32, 5, 5, 9, numbers.hDC, ((Int(tlines / 100) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt slines.hDC, 32, 5, 5, 9, numbers.hDC, ((Int(tlines / 100) Mod 10) * 5), 0, vbSrcPaint
        
        tlines = tlines - ((Int(tlines / 100)) * 100)
        'tens
        BitBlt slines.hDC, 37, 5, 5, 9, numbers.hDC, ((Int(tlines / 10) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt slines.hDC, 37, 5, 5, 9, numbers.hDC, ((Int(tlines / 10) Mod 10) * 5), 0, vbSrcPaint
        
        tlines = tlines - ((Int(tlines / 10)) * 10)
        'ones
        BitBlt slines.hDC, 42, 5, 5, 9, numbers.hDC, (tlines Mod 10) * 5 + 55, 0, vbSrcAnd
        BitBlt slines.hDC, 42, 5, 5, 9, numbers.hDC, (tlines Mod 10) * 5, 0, vbSrcPaint


    'level------------------------------
        tlevel = Int(lines / 10)
        'hundreds
        BitBlt slevel.hDC, 32, 5, 5, 9, numbers.hDC, ((Int(tlevel / 100) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt slevel.hDC, 32, 5, 5, 9, numbers.hDC, ((Int(tlevel / 100) Mod 10) * 5), 0, vbSrcPaint
        
        tlevel = tlevel - ((Int(tlevel / 100)) * 100)
        'tens
        BitBlt slevel.hDC, 37, 5, 5, 9, numbers.hDC, ((Int(tlevel / 10) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt slevel.hDC, 37, 5, 5, 9, numbers.hDC, ((Int(tlevel / 10) Mod 10) * 5), 0, vbSrcPaint
        
        tlevel = tlevel - ((Int(tlevel / 10)) * 10)
        'ones
        BitBlt slevel.hDC, 42, 5, 5, 9, numbers.hDC, (tlevel Mod 10) * 5 + 55, 0, vbSrcAnd
        BitBlt slevel.hDC, 42, 5, 5, 9, numbers.hDC, (tlevel Mod 10) * 5, 0, vbSrcPaint


    'score------------------------------
    tscore = score
        'millions
        BitBlt sscore.hDC, 20, 5, 5, 9, numbers.hDC, ((Int(tscore / 1000000) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt sscore.hDC, 20, 5, 5, 9, numbers.hDC, ((Int(tscore / 1000000) Mod 10) * 5), 0, vbSrcPaint
        
        tscore = tscore - ((Int(tscore / 1000000)) * 1000000)
        'hundred thousands
        BitBlt sscore.hDC, 25, 5, 5, 9, numbers.hDC, ((Int(tscore / 100000) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt sscore.hDC, 25, 5, 5, 9, numbers.hDC, ((Int(tscore / 100000) Mod 10) * 5), 0, vbSrcPaint
        
        tscore = tscore - ((Int(tscore / 100000)) * 100000)
        'ten thousands
        BitBlt sscore.hDC, 30, 5, 5, 9, numbers.hDC, ((Int(tscore / 10000) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt sscore.hDC, 30, 5, 5, 9, numbers.hDC, ((Int(tscore / 10000) Mod 10) * 5), 0, vbSrcPaint
        
        tscore = tscore - ((Int(tscore / 10000)) * 10000)
        'thousands
        BitBlt sscore.hDC, 35, 5, 5, 9, numbers.hDC, ((Int(tscore / 1000) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt sscore.hDC, 35, 5, 5, 9, numbers.hDC, ((Int(tscore / 1000) Mod 10) * 5), 0, vbSrcPaint
        
        tscore = tscore - ((Int(tscore / 1000)) * 1000)
        'hundreds
        BitBlt sscore.hDC, 40, 5, 5, 9, numbers.hDC, ((Int(tscore / 100) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt sscore.hDC, 40, 5, 5, 9, numbers.hDC, ((Int(tscore / 100) Mod 10) * 5), 0, vbSrcPaint
        
        tscore = tscore - ((Int(tscore / 100)) * 100)
        'tens
        BitBlt sscore.hDC, 45, 5, 5, 9, numbers.hDC, ((Int(tscore / 10) Mod 10) * 5) + 55, 0, vbSrcAnd
        BitBlt sscore.hDC, 45, 5, 5, 9, numbers.hDC, ((Int(tscore / 10) Mod 10) * 5), 0, vbSrcPaint
        
        tscore = tscore - ((Int(tscore / 10)) * 10)
        'ones
        BitBlt sscore.hDC, 50, 5, 5, 9, numbers.hDC, (tscore Mod 10) * 5 + 55, 0, vbSrcAnd
        BitBlt sscore.hDC, 50, 5, 5, 9, numbers.hDC, (tscore Mod 10) * 5, 0, vbSrcPaint



End Function

Private Sub howto_Click()
    paused = True
    ipause.Visible = True
    howtoplay.Visible = True
    aboutme.Visible = False
End Sub

Private Sub newGame_Click()
    Erase board
    Erase bcolor
    gameover = False
    time = 0
    score = 0
    lines = 0
    nextup.Cls
    Call gameLoop
End Sub

Private Sub quit_Click()
    End
End Sub
