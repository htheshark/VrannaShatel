VERSION 5.00
Begin VB.Form SinglePlayerFight 
   BorderStyle     =   0  'None
   Caption         =   "Fight"
   ClientHeight    =   6384
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8148
   LinkTopic       =   "Form2"
   ScaleHeight     =   6384
   ScaleWidth      =   8148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDie 
      Caption         =   "Iv'e had enough"
      Height          =   612
      Left            =   4920
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.CommandButton cmdAgain 
      Caption         =   "Again"
      Height          =   612
      Left            =   2160
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Timer CompMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6240
      Top             =   4080
   End
   Begin VB.Timer PlayerMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   3960
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   492
      Left            =   240
      TabIndex        =   4
      Top             =   5760
      Width           =   1572
   End
   Begin VB.Label lblWinner 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2640
      TabIndex        =   1
      Top             =   4080
      Width           =   3252
   End
   Begin VB.Label lblVS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V.S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3840
      TabIndex        =   0
      Top             =   2040
      Width           =   732
   End
   Begin VB.Image imgCompC 
      Height          =   2172
      Left            =   4968
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2652
   End
   Begin VB.Image imgPlayerC 
      Height          =   2172
      Left            =   768
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2652
   End
   Begin VB.Image P 
      Height          =   732
      Left            =   6240
      Picture         =   "SinglePlayerFight.frx":0000
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Image S 
      Height          =   732
      Left            =   7200
      Picture         =   "SinglePlayerFight.frx":2193
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Image R 
      Height          =   732
      Left            =   5280
      Picture         =   "SinglePlayerFight.frx":55DC
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Image BackGround 
      Height          =   6372
      Left            =   0
      Picture         =   "SinglePlayerFight.frx":14B29
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8172
   End
End
Attribute VB_Name = "SinglePlayerFight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CompC As Integer
Dim mintResults As Integer
Dim strPath As String

Private Sub cmdAgain_Click()
    Player1.Show
    Unload Me
End Sub

Private Sub cmdDie_Click()
    MainMenu.Show
    Unload Me
End Sub

Private Sub CompMove_Timer()
    imgCompC.Left = imgCompC.Left - 4
    If imgCompC.Left < 4547 Then CompMove.Enabled = False

End Sub

'1 = Rock
'2 = Paper
'3 = Scissors

Private Sub Form_Load()
    Randomize Timer
    
    strPath = "C:\Gaem\Scores.txt"
    
    mintResults = 0
    CompC = Int(Rnd * 3) + 1
    'CompC = 1
    
    'mint Results
    '1 = tie
    '2 = lose
    '3 = win
    
'Main Winner Test Loop
    If gSinglePlayerC = 1 And CompC = 1 Then 'Tie with Rocks               ROCK
        
        imgPlayerC.Picture = R.Picture
        imgCompC.Picture = R.Picture
        
        PlayerMove.Enabled = True
        CompMove.Enabled = True
        
        mintResults = 1
        
    ElseIf gSinglePlayerC = 1 And CompC = 2 Then 'Human loses to Paper     ROCK
        
        imgPlayerC.Picture = R.Picture
        imgCompC.Picture = P.Picture
        
        PlayerMove.Enabled = True
        CompMove.Enabled = True
        
        mintResults = 2
        
    ElseIf gSinglePlayerC = 1 And CompC = 3 Then 'Human Wins to Scissors     ROCK
        
        imgPlayerC.Picture = R.Picture
        imgCompC.Picture = S.Picture
        
        PlayerMove.Enabled = True
        CompMove.Enabled = True
        
        mintResults = 3
    
    ElseIf gSinglePlayerC = 2 And CompC = 1 Then 'Human wins to Rock          PAPER
        
        imgPlayerC.Picture = P.Picture
        imgCompC.Picture = R.Picture
        
        PlayerMove.Enabled = True
        CompMove.Enabled = True
        
        mintResults = 3
        
    ElseIf gSinglePlayerC = 2 And CompC = 2 Then 'Ties with Paper          PAPER
        
        imgPlayerC.Picture = P.Picture
        imgCompC.Picture = P.Picture
        
        PlayerMove.Enabled = True
        CompMove.Enabled = True
        
        mintResults = 1
        
    ElseIf gSinglePlayerC = 2 And CompC = 3 Then 'Human Loses to Scissors          PAPER
        
        imgPlayerC.Picture = P.Picture
        imgCompC.Picture = S.Picture
        
        PlayerMove.Enabled = True
        CompMove.Enabled = True
        
        mintResults = 2
        
    ElseIf gSinglePlayerC = 3 And CompC = 1 Then 'Loses to Rock          SCISSORS
        
        imgPlayerC.Picture = S.Picture
        imgCompC.Picture = R.Picture
        
        PlayerMove.Enabled = True
        CompMove.Enabled = True
        
        mintResults = 2
    
    ElseIf gSinglePlayerC = 3 And CompC = 2 Then 'Human wins to Paper          SCISSORS
        
        imgPlayerC.Picture = S.Picture
        imgCompC.Picture = P.Picture
        
        PlayerMove.Enabled = True
        CompMove.Enabled = True
        
        mintResults = 3
        
    ElseIf gSinglePlayerC = 3 And CompC = 3 Then 'Ties         SCISSORS
        
        imgPlayerC.Picture = S.Picture
        imgCompC.Picture = S.Picture
        
        PlayerMove.Enabled = True
        CompMove.Enabled = True
        
        mintResults = 1
    End If
      
End Sub

Private Sub PlayerMove_Timer()
    imgPlayerC.Left = imgPlayerC.Left + 4
    
    If imgPlayerC.Left > (3817 - imgPlayerC.Width) Then
        PlayerMove.Enabled = False
        Call Results
    End If
End Sub

Private Sub Results()

'Handles Labels and Updates Scores
    
    Dim Win As Integer
    Dim Tie As Integer
    Dim Lose As Integer
    
    Open strPath For Input As #1
    Input #1, Win, Tie, Lose
    Close #1
    
    Select Case mintResults
        Case 1
            lblWinner.Caption = "Tie!"
            Tie = Tie + 1
        Case 2
            lblWinner.Caption = "Loser! :)"
            Lose = Lose + 1
        Case 3
            lblWinner.Caption = "Winner. :("
            Win = Win + 1
    End Select
    
    Open strPath For Output As #1
    Write #1, Win, Tie, Lose,
    Close #1
    
    lblScores.Caption = Win & " - " & Tie & " - " & Lose
    cmdAgain.Visible = True
    cmdDie.Visible = True
End Sub
