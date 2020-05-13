VERSION 5.00
Begin VB.Form MainMenu 
   Caption         =   "Main Menu"
   ClientHeight    =   6384
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8184
   LinkTopic       =   "Form1"
   ScaleHeight     =   6384
   ScaleWidth      =   8184
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "2 (too) Palyers"
      Height          =   972
      Left            =   3120
      TabIndex        =   2
      Top             =   4800
      Width           =   2412
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1 (one) Palyer"
      Height          =   972
      Left            =   3120
      TabIndex        =   1
      Top             =   3600
      Width           =   2412
   End
   Begin VB.Image Image4 
      Height          =   972
      Left            =   6480
      Picture         =   "Multiplayer game.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1332
   End
   Begin VB.Image Image3 
      Height          =   852
      Left            =   5520
      Picture         =   "Multiplayer game.frx":3449
      Stretch         =   -1  'True
      Top             =   120
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   600
      Picture         =   "Multiplayer game.frx":55DC
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Rook Pepar and Skisers"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   792
      TabIndex        =   0
      Top             =   720
      Width           =   6612
   End
   Begin VB.Image Image2 
      Height          =   6372
      Left            =   0
      Picture         =   "Multiplayer game.frx":14B29
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8172
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
Player1.Show
End Sub

Private Sub Command2_Click()
Unload Me
Player2.Show
End Sub

Private Sub Form_Load()
    
    On Error GoTo AllGood
    MkDir "C:\Gaem"
        
    Dim strPath As String
    strPath = "C:\Gaem\Scores.txt"
    Open strPath For Output As #1
    Write #1, 0, 0, 0
    Close #1
    
AllGood:
End Sub
