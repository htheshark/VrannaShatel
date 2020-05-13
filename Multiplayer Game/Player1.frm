VERSION 5.00
Begin VB.Form Player1 
   BorderStyle     =   0  'None
   Caption         =   "1 (one) Palyer"
   ClientHeight    =   6384
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8148
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6384
   ScaleWidth      =   8148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   732
      Left            =   3000
      TabIndex        =   1
      Top             =   5160
      Width           =   2172
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chose!"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   25.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1932
   End
   Begin VB.Image S 
      Height          =   1692
      Left            =   5400
      Picture         =   "Player1.frx":0000
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1932
   End
   Begin VB.Shape SC 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   1932
      Left            =   5280
      Top             =   2160
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.Image P 
      Height          =   1692
      Left            =   3000
      Picture         =   "Player1.frx":3449
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1932
   End
   Begin VB.Shape PC 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   1932
      Left            =   2880
      Top             =   2160
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.Image R 
      Height          =   1692
      Left            =   600
      Picture         =   "Player1.frx":55DC
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1932
   End
   Begin VB.Shape RC 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   1932
      Left            =   480
      Top             =   2160
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.Image BackGround 
      Height          =   6372
      Left            =   0
      Picture         =   "Player1.frx":14B29
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8172
   End
End
Attribute VB_Name = "Player1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    gSinglePlayerC = 0
End Sub

Private Sub P_Click()
    RC.Visible = False
    PC.Visible = True
    SC.Visible = False
    gSinglePlayerC = 2
    Unload Me
    SinglePlayerFight.Show
End Sub

Private Sub R_Click()
    RC.Visible = True
    PC.Visible = False
    SC.Visible = False
    gSinglePlayerC = 1
    Unload Me
    SinglePlayerFight.Show
End Sub

Private Sub S_Click()
    RC.Visible = False
    PC.Visible = False
    SC.Visible = True
    gSinglePlayerC = 3
    Unload Me
    SinglePlayerFight.Show
End Sub

