VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "What does this button do?"
      Height          =   855
      Left            =   5880
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   3480
      Picture         =   "VrannaShatel.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'hehe you lookin at my code
Private Sub Command1_Click()
    Image1.Visible = True
    Command1.Visible = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
