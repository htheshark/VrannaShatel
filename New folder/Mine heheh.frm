VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7392
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13668
   LinkTopic       =   "Form1"
   ScaleHeight     =   7392
   ScaleWidth      =   13668
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
'my code now
Private Sub Command1_Click()
    Dim a As Integer
    Const b = 0
    Dim c As Integer
    
    c = a / b
    
    End ' or try to
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
