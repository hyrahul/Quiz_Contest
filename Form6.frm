VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   4920
      TabIndex        =   9
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Average"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   5160
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5160
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Total Sum of Three Subject is"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Maths"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Computer Science"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "English"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text4.Text = Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text)
Text5.Text = Text4.Text / 3
End Sub
