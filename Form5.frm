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
   Begin VB.CommandButton Command1 
      Caption         =   "Find Length of Enter Chracter"
      Height          =   615
      Left            =   6480
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Size Of Your Enter Chracter"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your Favourite Food Name"
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = Len(Text1.Text)
End Sub
