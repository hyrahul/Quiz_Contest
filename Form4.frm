VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form "
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Click Here to Check Percentage of Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   0
      Top             =   4320
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command5_Click()
Dim a As Integer
 Dim b As Integer
 Dim c As Integer
 a = InputBox("Enter First Subject Number is ")
 Print a
b = InputBox("Enter Second Subject Number")
Print b
c = InputBox("Enter Third Subject Number")
Print c
MsgBox ("The Total of Three subject is " & a + b + c)
End Sub
