VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form HTMLQues 
   Caption         =   "HTML Language"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   16305
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   14520
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=root;User ID=system;Data Source=localhost;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=root;User ID=system;Data Source=localhost;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "system"
      Password        =   "root"
      RecordSource    =   "HTMLQues"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Score Board"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   14040
      TabIndex        =   13
      Top             =   1320
      Width           =   3735
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1800
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Attempt"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Correct Answer"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   12
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   11
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change_Lan"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   10
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   9
      Top             =   8160
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   795
      Left            =   7200
      TabIndex        =   8
      Top             =   7080
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   7
      Top             =   7200
      Width           =   1695
   End
   Begin VB.OptionButton Option4 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   6
      Top             =   5880
      Width           =   3975
   End
   Begin VB.OptionButton Option3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   5
      Top             =   5880
      Width           =   3855
   End
   Begin VB.OptionButton Option2 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   4
      Top             =   4560
      Width           =   3975
   End
   Begin VB.OptionButton Option1 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   3
      Top             =   4560
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   9975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "User Request"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HTML Question Page"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   18
      Top             =   240
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   5640
      Picture         =   "HTMLQues.frx":0000
      Top             =   7200
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      Height          =   2895
      Left            =   2880
      Top             =   1200
      Width           =   10215
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   2160
      Picture         =   "HTMLQues.frx":2933
      Top             =   0
      Width           =   16500
   End
End
Attribute VB_Name = "HTMLQues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connbook As ADODB.Connection
Dim cmd As ADODB.Command
Dim book As ADODB.Recordset
Dim s As Integer
Dim o As Integer
Dim l As Integer
Private Sub Command1_Click()
Text2.Text = ""
Option1.Caption = ""
Option2.Caption = ""
Option3.Caption = ""
Option4.Caption = ""
Option1.BackColor = vbWhite
Option2.BackColor = vbWhite
Option3.BackColor = vbWhite
Option4.BackColor = vbWhite
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Text1.Enabled = True
Text3.Visible = False
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Command2.Enabled = False
Frame1.Visible = False

Set book = New ADODB.Recordset
book.Open "select * from HTMLQues where sr_no = '" & Text1.Text & "'", connbook, adOpenKeyset, adLockReadOnly, adCmdText
If book.RecordCount <> 0 Then
Text1.Text = Val(book!sr_no)
Text2.Text = Trim(book!ques)
Option1.Caption = Trim(book!a)
Option2.Caption = Trim(book!b)
Option3.Caption = Trim(book!c)
Option4.Caption = Trim(book!d)
Text3.Text = Trim(book!ans)
connbook.Execute "commit"
Else
MsgBox "sr_no  Not Exists..."
Text2.Text = ""
Text3.Text = ""
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
End If
book.Close
Set book = Nothing
End Sub

Private Sub Command2_Click()
Text3.Visible = True
Text1.Enabled = False
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Option1.Caption = ""
Option2.Caption = ""
Option3.Caption = ""
Option4.Caption = ""
Option1.BackColor = vbWhite
Option2.BackColor = vbWhite
Option3.BackColor = vbWhite
Option4.BackColor = vbWhite
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Text3.Text = ""
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
End Sub

Private Sub Command4_Click()
Sel_Lan.Show
Unload Me
End Sub

Private Sub Command5_Click()
Text4.Text = s
Text5.Text = o
Frame1.Visible = True
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set connbook = New ADODB.Connection
connbook.Open "Provider=MSDAORA.1;Password=root;User ID=system;Data Source=localhost;Persist Security Info=True"
connbook.CursorLocation = adUseClient
MsgBox "Connection Established..."
Frame1.Visible = False
Text2.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option1.BackColor = vbWhite
Option2.BackColor = vbWhite
Option3.BackColor = vbWhite
Option4.BackColor = vbWhite
End Sub

Private Sub Option1_Click()
If Option1.Enabled Then
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Command2.Enabled = True
End If
If Text3.Text = Option1.Caption Then
s = s + 1
o = o + 1
Option1.BackColor = vbGreen
Else
o = o + 1
Option1.BackColor = vbRed
End If
End Sub

Private Sub Option2_Click()
If Option2.Enabled Then
Option1.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Command2.Enabled = True
End If
If Text3.Text = Option2.Caption Then
s = s + 1
o = o + 1
Option2.BackColor = vbGreen
Else
o = o + 1
Option2.BackColor = vbRed
End If
End Sub

Private Sub Option3_Click()
If Option3.Enabled Then
Option1.Enabled = False
Option2.Enabled = False
Option4.Enabled = False
Command2.Enabled = True
End If
If Text3.Text = Option3.Caption Then
s = s + 1
o = o + 1
Option3.BackColor = vbGreen
Else
o = o + 1
Option3.BackColor = vbRed
End If
End Sub

Private Sub Option4_Click()
If Option4.Enabled Then
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Command2.Enabled = True
End If
If Text3.Text = Option4.Caption Then
s = s + 1
o = o + 1
Option4.BackColor = vbGreen
Else
o = o + 1
Option4.BackColor = vbRed
End If
End Sub

