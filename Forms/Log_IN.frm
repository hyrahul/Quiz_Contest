VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Log_in 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Log_In Into Quiz"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   4560
   FillColor       =   &H00FFFF00&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7815
      Left            =   4320
      TabIndex        =   0
      Top             =   3000
      Width           =   12105
      Begin VB.CommandButton Command3 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         TabIndex        =   8
         Top             =   5160
         Width           =   1935
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   0
         Top             =   7320
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1085
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
         RecordSource    =   "log_reg"
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
      Begin VB.CommandButton Command2 
         Caption         =   " Register Now"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   8520
         TabIndex        =   6
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Log_In"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   5
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   675
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   4080
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         DataField       =   "USER_NAME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   675
         Left            =   2880
         TabIndex        =   2
         Top             =   2880
         Width           =   3975
      End
      Begin VB.Line Line1 
         X1              =   7680
         X2              =   7680
         Y1              =   2280
         Y2              =   6240
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "New User ?"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   8400
         TabIndex        =   9
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Quiz System"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   1575
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   8295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   " Password    :"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "  User_Name   :"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   360
         TabIndex        =   1
         Top             =   2880
         Width           =   2295
      End
   End
   Begin VB.Image Image4 
      Height          =   8100
      Left            =   16440
      Picture         =   "Log_IN.frx":0000
      Top             =   3000
      Width           =   4200
   End
   Begin VB.Image Image3 
      Height          =   7935
      Left            =   5400
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   8100
      Left            =   0
      Picture         =   "Log_IN.frx":6214
      Top             =   3000
      Width           =   4350
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "Log_IN.frx":D692
      Top             =   0
      Width           =   21000
   End
End
Attribute VB_Name = "Log_in"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connbook As ADODB.Connection
Dim cmd As ADODB.Command
Dim book As ADODB.Recordset

Private Sub Command1_Click()

Set book = New ADODB.Recordset
book.Open "select * from log_reg where user_name = '" & Text1.Text & "'", connbook, adOpenKeyset, adLockReadOnly, adCmdText
If book.RecordCount <> 0 Then
Text1.Text = Trim(book!User_name)
Text2.Text = Trim(book!Password)
connbook.Execute "commit"
MsgBox "Login Success"
Sel_Lan.Show
Unload Me

Else
MsgBox "Login Denied"
Text1.Text = ""
Text2.Text = ""
Log_in.Show
End If

book.Close
Set book = Nothing
End Sub




Private Sub Command2_Click()
Reg.Show
Unload Me

End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Form_Load()
Set connbook = New ADODB.Connection
connbook.Open "Provider=MSDAORA.1;Password=root;User ID=system;Data Source=localhost;Persist Security Info=True"
connbook.CursorLocation = adUseClient
MsgBox "Connection Established..."
End Sub

