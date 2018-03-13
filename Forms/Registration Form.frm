VERSION 5.00
Begin VB.Form Reg 
   Caption         =   "Reg"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Frame1"
      Height          =   9615
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   13935
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   6480
         TabIndex        =   7
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Height          =   615
         Left            =   6480
         TabIndex        =   6
         Top             =   6360
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3240
         TabIndex        =   5
         Top             =   7680
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Height          =   615
         Left            =   6480
         TabIndex        =   4
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         Height          =   615
         Left            =   6480
         TabIndex        =   3
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   6480
         TabIndex        =   2
         Top             =   5400
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   6480
         TabIndex        =   1
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Registration Form"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3360
         TabIndex        =   14
         Top             =   480
         Width           =   7335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "            Name        :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   13
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "User_Name  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3660
         TabIndex        =   12
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "City              :"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Address     :"
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
         Left            =   4080
         TabIndex        =   10
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Contact_No   :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   9
         Top             =   6480
         Width           =   3015
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "      Password     :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   8
         Top             =   2640
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connbook As ADODB.Connection
Dim cmd As ADODB.Command
Dim book As ADODB.Recordset
Private Sub Command1_Click()
Set book = New ADODB.Recordset
book.Open "select * from log_reg where user_name = '" & Text2.Text & "'", connbook, adOpenKeyset, adLockReadOnly, adCmdText
If book.RecordCount <> 0 Then
MsgBox "User_Name Already Exists..."
book.Close
Set book = Nothing
Exit Sub
Else
Set book = New ADODB.Recordset
book.Open "select * from log_reg where user_name = '" & Text2.Text & "'", connbook, adOpenKeyset, adLockPessimistic, adCmdText
book.AddNew
book!Name = Trim(Text1.Text)
book!User_name = Trim(Text2.Text)
book!City = Trim(Text3.Text)
book!Address = Trim(Text4.Text)
book!Contact_No = Val(Text5.Text)
book!Password = Trim(Text6.Text)
book.Update
connbook.Execute "commit"
book.Close
Set book = Nothing
MsgBox "Added Succesfully..."
Log_in.Show
End If
End Sub

Private Sub Form_Load()
Set connbook = New ADODB.Connection
connbook.Open "Provider=MSDAORA.1;Password=root;User ID=system;Data Source=localhost;Persist Security Info=True"
connbook.CursorLocation = adUseClient
MsgBox "Connection Established..."
End Sub

