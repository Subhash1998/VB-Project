VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10515
   LinkTopic       =   "Form7"
   ScaleHeight     =   7425
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Update Lab Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   7
      Top             =   6480
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   5280
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   4440
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   2640
      Picture         =   "Form7.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   8520
      Picture         =   "Form7.frx":2268
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   3360
      Picture         =   "Form7.frx":B116
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Department    :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name     :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name     :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   -840
      Picture         =   "Form7.frx":1394C
      Stretch         =   -1  'True
      Top             =   -2640
      Width           =   14775
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()
Unload Me
Form4.Show

End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\USER\Desktop\signup.mdb")
Set rs = db.OpenRecordset("select * from signup")
Text4.Text = username
Text1.Text = f_name
Text2.Text = l_name
Text3.Text = dept
End Sub

Private Sub Image3_Click()
Unload Me
Form2.Show

End Sub

Private Sub Image4_Click()
Unload Me
homepage.Show
End Sub
