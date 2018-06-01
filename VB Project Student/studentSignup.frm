VERSION 5.00
Begin VB.Form student_signup 
   Caption         =   "Form3"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13935
   LinkTopic       =   "Form3"
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   285
      Left            =   8280
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   3960
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8280
      TabIndex        =   12
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIGN UP"
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
      Left            =   6960
      TabIndex        =   10
      Top             =   6840
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   8280
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   5640
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   8280
      TabIndex        =   8
      Top             =   4680
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8280
      TabIndex        =   7
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   8280
      TabIndex        =   6
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ROLL NUMBER    :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   3720
      Picture         =   "studentSignup.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD         :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT ID       :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT    :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME          :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME         :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT   SIGN   UP"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   10695
      Left            =   -2040
      Picture         =   "studentSignup.frx":2268
      Stretch         =   -1  'True
      Top             =   -2160
      Width           =   21135
   End
End
Attribute VB_Name = "student_signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db2 As Database
Public rs2 As Recordset

Private Sub Command1_Click()
rs2.AddNew
rs2.Fields(0).Value = Text1.Text
rs2.Fields(1).Value = Text2.Text
rs2.Fields(2).Value = Text5.Text
rs2.Fields(3).Value = List1.Text
rs2.Fields(4).Value = Text3.Text
rs2.Fields(5).Value = Text4.Text
rs2.Fields(6).Value = 0
rs2.Update
MsgBox ("Signed up successfully!!!")
Unload Me
homepage.Show

End Sub

Private Sub Form_Load()
List1.AddItem "Architecture"
List1.AddItem "Biomedical Engineering"
List1.AddItem "Biotech Engineering"
List1.AddItem "Civil Engineering"
List1.AddItem "Computer Science Engineering"
List1.AddItem "Chemicsl Engineering"
List1.AddItem "Electrical Engineering"
List1.AddItem "Electronics And Telecommunication Engineering"
List1.AddItem "Mining Engineering"
List1.AddItem "Metallurgical Engineering"
List1.AddItem "Information Technology"

Set db2 = OpenDatabase("C:\Users\USER\Desktop\student.mdb")

Set rs2 = db2.OpenRecordset("select * from student")
End Sub

Private Sub Image1_Click()
Unload Me
homepage.Show
End Sub

Private Sub Image2_Click()
Unload Me
homepage.Show

End Sub
