VERSION 5.00
Begin VB.Form student_login 
   Caption         =   "Form2"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13125
   LinkTopic       =   "Form2"
   ScaleHeight     =   7500
   ScaleWidth      =   13125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   8280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3840
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   3960
      Picture         =   "STudentLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT   LOGIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   6000
      TabIndex        =   8
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD          :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   5040
      TabIndex        =   7
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT ID       :  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   10395
      Left            =   0
      Picture         =   "STudentLogin.frx":2268
      Stretch         =   -1  'True
      Top             =   -2040
      Width           =   15810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TEACHER LOGIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   5280
      TabIndex        =   5
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME         :  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD          :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   4200
      TabIndex        =   3
      Top             =   3360
      Width           =   3375
   End
End
Attribute VB_Name = "student_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()
Dim strName As String
Dim strPass As String
Dim pesan As String



strName = Text1.Text
strPass = Text2.Text


Do Until rs.EOF
If rs.Fields(4).Value = strName And rs.Fields(5).Value = strPass Then
f_name = rs.Fields(0).Value
rollno = rs.Fields(2).Value
Unload Me
student_lab.Show
'if the login succeed then form that contain employee info shown
Exit Sub

Else
rs.MoveNext
End If

Loop

pesan = MsgBox("Invalid password, try again!")
If (pesan = 1) Then
Unload Me
Form2.Show
Text1.Text = ""
Text2.Text = ""

Else
End
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\USER\Desktop\student.mdb")
Set rs = db.OpenRecordset("select * from student")
End Sub

Private Sub Image4_Click()
Unload Me
homepage.Show
End Sub
