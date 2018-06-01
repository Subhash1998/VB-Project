VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11220
   LinkTopic       =   "Form2"
   ScaleHeight     =   5955
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
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
      Left            =   5880
      TabIndex        =   5
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   3840
      Picture         =   "TEACHER LOGIN 2.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
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
      TabIndex        =   2
      Top             =   3360
      Width           =   3375
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
      TabIndex        =   1
      Top             =   2160
      Width           =   3015
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
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   12195
      Left            =   -2280
      Picture         =   "TEACHER LOGIN 2.frx":2268
      Stretch         =   -1  'True
      Top             =   -3120
      Width           =   22770
   End
End
Attribute VB_Name = "Form2"
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
'pesan is Indonesian and it is the same with Message in english


strName = Text1.Text
strPass = Text2.Text


Do Until rs.EOF
If rs.Fields(4).Value = strName And rs.Fields(5).Value = strPass Then
username = rs.Fields(4).Value
f_name = rs.Fields(0).Value
l_name = rs.Fields(1).Value
dept = rs.Fields(2).Value
Unload Me
Form7.Show
'if the login succeed then form that contain employee info shown
Exit Sub

Else
rs.MoveNext
End If

Loop

pesan = MsgBox("Invalid password, try again!")
If (pesan = 1) Then
Form2.Show
Text1.Text = ""
Text2.Text = ""

Else
End
End If

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\USER\Desktop\signup.mdb")
Set rs = db.OpenRecordset("select * from signup")
End Sub

Private Sub Image2_Click()
Unload Me
homepage.Show
End Sub
