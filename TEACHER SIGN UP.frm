VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14295
   DrawMode        =   8  'Xor Pen
   DrawStyle       =   6  'Inside Solid
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   285
      Left            =   9000
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   9000
      TabIndex        =   12
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "SIGN UP"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11280
      TabIndex        =   10
      Top             =   7560
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   9000
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   9000
      TabIndex        =   7
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9000
      TabIndex        =   5
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   3120
      Picture         =   "TEACHER SIGN UP.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "E-MAIL ID             :"
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
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TEACHER SIGN UP DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   735
      Left            =   5160
      TabIndex        =   9
      Top             =   480
      Width           =   7575
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
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
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT     :"
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
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "TEACHER ID        :"
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
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME          :"
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
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME        :"
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
      Left            =   6120
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   11835
      Left            =   -1080
      Picture         =   "TEACHER SIGN UP.frx":2268
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   18330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub HScroll1_Change()

End Sub
Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub Command1_Click()
 rs.AddNew
 rs.Fields(0).Value = Text1.Text
 rs.Fields(1).Value = Text2.Text
 rs.Fields(2).Value = List1.Text
 rs.Fields(3).Value = Text5.Text
 rs.Fields(4).Value = Text3.Text
 rs.Fields(5).Value = Text4.Text
 rs.Update
 MsgBox ("Signed Up Successfully!!!")
 Unload Me
 homepage.Show
End Sub

Private Sub Form_Load()
List1.AddItem "COMPUTER SCIENCE AND ENGINEERING"
List1.AddItem "ELECTRONICS & TELECOMMUNICATION ENGINEERING"
List1.AddItem "CIVIL ENGINEERING"
List1.AddItem "MECHANICAL ENGINEERING"
List1.AddItem "ELECTRICAL ENGINEERING"
List1.AddItem "MINING ENGINEERING"
List1.AddItem "METALLURGICAL ENGINEERING"
List1.AddItem "CHEMICAL ENGINEERING"
Set db = OpenDatabase("C:\Users\USER\Desktop\signup.mdb")
Set rs = db.OpenRecordset("select * from signup")
End Sub

Private Sub Image2_Click()
Unload Me
homepage.Show

End Sub

