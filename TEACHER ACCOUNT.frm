VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   LinkTopic       =   "Form4"
   ScaleHeight     =   7770
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "VIEW TIME TABLE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   6960
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Height          =   285
      Left            =   6480
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   3240
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   285
      Left            =   6480
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   11
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ATTENDANCE SHEET"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   10
      Top             =   6960
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   4920
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "Option1"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   6600
      TabIndex        =   6
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   2760
      Picture         =   "TEACHER ACCOUNT.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   9240
      Picture         =   "TEACHER ACCOUNT.frx":2268
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   9240
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "BATCH 2"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "BATCH 1"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7320
      TabIndex        =   7
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BATCH                 :"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE                    :"
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
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SEMESTER         :"
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
      Height          =   735
      Left            =   3000
      TabIndex        =   2
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LAB NAME         :"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TEACHER ACCOUNT"
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
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   10935
      Left            =   -720
      Picture         =   "TEACHER ACCOUNT.frx":B116
      Stretch         =   -1  'True
      Top             =   -3120
      Width           =   15495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db1 As Database
Public rs As Recordset

Private Sub Command1_Click()
Form5.Show

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo2_Change()

End Sub

Private Sub Command2_Click()
Form6.Show
End Sub

Private Sub Command3_Click()
rs.AddNew
rs.Fields(0).Value = List1.Text
rs.Fields(1).Value = List2.Text
rs.Fields(2).Value = Text1.Text
If Option1 Then
rs.Fields(3).Value = "Batch1"
Else
rs.Fields(3).Value = "Batch2"
End If
Unload Me
Form4.Show
End Sub

Private Sub Form_Load()
List1.AddItem "PROGRAMMING LAB"
List1.AddItem "DIGITAL LOGIC & DESIGN"
List1.AddItem "C PROGRAMMING LAB"

List2.AddItem "1"
List2.AddItem "2"
List2.AddItem "3"
List2.AddItem "4"
List2.AddItem "5"
List2.AddItem "6"
List2.AddItem "7"
List2.AddItem "8"
List2.AddItem "9"
List2.AddItem "10"
Set db = OpenDatabase("C:\Users\USER\Desktop\teacher_account.mdb")
Set rs = db.OpenRecordset("select * from account")
End Sub

Private Sub Image3_Click()
Unload Me
Form2.Show
End Sub

Private Sub Image4_Click()
Unload Me
homepage.Show
End Sub
