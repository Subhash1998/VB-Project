VERSION 5.00
Begin VB.Form student_lab 
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17625
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   17625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7695
      Left            =   4080
      TabIndex        =   0
      Top             =   1080
      Width           =   11295
      Begin VB.CommandButton Command2 
         Caption         =   "VIEW TIME TABLE"
         Height          =   735
         Left            =   3600
         TabIndex        =   6
         Top             =   6600
         Width           =   3855
      End
      Begin VB.ListBox List1 
         Height          =   285
         Left            =   6480
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Submit"
         Height          =   735
         Left            =   8400
         TabIndex        =   4
         Top             =   6600
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "d/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   2
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6480
         TabIndex        =   1
         Top             =   5040
         Width           =   2775
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   2880
         Picture         =   "student_lab.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8280
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   7620
         Left            =   -360
         Picture         =   "student_lab.frx":8EAE
         Stretch         =   -1  'True
         Top             =   120
         Width           =   12180
      End
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   4080
      Picture         =   "student_lab.frx":152DD
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   1575
   End
End
Attribute VB_Name = "student_lab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset

Private Sub Command1_Click()
rs.AddNew
'If Text2.Text = "dd/mm/yyyy" Then MsgBox ("Enter lab date")

'ElseIf Text1.Text = "" Then MsgBox ("Enter Computer no. assigned")
'End If



rs.Fields(0).Value = List1.Text
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text1.Text
rs.Fields(3).Value = rollno
rs.Update

MsgBox ("Record added successfully")
Unload Me
student_lab.Show

End Sub

Private Sub Command2_Click()
Form5.Show

End Sub

Private Sub Form_Load()
'Label1. = rs1.Fields(0).Value
Set db = OpenDatabase("C:\Users\USER\Desktop\student_lab.mdb")

Set rs = db.OpenRecordset("select * from s_lab")

'Set db1 = OpenDatabase("C:\Users\USER\Desktop\student.mdb")

'Set rs1 = db1.OpenRecordset("select * from student")

List1.AddItem "DLD Lab"
List1.AddItem "VB Lab"
List1.AddItem "C Lab"
End Sub

Private Sub Image2_Click()
Unload Me
homepage.Show

End Sub

Private Sub Label1_Click()
Label1.Caption = f_name
End Sub
