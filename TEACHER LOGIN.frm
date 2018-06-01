VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "VIEW STUDENT DETAILS"
      Height          =   495
      Left            =   10320
      TabIndex        =   13
      Top             =   7560
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   6240
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   4800
      Width           =   2895
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9960
      TabIndex        =   8
      Top             =   5520
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "Option1"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   7800
      TabIndex        =   7
      Top             =   5520
      Width           =   255
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7800
      TabIndex        =   6
      Top             =   4080
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7800
      TabIndex        =   5
      Top             =   3360
      Width           =   2895
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
      Left            =   5160
      TabIndex        =   14
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "BATCH 2"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   10440
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "BATCH 1"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENDANCE    :        (ENTER STUDENT'S ID)"
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
      Height          =   855
      Left            =   4200
      TabIndex        =   4
      Top             =   6360
      Width           =   2775
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
      Left            =   4200
      TabIndex        =   3
      Top             =   5640
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
      Left            =   4200
      TabIndex        =   2
      Top             =   4920
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
      Left            =   4200
      TabIndex        =   1
      Top             =   4200
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
      Left            =   4200
      TabIndex        =   0
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   10395
      Left            =   0
      Picture         =   "TEACHER LOGIN.frx":0000
      Top             =   -120
      Width           =   15090
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()

End Sub

Private Sub Form_Load()
Combo1.AddItem "PROGRAMMING LAB"
Combo1.AddItem "DIGITAL LOGIC & DESIGN"
Combo1.AddItem "C PROGRAMMING LAB"

Combo2.AddItem "1"
Combo2.AddItem "2"
Combo2.AddItem "3"
Combo2.AddItem "4"
Combo2.AddItem "5"
Combo2.AddItem "6"
Combo2.AddItem "7"
Combo2.AddItem "8"
Combo2.AddItem "9"
Combo2.AddItem "10"
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Text1_Change()
Dim MyDate As Date
  MyDate = DateSerial(intYear, intMonth, intDay)
End Sub
