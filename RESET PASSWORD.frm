VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   LinkTopic       =   "Form3"
   ScaleHeight     =   6060
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SET PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NEW PASSWORD   :"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME            :"
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
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   0
      Picture         =   "RESET PASSWORD.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub
