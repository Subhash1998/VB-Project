VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9855
   LinkTopic       =   "Form3"
   ScaleHeight     =   6030
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   120
      Top             =   600
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "KISHAN DEWANGAN   ,   PIYUSH OHRI"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   7335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ASHLY JUSTIN  ,   SUBHASH KSHATRI"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DESIGNED BY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800080&
      Height          =   615
      Left            =   480
      Top             =   4200
      Width           =   90
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   480
      Top             =   4200
      Width           =   9000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LAB MANAGEMENT SYSTEM "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   7815
      Left            =   -2040
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   13335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Shape2.Width = Shape2.Width + 90
If Shape2.Width = 9000 Then
Unload Me
homepage.Show
End If
End Sub
