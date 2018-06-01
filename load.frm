VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12900
   LinkTopic       =   "Form3"
   ScaleHeight     =   9150
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   11280
      Top             =   480
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "Label6"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   4560
      Width           =   7335
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   4440
      Width           =   15
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   4440
      Width           =   15
   End
   Begin VB.Label Label2 
      Height          =   15
      Left            =   6240
      TabIndex        =   1
      Top             =   4920
      Width           =   5175
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   7080
      Width           =   6615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr, ctr2, r As Double
Dim ctr3 As String

Private Sub Image1_Click()

End Sub

Private Sub Timer1_Timer()
ctr = 0
If ctr2 <= 100 Then
Randomize
   r = Int((200 - 100 + 1) * Rnd + 100)
   ctr = r / 50
   ctr = Round(ctr, 0)
   ctr = ctr2 + ctr
   ctr3 = Str(Str)
     If ctr3 >= 100 Then
       Label6.Width = 5000
       homepage.Show
       Unload Me
     Else
       Label6.Width = Label6.Width + 1
       ctr2 = Int(ctr3)
     End If
End If

 
End Sub
