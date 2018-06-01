VERSION 5.00
Begin VB.Form homepage 
   Caption         =   "Form1"
   ClientHeight    =   10575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16200
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   16200
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   9600
      TabIndex        =   1
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   2160
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   2
      Top             =   2040
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   495
      Index           =   1
      Left            =   9480
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   5640
      Picture         =   "vb project.frx":0000
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2085
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   2055
      Index           =   0
      Left            =   9720
      Picture         =   "vb project.frx":8836
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   10800
      Left            =   -360
      Picture         =   "vb project.frx":11C14
      Top             =   -360
      Width           =   15660
   End
End
Attribute VB_Name = "homepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub option1_Click(Index As Integer)

End Sub

Private Sub Image2_Click(Index As Integer)
If Check1 Then
If Check2 Then
MsgBox ("SELECT ONE CHOICE")
Unload Me
homepage.Show
End If
End If
If Check1 Then
Unload Me
student_signup.Show
ElseIf Check2 Then
Unload Me
student_login.Show
End If
End Sub

Private Sub Image3_Click()
If Check1 Then
If Check2 Then
MsgBox ("SELECT ONE CHOICE")
Unload Me
homepage.Show
End If
End If
If Check1 Then
Unload Me
Form1.Show
ElseIf Check2 Then
Unload Me
Form2.Show
End If
End Sub
