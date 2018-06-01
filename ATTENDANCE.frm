VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   10485
   ClientLeft      =   135
   ClientTop       =   570
   ClientWidth     =   16755
   LinkTopic       =   "Form6"
   ScaleHeight     =   10485
   ScaleWidth      =   16755
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ATTENDANCE.frx":0000
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483647
      HeadLines       =   1
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD ONE ATTENDANCE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   3000
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   6480
      Top             =   5280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USER\Desktop\student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USER\Desktop\student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "student"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   720
      Picture         =   "ATTENDANCE.frx":0015
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER STUDENT'S ROLL NUMBER  :"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT ATTENDANCE"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   12615
      Left            =   -4560
      Picture         =   "ATTENDANCE.frx":227D
      Stretch         =   -1  'True
      Top             =   -2160
      Width           =   21375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()
Dim pesan As String
Do Until rs.EOF
If rs.Fields(2).Value = Text1.Text Then
db.Execute "update student set attendance = attendance+1"
Unload Me
Form4.Show
Exit Sub
Else
rs.MoveNext

End If
Loop
pesan = MsgBox("Invalid Roll Number, try again!")
If (pesan = 1) Then
Unload Me
Form6.Show
Text1.Text = ""
Else
End
End If
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\USER\Desktop\student.mdb")
Set rs = db.OpenRecordset("select * from student")
End Sub

Private Sub Image2_Click()
Unload Me
homepage.Show
End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub
