VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form3"
   ClientHeight    =   9885
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15180
   LinkTopic       =   "Form3"
   ScaleHeight     =   9885
   ScaleWidth      =   15180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "&Log-out"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "&Student Records"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9000
      Picture         =   "Form3main.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "&Student Registration"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2520
      Picture         =   "Form3main.frx":0D99
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   2730
      Left            =   9000
      Picture         =   "Form3main.frx":19B8
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   4995
   End
   Begin VB.Image Image2 
      Height          =   2700
      Left            =   1200
      Picture         =   "Form3main.frx":3A9C
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   5280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MSIT Online Management Portal"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   1560
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   2220
      Left            =   120
      Picture         =   "Form3main.frx":222B9
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2700
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
   End
   Begin VB.Menu About 
      Caption         =   "About Us"
   End
   Begin VB.Menu Window 
      Caption         =   "Window"
      Begin VB.Menu Maximize 
         Caption         =   "Maximize"
      End
      Begin VB.Menu Minimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu Restore 
         Caption         =   "Restore"
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
Unload Me
Me.Hide
Load Form4
Form4.Show

End Sub

Private Sub Command1_Click()
Form5.Show

End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
Load Form6
Form6.Show

End Sub

Private Sub Command3_Click()
Unload Me
Me.Hide
Load Form2
Form2.Show

End Sub

Private Sub Exit_Click()
Unload Me
Me.Hide
End

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = 12000
Me.Width = 15000
End Sub

Private Sub Home_Click()
Load Form3
Form3.Show

End Sub

Private Sub Maximize_Click()
Form3.WindowState = 2

End Sub

Private Sub Minimize_Click()
Form3.WindowState = 1

End Sub

Private Sub Restore_Click()
Form3.WindowState = 0

End Sub
