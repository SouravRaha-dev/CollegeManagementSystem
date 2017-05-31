VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Text            =   "Developed By- Sourav Raha, Subha Ganguly and Srijita Misra"
      Top             =   5520
      Width           =   14175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form4abtus.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   14415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About Us"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   24000
      Left            =   0
      Picture         =   "Form4abtus.frx":03C4
      Top             =   0
      Width           =   38400
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
Attribute VB_Name = "Form4"
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

Private Sub Form_Resize()
Image1.Top = 0
Image1.Left = 0
Image1.Width = Me.ScaleWidth
Image1.Height = Me.ScaleHeight
End Sub

Private Sub Home_Click()
Unload Me
Me.Hide
Load Form3
Form3.Show

End Sub

Private Sub Maximize_Click()
Form4.WindowState = 2

End Sub

Private Sub Minimize_Click()
Form4.WindowState = 1
End Sub

Private Sub Restore_Click()
Form4.WindowState = 0
End Sub

