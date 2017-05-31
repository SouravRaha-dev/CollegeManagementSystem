VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   17730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19650
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   51891.85
   ScaleMode       =   0  'User
   ScaleWidth      =   19650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   2160
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   70
      Left            =   4560
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   3360
      Top             =   240
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to College Management System"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   -1560
      TabIndex        =   5
      Top             =   1680
      Width           =   9135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      TabIndex        =   4
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   11640
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11400
      TabIndex        =   2
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11400
      TabIndex        =   1
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   0
      Top             =   4440
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, j, a, b, r, o As Integer

Private Sub Form_Click()
Unload Me
Me.Hide
Load Form2
Form2.Show

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
Me.Hide
Load Form2
Form2.Show
End Sub

Private Sub Form_Load()
Form1.BackColor = vbBlack
Form1.WindowState = 2
Label6.Top = 8500
Label6.Left = 11000

r = 300
End Sub



Private Sub Timer1_Timer()
With Label6
        If .Left <= 11000 Then .Left = .Left - 100
       If .Left < -5000 Then .Left = 11000
End With
End Sub


Private Sub Timer2_Timer()
Form1.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Timer4_Timer()
Dim Today As Variant
Today = Now
Label1.Caption = Format(Today, "dddd")
Label2.Caption = Format(Today, "mmmm")
Label3.Caption = Format(Today, "yyyy")
Label4.Caption = Format(Today, "d")
Label5.Caption = Format(Today, "hh:mm:ss ampm")


End Sub
