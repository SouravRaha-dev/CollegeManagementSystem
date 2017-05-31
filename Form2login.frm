VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   6675
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   19845
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   19845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Text            =   "Password"
      Top             =   4680
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Text            =   "Username"
      Top             =   3840
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Log-In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      Height          =   2055
      Left            =   10800
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   2055
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM LOG-IN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   2280
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   8820
      Left            =   0
      Picture         =   "Form2login.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   11040
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag1 As Boolean
Dim flag2 As Boolean
Dim cnt As Integer
Public Sub Clear()

Text1.Text = "Username"
Text1.FontItalic = True
Text1.ForeColor = &H808080

Text2.Text = "Password"
Text2.FontItalic = True
Text2.ForeColor = &H808080
Text2.PasswordChar = ""

flag1 = True
flag2 = True
End Sub

Private Sub Command1_Click()
cnt = cnt + 1
If Text1.Text = "admin" And Text2.Text = "12345" And cnt <= 4 Then
    Form3.Show
    Unload Me
    Me.Hide
Else
    If cnt >= 3 Then
    MsgBox "Too may tries, Sorry try again! "
    Unload Me
    Me.Hide
    End
    Else
    If cnt >= 2 Then
    MsgBox ("Invalid Username or Password-" & (3 - cnt) & " try left !")
    Else
    MsgBox ("Invalid Username or Password-" & (3 - cnt) & " tries left !")
    End If
    Call Clear
    End If
End If
End Sub

Private Sub Command2_Click()
Call Clear
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = 12000
Me.Width = 15000
Text1.FontItalic = True
Text1.ForeColor = &H808080
Text1.FontSize = "12"
Text2.FontItalic = True
Text2.ForeColor = &H808080
Text2.FontSize = "12"
flag1 = True
flag2 = True
cnt = 0
End Sub

Private Sub Form_Resize()
 Image1.Top = 0
    Image1.Left = 0
    Image1.Width = Me.ScaleWidth
    Image1.Height = Me.ScaleHeight
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 8 And flag1 Then
Text1.Text = ""
Text1.FontItalic = False
Text1.ForeColor = &H80000007
flag1 = False
End If
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) > 8 And flag2 Then
Text2.Text = ""
Text2.PasswordChar = "*"
Text2.FontItalic = False
Text2.ForeColor = &H80000007
flag2 = False
End If
End Sub
