VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   6825
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14640
   FillColor       =   &H00C0C000&
   ForeColor       =   &H00C0C000&
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   3840
      TabIndex        =   27
      Top             =   4320
      Width           =   9735
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   10800
      TabIndex        =   26
      Top             =   3840
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3840
      TabIndex        =   25
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   10800
      TabIndex        =   24
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   23
      Top             =   3240
      Width           =   3975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   10800
      TabIndex        =   22
      Top             =   2760
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3840
      TabIndex        =   21
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   10800
      TabIndex        =   20
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   19
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   10800
      TabIndex        =   18
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   17
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   10800
      TabIndex        =   16
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   15
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "&Submit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Picture         =   "Form5regt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      Height          =   4455
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   13095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   735
      Left            =   1080
      Top             =   720
      Width           =   13095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Group :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   9120
      TabIndex        =   11
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian's Contact No. :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Student's Contact No. :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Discipline :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   9480
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Stream :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian's Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian's Occupation :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   9120
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Session :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "WBJEE/JEE Rank :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Admission"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   120
      Width           =   2895
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
Attribute VB_Name = "Form5"
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
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Then
    MsgBox ("Please fill up the vacant space(s) !")
    Else
    ReDim Preserve std(stdlen + 1)
    stdlen = stdlen + 1
    std(stdlen).Rank = Text1.Text
    std(stdlen).Ses = Text2.Text
    std(stdlen).Name = Text3.Text
    std(stdlen).Dob = Text4.Text
    std(stdlen).GName = Text5.Text
    std(stdlen).GOcc = Text6.Text
    std(stdlen).Contact = Text7.Text
    std(stdlen).GContact = Text8.Text
    std(stdlen).Address = Text9.Text
    std(stdlen).Stream = Combo1.Text
    std(stdlen).Disc = Combo2.Text
    std(stdlen).Sex = Combo3.Text
    std(stdlen).BGroup = Combo4.Text
Form7.Show
Unload Me
Me.Hide

End If


End Sub

Private Sub Exit_Click()
Unload Me
Me.Hide
End

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 12000
Me.Width = 15000
Combo1.AddItem ("B.Tech")
Combo1.AddItem ("M.Tech")
Combo1.AddItem ("BCA")
Combo1.AddItem ("MCA")
Combo1.AddItem ("BBA")
Combo1.AddItem ("MBA")
Combo1.FontSize = 10

Combo2.AddItem ("CSE")
Combo2.AddItem ("IT")
Combo2.AddItem ("ECE")
Combo2.AddItem ("EE")
Combo2.AddItem ("CE")
Combo2.AddItem ("ME")
Combo2.AddItem ("GeoTech")
Combo2.FontSize = 10

Combo3.AddItem ("Male")
Combo3.AddItem ("Female")
Combo3.AddItem ("Transgender")
Combo3.FontSize = 10

Combo4.AddItem ("O+")
Combo4.AddItem ("O-")
Combo4.AddItem ("A+")
Combo4.AddItem ("A-")
Combo4.AddItem ("B+")
Combo4.AddItem ("B-")
Combo4.AddItem ("AB-")
Combo4.FontSize = 10

Text1.FontSize = 10
Text2.FontSize = 10
Text3.FontSize = 10
Text4.FontSize = 10
Text5.FontSize = 10
Text6.FontSize = 10
Text7.FontSize = 10
Text8.FontSize = 10
Text9.FontSize = 10
End Sub

Private Sub Home_Click()
Unload Me
Me.Hide
Load Form3
Form3.Show

End Sub

Private Sub Maximize_Click()
Form5.WindowState = 2

End Sub

Private Sub Minimize_Click()
Form5.WindowState = 1

End Sub

Private Sub Restore_Click()
Form5.WindowState = 0
End Sub

Private Sub Text1_Change()
If IsNumeric(Text1.Text) Then
Else
Text1.Text = ""
End If
End Sub

Private Sub Text7_Change()
If IsNumeric(Text7.Text) Then
Else
Text7.Text = ""
End If

End Sub

Private Sub Text8_Change()
If IsNumeric(Text8.Text) Then
Else
Text8.Text = ""
End If

End Sub
