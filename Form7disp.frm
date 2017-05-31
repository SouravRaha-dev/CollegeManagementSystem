VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H8000000B&
   Caption         =   "Form7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label27 
      Height          =   735
      Left            =   3000
      TabIndex        =   26
      Top             =   6000
      Width           =   11295
   End
   Begin VB.Label Label26 
      Height          =   495
      Left            =   1320
      TabIndex        =   25
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label25 
      Height          =   495
      Left            =   10920
      TabIndex        =   24
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label24 
      Height          =   495
      Left            =   8520
      TabIndex        =   23
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label23 
      Height          =   495
      Left            =   3000
      TabIndex        =   22
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label22 
      Height          =   495
      Left            =   1320
      TabIndex        =   21
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label21 
      Height          =   495
      Left            =   10920
      TabIndex        =   20
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label20 
      Height          =   615
      Left            =   7440
      TabIndex        =   19
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label Label19 
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Label Label18 
      Height          =   615
      Left            =   -120
      TabIndex        =   17
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label17 
      Height          =   495
      Left            =   10920
      TabIndex        =   16
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label16 
      Height          =   495
      Left            =   8400
      TabIndex        =   15
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label15 
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label14 
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label13 
      Height          =   495
      Left            =   10920
      TabIndex        =   12
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label12 
      Height          =   495
      Left            =   7800
      TabIndex        =   11
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label11 
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label10 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label9 
      Height          =   495
      Left            =   10920
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label8 
      Height          =   495
      Left            =   9120
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label6 
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   10920
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Details Saved!!!!!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 12000
Me.Width = 15000
Form5.BackColor = &H80FFFF
Label2.Caption = "RANK:"
Label2.FontBold = True
Label2.FontSize = 12
Label2.BackColor = &HC0E0FF
Label3.Caption = std(stdlen).Rank
Label3.FontBold = True
Label3.FontSize = 12
Label3.BackColor = &HC0E0FF
Label4.Caption = "SESSION:"
Label4.FontBold = True
Label4.FontSize = 12
Label4.BackColor = &HC0E0FF
Label5.Caption = std(stdlen).Ses
Label5.FontBold = True
Label5.FontSize = 12
Label5.BackColor = &HC0E0FF
Label6.Caption = "NAME:"
Label6.FontBold = True
Label6.FontSize = 12
Label6.BackColor = &HC0E0FF
Label7.Caption = std(stdlen).Name
Label7.FontBold = True
Label7.FontSize = 12
Label7.BackColor = &HC0E0FF
Label8.Caption = "DOB:"
Label8.FontBold = True
Label8.FontSize = 12
Label8.BackColor = &HC0E0FF
Label9.Caption = std(stdlen).Dob
Label9.FontBold = True
Label9.FontSize = 12
Label9.BackColor = &HC0E0FF
Label10.Caption = "GUARDIAN'S NAME:"
Label10.FontBold = True
Label10.FontSize = 12
Label10.BackColor = &HC0E0FF
Label11.Caption = std(stdlen).GName
Label11.FontBold = True
Label11.FontSize = 12
Label11.BackColor = &HC0E0FF
Label12.Caption = "GUARDIAN'S OCC.:"
Label12.FontBold = True
Label12.FontSize = 12
Label12.BackColor = &HC0E0FF
Label13.Caption = std(stdlen).GOcc
Label13.FontBold = True
Label13.FontSize = 12
Label13.BackColor = &HC0E0FF
Label14.Caption = "STREAM:"
Label14.FontBold = True
Label14.FontSize = 12
Label14.BackColor = &HC0E0FF
Label15.Caption = std(stdlen).Stream
Label15.FontBold = True
Label15.FontSize = 12
Label15.BackColor = &HC0E0FF
Label16.Caption = "DISC.:"
Label16.FontBold = True
Label16.FontSize = 12
Label16.BackColor = &HC0E0FF
Label17.Caption = std(stdlen).Disc
Label17.FontBold = True
Label17.FontSize = 12
Label17.BackColor = &HC0E0FF
Label18.Caption = "STUDENT'S CONTACT NO.:"
Label18.FontBold = True
Label18.FontSize = 12
Label18.BackColor = &HC0E0FF
Label19.Caption = std(stdlen).Contact
Label19.FontBold = True
Label19.FontSize = 12
Label19.BackColor = &HC0E0FF
Label20.Caption = "GUARDIAN'S CONTACT NO.:"
Label20.FontBold = True
Label20.FontSize = 12
Label20.BackColor = &HC0E0FF
Label21.Caption = std(stdlen).GContact
Label21.FontBold = True
Label21.FontSize = 12
Label21.BackColor = &HC0E0FF
Label22.Caption = "SEX:"
Label22.FontBold = True
Label22.FontSize = 12
Label22.BackColor = &HC0E0FF
Label23.Caption = std(stdlen).Sex
Label23.FontBold = True
Label23.FontSize = 12
Label23.BackColor = &HC0E0FF
Label24.Caption = "BLOOD GROUP:"
Label24.FontBold = True
Label24.FontSize = 12
Label24.BackColor = &HC0E0FF
Label25.Caption = std(stdlen).BGroup
Label25.FontBold = True
Label25.FontSize = 12
Label25.BackColor = &HC0E0FF
Label26.Caption = "ADDRESS:"
Label26.FontBold = True
Label26.FontSize = 12
Label26.BackColor = &HC0E0FF
Label27.Caption = std(stdlen).Address
Label27.FontBold = True
Label27.FontSize = 12
Label27.BackColor = &HC0E0FF
End Sub

