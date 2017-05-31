VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   10635
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Delete Record"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   5040
      Width           =   5535
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   7858
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rank"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Session"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dob"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Guardian's Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Guardian Ocuupation"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Contact No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Guardian No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Stream"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Discipline"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Sex"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Blood Group"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
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
Attribute VB_Name = "Form6"
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
Dim i, j As Integer
If Text1.Text = "" Then
    MsgBox "Enter name !"
Else
    For i = 1 To stdlen
        If std(i).Name = Text1.Text Then
            Exit For
        End If
    Next
    If i <= stdlen Then
        For j = i To stdlen - 1
            std(i) = std(i + 1)
        Next
        stdlen = stdlen - 1
        Call LoadList
        ReDim Preserve std(stdlen)
    Else
        MsgBox "Name not found!"
    End If
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub LoadList()
Dim list As ListItem
Dim i As Integer
ListView1.ListItems.Clear
For i = 1 To stdlen
    Set list = ListView1.ListItems.Add(, , std(i).Rank)
    list.SubItems(1) = std(i).Ses
    list.SubItems(2) = std(i).Name
    list.SubItems(3) = std(i).Dob
    list.SubItems(4) = std(i).GName
    list.SubItems(5) = std(i).GOcc
    list.SubItems(6) = std(i).Contact
    list.SubItems(7) = std(i).GContact
    list.SubItems(8) = std(i).Address
    list.SubItems(9) = std(i).Stream
    list.SubItems(10) = std(i).Disc
    list.SubItems(11) = std(i).Sex
    list.SubItems(12) = std(i).BGroup

Next

End Sub

Private Sub Exit_Click()
Unload Me
Me.Hide
End

End Sub

Private Sub Form_Load()


Call LoadList
Me.Top = 0
Me.Left = 0
Me.Height = 12000
Me.Width = 15000

Label1.Caption = "Enter name :"
Label1.FontBold = True
Label1.FontSize = 16

Text1.FontSize = 16

End Sub

Private Sub Home_Click()
Unload Me
Me.Hide
Load Form3
Form3.Show

End Sub

Private Sub Maximize_Click()
Form6.WindowState = 2

End Sub

Private Sub Minimize_Click()
Form6.WindowState = 1
End Sub

Private Sub Restore_Click()
Form6.WindowState = 0
End Sub
