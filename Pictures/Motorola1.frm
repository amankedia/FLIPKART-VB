VERSION 5.00
Begin VB.Form Motorola1 
   BackColor       =   &H80000014&
   Caption         =   "Motorola"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9480
      Top             =   5280
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shop Now >"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Text            =   "Search any product"
      Top             =   480
      Width           =   4215
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   360
      Picture         =   "Motorola1.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "Explore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   495
      Left            =   15840
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "Explore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   615
      Left            =   15720
      TabIndex        =   5
      Top             =   10080
      Width           =   1455
   End
   Begin VB.Image Image7 
      Height          =   3375
      Left            =   14880
      Picture         =   "Motorola1.frx":0B5A
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   3375
      Left            =   14760
      Picture         =   "Motorola1.frx":4BEE
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   4815
      Left            =   2040
      Picture         =   "Motorola1.frx":7CD2
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   10575
   End
   Begin VB.Image Image4 
      Height          =   1575
      Left            =   1680
      Picture         =   "Motorola1.frx":19AD5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   7440
      Picture         =   "Motorola1.frx":1B35B
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   14040
      TabIndex        =   3
      Top             =   480
      Width           =   465
   End
   Begin VB.Image Image6 
      Height          =   555
      Left            =   12840
      Picture         =   "Motorola1.frx":1C230
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   4920
      Picture         =   "Motorola1.frx":1D6B7
      Stretch         =   -1  'True
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "Label4"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "Motorola1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontBold = False
Label1.FontSize = 12
Label2.FontBold = False
Label2.FontSize = 12
End Sub

Private Sub Image17_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Image3_Click()
Motog.Show
Me.Hide
End Sub

Private Sub Image4_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Image7_Click()
Motox.Show
Me.Hide
End Sub

Private Sub Label1_Click()
Motog.Show
Me.Hide
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontBold = True
Label1.FontSize = 14
End Sub

Private Sub Label2_Click()
Motox.Show
Me.Hide
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontBold = True
Label2.FontSize = 14
End Sub

Private Sub Timer1_Timer()
Command1.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub
Private Sub Text1_Click()
Text1.Text = ""
End Sub
