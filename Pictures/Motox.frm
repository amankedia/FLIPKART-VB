VERSION 5.00
Begin VB.Form Motox 
   BackColor       =   &H80000014&
   Caption         =   "Moto X"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Buy Now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8880
      Width           =   2295
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
      Left            =   6960
      TabIndex        =   1
      Text            =   "Search any product"
      Top             =   480
      Width           =   4215
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
      Height          =   495
      Left            =   15120
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   840
      Picture         =   "Motox.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs.23999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Available Exclusively at Flipkart"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   7920
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Height          =   735
      Left            =   11400
      TabIndex        =   3
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   3975
      Left            =   5280
      Picture         =   "Motox.frx":0B5A
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   10695
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   8280
      Picture         =   "Motox.frx":1295D
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   13920
      Picture         =   "Motox.frx":17124
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
      Left            =   11160
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   6360
      Picture         =   "Motox.frx":185AB
      Stretch         =   -1  'True
      Top             =   480
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   2520
      Picture         =   "Motox.frx":19106
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "Label4"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "Motox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As Integer

Private Sub Command1_Click()
MsgBox "You have added MotoX to your cart", vbOKOnly, "Product Bought"
d = InputBox("Enter the quanity", "Quantity")
Flipkart.Label6.Caption = Flipkart.Label6.Caption + d
Motog.Label6.Caption = Motog.Label6.Caption + d
Motorola1.Label6.Caption = Motorola1.Label6.Caption + d
xolophone.Label4.Caption = xolophone.Label4.Caption + d
Label6.Caption = Label6.Caption + d
End Sub

Private Sub Image17_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Image4_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub
