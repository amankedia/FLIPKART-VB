VERSION 5.00
Begin VB.Form Motog 
   BackColor       =   &H80000014&
   Caption         =   "Moto G"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
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
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9840
      Width           =   2055
   End
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9840
      Width           =   2055
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
      Left            =   6360
      TabIndex        =   1
      Text            =   "Search any product"
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SignIn"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   16560
      TabIndex        =   9
      Top             =   0
      Width           =   1335
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
      Left            =   14520
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   960
      Picture         =   "Motog.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs.13999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   5
      Top             =   9240
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs.12499"
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
      Left            =   8280
      TabIndex        =   4
      Top             =   9240
      Width           =   2055
   End
   Begin VB.Image Image7 
      Height          =   1695
      Left            =   12480
      Picture         =   "Motog.frx":0B5A
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   1695
      Left            =   6360
      Picture         =   "Motog.frx":2138
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Take Your Pick"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   8160
      TabIndex        =   3
      Top             =   8160
      Width           =   5415
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   13320
      Picture         =   "Motog.frx":3505
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
      Left            =   10560
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   5760
      Picture         =   "Motog.frx":498C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   2400
      Picture         =   "Motog.frx":54E7
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
   Begin VB.Image Image2 
      Height          =   3855
      Left            =   5040
      Picture         =   "Motog.frx":6D6D
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   12375
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   8280
      Picture         =   "Motog.frx":13AAE
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   5655
   End
End
Attribute VB_Name = "Motog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Dim b As Integer
Private Sub Command1_Click()
MsgBox "You have added MotoG(8GB) to your cart", vbOKOnly, "Product Bought"
b = InputBox("Enter the quanity", "Quantity")
Flipkart.Label6.Caption = Flipkart.Label6.Caption + b
Label6.Caption = Label6.Caption + b
Motorola1.Label6.Caption = Motorola1.Label6.Caption + b
xolophone.Label4.Caption = xolophone.Label4.Caption + b
Motox.Label6.Caption = Motox.Label6.Caption + b
End Sub

Private Sub Command2_Click()
MsgBox "You have added MotoG(16GB) to your cart", vbOKOnly, "Product Bought"
c = InputBox("Enter the quanity", "Quantity")
Flipkart.Label6.Caption = Flipkart.Label6.Caption + c
Label6.Caption = Label6.Caption + c
Motorola1.Label6.Caption = Motorola1.Label6.Caption + c
xolophone.Label4.Caption = xolophone.Label4.Caption + c
Motox.Label6.Caption = Motox.Label6.Caption + c
End Sub

Private Sub Image17_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Image4_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Label17_Click()
SignUp.Show
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.FontBold = True
Label17.FontUnderline = True
End Sub

Private Sub Text1_Change()
Text1.Text = ""
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub
