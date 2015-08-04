VERSION 5.00
Begin VB.Form xolophone 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Myriad Pro"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Buy Now"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7920
      Width           =   2535
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   600
      Picture         =   "xolophone.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash On Delivery"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   16
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Image Image16 
      Height          =   375
      Left            =   10080
      Picture         =   "xolophone.frx":0B5A
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "30 Day Replacement Guarantee"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   15
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Image Image15 
      Height          =   375
      Left            =   10080
      Picture         =   "xolophone.frx":191E
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Free Home Delivery"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Easy EMI AT 1059"
      BeginProperty Font 
         Name            =   "Kozuka Mincho Pro EL"
         Size            =   12
         Charset         =   0
         Weight          =   250
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   12
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Inclusive of Taxes"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs.11997"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   10
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   ">5-Inch HD IPS OGS Dislpay"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   9
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   ">Android Jelly Bean 4.3"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   ">Qualcomm Snapdragon 400"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   7
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   ">1.4GHz Quad Core Processor"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Image Image14 
      Height          =   975
      Left            =   13320
      Picture         =   "xolophone.frx":26E2
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   975
   End
   Begin VB.Image Image13 
      Height          =   615
      Left            =   11760
      Picture         =   "xolophone.frx":351A
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image Image12 
      Height          =   375
      Left            =   10200
      Picture         =   "xolophone.frx":3AA3
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(Black)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "XOLO Q100"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Image Image11 
      Height          =   1935
      Left            =   5040
      Picture         =   "xolophone.frx":4F44
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image Image10 
      Height          =   1935
      Left            =   4200
      Picture         =   "xolophone.frx":5118
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image Image9 
      Height          =   1935
      Left            =   2400
      Picture         =   "xolophone.frx":52F2
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Image Image8 
      Height          =   1935
      Left            =   360
      Picture         =   "xolophone.frx":5C0A
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Image Image7 
      Height          =   5640
      Left            =   3120
      Picture         =   "xolophone.frx":69A2
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   450
   End
   Begin VB.Image Image6 
      Height          =   5760
      Left            =   3120
      Picture         =   "xolophone.frx":748A
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   450
   End
   Begin VB.Image Image5 
      Height          =   5760
      Left            =   1800
      Picture         =   "xolophone.frx":7F02
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2820
   End
   Begin VB.Image Image4 
      Height          =   5760
      Left            =   1800
      Picture         =   "xolophone.frx":C820
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2850
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      Left            =   13680
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   12840
      Picture         =   "xolophone.frx":131DC
      Stretch         =   -1  'True
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Search"
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
      Left            =   10320
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Search any product"
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
      Left            =   6120
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   5400
      Picture         =   "xolophone.frx":14663
      Stretch         =   -1  'True
      Top             =   480
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   2520
      Picture         =   "xolophone.frx":151BE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2205
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "xolophone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Command1_Click()
MsgBox "You have added Xolo Q110 to your cart", vbOKOnly, "Product Bought"
a = InputBox("Enter the quanity", "Quantity")
Flipkart.Label6.Caption = Flipkart.Label6.Caption + a
Label4.Caption = Label4.Caption + a
Motorola1.Label6.Caption = Motorola1.Label6.Caption + a
Motog.Label6.Caption = Motog.Label6.Caption + a
Motox.Label6.Caption = Motox.Label6.Caption + a
End Sub

Private Sub Form_Load()
Image4.Visible = True
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
End Sub

Private Sub Image1_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
Image4.Visible = False
Image5.Visible = False
Image7.Visible = False
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image4.Visible = False
Image6.Visible = False
Image5.Visible = False
End Sub

Private Sub Image17_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
End Sub
Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = True
Image4.Visible = False
Image6.Visible = False
Image7.Visible = False
End Sub


Private Sub Label2_Click()
Text1.Text = ""
End Sub

