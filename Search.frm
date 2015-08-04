VERSION 5.00
Begin VB.Form Sunglasses 
   BackColor       =   &H8000000E&
   Caption         =   "Fastrack Aviator"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5220
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
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
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7920
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
      Left            =   6600
      TabIndex        =   1
      Text            =   "Search any product"
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   13320
      TabIndex        =   17
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Availlable in Black and"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   16
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Image Image13 
      Height          =   375
      Left            =   14520
      Picture         =   "Search.frx":0000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "+ Rs.60 Delivery Charges"
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
      Left            =   9360
      TabIndex        =   15
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Image Image15 
      Height          =   375
      Left            =   14520
      Picture         =   "Search.frx":0DC4
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   375
   End
   Begin VB.Image Image12 
      Height          =   615
      Left            =   14880
      Picture         =   "Search.frx":1B88
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Image Image16 
      Height          =   975
      Left            =   18360
      Picture         =   "Search.frx":3A06
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   975
   End
   Begin VB.Image Image14 
      Height          =   615
      Left            =   16800
      Picture         =   "Search.frx":483E
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash On Delivery available"
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
      Left            =   15240
      TabIndex        =   14
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "30 Day Replacement Policy"
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
      Left            =   15240
      TabIndex        =   13
      Top             =   5760
      Width           =   3735
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
      Left            =   9480
      TabIndex        =   12
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs.799"
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
      Left            =   9360
      TabIndex        =   11
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fastrack Aviator Sunglasses"
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
      Left            =   8640
      TabIndex        =   10
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Image Image11 
      Height          =   2175
      Left            =   6360
      Picture         =   "Search.frx":4DC7
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Image Image10 
      Height          =   2175
      Left            =   4680
      Picture         =   "Search.frx":580C
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Image Image9 
      Height          =   2175
      Left            =   2760
      Picture         =   "Search.frx":614E
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Image Image8 
      Height          =   2175
      Left            =   960
      Picture         =   "Search.frx":6B15
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Image Image7 
      Height          =   4575
      Left            =   2160
      Picture         =   "Search.frx":73D7
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5295
   End
   Begin VB.Image Image3 
      Height          =   4695
      Left            =   2040
      Picture         =   "Search.frx":C3B4
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   4575
      Left            =   2040
      Picture         =   "Search.frx":100EB
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   2040
      Picture         =   "Search.frx":13C7D
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   5055
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
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
      Height          =   615
      Left            =   16920
      TabIndex        =   8
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hi!"
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
      Height          =   495
      Left            =   16080
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
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
      Height          =   495
      Left            =   17160
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "LogOut"
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
      Left            =   17280
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Signup"
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
      Left            =   15840
      TabIndex        =   4
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
      Left            =   14400
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   13200
      Picture         =   "Search.frx":1738B
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
      Left            =   10800
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   6000
      Picture         =   "Search.frx":18812
      Stretch         =   -1  'True
      Top             =   480
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   3000
      Picture         =   "Search.frx":1936D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2085
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   1320
      Picture         =   "Search.frx":1ABF3
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "Label4"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20730
   End
End
Attribute VB_Name = "Sunglasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Integer
Dim c As Integer
Private Sub Command1_Click()
If login1 = 0 Then
MsgBox "You need to log in first", vbInformation, "Login"
Else
If i = 10 Then
MsgBox "You can't buy more items in one login", vbInformation, "Ophs"
Cart.Show
Me.Hide
Else
f = MsgBox("You have added Fastrack Aviator sunglasses to your cart", vbOKCancel, "Product Bought")
If f = 1 Then
c = InputBox("Enter the quanity", "Quantity")
If c = 0 Then
MsgBox "Add Proper number of items" & vbCr & "Click to buy again", vbCritical, "Check Your Selection"
Else
Flipkart.Label6.Caption = Flipkart.Label6.Caption + c
Label6.Caption = Label6.Caption + c
Motorola1.Label6.Caption = Motorola1.Label6.Caption + c
xolophone.Label4.Caption = xolophone.Label4.Caption + c
Motox.Label6.Caption = Motox.Label6.Caption + c
Motog.Label6.Caption = Motog.Label6.Caption + c
Kangaroo.Label6.Caption = Kangaroo.Label6.Caption + c
samsungduos2.Label6.Caption = samsungduos2.Label6.Caption + c
Tab3.Label6.Caption = Tab3.Label6.Caption + c
Cloth1.Label6.Caption = Cloth1.Label6.Caption + c
Mobileacc1.Label6.Caption = Mobileacc1.Label6.Caption + c
Mobilelist.Label6.Caption = Mobilelist.Label6.Caption + c
Cloth2.Label6.Caption = Cloth2.Label6.Caption + c
Clothes.Label6.Caption = Clothes.Label6.Caption + c
Sunglasses2.Label6.Caption = Sunglasses2.Label6.Caption + c
Sunlist.Label6.Caption = Sunlist.Label6.Caption + c
ct(i) = c
prodval(i) = 799
ce(i, i) = "Fastrack Aviator Sunglasses"
amount = amount + (c * 799)
i = i + 1
End If
End If
End If
End If
End Sub
Private Sub search()
Module1.GetConnected
Dim res As New ADODB.Recordset
Dim flag As Integer
flag = 0
res.Open "select* from Products where ProductName = '" & fr & "'", cnn, adOpenStatic, adLockReadOnly
If res.RecordCount < 1 Then
MsgBox "Product Invalid", vbInformation, "Search"
Me.Show
Exit Sub
Else
If fr = res!ProductName Then
MsgBox "PRODUCT FOUND", vbInformation, "Results"
flag = res.Fields(3)
Select Case flag
Case Is = 1
xolophone.Show
Me.Hide
Case Is = 2
Motox.Show
Me.Hide
Case Is = 3
Motog.Show
Me.Hide
Case Is = 4
Motorola1.Show
Me.Hide
Case Is = 5
Sunlist.Show
Me.Hide
Case Is = 6
Sunglasses.Show
Me.Hide
Case Is = 7
Sunglasses2.Show
Me.Hide
Case Is = 8
Kangaroo.Show
Me.Hide
Case Is = 9
Cloth1.Show
Me.Hide
Case Is = 10
Cloth2.Show
Me.Hide
Case Is = 11
Clothes.Show
Me.Hide
Case Is = 12
Mobilelist.Show
Me.Hide
Case Is = 13
Tab3.Show
Me.Hide
Case Is = 14
samsungduos2.Show
Me.Hide
Case Is = 15
Mobileacc1.Show
Me.Hide
End Select
Exit Sub
Else
MsgBox "Product Not found", vbInformation, "Search"
Me.Show
Exit Sub
End If
End If
Set res = Nothing
End Sub
Private Sub Form_Activate()
If login1 = 1 Then
Label18.Visible = False
Label19.Visible = True
End If
Text1.Text = "Seach any product"
End Sub

Private Sub Form_Load()
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Image7.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.FontBold = False
Label17.FontUnderline = False
Label18.FontBold = False
Label18.FontUnderline = False
Label19.FontBold = False
Label19.FontUnderline = False
End Sub

Private Sub Image16_Click()
Review.Show
Me.Hide
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Image7.Visible = False
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image1.Visible = False
Image3.Visible = False
Image7.Visible = False
End Sub
Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image1.Visible = False
Image2.Visible = False
Image7.Visible = False
End Sub
Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image1.Visible = False
Image3.Visible = False
Image2.Visible = False
End Sub

Private Sub Label18_Click()
login.Show
Me.Hide
End Sub

Private Sub Label19_Click()
lg = MsgBox("Are you sure you want to logout?", vbYesNo, "Logout")
If lg = 6 Then
Logout.Show
Me.Hide
End If
End Sub

Private Sub Label5_Click()
fr = Text1.Text
Call search
End Sub


Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Image17_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Image4_Click()
Flipkart.Show
Me.Hide
End Sub
Private Sub Image6_Click()
Cart.Show
End Sub
Private Sub Label17_Click()
SignUp.Show
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.FontBold = True
Label17.FontUnderline = True
Label18.FontBold = False
Label18.FontUnderline = False
End Sub
Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.FontBold = True
Label18.FontUnderline = True
Label17.FontBold = False
Label17.FontUnderline = False
End Sub
Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.FontBold = True
Label19.FontUnderline = True
End Sub



