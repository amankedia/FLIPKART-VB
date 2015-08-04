VERSION 5.00
Begin VB.Form Flipkart 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Welcome to Flipkart"
   ClientHeight    =   10095
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20400
   DrawMode        =   16  'Merge Pen
   FillColor       =   &H00C0FFC0&
   FillStyle       =   0  'Solid
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Shop Now >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   2880
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
      Left            =   5760
      TabIndex        =   6
      Text            =   "Search any product"
      Top             =   480
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Shop Now >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Shop Now >"
      DisabledPicture =   "Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Download the Flipkart App"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14520
      TabIndex        =   25
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Image Image12 
      BorderStyle     =   1  'Fixed Single
      Height          =   4095
      Left            =   14520
      Picture         =   "Form1.frx":14A1
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   3375
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
      Left            =   16560
      TabIndex        =   24
      Top             =   720
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
      Left            =   15840
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   735
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
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   16920
      TabIndex        =   21
      Top             =   0
      Width           =   1455
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
      Left            =   15600
      TabIndex        =   20
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Samsung Galaxy Tab 3(White) Rs.9999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   19
      Top             =   9720
      Width           =   2175
   End
   Begin VB.Image Image11 
      Height          =   1935
      Left            =   11520
      Picture         =   "Form1.frx":3AEE
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "              Samsung Galaxy S Duos 2          Rs.6350"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   18
      Top             =   9720
      Width           =   3015
   End
   Begin VB.Image Image10 
      Height          =   1935
      Left            =   8880
      Picture         =   "Form1.frx":4F2A
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "        Moto G Black(with 8 GB)       Rs. 4499"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   17
      Top             =   9720
      Width           =   2655
   End
   Begin VB.Image Image9 
      Height          =   1935
      Left            =   6000
      Picture         =   "Form1.frx":5FB1
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moto G Black(with 16 GB)    Rs.4999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   16
      Top             =   9720
      Width           =   2655
   End
   Begin VB.Image Image8 
      Height          =   1875
      Left            =   3480
      Picture         =   "Form1.frx":6D40
      Top             =   7800
      Width           =   960
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BestSellers"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pr6N M"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   7440
      Width           =   10575
   End
   Begin VB.Image Image3 
      Height          =   4035
      Left            =   2640
      Picture         =   "Form1.frx":7ACF
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   10575
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3525
      Left            =   14760
      Picture         =   "Form1.frx":1331E
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2880
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mobile Accessories"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   9960
      TabIndex        =   13
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sunglasses"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8160
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clothing"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stationery"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mobiles"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
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
      Left            =   13080
      TabIndex        =   8
      Top             =   480
      Width           =   465
   End
   Begin VB.Image Image6 
      Height          =   555
      Left            =   11880
      Picture         =   "Form1.frx":15543
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
      Left            =   9960
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   5160
      Picture         =   "Form1.frx":169CA
      Stretch         =   -1  'True
      Top             =   480
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   2640
      Picture         =   "Form1.frx":17525
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "Label4"
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   20415
   End
   Begin VB.Image Image1 
      Height          =   4020
      Left            =   2640
      Picture         =   "Form1.frx":18DAB
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   10575
   End
   Begin VB.Image Image2 
      Height          =   3945
      Left            =   2640
      Picture         =   "Form1.frx":25213
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "              Staple Easy                      With Kangaroo Staplers"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   9840
      TabIndex        =   2
      Top             =   6240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "            Summer Sale                        Minimum 40% Off"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   6360
      TabIndex        =   1
      Top             =   6240
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                Introducing                                  XOLO Phones"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   855
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   6240
      Width           =   3735
   End
End
Attribute VB_Name = "Flipkart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim a As String

Private Sub Command1_Click()
Kangaroo.Show
Me.Hide
End Sub

Private Sub Command2_Click()
Clothes.Show
Me.Hide
End Sub

Private Sub Command3_Click()
xolophone.Show
Me.Hide
End Sub
Private Sub Form_Activate()
If login1 = 1 Then
Label18.Visible = False
Label19.Visible = True
End If
Text1.Text = "Seach any product"
End Sub
Private Sub search()
Module1.GetConnected
Dim res As New ADODB.Recordset
Dim flag As Integer
flag = 0
res.Open "select* from Products where ProductName = '" & fr & "'", cnn, adOpenStatic, adLockReadOnly
If res.RecordCount < 1 Then
MsgBox "Product Invalid", vbInformation, "Search"
Flipkart.Show
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
Flipkart.Show
Exit Sub
End If
End If
Set res = Nothing
End Sub
Private Sub Form_Load()
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Command2.Visible = True
Command3.Visible = False
Command1.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label7.BackColor = &H80000014
 Label8.BackColor = &H80000014
 Label9.BackColor = &H80000014
 Label10.BackColor = &H80000014
 Label11.BackColor = &H80000014
 Label16.FontBold = False
 Label16.FontUnderline = False
  Label15.FontBold = False
 Label15.FontUnderline = False
  Label14.FontBold = False
 Label14.FontUnderline = False
  Label13.FontBold = False
 Label13.FontUnderline = False
 Label12.FontBold = False
 Label17.FontBold = False
Label17.FontUnderline = False
Label18.FontBold = False
Label18.FontUnderline = False
Label19.FontBold = False
Label19.FontUnderline = False
 End Sub

Private Sub Image1_Click()
Clothes.Show
Me.Hide
End Sub

Private Sub Image12_Click()
Me.Show
End Sub

Private Sub Image2_Click()
xolophone.Show
Me.Hide
End Sub

Private Sub Image3_Click()
Kangaroo.Show
Me.Hide
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.FontBold = False
Label15.FontBold = False
Label15.FontBold = False
Label13.FontBold = False
Label7.BackColor = &H80000014
Label8.BackColor = &H80000014
Label9.BackColor = &H80000014
Label10.BackColor = &H80000014
Label11.BackColor = &H80000014
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.FontBold = False
Label15.FontBold = False
Label15.FontBold = False
Label13.FontBold = False
Label7.BackColor = &H80000014
Label8.BackColor = &H80000014
Label9.BackColor = &H80000014
Label10.BackColor = &H80000014
Label11.BackColor = &H80000014
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.FontBold = False
Label15.FontBold = False
Label15.FontBold = False
Label13.FontBold = False
Label7.BackColor = &H80000014
Label8.BackColor = &H80000014
Label9.BackColor = &H80000014
Label10.BackColor = &H80000014
Label11.BackColor = &H80000014
End Sub


Private Sub Image6_Click()
Cart.Show
End Sub

Private Sub Image7_Click()
Motorola1.Show
Me.Hide
End Sub

Private Sub Image8_Click()
Motorola1.Show
Me.Hide
End Sub

Private Sub Image9_Click()
Motorola1.Show
Me.Hide
End Sub

Private Sub Label1_Click(Index As Integer)
xolophone.Show
Me.Hide
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image1.Visible = False
Image3.Visible = False
Command3.Visible = True
Command2.Visible = False
Command1.Visible = False
Label1(0).BackColor = &H80000014
Label1(0).ForeColor = &H80000012
Label2.ForeColor = &H80000014
Label3.ForeColor = &H80000014
Label2.BackColor = &H80000012
Label3.BackColor = &H80000012
End Sub
Private Sub HighlightLabel(blnHighlight As Boolean)
  Label7.FontBold = blnHighlight
  Label7.FontSize = 18
  Label7.BackColor = &HC0C0FF
  Label8.BackColor = &H80000014
  Label9.BackColor = &H80000014
  Label10.BackColor = &H80000014
  Label11.BackColor = &H80000014
  Label8.FontBold = False
  Label9.FontBold = False
  Label10.FontBold = False
  Label11.FontBold = False
  Label13.FontBold = False
  Label14.FontBold = False
  Label12.FontBold = False
  Label15.FontBold = False
  Label16.FontBold = False
  End Sub
  Private Sub HighlightLabelA(blnHighlight As Boolean)
  Label8.FontBold = blnHighlight
  Label8.FontSize = 18
  Label8.BackColor = &HC0C0FF
  Label7.BackColor = &H80000014
  Label9.BackColor = &H80000014
  Label10.BackColor = &H80000014
  Label11.BackColor = &H80000014
  Label7.FontBold = False
  Label9.FontBold = False
  Label10.FontBold = False
  Label11.FontBold = False
  Label12.FontBold = False
  Label13.FontBold = False
  Label14.FontBold = False
  Label16.FontBold = False
  Label15.FontBold = False
  End Sub
  
Private Sub Label10_Click()
Sunlist.Show
Me.Hide
End Sub

Private Sub Label11_Click()
Mobileacc1.Show
Me.Hide
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.FontBold = True
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.FontBold = True
Label13.FontUnderline = True
Label14.FontBold = False
Label14.FontUnderline = False
Label15.FontBold = False
Label15.FontUnderline = False
Label16.FontBold = False
Label16.FontUnderline = False
End Sub

Private Sub Label14_Click()
Motorola1.Show
Me.Hide
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.FontBold = True
Label14.FontUnderline = True
Label13.FontBold = False
Label13.FontUnderline = False
Label15.FontBold = False
Label15.FontUnderline = False
Label16.FontBold = False
Label16.FontUnderline = False
End Sub

Private Sub Label15_Click()
samsungduos2.Show
Me.Hide
End Sub

Private Sub Label16_Click()
Tab3.Show
Me.Hide
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.FontBold = True
Label16.FontUnderline = True
Label13.FontBold = False
Label13.FontUnderline = False
Label14.FontBold = False
Label14.FontUnderline = False
Label15.FontBold = False
Label15.FontUnderline = False
End Sub
Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.FontBold = True
Label15.FontUnderline = True
Label13.FontBold = False
Label13.FontUnderline = False
Label14.FontBold = False
Label14.FontUnderline = False
Label16.FontBold = False
Label16.FontUnderline = False
End Sub
Private Sub Label13_Click()
Motorola1.Show
Me.Hide
End Sub


Private Sub Label17_Click()
SignUp.Show
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.FontBold = True
Label17.FontUnderline = True
Label18.FontBold = False
Label18.FontUnderline = False
Label19.FontBold = False
Label19.FontUnderline = False
End Sub

Private Sub Label18_Click()
login.Show
Me.Hide
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.FontBold = True
Label18.FontUnderline = True
Label17.FontBold = False
Label17.FontUnderline = False
End Sub

Private Sub Label19_Click()
lg = MsgBox("Are you sure you want to logout?", vbYesNo, "Logout")
If lg = 6 Then
Logout.Show
Me.Hide
End If
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.FontBold = True
Label19.FontUnderline = True
End Sub

Private Sub Label2_Click()
Clothes.Show
Me.Hide
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.Visible = True
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Command3.Visible = False
Command1.Visible = False
Label2.BackColor = &H80000014
Label2.ForeColor = &H80000012
Label1(0).BackColor = &H80000012
Label3.BackColor = &H80000012
Label1(0).ForeColor = &H80000014
Label3.ForeColor = &H80000014
End Sub

Private Sub Label3_Click()
Kangaroo.Show
Me.Hide
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image1.Visible = False
Image2.Visible = False
Command1.Visible = True
Command3.Visible = False
Command2.Visible = False
Label3.BackColor = &H80000014
Label3.ForeColor = &H80000012
Label1(0).BackColor = &H80000012
Label2.BackColor = &H80000012
Label2.ForeColor = &H80000014
Label1(0).ForeColor = &H80000014
End Sub



Private Sub Label5_Click()
fr = Text1.Text
Call search
End Sub


Private Sub Label7_Click()
Mobilelist.Show
Me.Hide
End Sub

Private Sub Label8_Click()
Kangaroo.Show
Me.Hide
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightLabelA True
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightLabel True
End Sub


Private Sub Label9_Click()
Clothes.Show
Me.Hide
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub
Private Sub HighlightLabelB(blnHighlight As Boolean)
  Label9.FontBold = blnHighlight
  Label9.FontSize = 18
  Label9.BackColor = &HC0C0FF
  Label7.BackColor = &H80000014
  Label8.BackColor = &H80000014
  Label10.BackColor = &H80000014
  Label11.BackColor = &H80000014
  Label8.FontBold = False
  Label7.FontBold = False
  Label10.FontBold = False
  Label11.FontBold = False
  Label12.FontBold = False
  Label13.FontBold = False
  Label14.FontBold = False
  Label15.FontBold = False
  Label16.FontBold = False
  End Sub
  Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightLabelB True
End Sub

Private Sub HighlightLabelC(blnHighlight As Boolean)
  Label10.FontBold = blnHighlight
  Label10.FontSize = 18
  Label10.BackColor = &HC0C0FF
  Label8.BackColor = &H80000014
  Label9.BackColor = &H80000014
  Label7.BackColor = &H80000014
  Label11.BackColor = &H80000014
 Label16.FontBold = False
  Label8.FontBold = False
  Label9.FontBold = False
  Label15.FontBold = False
  Label7.FontBold = False
  Label11.FontBold = False
  Label12.FontBold = False
Label13.FontBold = False
  Label14.FontBold = False
  End Sub
  Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightLabelC True
End Sub
Private Sub HighlightLabelD(blnHighlight As Boolean)
  Label11.FontBold = blnHighlight
  Label11.FontSize = 18
  Label11.BackColor = &HC0C0FF
  Label15.FontBold = False
  Label16.FontBold = False
  Label8.BackColor = &H80000014
  Label9.BackColor = &H80000014
  Label7.BackColor = &H80000014
  Label10.BackColor = &H80000014
  Label8.FontBold = False
  Label9.FontBold = False
  Label7.FontBold = False
  Label10.FontBold = False
  Label12.FontBold = False
  Label13.FontBold = False
  Label14.FontBold = False
  End Sub
  Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightLabelD True
End Sub
Private Sub Timer1_Timer()
Label8.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label7.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label9.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label10.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label11.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Command1.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Command2.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Command3.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label22.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label13.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label14.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label15.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label16.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label12.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub


