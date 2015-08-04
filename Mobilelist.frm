VERSION 5.00
Begin VB.Form Mobilelist 
   BackColor       =   &H8000000E&
   Caption         =   "Mobiles"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   6120
      TabIndex        =   1
      Text            =   "Search any product"
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Skull Candy Headset"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   13
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Moto G (16 GB/8 GB)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   12
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Xolophone Q110"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   11
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Samsung S Duos 2"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   10
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Samsung Tab 3"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Image Image9 
      Height          =   2895
      Left            =   13560
      Picture         =   "Mobilelist.frx":0000
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Image Image8 
      Height          =   3015
      Left            =   9000
      Picture         =   "Mobilelist.frx":3252
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Image Image7 
      Height          =   2895
      Left            =   3840
      Picture         =   "Mobilelist.frx":3FE1
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   2895
      Left            =   13680
      Picture         =   "Mobilelist.frx":8075
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   2895
      Left            =   9000
      Picture         =   "Mobilelist.frx":EA31
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   3720
      Picture         =   "Mobilelist.frx":13C2B
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2775
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
      Left            =   15360
      TabIndex        =   8
      Top             =   0
      Width           =   1335
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
      Left            =   16680
      TabIndex        =   7
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
      Left            =   16680
      TabIndex        =   6
      Top             =   0
      Width           =   1455
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
      Left            =   15720
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   735
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
      Left            =   16440
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   960
      Picture         =   "Mobilelist.frx":1AA6E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   2640
      Picture         =   "Mobilelist.frx":1B5C8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2085
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   5520
      Picture         =   "Mobilelist.frx":1CE4E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   645
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
      Left            =   10320
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   12600
      Picture         =   "Mobilelist.frx":1D9A9
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1260
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
      Left            =   13800
      TabIndex        =   2
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
      Width           =   20295
   End
End
Attribute VB_Name = "Mobilelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.FontBold = False
Label18.FontBold = False
Label19.FontBold = False
Label17.FontUnderline = False
Label18.FontUnderline = False
Label19.FontUnderline = False
End Sub

Private Sub Label5_Click()
fr = Text1.Text
Call search
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
Private Sub Image1_Click()
Tab3.Show
Me.Hide
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image8.BorderStyle = 0
Image7.BorderStyle = 0
Image9.BorderStyle = 0
End Sub

Private Sub Image2_Click()
samsungduos2.Show
Me.Hide
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
Image2.BorderStyle = 1
Image3.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0
Image9.BorderStyle = 0
End Sub
Private Sub Image3_Click()
xolophone.Show
Me.Hide
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 1
Image7.BorderStyle = 0
Image8.BorderStyle = 0
Image9.BorderStyle = 0
End Sub
Private Sub Image7_Click()
Motox.Show
Me.Hide
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image7.BorderStyle = 1
Image8.BorderStyle = 0
Image9.BorderStyle = 0
End Sub
Private Sub Image8_Click()
Motog.Show
Me.Hide
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 1
Image9.BorderStyle = 0
End Sub
Private Sub Image9_Click()
Mobileacc1.Show
Me.Hide
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0
Image9.BorderStyle = 1
End Sub
