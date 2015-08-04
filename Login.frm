VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H80000014&
   Caption         =   "Log In"
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
      BackColor       =   &H008080FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   3135
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   7
      Text            =   "Search any product"
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   11760
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   5160
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      TabIndex        =   4
      Top             =   4200
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      TabIndex        =   3
      Top             =   3240
      Width           =   4335
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
      Left            =   17760
      TabIndex        =   14
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
      Left            =   16320
      TabIndex        =   13
      Top             =   0
      Width           =   1335
   End
   Begin VB.Image Image17 
      Height          =   495
      Left            =   840
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't have an account?SignUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   7800
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot your password?"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12960
      TabIndex        =   11
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14880
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   555
      Left            =   13680
      Picture         =   "Login.frx":0B5A
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
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   6000
      Picture         =   "Login.frx":1FE1
      Stretch         =   -1  'True
      Top             =   480
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   2760
      Picture         =   "Login.frx":2B3C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "Label4"
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   20295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID:"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   6720
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Name Empty"
Text1.SetFocus
Exit Sub
ElseIf Text3.Text = "" Then
MsgBox "Password Empty"
Text3.SetFocus
Exit Sub
ElseIf Text2.Text = "" Then
MsgBox "EmailID Empty"
Text2.SetFocus
Exit Sub
Else
Call login
End If
End Sub

Private Sub login()
Module1.GetConnected
Dim rs As New ADODB.Recordset
rs.Open "select* from Users where Name = '" & Text1.Text & "'", cnn, adOpenStatic, adLockReadOnly

If rs.RecordCount < 1 Then
MsgBox "UserName is Invalid", vbInformation, "login"
Text1.SetFocus
Exit Sub
Else
If Text3.Text = rs!Password Then
STR1 = Text1.Text
MsgBox "You have successfully logged in", vbInformation, "Login Successful"
Flipkart.Label18.Visible = False
Flipkart.Label17.Visible = False
Flipkart.Label19.Visible = True
Flipkart.Label20.Visible = True
Flipkart.Label21.Caption = STR1
Motog.Label18.Visible = False
Motog.Label17.Visible = False
Motog.Label19.Visible = True
Motog.Label20.Visible = True
Motog.Label21.Caption = STR1
Motox.Label18.Visible = False
Motox.Label17.Visible = False
Motox.Label19.Visible = True
Motox.Label20.Visible = True
Motox.Label21.Caption = STR1
Motorola1.Label18.Visible = False
Motorola1.Label17.Visible = False
Motorola1.Label19.Visible = True
Motorola1.Label20.Visible = True
Motorola1.Label21.Caption = STR1
xolophone.Label18.Visible = False
xolophone.Label17.Visible = False
xolophone.Label19.Visible = True
xolophone.Label20.Visible = True
xolophone.Label21.Caption = STR1
samsungduos2.Label18.Visible = False
samsungduos2.Label17.Visible = False
samsungduos2.Label19.Visible = True
samsungduos2.Label20.Visible = True
samsungduos2.Label21.Caption = STR1
Kangaroo.Label18.Visible = False
Kangaroo.Label17.Visible = False
Kangaroo.Label19.Visible = True
Kangaroo.Label20.Visible = True
Kangaroo.Label21.Caption = STR1
Tab3.Label18.Visible = False
Tab3.Label17.Visible = False
Tab3.Label19.Visible = True
Tab3.Label20.Visible = True
Tab3.Label21.Caption = STR1
Sunglasses.Label18.Visible = False
Sunglasses.Label17.Visible = False
Sunglasses.Label19.Visible = True
Sunglasses.Label20.Visible = True
Sunglasses.Label21.Caption = STR1
Sunlist.Label18.Visible = False
Sunlist.Label17.Visible = False
Sunlist.Label19.Visible = True
Sunlist.Label20.Visible = True
Sunlist.Label21.Caption = STR1
Sunglasses2.Label18.Visible = False
Sunglasses2.Label17.Visible = False
Sunglasses2.Label19.Visible = True
Sunglasses2.Label20.Visible = True
Sunglasses2.Label21.Caption = STR1
Cloth1.Label18.Visible = False
Cloth1.Label17.Visible = False
Cloth1.Label19.Visible = True
Cloth1.Label20.Visible = True
Cloth1.Label21.Caption = STR1
Mobileacc1.Label18.Visible = False
Mobileacc1.Label17.Visible = False
Mobileacc1.Label19.Visible = True
Mobileacc1.Label20.Visible = True
Mobileacc1.Label21.Caption = STR1
Mobilelist.Label18.Visible = False
Mobilelist.Label17.Visible = False
Mobilelist.Label19.Visible = True
Mobilelist.Label20.Visible = True
Mobilelist.Label21.Caption = STR1
Cloth2.Label18.Visible = False
Cloth2.Label17.Visible = False
Cloth2.Label19.Visible = True
Cloth2.Label20.Visible = True
Cloth2.Label21.Caption = STR1
Clothes.Label18.Visible = False
Clothes.Label17.Visible = False
Clothes.Label19.Visible = True
Clothes.Label20.Visible = True
Clothes.Label21.Caption = STR1
Cart.Label20.Visible = True
Cart.Label21.Caption = STR1
Cart.Label1.Visible = True
login1 = 1
Unload Me
Load Flipkart
Flipkart.Show
Exit Sub
Else
MsgBox "Password invalid", vbInformation, "login"
Text3.SetFocus
Exit Sub
End If
End If
Set rs = Nothing
End Sub

Private Sub Image17_Click()
Flipkart.Show
Me.Hide
End Sub
Private Sub Image4_Click()
Flipkart.Show
Me.Hide
End Sub

Private Sub Label5_Click()
fr = Text4.Text
Call search
End Sub
Private Sub Label17_Click()
SignUp.Show
End Sub
Private Sub Form_Activate()
Text4.Text = "Seach any product"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.FontBold = False
Label17.FontUnderline = False
Label18.FontBold = False
Label18.FontUnderline = False
End Sub
Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.FontBold = True
Label17.FontUnderline = True
Label18.FontBold = False
Label18.FontUnderline = False
End Sub
Private Sub Label18_Click()
Me.Show
End Sub
Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.FontBold = True
Label18.FontUnderline = True
Label17.FontBold = False
Label17.FontUnderline = False
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

Private Sub Label7_Click()
ForgotID.Show
Me.Hide
End Sub

Private Sub Label8_Click()
SignUp.Show
Me.Hide
End Sub

Private Sub Text4_Click()
Text4.Text = ""
End Sub
