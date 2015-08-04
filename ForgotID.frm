VERSION 5.00
Begin VB.Form ForgotID 
   BackColor       =   &H8000000E&
   Caption         =   "Reset Password"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Submit"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   1920
      Picture         =   "ForgotID.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "Label4"
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   20295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Email ID:"
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
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
   End
End
Attribute VB_Name = "ForgotID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "A reset link has been sent to your ID", vbInformation, "Password Reset"
Flipkart.Show
Me.Hide
End Sub

Private Sub Image4_Click()
Flipkart.Show
Me.Hide
End Sub
