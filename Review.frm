VERSION 5.00
Begin VB.Form Review 
   BackColor       =   &H80000018&
   Caption         =   "Review"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Submit"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   1320
      Picture         =   "Review.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Write Your Review:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   4575
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
Attribute VB_Name = "Review"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "Thanks for your review,highly appreciated", vbOKOnly, "Thanks"
Flipkart.Show
Me.Hide
End Sub
